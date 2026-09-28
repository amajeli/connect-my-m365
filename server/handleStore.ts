// Short ids for the v3 MCP server: "m3" instead of a ~150-character Graph id.
// Stored in Postgres so they survive serverless cold starts and work across
// parallel instances (Netlify). A small in-process cache avoids repeat lookups.
// Without DATABASE_URL (tests, quick local runs) an in-memory store is used.
import { createHash } from "crypto";
import { and, eq, inArray, max, sql } from "drizzle-orm";
import { mcpHandles } from "@shared/schema";

export type HandleKind = "m" | "e" | "f" | "c";

export interface HandleTarget {
  accountId: string;
  accountEmail: string;
  ref: string;
}

interface Backend {
  findByRef(userId: string, refKeys: string[]): Promise<Map<string, string>>;
  insert(userId: string, kind: HandleKind, rows: { refKey: string; target: HandleTarget }[]): Promise<Map<string, string>>;
  get(userId: string, handle: string): Promise<HandleTarget | null>;
}

function refKeyFor(kind: HandleKind, target: HandleTarget): string {
  const digest = createHash("sha256").update(`${target.accountId}\0${target.ref}`).digest("base64url");
  return `${kind}:${digest.slice(0, 24)}`;
}

export function memoryBackend(): Backend {
  const byHandle = new Map<string, HandleTarget>();
  const byRef = new Map<string, string>();
  const next = new Map<string, number>();
  return {
    async findByRef(userId, refKeys) {
      const found = new Map<string, string>();
      for (const k of refKeys) {
        const h = byRef.get(`${userId}|${k}`);
        if (h) found.set(k, h);
      }
      return found;
    },
    async insert(userId, kind, rows) {
      const out = new Map<string, string>();
      for (const { refKey, target } of rows) {
        const counterKey = `${userId}|${kind}`;
        const n = (next.get(counterKey) ?? 0) + 1;
        next.set(counterKey, n);
        const handle = `${kind}${n}`;
        byHandle.set(`${userId}|${handle}`, target);
        byRef.set(`${userId}|${refKey}`, handle);
        out.set(refKey, handle);
      }
      return out;
    },
    async get(userId, handle) {
      return byHandle.get(`${userId}|${handle}`) ?? null;
    },
  };
}

function postgresBackend(): Backend {
  // Imported lazily so modules that never touch handles don't open a pool.
  const dbPromise = import("./db").then((m) => m.db);
  const self: Backend = {
    async findByRef(userId, refKeys) {
      const db = await dbPromise;
      const rows = await db
        .select({ handle: mcpHandles.handle, refKey: mcpHandles.refKey })
        .from(mcpHandles)
        .where(and(eq(mcpHandles.userId, userId), inArray(mcpHandles.refKey, refKeys)));
      return new Map(rows.map((r) => [r.refKey, r.handle]));
    },
    async insert(userId, kind, rows) {
      const db = await dbPromise;
      const out = new Map<string, string>();
      // Parallel requests for one user take turns (a transaction-scoped
      // advisory lock), so numbering never races and nothing needs retrying.
      await db.transaction(async (tx) => {
        await tx.execute(sql`select pg_advisory_xact_lock(hashtext(${`mcp_handles:${userId}:${kind}`}))`);
        const refKeys = rows.map((r) => r.refKey);
        const existing = await tx
          .select({ handle: mcpHandles.handle, refKey: mcpHandles.refKey })
          .from(mcpHandles)
          .where(and(eq(mcpHandles.userId, userId), inArray(mcpHandles.refKey, refKeys)));
        for (const r of existing) out.set(r.refKey, r.handle);
        const fresh = rows.filter((r) => !out.has(r.refKey));
        if (!fresh.length) return;
        const [{ top }] = await tx
          .select({ top: max(mcpHandles.num) })
          .from(mcpHandles)
          .where(and(eq(mcpHandles.userId, userId), eq(mcpHandles.kind, kind)));
        const start = (top ?? 0) + 1;
        const values = fresh.map((r, i) => ({
          userId,
          kind,
          num: start + i,
          handle: `${kind}${start + i}`,
          refKey: r.refKey,
          accountId: r.target.accountId,
          accountEmail: r.target.accountEmail,
          ref: r.target.ref,
        }));
        await tx.insert(mcpHandles).values(values);
        for (const v of values) out.set(v.refKey, v.handle);
      });

      if (Math.random() < 0.02) {
        // Drop ids older than 30 days, but keep each kind's highest number so
        // numbering never restarts: an old "e1" must never point at a new event.
        db.execute(sql`
          delete from mcp_handles h
          where h.user_id = ${userId}
            and h.created_at < now() - interval '30 days'
            and h.num < (select max(m.num) from mcp_handles m where m.user_id = h.user_id and m.kind = h.kind)
        `).catch((err) => console.error("mcp_handles prune failed:", err.message));
      }
      return out;
    },
    async get(userId, handle) {
      const db = await dbPromise;
      const [row] = await db
        .select({ accountId: mcpHandles.accountId, accountEmail: mcpHandles.accountEmail, ref: mcpHandles.ref })
        .from(mcpHandles)
        .where(and(eq(mcpHandles.userId, userId), eq(mcpHandles.handle, handle)));
      return row ?? null;
    },
  };
  return self;
}

let backend: Backend = process.env.DATABASE_URL ? postgresBackend() : memoryBackend();

export function setHandleBackendForTest(b: Backend): void {
  backend = b;
  cache.clear();
}

const CACHE_LIMIT = 5000;
const cache = new Map<string, string | HandleTarget>();

function remember(key: string, value: string | HandleTarget) {
  cache.set(key, value);
  if (cache.size > CACHE_LIMIT) cache.delete(cache.keys().next().value as string);
}

/** Returns a short id for each target, in order, reusing existing ids. */
export async function toHandles(userId: string, kind: HandleKind, targets: HandleTarget[]): Promise<string[]> {
  const keys = targets.map((t) => refKeyFor(kind, t));
  const resolved = new Map<string, string>();
  const unknown: string[] = [];
  for (const k of new Set(keys)) {
    const hit = cache.get(`${userId}|${k}`);
    if (typeof hit === "string") resolved.set(k, hit);
    else unknown.push(k);
  }
  if (unknown.length) {
    const found = await backend.findByRef(userId, unknown);
    found.forEach((h, k) => resolved.set(k, h));
    const toInsert = unknown
      .filter((k) => !resolved.has(k))
      .map((k) => ({ refKey: k, target: targets[keys.indexOf(k)] }));
    if (toInsert.length) {
      const inserted = await backend.insert(userId, kind, toInsert);
      inserted.forEach((h, k) => resolved.set(k, h));
    }
  }
  return keys.map((k, i) => {
    const handle = resolved.get(k)!;
    remember(`${userId}|${k}`, handle);
    remember(`${userId}|${handle}`, targets[i]);
    return handle;
  });
}

export async function toHandle(userId: string, kind: HandleKind, target: HandleTarget): Promise<string> {
  return (await toHandles(userId, kind, [target]))[0];
}

export async function fromHandle(userId: string, kind: HandleKind, handle: string): Promise<HandleTarget | null> {
  if (!new RegExp(`^${kind}\\d+$`).test(handle)) return null;
  const hit = cache.get(`${userId}|${handle}`);
  if (hit && typeof hit !== "string") return hit;
  const target = await backend.get(userId, handle);
  if (target) remember(`${userId}|${handle}`, target);
  return target;
}
