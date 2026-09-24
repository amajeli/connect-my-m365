// Pure helpers for the v3 MCP server. v3 returns pipe-delimited rows instead of
// JSON (no repeated keys) and swaps long Graph IDs for short per-user handles
// like "m3" — a Graph message ID alone is ~150 characters of hard-to-tokenize text.

export type HandleKind = "m" | "e" | "f" | "c";

export interface HandleTarget {
  accountId: string;
  accountEmail: string;
  ref: string;
}

const MAX_HANDLES_PER_USER = 2000;

interface UserHandles {
  next: Record<HandleKind, number>;
  byHandle: Map<string, HandleTarget>;
  byRef: Map<string, string>;
}

// In memory on purpose: handles are a token-saving convenience, not durable
// state. After a restart, tools say the handle expired and to search again.
const handleStore = new Map<string, UserHandles>();

function userHandles(userId: string): UserHandles {
  let h = handleStore.get(userId);
  if (!h) {
    h = { next: { m: 1, e: 1, f: 1, c: 1 }, byHandle: new Map(), byRef: new Map() };
    handleStore.set(userId, h);
  }
  return h;
}

export function toHandle(userId: string, kind: HandleKind, target: HandleTarget): string {
  const h = userHandles(userId);
  const key = `${kind}:${target.accountId}:${target.ref}`;
  const existing = h.byRef.get(key);
  if (existing && h.byHandle.has(existing)) return existing;

  const handle = `${kind}${h.next[kind]++}`;
  h.byHandle.set(handle, target);
  h.byRef.set(key, handle);

  if (h.byHandle.size > MAX_HANDLES_PER_USER) {
    const oldest = h.byHandle.keys().next().value as string;
    const t = h.byHandle.get(oldest)!;
    h.byHandle.delete(oldest);
    h.byRef.delete(`${oldest[0]}:${t.accountId}:${t.ref}`);
  }
  return handle;
}

export function fromHandle(userId: string, kind: HandleKind, handle: string): HandleTarget | null {
  if (!new RegExp(`^${kind}\\d+$`).test(handle)) return null;
  return userHandles(userId).byHandle.get(handle) ?? null;
}

export function resetHandlesForTest(): void {
  handleStore.clear();
}

/** Makes a value safe for a pipe-delimited row. */
export function cell(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).replace(/[|\r\n\t]+/g, " ").replace(/\s{2,}/g, " ").trim();
}

export function row(values: unknown[]): string {
  return values.map(cell).join("|").replace(/\|+$/, "");
}

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

function pad(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/** Graph returns UTC datetimes, sometimes without a trailing "Z". */
function parseUtc(iso: string): Date {
  return new Date(/[zZ]|[+-]\d{2}:?\d{2}$/.test(iso) ? iso : `${iso}Z`);
}

/** "09-24 14:05" (UTC). */
export function shortDateTime(iso: string | null | undefined): string {
  if (!iso) return "";
  const d = parseUtc(iso);
  if (isNaN(d.getTime())) return "";
  return `${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
}

/** "Wed 09-24 14:00-15:00", "Wed 09-24 all-day", or a span across days. */
export function eventWhen(start: string, end: string, allDay: boolean): string {
  const s = parseUtc(start);
  const e = parseUtc(end);
  const day = (d: Date) => `${WEEKDAYS[d.getUTCDay()]} ${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  const time = (d: Date) => `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`;
  if (allDay) {
    const lastDay = new Date(e.getTime() - 1);
    return day(s) === day(lastDay) ? `${day(s)} all-day` : `${day(s)}..${day(lastDay)} all-day`;
  }
  return day(s) === day(e) ? `${day(s)} ${time(s)}-${time(e)}` : `${day(s)} ${time(s)}-${day(e)} ${time(e)}`;
}

export function person(p: { name?: string; address?: string } | null | undefined): string {
  if (!p) return "";
  const name = (p.name || "").trim();
  const addr = (p.address || "").trim();
  if (!addr) return name;
  if (!name || name.toLowerCase() === addr.toLowerCase()) return addr;
  return `${name} <${addr}>`;
}

export function humanSize(bytes: number | null | undefined): string {
  if (bytes === null || bytes === undefined) return "";
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)}K`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)}M`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)}G`;
}

/**
 * Shrinks email/chat text: collapses whitespace and replaces long (usually
 * tracking) URLs with their domain.
 */
export function compactText(text: string | null | undefined, maxChars: number): string {
  if (!text) return "";
  let t = text
    .replace(/https?:\/\/[^\s<>"')\]]+/g, (url) => {
      if (url.length <= 80) return url;
      try {
        return `[link:${new URL(url).hostname}]`;
      } catch {
        return "[link]";
      }
    })
    .replace(/[ \t\u00a0]+/g, " ")
    .replace(/ *\n[ \n]*/g, (m) => (m.split("\n").length > 2 ? "\n\n" : "\n"))
    .trim();
  if (t.length > maxChars) {
    t = `${t.slice(0, maxChars)}…[+${t.length - maxChars} chars]`;
  }
  return t;
}

/** Strips characters that would break out of a quoted KQL $search string. */
export function kqlTerm(value: string): string {
  return value.replace(/["\\]/g, " ").trim();
}
