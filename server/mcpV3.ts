import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import type { Express, Request, Response } from "express";
import type { ConnectedAccount, User } from "@shared/schema";
import { storage } from "./storage";
import {
  getValidAccessToken,
  createEmailDraft,
  listDriveItems,
  searchDriveItems,
  getDriveItemContent,
  getTeamsChats,
  getTeamsChatMessages,
  searchTeamsMessages,
} from "./microsoft";
import {
  validateMcpAuth,
  checkSubscription,
  checkFullAccessRequired,
  isFullAccess,
  getFreeEmailsUsedToday,
  incrementFreeEmailCount,
  FREE_EMAIL_DAILY_LIMIT,
  MAX_BODY_CHARS,
  MAX_FILE_CHARS,
} from "./mcp";
import {
  stripHtml,
  isBlockedSender,
  isNewsletter,
  parseDateInput,
  isWithinAllowed,
  senderMatches,
  getEffectiveSettings,
  emailAllowedBySettings,
} from "./mcpShared";
import {
  listInboxMessages,
  searchMessages,
  getMessageText,
  createReplyDraft,
  getCalendarView,
  createEvent,
  updateEvent,
  deleteEvent,
} from "./graphV3";
import {
  toHandle,
  fromHandle,
  row,
  shortDateTime,
  eventWhen,
  person,
  humanSize,
  compactText,
  kqlTerm,
  type HandleKind,
} from "./mcpV3Format";

// MCP v3: built on v2's filters and allow-lists, tuned further for tokens.
//  - 10 short tools instead of 16. Tool definitions are re-sent every turn.
//  - Pipe-delimited rows instead of JSON, and short handles (m1, e1, f1, c1) instead of Graph IDs.
//  - mail_read returns only the new text of a message (Graph uniqueBody), not the quoted thread.
//  - Stateless transport: no in-memory sessions to lose on restart or across instances.
//  - Free tier only sees the tool it can use.

type ToolResult = { content: { type: "text"; text: string }[]; isError?: boolean };

const ok = (text: string): ToolResult => ({ content: [{ type: "text", text }] });
const fail = (text: string): ToolResult => ({ content: [{ type: "text", text }], isError: true });

const READ = { readOnlyHint: true, openWorldHint: true } as const;
const WRITE = { readOnlyHint: false, destructiveHint: false, openWorldHint: true } as const;
const DESTROY = { readOnlyHint: false, destructiveHint: true, openWorldHint: true } as const;

const DAY_MS = 86400000;

function accountList(accounts: ConnectedAccount[]): string {
  return accounts.map((a) => a.email).join(", ");
}

/** Accounts a read tool should cover: the named one, or all of them. */
async function targetAccounts(user: User, account?: string): Promise<ConnectedAccount[] | string> {
  const all = await storage.getAccountsByUserId(user.id);
  if (all.length === 0) return "No Microsoft accounts connected. Connect one in the web app.";
  if (!account) return all;
  const match = all.filter((a) => a.email.toLowerCase() === account.toLowerCase());
  return match.length ? match : `Unknown account ${account}. Connected: ${accountList(all)}`;
}

/** The single account a write tool acts on. */
async function oneAccount(user: User, account?: string): Promise<ConnectedAccount | string> {
  const all = await storage.getAccountsByUserId(user.id);
  if (all.length === 0) return "No Microsoft accounts connected. Connect one in the web app.";
  if (account) {
    const match = all.find((a) => a.email.toLowerCase() === account.toLowerCase());
    return match ?? `Unknown account ${account}. Connected: ${accountList(all)}`;
  }
  if (all.length === 1) return all[0];
  return `Several accounts connected; pass account (${accountList(all)}).`;
}

/** Resolves a handle like "m3", or a raw Graph ID plus account. */
async function resolveRef(
  user: User,
  kind: HandleKind,
  id: string,
  account?: string,
): Promise<{ accountId: string; accountEmail: string; ref: string } | string> {
  const target = fromHandle(user.id, kind, id);
  if (target) return target;
  if (new RegExp(`^${kind}\\d+$`).test(id)) {
    return `${id} has expired (handles reset when the server restarts). Search again to get a new id.`;
  }
  const acct = await oneAccount(user, account);
  if (typeof acct === "string") return acct;
  return { accountId: acct.id, accountEmail: acct.email, ref: id };
}

function gate(user: User, tool: string): string | null {
  return checkSubscription(user) || checkFullAccessRequired(user, tool);
}

function errMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

async function buildInstructions(user: User): Promise<string> {
  const [accounts, settings] = await Promise.all([
    storage.getAccountsByUserId(user.id),
    getEffectiveSettings(user.id),
  ]);
  const lines = [
    `Microsoft 365 for ${user.email}.`,
    accounts.length
      ? `Accounts: ${accounts.map((a) => `${a.email} (${a.accountType})`).join(", ")}.`
      : "No accounts connected yet.",
    "All times are UTC. Results are pipe-delimited rows; the first row names the columns.",
    "Ids like m1/e1/f1/c1 are short handles for emails/events/files/chats. Pass them back as-is.",
    "Email flags: U=unread A=attachment !=high importance.",
  ];
  if (settings.focusedOnly) lines.push("Mail is limited to the Focused inbox.");
  if (settings.approvedSenders.length) lines.push(`Mail is limited to ${settings.approvedSenders.length} approved senders.`);
  if (settings.allowedFolders.length) lines.push(`OneDrive is limited to: ${settings.allowedFolders.join(", ")}.`);
  if (!isFullAccess(user)) lines.push(`Free tier: mail_find only, ${FREE_EMAIL_DAILY_LIMIT} emails/day.`);
  return lines.join("\n");
}

function createMcpV3Server(user: User, instructions?: string): McpServer {
  const server = new McpServer(
    { name: "microsoft365-connector-v3", version: "3.0.0" },
    instructions ? { instructions } : undefined,
  );
  const full = isFullAccess(user);

  // ─── Mail ──────────────────────────────────────────────────────────────────

  server.registerTool(
    "mail_find",
    {
      description:
        "Find emails. With no q or from: an inbox briefing of the last 18h that hides automated senders and newsletters. With q: searches every folder (KQL keywords). Returns rows; read one with mail_read.",
      inputSchema: {
        q: z.string().optional().describe("Search keywords"),
        from: z.string().optional().describe("Sender address or name"),
        after: z.string().optional().describe("'today', 'yesterday', YYYY-MM-DD or ISO datetime"),
        before: z.string().optional(),
        unread: z.boolean().optional(),
        account: z.string().optional(),
        limit: z.number().int().min(1).max(25).default(12),
        preview: z.number().int().min(0).max(255).default(100).describe("Preview chars, 0 for none"),
        include_noise: z.boolean().default(false),
      },
      annotations: READ,
    },
    async ({ q, from, after, before, unread, account, limit, preview, include_noise }) => {
      const subErr = checkSubscription(user);
      if (subErr) return fail(subErr);
      if (q && !full) return fail(checkFullAccessRequired(user, "mail_find with q")!);

      if (!full) {
        const used = getFreeEmailsUsedToday(user.id);
        if (used >= FREE_EMAIL_DAILY_LIMIT) {
          return fail(`Free tier daily limit reached (${FREE_EMAIL_DAILY_LIMIT} emails). Resets at midnight UTC.`);
        }
        limit = Math.min(limit, FREE_EMAIL_DAILY_LIMIT - used);
      }

      const accounts = await targetAccounts(user, account);
      if (typeof accounts === "string") return fail(accounts);
      const settings = await getEffectiveSettings(user.id);

      const briefing = !q && !from;
      const afterIso = parseDateInput(after ?? (q ? undefined : new Date(Date.now() - 18 * 3600000).toISOString()), "start");
      const beforeIso = parseDateInput(before, "end");
      const fromIsAddress = Boolean(from && from.includes("@") && !from.includes(" "));
      const fetchTop = Math.min(50, Math.max(limit * 3, 25));

      const errors: string[] = [];
      let all: any[] = [];
      for (const acct of accounts) {
        try {
          const token = await getValidAccessToken(acct.id);
          let messages: any[];
          if (q) {
            const kql = [kqlTerm(q)];
            if (from) kql.push(`from:${kqlTerm(from)}`);
            if (afterIso) kql.push(`received>=${afterIso.slice(0, 10)}`);
            if (beforeIso) kql.push(`received<=${beforeIso.slice(0, 10)}`);
            messages = await searchMessages(token, kql.join(" "), fetchTop);
          } else {
            messages = await listInboxMessages(token, {
              top: fetchTop,
              after: afterIso,
              before: beforeIso,
              from: fromIsAddress ? from : undefined,
              unreadOnly: unread === true,
              focusedOnly: settings.focusedOnly,
            });
          }
          for (const m of messages) m._acct = acct;
          all.push(...messages);
        } catch (err) {
          errors.push(`! ${acct.email}: ${errMessage(err)}`);
        }
      }

      const afterMs = afterIso ? Date.parse(afterIso) : -Infinity;
      const beforeMs = beforeIso ? Date.parse(beforeIso) : Infinity;
      const fromLower = from?.toLowerCase();
      all = all.filter((m) => {
        const addr = m.from?.emailAddress?.address || "";
        const t = Date.parse(m.receivedDateTime);
        if (t < afterMs || t > beforeMs) return false;
        if (unread !== undefined && m.isRead === unread) return false;
        if (fromLower && !fromIsAddress && !`${m.from?.emailAddress?.name || ""} ${addr}`.toLowerCase().includes(fromLower)) return false;
        if (settings.focusedOnly && m.inferenceClassification && m.inferenceClassification !== "focused") return false;
        if (!senderMatches(addr, settings.approvedSenders)) return false;
        return true;
      });

      let hiddenNoise = 0;
      if (briefing && !include_noise) {
        const before = all.length;
        all = all.filter((m) => !isBlockedSender(m.from?.emailAddress?.address || "") && !isNewsletter(m));
        hiddenNoise = before - all.length;
      }

      all.sort((a, b) => Date.parse(b.receivedDateTime) - Date.parse(a.receivedDateTime));
      const shown = all.slice(0, limit);
      if (!full && shown.length) incrementFreeEmailCount(user.id, shown.length);

      const out: string[] = [];
      if (shown.length === 0) {
        out.push("No matching emails.");
      } else {
        out.push(preview ? "id|received|from|subject|flags|preview" : "id|received|from|subject|flags");
        const multi = new Set(shown.map((m) => m._acct.id)).size > 1;
        let currentAcct = "";
        const ordered = multi ? [...shown].sort((a, b) => a._acct.email.localeCompare(b._acct.email)) : shown;
        for (const m of ordered) {
          if (multi && m._acct.email !== currentAcct) {
            currentAcct = m._acct.email;
            out.push(`# ${currentAcct}`);
          }
          const flags = `${m.isRead ? "" : "U"}${m.hasAttachments ? "A" : ""}${m.importance === "high" ? "!" : ""}`;
          const values: unknown[] = [
            toHandle(user.id, "m", { accountId: m._acct.id, accountEmail: m._acct.email, ref: m.id }),
            shortDateTime(m.receivedDateTime),
            person(m.from?.emailAddress),
            m.subject || "(no subject)",
            flags,
          ];
          if (preview) values.push(compactText(m.bodyPreview, preview));
          out.push(row(values));
        }
      }
      if (all.length > shown.length) out.push(`(${all.length - shown.length} more; raise limit or narrow the search)`);
      if (hiddenNoise) out.push(`(${hiddenNoise} automated/newsletter hidden; include_noise=true shows them)`);
      out.push(...errors);
      return ok(out.join("\n"));
    },
  );

  if (full) {
    server.registerTool(
      "mail_read",
      {
        description:
          "Read up to 5 emails by id. Returns only the new text of each message and leaves out the quoted thread; set full=true to include it.",
        inputSchema: {
          ids: z.array(z.string()).min(1).max(5).describe("Handles from mail_find, e.g. [\"m1\",\"m4\"]"),
          full: z.boolean().default(false),
          max_chars: z.number().int().min(200).max(20000).default(MAX_BODY_CHARS),
          account: z.string().optional().describe("Only needed for raw Graph ids"),
        },
        annotations: READ,
      },
      async ({ ids, full: wantFull, max_chars, account }) => {
        const settings = await getEffectiveSettings(user.id);
        const parts: string[] = [];
        for (const id of ids) {
          const target = await resolveRef(user, "m", id, account);
          if (typeof target === "string") {
            parts.push(`## ${id}\n${target}`);
            continue;
          }
          try {
            const token = await getValidAccessToken(target.accountId);
            const m = await getMessageText(token, target.ref);
            const denied = emailAllowedBySettings(m, settings);
            if (denied) {
              parts.push(`## ${id}\n${denied}`);
              continue;
            }
            const recipients = (list: any[] | undefined) => {
              const names = (list || []).map((r) => person(r.emailAddress));
              return names.length > 5 ? `${names.slice(0, 5).join(", ")} +${names.length - 5}` : names.join(", ");
            };
            const unique = m.uniqueBody?.content?.trim();
            const raw = !wantFull && unique ? unique : m.body?.content || "";
            const text = m.body?.contentType === "html" && raw === m.body?.content ? stripHtml(raw) : raw;
            const meta = [
              `from: ${person(m.from?.emailAddress)}`,
              `to: ${recipients(m.toRecipients)}`,
              m.ccRecipients?.length ? `cc: ${recipients(m.ccRecipients)}` : "",
              shortDateTime(m.receivedDateTime),
              m.hasAttachments ? "attachments" : "",
            ].filter(Boolean);
            parts.push(`## ${id} · ${m.subject || "(no subject)"}\n${meta.join(" | ")}\n${compactText(text, max_chars)}`);
          } catch (err) {
            parts.push(`## ${id}\nError: ${errMessage(err)}`);
          }
        }
        return ok(parts.join("\n\n"));
      },
    );

    server.registerTool(
      "mail_draft",
      {
        description:
          "Save an email draft; this never sends. Pass reply_to (an m-handle) to draft a reply that includes the quoted thread, or pass to plus subject for a new email. The user reviews and sends it from Outlook.",
        inputSchema: {
          body: z.string().describe("Plain text"),
          reply_to: z.string().optional(),
          reply_all: z.boolean().default(false),
          to: z.array(z.string()).optional(),
          cc: z.array(z.string()).optional(),
          subject: z.string().optional(),
          account: z.string().optional(),
        },
        annotations: WRITE,
      },
      async ({ body, reply_to, reply_all, to, cc, subject, account }) => {
        try {
          if (reply_to) {
            const target = await resolveRef(user, "m", reply_to, account);
            if (typeof target === "string") return fail(target);
            const token = await getValidAccessToken(target.accountId);
            const draft = await createReplyDraft(token, target.ref, body, reply_all);
            return ok(`Draft saved, not sent (${target.accountEmail}): "${draft.subject}"\n${draft.webLink || ""}`.trim());
          }
          if (!to?.length || !subject) return fail("A new draft needs to and subject (or pass reply_to).");
          const acct = await oneAccount(user, account);
          if (typeof acct === "string") return fail(acct);
          const token = await getValidAccessToken(acct.id);
          const draft = await createEmailDraft(token, { to, cc, subject, body, bodyType: "Text" });
          return ok(`Draft saved, not sent (${acct.email}): "${subject}"\n${draft.webLink || ""}`.trim());
        } catch (err) {
          return fail(`Draft failed: ${errMessage(err)}. If this is a permissions error, reconnect the account to grant Mail.ReadWrite.`);
        }
      },
    );

    // ─── Calendar ────────────────────────────────────────────────────────────

    server.registerTool(
      "calendar_find",
      {
        description:
          "List calendar events in a time window. Defaults to today through the next 7 days. Flags: O=online meeting T=tentative F=free W=working elsewhere OOF=out of office X=cancelled.",
        inputSchema: {
          after: z.string().optional().describe("'today', 'tomorrow', YYYY-MM-DD or ISO datetime"),
          before: z.string().optional(),
          q: z.string().optional().describe("Matches subject, location or organizer"),
          account: z.string().optional(),
          limit: z.number().int().min(1).max(100).default(25),
          links: z.boolean().default(false).describe("Include meeting join URLs"),
        },
        annotations: READ,
      },
      async ({ after, before, q, account, limit, links }) => {
        const accounts = await targetAccounts(user, account);
        if (typeof accounts === "string") return fail(accounts);
        const start = parseDateInput(after ?? "today", "start") || new Date().toISOString();
        const end = parseDateInput(before, "end") || new Date(Date.parse(start) + 7 * DAY_MS).toISOString();
        const needle = q?.toLowerCase();

        const errors: string[] = [];
        let all: any[] = [];
        for (const acct of accounts) {
          try {
            const token = await getValidAccessToken(acct.id);
            const events = await getCalendarView(token, start, end, 200);
            for (const e of events) e._acct = acct;
            all.push(...events);
          } catch (err) {
            errors.push(`! ${acct.email}: ${errMessage(err)}`);
          }
        }
        if (needle) {
          all = all.filter((e) =>
            `${e.subject || ""} ${e.location?.displayName || ""} ${e.organizer?.emailAddress?.name || ""} ${e.organizer?.emailAddress?.address || ""}`
              .toLowerCase()
              .includes(needle),
          );
        }
        all.sort((a, b) => Date.parse(`${a.start.dateTime}Z`) - Date.parse(`${b.start.dateTime}Z`));
        const shown = all.slice(0, limit);

        const out: string[] = [];
        if (shown.length === 0) {
          out.push("No events in that window.");
        } else {
          const multi = accounts.length > 1;
          out.push(`id|when|subject|location|organizer|flags${multi ? "|account" : ""}${links ? "|join" : ""}`);
          for (const e of shown) {
            const showAs: Record<string, string> = { free: "F", tentative: "T", oof: "OOF", workingElsewhere: "W" };
            const flags = `${e.isOnlineMeeting ? "O" : ""}${showAs[e.showAs] || ""}${e.isCancelled ? "X" : ""}`;
            const orgAddr = e.organizer?.emailAddress?.address || "";
            const values: unknown[] = [
              toHandle(user.id, "e", { accountId: e._acct.id, accountEmail: e._acct.email, ref: e.id }),
              eventWhen(e.start.dateTime, e.end.dateTime, e.isAllDay),
              e.subject || "(no subject)",
              e.location?.displayName || "",
              orgAddr.toLowerCase() === e._acct.email.toLowerCase() ? "me" : person(e.organizer?.emailAddress),
              flags,
            ];
            if (multi) values.push(e._acct.email);
            if (links) values.push(e.onlineMeeting?.joinUrl || e.onlineMeetingUrl || "");
            out.push(row(values));
          }
        }
        if (all.length > shown.length) out.push(`(${all.length - shown.length} more; raise limit or narrow the window)`);
        out.push(...errors);
        return ok(out.join("\n"));
      },
    );

    server.registerTool(
      "event_save",
      {
        description:
          "Create a calendar event, or update one when id is given (only the fields you pass change). Adding attendees sends them invitations.",
        inputSchema: {
          id: z.string().optional().describe("e-handle; if omitted, creates a new event"),
          account: z.string().optional(),
          subject: z.string().optional(),
          start: z.string().optional().describe("ISO datetime, or YYYY-MM-DD for all-day"),
          end: z.string().optional(),
          tz: z.string().default("UTC").describe("IANA or Windows time zone for start/end"),
          body: z.string().optional(),
          location: z.string().optional(),
          attendees: z.array(z.string()).optional().describe("Replaces the attendee list"),
          all_day: z.boolean().optional(),
          online: z.boolean().optional().describe("Add a Teams link"),
          reminder_min: z.number().int().min(0).max(40320).optional(),
        },
        annotations: WRITE,
      },
      async ({ id, account, subject, start, end, tz, body, location, attendees, all_day, online, reminder_min }) => {
        const fields: any = {};
        if (subject !== undefined) fields.subject = subject;
        if (start !== undefined) fields.start = { dateTime: start, timeZone: tz };
        if (end !== undefined) fields.end = { dateTime: end, timeZone: tz };
        if (body !== undefined) fields.body = { contentType: "text", content: body };
        if (location !== undefined) fields.location = { displayName: location };
        if (attendees !== undefined) {
          fields.attendees = attendees.map((address) => ({ emailAddress: { address }, type: "required" }));
        }
        if (all_day !== undefined) fields.isAllDay = all_day;
        if (online !== undefined) fields.isOnlineMeeting = online;
        if (reminder_min !== undefined) {
          fields.isReminderOn = true;
          fields.reminderMinutesBeforeStart = reminder_min;
        }

        try {
          let saved: any;
          let target: { accountId: string; accountEmail: string; ref: string };
          if (id) {
            const resolved = await resolveRef(user, "e", id, account);
            if (typeof resolved === "string") return fail(resolved);
            target = resolved;
            if (Object.keys(fields).length === 0) return fail("Nothing to update.");
            saved = await updateEvent(await getValidAccessToken(target.accountId), target.ref, fields);
          } else {
            if (!subject || !start || !end) return fail("A new event needs subject, start and end.");
            const acct = await oneAccount(user, account);
            if (typeof acct === "string") return fail(acct);
            target = { accountId: acct.id, accountEmail: acct.email, ref: "" };
            saved = await createEvent(await getValidAccessToken(acct.id), fields);
          }
          const handle = toHandle(user.id, "e", { ...target, ref: saved.id });
          const invited = attendees?.length ? ` · invitations sent to ${attendees.length}` : "";
          return ok(
            `${id ? "Updated" : "Created"} ${row([handle, eventWhen(saved.start.dateTime, saved.end.dateTime, saved.isAllDay), saved.subject])}${invited}`,
          );
        } catch (err) {
          return fail(`Event save failed: ${errMessage(err)}`);
        }
      },
    );

    server.registerTool(
      "event_delete",
      {
        description: "Delete a calendar event permanently. If the user organized it, attendees get a cancellation.",
        inputSchema: { id: z.string().describe("e-handle"), account: z.string().optional() },
        annotations: DESTROY,
      },
      async ({ id, account }) => {
        const target = await resolveRef(user, "e", id, account);
        if (typeof target === "string") return fail(target);
        try {
          await deleteEvent(await getValidAccessToken(target.accountId), target.ref);
          return ok(`Deleted ${id}.`);
        } catch (err) {
          return fail(`Delete failed: ${errMessage(err)}`);
        }
      },
    );

    // ─── OneDrive ────────────────────────────────────────────────────────────

    server.registerTool(
      "files_find",
      {
        description: "Browse a OneDrive folder (path) or search by name/keyword (q). Read a file with file_read.",
        inputSchema: {
          q: z.string().optional(),
          path: z.string().optional().describe("Folder path, e.g. 'Documents/Deals'; omit for the root"),
          account: z.string().optional(),
          limit: z.number().int().min(1).max(100).default(25),
        },
        annotations: READ,
      },
      async ({ q, path, account, limit }) => {
        const accounts = await targetAccounts(user, account);
        if (typeof accounts === "string") return fail(accounts);
        const { allowedFolders } = await getEffectiveSettings(user.id);
        const folder = path && path !== "root" && path !== "/" ? path.replace(/^\/+|\/+$/g, "") : "";

        if (!q && allowedFolders.length) {
          if (!folder) return ok(`OneDrive is limited to these folders; pass one as path:\n${allowedFolders.join("\n")}`);
          if (!isWithinAllowed(folder, allowedFolders)) {
            return fail(`${folder} is outside the allowed OneDrive folders: ${allowedFolders.join(", ")}`);
          }
        }

        const errors: string[] = [];
        const out: string[] = [];
        let total = 0;
        for (const acct of accounts) {
          try {
            const token = await getValidAccessToken(acct.id);
            let items: any[] = q
              ? await searchDriveItems(token, q, allowedFolders.length ? 50 : limit)
              : await listDriveItems(token, folder || "root", limit);
            const parentOf = (item: any) => (item.parentReference?.path || "").replace(/^\/drive\/root:/, "") || "/";
            if (allowedFolders.length) {
              items = items.filter((item) => isWithinAllowed(`${parentOf(item)}/${item.name}`, allowedFolders));
            }
            items = items.slice(0, limit);
            if (!items.length) continue;
            if (accounts.length > 1) out.push(`# ${acct.email}`);
            for (const item of items) {
              total++;
              out.push(
                row([
                  toHandle(user.id, "f", { accountId: acct.id, accountEmail: acct.email, ref: item.id }),
                  item.name,
                  item.folder ? "dir" : "file",
                  item.folder ? `${item.folder.childCount ?? ""} items` : humanSize(item.size),
                  shortDateTime(item.lastModifiedDateTime),
                  parentOf(item),
                ]),
              );
            }
          } catch (err) {
            errors.push(`! ${acct.email}: ${errMessage(err)}`);
          }
        }
        const header = total ? ["id|name|kind|size|modified|folder"] : [q ? `No files match "${q}".` : "Folder is empty."];
        return ok([...header, ...out, ...errors].join("\n"));
      },
    );

    server.registerTool(
      "file_read",
      {
        description: "Read a OneDrive text file (txt, md, csv, json and similar). For other types, returns a web link to open it.",
        inputSchema: {
          id: z.string().describe("f-handle from files_find"),
          max_chars: z.number().int().min(200).max(50000).default(MAX_FILE_CHARS),
          account: z.string().optional().describe("Only needed for raw Graph ids"),
        },
        annotations: READ,
      },
      async ({ id, max_chars, account }) => {
        const target = await resolveRef(user, "f", id, account);
        if (typeof target === "string") return fail(target);
        const { allowedFolders } = await getEffectiveSettings(user.id);
        try {
          const token = await getValidAccessToken(target.accountId);
          const { metadata, content, mimeType } = await getDriveItemContent(token, target.ref);
          const parent = (metadata.parentReference?.path || "").replace(/^\/drive\/root:/, "");
          if (!isWithinAllowed(`${parent}/${metadata.name}`, allowedFolders)) {
            return fail(`Access denied: this file is outside the allowed OneDrive folders.`);
          }
          const title = `## ${metadata.name} (${humanSize(metadata.size)}, ${shortDateTime(metadata.lastModifiedDateTime)})`;
          if (content !== null) return ok(`${title}\n${compactText(content, max_chars)}`);
          const why = metadata.folder ? "This is a folder; browse it with files_find path." : `Not readable as text (${mimeType}).`;
          return ok(`${title}\n${why}\n${metadata.webUrl || ""}`.trim());
        } catch (err) {
          return fail(`File read failed: ${errMessage(err)}`);
        }
      },
    );

    // ─── Teams (work accounts only) ──────────────────────────────────────────

    server.registerTool(
      "teams_find",
      {
        description: "Without q: list recent Teams chats. With q: search Teams messages. Work accounts only. Read a chat with teams_read.",
        inputSchema: {
          q: z.string().optional(),
          account: z.string().optional(),
          limit: z.number().int().min(1).max(50).default(15),
        },
        annotations: READ,
      },
      async ({ q, account, limit }) => {
        const accounts = await targetAccounts(user, account);
        if (typeof accounts === "string") return fail(accounts);
        const work = accounts.filter((a) => a.accountType !== "personal");
        if (!work.length) return fail("Teams needs a work account; none is connected.");

        const out: string[] = [];
        const errors: string[] = [];
        for (const acct of work) {
          try {
            const token = await getValidAccessToken(acct.id);
            const rows: string[] = [];
            if (q) {
              const hits = await searchTeamsMessages(token, q, limit);
              for (const hit of hits) {
                const r = hit.resource || {};
                if (!r.chatId) continue;
                rows.push(
                  row([
                    toHandle(user.id, "c", { accountId: acct.id, accountEmail: acct.email, ref: r.chatId }),
                    shortDateTime(r.createdDateTime),
                    r.from?.user?.displayName || "",
                    compactText(stripHtml(hit.summary || r.body?.content || ""), 200),
                  ]),
                );
              }
            } else {
              const chats = await getTeamsChats(token, limit);
              for (const chat of chats) {
                const others = (chat.members || [])
                  .filter((m: any) => (m.email || "").toLowerCase() !== acct.email.toLowerCase())
                  .map((m: any) => m.displayName)
                  .filter(Boolean);
                const who = others.length > 4 ? `${others.slice(0, 4).join(", ")} +${others.length - 4}` : others.join(", ");
                rows.push(
                  row([
                    toHandle(user.id, "c", { accountId: acct.id, accountEmail: acct.email, ref: chat.id }),
                    chat.chatType,
                    chat.topic || who,
                    shortDateTime(chat.lastUpdatedDateTime),
                  ]),
                );
              }
            }
            if (rows.length) {
              if (work.length > 1) out.push(`# ${acct.email}`);
              out.push(...rows);
            }
          } catch (err) {
            errors.push(`! ${acct.email}: ${errMessage(err)}`);
          }
        }
        const header = out.length ? [q ? "chat|sent|sender|snippet" : "chat|type|topic or members|last active"] : [q ? `No Teams messages match "${q}".` : "No Teams chats."];
        return ok([...header, ...out, ...errors].join("\n"));
      },
    );

    server.registerTool(
      "teams_read",
      {
        description: "Read recent messages in a Teams chat, oldest first.",
        inputSchema: {
          id: z.string().describe("c-handle from teams_find"),
          limit: z.number().int().min(1).max(50).default(20),
          account: z.string().optional().describe("Only needed for raw Graph ids"),
        },
        annotations: READ,
      },
      async ({ id, limit, account }) => {
        const target = await resolveRef(user, "c", id, account);
        if (typeof target === "string") return fail(target);
        try {
          const token = await getValidAccessToken(target.accountId);
          const messages = await getTeamsChatMessages(token, target.ref, limit);
          const lines = messages
            .filter((m: any) => m.messageType === "message" && !m.deletedDateTime)
            .reverse()
            .map((m: any) => {
              const body = m.body?.contentType === "html" ? stripHtml(m.body.content) : m.body?.content || "";
              return `${shortDateTime(m.createdDateTime)} ${m.from?.user?.displayName || "?"}: ${compactText(body, 600).replace(/\n+/g, " ")}`;
            });
          return ok(lines.length ? lines.join("\n") : "No messages in this chat.");
        } catch (err) {
          return fail(`Teams read failed: ${errMessage(err)}`);
        }
      },
    );
  }

  return server;
}

function methodNotAllowed(_req: Request, res: Response) {
  res.setHeader("Allow", "POST");
  res.status(405).json({ jsonrpc: "2.0", error: { code: -32000, message: "Method not allowed; v3 is stateless and uses POST only" }, id: null });
}

export function registerMcpV3Routes(app: Express): void {
  app.post("/mcp/v3", async (req: Request, res: Response) => {
    // Points clients at OAuth discovery if auth fails (RFC 9728).
    const base = `${req.protocol}://${req.get("host")}`;
    res.setHeader("WWW-Authenticate", `Bearer resource_metadata="${base}/.well-known/oauth-protected-resource/mcp/v3"`);
    const user = await validateMcpAuth(req, res);
    if (!user) return;
    res.removeHeader("WWW-Authenticate");

    // A fresh server per request: nothing held in memory between calls except short handles.
    const instructions = isInitializeRequest(req.body) ? await buildInstructions(user) : undefined;
    const server = createMcpV3Server(user, instructions);
    const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });
    res.on("close", () => {
      transport.close();
      server.close();
    });
    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err) {
      console.error("MCP v3 request error:", errMessage(err));
      if (!res.headersSent) {
        res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal server error" }, id: null });
      }
    }
  });

  app.get("/mcp/v3", methodNotAllowed);
  app.delete("/mcp/v3", methodNotAllowed);

  console.log("MCP v3 server registered at /mcp/v3 endpoint");
}
