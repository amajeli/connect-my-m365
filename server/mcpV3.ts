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
  sendChatMessage,
  createChat,
} from "./graphV3";
import {
  row,
  shortDateTime,
  eventWhen,
  person,
  humanSize,
  compactText,
  kqlTerm,
  findConflicts,
  type CalendarItem,
} from "./mcpV3Format";
import { toHandles, toHandle, fromHandle, type HandleKind, type HandleTarget } from "./handleStore";

// MCP v3: built on v2's filters and allow-lists, tuned further for tokens.
//  - Few short tools. Tool definitions are re-sent every turn.
//  - Pipe-delimited rows instead of JSON, and short ids (m1, e1, f1, c1) instead of Graph IDs.
//  - mail_read returns only the new text of a message (Graph uniqueBody), not the quoted thread.
//  - Stateless transport and database-backed short ids, so it runs on serverless hosts.
//  - All connected accounts are queried in parallel.
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

function sameEmail(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

/** Accounts a read tool should cover: the named one, or all of them. */
async function targetAccounts(user: User, account?: string): Promise<ConnectedAccount[] | string> {
  const all = await storage.getAccountsByUserId(user.id);
  if (all.length === 0) return "No Microsoft accounts connected. Connect one in the web app.";
  if (!account) return all;
  const match = all.filter((a) => sameEmail(a.email, account));
  return match.length ? match : `Unknown account ${account}. Connected: ${accountList(all)}`;
}

/** The single account a write tool acts on. */
async function oneAccount(user: User, account?: string): Promise<ConnectedAccount | string> {
  const all = await storage.getAccountsByUserId(user.id);
  if (all.length === 0) return "No Microsoft accounts connected. Connect one in the web app.";
  if (account) {
    const match = all.find((a) => sameEmail(a.email, account));
    return match ?? `Unknown account ${account}. Connected: ${accountList(all)}`;
  }
  if (all.length === 1) return all[0];
  return `Several accounts connected; ask the user which one to use (${accountList(all)}).`;
}

/** Resolves a short id like "m3", or a raw Graph ID plus account. */
async function resolveRef(user: User, kind: HandleKind, id: string, account?: string): Promise<HandleTarget | string> {
  const target = await fromHandle(user.id, kind, id);
  if (target) return target;
  if (new RegExp(`^${kind}\\d+$`).test(id)) return `Unknown id ${id}. Search again to get a current id.`;
  const acct = await oneAccount(user, account);
  if (typeof acct === "string") return acct;
  return { accountId: acct.id, accountEmail: acct.email, ref: id };
}

function errMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

/** Runs fn for every account at once. Failures become "! account: reason" lines. */
async function perAccount<T>(
  accounts: ConnectedAccount[],
  fn: (acct: ConnectedAccount, token: string) => Promise<T>,
): Promise<{ results: { acct: ConnectedAccount; value: T }[]; errors: string[] }> {
  const settled = await Promise.all(
    accounts.map(async (acct) => {
      try {
        return { acct, value: await fn(acct, await getValidAccessToken(acct.id)) };
      } catch (err) {
        return { acct, error: `! ${acct.email}: ${errMessage(err)}` };
      }
    }),
  );
  return {
    results: settled.filter((s): s is { acct: ConnectedAccount; value: T } => !("error" in s)),
    errors: settled.filter((s) => "error" in s).map((s) => (s as { error: string }).error),
  };
}

function target(acct: ConnectedAccount, ref: string): HandleTarget {
  return { accountId: acct.id, accountEmail: acct.email, ref };
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
    "Ids like m1/e1/f1/c1 are short ids for emails/events/files/chats. Pass them back as-is.",
    "Email flags: U=unread A=attachment !=high importance.",
    "Before mail_draft, always ask the user which account to draft from. Never choose it yourself.",
    "Before teams_send, show the user the exact message and get their OK; it sends immediately.",
  ];
  if (settings.focusedOnly) lines.push("Mail is limited to the Focused inbox.");
  if (settings.approvedSenders.length) lines.push(`Mail is limited to ${settings.approvedSenders.length} approved senders.`);
  if (settings.allowedFolders.length) lines.push(`OneDrive is limited to: ${settings.allowedFolders.join(", ")}.`);
  if (!isFullAccess(user)) lines.push(`Free tier: mail_find only, ${FREE_EMAIL_DAILY_LIMIT} emails/day.`);
  return lines.join("\n");
}

function calendarItem(e: any, key: string): CalendarItem {
  return {
    key,
    subject: e.subject || "(no subject)",
    account: e._acct.email,
    start: e.start.dateTime,
    end: e.end.dateTime,
    isAllDay: e.isAllDay,
    showAs: e.showAs,
    isCancelled: e.isCancelled,
    declined: e.responseStatus?.response === "declined",
  };
}

async function fetchCalendars(accounts: ConnectedAccount[], start: string, end: string) {
  const { results, errors } = await perAccount(accounts, (_acct, token) => getCalendarView(token, start, end, 250));
  const events: any[] = [];
  for (const { acct, value } of results) {
    for (const e of value) {
      e._acct = acct;
      events.push(e);
    }
  }
  events.sort((a, b) => Date.parse(`${a.start.dateTime}Z`) - Date.parse(`${b.start.dateTime}Z`));
  return { events, errors };
}

function createMcpV3Server(user: User, instructions?: string): McpServer {
  const server = new McpServer(
    { name: "microsoft365-connector-v3", version: "3.1.0" },
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

      const { results, errors } = await perAccount(accounts, (_acct, token) => {
        if (q) {
          const kql = [kqlTerm(q)];
          if (from) kql.push(`from:${kqlTerm(from)}`);
          if (afterIso) kql.push(`received>=${afterIso.slice(0, 10)}`);
          if (beforeIso) kql.push(`received<=${beforeIso.slice(0, 10)}`);
          return searchMessages(token, kql.join(" "), fetchTop);
        }
        return listInboxMessages(token, {
          top: fetchTop,
          after: afterIso,
          before: beforeIso,
          from: fromIsAddress ? from : undefined,
          unreadOnly: unread === true,
          focusedOnly: settings.focusedOnly,
        });
      });
      let all: any[] = [];
      for (const { acct, value } of results) {
        for (const m of value) m._acct = acct;
        all.push(...value);
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
        const count = all.length;
        all = all.filter((m) => !isBlockedSender(m.from?.emailAddress?.address || "") && !isNewsletter(m));
        hiddenNoise = count - all.length;
      }

      all.sort((a, b) => Date.parse(b.receivedDateTime) - Date.parse(a.receivedDateTime));
      let shown = all.slice(0, limit);
      if (!full && shown.length) incrementFreeEmailCount(user.id, shown.length);

      const out: string[] = [];
      if (shown.length === 0) {
        out.push("No matching emails.");
      } else {
        const multi = new Set(shown.map((m) => m._acct.id)).size > 1;
        if (multi) shown = [...shown].sort((a, b) => a._acct.email.localeCompare(b._acct.email));
        const ids = await toHandles(user.id, "m", shown.map((m) => target(m._acct, m.id)));
        out.push(preview ? "id|received|from|subject|flags|preview" : "id|received|from|subject|flags");
        let currentAcct = "";
        shown.forEach((m, i) => {
          if (multi && m._acct.email !== currentAcct) {
            currentAcct = m._acct.email;
            out.push(`# ${currentAcct}`);
          }
          const flags = `${m.isRead ? "" : "U"}${m.hasAttachments ? "A" : ""}${m.importance === "high" ? "!" : ""}`;
          const values: unknown[] = [ids[i], shortDateTime(m.receivedDateTime), person(m.from?.emailAddress), m.subject || "(no subject)", flags];
          if (preview) values.push(compactText(m.bodyPreview, preview));
          out.push(row(values));
        });
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
          ids: z.array(z.string()).min(1).max(5).describe("Ids from mail_find, e.g. [\"m1\",\"m4\"]"),
          full: z.boolean().default(false),
          max_chars: z.number().int().min(200).max(20000).default(MAX_BODY_CHARS),
          account: z.string().optional().describe("Only needed for raw Graph ids"),
        },
        annotations: READ,
      },
      async ({ ids, full: wantFull, max_chars, account }) => {
        const settings = await getEffectiveSettings(user.id);
        const parts = await Promise.all(
          ids.map(async (id) => {
            const t = await resolveRef(user, "m", id, account);
            if (typeof t === "string") return `## ${id}\n${t}`;
            try {
              const m = await getMessageText(await getValidAccessToken(t.accountId), t.ref);
              const denied = emailAllowedBySettings(m, settings);
              if (denied) return `## ${id}\n${denied}`;
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
                `in: ${t.accountEmail}`,
                m.hasAttachments ? "attachments" : "",
              ].filter(Boolean);
              return `## ${id} · ${m.subject || "(no subject)"}\n${meta.join(" | ")}\n${compactText(text, max_chars)}`;
            } catch (err) {
              return `## ${id}\nError: ${errMessage(err)}`;
            }
          }),
        );
        return ok(parts.join("\n\n"));
      },
    );

    server.registerTool(
      "mail_draft",
      {
        description:
          "Save an email draft; this never sends. ALWAYS ask the user which account to draft from before calling. Never pick one yourself. For a reply, pass reply_to (an m-id); it can be drafted from any account. For a new email, pass to and subject. The user reviews and sends it from Outlook.",
        inputSchema: {
          account: z.string().describe("The account the user chose to draft from"),
          body: z.string().describe("Plain text"),
          reply_to: z.string().optional(),
          reply_all: z.boolean().default(false),
          to: z.array(z.string()).optional(),
          cc: z.array(z.string()).optional(),
          subject: z.string().optional(),
        },
        annotations: WRITE,
      },
      async ({ account, body, reply_to, reply_all, to, cc, subject }) => {
        const acct = await oneAccount(user, account);
        if (typeof acct === "string") return fail(acct);
        try {
          const token = await getValidAccessToken(acct.id);
          if (reply_to) {
            const original = await resolveRef(user, "m", reply_to, account);
            if (typeof original === "string") return fail(original);

            if (original.accountId === acct.id) {
              const draft = await createReplyDraft(token, original.ref, body, reply_all);
              return ok(`Reply draft saved in ${acct.email}, not sent: "${draft.subject}"\n${draft.webLink || ""}`.trim());
            }

            // Replying from a different account than the one that received it:
            // a new draft addressed to the sender, with the original quoted below.
            const m = await getMessageText(await getValidAccessToken(original.accountId), original.ref);
            const denied = emailAllowedBySettings(m, await getEffectiveSettings(user.id));
            if (denied) return fail(denied);
            const sender = m.from?.emailAddress?.address;
            if (!sender) return fail("The original email has no sender address to reply to.");
            const others = reply_all
              ? [...(m.toRecipients || []), ...(m.ccRecipients || [])]
                  .map((r: any) => r.emailAddress?.address)
                  .filter((a: string | undefined): a is string => !!a && !sameEmail(a, acct.email) && !sameEmail(a, original.accountEmail) && !sameEmail(a, sender))
              : [];
            const quoted = compactText(m.body?.content || "", MAX_BODY_CHARS);
            const replySubject = /^re:/i.test(m.subject || "") ? m.subject : `RE: ${m.subject || ""}`;
            const draft = await createEmailDraft(token, {
              to: [sender],
              cc: others.length ? others : undefined,
              subject: replySubject,
              body: `${body}\n\n-----Original message-----\nFrom: ${person(m.from?.emailAddress)}\nSent: ${m.receivedDateTime}\nSubject: ${m.subject || ""}\n\n${quoted}`,
              bodyType: "Text",
            });
            return ok(`Reply draft saved in ${acct.email} (original arrived in ${original.accountEmail}), not sent: "${replySubject}"\n${draft.webLink || ""}`.trim());
          }

          if (!to?.length || !subject) return fail("A new draft needs to and subject (or pass reply_to).");
          const draft = await createEmailDraft(token, { to, cc, subject, body, bodyType: "Text" });
          return ok(`Draft saved in ${acct.email}, not sent: "${subject}"\n${draft.webLink || ""}`.trim());
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
          "List calendar events across all accounts. Defaults to today through the next 7 days. Flags: C=double-booked O=online T=tentative F=free W=working elsewhere OOF=out of office X=cancelled D=declined.",
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

        const { events, errors } = await fetchCalendars(accounts, start, end);
        // Conflicts are judged against the whole window, before the q filter.
        const { conflicts } = findConflicts(events.map((e, i) => calendarItem(e, String(i))));
        const conflicted = new Set(conflicts.flatMap((c) => [c.a.key, c.b.key]));

        const needle = q?.toLowerCase();
        const matching = events
          .map((e, i) => ({ e, key: String(i) }))
          .filter(({ e }) =>
            !needle ||
            `${e.subject || ""} ${e.location?.displayName || ""} ${e.organizer?.emailAddress?.name || ""} ${e.organizer?.emailAddress?.address || ""}`
              .toLowerCase()
              .includes(needle),
          );
        const shown = matching.slice(0, limit);

        const out: string[] = [];
        if (shown.length === 0) {
          out.push("No events in that window.");
        } else {
          const multi = accounts.length > 1;
          const ids = await toHandles(user.id, "e", shown.map(({ e }) => target(e._acct, e.id)));
          out.push(`id|when|subject|location|organizer|flags${multi ? "|account" : ""}${links ? "|join" : ""}`);
          const showAs: Record<string, string> = { free: "F", tentative: "T", oof: "OOF", workingElsewhere: "W" };
          shown.forEach(({ e, key }, i) => {
            const flags = [
              conflicted.has(key) ? "C" : "",
              e.isOnlineMeeting ? "O" : "",
              showAs[e.showAs] || "",
              e.isCancelled ? "X" : "",
              e.responseStatus?.response === "declined" ? "D" : "",
            ].join("");
            const orgAddr = e.organizer?.emailAddress?.address || "";
            const values: unknown[] = [
              ids[i],
              eventWhen(e.start.dateTime, e.end.dateTime, e.isAllDay),
              e.subject || "(no subject)",
              e.location?.displayName || "",
              sameEmail(orgAddr, e._acct.email) ? "me" : person(e.organizer?.emailAddress),
              flags,
            ];
            if (multi) values.push(e._acct.email);
            if (links) values.push(e.onlineMeeting?.joinUrl || e.onlineMeetingUrl || "");
            out.push(row(values));
          });
        }
        if (matching.length > shown.length) out.push(`(${matching.length - shown.length} more; raise limit or narrow the window)`);
        if (conflicted.size) out.push(`(${conflicts.length} double-booking${conflicts.length > 1 ? "s" : ""} in this window; calendar_conflicts lists them)`);
        out.push(...errors);
        return ok(out.join("\n"));
      },
    );

    server.registerTool(
      "calendar_conflicts",
      {
        description:
          "Find double bookings across all calendars (defaults to today through the next 14 days). Ignores free, cancelled and declined events, and the same meeting showing up on two calendars.",
        inputSchema: {
          after: z.string().optional().describe("'today', YYYY-MM-DD or ISO datetime"),
          before: z.string().optional(),
        },
        annotations: READ,
      },
      async ({ after, before }) => {
        const accounts = await targetAccounts(user);
        if (typeof accounts === "string") return fail(accounts);
        const start = parseDateInput(after ?? "today", "start") || new Date().toISOString();
        const end = parseDateInput(before, "end") || new Date(Date.parse(start) + 14 * DAY_MS).toISOString();

        const { events, errors } = await fetchCalendars(accounts, start, end);
        const { conflicts, duplicates } = findConflicts(events.map((e, i) => calendarItem(e, String(i))));

        const out: string[] = [];
        if (!conflicts.length) {
          out.push(`No double bookings across ${accounts.length} calendar${accounts.length > 1 ? "s" : ""}.`);
        } else {
          const involved = [...new Set(conflicts.flatMap((c) => [Number(c.a.key), Number(c.b.key)]))];
          const ids = await toHandles(user.id, "e", involved.map((i) => target(events[i]._acct, events[i].id)));
          const idOf = new Map(involved.map((i, n) => [String(i), ids[n]]));
          const multi = accounts.length > 1;
          const label = (c: CalendarItem) => `${idOf.get(c.key)} ${c.subject}${multi ? ` (${c.account})` : ""}`;
          out.push("overlap|event|clashes with|scope");
          for (const c of conflicts) {
            out.push(row([eventWhen(c.overlapStart, c.overlapEnd, false), label(c.a), label(c.b), c.crossCalendar ? "across calendars" : "same calendar"]));
          }
        }
        if (duplicates) out.push(`(${duplicates} meeting${duplicates > 1 ? "s" : ""} on more than one calendar ignored as duplicates)`);
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
          id: z.string().optional().describe("e-id; if omitted, creates a new event"),
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
          let where: HandleTarget;
          if (id) {
            const resolved = await resolveRef(user, "e", id, account);
            if (typeof resolved === "string") return fail(resolved);
            where = resolved;
            if (Object.keys(fields).length === 0) return fail("Nothing to update.");
            saved = await updateEvent(await getValidAccessToken(where.accountId), where.ref, fields);
          } else {
            if (!subject || !start || !end) return fail("A new event needs subject, start and end.");
            const acct = await oneAccount(user, account);
            if (typeof acct === "string") return fail(acct);
            where = target(acct, "");
            saved = await createEvent(await getValidAccessToken(acct.id), fields);
          }
          const handle = await toHandle(user.id, "e", { ...where, ref: saved.id });
          const invited = attendees?.length ? ` · invitations sent to ${attendees.length}` : "";
          return ok(
            `${id ? "Updated" : "Created"} in ${where.accountEmail}: ${row([handle, eventWhen(saved.start.dateTime, saved.end.dateTime, saved.isAllDay), saved.subject])}${invited}`,
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
        inputSchema: { id: z.string().describe("e-id"), account: z.string().optional() },
        annotations: DESTROY,
      },
      async ({ id, account }) => {
        const t = await resolveRef(user, "e", id, account);
        if (typeof t === "string") return fail(t);
        try {
          await deleteEvent(await getValidAccessToken(t.accountId), t.ref);
          return ok(`Deleted ${id} from ${t.accountEmail}.`);
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

        const parentOf = (item: any) => (item.parentReference?.path || "").replace(/^\/drive\/root:/, "") || "/";
        const { results, errors } = await perAccount(accounts, async (_acct, token) => {
          let items: any[] = q
            ? await searchDriveItems(token, q, allowedFolders.length ? 50 : limit)
            : await listDriveItems(token, folder || "root", limit);
          if (allowedFolders.length) {
            items = items.filter((item) => isWithinAllowed(`${parentOf(item)}/${item.name}`, allowedFolders));
          }
          return items.slice(0, limit);
        });

        const out: string[] = [];
        for (const { acct, value: items } of results) {
          if (!items.length) continue;
          if (accounts.length > 1) out.push(`# ${acct.email}`);
          const ids = await toHandles(user.id, "f", items.map((item) => target(acct, item.id)));
          items.forEach((item, i) => {
            out.push(
              row([
                ids[i],
                item.name,
                item.folder ? "dir" : "file",
                item.folder ? `${item.folder.childCount ?? ""} items` : humanSize(item.size),
                shortDateTime(item.lastModifiedDateTime),
                parentOf(item),
              ]),
            );
          });
        }
        const header = out.length ? ["id|name|kind|size|modified|folder"] : [q ? `No files match "${q}".` : "Folder is empty."];
        return ok([...header, ...out, ...errors].join("\n"));
      },
    );

    server.registerTool(
      "file_read",
      {
        description: "Read a OneDrive text file (txt, md, csv, json and similar). For other types, returns a web link to open it.",
        inputSchema: {
          id: z.string().describe("f-id from files_find"),
          max_chars: z.number().int().min(200).max(50000).default(MAX_FILE_CHARS),
          account: z.string().optional().describe("Only needed for raw Graph ids"),
        },
        annotations: READ,
      },
      async ({ id, max_chars, account }) => {
        const t = await resolveRef(user, "f", id, account);
        if (typeof t === "string") return fail(t);
        const { allowedFolders } = await getEffectiveSettings(user.id);
        try {
          const { metadata, content, mimeType } = await getDriveItemContent(await getValidAccessToken(t.accountId), t.ref);
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

    // ─── Teams chats (work accounts only) ────────────────────────────────────

    const workAccounts = async (account?: string): Promise<ConnectedAccount[] | string> => {
      const accounts = await targetAccounts(user, account);
      if (typeof accounts === "string") return accounts;
      const work = accounts.filter((a) => a.accountType !== "personal");
      return work.length ? work : "Teams needs a work account; none is connected.";
    };

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
        const work = await workAccounts(account);
        if (typeof work === "string") return fail(work);

        const { results, errors } = await perAccount(work, (_acct, token) =>
          q ? searchTeamsMessages(token, q, limit) : getTeamsChats(token, limit),
        );
        const out: string[] = [];
        for (const { acct, value } of results) {
          const items = q ? value.filter((hit: any) => hit.resource?.chatId) : value;
          if (!items.length) continue;
          if (work.length > 1) out.push(`# ${acct.email}`);
          const ids = await toHandles(user.id, "c", items.map((x: any) => target(acct, q ? x.resource.chatId : x.id)));
          items.forEach((x: any, i: number) => {
            if (q) {
              const r = x.resource;
              out.push(row([ids[i], shortDateTime(r.createdDateTime), r.from?.user?.displayName || "", compactText(stripHtml(x.summary || r.body?.content || ""), 200)]));
            } else {
              const others = (x.members || [])
                .filter((m: any) => !sameEmail(m.email || "", acct.email))
                .map((m: any) => m.displayName)
                .filter(Boolean);
              const who = others.length > 4 ? `${others.slice(0, 4).join(", ")} +${others.length - 4}` : others.join(", ");
              out.push(row([ids[i], x.chatType, x.topic || who, shortDateTime(x.lastUpdatedDateTime)]));
            }
          });
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
          id: z.string().describe("c-id from teams_find"),
          limit: z.number().int().min(1).max(50).default(20),
          account: z.string().optional().describe("Only needed for raw Graph ids"),
        },
        annotations: READ,
      },
      async ({ id, limit, account }) => {
        const t = await resolveRef(user, "c", id, account);
        if (typeof t === "string") return fail(t);
        try {
          const messages = await getTeamsChatMessages(await getValidAccessToken(t.accountId), t.ref, limit);
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

    server.registerTool(
      "teams_send",
      {
        description:
          "Send a Teams chat message as the user. It sends immediately, so show the user the exact text and get their OK first. Pass chat (a c-id) to post in an existing chat, or people (emails in the same organization) to start or reuse a chat with them.",
        inputSchema: {
          text: z.string().min(1),
          chat: z.string().optional().describe("c-id from teams_find"),
          people: z.array(z.string()).optional().describe("Emails; one person = 1:1 chat"),
          topic: z.string().optional().describe("Name for a new group chat"),
          account: z.string().optional().describe("Work account to send from when starting a chat"),
        },
        annotations: WRITE,
      },
      async ({ text, chat, people, topic, account }) => {
        if (!chat && !people?.length) return fail("Pass chat (a c-id) or people.");
        try {
          let where: HandleTarget;
          if (chat) {
            const t = await resolveRef(user, "c", chat, account);
            if (typeof t === "string") return fail(t);
            where = t;
          } else {
            const work = await workAccounts(account);
            if (typeof work === "string") return fail(work);
            if (work.length > 1 && !account) {
              return fail(`Several work accounts connected; ask the user which one to send from (${accountList(work)}).`);
            }
            const created = await createChat(await getValidAccessToken(work[0].id), people!, topic);
            where = target(work[0], created.id);
          }
          const sent = await sendChatMessage(await getValidAccessToken(where.accountId), where.ref, text);
          const handle = await toHandle(user.id, "c", where);
          return ok(`Sent in ${handle} from ${where.accountEmail} at ${shortDateTime(sent.createdDateTime)}.`);
        } catch (err) {
          return fail(`Teams send failed: ${errMessage(err)}. If this is a permissions error, reconnect the account to grant Chat.ReadWrite.`);
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

    // A fresh server per request: nothing is kept in memory between calls.
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
