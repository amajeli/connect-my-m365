import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { randomUUID } from "crypto";
import type { Express, Request, Response } from "express";
import { storage } from "./storage";
import {
  getValidAccessToken,
  getFilteredEmails,
  getEmailById,
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
  truncate,
  MAX_BODY_CHARS,
  MAX_FILE_CHARS,
} from "./mcp";
import type { User } from "@shared/schema";
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

function formatEmailBriefing(e: any, opts: { metadataOnly: boolean; includeResource: boolean }): any {
  const preview = opts.metadataOnly
    ? (e.bodyPreview || "").slice(0, 100)
    : (e.bodyPreview || "");
  const result: any = {
    id: e.id,
    account: e.accountEmail,
    subject: e.subject,
    from: e.from?.emailAddress
      ? `${e.from.emailAddress.name} <${e.from.emailAddress.address}>`
      : null,
    date: e.receivedDateTime,
    isRead: e.isRead,
    hasAttachments: e.hasAttachments,
    preview,
  };
  if (opts.includeResource) {
    result.resource = emailUri(e.accountEmail, e.id);
  }
  return result;
}

function emailUri(accountEmail: string, messageId: string): string {
  return `outlook-email://${encodeURIComponent(accountEmail)}/${encodeURIComponent(messageId)}`;
}

function leanEmail(e: any) {
  return {
    id: e.id,
    account: e.accountEmail,
    subject: e.subject,
    from: e.from?.emailAddress
      ? `${e.from.emailAddress.name} <${e.from.emailAddress.address}>`
      : null,
    date: e.receivedDateTime,
    preview: e.bodyPreview,
    isRead: e.isRead,
    hasAttachments: e.hasAttachments,
    resource: emailUri(e.accountEmail, e.id),
  };
}

function emailResourceLink(e: any) {
  return {
    type: "resource_link" as const,
    uri: emailUri(e.accountEmail, e.id),
    name: e.subject || "(no subject)",
    description: `From ${e.from?.emailAddress?.address || "unknown"} · ${e.receivedDateTime}`,
    mimeType: "text/plain",
  };
}

function createMcpV2Server(user: User): McpServer {
  const server = new McpServer({
    name: "microsoft365-connector-v2",
    version: "2.0.0",
  });

  server.tool(
    "list_accounts",
    "List all connected Microsoft Outlook accounts for this user. Returns each account's email, display name, and type (work or personal).",
    {},
    async () => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      if (accounts.length === 0) {
        return { content: [{ type: "text" as const, text: "No Microsoft accounts are connected." }] };
      }
      const formatted = accounts.map((a) => ({
        email: a.email,
        displayName: a.displayName,
        type: a.accountType,
      }));
      return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
    }
  );

  server.tool(
    "get_recent_emails",
    "Fetch recent emails as lean metadata for briefing — subject, sender, date, isRead, hasAttachments, and a short preview. Defaults to the last 18 hours, limit 12, metadata-only mode, no resource URIs. Call get_email by ID for full body. Automated/transactional senders are filtered out automatically. Honors Focused-inbox-only and approved-sender connector settings.",
    {
      account_email: z.string().optional().describe("Filter to a specific connected account email. If omitted, fetches from all accounts."),
      after_datetime: z.string().optional().describe("Only emails received on/after this time. Defaults to 18 hours ago. Accepts 'today', 'yesterday', an ISO date (YYYY-MM-DD), or full ISO datetime."),
      before_datetime: z.string().optional().describe("Only emails received before this time. Accepts 'today', an ISO date (YYYY-MM-DD), or full ISO datetime."),
      sender: z.string().optional().describe("Only emails from this exact sender email address."),
      limit: z.number().min(1).max(25).default(12).describe("Maximum number of emails to return (default: 12, max: 25)."),
      metadata_only: z.boolean().default(true).describe("When true (default), returns slim fields only: subject, from, date, isRead, hasAttachments, preview capped at 100 chars. Set false for full preview text."),
      include_resource_uris: z.boolean().default(false).describe("When true, includes outlook-email:// resource URIs in the response. Off by default to save tokens."),
      skip_newsletters: z.boolean().default(true).describe("When true (default), filters out newsletters, digests, roundups, and marketing emails. Set false to include them."),
    },
    async ({ account_email, after_datetime, before_datetime, sender, limit, metadata_only, include_resource_uris, skip_newsletters }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };

      const settings = await getEffectiveSettings(user.id);
      const isFree = !isFullAccess(user);
      if (isFree) {
        const usedToday = getFreeEmailsUsedToday(user.id);
        if (usedToday >= FREE_EMAIL_DAILY_LIMIT) {
          return {
            content: [{
              type: "text" as const,
              text: `Daily email limit reached. Free tier allows ${FREE_EMAIL_DAILY_LIMIT} emails per day. ${usedToday} fetched today. Resets at midnight UTC. Upgrade for unlimited access.`,
            }],
          };
        }
        limit = Math.min(limit, FREE_EMAIL_DAILY_LIMIT - usedToday);
      }

      const accounts = await storage.getAccountsByUserId(user.id);
      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;
      if (targetAccounts.length === 0) {
        return { content: [{ type: "text" as const, text: account_email ? `No connected account found with email: ${account_email}` : "No Microsoft accounts are connected." }] };
      }

      // Default to 18-hour rolling window if no after_datetime provided
      const effectiveAfterDatetime = after_datetime || new Date(Date.now() - 18 * 60 * 60 * 1000).toISOString();
      const after = parseDateInput(effectiveAfterDatetime, "start");
      const before = parseDateInput(before_datetime, "end");

      const needsClientFilter = settings.focusedOnly || settings.approvedSenders.length > 0 || Boolean(sender);
      const fetchTop = needsClientFilter ? Math.max(limit * 4, 50) : Math.max(limit * 2, 25);

      let all: any[] = [];
      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const emails = await getFilteredEmails(accessToken, {
            top: fetchTop,
            afterDateTime: after,
            beforeDateTime: before,
            sender,
            focusedOnly: settings.focusedOnly,
          });
          for (const e of emails) {
            e.accountEmail = account.email;
            e.accountId = account.id;
          }
          all.push(...emails);
        } catch (err: any) {
          console.error(`MCP v2: Failed to fetch emails for ${account.email}:`, err.message);
        }
      }

      // Apply approved-sender allow-list
      if (settings.approvedSenders.length > 0) {
        all = all.filter((e) => senderMatches(e.from?.emailAddress?.address || "", settings.approvedSenders));
      }

      // Apply sender blocklist (automated/transactional noise)
      all = all.filter((e) => !isBlockedSender(e.from?.emailAddress?.address || ""));

      // Apply newsletter filter
      if (skip_newsletters) {
        all = all.filter((e) => !isNewsletter(e));
      }

      all.sort((a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime());
      const limited = all.slice(0, limit);

      if (isFree && limited.length > 0) incrementFreeEmailCount(user.id, limited.length);

      if (limited.length === 0) {
        return { content: [{ type: "text" as const, text: "No emails found matching the given filters and connector settings." }] };
      }

      const formatted = limited.map((e) => formatEmailBriefing(e, { metadataOnly: metadata_only, includeResource: include_resource_uris }));
      const resourceLinks = include_resource_uris ? limited.map(emailResourceLink) : [];

      return {
        content: [
          { type: "text" as const, text: JSON.stringify(formatted) },
          ...resourceLinks,
        ],
      };
    }
  );

  server.tool(
    "search_emails",
    "Search emails by keyword across subject, sender, and preview, returning lean metadata (no full bodies). Read each result's `resource` URI (or call get_email) for the full body. Search runs over recent emails (most recent 50 per account) unless you narrow with a date range. Honors Focused-inbox-only and approved-sender connector settings.",
    {
      query: z.string().describe("Keyword to match against subject, sender, and preview."),
      account_email: z.string().optional().describe("Filter to a specific connected account email."),
      after_datetime: z.string().optional().describe("Only emails received on/after this time ('today', YYYY-MM-DD, or ISO datetime)."),
      before_datetime: z.string().optional().describe("Only emails received before this time ('today', YYYY-MM-DD, or ISO datetime)."),
      limit: z.number().min(1).max(25).default(5).describe("Maximum number of results to return (default: 5, max: 25)."),
    },
    async ({ query, account_email, after_datetime, before_datetime, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "search_emails");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const settings = await getEffectiveSettings(user.id);
      const accounts = await storage.getAccountsByUserId(user.id);
      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;
      if (targetAccounts.length === 0) {
        return { content: [{ type: "text" as const, text: account_email ? `No connected account found with email: ${account_email}` : "No Microsoft accounts are connected." }] };
      }

      const after = parseDateInput(after_datetime, "start");
      const before = parseDateInput(before_datetime, "end");
      const q = query.toLowerCase();

      let all: any[] = [];
      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const emails = await getFilteredEmails(accessToken, {
            top: Math.max(50, limit),
            afterDateTime: after,
            beforeDateTime: before,
            focusedOnly: settings.focusedOnly,
          });
          for (const e of emails) {
            e.accountEmail = account.email;
            e.accountId = account.id;
          }
          all.push(...emails);
        } catch (err: any) {
          console.error(`MCP v2: Failed to search emails for ${account.email}:`, err.message);
        }
      }

      if (settings.approvedSenders.length > 0) {
        all = all.filter((e) => senderMatches(e.from?.emailAddress?.address || "", settings.approvedSenders));
      }
      all = all.filter((e) => {
        const haystack = `${e.subject || ""} ${e.bodyPreview || ""} ${e.from?.emailAddress?.address || ""} ${e.from?.emailAddress?.name || ""}`.toLowerCase();
        return haystack.includes(q);
      });
      all.sort((a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime());
      const limited = all.slice(0, limit);

      if (limited.length === 0) {
        return { content: [{ type: "text" as const, text: `No emails found matching "${query}".` }] };
      }

      return {
        content: [
          { type: "text" as const, text: JSON.stringify(limited.map(leanEmail)) },
          ...limited.map(emailResourceLink),
        ],
      };
    }
  );

  server.tool(
    "get_email",
    "Load the full content of a single email by its ID (from get_recent_emails / search_emails). Returns the complete body (HTML stripped, truncated to conserve context), recipients, and metadata. Equivalent to reading the email's `resource` URI.",
    {
      account_email: z.string().describe("The email address of the account that contains this email."),
      email_id: z.string().describe("The ID of the email to retrieve."),
    },
    async ({ account_email, email_id }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "get_email");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const email = await getEmailById(accessToken, email_id);
        const settings = await getEffectiveSettings(user.id);
        const denied = emailAllowedBySettings(email, settings);
        if (denied) return { content: [{ type: "text" as const, text: denied }] };
        const formatted = {
          subject: email.subject,
          from: email.from?.emailAddress ? `${email.from.emailAddress.name} <${email.from.emailAddress.address}>` : null,
          to: email.toRecipients?.map((r: any) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", "),
          cc: email.ccRecipients?.map((r: any) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", "),
          date: email.receivedDateTime,
          importance: email.importance,
          hasAttachments: email.hasAttachments,
          webLink: email.webLink,
          body: truncate(email.body?.contentType === "html" ? stripHtml(email.body.content) : (email.body?.content || ""), MAX_BODY_CHARS),
        };
        return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error fetching email: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "create_email_draft",
    "Create a NEW email DRAFT in the account's Drafts folder. This NEVER sends mail — it only saves a draft the user can review and send manually from Outlook. Returns the draft ID and a web link to open it. Requires the Mail.ReadWrite permission.",
    {
      account_email: z.string().describe("The connected account that should own the draft."),
      to: z.array(z.string()).min(1).describe("Recipient email addresses."),
      subject: z.string().describe("The email subject line."),
      body: z.string().describe("The email body content."),
      cc: z.array(z.string()).optional().describe("Optional CC recipient email addresses."),
      body_type: z.enum(["Text", "HTML"]).default("Text").describe("Body format (default: Text)."),
    },
    async ({ account_email, to, subject, body, cc, body_type }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "create_email_draft");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const draft = await createEmailDraft(accessToken, { to, cc, subject, body, bodyType: body_type });
        const result = {
          status: "draft_created",
          note: "Saved to Drafts only — nothing was sent. Open it in Outlook to review and send.",
          id: draft.id,
          webLink: draft.webLink,
          account: account.email,
        };
        return { content: [{ type: "text" as const, text: JSON.stringify(result) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error creating draft: ${err.message}. If this mentions permissions, the account may need to re-consent with the Mail.ReadWrite scope.` }] };
      }
    }
  );

  server.tool(
    "outlook_calendar_search",
    "Search calendar events within a specific time window. Returns compact event metadata (subject, start/end, location, organizer, meeting link). Use this for targeted lookups (e.g. 'today', a specific date range); use get_calendar_events for a general upcoming overview.",
    {
      account_email: z.string().optional().describe("Filter to a specific connected account email."),
      after_datetime: z.string().optional().describe("Window start ('today', YYYY-MM-DD, or ISO datetime). Defaults to now."),
      before_datetime: z.string().optional().describe("Window end ('tomorrow', YYYY-MM-DD, or ISO datetime). Defaults to 30 days from start."),
      limit: z.number().min(1).max(50).default(10).describe("Maximum number of events to return (default: 10, max: 50)."),
    },
    async ({ account_email, after_datetime, before_datetime, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "outlook_calendar_search");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;
      if (targetAccounts.length === 0) {
        return { content: [{ type: "text" as const, text: account_email ? `No connected account found with email: ${account_email}` : "No Microsoft accounts are connected." }] };
      }

      const start = parseDateInput(after_datetime, "start") || new Date().toISOString();
      const end = parseDateInput(before_datetime, "end") || new Date(new Date(start).getTime() + 30 * 86400000).toISOString();

      let all: any[] = [];
      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const response = await fetch(
            `https://graph.microsoft.com/v1.0/me/calendarview?startDateTime=${start}&endDateTime=${end}&$select=id,subject,start,end,location,organizer,isAllDay,showAs,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime&$top=100`,
            { headers: { Authorization: `Bearer ${accessToken}`, Prefer: 'outlook.timezone="UTC"' } }
          );
          if (!response.ok) throw new Error(await response.text());
          const data = await response.json();
          for (const e of data.value || []) {
            e.accountEmail = account.email;
          }
          all.push(...(data.value || []));
        } catch (err: any) {
          console.error(`MCP v2: Failed to search calendar for ${account.email}:`, err.message);
        }
      }

      all.sort((a, b) => new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime());
      const limited = all.slice(0, limit);

      if (limited.length === 0) {
        return { content: [{ type: "text" as const, text: "No calendar events found in the specified window." }] };
      }

      const formatted = limited.map((e) => ({
        subject: e.subject,
        start: e.start.dateTime,
        end: e.end.dateTime,
        isAllDay: e.isAllDay,
        location: e.location?.displayName || null,
        organizer: e.organizer ? `${e.organizer.emailAddress.name} <${e.organizer.emailAddress.address}>` : null,
        showAs: e.showAs,
        meetingUrl: e.onlineMeetingUrl || null,
        account: e.accountEmail,
      }));
      return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
    }
  );

  server.tool(
    "get_calendar_events",
    "Fetch an overview of upcoming and recent calendar events across connected accounts. Returns subject, time, location, organizer, and online meeting links.",
    {
      account_email: z.string().optional().describe("Filter events to a specific account email. If omitted, fetches from all accounts."),
      days_ahead: z.number().min(1).max(90).default(30).describe("Number of days ahead to fetch events (default: 30, max: 90)."),
      days_behind: z.number().min(0).max(30).default(7).describe("Number of days in the past to include (default: 7, max: 30)."),
    },
    async ({ account_email, days_ahead, days_behind }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "get_calendar_events");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;
      if (targetAccounts.length === 0) {
        return { content: [{ type: "text" as const, text: account_email ? `No connected account found with email: ${account_email}` : "No Microsoft accounts are connected." }] };
      }

      let all: any[] = [];
      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const now = new Date();
          const startDate = new Date(now);
          startDate.setDate(startDate.getDate() - days_behind);
          const endDate = new Date(now);
          endDate.setDate(endDate.getDate() + days_ahead);
          const response = await fetch(
            `https://graph.microsoft.com/v1.0/me/calendarview?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$select=id,subject,start,end,location,organizer,isAllDay,showAs,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime&$top=100`,
            { headers: { Authorization: `Bearer ${accessToken}`, Prefer: 'outlook.timezone="UTC"' } }
          );
          if (!response.ok) throw new Error(await response.text());
          const data = await response.json();
          for (const e of data.value || []) {
            e.accountEmail = account.email;
          }
          all.push(...(data.value || []));
        } catch (err: any) {
          console.error(`MCP v2: Failed to fetch calendar for ${account.email}:`, err.message);
        }
      }

      all.sort((a, b) => new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime());
      const formatted = all.map((e) => ({
        subject: e.subject,
        start: e.start.dateTime,
        end: e.end.dateTime,
        isAllDay: e.isAllDay,
        location: e.location?.displayName || null,
        organizer: e.organizer ? `${e.organizer.emailAddress.name} <${e.organizer.emailAddress.address}>` : null,
        showAs: e.showAs,
        meetingUrl: e.onlineMeetingUrl || null,
        account: e.accountEmail,
      }));
      return {
        content: [{
          type: "text" as const,
          text: formatted.length > 0 ? JSON.stringify(formatted) : "No calendar events found in the specified date range.",
        }],
      };
    }
  );

  server.tool(
    "create_calendar_event",
    "Create a new calendar event on the specified Microsoft 365 account. Returns the created event's id, subject, and start/end times.",
    {
      account_email: z.string().describe("The connected account email to create the event on."),
      subject: z.string().describe("Title/subject of the event."),
      start_datetime: z.string().describe("Event start time as ISO 8601 datetime (e.g. '2026-08-10T14:00:00')."),
      end_datetime: z.string().describe("Event end time as ISO 8601 datetime (e.g. '2026-08-10T15:00:00')."),
      timezone: z.string().default("UTC").describe("IANA timezone for start/end times (e.g. 'America/New_York'). Defaults to UTC."),
      body: z.string().optional().describe("Body/description of the event (plain text)."),
      location: z.string().optional().describe("Location of the event."),
      attendees: z.array(z.string()).optional().describe("List of attendee email addresses to invite."),
      is_all_day: z.boolean().default(false).describe("Whether this is an all-day event. If true, start/end should be dates only (YYYY-MM-DD)."),
      is_online_meeting: z.boolean().default(false).describe("Whether to add a Teams online meeting link."),
      reminder_on: z.boolean().default(true).describe("Whether the event reminder is enabled. Defaults to true."),
      reminder_minutes: z.number().int().min(0).max(40320).default(15).describe("Minutes before the event start to show the reminder. Defaults to 15."),
    },
    async ({ account_email, subject, start_datetime, end_datetime, timezone, body, location, attendees, is_all_day, is_online_meeting, reminder_on, reminder_minutes }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "create_calendar_event");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      const eventBody: any = {
        subject,
        start: { dateTime: start_datetime, timeZone: timezone },
        end: { dateTime: end_datetime, timeZone: timezone },
        isAllDay: is_all_day,
        isOnlineMeeting: is_online_meeting,
        isReminderOn: reminder_on,
        reminderMinutesBeforeStart: reminder_minutes,
      };
      if (body) {
        eventBody.body = { contentType: "text", content: body };
      }
      if (location) {
        eventBody.location = { displayName: location };
      }
      if (attendees && attendees.length > 0) {
        eventBody.attendees = attendees.map((email) => ({
          emailAddress: { address: email },
          type: "required",
        }));
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(eventBody),
        });
        if (!response.ok) throw new Error(await response.text());
        const created = await response.json();
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              id: created.id,
              subject: created.subject,
              start: created.start,
              end: created.end,
              webLink: created.webLink,
              onlineMeetingUrl: created.onlineMeetingUrl || null,
            }),
          }],
        };
      } catch (err: any) {
        console.error(`MCP v2: Failed to create calendar event for ${account_email}:`, err.message);
        return { content: [{ type: "text" as const, text: `Failed to create event: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "update_calendar_event",
    "Update an existing calendar event. Only the fields you provide will be changed. Use get_calendar_events or outlook_calendar_search to find the event id first.",
    {
      account_email: z.string().describe("The connected account email that owns the event."),
      event_id: z.string().describe("The event id to update (from get_calendar_events or outlook_calendar_search)."),
      subject: z.string().optional().describe("New title/subject for the event."),
      start_datetime: z.string().optional().describe("New start time as ISO 8601 datetime."),
      end_datetime: z.string().optional().describe("New end time as ISO 8601 datetime."),
      timezone: z.string().optional().describe("IANA timezone for updated start/end times (e.g. 'America/New_York')."),
      body: z.string().optional().describe("New body/description (plain text)."),
      location: z.string().optional().describe("New location."),
      attendees: z.array(z.string()).optional().describe("Replacement list of attendee email addresses (replaces existing attendees)."),
      is_online_meeting: z.boolean().optional().describe("Set to true to add a Teams meeting link, false to remove it."),
      reminder_on: z.boolean().optional().describe("Enable or disable the event reminder."),
      reminder_minutes: z.number().int().min(0).max(40320).optional().describe("Minutes before the event start to show the reminder."),
    },
    async ({ account_email, event_id, subject, start_datetime, end_datetime, timezone, body, location, attendees, is_online_meeting, reminder_on, reminder_minutes }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "update_calendar_event");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      const patch: any = {};
      if (subject !== undefined) patch.subject = subject;
      if (start_datetime !== undefined) patch.start = { dateTime: start_datetime, timeZone: timezone || "UTC" };
      if (end_datetime !== undefined) patch.end = { dateTime: end_datetime, timeZone: timezone || "UTC" };
      if (body !== undefined) patch.body = { contentType: "text", content: body };
      if (location !== undefined) patch.location = { displayName: location };
      if (attendees !== undefined) patch.attendees = attendees.map((email) => ({ emailAddress: { address: email }, type: "required" }));
      if (is_online_meeting !== undefined) patch.isOnlineMeeting = is_online_meeting;
      if (reminder_on !== undefined) patch.isReminderOn = reminder_on;
      if (reminder_minutes !== undefined) patch.reminderMinutesBeforeStart = reminder_minutes;

      if (Object.keys(patch).length === 0) {
        return { content: [{ type: "text" as const, text: "No fields to update were provided." }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(event_id)}`, {
          method: "PATCH",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(patch),
        });
        if (!response.ok) throw new Error(await response.text());
        const updated = await response.json();
        return {
          content: [{
            type: "text" as const,
            text: JSON.stringify({
              id: updated.id,
              subject: updated.subject,
              start: updated.start,
              end: updated.end,
              webLink: updated.webLink,
            }),
          }],
        };
      } catch (err: any) {
        console.error(`MCP v2: Failed to update calendar event for ${account_email}:`, err.message);
        return { content: [{ type: "text" as const, text: `Failed to update event: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "delete_calendar_event",
    "Permanently delete a calendar event. Use get_calendar_events or outlook_calendar_search to find the event id first. This cannot be undone.",
    {
      account_email: z.string().describe("The connected account email that owns the event."),
      event_id: z.string().describe("The event id to delete."),
    },
    async ({ account_email, event_id }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "delete_calendar_event");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(event_id)}`, {
          method: "DELETE",
          headers: { Authorization: `Bearer ${accessToken}` },
        });
        if (!response.ok) throw new Error(await response.text());
        return { content: [{ type: "text" as const, text: `Event deleted successfully.` }] };
      } catch (err: any) {
        console.error(`MCP v2: Failed to delete calendar event for ${account_email}:`, err.message);
        return { content: [{ type: "text" as const, text: `Failed to delete event: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "list_onedrive_files",
    "List files and folders in OneDrive. If the user has designated allowed folders in their connector settings, browsing is restricted to those folders (and their subfolders). Returns names, types, sizes, and item IDs for use with get_onedrive_file_content.",
    {
      account_email: z.string().describe("The email address of the Microsoft account whose OneDrive to browse."),
      folder_path: z.string().optional().describe("Folder path to list (e.g. 'Documents/Projects'). Omit or 'root' for the OneDrive root."),
      limit: z.number().min(1).max(100).default(50).describe("Maximum number of items to return (default: 50, max: 100)."),
    },
    async ({ account_email, folder_path, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "list_onedrive_files");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const settings = await getEffectiveSettings(user.id);
      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      const requested = folder_path && folder_path !== "root" ? folder_path : "";
      if (settings.allowedFolders.length > 0) {
        if (requested === "") {
          return { content: [{ type: "text" as const, text: JSON.stringify({ note: "OneDrive access is restricted to designated folders. Call list_onedrive_files again with one of these folder_path values.", allowedFolders: settings.allowedFolders }) }] };
        }
        if (!isWithinAllowed(requested, settings.allowedFolders)) {
          return { content: [{ type: "text" as const, text: JSON.stringify({ error: "Access denied. This folder is outside the designated OneDrive folders for this connector.", allowedFolders: settings.allowedFolders }) }] };
        }
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const items = await listDriveItems(accessToken, requested || "root", limit);
        if (items.length === 0) {
          return { content: [{ type: "text" as const, text: `No files found in ${requested || "root"}.` }] };
        }
        const formatted = items.map((item: any) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          mimeType: item.file?.mimeType || null,
          size: item.size != null ? item.size : null,
          lastModified: item.lastModifiedDateTime,
          parentPath: item.parentReference?.path?.replace("/drive/root:", "") || "/",
        }));
        return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error listing OneDrive files: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "search_onedrive_files",
    "Search OneDrive files by name or keyword. If the user has designated allowed folders in their connector settings, results are limited to those folders. Returns matching items with IDs, types, sizes, and locations.",
    {
      account_email: z.string().describe("The email address of the Microsoft account whose OneDrive to search."),
      query: z.string().describe("Search keyword to find files by name or content."),
      limit: z.number().min(1).max(50).default(25).describe("Maximum number of results to return (default: 25, max: 50)."),
    },
    async ({ account_email, query, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "search_onedrive_files");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const settings = await getEffectiveSettings(user.id);
      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        let items = await searchDriveItems(accessToken, query, Math.max(limit, settings.allowedFolders.length > 0 ? 50 : limit));
        if (settings.allowedFolders.length > 0) {
          items = items.filter((item: any) => {
            const parent = (item.parentReference?.path || "").replace("/drive/root:", "");
            const full = `${parent}/${item.name}`;
            return isWithinAllowed(full, settings.allowedFolders);
          });
        }
        items = items.slice(0, limit);
        if (items.length === 0) {
          return { content: [{ type: "text" as const, text: `No files found matching "${query}"${settings.allowedFolders.length > 0 ? " within the designated folders" : ""}.` }] };
        }
        const formatted = items.map((item: any) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          mimeType: item.file?.mimeType || null,
          size: item.size != null ? item.size : null,
          lastModified: item.lastModifiedDateTime,
          parentPath: item.parentReference?.path?.replace("/drive/root:", "") || "/",
        }));
        return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error searching OneDrive: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "get_onedrive_file_content",
    `Read a OneDrive file by its ID. For plain-text files under 512 KB, returns the content (truncated to ~${MAX_FILE_CHARS} characters). For other types, returns metadata and a web link. Respects the user's designated OneDrive folders — files outside them are denied.`,
    {
      account_email: z.string().describe("The email address of the Microsoft account that owns the file."),
      item_id: z.string().describe("The ID of the OneDrive file (from list_onedrive_files or search_onedrive_files)."),
    },
    async ({ account_email, item_id }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "get_onedrive_file_content");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const settings = await getEffectiveSettings(user.id);
      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const { metadata, content, mimeType } = await getDriveItemContent(accessToken, item_id);

        if (settings.allowedFolders.length > 0) {
          const parent = (metadata.parentReference?.path || "").replace("/drive/root:", "");
          const full = `${parent}/${metadata.name}`;
          if (!isWithinAllowed(full, settings.allowedFolders)) {
            return { content: [{ type: "text" as const, text: JSON.stringify({ error: "Access denied. This file is outside the designated OneDrive folders for this connector.", allowedFolders: settings.allowedFolders }) }] };
          }
        }

        if (content !== null) {
          const result = {
            name: metadata.name,
            mimeType,
            size: metadata.size,
            lastModified: metadata.lastModifiedDateTime,
            webUrl: metadata.webUrl,
            content: truncate(content, MAX_FILE_CHARS),
          };
          return { content: [{ type: "text" as const, text: JSON.stringify(result) }] };
        }

        const result = {
          name: metadata.name,
          mimeType,
          size: metadata.size,
          lastModified: metadata.lastModifiedDateTime,
          webUrl: metadata.webUrl,
          note: metadata.folder
            ? "This is a folder, not a file. Use list_onedrive_files to browse its contents."
            : `This file type (${mimeType}) cannot be read as text. Open it at the web URL above.`,
        };
        return { content: [{ type: "text" as const, text: JSON.stringify(result) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error fetching OneDrive file: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "list_teams_chats",
    "List recent Microsoft Teams chats for a connected work account. Returns each chat's ID, type, topic, and participants. Only available for work (Azure AD) accounts.",
    {
      account_email: z.string().describe("The work account whose Teams chats to list."),
      limit: z.number().min(1).max(50).default(20).describe("Maximum number of chats to return (default: 20, max: 50)."),
    },
    async ({ account_email, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "list_teams_chats");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const chats = await getTeamsChats(accessToken, limit);
        if (chats.length === 0) {
          return { content: [{ type: "text" as const, text: "No Teams chats found. Teams chats are only available for work (Azure AD) accounts." }] };
        }
        const formatted = chats.map((chat: any) => ({
          id: chat.id,
          type: chat.chatType,
          topic: chat.topic || null,
          lastUpdated: chat.lastUpdatedDateTime,
          members: (chat.members || []).map((m: any) => ({ displayName: m.displayName, email: m.email || null })),
        }));
        return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error listing Teams chats: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "get_teams_chat_messages",
    "Read messages from a specific Microsoft Teams chat. Returns message content (HTML stripped, truncated), sender, and timestamp. Use list_teams_chats first to get the chat ID. Work accounts only.",
    {
      account_email: z.string().describe("The work Microsoft account."),
      chat_id: z.string().describe("The ID of the Teams chat (from list_teams_chats)."),
      limit: z.number().min(1).max(50).default(25).describe("Maximum number of messages, newest first (default: 25, max: 50)."),
    },
    async ({ account_email, chat_id, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "get_teams_chat_messages");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const messages = await getTeamsChatMessages(accessToken, chat_id, limit);
        if (messages.length === 0) {
          return { content: [{ type: "text" as const, text: "No messages found in this chat." }] };
        }
        const formatted = messages
          .filter((m: any) => m.messageType === "message")
          .map((m: any) => ({
            id: m.id,
            sender: m.from?.user ? { name: m.from.user.displayName, email: m.from.user.userPrincipalName } : { name: "Unknown", email: null },
            sentAt: m.createdDateTime,
            body: truncate(m.body?.contentType === "html" ? stripHtml(m.body.content) : (m.body?.content || ""), MAX_BODY_CHARS),
            mentions: (m.mentions || []).map((mention: any) => mention.mentioned?.user?.displayName).filter(Boolean),
          }));
        return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error fetching Teams messages: ${err.message}` }] };
      }
    }
  );

  server.tool(
    "search_teams_messages",
    "Search across Microsoft Teams chats by keyword. Returns matching messages with sender, timestamp, and a snippet. Work accounts only.",
    {
      account_email: z.string().describe("The work Microsoft account whose Teams to search."),
      query: z.string().describe("Search keyword to find in Teams messages."),
      limit: z.number().min(1).max(25).default(15).describe("Maximum number of results to return (default: 15, max: 25)."),
    },
    async ({ account_email, query, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "search_teams_messages");
      if (fullErr) return { content: [{ type: "text" as const, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());
      if (!account) {
        return { content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const hits = await searchTeamsMessages(accessToken, query, limit);
        if (hits.length === 0) {
          return { content: [{ type: "text" as const, text: `No Teams messages found matching "${query}".` }] };
        }
        const formatted = hits.map((hit: any) => {
          const resource = hit.resource || {};
          return {
            chatId: resource.chatId || null,
            messageId: resource.id || null,
            sender: resource.from?.user ? { name: resource.from.user.displayName, email: resource.from.user.userPrincipalName } : { name: "Unknown", email: null },
            sentAt: resource.createdDateTime,
            summary: stripHtml(hit.summary || resource.body?.content || ""),
          };
        });
        return { content: [{ type: "text" as const, text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { content: [{ type: "text" as const, text: `Error searching Teams messages: ${err.message}` }] };
      }
    }
  );

  server.registerResource(
    "outlook-email",
    new ResourceTemplate("outlook-email://{account}/{messageId}", { list: undefined }),
    {
      title: "Outlook Email",
      description: "Full content of an Outlook email, loaded on demand by URI.",
      mimeType: "text/plain",
    },
    async (uri, variables) => {
      const account = Array.isArray(variables.account) ? variables.account[0] : variables.account;
      const messageId = Array.isArray(variables.messageId) ? variables.messageId[0] : variables.messageId;

      const subError = checkSubscription(user);
      if (subError) return { contents: [{ uri: uri.href, text: subError }] };
      const fullErr = checkFullAccessRequired(user, "read email resource");
      if (fullErr) return { contents: [{ uri: uri.href, text: fullErr }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const acct = accounts.find((a) => a.email.toLowerCase() === String(account).toLowerCase());
      if (!acct) {
        return { contents: [{ uri: uri.href, text: `No connected account found with email: ${account}` }] };
      }

      try {
        const accessToken = await getValidAccessToken(acct.id);
        const email = await getEmailById(accessToken, String(messageId));
        const settings = await getEffectiveSettings(user.id);
        const denied = emailAllowedBySettings(email, settings);
        if (denied) return { contents: [{ uri: uri.href, text: denied }] };
        const formatted = {
          subject: email.subject,
          from: email.from?.emailAddress ? `${email.from.emailAddress.name} <${email.from.emailAddress.address}>` : null,
          to: email.toRecipients?.map((r: any) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", "),
          cc: email.ccRecipients?.map((r: any) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", "),
          date: email.receivedDateTime,
          importance: email.importance,
          hasAttachments: email.hasAttachments,
          webLink: email.webLink,
          body: truncate(email.body?.contentType === "html" ? stripHtml(email.body.content) : (email.body?.content || ""), MAX_BODY_CHARS),
        };
        return { contents: [{ uri: uri.href, mimeType: "text/plain", text: JSON.stringify(formatted) }] };
      } catch (err: any) {
        return { contents: [{ uri: uri.href, text: `Error fetching email: ${err.message}` }] };
      }
    }
  );

  return server;
}

export function registerMcpV2Routes(app: Express): void {
  const sessions: Record<string, { transport: StreamableHTTPServerTransport; server: McpServer; userId: string }> = {};

  app.post("/mcp/v2", async (req: Request, res: Response) => {
    const user = await validateMcpAuth(req, res);
    if (!user) return;

    const sessionId = req.headers["mcp-session-id"] as string | undefined;

    if (sessionId && sessions[sessionId]) {
      if (sessions[sessionId].userId !== user.id) {
        return res.status(403).json({ jsonrpc: "2.0", error: { code: -32000, message: "Forbidden: Session belongs to another user" }, id: null });
      }
      await sessions[sessionId].transport.handleRequest(req, res, req.body);
      return;
    }

    if (isInitializeRequest(req.body)) {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
      });
      const server = createMcpV2Server(user);
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);

      const sid = transport.sessionId;
      if (sid) {
        sessions[sid] = { transport, server, userId: user.id };
        transport.onclose = () => {
          delete sessions[sid];
        };
      }
      return;
    }

    res.status(400).json({ jsonrpc: "2.0", error: { code: -32000, message: "Bad Request: No valid session ID provided" }, id: null });
  });

  app.get("/mcp/v2", async (req: Request, res: Response) => {
    const user = await validateMcpAuth(req, res);
    if (!user) return;

    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (sessionId && sessions[sessionId]) {
      if (sessions[sessionId].userId !== user.id) {
        return res.status(403).json({ jsonrpc: "2.0", error: { code: -32000, message: "Forbidden: Session belongs to another user" }, id: null });
      }
      await sessions[sessionId].transport.handleRequest(req, res);
    } else {
      res.status(400).json({ jsonrpc: "2.0", error: { code: -32000, message: "Bad Request: No valid session ID provided" }, id: null });
    }
  });

  app.delete("/mcp/v2", async (req: Request, res: Response) => {
    const user = await validateMcpAuth(req, res);
    if (!user) return;

    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (sessionId && sessions[sessionId]) {
      if (sessions[sessionId].userId !== user.id) {
        return res.status(403).json({ jsonrpc: "2.0", error: { code: -32000, message: "Forbidden: Session belongs to another user" }, id: null });
      }
      await sessions[sessionId].transport.handleRequest(req, res);
      await sessions[sessionId].server.close();
      delete sessions[sessionId];
    } else {
      res.status(400).json({ jsonrpc: "2.0", error: { code: -32000, message: "Bad Request: No valid session ID provided" }, id: null });
    }
  });

  console.log("MCP v2 server registered at /mcp/v2 endpoint");
}
