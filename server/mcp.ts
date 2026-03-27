import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { randomUUID } from "crypto";
import type { Express, Request, Response } from "express";
import { storage } from "./storage";
import {
  getValidAccessToken,
  getEmails,
  getCalendarEvents,
  listDriveItems,
  searchDriveItems,
  getDriveItemContent,
} from "./microsoft";
import type { User, OutlookEmail, OutlookCalendarEvent } from "@shared/schema";

function checkSubscription(user: User): string | null {
  const activeStatuses = ["active", "trialing", "free"];
  if (!activeStatuses.includes(user.subscriptionStatus || "")) {
    return "Active subscription required to use Outlook Connector tools. Please visit the Outlook Connector web app to subscribe first.";
  }
  return null;
}

function createMcpServer(user: User): McpServer {
  const server = new McpServer({
    name: "outlook-connector",
    version: "1.0.0",
  });

  server.tool(
    "list_accounts",
    "List all connected Microsoft Outlook accounts. Returns account email, display name, and type (work or personal).",
    {},
    async () => {
      const subError = checkSubscription(user);
      if (subError) {
        return { content: [{ type: "text" as const, text: subError }] };
      }
      const accounts = await storage.getAccountsByUserId(user.id);
      if (accounts.length === 0) {
        return {
          content: [
            {
              type: "text" as const,
              text: "No Microsoft accounts are connected. Please visit the Outlook Connector web app to connect your accounts first.",
            },
          ],
        };
      }

      const accountList = accounts.map((a) => ({
        email: a.email,
        displayName: a.displayName,
        accountType: a.accountType,
        isActive: a.isActive,
      }));

      return {
        content: [
          {
            type: "text" as const,
            text: JSON.stringify(accountList, null, 2),
          },
        ],
      };
    }
  );

  server.tool(
    "get_recent_emails",
    "Fetch recent emails from connected Outlook accounts. Returns subject, sender, date, and body preview for each email. Optionally filter by a specific account email or limit the number of results.",
    {
      account_email: z
        .string()
        .optional()
        .describe(
          "Filter emails to a specific account email address. If omitted, fetches from all accounts."
        ),
      limit: z
        .number()
        .min(1)
        .max(50)
        .default(20)
        .describe("Maximum number of emails to return (default: 20, max: 50)"),
    },
    async ({ account_email, limit }) => {
      const subError = checkSubscription(user);
      if (subError) {
        return { content: [{ type: "text" as const, text: subError }] };
      }
      const accounts = await storage.getAccountsByUserId(user.id);
      if (accounts.length === 0) {
        return {
          content: [{ type: "text" as const, text: "No Microsoft accounts are connected." }],
        };
      }

      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;

      if (targetAccounts.length === 0) {
        return {
          content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }],
        };
      }

      const allEmails: OutlookEmail[] = [];

      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const emails = await getEmails(accessToken, limit);
          const enriched = emails.map((email: any) => ({
            ...email,
            accountId: account.id,
            accountEmail: account.email,
          }));
          allEmails.push(...enriched);
        } catch (err: any) {
          console.error(`MCP: Failed to fetch emails for ${account.email}:`, err.message);
        }
      }

      allEmails.sort(
        (a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
      );

      const limited = allEmails.slice(0, limit);

      const formatted = limited.map((e) => ({
        subject: e.subject,
        from: `${e.from.emailAddress.name} <${e.from.emailAddress.address}>`,
        to: e.toRecipients.map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", "),
        date: e.receivedDateTime,
        preview: e.bodyPreview,
        isRead: e.isRead,
        importance: e.importance,
        hasAttachments: e.hasAttachments,
        account: e.accountEmail,
      }));

      return {
        content: [{ type: "text" as const, text: JSON.stringify(formatted, null, 2) }],
      };
    }
  );

  server.tool(
    "search_emails",
    "Search emails across connected Outlook accounts by keyword. Searches in subject lines and body previews. Returns matching emails with subject, sender, date, and preview.",
    {
      query: z.string().describe("Search keyword to find in email subjects and body previews"),
      account_email: z.string().optional().describe("Filter search to a specific account email. If omitted, searches all accounts."),
      limit: z.number().min(1).max(50).default(20).describe("Maximum number of results to return (default: 20, max: 50)"),
    },
    async ({ query, account_email, limit }) => {
      const subError = checkSubscription(user);
      if (subError) {
        return { content: [{ type: "text" as const, text: subError }] };
      }
      const accounts = await storage.getAccountsByUserId(user.id);
      if (accounts.length === 0) {
        return {
          content: [{ type: "text" as const, text: "No Microsoft accounts are connected." }],
        };
      }

      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;

      const allEmails: OutlookEmail[] = [];

      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const emails = await getEmails(accessToken, 50);
          const enriched = emails.map((email: any) => ({
            ...email,
            accountId: account.id,
            accountEmail: account.email,
          }));
          allEmails.push(...enriched);
        } catch (err: any) {
          console.error(`MCP: Failed to fetch emails for ${account.email}:`, err.message);
        }
      }

      const queryLower = query.toLowerCase();
      const filtered = allEmails.filter(
        (e) =>
          e.subject.toLowerCase().includes(queryLower) ||
          e.bodyPreview.toLowerCase().includes(queryLower) ||
          e.from.emailAddress.name.toLowerCase().includes(queryLower) ||
          e.from.emailAddress.address.toLowerCase().includes(queryLower)
      );

      filtered.sort(
        (a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
      );

      const limited = filtered.slice(0, limit);

      const formatted = limited.map((e) => ({
        subject: e.subject,
        from: `${e.from.emailAddress.name} <${e.from.emailAddress.address}>`,
        to: e.toRecipients.map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", "),
        date: e.receivedDateTime,
        preview: e.bodyPreview,
        isRead: e.isRead,
        importance: e.importance,
        hasAttachments: e.hasAttachments,
        account: e.accountEmail,
      }));

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? JSON.stringify(formatted, null, 2)
              : `No emails found matching "${query}"`,
          },
        ],
      };
    }
  );

  server.tool(
    "get_email_detail",
    "Get the full content of a specific email by its ID. Returns the complete email body, all recipients, and metadata.",
    {
      email_id: z.string().describe("The ID of the email to retrieve"),
      account_email: z.string().describe("The email address of the account that contains this email"),
    },
    async ({ email_id, account_email }) => {
      const subError = checkSubscription(user);
      if (subError) {
        return { content: [{ type: "text" as const, text: subError }] };
      }
      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find(
        (a) => a.email.toLowerCase() === account_email.toLowerCase()
      );

      if (!account) {
        return {
          content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }],
        };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const response = await fetch(
          `https://graph.microsoft.com/v1.0/me/messages/${email_id}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,hasAttachments,importance`,
          {
            headers: { Authorization: `Bearer ${accessToken}` },
          }
        );

        if (!response.ok) {
          throw new Error(`Failed to fetch email: ${response.statusText}`);
        }

        const email = await response.json();

        const stripHtml = (html: string) =>
          html.replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim();

        const formatted = {
          subject: email.subject,
          from: `${email.from.emailAddress.name} <${email.from.emailAddress.address}>`,
          to: email.toRecipients
            ?.map((r: any) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
            .join(", "),
          cc: email.ccRecipients
            ?.map((r: any) => `${r.emailAddress.name} <${r.emailAddress.address}>`)
            .join(", "),
          date: email.receivedDateTime,
          importance: email.importance,
          hasAttachments: email.hasAttachments,
          body: email.body.contentType === "html" ? stripHtml(email.body.content) : email.body.content,
        };

        return {
          content: [{ type: "text" as const, text: JSON.stringify(formatted, null, 2) }],
        };
      } catch (err: any) {
        return {
          content: [{ type: "text" as const, text: `Error fetching email: ${err.message}` }],
        };
      }
    }
  );

  server.tool(
    "get_calendar_events",
    "Fetch calendar events from connected Outlook accounts. Returns upcoming and recent events with subject, time, location, attendees, and online meeting links.",
    {
      account_email: z.string().optional().describe("Filter events to a specific account email. If omitted, fetches from all accounts."),
      days_ahead: z.number().min(1).max(90).default(30).describe("Number of days ahead to fetch events (default: 30, max: 90)"),
      days_behind: z.number().min(0).max(30).default(7).describe("Number of days in the past to include (default: 7, max: 30)"),
    },
    async ({ account_email, days_ahead, days_behind }) => {
      const subError = checkSubscription(user);
      if (subError) {
        return { content: [{ type: "text" as const, text: subError }] };
      }
      const accounts = await storage.getAccountsByUserId(user.id);
      if (accounts.length === 0) {
        return {
          content: [{ type: "text" as const, text: "No Microsoft accounts are connected." }],
        };
      }

      const targetAccounts = account_email
        ? accounts.filter((a) => a.email.toLowerCase() === account_email.toLowerCase())
        : accounts;

      if (targetAccounts.length === 0) {
        return {
          content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }],
        };
      }

      const allEvents: OutlookCalendarEvent[] = [];

      for (const account of targetAccounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);

          const now = new Date();
          const startDate = new Date(now);
          startDate.setDate(startDate.getDate() - days_behind);
          const endDate = new Date(now);
          endDate.setDate(endDate.getDate() + days_ahead);

          const response = await fetch(
            `https://graph.microsoft.com/v1.0/me/calendarview?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$select=id,subject,start,end,location,organizer,attendees,isAllDay,bodyPreview,showAs,importance,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime&$top=100`,
            {
              headers: {
                Authorization: `Bearer ${accessToken}`,
                Prefer: 'outlook.timezone="UTC"',
              },
            }
          );

          if (!response.ok) throw new Error(await response.text());
          const data = await response.json();

          const enriched = (data.value || []).map((event: any) => ({
            ...event,
            accountId: account.id,
            accountEmail: account.email,
          }));
          allEvents.push(...enriched);
        } catch (err: any) {
          console.error(`MCP: Failed to fetch calendar for ${account.email}:`, err.message);
        }
      }

      allEvents.sort(
        (a, b) => new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime()
      );

      const formatted = allEvents.map((e) => ({
        subject: e.subject,
        start: e.start.dateTime,
        end: e.end.dateTime,
        isAllDay: e.isAllDay,
        location: e.location?.displayName || null,
        organizer: e.organizer
          ? `${e.organizer.emailAddress.name} <${e.organizer.emailAddress.address}>`
          : null,
        attendees: e.attendees?.map((a) => `${a.emailAddress.name} <${a.emailAddress.address}>`) || [],
        preview: e.bodyPreview,
        showAs: e.showAs,
        isOnlineMeeting: e.isOnlineMeeting,
        meetingUrl: e.onlineMeetingUrl || null,
        account: e.accountEmail,
      }));

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? JSON.stringify(formatted, null, 2)
              : "No calendar events found in the specified date range.",
          },
        ],
      };
    }
  );

  server.tool(
    "list_onedrive_files",
    "List files and folders in the user's OneDrive. Browse any folder path or list the root. Returns file names, types, sizes, last modified dates, and web links. Use the item ID with get_onedrive_file_content to read a specific file.",
    {
      account_email: z.string().describe("The email address of the Microsoft account whose OneDrive to browse"),
      folder_path: z
        .string()
        .optional()
        .describe("Folder path to list (e.g. 'Documents', 'Documents/Projects'). Omit or use 'root' for the OneDrive root."),
      limit: z.number().min(1).max(100).default(50).describe("Maximum number of items to return (default: 50, max: 100)"),
    },
    async ({ account_email, folder_path, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());

      if (!account) {
        return {
          content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }],
        };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const items = await listDriveItems(accessToken, folder_path || "root", limit);

        if (items.length === 0) {
          return {
            content: [{ type: "text" as const, text: `No files found in ${folder_path || "root"}.` }],
          };
        }

        const formatted = items.map((item: any) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          mimeType: item.file?.mimeType || null,
          size: item.size != null ? item.size : null,
          lastModified: item.lastModifiedDateTime,
          created: item.createdDateTime,
          webUrl: item.webUrl,
          parentPath: item.parentReference?.path?.replace("/drive/root:", "") || "/",
        }));

        return {
          content: [{ type: "text" as const, text: JSON.stringify(formatted, null, 2) }],
        };
      } catch (err: any) {
        return {
          content: [{ type: "text" as const, text: `Error listing OneDrive files: ${err.message}` }],
        };
      }
    }
  );

  server.tool(
    "search_onedrive_files",
    "Search for files in the user's OneDrive by name or keyword. Returns matching files and folders with their IDs, types, sizes, and web links.",
    {
      account_email: z.string().describe("The email address of the Microsoft account whose OneDrive to search"),
      query: z.string().describe("Search keyword to find files by name or content"),
      limit: z.number().min(1).max(50).default(25).describe("Maximum number of results to return (default: 25, max: 50)"),
    },
    async ({ account_email, query, limit }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());

      if (!account) {
        return {
          content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }],
        };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const items = await searchDriveItems(accessToken, query, limit);

        if (items.length === 0) {
          return {
            content: [{ type: "text" as const, text: `No files found matching "${query}".` }],
          };
        }

        const formatted = items.map((item: any) => ({
          id: item.id,
          name: item.name,
          type: item.folder ? "folder" : "file",
          mimeType: item.file?.mimeType || null,
          size: item.size != null ? item.size : null,
          lastModified: item.lastModifiedDateTime,
          webUrl: item.webUrl,
          parentPath: item.parentReference?.path?.replace("/drive/root:", "") || "/",
        }));

        return {
          content: [{ type: "text" as const, text: JSON.stringify(formatted, null, 2) }],
        };
      } catch (err: any) {
        return {
          content: [{ type: "text" as const, text: `Error searching OneDrive: ${err.message}` }],
        };
      }
    }
  );

  server.tool(
    "get_onedrive_file_content",
    "Get the content of a specific OneDrive file by its ID. For plain text files (txt, md, csv, json, xml, etc.) under 512 KB, returns the full file content. For other file types (Office documents, PDFs, images), returns metadata and a web link to open the file.",
    {
      account_email: z.string().describe("The email address of the Microsoft account that owns the file"),
      item_id: z.string().describe("The ID of the OneDrive file (obtained from list_onedrive_files or search_onedrive_files)"),
    },
    async ({ account_email, item_id }) => {
      const subError = checkSubscription(user);
      if (subError) return { content: [{ type: "text" as const, text: subError }] };

      const accounts = await storage.getAccountsByUserId(user.id);
      const account = accounts.find((a) => a.email.toLowerCase() === account_email.toLowerCase());

      if (!account) {
        return {
          content: [{ type: "text" as const, text: `No connected account found with email: ${account_email}` }],
        };
      }

      try {
        const accessToken = await getValidAccessToken(account.id);
        const { metadata, content, mimeType } = await getDriveItemContent(accessToken, item_id);

        if (content !== null) {
          const result = {
            name: metadata.name,
            mimeType,
            size: metadata.size,
            lastModified: metadata.lastModifiedDateTime,
            webUrl: metadata.webUrl,
            content,
          };
          return {
            content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
          };
        }

        const result = {
          name: metadata.name,
          mimeType,
          size: metadata.size,
          lastModified: metadata.lastModifiedDateTime,
          webUrl: metadata.webUrl,
          note: metadata.folder
            ? "This is a folder, not a file. Use list_onedrive_files with the parent path to browse its contents."
            : `This file type (${mimeType}) cannot be read as text. Open it directly at the web URL above, or use Microsoft Graph to download it.`,
        };
        return {
          content: [{ type: "text" as const, text: JSON.stringify(result, null, 2) }],
        };
      } catch (err: any) {
        return {
          content: [{ type: "text" as const, text: `Error fetching OneDrive file: ${err.message}` }],
        };
      }
    }
  );

  return server;
}

async function validateMcpAuth(req: Request, res: Response): Promise<User | null> {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith("Bearer ")) {
    res.status(401).json({
      jsonrpc: "2.0",
      error: { code: -32000, message: "Unauthorized: Bearer token required" },
      id: null,
    });
    return null;
  }

  const token = authHeader.slice(7);

  const user = await storage.getUserByMcpApiKey(token);
  if (user) {
    console.log(`MCP auth: Legacy API key matched user ${user.id} (${user.email}), subscription: ${user.subscriptionStatus}`);
    return user;
  }

  const oauthToken = await storage.getOAuthTokenByAccessToken(token);
  if (oauthToken) {
    if (new Date() > oauthToken.expiresAt) {
      console.log(`MCP auth: OAuth token expired for user ${oauthToken.userId}`);
      res.status(401).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Unauthorized: OAuth access token has expired" },
        id: null,
      });
      return null;
    }

    const oauthUser = await storage.getUser(oauthToken.userId);
    if (!oauthUser) {
      console.log(`MCP auth: User ${oauthToken.userId} not found for OAuth token`);
      res.status(403).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Forbidden: User not found" },
        id: null,
      });
      return null;
    }

    console.log(`MCP auth: OAuth token matched user ${oauthUser.id} (${oauthUser.email}), subscription: ${oauthUser.subscriptionStatus}`);
    return oauthUser;
  }

  console.log(`MCP auth: No matching API key or OAuth token found`);
  res.status(401).json({
    jsonrpc: "2.0",
    error: { code: -32000, message: "Unauthorized: Invalid API key or OAuth token" },
    id: null,
  });
  return null;
}

export function registerMcpRoutes(app: Express): void {
  const sessions: Record<string, { transport: StreamableHTTPServerTransport; server: McpServer; userId: string }> = {};

  app.post("/mcp", async (req: Request, res: Response) => {
    const user = await validateMcpAuth(req, res);
    if (!user) return;

    const sessionId = req.headers["mcp-session-id"] as string | undefined;

    if (sessionId && sessions[sessionId]) {
      if (sessions[sessionId].userId !== user.id) {
        return res.status(403).json({
          jsonrpc: "2.0",
          error: { code: -32000, message: "Forbidden: Session belongs to another user" },
          id: null,
        });
      }
      await sessions[sessionId].transport.handleRequest(req, res, req.body);
      return;
    }

    if (isInitializeRequest(req.body)) {
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
      });

      const server = createMcpServer(user);
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

    res.status(400).json({
      jsonrpc: "2.0",
      error: { code: -32000, message: "Bad Request: No valid session ID provided" },
      id: null,
    });
  });

  app.get("/mcp", async (req: Request, res: Response) => {
    const user = await validateMcpAuth(req, res);
    if (!user) return;

    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (sessionId && sessions[sessionId]) {
      if (sessions[sessionId].userId !== user.id) {
        return res.status(403).json({
          jsonrpc: "2.0",
          error: { code: -32000, message: "Forbidden: Session belongs to another user" },
          id: null,
        });
      }
      await sessions[sessionId].transport.handleRequest(req, res);
    } else {
      res.status(400).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Bad Request: No valid session ID provided" },
        id: null,
      });
    }
  });

  app.delete("/mcp", async (req: Request, res: Response) => {
    const user = await validateMcpAuth(req, res);
    if (!user) return;

    const sessionId = req.headers["mcp-session-id"] as string | undefined;
    if (sessionId && sessions[sessionId]) {
      if (sessions[sessionId].userId !== user.id) {
        return res.status(403).json({
          jsonrpc: "2.0",
          error: { code: -32000, message: "Forbidden: Session belongs to another user" },
          id: null,
        });
      }
      await sessions[sessionId].transport.handleRequest(req, res);
      await sessions[sessionId].server.close();
      delete sessions[sessionId];
    } else {
      res.status(400).json({
        jsonrpc: "2.0",
        error: { code: -32000, message: "Bad Request: No valid session ID provided" },
        id: null,
      });
    }
  });

  console.log("MCP server registered at /mcp endpoint");
}
