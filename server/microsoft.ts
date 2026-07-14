import { storage } from "./storage";

const MICROSOFT_CLIENT_ID = process.env.MICROSOFT_CLIENT_ID;
const MICROSOFT_CLIENT_SECRET = process.env.MICROSOFT_CLIENT_SECRET;

if (!MICROSOFT_CLIENT_ID || !MICROSOFT_CLIENT_SECRET) {
  console.warn("Warning: MICROSOFT_CLIENT_ID or MICROSOFT_CLIENT_SECRET not set. OAuth flow will not work.");
}
const SCOPES = [
  "openid",
  "email",
  "profile",
  "offline_access",
  "Mail.Read",
  "Mail.ReadWrite",
  "Calendars.Read",
  "User.Read",
  "Files.Read",
  "Chat.Read",
];

function getRedirectUri(requestHost?: string): string {
  const host = requestHost || process.env.REPLIT_DEPLOYMENT_URL || process.env.REPLIT_DEV_DOMAIN;
  if (host) {
    const cleanHost = host.replace(/^https?:\/\//, '');
    const protocol = cleanHost.includes("localhost") ? "http" : "https";
    return `${protocol}://${cleanHost}/api/auth/callback`;
  }
  return `http://localhost:5000/api/auth/callback`;
}

export function getAuthUrl(state: string, requestHost?: string): string {
  const redirectUri = getRedirectUri(requestHost);
  const params = new URLSearchParams({
    client_id: MICROSOFT_CLIENT_ID || "",
    redirect_uri: redirectUri,
    response_type: "code",
    scope: SCOPES.join(" "),
    response_mode: "query",
    state: state,
    prompt: "select_account",
  });

  return `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params}`;
}

export async function exchangeCodeForTokens(code: string, requestHost?: string) {
  const redirectUri = getRedirectUri(requestHost);
  console.log("TOKEN EXCHANGE - redirect_uri:", redirectUri);
  console.log("TOKEN EXCHANGE - requestHost:", requestHost);
  const params = new URLSearchParams({
    client_id: MICROSOFT_CLIENT_ID || "",
    client_secret: MICROSOFT_CLIENT_SECRET || "",
    code: code,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
    scope: SCOPES.join(" "),
  });

  const response = await fetch(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Token exchange failed: ${error}`);
  }

  return await response.json();
}

export async function refreshAccessToken(refreshToken: string) {
  const params = new URLSearchParams({
    client_id: MICROSOFT_CLIENT_ID || "",
    client_secret: MICROSOFT_CLIENT_SECRET || "",
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: SCOPES.join(" "),
  });

  const response = await fetch(
    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Token refresh failed: ${error}`);
  }

  return await response.json();
}

export async function getValidAccessToken(accountId: string): Promise<string> {
  const account = await storage.getAccount(accountId);
  if (!account) throw new Error("Account not found");

  const now = new Date();
  const expiryBuffer = new Date(account.tokenExpiry.getTime() - 5 * 60 * 1000);

  if (now < expiryBuffer) {
    return account.accessToken;
  }

  const tokens = await refreshAccessToken(account.refreshToken);
  const newExpiry = new Date(Date.now() + tokens.expires_in * 1000);
  await storage.updateAccountTokens(
    accountId,
    tokens.access_token,
    tokens.refresh_token || account.refreshToken,
    newExpiry
  );

  return tokens.access_token;
}

export async function getUserProfile(accessToken: string) {
  const response = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${accessToken}` },
  });

  if (!response.ok) throw new Error("Failed to fetch user profile");
  return await response.json();
}

export async function getEmails(accessToken: string, top: number = 50, includeBody: boolean = true) {
  const fields = includeBody
    ? "id,subject,from,toRecipients,receivedDateTime,bodyPreview,body,isRead,hasAttachments,importance,conversationId,parentFolderId"
    : "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments,importance";
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages?$select=${fields}&$top=${top}&$orderby=receivedDateTime DESC`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to fetch emails: ${error}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function getCalendarEvents(accessToken: string) {
  const now = new Date();
  const startDate = new Date(now);
  startDate.setDate(startDate.getDate() - 7);
  const endDate = new Date(now);
  endDate.setDate(endDate.getDate() + 30);

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/calendarview?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$select=id,subject,start,end,location,organizer,attendees,isAllDay,bodyPreview,showAs,importance,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime&$top=100`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Prefer: 'outlook.timezone="UTC"',
      },
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to fetch calendar: ${error}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function getFilteredEmails(
  accessToken: string,
  opts: {
    top?: number;
    afterDateTime?: string;
    beforeDateTime?: string;
    sender?: string;
    focusedOnly?: boolean;
  } = {}
) {
  const { top = 25, afterDateTime, beforeDateTime, sender, focusedOnly } = opts;
  const filters: string[] = [];
  if (afterDateTime) filters.push(`receivedDateTime ge ${afterDateTime}`);
  if (beforeDateTime) filters.push(`receivedDateTime lt ${beforeDateTime}`);
  if (sender) filters.push(`from/emailAddress/address eq '${sender.replace(/'/g, "''")}'`);
  if (focusedOnly) filters.push(`inferenceClassification eq 'focused'`);

  const fields = "id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments,importance,inferenceClassification";
  const params = new URLSearchParams();
  params.set("$select", fields);
  params.set("$top", String(top));
  if (filters.length) params.set("$filter", filters.join(" and "));
  // Graph rejects $orderby when $filter uses from/toRecipients.
  // inferenceClassification (focused) works fine with ordering.
  const hasSenderFilter = Boolean(sender);
  if (!hasSenderFilter) params.set("$orderby", "receivedDateTime DESC");

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages?${params.toString()}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to fetch emails: ${error}`);
  }

  const data = await response.json();
  const values = data.value || [];
  if (hasSenderFilter) {
    values.sort(
      (a: any, b: any) =>
        new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
    );
  }
  return values;
}

export async function getEmailById(accessToken: string, messageId: string) {
  const fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,isRead,hasAttachments,importance,webLink,inferenceClassification";
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}?$select=${fields}`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to fetch email: ${error}`);
  }

  return await response.json();
}

export async function createEmailDraft(
  accessToken: string,
  opts: { to: string[]; cc?: string[]; subject: string; body: string; bodyType?: "Text" | "HTML" }
) {
  const message: any = {
    subject: opts.subject,
    body: { contentType: opts.bodyType || "Text", content: opts.body },
    toRecipients: opts.to.map((address) => ({ emailAddress: { address } })),
  };
  if (opts.cc && opts.cc.length) {
    message.ccRecipients = opts.cc.map((address) => ({ emailAddress: { address } }));
  }

  const response = await fetch("https://graph.microsoft.com/v1.0/me/messages", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(message),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to create draft: ${error}`);
  }

  return await response.json();
}

export async function listDriveItems(accessToken: string, path: string = "root", top: number = 50) {
  const endpoint = path === "root" || path === ""
    ? `https://graph.microsoft.com/v1.0/me/drive/root/children`
    : `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path)}:/children`;

  const response = await fetch(
    `${endpoint}?$select=id,name,size,file,folder,lastModifiedDateTime,createdDateTime,webUrl,parentReference&$top=${top}&$orderby=name`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to list OneDrive items: ${error}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function searchDriveItems(accessToken: string, query: string, top: number = 25) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/root/search(q='${encodeURIComponent(query)}')?$select=id,name,size,file,folder,lastModifiedDateTime,createdDateTime,webUrl,parentReference&$top=${top}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to search OneDrive: ${error}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function getDriveItemContent(accessToken: string, itemId: string): Promise<{ metadata: any; content: string | null; mimeType: string }> {
  const metaResponse = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}?$select=id,name,size,file,folder,lastModifiedDateTime,createdDateTime,webUrl,parentReference`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!metaResponse.ok) {
    const error = await metaResponse.text();
    throw new Error(`Failed to fetch OneDrive item: ${error}`);
  }

  const metadata = await metaResponse.json();
  const mimeType: string = metadata.file?.mimeType || "application/octet-stream";

  const TEXT_MIMETYPES = [
    "text/plain",
    "text/csv",
    "text/markdown",
    "text/html",
    "text/css",
    "text/javascript",
    "application/json",
    "application/xml",
    "text/xml",
  ];

  const isTextFile = TEXT_MIMETYPES.some((t) => mimeType.startsWith(t)) ||
    metadata.name?.match(/\.(txt|md|csv|json|xml|html|htm|css|js|ts|py|sh|yaml|yml|log)$/i);

  const MAX_TEXT_SIZE = 512 * 1024;

  if (isTextFile && metadata.size && metadata.size <= MAX_TEXT_SIZE) {
    const contentResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`,
      {
        headers: { Authorization: `Bearer ${accessToken}` },
      }
    );

    if (contentResponse.ok) {
      const text = await contentResponse.text();
      return { metadata, content: text, mimeType };
    }
  }

  return { metadata, content: null, mimeType };
}

export async function getTeamsChats(accessToken: string, top: number = 20) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/chats?$expand=members&$top=${top}`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to fetch Teams chats: ${error}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function getTeamsChatMessages(accessToken: string, chatId: string, top: number = 50) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/chats/${encodeURIComponent(chatId)}/messages?$top=${top}&$orderby=createdDateTime desc`,
    {
      headers: { Authorization: `Bearer ${accessToken}` },
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to fetch Teams chat messages: ${error}`);
  }

  const data = await response.json();
  return data.value || [];
}

export async function searchTeamsMessages(accessToken: string, query: string, top: number = 25) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/search/query`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        requests: [
          {
            entityTypes: ["chatMessage"],
            query: { queryString: query },
            from: 0,
            size: top,
          },
        ],
      }),
    }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to search Teams messages: ${error}`);
  }

  const data = await response.json();
  const hits = data.value?.[0]?.hitsContainers?.[0]?.hits || [];
  return hits;
}
