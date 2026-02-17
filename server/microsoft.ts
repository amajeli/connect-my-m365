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
  "Calendars.Read",
  "User.Read",
];

function getRedirectUri(): string {
  const host = process.env.REPLIT_DEPLOYMENT_URL || process.env.REPLIT_DEV_DOMAIN;
  if (host) {
    const protocol = host.includes("localhost") ? "http" : "https";
    return `${protocol}://${host}/api/auth/callback`;
  }
  return `http://localhost:5000/api/auth/callback`;
}

export function getAuthUrl(state: string): string {
  const redirectUri = getRedirectUri();
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

export async function exchangeCodeForTokens(code: string) {
  const redirectUri = getRedirectUri();
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

export async function getEmails(accessToken: string, top: number = 50) {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages?$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,body,isRead,hasAttachments,importance,conversationId,parentFolderId&$top=${top}&$orderby=receivedDateTime DESC`,
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
    `https://graph.microsoft.com/v1.0/me/calendarview?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$select=id,subject,start,end,location,organizer,attendees,isAllDay,bodyPreview,body,showAs,importance,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime&$top=100`,
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
