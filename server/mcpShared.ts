import { storage } from "./storage";

// Filters and allow-list checks shared by the v2 and v3 MCP servers.

export function stripHtml(html: string): string {
  return html ? html.replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim() : "";
}

export const DEFAULT_SENDER_BLOCKLIST = [
  "no-reply@", "noreply@", "donotreply@", "do-not-reply@", "no_reply@",
  "notifications@", "notification@", "receipts@", "billing@",
  "alerts@", "alert@", "auto-confirm@", "mailer-daemon@",
];

export function getSenderBlocklist(): string[] {
  const envList = process.env.EMAIL_SENDER_BLOCKLIST;
  if (envList) {
    return envList.split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  }
  return DEFAULT_SENDER_BLOCKLIST;
}

export function isBlockedSender(address: string): boolean {
  const addr = (address || "").toLowerCase();
  return getSenderBlocklist().some((blocked) => addr.includes(blocked));
}

export const NEWSLETTER_SENDER_PATTERNS = [
  "newsletter@", "newsletters@", "digest@", "weekly@",
  "marketing@", "promo@", "promotions@", "campaign@",
  "deals@", "offers@", "updates@",
];

export const NEWSLETTER_SUBJECT_KEYWORDS = [
  "newsletter", "weekly digest", "daily digest", "weekly roundup",
  "roundup", "this week in", "your weekly", "your daily",
  "weekly update", "monthly update", "weekly report", "daily brief",
];

export function isNewsletter(email: any): boolean {
  const from = (email.from?.emailAddress?.address || "").toLowerCase();
  if (NEWSLETTER_SENDER_PATTERNS.some((p) => from.includes(p))) return true;
  const subject = (email.subject || "").toLowerCase();
  return NEWSLETTER_SUBJECT_KEYWORDS.some((k) => subject.includes(k));
}

export function parseDateInput(input: string | undefined, boundary: "start" | "end"): string | undefined {
  if (!input) return undefined;
  const lower = input.trim().toLowerCase();
  const dayMs = 86400000;
  const now = new Date();
  let base: Date | null = null;
  if (lower === "today") base = now;
  else if (lower === "yesterday") base = new Date(now.getTime() - dayMs);
  else if (lower === "tomorrow") base = new Date(now.getTime() + dayMs);

  const toBoundaryIso = (d: Date) => {
    const y = d.getUTCFullYear();
    const m = d.getUTCMonth();
    const day = d.getUTCDate();
    return boundary === "start"
      ? new Date(Date.UTC(y, m, day, 0, 0, 0, 0)).toISOString()
      : new Date(Date.UTC(y, m, day, 23, 59, 59, 999)).toISOString();
  };

  if (base) return toBoundaryIso(base);

  const parsed = new Date(input);
  if (isNaN(parsed.getTime())) return undefined;
  if (/^\d{4}-\d{2}-\d{2}$/.test(input.trim())) return toBoundaryIso(parsed);
  return parsed.toISOString();
}

export function normalizePath(p: string): string {
  return (p || "")
    .replace(/^\/drive\/root:/i, "")
    .replace(/^\/+/, "")
    .replace(/\/+$/, "")
    .toLowerCase();
}

export function isWithinAllowed(itemPath: string, allowed: string[]): boolean {
  if (!allowed || allowed.length === 0) return true;
  const np = normalizePath(itemPath);
  return allowed.some((a) => {
    const na = normalizePath(a);
    return na === "" || np === na || np.startsWith(na + "/");
  });
}

export function senderMatches(address: string, approved: string[]): boolean {
  if (!approved || approved.length === 0) return true;
  const addr = (address || "").toLowerCase();
  return approved.some((raw) => {
    const e = raw.trim().toLowerCase();
    if (!e) return false;
    if (e.startsWith("@")) return addr.endsWith(e);
    if (e.includes("@")) return addr === e;
    return addr.endsWith("@" + e) || addr === e;
  });
}

export async function getEffectiveSettings(userId: string) {
  const s = await storage.getConnectorSettings(userId);
  return {
    focusedOnly: s?.focusedOnly ?? false,
    approvedSenders: s?.approvedSenders ?? [],
    allowedFolders: s?.allowedOnedriveFolders ?? [],
  };
}

export function emailAllowedBySettings(
  email: any,
  settings: { focusedOnly: boolean; approvedSenders: string[] },
): string | null {
  if (settings.approvedSenders.length > 0) {
    const from = email.from?.emailAddress?.address || "";
    if (!senderMatches(from, settings.approvedSenders)) {
      return "This email is from a sender outside your approved-sender allow-list. Update Connector Settings to access it.";
    }
  }
  if (settings.focusedOnly && email.inferenceClassification && email.inferenceClassification !== "focused") {
    return "This email is in the Other inbox, but your Connector Settings restrict access to Focused inbox only.";
  }
  return null;
}
