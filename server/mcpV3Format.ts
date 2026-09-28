// Pure helpers for the v3 MCP server. v3 returns pipe-delimited rows instead of
// JSON (no repeated keys); short ids like "m3" come from handleStore.ts.

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

export interface CalendarItem {
  key: string;
  subject: string;
  account: string;
  start: string;
  end: string;
  isAllDay: boolean;
  showAs?: string;
  isCancelled?: boolean;
  declined?: boolean;
}

export interface Conflict {
  a: CalendarItem;
  b: CalendarItem;
  overlapStart: string;
  overlapEnd: string;
  crossCalendar: boolean;
}

/**
 * Finds double bookings. Free, cancelled and declined events don't block time.
 * The same meeting on two calendars (same subject and times on different
 * accounts, e.g. an invite sent to both work and personal) is a duplicate,
 * not a conflict.
 */
export function findConflicts(items: CalendarItem[]): { conflicts: Conflict[]; duplicates: number } {
  const busy = items
    .filter((e) => !e.isCancelled && !e.declined && e.showAs !== "free")
    .map((e) => ({ e, s: parseUtc(e.start).getTime(), t: parseUtc(e.end).getTime() }))
    .filter((x) => x.t > x.s)
    .sort((x, y) => x.s - y.s);

  const norm = (s: string) => s.trim().toLowerCase().replace(/^(re|fw|fwd):\s*/i, "");
  const conflicts: Conflict[] = [];
  let duplicates = 0;
  for (let i = 0; i < busy.length; i++) {
    for (let j = i + 1; j < busy.length && busy[j].s < busy[i].t; j++) {
      const x = busy[i];
      const y = busy[j];
      const sameMeeting =
        x.e.account !== y.e.account && x.s === y.s && x.t === y.t && norm(x.e.subject) === norm(y.e.subject);
      if (sameMeeting) {
        duplicates++;
        continue;
      }
      conflicts.push({
        a: x.e,
        b: y.e,
        overlapStart: new Date(Math.max(x.s, y.s)).toISOString(),
        overlapEnd: new Date(Math.min(x.t, y.t)).toISOString(),
        crossCalendar: x.e.account !== y.e.account,
      });
    }
  }
  return { conflicts, duplicates };
}
