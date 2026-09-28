import { test } from "node:test";
import assert from "node:assert/strict";
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
} from "../mcpV3Format";
import { toHandles, fromHandle, memoryBackend, setHandleBackendForTest } from "../handleStore";

test("short ids are stable per item and scoped by kind and user", async () => {
  setHandleBackendForTest(memoryBackend());
  const t = { accountId: "a1", accountEmail: "a@x.com", ref: "AAMkAG-long-graph-id" };
  assert.deepEqual(await toHandles("u1", "m", [t, { ...t, ref: "other" }, t]), ["m1", "m2", "m1"]);
  assert.deepEqual(await toHandles("u1", "m", [{ ...t, accountId: "a2" }]), ["m3"]);
  assert.deepEqual(await toHandles("u1", "e", [t]), ["e1"]);
  assert.deepEqual(await fromHandle("u1", "m", "m1"), t);
  assert.equal(await fromHandle("u1", "e", "m1"), null);
  assert.equal(await fromHandle("u2", "m", "m1"), null);
});

test("row flattens separators and drops trailing empty cells", () => {
  assert.equal(row(["m1", "a|b", "line1\nline2", "", null]), "m1|a b|line1 line2");
});

test("dates render as short UTC, with or without a Z suffix", () => {
  assert.equal(shortDateTime("2026-09-24T14:05:00Z"), "09-24 14:05");
  assert.equal(shortDateTime("2026-09-24T14:05:00.0000000"), "09-24 14:05");
  assert.equal(shortDateTime(undefined), "");
  assert.equal(eventWhen("2026-09-24T14:00:00.0000000", "2026-09-24T15:30:00.0000000", false), "Thu 09-24 14:00-15:30");
  assert.equal(eventWhen("2026-09-24T00:00:00", "2026-09-25T00:00:00", true), "Thu 09-24 all-day");
  assert.equal(eventWhen("2026-09-24T00:00:00", "2026-09-27T00:00:00", true), "Thu 09-24..Sat 09-26 all-day");
});

test("person skips a name that repeats the address", () => {
  assert.equal(person({ name: "Jane Doe", address: "jane@acme.com" }), "Jane Doe <jane@acme.com>");
  assert.equal(person({ name: "jane@acme.com", address: "jane@acme.com" }), "jane@acme.com");
  assert.equal(person(null), "");
});

test("humanSize", () => {
  assert.equal(humanSize(512), "512B");
  assert.equal(humanSize(2048), "2K");
  assert.equal(humanSize(5 * 1024 * 1024), "5.0M");
});

test("compactText shortens tracking links, whitespace, and length", () => {
  const long = `https://click.example.com/track?u=${"x".repeat(100)}`;
  assert.equal(compactText(`See ${long} now`, 100), "See [link:click.example.com] now");
  assert.equal(compactText("keep https://acme.com/a", 100), "keep https://acme.com/a");
  assert.equal(compactText("a   b\n\n\n\nc", 100), "a b\n\nc");
  assert.equal(compactText("abcdefghij", 4), "abcd…[+6 chars]");
});

test("kqlTerm strips quote breakouts", () => {
  assert.equal(kqlTerm('budget" OR "x'), "budget  OR  x");
});

const ev = (key: string, account: string, start: string, end: string, extra: Partial<CalendarItem> = {}): CalendarItem => ({
  key, subject: `Meeting ${key}`, account, start, end, isAllDay: false, showAs: "busy", ...extra,
});

test("findConflicts finds overlaps within and across calendars", () => {
  const { conflicts, duplicates } = findConflicts([
    ev("1", "work@a.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00"),
    ev("2", "home@b.com", "2026-09-24T15:30:00", "2026-09-24T16:30:00"),
    ev("3", "work@a.com", "2026-09-24T16:00:00", "2026-09-24T17:00:00"),
  ]);
  assert.equal(duplicates, 0);
  assert.deepEqual(
    conflicts.map((c) => [c.a.key, c.b.key, c.crossCalendar, c.overlapStart]),
    [
      ["1", "2", true, "2026-09-24T15:30:00.000Z"],
      ["2", "3", true, "2026-09-24T16:00:00.000Z"],
    ],
  );
});

test("findConflicts ignores free, cancelled, declined and the same meeting on two calendars", () => {
  const { conflicts, duplicates } = findConflicts([
    ev("1", "work@a.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00", { subject: "Board call" }),
    ev("2", "home@b.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00", { subject: "FW: Board call" }),
    ev("3", "work@a.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00", { showAs: "free" }),
    ev("4", "work@a.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00", { isCancelled: true }),
    ev("5", "home@b.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00", { declined: true }),
  ]);
  assert.equal(conflicts.length, 0);
  assert.equal(duplicates, 1);
});

test("findConflicts treats an out-of-office day as blocking", () => {
  const { conflicts } = findConflicts([
    ev("1", "work@a.com", "2026-09-24T00:00:00", "2026-09-25T00:00:00", { isAllDay: true, showAs: "oof" }),
    ev("2", "work@a.com", "2026-09-24T15:00:00", "2026-09-24T16:00:00"),
  ]);
  assert.equal(conflicts.length, 1);
  assert.equal(conflicts[0].crossCalendar, false);
});
