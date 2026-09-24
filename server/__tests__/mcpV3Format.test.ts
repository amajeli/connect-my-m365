import { test } from "node:test";
import assert from "node:assert/strict";
import {
  toHandle,
  fromHandle,
  resetHandlesForTest,
  row,
  shortDateTime,
  eventWhen,
  person,
  humanSize,
  compactText,
  kqlTerm,
} from "../mcpV3Format";

test("handles are short, stable per item, and scoped by kind and user", () => {
  resetHandlesForTest();
  const target = { accountId: "a1", accountEmail: "a@x.com", ref: "AAMkAG-long-graph-id" };
  const h = toHandle("u1", "m", target);
  assert.equal(h, "m1");
  assert.equal(toHandle("u1", "m", target), "m1");
  assert.equal(toHandle("u1", "m", { ...target, ref: "other" }), "m2");
  assert.equal(toHandle("u1", "m", { ...target, accountId: "a2" }), "m3");
  assert.deepEqual(fromHandle("u1", "m", "m1"), target);
  assert.equal(fromHandle("u1", "e", "m1"), null);
  assert.equal(fromHandle("u2", "m", "m1"), null);
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
