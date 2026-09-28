// Runs the real v3 server over HTTP with an MCP client. Storage and Microsoft
// Graph are stubbed, so no database or Microsoft account is needed. Two
// accounts are connected (work and personal) to cover cross-account behavior.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import express from "express";
import type { AddressInfo } from "node:net";
import type { Server } from "node:http";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { storage } from "../storage";
import { registerMcpV3Routes } from "../mcpV3";
import { memoryBackend, setHandleBackendForTest } from "../handleStore";

const API_KEY = "test-key";
const baseUser: any = { id: "u1", email: "amanda@example.com", subscriptionStatus: "active" };
let user: any = { ...baseUser };
const future = new Date(Date.now() + 3600_000);
const work: any = { id: "acct-work", userId: "u1", email: "amanda@example.com", displayName: "Amanda", accountType: "work", accessToken: "t-work", refreshToken: "r", tokenExpiry: future };
const home: any = { id: "acct-home", userId: "u1", email: "amanda@home.com", displayName: "Amanda", accountType: "personal", accessToken: "t-home", refreshToken: "r", tokenExpiry: future };
let settings: any = undefined;

const now = Date.now();
const iso = (msAgo: number) => new Date(now - msAgo).toISOString();
const GRAPH_ID = "AAMkAGI2TG93AAA=" + "x".repeat(120);
const inbox = [
  { id: GRAPH_ID, subject: "Deal memo", from: { emailAddress: { name: "Jane Doe", address: "jane@acme.com" } }, receivedDateTime: iso(3600_000), bodyPreview: "Here is the memo", isRead: false, hasAttachments: true, importance: "high", inferenceClassification: "focused" },
  { id: "id-2", subject: "Your receipt", from: { emailAddress: { name: "Shop", address: "receipts@shop.com" } }, receivedDateTime: iso(7200_000), bodyPreview: "Thanks", isRead: true, hasAttachments: false, importance: "normal", inferenceClassification: "focused" },
  { id: "id-3", subject: "Weekly digest", from: { emailAddress: { name: "News", address: "news@site.com" } }, receivedDateTime: iso(9000_000), bodyPreview: "This week", isRead: true, hasAttachments: false, importance: "normal", inferenceClassification: "other" },
];
const event = (id: string, subject: string, start: string, end: string) => ({
  id, subject, start: { dateTime: `2026-09-24T${start}:00.0000000` }, end: { dateTime: `2026-09-24T${end}:00.0000000` },
  isAllDay: false, showAs: "busy", isOnlineMeeting: false, organizer: { emailAddress: { name: "Amanda", address: "amanda@example.com" } }, location: { displayName: "" },
});
const calendars: Record<string, any[]> = {
  "t-work": [event("w1", "Lender call", "15:00", "15:30"), event("w2", "Board call", "16:00", "17:00")],
  "t-home": [event("h1", "Dentist", "15:15", "16:00"), event("h2", "Board call", "16:00", "17:00")],
};
const graphCalls: { method: string; path: string; token: string; body?: any }[] = [];

let server: Server;
let url: URL;
const realFetch = globalThis.fetch;

before(async () => {
  setHandleBackendForTest(memoryBackend());
  Object.assign(storage, {
    getUserByMcpApiKey: async (key: string) => (key === API_KEY ? user : undefined),
    getOAuthTokenByAccessToken: async () => undefined,
    getAccountsByUserId: async () => [work, home],
    getAccount: async (id: string) => [work, home].find((a) => a.id === id),
    getConnectorSettings: async () => settings,
  });

  globalThis.fetch = (async (input: any, init: any = {}) => {
    const target = String(input instanceof Request ? input.url : input);
    if (!target.startsWith("https://graph.microsoft.com/")) return realFetch(input, init);
    const method = init.method || "GET";
    const token = String(init.headers?.Authorization || "").replace("Bearer ", "");
    const path = decodeURIComponent(new URL(target).pathname).replace("/v1.0", "");
    graphCalls.push({ method, path, token, body: init.body ? JSON.parse(init.body) : undefined });
    const json = (data: any) => new Response(JSON.stringify(data), { status: 200, headers: { "Content-Type": "application/json" } });

    if (path.endsWith("/mailFolders/inbox/messages")) return json({ value: token === "t-work" ? inbox : [] });
    if (path.endsWith("/createReply")) return json({ id: "draft1", subject: "RE: Deal memo", webLink: "https://outlook/draft1" });
    if (path === "/me/messages" && method === "POST") return json({ id: "draft2", webLink: "https://outlook/draft2" });
    if (path.startsWith("/me/messages/")) {
      return json({
        ...inbox[0],
        toRecipients: [{ emailAddress: { name: "Amanda", address: "amanda@example.com" } }, { emailAddress: { name: "Bob", address: "bob@acme.com" } }],
        uniqueBody: { contentType: "text", content: "New reply text only." },
        body: { contentType: "text", content: "New reply text only.\n\n> quoted history ..." },
      });
    }
    if (path.endsWith("/calendarView")) return json({ value: calendars[token] });
    if (path === "/me/chats") return json({ value: [{ id: "chat-1", chatType: "oneOnOne", topic: null, lastUpdatedDateTime: iso(60_000), members: [{ displayName: "Amanda", email: "amanda@example.com" }, { displayName: "Bob", email: "bob@acme.com" }] }] });
    if (path === "/me") return json({ id: "me-id" });
    if (path === "/chats" && method === "POST") return json({ id: "chat-new" });
    if (/^\/me\/chats\/[^/]+\/messages$/.test(path) && method === "POST") return json({ id: "msg-1", createdDateTime: "2026-09-28T14:05:00Z" });
    return new Response("not mocked", { status: 404 });
  }) as typeof fetch;

  const app = express();
  app.use(express.json());
  registerMcpV3Routes(app);
  server = app.listen(0);
  await new Promise((r) => server.once("listening", r));
  url = new URL(`http://127.0.0.1:${(server.address() as AddressInfo).port}/mcp/v3`);
});

after(() => {
  globalThis.fetch = realFetch;
  server.close();
});

async function connect(): Promise<Client> {
  const client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(
    new StreamableHTTPClientTransport(url, { requestInit: { headers: { Authorization: `Bearer ${API_KEY}` } } }),
  );
  return client;
}

const textOf = (r: any) => r.content.map((c: any) => c.text).join("\n");

async function firstMailId(client: Client): Promise<string> {
  const list = textOf(await client.callTool({ name: "mail_find", arguments: {} }));
  return list.split("\n")[1].split("|")[0];
}

test("rejects requests without a token and points to OAuth discovery", async () => {
  const res = await realFetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
  assert.equal(res.status, 401);
  assert.match(res.headers.get("www-authenticate") || "", /oauth-protected-resource\/mcp\/v3/);
});

test("paid users get 12 tools and instructions that say to ask for the draft account", async () => {
  user = { ...baseUser };
  const client = await connect();
  const { tools } = await client.listTools();
  assert.deepEqual(tools.map((t) => t.name).sort(), [
    "calendar_conflicts", "calendar_find", "event_delete", "event_save", "file_read", "files_find",
    "mail_draft", "mail_find", "mail_read", "teams_find", "teams_read", "teams_send",
  ]);
  const instructions = client.getInstructions() || "";
  assert.match(instructions, /amanda@example\.com \(work\), amanda@home\.com \(personal\)/);
  assert.match(instructions, /always ask the user which account to draft from/);
  await client.close();
});

test("free tier only sees mail_find", async () => {
  user = { ...baseUser, subscriptionStatus: "free" };
  const client = await connect();
  const { tools } = await client.listTools();
  assert.deepEqual(tools.map((t) => t.name), ["mail_find"]);
  await client.close();
  user = { ...baseUser };
});

test("mail_find briefing: compact rows, short ids, noise hidden", async () => {
  const client = await connect();
  const text = textOf(await client.callTool({ name: "mail_find", arguments: {} }));
  const lines = text.split("\n");
  assert.equal(lines[0], "id|received|from|subject|flags|preview");
  assert.match(lines[1], /^m\d+\|\d\d-\d\d \d\d:\d\d\|Jane Doe <jane@acme\.com>\|Deal memo\|UA!\|Here is the memo$/);
  assert.ok(!text.includes(GRAPH_ID), "long Graph ids never reach the model");
  assert.match(text, /2 automated\/newsletter hidden/);
  await client.close();
});

test("mail_read returns only the new text", async () => {
  const client = await connect();
  const id = await firstMailId(client);
  const read = textOf(await client.callTool({ name: "mail_read", arguments: { ids: [id] } }));
  assert.match(read, new RegExp(`^## ${id} · Deal memo`));
  assert.match(read, /New reply text only\./);
  assert.match(read, /in: amanda@example\.com/);
  assert.ok(!read.includes("quoted history"));
  await client.close();
});

test("mail_draft requires an account", async () => {
  const client = await connect();
  let rejected = false;
  try {
    const res: any = await client.callTool({ name: "mail_draft", arguments: { body: "Hi", to: ["x@y.com"], subject: "Hi" } });
    rejected = res.isError === true;
  } catch {
    rejected = true;
  }
  assert.ok(rejected, "a draft without an account is refused");
  assert.ok(!graphCalls.some((c) => c.method === "POST" && c.path === "/me/messages"));
  await client.close();
});

test("a reply from the receiving account uses Outlook's reply draft", async () => {
  const client = await connect();
  const id = await firstMailId(client);
  const res = await client.callTool({ name: "mail_draft", arguments: { account: "amanda@example.com", reply_to: id, body: "Thanks Jane" } });
  assert.match(textOf(res), /Reply draft saved in amanda@example\.com, not sent/);
  const call = graphCalls.find((c) => c.path.endsWith("/createReply"));
  assert.ok(call && call.path.includes(GRAPH_ID) && call.token === "t-work");
  assert.deepEqual(call!.body, { comment: "Thanks Jane" });
  await client.close();
});

test("a reply can be drafted from a different account", async () => {
  const client = await connect();
  const id = await firstMailId(client);
  const res = await client.callTool({
    name: "mail_draft",
    arguments: { account: "amanda@home.com", reply_to: id, body: "Replying from home", reply_all: true },
  });
  assert.match(textOf(res), /Reply draft saved in amanda@home\.com \(original arrived in amanda@example\.com\), not sent/);
  const call = graphCalls.find((c) => c.method === "POST" && c.path === "/me/messages");
  assert.equal(call?.token, "t-home");
  assert.equal(call!.body.subject, "RE: Deal memo");
  assert.deepEqual(call!.body.toRecipients, [{ emailAddress: { address: "jane@acme.com" } }]);
  assert.deepEqual(call!.body.ccRecipients, [{ emailAddress: { address: "bob@acme.com" } }]);
  assert.match(call!.body.body.content, /^Replying from home\n\n-----Original message-----/);
  await client.close();
});

test("allow-lists are enforced on direct reads", async () => {
  const client = await connect();
  const id = await firstMailId(client);
  settings = { focusedOnly: false, approvedSenders: ["@other.com"], allowedOnedriveFolders: [] };
  try {
    const read = textOf(await client.callTool({ name: "mail_read", arguments: { ids: [id] } }));
    assert.match(read, /outside your approved-sender allow-list/);
    assert.ok(!read.includes("New reply text only"));
  } finally {
    settings = undefined;
  }
  await client.close();
});

test("unknown ids give a clear message", async () => {
  const client = await connect();
  const res = await client.callTool({ name: "mail_read", arguments: { ids: ["m9999"] } });
  assert.match(textOf(res), /Unknown id m9999/);
  await client.close();
});

test("calendar_find flags double bookings across calendars", async () => {
  const client = await connect();
  const text = textOf(await client.callTool({ name: "calendar_find", arguments: {} }));
  const lines = text.split("\n");
  assert.equal(lines[0], "id|when|subject|location|organizer|flags|account");
  assert.match(text, /^e\d+\|Thu 09-24 15:00-15:30\|Lender call\|\|me\|C\|amanda@example\.com$/m);
  assert.match(text, /^e\d+\|Thu 09-24 15:15-16:00\|Dentist\|\|.*\|C\|amanda@home\.com$/m);
  assert.match(text, /^e\d+\|Thu 09-24 16:00-17:00\|Board call\|\|me\|\|amanda@example\.com$/m);
  assert.match(text, /1 double-booking in this window/);
  await client.close();
});

test("calendar_conflicts lists the clash and ignores the shared meeting", async () => {
  const client = await connect();
  const text = textOf(await client.callTool({ name: "calendar_conflicts", arguments: {} }));
  const lines = text.split("\n");
  assert.equal(lines[0], "overlap|event|clashes with|scope");
  assert.match(lines[1], /^Thu 09-24 15:15-15:30\|e\d+ Lender call \(amanda@example\.com\)\|e\d+ Dentist \(amanda@home\.com\)\|across calendars$/);
  assert.equal(lines.length, 3);
  assert.match(lines[2], /1 meeting on more than one calendar ignored/);
  await client.close();
});

test("teams_find lists chats from work accounts only", async () => {
  const client = await connect();
  const text = textOf(await client.callTool({ name: "teams_find", arguments: {} }));
  assert.match(text, /^chat\|type\|topic or members\|last active\nc\d+\|oneOnOne\|Bob\|/);
  assert.ok(!graphCalls.some((c) => c.path === "/me/chats" && c.token === "t-home"));
  await client.close();
});

test("teams_send posts to an existing chat or starts one", async () => {
  const client = await connect();
  const list = textOf(await client.callTool({ name: "teams_find", arguments: {} }));
  const chat = list.split("\n")[1].split("|")[0];

  const sent = textOf(await client.callTool({ name: "teams_send", arguments: { chat, text: "On my way" } }));
  assert.match(sent, new RegExp(`^Sent in ${chat} from amanda@example\\.com at 09-28 14:05\\.$`));
  const post = graphCalls.find((c) => c.path === "/me/chats/chat-1/messages");
  assert.deepEqual(post?.body, { body: { contentType: "text", content: "On my way" } });

  const started = textOf(await client.callTool({ name: "teams_send", arguments: { people: ["bob@acme.com"], text: "Quick question" } }));
  assert.match(started, /^Sent in c\d+ from amanda@example\.com/);
  const create = graphCalls.find((c) => c.path === "/chats" && c.method === "POST");
  assert.equal(create?.body.chatType, "oneOnOne");
  assert.deepEqual(
    create!.body.members.map((m: any) => m["user@odata.bind"]),
    ["https://graph.microsoft.com/v1.0/users('me-id')", "https://graph.microsoft.com/v1.0/users('bob%40acme.com')"],
  );
  assert.ok(graphCalls.some((c) => c.path === "/me/chats/chat-new/messages"));
  await client.close();
});
