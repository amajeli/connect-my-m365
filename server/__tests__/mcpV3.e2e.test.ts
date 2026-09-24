// Runs the real v3 server over HTTP with an MCP client. Storage and Microsoft
// Graph are stubbed, so no database or Microsoft account is needed.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import express from "express";
import type { AddressInfo } from "node:net";
import type { Server } from "node:http";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";
import { storage } from "../storage";
import { registerMcpV3Routes } from "../mcpV3";

const API_KEY = "test-key";
const baseUser: any = { id: "u1", email: "amanda@example.com", subscriptionStatus: "active" };
let user: any = { ...baseUser };
const account: any = {
  id: "acct1",
  userId: "u1",
  email: "amanda@example.com",
  displayName: "Amanda",
  accountType: "work",
  accessToken: "graph-token",
  refreshToken: "refresh",
  tokenExpiry: new Date(Date.now() + 3600_000),
};
let settings: any = undefined;

const now = Date.now();
const iso = (msAgo: number) => new Date(now - msAgo).toISOString();
const GRAPH_ID = "AAMkAGI2TG93AAA=" + "x".repeat(120);
const inbox = [
  { id: GRAPH_ID, subject: "Deal memo", from: { emailAddress: { name: "Jane Doe", address: "jane@acme.com" } }, receivedDateTime: iso(3600_000), bodyPreview: "Here is the memo", isRead: false, hasAttachments: true, importance: "high", inferenceClassification: "focused" },
  { id: "id-2", subject: "Your receipt", from: { emailAddress: { name: "Shop", address: "receipts@shop.com" } }, receivedDateTime: iso(7200_000), bodyPreview: "Thanks", isRead: true, hasAttachments: false, importance: "normal", inferenceClassification: "focused" },
  { id: "id-3", subject: "Weekly digest", from: { emailAddress: { name: "News", address: "news@site.com" } }, receivedDateTime: iso(9000_000), bodyPreview: "This week", isRead: true, hasAttachments: false, importance: "normal", inferenceClassification: "other" },
];
const graphCalls: { method: string; url: string; body?: any }[] = [];

let server: Server;
let url: URL;
const realFetch = globalThis.fetch;

before(async () => {
  Object.assign(storage, {
    getUserByMcpApiKey: async (key: string) => (key === API_KEY ? user : undefined),
    getOAuthTokenByAccessToken: async () => undefined,
    getAccountsByUserId: async () => [account],
    getAccount: async () => account,
    getConnectorSettings: async () => settings,
  });

  globalThis.fetch = (async (input: any, init: any = {}) => {
    const target = String(input instanceof Request ? input.url : input);
    if (!target.startsWith("https://graph.microsoft.com/")) return realFetch(input, init);
    const method = init.method || "GET";
    graphCalls.push({ method, url: target, body: init.body ? JSON.parse(init.body) : undefined });
    const json = (data: any) => new Response(JSON.stringify(data), { status: 200, headers: { "Content-Type": "application/json" } });
    const path = decodeURIComponent(new URL(target).pathname);
    if (path.endsWith("/mailFolders/inbox/messages")) return json({ value: inbox });
    if (path.endsWith("/createReply")) return json({ id: "draft1", subject: "RE: Deal memo", webLink: "https://outlook/draft1" });
    if (path.startsWith("/v1.0/me/messages/")) {
      return json({
        ...inbox[0],
        toRecipients: [{ emailAddress: { name: "Amanda", address: "amanda@example.com" } }],
        uniqueBody: { contentType: "text", content: "New reply text only." },
        body: { contentType: "text", content: "New reply text only.\n\n> quoted history ..." },
      });
    }
    if (path.endsWith("/calendarView")) {
      return json({ value: [{ id: "evt-1", subject: "Lender call", start: { dateTime: "2026-09-24T15:00:00.0000000" }, end: { dateTime: "2026-09-24T15:30:00.0000000" }, isAllDay: false, showAs: "busy", isOnlineMeeting: true, organizer: { emailAddress: { name: "Amanda", address: "amanda@example.com" } }, location: { displayName: "" } }] });
    }
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

test("rejects requests without a token and points to OAuth discovery", async () => {
  const res = await realFetch(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: "{}" });
  assert.equal(res.status, 401);
  assert.match(res.headers.get("www-authenticate") || "", /oauth-protected-resource\/mcp\/v3/);
});

test("paid users get 10 tools with account info in the instructions", async () => {
  user = { ...baseUser };
  const client = await connect();
  const { tools } = await client.listTools();
  assert.deepEqual(tools.map((t) => t.name).sort(), [
    "calendar_find", "event_delete", "event_save", "file_read", "files_find",
    "mail_draft", "mail_find", "mail_read", "teams_find", "teams_read",
  ]);
  assert.match(client.getInstructions() || "", /amanda@example\.com \(work\)/);
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

test("mail_read returns only the new text, and mail_draft replies by handle", async () => {
  const client = await connect();
  const list = textOf(await client.callTool({ name: "mail_find", arguments: {} }));
  const handle = list.split("\n")[1].split("|")[0];

  const read = textOf(await client.callTool({ name: "mail_read", arguments: { ids: [handle] } }));
  assert.match(read, new RegExp(`^## ${handle} · Deal memo`));
  assert.match(read, /New reply text only\./);
  assert.ok(!read.includes("quoted history"));

  const draft = await client.callTool({ name: "mail_draft", arguments: { reply_to: handle, body: "Thanks Jane" } });
  assert.match(textOf(draft), /Draft saved, not sent/);
  const call = graphCalls.find((c) => c.url.includes("/createReply"));
  assert.ok(call && decodeURIComponent(call.url).includes(GRAPH_ID));
  assert.deepEqual(call!.body, { comment: "Thanks Jane" });
  await client.close();
});

test("allow-lists are enforced on direct reads", async () => {
  const client = await connect();
  const list = textOf(await client.callTool({ name: "mail_find", arguments: {} }));
  const handle = list.split("\n")[1].split("|")[0];
  settings = { focusedOnly: false, approvedSenders: ["@other.com"], allowedOnedriveFolders: [] };
  try {
    const read = textOf(await client.callTool({ name: "mail_read", arguments: { ids: [handle] } }));
    assert.match(read, /outside your approved-sender allow-list/);
    assert.ok(!read.includes("New reply text only"));
  } finally {
    settings = undefined;
  }
  await client.close();
});

test("expired handles give a clear message", async () => {
  const client = await connect();
  const res = await client.callTool({ name: "mail_read", arguments: { ids: ["m9999"] } });
  assert.match(textOf(res), /m9999 has expired/);
  await client.close();
});

test("calendar_find renders compact event rows", async () => {
  const client = await connect();
  const text = textOf(await client.callTool({ name: "calendar_find", arguments: {} }));
  assert.equal(text.split("\n")[0], "id|when|subject|location|organizer|flags");
  assert.match(text, /^e\d+\|Thu 09-24 15:00-15:30\|Lender call\|\|me\|O$/m);
  await client.close();
});
