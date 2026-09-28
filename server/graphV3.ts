// Microsoft Graph calls used only by the v3 MCP server. Each one selects just
// the fields v3 renders.

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graph(accessToken: string, path: string, init: RequestInit = {}, prefer?: string): Promise<any> {
  const headers: Record<string, string> = { Authorization: `Bearer ${accessToken}` };
  if (init.body) headers["Content-Type"] = "application/json";
  if (prefer) headers.Prefer = prefer;
  const response = await fetch(`${GRAPH}${path}`, { ...init, headers: { ...headers, ...(init.headers as any) } });
  if (!response.ok) {
    const text = await response.text();
    let message = text;
    try {
      message = JSON.parse(text).error?.message || text;
    } catch {}
    throw new Error(`Graph ${response.status}: ${message.slice(0, 300)}`);
  }
  if (response.status === 204) return null;
  return response.json();
}

const MESSAGE_LIST_FIELDS =
  "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,importance,inferenceClassification";

export async function listInboxMessages(
  accessToken: string,
  opts: { top: number; after?: string; before?: string; from?: string; unreadOnly?: boolean; focusedOnly?: boolean },
): Promise<any[]> {
  const filters: string[] = [];
  if (opts.after) filters.push(`receivedDateTime ge ${opts.after}`);
  if (opts.before) filters.push(`receivedDateTime lt ${opts.before}`);
  if (opts.from) filters.push(`from/emailAddress/address eq '${opts.from.replace(/'/g, "''")}'`);
  if (opts.unreadOnly) filters.push("isRead eq false");
  if (opts.focusedOnly) filters.push("inferenceClassification eq 'focused'");

  const params = new URLSearchParams({ $select: MESSAGE_LIST_FIELDS, $top: String(opts.top) });
  if (filters.length) params.set("$filter", filters.join(" and "));
  // Graph rejects $orderby combined with a from/ filter.
  if (!opts.from) params.set("$orderby", "receivedDateTime DESC");

  const data = await graph(accessToken, `/me/mailFolders/inbox/messages?${params}`);
  return data.value || [];
}

/** Mailbox-wide KQL search (all folders). Graph doesn't allow $filter or $orderby with $search. */
export async function searchMessages(accessToken: string, kql: string, top: number): Promise<any[]> {
  const params = new URLSearchParams({ $select: MESSAGE_LIST_FIELDS, $top: String(top), $search: `"${kql}"` });
  const data = await graph(accessToken, `/me/messages?${params}`);
  return data.value || [];
}

/**
 * Loads one message as plain text. `uniqueBody` holds only this message's new
 * content, without the quoted thread below it; that is usually a fraction of the size.
 * `body` is also selected as a fallback for messages where uniqueBody comes back empty.
 */
export async function getMessageText(accessToken: string, id: string): Promise<any> {
  const fields =
    "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,importance,inferenceClassification,uniqueBody,body";
  return graph(
    accessToken,
    `/me/messages/${encodeURIComponent(id)}?$select=${fields}`,
    {},
    'outlook.body-content-type="text"',
  );
}

/** Creates a reply (or reply-all) draft with `comment` above the quoted thread. Never sends. */
export async function createReplyDraft(accessToken: string, id: string, comment: string, replyAll: boolean): Promise<any> {
  const action = replyAll ? "createReplyAll" : "createReply";
  return graph(accessToken, `/me/messages/${encodeURIComponent(id)}/${action}`, {
    method: "POST",
    body: JSON.stringify({ comment }),
  });
}

export async function getCalendarView(accessToken: string, start: string, end: string, top: number): Promise<any[]> {
  const params = new URLSearchParams({
    startDateTime: start,
    endDateTime: end,
    $select: "id,subject,start,end,location,organizer,isAllDay,showAs,isCancelled,responseStatus,isOnlineMeeting,onlineMeeting,onlineMeetingUrl",
    $orderby: "start/dateTime",
    $top: String(top),
  });
  const data = await graph(accessToken, `/me/calendarView?${params}`, {}, 'outlook.timezone="UTC"');
  return data.value || [];
}

export async function createEvent(accessToken: string, event: any): Promise<any> {
  return graph(accessToken, "/me/events", { method: "POST", body: JSON.stringify(event) }, 'outlook.timezone="UTC"');
}

export async function updateEvent(accessToken: string, id: string, patch: any): Promise<any> {
  return graph(
    accessToken,
    `/me/events/${encodeURIComponent(id)}`,
    { method: "PATCH", body: JSON.stringify(patch) },
    'outlook.timezone="UTC"',
  );
}

export async function deleteEvent(accessToken: string, id: string): Promise<void> {
  await graph(accessToken, `/me/events/${encodeURIComponent(id)}`, { method: "DELETE" });
}

/** Posts a plain-text message to a Teams chat. Requires Chat.ReadWrite (or ChatMessage.Send). */
export async function sendChatMessage(accessToken: string, chatId: string, text: string): Promise<any> {
  return graph(accessToken, `/me/chats/${encodeURIComponent(chatId)}/messages`, {
    method: "POST",
    body: JSON.stringify({ body: { contentType: "text", content: text } }),
  });
}

/**
 * Opens a chat with people in the same organization: one-on-one for one
 * person (Graph returns the existing chat if there is one), a group chat otherwise.
 */
export async function createChat(accessToken: string, emails: string[], topic?: string): Promise<any> {
  const me = await graph(accessToken, "/me?$select=id");
  const member = (userRef: string) => ({
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    roles: ["owner"],
    "user@odata.bind": `${GRAPH}/users('${encodeURIComponent(userRef)}')`,
  });
  const chat: any = {
    chatType: emails.length === 1 ? "oneOnOne" : "group",
    members: [member(me.id), ...emails.map(member)],
  };
  if (topic && emails.length > 1) chat.topic = topic;
  return graph(accessToken, "/chats", { method: "POST", body: JSON.stringify(chat) });
}
