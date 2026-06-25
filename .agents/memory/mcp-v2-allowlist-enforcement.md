---
name: MCP v2 connector allow-list enforcement
description: Per-user content allow-lists must be enforced at every read path, not just list/search; v2 is a separate endpoint from v1.
---

# MCP v2 connector — allow-list enforcement

The token-efficient connector lives at a SEPARATE endpoint (`/mcp/v2`, `server/mcpV2.ts`) and v1 (`/mcp`, `server/mcp.ts`) must stay byte-for-byte behavior-identical — v2 only *imports* exported helpers from v1.

**Rule:** Per-user content allow-lists (approved senders, focused-inbox-only) must be enforced on EVERY path that returns content — not just the list/search tools.
**Why:** A direct-ID fetch (`get_email`) or a resource read (`outlook-email://{account}/{messageId}`) can return a disallowed email's body even when list/search correctly filter it out, because a message ID can be known from prior context. A code review caught this bypass.
**How to apply:** Any new tool/resource that returns email (or OneDrive) content must re-check the user's `connector_settings` after fetching and deny if outside the allow-list. For focused-only checks, `getEmailById` must `$select` `inferenceClassification`. OneDrive tools must constrain to `allowedOnedriveFolders` (folder + subfolders) on every tool, not a subset.
