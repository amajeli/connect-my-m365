# Microsoft 365 Connector for Claude.ai

## Overview
A multi-user SaaS web application that connects to Microsoft 365 (work and personal accounts) via Microsoft Graph API, pulling emails, calendar events, OneDrive files, and Teams chats. Functions as a **custom MCP connector** for Claude.ai with per-user subscription monetization via Stripe. Not affiliated with or endorsed by Microsoft Corporation.

## Architecture
- **Frontend**: React + Vite + Tailwind CSS + shadcn/ui components
- **Backend**: Express.js with Microsoft Graph API integration
- **MCP Server**: Streamable HTTP transport for Claude.ai custom connector integration (per-user scoped)
- **Database**: PostgreSQL (Drizzle ORM) for users, connected accounts, and tokens
- **Auth**: Microsoft OAuth2 doubles as both sign-in and account connection
- **Payments**: Stripe integration via stripe-replit-sync for subscriptions
- **Stripe Schema**: Managed automatically by stripe-replit-sync (never create/modify tables in stripe schema)

## Key Features
- **Multi-user system**: Each user signs in with Microsoft, gets their own isolated data
- Connect multiple Microsoft accounts per user (work + personal)
- View inbox emails with search and account filtering
- View calendar events with weekly navigation
- Copy formatted email/calendar data to clipboard for Claude.ai
- **MCP custom connector** - Claude.ai can directly access user's Microsoft 365 data via MCP protocol
- **Per-user MCP API keys** for secure, isolated connector access
- **Tiered subscription** model for MCP connector access
- Token auto-refresh for uninterrupted access
- Dark/light theme toggle
- Landing page with pricing

## Subscription Tiers
- **trialing**: 10-day trial started at Stripe checkout, full access to all MCP tools
- **active**: Paid subscriber, full access to all MCP tools
- **admin**: Internal/owner tier, full access, no expiry; granted via `POST /api/subscription/admin`; only eligible if user email is in ADMIN_EMAIL env var
- **free**: Post-trial fallback, email-only access (list_accounts + get_recent_emails), 20 emails/day cap (in-memory counter, resets midnight UTC)
- **none**: No access to MCP tools; must subscribe or activate free tier

## Auth Flow
1. User visits landing page (unauthenticated)
2. Clicks "Sign in with Microsoft" → Microsoft OAuth2
3. On callback: user record created/updated, first Outlook account auto-connected, session established
4. User can add more Outlook accounts via "Add Account" (uses `?connect=true` flag to only connect, not re-auth)
5. User subscribes via Stripe checkout to activate MCP connector

## MCP Connector (Claude.ai Integration)
The app exposes an MCP server at `/mcp` (v1) that Claude.ai can connect to as a custom connector. A second, token-efficient server is exposed at `/mcp/v2` (see "MCP v2" below). Both run side by side and share authentication and subscription gating; v1 is unchanged.

### MCP Tools Available
- `list_accounts` - List all connected Microsoft Outlook accounts (user-scoped)
- `get_recent_emails` - Fetch recent emails with optional account filter and limit
- `search_emails` - Search emails by keyword across subject, body, and sender
- `get_email_detail` - Get the full content of a specific email by ID
- `get_calendar_events` - Fetch calendar events with configurable date range
- `list_onedrive_files` - Browse files/folders in OneDrive at any path (account-scoped)
- `search_onedrive_files` - Search OneDrive by filename/keyword
- `get_onedrive_file_content` - Read text file content (≤512 KB) or get metadata + web link for other types
- `list_teams_chats` - List recent Teams chats with participants (work accounts only)
- `get_teams_chat_messages` - Read messages from a specific Teams chat by ID
- `search_teams_messages` - Search across Teams chats by keyword

### MCP Authentication
- **Primary: OAuth 2.0** - Claude.ai uses Dynamic Client Registration (RFC 7591) and OAuth authorization code flow
  - Discovery endpoint: `/.well-known/oauth-authorization-server`
  - Dynamic client registration: `POST /oauth/register`
  - Authorization: `GET /oauth/authorize` (redirects to Microsoft login if not authenticated)
  - Token exchange: `POST /oauth/token` (supports authorization_code and refresh_token grants)
  - PKCE (S256) supported for enhanced security
- **Fallback: Legacy API keys** - Per-user MCP API keys still work for manual testing
- User must have an active subscription (status: "active" or "trialing") to use MCP

### Setting Up in Claude.ai
1. Subscribe to activate the connector
2. Go to Billing page, copy the MCP Server URL (v1 `/mcp` or v2 `/mcp/v2`)
3. In Claude.ai Settings > Connectors, click "Add custom connector"
4. Paste the MCP Server URL - Claude handles OAuth authentication automatically

## MCP v2 (token-efficient connector at `/mcp/v2`)
A separate MCP server (`server/mcpV2.ts`, name `microsoft365-connector-v2`) built for minimal token usage. v1 at `/mcp` is left fully intact. v2 reuses v1's auth, subscription/tier gating, and free-tier email counter (exported from `server/mcp.ts`).

### v2 differences vs v1
- **Lean emails + on-demand bodies**: `get_recent_emails` / `search_emails` return compact metadata (subject, sender, date, preview) plus a `resource` URI per email — never full bodies. The body is loaded only when needed via the `outlook-email://{account}/{messageId}` resource or the `get_email` tool. Default email limit is 5.
- **Server-side filters**: email tools accept `after_datetime` / `before_datetime` (accept `today`/`yesterday`/`tomorrow`, ISO date, or ISO datetime) and `sender`.
- **`outlook_calendar_search`**: targeted calendar lookups by datetime window (in addition to the v1-style `get_calendar_events` overview).
- **`create_email_draft`**: creates a DRAFT only (saved to Drafts, never sends). Requires the `Mail.ReadWrite` Graph scope.
- **Per-user allow-lists (Connector Settings)** applied automatically:
  - Focused-inbox-only toggle (`inferenceClassification eq 'focused'`).
  - Approved senders (full address `jane@acme.com` or domain `@acme.com`); empty = allow all.
  - Designated OneDrive folders; if any are set, all OneDrive tools are restricted to those folders and subfolders; empty = all of OneDrive.
- Resource registered via `registerResource` + `ResourceTemplate`; account/messageId are URL-encoded into the URI.

### v2 tools
`list_accounts`, `get_recent_emails`, `search_emails`, `get_email`, `create_email_draft`, `get_calendar_events`, `outlook_calendar_search`, `list_onedrive_files`, `search_onedrive_files`, `get_onedrive_file_content`, `list_teams_chats`, `get_teams_chat_messages`, `search_teams_messages` + resource `outlook-email://{account}/{messageId}`.

## MCP v3 (leanest connector at `/mcp/v3`, recommended)
`server/mcpV3.ts` (name `microsoft365-connector-v3`). Built on v2's filters and allow-lists, which now live in `server/mcpShared.ts` (v2 imports them and behaves the same). v1 and v2 are unchanged. v3 is the only version that works reliably on serverless hosts (Netlify); see `NETLIFY.md`.

### v3 differences vs v2
- **Fewer, shorter tools** with `readOnlyHint`/`destructiveHint` annotations. Tool definitions are re-sent on every turn.
- **Pipe-delimited rows instead of JSON.** The first row names the columns once.
- **Short ids** (`m1`, `e1`, `f1`, `c1`) replace ~150-char Graph IDs. They're stored in the `mcp_handles` table (`server/handleStore.ts`), so they survive cold starts and work across parallel instances. Allocation is serialized per user with a Postgres advisory lock. Old ids are pruned after 30 days, but numbers are never reused. Without `DATABASE_URL` an in-memory store is used (tests).
- **Multi-account first**: every read tool queries all connected accounts in parallel and labels results by account.
- **`mail_read` returns Graph `uniqueBody`**: only the new text of each message, not the quoted thread (`full=true` includes it). Long tracking URLs become `[link:domain]`.
- **`mail_find`** handles both the inbox briefing (last 18h, noise hidden, reports how many were hidden) and mailbox-wide KQL search.
- **`mail_draft` requires `account`**, and the tool description and server instructions tell Claude to ask the user which account. It never sends. A reply can be drafted from a different account than the one that received it; this creates a new draft to the sender with the original quoted.
- **`calendar_find`** flags double bookings (`C`). **`calendar_conflicts`** lists every overlap across all calendars. Free, cancelled and declined events don't count, and the same meeting on two calendars (same subject and times) is treated as a duplicate, not a conflict.
- **`teams_send`** posts to an existing chat (c-id) or starts or reuses a chat with people (`people` emails). It sends immediately, so the description tells Claude to confirm the text first. Requires the `Chat.ReadWrite` scope; existing accounts must reconnect once.
- **Stateless Streamable HTTP**: a new server per POST and no session map. GET/DELETE return 405.
- **Free tier sees only `mail_find`**.
- A 401 includes `WWW-Authenticate` pointing to `/.well-known/oauth-protected-resource/mcp/v3`.

### v3 tools
`mail_find`, `mail_read`, `mail_draft`, `calendar_find`, `calendar_conflicts`, `event_save`, `event_delete`, `files_find`, `file_read`, `teams_find`, `teams_read`, `teams_send`. Accounts, active settings and the draft/send rules go in the server `instructions` at initialize.

### Connector Settings UI
`/settings` page (`client/src/pages/settings.tsx`): Focused-only switch, approved-sender badge editor, and OneDrive folder allow-list with a live folder browser. Backed by `GET/PUT /api/settings` and `GET /api/onedrive/browse`.

## Project Structure
- `shared/schema.ts` - Database schema (users, connected_accounts, oauth_clients, oauth_codes, oauth_tokens)
- `server/routes.ts` - API endpoints (auth, accounts, emails, calendar, Stripe checkout/billing)
- `server/oauth.ts` - OAuth 2.0 provider (discovery, DCR, authorize, token endpoints)
- `server/mcp.ts` - MCP v1 server with per-user scoping and OAuth/API key auth (exports shared auth/gating helpers)
- `server/mcpV2.ts` - MCP v2 server (token-efficient, allow-lists, drafts) at `/mcp/v2`
- `server/mcpV3.ts` - MCP v3 server (leanest, stateless) at `/mcp/v3`; `mcpV3Format.ts` (rows, handles), `graphV3.ts` (Graph calls)
- `server/mcpShared.ts` - filters and allow-list checks shared by v2 and v3
- `server/handleStore.ts` - database-backed short ids for v3
- `server/app.ts` - Express app factory used by `server/index.ts` (long-running server) and `netlify/functions/api.ts` (Netlify Function)
- `netlify.toml`, `script/build-netlify.ts`, `NETLIFY.md` - Netlify build and deploy guide
- `server/__tests__/` - `npm test` (node:test via tsx; v3 end-to-end with stubbed storage and Graph)
- `server/microsoft.ts` - Microsoft Graph API integration (OAuth, token management, API calls)
- `server/storage.ts` - Database storage layer + Stripe schema queries + OAuth storage
- `server/stripeClient.ts` - Stripe client via Replit connector credentials
- `server/webhookHandlers.ts` - Stripe webhook processing
- `server/seed-products.ts` - Script to create Stripe products/prices
- `server/db.ts` - Database connection
- `server/index.ts` - Express app setup with Stripe init, webhook route, session, MCP registration
- `client/src/lib/auth.tsx` - Auth context provider (useAuth hook)
- `client/src/pages/landing.tsx` - Public landing page with pricing
- `client/src/pages/inbox.tsx` - Email inbox view
- `client/src/pages/calendar-page.tsx` - Calendar view
- `client/src/pages/accounts.tsx` - Account management
- `client/src/pages/billing.tsx` - Subscription management and MCP setup instructions (v1 + v2 URLs)
- `client/src/pages/settings.tsx` - Connector Settings (focused-only, approved senders, OneDrive folder allow-list)
- `client/src/components/app-sidebar.tsx` - Navigation sidebar with user menu

## Database Schema
### users table
- id (varchar, PK, UUID)
- email (text, unique)
- displayName (text)
- microsoftId (text, unique) - Microsoft user ID for auth
- mcpApiKey (text, UUID) - per-user MCP bearer token
- stripeCustomerId (text, nullable)
- stripeSubscriptionId (text, nullable)
- subscriptionStatus (text, default "none") — "none" | "trialing" | "active" | "free" | "admin" | "past_due" | "canceled"
- createdAt (timestamp)

> UserInfo (returned by /api/auth/me) also includes: isAdmin (bool)

### connected_accounts table
- id (varchar, PK, UUID)
- userId (varchar, FK → users.id, cascade delete)
- email, displayName, accountType, accessToken, refreshToken, tokenExpiry, isActive, connectedAt

### oauth_clients table (Dynamic Client Registration)
- id (varchar, PK, UUID)
- clientId (text, unique), clientSecret (text), clientName (text), redirectUris (text[]), createdAt

### oauth_codes table (Authorization codes)
- id (varchar, PK, UUID)
- code (text, unique), clientId, userId (FK), redirectUri, scope, codeChallenge, codeChallengeMethod, expiresAt, used, createdAt

### oauth_tokens table (Access/refresh tokens)
- id (varchar, PK, UUID)
- accessToken (text, unique), refreshToken (text, unique), clientId, userId (FK), scope, expiresAt, createdAt

### mcp_handles table (v3 short ids)
- PK (userId, handle); unique (userId, refKey)
- handle (e.g. "m12"), kind (m/e/f/c), num, refKey (hash of account + Graph id), accountId, accountEmail, ref (Graph id), createdAt
- Created by `npm run db:migrate` (`script/migrate.ts`, idempotent)

### connector_settings table (per-user v2 allow-lists)
- id (varchar, PK, UUID)
- userId (varchar, unique FK → users.id, cascade delete)
- focusedOnly (bool, default false)
- approvedSenders (text[], default {})
- allowedOnedriveFolders (text[], default {})
- updatedAt (timestamp)

## API Routes
### Auth
- `GET /api/auth/me` - Get current user (returns null if not authenticated)
- `GET /api/auth/microsoft` - Start OAuth flow (sign-in or connect with ?connect=true)
- `GET /api/auth/callback` - OAuth callback handler
- `POST /api/auth/logout` - Destroy session

### Data (requires auth)
- `GET /api/accounts` - List user's connected accounts
- `DELETE /api/accounts/:id` - Remove connected account (verified ownership)
- `GET /api/emails` - Fetch emails from user's accounts
- `GET /api/calendar` - Fetch calendar events from user's accounts

### Stripe / Billing (requires auth)
- `GET /api/products-with-prices` - List products with prices (public)
- `GET /api/stripe/publishable-key` - Get Stripe publishable key
- `POST /api/checkout` - Create Stripe checkout session
- `POST /api/billing/portal` - Create Stripe billing portal session
- `GET /api/subscription` - Get user's subscription status
- `POST /api/subscription/sync` - Sync subscription from Stripe API
- `POST /api/mcp-key/regenerate` - Regenerate MCP API key
- `POST /api/stripe/webhook` - Stripe webhook handler (raw body, before express.json())

### OAuth (for Claude.ai MCP connector)
Security rules (codes are issued without a consent screen, so these are required):
- `redirect_uri` must be https on `claude.ai` / `claude.com`, loopback http (`localhost`, `127.0.0.1`), or a host listed in `OAUTH_ALLOWED_REDIRECT_HOSTS`. This is checked at registration and at every authorize, including for older or self-healed clients.
- PKCE S256 is mandatory.

- `GET /.well-known/oauth-authorization-server` - OAuth discovery metadata
- `POST /oauth/register` - Dynamic Client Registration (RFC 7591)
- `GET /oauth/authorize` - Authorization endpoint (redirects to Microsoft login if needed)
- `POST /oauth/token` - Token endpoint (authorization_code + refresh_token grants)

### MCP
- `POST /mcp` - MCP v1 Streamable HTTP endpoint (requires OAuth Bearer token or legacy API key + active subscription)
- `POST /mcp/v2`, `GET /mcp/v2`, `DELETE /mcp/v2` - MCP v2 endpoint (same auth/gating as v1)
- `GET /mcp` - MCP SSE notifications endpoint
- `DELETE /mcp` - MCP session termination

## Environment Variables
- `MICROSOFT_CLIENT_ID` - Azure AD app client ID
- `MICROSOFT_CLIENT_SECRET` - Azure AD app client secret
- `DATABASE_URL` - PostgreSQL connection string
- `SESSION_SECRET` - Express session secret
- `OAUTH_ALLOWED_REDIRECT_HOSTS` - optional, comma-separated extra hosts allowed as OAuth redirect targets
- `EMAIL_SENDER_BLOCKLIST` - optional, comma-separated override for the automated-sender noise filter (v2/v3)
- Stripe credentials managed via Replit connector (stripe-replit-sync)

## Setup Requirements
Azure AD app registration must have:
- Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
- Redirect URI: `https://<domain>/api/auth/callback`
- API permissions: Mail.Read, Mail.ReadWrite (for drafts), Calendars.ReadWrite, User.Read, Files.Read, Chat.ReadWrite (for v3 teams_send), offline_access
- NOTE: `Mail.ReadWrite` was added to the requested scopes — existing connected accounts must re-consent (reconnect) before `create_email_draft` will work.

## Important Notes
- Never create tables in the `stripe` schema — managed by stripe-replit-sync
- Query Stripe data from `stripe.products`, `stripe.prices`, `stripe.subscriptions` etc.
- Create Stripe products via API (seed script), not SQL
- Webhook route must be registered BEFORE express.json() middleware
- Session cookie maxAge is 7 days
