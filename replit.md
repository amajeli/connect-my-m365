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
The app exposes an MCP server at `/mcp` that Claude.ai can connect to as a custom connector.

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
2. Go to Billing page, copy the MCP Server URL
3. In Claude.ai Settings > Connectors, click "Add custom connector"
4. Paste the MCP Server URL - Claude handles OAuth authentication automatically

## Project Structure
- `shared/schema.ts` - Database schema (users, connected_accounts, oauth_clients, oauth_codes, oauth_tokens)
- `server/routes.ts` - API endpoints (auth, accounts, emails, calendar, Stripe checkout/billing)
- `server/oauth.ts` - OAuth 2.0 provider (discovery, DCR, authorize, token endpoints)
- `server/mcp.ts` - MCP server with per-user scoping and OAuth/API key auth
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
- `client/src/pages/billing.tsx` - Subscription management and MCP setup instructions
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
- `GET /.well-known/oauth-authorization-server` - OAuth discovery metadata
- `POST /oauth/register` - Dynamic Client Registration (RFC 7591)
- `GET /oauth/authorize` - Authorization endpoint (redirects to Microsoft login if needed)
- `POST /oauth/token` - Token endpoint (authorization_code + refresh_token grants)

### MCP
- `POST /mcp` - MCP Streamable HTTP endpoint (requires OAuth Bearer token or legacy API key + active subscription)
- `GET /mcp` - MCP SSE notifications endpoint
- `DELETE /mcp` - MCP session termination

## Environment Variables
- `MICROSOFT_CLIENT_ID` - Azure AD app client ID
- `MICROSOFT_CLIENT_SECRET` - Azure AD app client secret
- `DATABASE_URL` - PostgreSQL connection string
- `SESSION_SECRET` - Express session secret
- Stripe credentials managed via Replit connector (stripe-replit-sync)

## Setup Requirements
Azure AD app registration must have:
- Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
- Redirect URI: `https://<domain>/api/auth/callback`
- API permissions: Mail.Read, Calendars.Read, User.Read, Files.Read, Chat.Read, offline_access

## Important Notes
- Never create tables in the `stripe` schema — managed by stripe-replit-sync
- Query Stripe data from `stripe.products`, `stripe.prices`, `stripe.subscriptions` etc.
- Create Stripe products via API (seed script), not SQL
- Webhook route must be registered BEFORE express.json() middleware
- Session cookie maxAge is 7 days
