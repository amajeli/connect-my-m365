# Deploying to Netlify

The React client is served from Netlify's CDN, and the Express API (sign-in,
OAuth, MCP, billing) runs as one Netlify Function. `netlify.toml` wires it up;
`npm run build:netlify` builds both.

**Use the v3 connector URL (`/mcp/v3`) on Netlify.** v3 keeps nothing in memory
between requests. v1 (`/mcp`) and v2 (`/mcp/v2`) keep MCP sessions in memory,
which serverless functions don't preserve, so they will disconnect at random.

## 1. Database

Netlify doesn't host Postgres for this app. Use any hosted Postgres (Neon,
Supabase, or the existing Replit database if it's reachable from outside
Replit). Then create or upgrade the tables. This is safe to re-run:

```bash
DATABASE_URL="postgres://..." npm run db:migrate
```

This adds the `mcp_handles` table v3 needs. It stores the short ids like `m1`
and `e1`.

## 2. Netlify site

1. **Add new site → Import from Git**, pick this repo and branch.
2. The build command and publish directory come from `netlify.toml`, so
   leave those fields as they are.
3. **Site configuration → Environment variables**:

| Variable | Value |
|---|---|
| `DATABASE_URL` | Postgres connection string (use the pooled one if your host offers it) |
| `SESSION_SECRET` | A long random string |
| `MICROSOFT_CLIENT_ID` | Azure app (client) ID |
| `MICROSOFT_CLIENT_SECRET` | Azure client secret |
| `STRIPE_SECRET_KEY`, `STRIPE_PUBLISHABLE_KEY` | Only if billing is used (see below) |
| `ADMIN_EMAIL` | Your email, for the admin tier (full access, no billing) |
| `OAUTH_ALLOWED_REDIRECT_HOSTS` | Optional. Extra MCP client hosts besides claude.ai / claude.com |

## 3. Azure app registration

In the Azure portal, under **App registrations → your app**:

- **Authentication → Redirect URIs**: add
  `https://<your-site>.netlify.app/api/auth/callback`, and the same with your
  custom domain if you add one.
- **API permissions (delegated)**: `Mail.Read`, `Mail.ReadWrite`,
  `Calendars.ReadWrite`, `Files.Read`, **`Chat.ReadWrite`** (new, for sending
  Teams messages), `User.Read`, `offline_access`.

Existing connected accounts need to **reconnect once** (Accounts page → remove
and add) to grant `Chat.ReadWrite`. Until they do, reading Teams still works
but sending fails with a permissions error.

## 4. Connect Claude

In Claude: **Settings → Connectors → Add custom connector** →
`https://<your-site>.netlify.app/mcp/v3`.

Claude caches OAuth registrations per URL, so the new domain signs in fresh.
The old Replit connector keeps working until you remove it.

## Billing on Netlify

Checkout and the customer portal work when the Stripe keys are set. The
Replit-only Stripe sync (`stripe-replit-sync`) doesn't run on Netlify, so
subscription changes from Stripe webhooks aren't synced yet. Until that's
replaced, set your own account's tier directly (admin tier: your email in
`ADMIN_EMAIL`).

## Limits to know

- Each request can run up to 60 seconds. v3 queries all accounts in parallel,
  so normal calls take a second or two.
- The free-tier daily email counter lives in memory, so on serverless it
  undercounts. This only matters once there are free-tier users.
