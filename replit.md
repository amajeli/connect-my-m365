# Outlook Connector for Claude.ai

## Overview
A web application that connects to Microsoft Outlook (both work and personal accounts) via Microsoft Graph API, pulling emails and calendar events. Users can view their data and copy it formatted for use in Claude.ai. Also functions as a **custom MCP connector** that can be added directly to Claude.ai for seamless integration.

## Architecture
- **Frontend**: React + Vite + Tailwind CSS + shadcn/ui components
- **Backend**: Express.js with Microsoft Graph API integration
- **MCP Server**: Streamable HTTP transport for Claude.ai custom connector integration
- **Database**: PostgreSQL (Drizzle ORM) for storing connected account tokens
- **Auth**: Microsoft OAuth2 (authorization code flow) supporting both organizational and personal Microsoft accounts

## Key Features
- Connect multiple Microsoft accounts (work + personal)
- View inbox emails with search and account filtering
- View calendar events with weekly navigation
- Copy formatted email/calendar data to clipboard for Claude.ai
- **MCP custom connector** - Claude.ai can directly access Outlook data via MCP protocol
- Token auto-refresh for uninterrupted access
- Dark/light theme toggle

## MCP Connector (Claude.ai Integration)
The app exposes an MCP server at `/mcp` that Claude.ai can connect to as a custom connector.

### MCP Tools Available
- `list_accounts` - List all connected Microsoft Outlook accounts
- `get_recent_emails` - Fetch recent emails with optional account filter and limit
- `search_emails` - Search emails by keyword across subject, body, and sender
- `get_email_detail` - Get the full content of a specific email by ID
- `get_calendar_events` - Fetch calendar events with configurable date range

### Setting Up in Claude.ai
1. Publish the app to get a public URL
2. Go to Claude.ai Settings > Connectors
3. Click "Add custom connector"
4. Enter the MCP server URL: `https://<your-domain>/mcp`
5. (Optional) If MCP_API_KEY is set, add it as a Bearer token in Advanced settings

### MCP Authentication
- If `MCP_API_KEY` environment variable is set, all MCP requests require a Bearer token
- If not set, MCP endpoint is open (suitable for personal use)

## Project Structure
- `shared/schema.ts` - Database schema and TypeScript types
- `server/routes.ts` - API endpoints
- `server/mcp.ts` - MCP server with Streamable HTTP transport and tools
- `server/microsoft.ts` - Microsoft Graph API integration (OAuth, token management, API calls)
- `server/storage.ts` - Database storage layer
- `server/db.ts` - Database connection
- `server/index.ts` - Express app setup with session and MCP registration
- `client/src/pages/inbox.tsx` - Email inbox view
- `client/src/pages/calendar-page.tsx` - Calendar view
- `client/src/pages/accounts.tsx` - Account management
- `client/src/components/app-sidebar.tsx` - Navigation sidebar

## API Routes
- `GET /api/auth/microsoft` - Start OAuth flow
- `GET /api/auth/callback` - OAuth callback handler
- `GET /api/accounts` - List connected accounts
- `DELETE /api/accounts/:id` - Remove connected account
- `GET /api/emails` - Fetch emails from all accounts
- `GET /api/calendar` - Fetch calendar events from all accounts
- `POST /mcp` - MCP Streamable HTTP endpoint (initialize + tool calls)
- `GET /mcp` - MCP SSE notifications endpoint
- `DELETE /mcp` - MCP session termination

## Environment Variables
- `MICROSOFT_CLIENT_ID` - Azure AD app client ID
- `MICROSOFT_CLIENT_SECRET` - Azure AD app client secret
- `DATABASE_URL` - PostgreSQL connection string
- `SESSION_SECRET` - Express session secret
- `MCP_API_KEY` - (Optional) Bearer token for MCP endpoint authentication

## Setup Requirements
Azure AD app registration must have:
- Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
- Redirect URI: `https://<domain>/api/auth/callback`
- API permissions: Mail.Read, Calendars.Read, User.Read, offline_access
