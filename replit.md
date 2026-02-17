# Outlook Connector for Claude.ai

## Overview
A web application that connects to Microsoft Outlook (both work and personal accounts) via Microsoft Graph API, pulling emails and calendar events. Users can view their data and copy it formatted for use in Claude.ai.

## Architecture
- **Frontend**: React + Vite + Tailwind CSS + shadcn/ui components
- **Backend**: Express.js with Microsoft Graph API integration
- **Database**: PostgreSQL (Drizzle ORM) for storing connected account tokens
- **Auth**: Microsoft OAuth2 (authorization code flow) supporting both organizational and personal Microsoft accounts

## Key Features
- Connect multiple Microsoft accounts (work + personal)
- View inbox emails with search and account filtering
- View calendar events with weekly navigation
- Copy formatted email/calendar data to clipboard for Claude.ai
- Token auto-refresh for uninterrupted access
- Dark/light theme toggle

## Project Structure
- `shared/schema.ts` - Database schema and TypeScript types
- `server/routes.ts` - API endpoints
- `server/microsoft.ts` - Microsoft Graph API integration (OAuth, token management, API calls)
- `server/storage.ts` - Database storage layer
- `server/db.ts` - Database connection
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

## Environment Variables
- `MICROSOFT_CLIENT_ID` - Azure AD app client ID
- `MICROSOFT_CLIENT_SECRET` - Azure AD app client secret
- `DATABASE_URL` - PostgreSQL connection string
- `SESSION_SECRET` - Express session secret

## Setup Requirements
Azure AD app registration must have:
- Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
- Redirect URI: `https://<domain>/api/auth/callback`
- API permissions: Mail.Read, Calendars.Read, User.Read, offline_access
