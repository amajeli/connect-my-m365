import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { randomUUID } from "crypto";
import {
  getAuthUrl,
  exchangeCodeForTokens,
  getUserProfile,
  getValidAccessToken,
  getEmails,
  getCalendarEvents,
} from "./microsoft";
import type { OutlookEmail, OutlookCalendarEvent } from "@shared/schema";

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {

  // Start Microsoft OAuth flow
  app.get("/api/auth/microsoft", (req, res) => {
    const state = randomUUID();
    req.session.oauthState = state;
    const authUrl = getAuthUrl(state);
    res.redirect(authUrl);
  });

  // OAuth callback
  app.get("/api/auth/callback", async (req, res) => {
    const { code, error, error_description, state } = req.query;

    if (error) {
      console.error("OAuth error:", error, error_description);
      return res.redirect("/?error=auth_failed");
    }

    if (!state || state !== req.session.oauthState) {
      console.error("OAuth state mismatch");
      return res.redirect("/?error=state_mismatch");
    }

    delete req.session.oauthState;

    if (!code || typeof code !== "string") {
      return res.redirect("/?error=no_code");
    }

    try {
      const tokens = await exchangeCodeForTokens(code);
      const profile = await getUserProfile(tokens.access_token);

      const email = profile.mail || profile.userPrincipalName || "";
      const displayName = profile.displayName || email;

      // Determine account type based on the email domain
      const isPersonal =
        email.endsWith("@outlook.com") ||
        email.endsWith("@hotmail.com") ||
        email.endsWith("@live.com") ||
        email.endsWith("@msn.com");
      const accountType = isPersonal ? "personal" : "work";

      const tokenExpiry = new Date(Date.now() + tokens.expires_in * 1000);

      // Check if account already exists, update tokens if so
      const existingAccount = await storage.getAccountByEmail(email);
      if (existingAccount) {
        await storage.updateAccountTokens(
          existingAccount.id,
          tokens.access_token,
          tokens.refresh_token,
          tokenExpiry
        );
      } else {
        await storage.createAccount({
          email,
          displayName,
          accountType,
          accessToken: tokens.access_token,
          refreshToken: tokens.refresh_token,
          tokenExpiry,
          isActive: true,
        });
      }

      res.redirect("/accounts?connected=true");
    } catch (err: any) {
      console.error("OAuth callback error:", err.message);
      res.redirect("/?error=token_exchange_failed");
    }
  });

  // Get all connected accounts (excluding sensitive token data)
  app.get("/api/accounts", async (_req, res) => {
    try {
      const accounts = await storage.getAccounts();
      const safeAccounts = accounts.map((a) => ({
        id: a.id,
        email: a.email,
        displayName: a.displayName,
        accountType: a.accountType,
        isActive: a.isActive,
        connectedAt: a.connectedAt.toISOString(),
      }));
      res.json(safeAccounts);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  // Delete a connected account
  app.delete("/api/accounts/:id", async (req, res) => {
    try {
      await storage.deleteAccount(req.params.id);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  // Get emails from all connected accounts
  app.get("/api/emails", async (_req, res) => {
    try {
      const accounts = await storage.getAccounts();
      if (accounts.length === 0) {
        return res.json([]);
      }

      const allEmails: OutlookEmail[] = [];

      for (const account of accounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const emails = await getEmails(accessToken);
          const enriched = emails.map((email: any) => ({
            ...email,
            accountId: account.id,
            accountEmail: account.email,
          }));
          allEmails.push(...enriched);
        } catch (err: any) {
          console.error(`Failed to fetch emails for ${account.email}:`, err.message);
        }
      }

      // Sort all emails by date
      allEmails.sort(
        (a, b) =>
          new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
      );

      res.json(allEmails);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  // Get calendar events from all connected accounts
  app.get("/api/calendar", async (_req, res) => {
    try {
      const accounts = await storage.getAccounts();
      if (accounts.length === 0) {
        return res.json([]);
      }

      const allEvents: OutlookCalendarEvent[] = [];

      for (const account of accounts) {
        try {
          const accessToken = await getValidAccessToken(account.id);
          const events = await getCalendarEvents(accessToken);
          const enriched = events.map((event: any) => ({
            ...event,
            accountId: account.id,
            accountEmail: account.email,
          }));
          allEvents.push(...enriched);
        } catch (err: any) {
          console.error(`Failed to fetch calendar for ${account.email}:`, err.message);
        }
      }

      // Sort events by start time
      allEvents.sort(
        (a, b) =>
          new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime()
      );

      res.json(allEvents);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  return httpServer;
}
