import type { Express, Request, Response, NextFunction } from "express";
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
  listDriveItems,
} from "./microsoft";
import { getUncachableStripeClient, getStripePublishableKey } from "./stripeClient";
import type { OutlookEmail, OutlookCalendarEvent } from "@shared/schema";
import { connectorSettingsUpdateSchema } from "@shared/schema";

function requireAuth(req: Request, res: Response, next: NextFunction) {
  if (!req.session.userId) {
    return res.status(401).json({ message: "Not authenticated" });
  }
  next();
}

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {

  app.get("/api/auth/me", async (req, res) => {
    if (!req.session.userId) {
      return res.json({ user: null });
    }
    try {
      const user = await storage.getUser(req.session.userId);
      if (!user) {
        req.session.destroy(() => {});
        return res.json({ user: null });
      }
      const adminEmails = (process.env.ADMIN_EMAIL || "").split(",").map(e => e.trim().toLowerCase()).filter(Boolean);
      const isAdmin = adminEmails.includes(user.email.toLowerCase());
      res.json({
        user: {
          id: user.id,
          email: user.email,
          displayName: user.displayName,
          subscriptionStatus: user.subscriptionStatus,
          stripeCustomerId: user.stripeCustomerId,
          stripeSubscriptionId: user.stripeSubscriptionId,
          isAdmin,
        },
      });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/auth/microsoft", (req, res) => {
    const state = randomUUID();
    req.session.oauthState = state;
    if (req.query.connect === "true") {
      if (!req.session.userId) {
        return res.redirect("/?error=not_authenticated");
      }
      req.session.connectOnly = true;
    }
    const requestHost = req.get('host');
    const authUrl = getAuthUrl(state, requestHost);
    req.session.save((err) => {
      if (err) console.error("Session save error before OAuth redirect:", err);
      res.redirect(authUrl);
    });
  });

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
    const connectOnly = req.session.connectOnly;
    delete req.session.connectOnly;

    if (!code || typeof code !== "string") {
      return res.redirect("/?error=no_code");
    }

    try {
      const requestHost = req.get('host');
      const tokens = await exchangeCodeForTokens(code, requestHost);
      const profile = await getUserProfile(tokens.access_token);

      const email = profile.mail || profile.userPrincipalName || "";
      const displayName = profile.displayName || email;
      const microsoftId = profile.id;

      const isPersonal =
        email.endsWith("@outlook.com") ||
        email.endsWith("@hotmail.com") ||
        email.endsWith("@live.com") ||
        email.endsWith("@msn.com");
      const accountType = isPersonal ? "personal" : "work";

      const tokenExpiry = new Date(Date.now() + tokens.expires_in * 1000);

      if (connectOnly && req.session.userId) {
        const userId = req.session.userId;
        const existingAccount = await storage.getAccountByEmailAndUser(email, userId);
        if (existingAccount) {
          await storage.updateAccountTokens(
            existingAccount.id,
            tokens.access_token,
            tokens.refresh_token,
            tokenExpiry
          );
        } else {
          await storage.createAccount({
            userId,
            email,
            displayName,
            accountType,
            accessToken: tokens.access_token,
            refreshToken: tokens.refresh_token,
            tokenExpiry,
            isActive: true,
          });
        }
        return res.redirect("/accounts?connected=true");
      }

      let user = await storage.getUserByMicrosoftId(microsoftId);
      if (!user) {
        user = await storage.createUser({
          email,
          displayName,
          microsoftId,
        });
      } else {
        await storage.updateUser(user.id, { email, displayName });
      }

      req.session.userId = user.id;

      const existingAccount = await storage.getAccountByEmailAndUser(email, user.id);
      if (existingAccount) {
        await storage.updateAccountTokens(
          existingAccount.id,
          tokens.access_token,
          tokens.refresh_token,
          tokenExpiry
        );
      } else {
        await storage.createAccount({
          userId: user.id,
          email,
          displayName,
          accountType,
          accessToken: tokens.access_token,
          refreshToken: tokens.refresh_token,
          tokenExpiry,
          isActive: true,
        });
      }

      if (req.session.pendingOAuth) {
        const pending = req.session.pendingOAuth;
        delete req.session.pendingOAuth;
        const params = new URLSearchParams();
        if (pending.client_id) params.set("client_id", pending.client_id);
        if (pending.redirect_uri) params.set("redirect_uri", pending.redirect_uri);
        if (pending.response_type) params.set("response_type", pending.response_type);
        if (pending.state) params.set("state", pending.state);
        if (pending.scope) params.set("scope", pending.scope);
        if (pending.code_challenge) params.set("code_challenge", pending.code_challenge);
        if (pending.code_challenge_method) params.set("code_challenge_method", pending.code_challenge_method);
        return res.redirect(`/oauth/authorize?${params.toString()}`);
      }

      res.redirect("/inbox");
    } catch (err: any) {
      console.error("OAuth callback error:", err.message);
      res.redirect("/?error=token_exchange_failed");
    }
  });

  app.post("/api/auth/logout", (req, res) => {
    req.session.destroy(() => {
      res.json({ success: true });
    });
  });

  app.get("/api/accounts", requireAuth, async (req, res) => {
    try {
      const accounts = await storage.getAccountsByUserId(req.session.userId!);
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

  app.delete("/api/accounts/:id", requireAuth, async (req, res) => {
    try {
      const account = await storage.getAccount(req.params.id);
      if (!account || account.userId !== req.session.userId) {
        return res.status(404).json({ message: "Account not found" });
      }
      await storage.deleteAccount(req.params.id);
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/emails", requireAuth, async (req, res) => {
    try {
      const accounts = await storage.getAccountsByUserId(req.session.userId!);
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

      allEmails.sort(
        (a, b) =>
          new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
      );

      res.json(allEmails);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/calendar", requireAuth, async (req, res) => {
    try {
      const accounts = await storage.getAccountsByUserId(req.session.userId!);
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

      allEvents.sort(
        (a, b) =>
          new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime()
      );

      res.json(allEvents);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/stripe/publishable-key", async (_req, res) => {
    try {
      const key = await getStripePublishableKey();
      res.json({ publishableKey: key });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/products-with-prices", async (_req, res) => {
    try {
      const rows = await storage.listProductsWithPrices();
      const productsMap = new Map();
      for (const row of rows) {
        if (!productsMap.has(row.product_id)) {
          productsMap.set(row.product_id, {
            id: row.product_id,
            name: row.product_name,
            description: row.product_description,
            active: row.product_active,
            metadata: row.product_metadata,
            prices: []
          });
        }
        if (row.price_id) {
          productsMap.get(row.product_id).prices.push({
            id: row.price_id,
            unit_amount: row.unit_amount,
            currency: row.currency,
            recurring: row.recurring,
            active: row.price_active,
          });
        }
      }
      res.json({ data: Array.from(productsMap.values()) });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post("/api/checkout", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user) return res.status(401).json({ message: "User not found" });

      const { priceId } = req.body;
      if (!priceId) return res.status(400).json({ message: "priceId is required" });

      const stripe = await getUncachableStripeClient();

      let customerId = user.stripeCustomerId;
      if (!customerId) {
        const customer = await stripe.customers.create({
          email: user.email,
          metadata: { userId: user.id },
        });
        await storage.updateUser(user.id, { stripeCustomerId: customer.id });
        customerId = customer.id;
      }

      const session = await stripe.checkout.sessions.create({
        customer: customerId,
        payment_method_types: ['card'],
        line_items: [{ price: priceId, quantity: 1 }],
        mode: 'subscription',
        subscription_data: {
          trial_period_days: 10,
        },
        success_url: `${req.protocol}://${req.get('host')}/billing?success=true`,
        cancel_url: `${req.protocol}://${req.get('host')}/billing?canceled=true`,
      });

      res.json({ url: session.url });
    } catch (err: any) {
      console.error("Checkout error:", err);
      res.status(500).json({ message: err.message });
    }
  });

  app.post("/api/billing/portal", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user?.stripeCustomerId) {
        return res.status(400).json({ message: "No billing account found" });
      }

      const stripe = await getUncachableStripeClient();
      const session = await stripe.billingPortal.sessions.create({
        customer: user.stripeCustomerId,
        return_url: `${req.protocol}://${req.get('host')}/billing`,
      });

      res.json({ url: session.url });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/subscription", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user?.stripeSubscriptionId) {
        return res.json({ subscription: null, status: user?.subscriptionStatus || "none" });
      }

      const subscription = await storage.getSubscription(user.stripeSubscriptionId);
      res.json({ subscription, status: user.subscriptionStatus });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post("/api/subscription/sync", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user?.stripeCustomerId) {
        return res.json({ status: user?.subscriptionStatus || "none" });
      }

      const stripe = await getUncachableStripeClient();
      const subscriptions = await stripe.subscriptions.list({
        customer: user.stripeCustomerId,
        status: 'all',
        limit: 1,
      });

      if (subscriptions.data.length > 0) {
        const sub = subscriptions.data[0];
        await storage.updateUser(user.id, {
          stripeSubscriptionId: sub.id,
          subscriptionStatus: sub.status,
        });
        res.json({ status: sub.status });
      } else {
        const preservedStatuses = ["free", "admin"];
        const keepStatus = preservedStatuses.includes(user.subscriptionStatus || "") ? user.subscriptionStatus! : "none";
        await storage.updateUser(user.id, {
          stripeSubscriptionId: null,
          subscriptionStatus: keepStatus,
        });
        res.json({ status: keepStatus });
      }
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post("/api/subscription/free", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user) return res.status(404).json({ message: "User not found" });
      if (user.subscriptionStatus === "active" || user.subscriptionStatus === "trialing" || user.subscriptionStatus === "admin") {
        return res.json({ status: user.subscriptionStatus, message: "Already has active access" });
      }
      await storage.updateUser(user.id, {
        subscriptionStatus: "free",
      });
      res.json({ status: "free" });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post("/api/subscription/admin", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user) return res.status(404).json({ message: "User not found" });
      const adminEmails = (process.env.ADMIN_EMAIL || "").split(",").map(e => e.trim().toLowerCase()).filter(Boolean);
      if (!adminEmails.includes(user.email.toLowerCase())) {
        return res.status(403).json({ message: "Forbidden: Admin access only" });
      }
      await storage.updateUser(user.id, {
        subscriptionStatus: "admin",
      });
      res.json({ status: "admin" });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/mcp-key", requireAuth, async (req, res) => {
    try {
      const user = await storage.getUser(req.session.userId!);
      if (!user) return res.status(404).json({ message: "User not found" });
      res.json({ mcpApiKey: user.mcpApiKey });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post("/api/mcp-key/regenerate", requireAuth, async (req, res) => {
    try {
      const newKey = await storage.regenerateMcpApiKey(req.session.userId!);
      res.json({ mcpApiKey: newKey });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.get("/api/settings", requireAuth, async (req, res) => {
    try {
      const settings = await storage.getConnectorSettings(req.session.userId!);
      res.json({
        focusedOnly: settings?.focusedOnly ?? false,
        approvedSenders: settings?.approvedSenders ?? [],
        allowedOnedriveFolders: settings?.allowedOnedriveFolders ?? [],
      });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.put("/api/settings", requireAuth, async (req, res) => {
    try {
      const parsed = connectorSettingsUpdateSchema.parse(req.body);
      const settings = await storage.upsertConnectorSettings(req.session.userId!, parsed);
      res.json({
        focusedOnly: settings.focusedOnly,
        approvedSenders: settings.approvedSenders ?? [],
        allowedOnedriveFolders: settings.allowedOnedriveFolders ?? [],
      });
    } catch (err: any) {
      res.status(400).json({ message: err.message });
    }
  });

  app.get("/api/onedrive/browse", requireAuth, async (req, res) => {
    try {
      const accountEmail = req.query.account as string | undefined;
      const path = (req.query.path as string | undefined) || "root";
      if (!accountEmail) {
        return res.status(400).json({ message: "account query parameter is required" });
      }
      const accounts = await storage.getAccountsByUserId(req.session.userId!);
      const account = accounts.find((a) => a.email.toLowerCase() === accountEmail.toLowerCase());
      if (!account) {
        return res.status(404).json({ message: "Connected account not found" });
      }
      const accessToken = await getValidAccessToken(account.id);
      const items = await listDriveItems(accessToken, path, 100);
      const folders = items
        .filter((item: any) => item.folder)
        .map((item: any) => ({
          id: item.id,
          name: item.name,
          path: `${(item.parentReference?.path?.replace("/drive/root:", "") || "").replace(/^\/+/, "")}${item.parentReference?.path && item.parentReference.path !== "/drive/root:" ? "/" : ""}${item.name}`.replace(/^\/+/, ""),
        }));
      res.json({ folders });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  return httpServer;
}
