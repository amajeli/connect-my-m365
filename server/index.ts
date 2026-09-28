import { createServer } from "http";
import { runMigrations } from "stripe-replit-sync";
import { getStripeSync } from "./stripeClient";
import { createApp, log } from "./app";
import { serveStatic } from "./static";

async function initStripe() {
  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    throw new Error('DATABASE_URL environment variable is required for Stripe integration.');
  }

  try {
    console.log('Initializing Stripe schema...');
    await runMigrations({ databaseUrl, schema: 'stripe' });
    console.log('Stripe schema ready');

    const stripeSync = await getStripeSync();

    console.log('Setting up managed webhook...');
    const webhookBaseUrl = process.env.APP_BASE_URL || `https://${process.env.REPLIT_DOMAINS?.split(',')[0]}`;
    const result = await stripeSync.findOrCreateManagedWebhook(
      `${webhookBaseUrl}/api/stripe/webhook`
    );
    console.log(`Webhook configured: ${result?.webhook?.url || 'managed webhook ready'}`);

    console.log('Syncing Stripe data...');
    stripeSync.syncBackfill()
      .then(() => console.log('Stripe data synced'))
      .catch((err: any) => console.error('Error syncing Stripe data:', err));
  } catch (error) {
    console.error('Failed to initialize Stripe:', error);
    throw error;
  }
}

(async () => {
  const hasStripeConfig = process.env.STRIPE_SECRET_KEY || process.env.REPL_IDENTITY || process.env.WEB_REPL_RENEWAL;
  if (hasStripeConfig) {
    await initStripe();
  } else {
    console.warn('Stripe not initialized: set STRIPE_SECRET_KEY/STRIPE_PUBLISHABLE_KEY to enable.');
  }

  const app = await createApp();
  const httpServer = createServer(app);

  if (process.env.NODE_ENV === "production") {
    serveStatic(app);
  } else {
    const { setupVite } = await import("./vite");
    await setupVite(httpServer, app);
  }

  const port = parseInt(process.env.PORT || "5000", 10);
  httpServer.listen(
    {
      port,
      host: "0.0.0.0",
      ...(process.platform !== 'win32' && { reusePort: true }),
    },
    () => {
      log(`serving on port ${port}`);
    },
  );
})();
