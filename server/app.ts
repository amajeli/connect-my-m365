import express, { type Express, type Request, Response, NextFunction } from "express";
import session from "express-session";
import connectPgSimple from "connect-pg-simple";
import { createServer } from "http";
import { registerRoutes } from "./routes";
import { registerOAuthRoutes } from "./oauth";
import { registerMcpRoutes } from "./mcp";
import { registerMcpV2Routes } from "./mcpV2";
import { registerMcpV3Routes } from "./mcpV3";
import { WebhookHandlers } from "./webhookHandlers";

// Builds the Express app without listening or serving the client, so it runs
// both as a long-lived server (server/index.ts) and as a Netlify function
// (netlify/functions/api.ts).

declare module "express-session" {
  interface SessionData {
    oauthState?: string;
    userId?: string;
    connectOnly?: boolean;
    pendingOAuth?: {
      client_id?: string;
      redirect_uri?: string;
      response_type?: string;
      state?: string;
      scope?: string;
      code_challenge?: string;
      code_challenge_method?: string;
    };
  }
}

declare module "http" {
  interface IncomingMessage {
    rawBody: unknown;
  }
}

export function log(message: string, source = "express") {
  const formattedTime = new Date().toLocaleTimeString("en-US", {
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
    hour12: true,
  });

  console.log(`${formattedTime} [${source}] ${message}`);
}

export async function createApp(): Promise<Express> {
  const app = express();
  const httpServer = createServer(app);

  app.post(
    '/api/stripe/webhook',
    express.raw({ type: 'application/json' }),
    async (req, res) => {
      const signature = req.headers['stripe-signature'];
      if (!signature) {
        return res.status(400).json({ error: 'Missing stripe-signature' });
      }

      try {
        const sig = Array.isArray(signature) ? signature[0] : signature;
        if (!Buffer.isBuffer(req.body)) {
          console.error('STRIPE WEBHOOK ERROR: req.body is not a Buffer.');
          return res.status(500).json({ error: 'Webhook processing error' });
        }

        await WebhookHandlers.processWebhook(req.body as Buffer, sig);
        res.status(200).json({ received: true });
      } catch (error: any) {
        console.error('Webhook error:', error.message);
        res.status(400).json({ error: 'Webhook processing error' });
      }
    }
  );

  app.use(
    express.json({
      verify: (req, _res, buf) => {
        req.rawBody = buf;
      },
    }),
  );

  app.use(express.urlencoded({ extended: false }));

  const PgSession = connectPgSimple(session);
  app.set('trust proxy', 1);
  app.use(
    session({
      store: new PgSession({
        conString: process.env.DATABASE_URL,
        createTableIfMissing: true,
      }),
      secret: process.env.SESSION_SECRET || "outlook-connector-session-secret",
      resave: false,
      saveUninitialized: false,
      cookie: {
        secure: process.env.NODE_ENV === "production",
        httpOnly: true,
        maxAge: 7 * 24 * 60 * 60 * 1000,
        sameSite: "lax",
      },
    })
  );

  app.use((req, res, next) => {
    const start = Date.now();
    const path = req.path;
    let capturedJsonResponse: Record<string, any> | undefined = undefined;

    const originalResJson = res.json;
    res.json = function (bodyJson, ...args) {
      capturedJsonResponse = bodyJson;
      return originalResJson.apply(res, [bodyJson, ...args]);
    };

    res.on("finish", () => {
      const duration = Date.now() - start;
      if (path.startsWith("/api") || path.startsWith("/oauth") || path.startsWith("/mcp") || path.startsWith("/.well-known")) {
        let logLine = `${req.method} ${path} ${res.statusCode} in ${duration}ms`;
        if (capturedJsonResponse) {
          const SENSITIVE_KEYS = ["access_token", "refresh_token", "client_secret", "accessToken", "refreshToken", "clientSecret", "code"];
          const redacted = JSON.stringify(capturedJsonResponse, (key, value) =>
            SENSITIVE_KEYS.includes(key) ? "[REDACTED]" : value
          );
          logLine += ` :: ${redacted}`;
        }

        log(logLine);
      }
    });

    next();
  });

  registerOAuthRoutes(app);
  registerMcpRoutes(app);
  registerMcpV2Routes(app);
  registerMcpV3Routes(app);
  await registerRoutes(httpServer, app);

  app.use((err: any, _req: Request, res: Response, next: NextFunction) => {
    const status = err.status || err.statusCode || 500;
    const message = err.message || "Internal Server Error";

    console.error("Internal Server Error:", err);

    if (res.headersSent) {
      return next(err);
    }

    return res.status(status).json({ message });
  });

  return app;
}
