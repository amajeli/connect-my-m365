import crypto from "crypto";
import type { Express, Request, Response } from "express";
import { storage } from "./storage";

function getBaseUrl(req: Request): string {
  const protocol = req.protocol;
  const host = req.get("host");
  return `${protocol}://${host}`;
}

export function registerOAuthRoutes(app: Express): void {
  app.get("/.well-known/oauth-protected-resource", (req: Request, res: Response) => {
    const baseUrl = getBaseUrl(req);
    res.json({
      resource: `${baseUrl}/mcp`,
      authorization_servers: [baseUrl],
      scopes_supported: ["outlook:read"],
      bearer_methods_supported: ["header"],
    });
  });

  app.get("/.well-known/oauth-authorization-server", (req: Request, res: Response) => {
    const baseUrl = getBaseUrl(req);
    res.json({
      issuer: baseUrl,
      authorization_endpoint: `${baseUrl}/oauth/authorize`,
      token_endpoint: `${baseUrl}/oauth/token`,
      registration_endpoint: `${baseUrl}/oauth/register`,
      response_types_supported: ["code"],
      grant_types_supported: ["authorization_code", "refresh_token"],
      token_endpoint_auth_methods_supported: ["client_secret_post"],
      scopes_supported: ["outlook:read"],
      code_challenge_methods_supported: ["S256"],
    });
  });

  app.post("/oauth/register", async (req: Request, res: Response) => {
    try {
      const { client_name, redirect_uris } = req.body;

      if (!redirect_uris || !Array.isArray(redirect_uris) || redirect_uris.length === 0) {
        return res.status(400).json({
          error: "invalid_client_metadata",
          error_description: "redirect_uris is required and must be a non-empty array",
        });
      }

      const clientId = crypto.randomUUID();
      const clientSecret = crypto.randomBytes(32).toString("hex");

      const client = await storage.createOAuthClient({
        clientId,
        clientSecret,
        clientName: client_name || null,
        redirectUris: redirect_uris,
      });

      res.status(201).json({
        client_id: client.clientId,
        client_secret: client.clientSecret,
        client_name: client.clientName,
        redirect_uris: client.redirectUris,
        grant_types: ["authorization_code", "refresh_token"],
        response_types: ["code"],
        token_endpoint_auth_method: "client_secret_post",
      });
    } catch (err: any) {
      console.error("OAuth register error:", err.message);
      res.status(500).json({ error: "server_error", error_description: "Failed to register client" });
    }
  });

  app.get("/oauth/authorize", async (req: Request, res: Response) => {
    try {
      const {
        client_id,
        redirect_uri,
        response_type,
        state,
        scope,
        code_challenge,
        code_challenge_method,
      } = req.query as Record<string, string>;

      if (response_type !== "code") {
        return res.status(400).json({
          error: "unsupported_response_type",
          error_description: "Only response_type=code is supported",
        });
      }

      if (!client_id || !redirect_uri) {
        return res.status(400).json({
          error: "invalid_request",
          error_description: "client_id and redirect_uri are required",
        });
      }

      const client = await storage.getOAuthClient(client_id);
      if (!client) {
        return res.status(400).json({
          error: "invalid_client",
          error_description: "Unknown client_id",
        });
      }

      if (!client.redirectUris.includes(redirect_uri)) {
        return res.status(400).json({
          error: "invalid_request",
          error_description: "redirect_uri does not match any registered URIs",
        });
      }

      if (!req.session.userId) {
        req.session.pendingOAuth = {
          client_id,
          redirect_uri,
          response_type,
          state,
          scope,
          code_challenge,
          code_challenge_method,
        };
        return req.session.save((err) => {
          if (err) console.error("Session save error before OAuth redirect:", err);
          res.redirect("/api/auth/microsoft");
        });
      }

      const code = crypto.randomBytes(32).toString("hex");
      const expiresAt = new Date(Date.now() + 10 * 60 * 1000);

      await storage.createOAuthCode({
        code,
        clientId: client_id,
        userId: req.session.userId,
        redirectUri: redirect_uri,
        scope: scope || null,
        codeChallenge: code_challenge || null,
        codeChallengeMethod: code_challenge_method || null,
        expiresAt,
      });

      const redirectUrl = new URL(redirect_uri);
      redirectUrl.searchParams.set("code", code);
      if (state) {
        redirectUrl.searchParams.set("state", state);
      }

      res.redirect(redirectUrl.toString());
    } catch (err: any) {
      console.error("OAuth authorize error:", err.message);
      res.status(500).json({ error: "server_error", error_description: "Authorization failed" });
    }
  });

  app.post("/oauth/token", async (req: Request, res: Response) => {
    try {
      const grantType = req.body.grant_type;

      if (grantType === "authorization_code") {
        const { code, redirect_uri, client_id, client_secret, code_verifier } = req.body;

        if (!code || !redirect_uri || !client_id || !client_secret) {
          return res.status(400).json({
            error: "invalid_request",
            error_description: "code, redirect_uri, client_id, and client_secret are required",
          });
        }

        const client = await storage.getOAuthClient(client_id);
        if (!client || client.clientSecret !== client_secret) {
          return res.status(401).json({
            error: "invalid_client",
            error_description: "Invalid client credentials",
          });
        }

        const authCode = await storage.getOAuthCode(code);
        if (!authCode) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "Invalid authorization code",
          });
        }

        if (authCode.used) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "Authorization code has already been used",
          });
        }

        if (new Date() > authCode.expiresAt) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "Authorization code has expired",
          });
        }

        if (authCode.clientId !== client_id) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "Authorization code was not issued to this client",
          });
        }

        if (authCode.redirectUri !== redirect_uri) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "redirect_uri does not match",
          });
        }

        if (authCode.codeChallenge) {
          if (!code_verifier) {
            return res.status(400).json({
              error: "invalid_grant",
              error_description: "code_verifier is required for PKCE",
            });
          }

          const expectedChallenge = crypto
            .createHash("sha256")
            .update(code_verifier)
            .digest("base64url");

          if (expectedChallenge !== authCode.codeChallenge) {
            return res.status(400).json({
              error: "invalid_grant",
              error_description: "code_verifier validation failed",
            });
          }
        }

        await storage.markOAuthCodeUsed(code);

        const accessToken = crypto.randomBytes(32).toString("hex");
        const refreshToken = crypto.randomBytes(32).toString("hex");
        const expiresIn = 3600;
        const expiresAt = new Date(Date.now() + expiresIn * 1000);

        await storage.createOAuthToken({
          accessToken,
          refreshToken,
          clientId: client_id,
          userId: authCode.userId,
          scope: authCode.scope,
          expiresAt,
        });

        res.json({
          access_token: accessToken,
          token_type: "Bearer",
          expires_in: expiresIn,
          refresh_token: refreshToken,
          scope: authCode.scope || "outlook:read",
        });
      } else if (grantType === "refresh_token") {
        const { refresh_token, client_id, client_secret } = req.body;

        if (!refresh_token || !client_id || !client_secret) {
          return res.status(400).json({
            error: "invalid_request",
            error_description: "refresh_token, client_id, and client_secret are required",
          });
        }

        const client = await storage.getOAuthClient(client_id);
        if (!client || client.clientSecret !== client_secret) {
          return res.status(401).json({
            error: "invalid_client",
            error_description: "Invalid client credentials",
          });
        }

        const existingToken = await storage.getOAuthTokenByRefreshToken(refresh_token);
        if (!existingToken) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "Invalid refresh token",
          });
        }

        if (existingToken.clientId !== client_id) {
          return res.status(400).json({
            error: "invalid_grant",
            error_description: "Refresh token was not issued to this client",
          });
        }

        await storage.deleteOAuthToken(existingToken.id);

        const accessToken = crypto.randomBytes(32).toString("hex");
        const refreshTokenNew = crypto.randomBytes(32).toString("hex");
        const expiresIn = 3600;
        const expiresAt = new Date(Date.now() + expiresIn * 1000);

        await storage.createOAuthToken({
          accessToken,
          refreshToken: refreshTokenNew,
          clientId: client_id,
          userId: existingToken.userId,
          scope: existingToken.scope,
          expiresAt,
        });

        res.json({
          access_token: accessToken,
          token_type: "Bearer",
          expires_in: expiresIn,
          refresh_token: refreshTokenNew,
          scope: existingToken.scope || "outlook:read",
        });
      } else {
        return res.status(400).json({
          error: "unsupported_grant_type",
          error_description: "Only authorization_code and refresh_token grant types are supported",
        });
      }
    } catch (err: any) {
      console.error("OAuth token error:", err.message);
      res.status(500).json({ error: "server_error", error_description: "Token exchange failed" });
    }
  });

  console.log("OAuth routes registered");
}
