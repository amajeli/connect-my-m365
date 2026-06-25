import { type ConnectedAccount, type InsertConnectedAccount, type User, type InsertUser, type ConnectorSettings, type ConnectorSettingsUpdate, connectedAccounts, users, oauthClients, oauthCodes, oauthTokens, connectorSettings } from "@shared/schema";
import { db } from "./db";
import { eq, and, sql } from "drizzle-orm";

export interface IStorage {
  getUser(id: string): Promise<User | undefined>;
  getUserByMicrosoftId(microsoftId: string): Promise<User | undefined>;
  getUserByEmail(email: string): Promise<User | undefined>;
  getUserByMcpApiKey(apiKey: string): Promise<User | undefined>;
  createUser(user: InsertUser): Promise<User>;
  updateUser(id: string, data: Partial<User>): Promise<User>;
  regenerateMcpApiKey(userId: string): Promise<string>;

  getAccountsByUserId(userId: string): Promise<ConnectedAccount[]>;
  getAccount(id: string): Promise<ConnectedAccount | undefined>;
  getAccountByEmailAndUser(email: string, userId: string): Promise<ConnectedAccount | undefined>;
  createAccount(account: InsertConnectedAccount): Promise<ConnectedAccount>;
  updateAccountTokens(id: string, accessToken: string, refreshToken: string, tokenExpiry: Date): Promise<void>;
  deleteAccount(id: string): Promise<void>;

  getProduct(productId: string): Promise<any>;
  listProducts(active?: boolean): Promise<any[]>;
  listProductsWithPrices(active?: boolean): Promise<any[]>;
  getPrice(priceId: string): Promise<any>;
  getSubscription(subscriptionId: string): Promise<any>;

  createOAuthClient(data: { clientId: string; clientSecret: string; clientName?: string; redirectUris: string[] }): Promise<any>;
  getOAuthClient(clientId: string): Promise<any | undefined>;
  updateOAuthClientSecret(clientId: string, clientSecret: string): Promise<void>;
  createOAuthCode(data: { code: string; clientId: string; userId: string; redirectUri: string; scope?: string; expiresAt: Date; codeChallenge?: string; codeChallengeMethod?: string }): Promise<any>;
  getOAuthCode(code: string): Promise<any | undefined>;
  markOAuthCodeUsed(code: string): Promise<void>;
  createOAuthToken(data: { accessToken: string; refreshToken?: string; clientId: string; userId: string; scope?: string; expiresAt: Date }): Promise<any>;
  getOAuthTokenByAccessToken(accessToken: string): Promise<any | undefined>;
  getOAuthTokenByRefreshToken(refreshToken: string): Promise<any | undefined>;
  deleteOAuthToken(id: string): Promise<void>;

  getConnectorSettings(userId: string): Promise<ConnectorSettings | undefined>;
  upsertConnectorSettings(userId: string, data: ConnectorSettingsUpdate): Promise<ConnectorSettings>;
}

export class DatabaseStorage implements IStorage {
  async getUser(id: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.id, id));
    return user;
  }

  async getUserByMicrosoftId(microsoftId: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.microsoftId, microsoftId));
    return user;
  }

  async getUserByEmail(email: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.email, email));
    return user;
  }

  async getUserByMcpApiKey(apiKey: string): Promise<User | undefined> {
    const [user] = await db.select().from(users).where(eq(users.mcpApiKey, apiKey));
    return user;
  }

  async createUser(user: InsertUser): Promise<User> {
    const [created] = await db.insert(users).values(user).returning();
    return created;
  }

  async updateUser(id: string, data: Partial<User>): Promise<User> {
    const [updated] = await db.update(users).set(data).where(eq(users.id, id)).returning();
    return updated;
  }

  async regenerateMcpApiKey(userId: string): Promise<string> {
    const [updated] = await db.update(users)
      .set({ mcpApiKey: sql`gen_random_uuid()` })
      .where(eq(users.id, userId))
      .returning();
    return updated.mcpApiKey;
  }

  async getAccountsByUserId(userId: string): Promise<ConnectedAccount[]> {
    return await db.select().from(connectedAccounts).where(eq(connectedAccounts.userId, userId));
  }

  async getAccount(id: string): Promise<ConnectedAccount | undefined> {
    const [account] = await db.select().from(connectedAccounts).where(eq(connectedAccounts.id, id));
    return account;
  }

  async getAccountByEmailAndUser(email: string, userId: string): Promise<ConnectedAccount | undefined> {
    const [account] = await db.select().from(connectedAccounts)
      .where(and(eq(connectedAccounts.email, email), eq(connectedAccounts.userId, userId)));
    return account;
  }

  async createAccount(account: InsertConnectedAccount): Promise<ConnectedAccount> {
    const [created] = await db.insert(connectedAccounts).values(account).returning();
    return created;
  }

  async updateAccountTokens(id: string, accessToken: string, refreshToken: string, tokenExpiry: Date): Promise<void> {
    await db.update(connectedAccounts).set({ accessToken, refreshToken, tokenExpiry }).where(eq(connectedAccounts.id, id));
  }

  async deleteAccount(id: string): Promise<void> {
    await db.delete(connectedAccounts).where(eq(connectedAccounts.id, id));
  }

  async getProduct(productId: string) {
    const result = await db.execute(
      sql`SELECT * FROM stripe.products WHERE id = ${productId}`
    );
    return result.rows[0] || null;
  }

  async listProducts(active = true) {
    const result = await db.execute(
      sql`SELECT * FROM stripe.products WHERE active = ${active}`
    );
    return result.rows;
  }

  async listProductsWithPrices(active = true) {
    const result = await db.execute(
      sql`
        SELECT 
          p.id as product_id,
          p.name as product_name,
          p.description as product_description,
          p.active as product_active,
          p.metadata as product_metadata,
          pr.id as price_id,
          pr.unit_amount,
          pr.currency,
          pr.recurring,
          pr.active as price_active
        FROM stripe.products p
        LEFT JOIN stripe.prices pr ON pr.product = p.id AND pr.active = true
        WHERE p.active = ${active}
        ORDER BY p.id, pr.unit_amount
      `
    );
    return result.rows;
  }

  async getPrice(priceId: string) {
    const result = await db.execute(
      sql`SELECT * FROM stripe.prices WHERE id = ${priceId}`
    );
    return result.rows[0] || null;
  }

  async getSubscription(subscriptionId: string) {
    const result = await db.execute(
      sql`SELECT * FROM stripe.subscriptions WHERE id = ${subscriptionId}`
    );
    return result.rows[0] || null;
  }

  async createOAuthClient(data: { clientId: string; clientSecret: string; clientName?: string; redirectUris: string[] }): Promise<any> {
    const [created] = await db.insert(oauthClients).values(data).returning();
    return created;
  }

  async getOAuthClient(clientId: string): Promise<any | undefined> {
    const [client] = await db.select().from(oauthClients).where(eq(oauthClients.clientId, clientId));
    return client;
  }

  async updateOAuthClientSecret(clientId: string, clientSecret: string): Promise<void> {
    await db.update(oauthClients).set({ clientSecret }).where(eq(oauthClients.clientId, clientId));
  }

  async createOAuthCode(data: { code: string; clientId: string; userId: string; redirectUri: string; scope?: string; expiresAt: Date; codeChallenge?: string; codeChallengeMethod?: string }): Promise<any> {
    const [created] = await db.insert(oauthCodes).values(data).returning();
    return created;
  }

  async getOAuthCode(code: string): Promise<any | undefined> {
    const [record] = await db.select().from(oauthCodes).where(eq(oauthCodes.code, code));
    return record;
  }

  async markOAuthCodeUsed(code: string): Promise<void> {
    await db.update(oauthCodes).set({ used: true }).where(eq(oauthCodes.code, code));
  }

  async createOAuthToken(data: { accessToken: string; refreshToken?: string; clientId: string; userId: string; scope?: string; expiresAt: Date }): Promise<any> {
    const [created] = await db.insert(oauthTokens).values(data).returning();
    return created;
  }

  async getOAuthTokenByAccessToken(accessToken: string): Promise<any | undefined> {
    const [token] = await db.select().from(oauthTokens).where(eq(oauthTokens.accessToken, accessToken));
    return token;
  }

  async getOAuthTokenByRefreshToken(refreshToken: string): Promise<any | undefined> {
    const [token] = await db.select().from(oauthTokens).where(eq(oauthTokens.refreshToken, refreshToken));
    return token;
  }

  async deleteOAuthToken(id: string): Promise<void> {
    await db.delete(oauthTokens).where(eq(oauthTokens.id, id));
  }

  async getConnectorSettings(userId: string): Promise<ConnectorSettings | undefined> {
    const [settings] = await db.select().from(connectorSettings).where(eq(connectorSettings.userId, userId));
    return settings;
  }

  async upsertConnectorSettings(userId: string, data: ConnectorSettingsUpdate): Promise<ConnectorSettings> {
    const [settings] = await db
      .insert(connectorSettings)
      .values({ userId, ...data, updatedAt: new Date() })
      .onConflictDoUpdate({
        target: connectorSettings.userId,
        set: { ...data, updatedAt: new Date() },
      })
      .returning();
    return settings;
  }
}

export const storage = new DatabaseStorage();
