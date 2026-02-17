import { type ConnectedAccount, type InsertConnectedAccount, connectedAccounts } from "@shared/schema";
import { db } from "./db";
import { eq } from "drizzle-orm";

export interface IStorage {
  getAccounts(): Promise<ConnectedAccount[]>;
  getAccount(id: string): Promise<ConnectedAccount | undefined>;
  getAccountByEmail(email: string): Promise<ConnectedAccount | undefined>;
  createAccount(account: InsertConnectedAccount): Promise<ConnectedAccount>;
  updateAccountTokens(id: string, accessToken: string, refreshToken: string, tokenExpiry: Date): Promise<void>;
  deleteAccount(id: string): Promise<void>;
}

export class DatabaseStorage implements IStorage {
  async getAccounts(): Promise<ConnectedAccount[]> {
    return await db.select().from(connectedAccounts);
  }

  async getAccount(id: string): Promise<ConnectedAccount | undefined> {
    const [account] = await db.select().from(connectedAccounts).where(eq(connectedAccounts.id, id));
    return account;
  }

  async getAccountByEmail(email: string): Promise<ConnectedAccount | undefined> {
    const [account] = await db.select().from(connectedAccounts).where(eq(connectedAccounts.email, email));
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
}

export const storage = new DatabaseStorage();
