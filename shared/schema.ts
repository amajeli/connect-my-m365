import { sql } from "drizzle-orm";
import { pgTable, text, varchar, timestamp, boolean } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const users = pgTable("users", {
  id: varchar("id").primaryKey().default(sql`gen_random_uuid()`),
  email: text("email").notNull().unique(),
  displayName: text("display_name").notNull(),
  microsoftId: text("microsoft_id").notNull().unique(),
  mcpApiKey: text("mcp_api_key").notNull().default(sql`gen_random_uuid()`),
  stripeCustomerId: text("stripe_customer_id"),
  stripeSubscriptionId: text("stripe_subscription_id"),
  subscriptionStatus: text("subscription_status").default("none"),
  createdAt: timestamp("created_at").notNull().default(sql`now()`),
});

export const insertUserSchema = createInsertSchema(users).omit({
  id: true,
  mcpApiKey: true,
  createdAt: true,
});

export type InsertUser = z.infer<typeof insertUserSchema>;
export type User = typeof users.$inferSelect;

export const connectedAccounts = pgTable("connected_accounts", {
  id: varchar("id").primaryKey().default(sql`gen_random_uuid()`),
  userId: varchar("user_id").notNull().references(() => users.id, { onDelete: "cascade" }),
  email: text("email").notNull(),
  displayName: text("display_name").notNull(),
  accountType: text("account_type").notNull(),
  accessToken: text("access_token").notNull(),
  refreshToken: text("refresh_token").notNull(),
  tokenExpiry: timestamp("token_expiry").notNull(),
  isActive: boolean("is_active").notNull().default(true),
  connectedAt: timestamp("connected_at").notNull().default(sql`now()`),
});

export const insertConnectedAccountSchema = createInsertSchema(connectedAccounts).omit({
  id: true,
  connectedAt: true,
});

export type InsertConnectedAccount = z.infer<typeof insertConnectedAccountSchema>;
export type ConnectedAccount = typeof connectedAccounts.$inferSelect;

export type OutlookEmail = {
  id: string;
  subject: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  toRecipients: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  receivedDateTime: string;
  bodyPreview: string;
  body: {
    contentType: string;
    content: string;
  };
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;
  conversationId: string;
  parentFolderId: string;
  accountId: string;
  accountEmail: string;
};

export type OutlookCalendarEvent = {
  id: string;
  subject: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  location?: {
    displayName: string;
  };
  organizer?: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  attendees?: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
    status?: {
      response: string;
    };
    type: string;
  }>;
  isAllDay: boolean;
  bodyPreview: string;
  body?: {
    contentType: string;
    content: string;
  };
  showAs: string;
  importance: string;
  isOnlineMeeting: boolean;
  onlineMeetingUrl?: string;
  accountId: string;
  accountEmail: string;
};

export type AccountInfo = {
  id: string;
  email: string;
  displayName: string;
  accountType: string;
  isActive: boolean;
  connectedAt: string;
};

export type UserInfo = {
  id: string;
  email: string;
  displayName: string;
  mcpApiKey: string;
  subscriptionStatus: string;
  stripeCustomerId: string | null;
  stripeSubscriptionId: string | null;
};
