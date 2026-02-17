import { sql } from "drizzle-orm";
import { pgTable, text, varchar, timestamp, boolean, jsonb } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";

export const connectedAccounts = pgTable("connected_accounts", {
  id: varchar("id").primaryKey().default(sql`gen_random_uuid()`),
  email: text("email").notNull(),
  displayName: text("display_name").notNull(),
  accountType: text("account_type").notNull(), // 'work' or 'personal'
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
