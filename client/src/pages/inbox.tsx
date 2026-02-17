import { useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { format, parseISO, isToday, isYesterday } from "date-fns";
import type { OutlookEmail, AccountInfo } from "@shared/schema";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Skeleton } from "@/components/ui/skeleton";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Separator } from "@/components/ui/separator";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Mail,
  Search,
  Star,
  Paperclip,
  ArrowLeft,
  Clock,
  AlertCircle,
  Inbox as InboxIcon,
  RefreshCw,
  Copy,
  Check,
} from "lucide-react";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";

function formatEmailDate(dateStr: string): string {
  const date = parseISO(dateStr);
  if (isToday(date)) return format(date, "h:mm a");
  if (isYesterday(date)) return "Yesterday";
  return format(date, "MMM d");
}

function getInitials(name: string): string {
  return name
    .split(" ")
    .map((n) => n[0])
    .join("")
    .toUpperCase()
    .slice(0, 2);
}

function generateColor(str: string): string {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  const hue = Math.abs(hash % 360);
  return `hsl(${hue}, 45%, 55%)`;
}

export default function InboxPage() {
  const [selectedEmail, setSelectedEmail] = useState<OutlookEmail | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [accountFilter, setAccountFilter] = useState("all");
  const [copied, setCopied] = useState(false);
  const { toast } = useToast();

  const { data: accounts } = useQuery<AccountInfo[]>({
    queryKey: ["/api/accounts"],
  });

  const {
    data: emails,
    isLoading,
    isError,
    refetch,
    isFetching,
  } = useQuery<OutlookEmail[]>({
    queryKey: ["/api/emails"],
    enabled: !!accounts && accounts.length > 0,
  });

  const filteredEmails = emails?.filter((email) => {
    const matchesSearch =
      !searchQuery ||
      email.subject?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      email.from?.emailAddress?.name?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      email.bodyPreview?.toLowerCase().includes(searchQuery.toLowerCase());
    const matchesAccount = accountFilter === "all" || email.accountId === accountFilter;
    return matchesSearch && matchesAccount;
  });

  const handleCopyForClaude = async () => {
    if (!filteredEmails || filteredEmails.length === 0) return;

    const emailText = filteredEmails
      .slice(0, 25)
      .map(
        (e) =>
          `From: ${e.from?.emailAddress?.name} <${e.from?.emailAddress?.address}>\nSubject: ${e.subject}\nDate: ${e.receivedDateTime}\nPreview: ${e.bodyPreview}\n---`
      )
      .join("\n\n");

    const text = `# Outlook Inbox Summary (${filteredEmails.length} emails)\n\n${emailText}`;

    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      toast({ title: "Copied to clipboard", description: "Paste this into Claude.ai" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  if (!accounts || accounts.length === 0) {
    return <EmptyAccountState />;
  }

  if (isError) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-4 p-8">
        <AlertCircle className="w-12 h-12 text-destructive" />
        <h2 className="text-lg font-semibold">Failed to load emails</h2>
        <p className="text-sm text-muted-foreground text-center max-w-md">
          There was a problem fetching your emails. Your account tokens may have expired.
        </p>
        <Button onClick={() => refetch()} data-testid="button-retry-emails">
          Try Again
        </Button>
      </div>
    );
  }

  return (
    <div className="flex h-full">
      <div
        className={`flex flex-col border-r ${
          selectedEmail ? "hidden md:flex md:w-[380px] lg:w-[420px]" : "w-full md:w-[380px] lg:w-[420px]"
        } shrink-0`}
      >
        <div className="p-3 space-y-2 border-b">
          <div className="flex items-center gap-2">
            <div className="relative flex-1">
              <Search className="absolute left-2.5 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <Input
                placeholder="Search emails..."
                className="pl-8"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                data-testid="input-search-emails"
              />
            </div>
            <Button
              size="icon"
              variant="ghost"
              onClick={() => refetch()}
              disabled={isFetching}
              data-testid="button-refresh-emails"
            >
              <RefreshCw className={`w-4 h-4 ${isFetching ? "animate-spin" : ""}`} />
            </Button>
            <Button
              size="icon"
              variant="ghost"
              onClick={handleCopyForClaude}
              data-testid="button-copy-for-claude"
            >
              {copied ? <Check className="w-4 h-4 text-green-500" /> : <Copy className="w-4 h-4" />}
            </Button>
          </div>
          {accounts.length > 1 && (
            <Select value={accountFilter} onValueChange={setAccountFilter}>
              <SelectTrigger data-testid="select-account-filter">
                <SelectValue placeholder="All accounts" />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="all">All accounts</SelectItem>
                {accounts.map((acc) => (
                  <SelectItem key={acc.id} value={acc.id}>
                    {acc.email}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          )}
        </div>

        <ScrollArea className="flex-1">
          {isLoading ? (
            <div className="space-y-1 p-2">
              {Array.from({ length: 8 }).map((_, i) => (
                <div key={i} className="flex items-start gap-3 p-3">
                  <Skeleton className="h-9 w-9 rounded-full shrink-0" />
                  <div className="flex-1 space-y-2">
                    <Skeleton className="h-4 w-3/4" />
                    <Skeleton className="h-3 w-full" />
                    <Skeleton className="h-3 w-2/3" />
                  </div>
                </div>
              ))}
            </div>
          ) : filteredEmails && filteredEmails.length > 0 ? (
            <div className="divide-y">
              {filteredEmails.map((email) => (
                <button
                  key={`${email.accountId}-${email.id}`}
                  className={`w-full text-left p-3 hover-elevate transition-colors ${
                    selectedEmail?.id === email.id && selectedEmail?.accountId === email.accountId
                      ? "bg-accent"
                      : ""
                  } ${!email.isRead ? "bg-primary/[0.03]" : ""}`}
                  onClick={() => setSelectedEmail(email)}
                  data-testid={`button-email-${email.id}`}
                >
                  <div className="flex items-start gap-3">
                    <Avatar className="h-9 w-9 shrink-0">
                      <AvatarFallback
                        style={{ backgroundColor: generateColor(email.from?.emailAddress?.address || "") }}
                        className="text-white text-[11px]"
                      >
                        {getInitials(email.from?.emailAddress?.name || "?")}
                      </AvatarFallback>
                    </Avatar>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-1 justify-between flex-wrap">
                        <span className={`text-sm truncate ${!email.isRead ? "font-semibold" : ""}`}>
                          {email.from?.emailAddress?.name || email.from?.emailAddress?.address}
                        </span>
                        <span className="text-[11px] text-muted-foreground shrink-0">
                          {formatEmailDate(email.receivedDateTime)}
                        </span>
                      </div>
                      <p className={`text-sm truncate ${!email.isRead ? "font-medium" : "text-muted-foreground"}`}>
                        {email.subject || "(No subject)"}
                      </p>
                      <p className="text-xs text-muted-foreground truncate mt-0.5">
                        {email.bodyPreview}
                      </p>
                      <div className="flex items-center gap-1 mt-1">
                        {email.hasAttachments && (
                          <Paperclip className="w-3 h-3 text-muted-foreground" />
                        )}
                        {email.importance === "high" && (
                          <Star className="w-3 h-3 text-yellow-500 fill-yellow-500" />
                        )}
                        {accounts.length > 1 && (
                          <Badge variant="secondary" className="text-[9px] h-4">
                            {email.accountEmail}
                          </Badge>
                        )}
                      </div>
                    </div>
                  </div>
                </button>
              ))}
            </div>
          ) : (
            <div className="flex flex-col items-center justify-center p-8 text-center">
              <InboxIcon className="w-10 h-10 text-muted-foreground mb-3" />
              <p className="text-sm font-medium">No emails found</p>
              <p className="text-xs text-muted-foreground mt-1">
                {searchQuery ? "Try a different search term" : "Your inbox is empty"}
              </p>
            </div>
          )}
        </ScrollArea>
      </div>

      <div className={`flex-1 ${selectedEmail ? "flex flex-col" : "hidden md:flex md:flex-col"}`}>
        {selectedEmail ? (
          <EmailDetail email={selectedEmail} onBack={() => setSelectedEmail(null)} />
        ) : (
          <div className="flex flex-col items-center justify-center h-full text-center p-8">
            <Mail className="w-12 h-12 text-muted-foreground mb-4" />
            <h3 className="text-base font-medium">Select an email</h3>
            <p className="text-sm text-muted-foreground mt-1">
              Choose an email from the list to read it
            </p>
          </div>
        )}
      </div>
    </div>
  );
}

function EmailDetail({ email, onBack }: { email: OutlookEmail; onBack: () => void }) {
  const [copied, setCopied] = useState(false);
  const { toast } = useToast();

  const handleCopyEmail = async () => {
    const text = `From: ${email.from?.emailAddress?.name} <${email.from?.emailAddress?.address}>
To: ${email.toRecipients?.map((r) => `${r.emailAddress.name} <${r.emailAddress.address}>`).join(", ")}
Subject: ${email.subject}
Date: ${email.receivedDateTime}

${email.bodyPreview}`;

    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
      toast({ title: "Email copied", description: "Paste into Claude.ai" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  return (
    <>
      <div className="flex items-center gap-2 p-3 border-b">
        <Button size="icon" variant="ghost" onClick={onBack} className="md:hidden" data-testid="button-back-to-list">
          <ArrowLeft className="w-4 h-4" />
        </Button>
        <div className="flex-1" />
        <Button size="sm" variant="outline" onClick={handleCopyEmail} className="gap-1.5" data-testid="button-copy-email">
          {copied ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
          {copied ? "Copied" : "Copy for Claude"}
        </Button>
      </div>
      <ScrollArea className="flex-1">
        <div className="p-4 md:p-6 max-w-3xl">
          <h1 className="text-xl font-semibold mb-4" data-testid="text-email-subject">
            {email.subject || "(No subject)"}
          </h1>
          <div className="flex items-start gap-3 mb-6">
            <Avatar className="h-10 w-10">
              <AvatarFallback
                style={{ backgroundColor: generateColor(email.from?.emailAddress?.address || "") }}
                className="text-white text-xs"
              >
                {getInitials(email.from?.emailAddress?.name || "?")}
              </AvatarFallback>
            </Avatar>
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-sm font-medium">
                  {email.from?.emailAddress?.name}
                </span>
                <span className="text-xs text-muted-foreground">
                  &lt;{email.from?.emailAddress?.address}&gt;
                </span>
              </div>
              <div className="flex items-center gap-1 text-xs text-muted-foreground mt-0.5">
                <Clock className="w-3 h-3" />
                <span>{format(parseISO(email.receivedDateTime), "EEEE, MMMM d, yyyy 'at' h:mm a")}</span>
              </div>
              {email.toRecipients && email.toRecipients.length > 0 && (
                <p className="text-xs text-muted-foreground mt-1">
                  To: {email.toRecipients.map((r) => r.emailAddress.name || r.emailAddress.address).join(", ")}
                </p>
              )}
            </div>
          </div>
          <Separator className="mb-4" />
          {email.body?.contentType === "html" ? (
            <div
              className="prose prose-sm dark:prose-invert max-w-none"
              dangerouslySetInnerHTML={{ __html: email.body.content }}
              data-testid="text-email-body"
            />
          ) : (
            <div className="text-sm whitespace-pre-wrap" data-testid="text-email-body">
              {email.body?.content || email.bodyPreview}
            </div>
          )}
        </div>
      </ScrollArea>
    </>
  );
}

function EmptyAccountState() {
  return (
    <div className="flex flex-col items-center justify-center h-full p-8 text-center">
      <div className="w-16 h-16 rounded-full bg-primary/10 flex items-center justify-center mb-4">
        <Mail className="w-8 h-8 text-primary" />
      </div>
      <h2 className="text-xl font-semibold mb-2">Connect your Outlook</h2>
      <p className="text-sm text-muted-foreground max-w-md mb-6">
        Connect your work or personal Microsoft account to view your emails and
        calendar events. You can then copy the data for use in Claude.ai.
      </p>
      <a href="/api/auth/microsoft" data-testid="button-connect-first-account">
        <Button className="gap-2">
          <Mail className="w-4 h-4" />
          Connect Microsoft Account
        </Button>
      </a>
    </div>
  );
}
