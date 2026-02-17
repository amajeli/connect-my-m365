import { useQuery } from "@tanstack/react-query";
import { format, parseISO } from "date-fns";
import type { AccountInfo } from "@shared/schema";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Separator } from "@/components/ui/separator";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from "@/components/ui/alert-dialog";
import {
  Plus,
  LogOut,
  Mail,
  Calendar,
  Shield,
  CheckCircle,
} from "lucide-react";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";

export default function AccountsPage() {
  const { toast } = useToast();

  const {
    data: accounts,
    isLoading,
  } = useQuery<AccountInfo[]>({
    queryKey: ["/api/accounts"],
  });

  const handleDisconnect = async (accountId: string) => {
    try {
      await apiRequest("DELETE", `/api/accounts/${accountId}`);
      queryClient.invalidateQueries({ queryKey: ["/api/accounts"] });
      queryClient.invalidateQueries({ queryKey: ["/api/emails"] });
      queryClient.invalidateQueries({ queryKey: ["/api/calendar"] });
      toast({ title: "Account disconnected successfully" });
    } catch {
      toast({ title: "Failed to disconnect account", variant: "destructive" });
    }
  };

  const getInitials = (name: string) => {
    return name
      .split(" ")
      .map((n) => n[0])
      .join("")
      .toUpperCase()
      .slice(0, 2);
  };

  return (
    <div className="p-4 md:p-6 max-w-2xl mx-auto">
      <div className="flex items-center justify-between gap-2 mb-6 flex-wrap">
        <div>
          <h1 className="text-xl font-semibold" data-testid="text-accounts-title">Connected Accounts</h1>
          <p className="text-sm text-muted-foreground mt-0.5">
            Manage your Microsoft Outlook connections
          </p>
        </div>
        <a href="/api/auth/microsoft" data-testid="button-add-new-account">
          <Button className="gap-2">
            <Plus className="w-4 h-4" />
            Add Account
          </Button>
        </a>
      </div>

      {isLoading ? (
        <div className="space-y-3">
          {Array.from({ length: 2 }).map((_, i) => (
            <Skeleton key={i} className="h-28 w-full rounded-md" />
          ))}
        </div>
      ) : accounts && accounts.length > 0 ? (
        <div className="space-y-3">
          {accounts.map((account) => (
            <Card key={account.id} className="p-4" data-testid={`card-account-${account.id}`}>
              <div className="flex items-start gap-4">
                <Avatar className="h-12 w-12">
                  <AvatarFallback className="bg-primary/10 text-primary text-sm">
                    {getInitials(account.displayName)}
                  </AvatarFallback>
                </Avatar>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <h3 className="font-medium" data-testid={`text-account-name-${account.id}`}>
                      {account.displayName}
                    </h3>
                    <Badge variant="secondary">{account.accountType}</Badge>
                    <div className="flex items-center gap-1 text-green-600 dark:text-green-400">
                      <CheckCircle className="w-3.5 h-3.5" />
                      <span className="text-xs">Connected</span>
                    </div>
                  </div>
                  <p className="text-sm text-muted-foreground mt-0.5">{account.email}</p>
                  <div className="flex items-center gap-4 mt-2 text-xs text-muted-foreground flex-wrap">
                    <span className="flex items-center gap-1">
                      <Calendar className="w-3 h-3" />
                      Connected {format(parseISO(account.connectedAt), "MMM d, yyyy")}
                    </span>
                    <span className="flex items-center gap-1">
                      <Shield className="w-3 h-3" />
                      Read-only access
                    </span>
                  </div>
                  <div className="flex items-center gap-2 mt-2">
                    <Badge variant="outline" className="gap-1 text-xs">
                      <Mail className="w-3 h-3" />
                      Mail.Read
                    </Badge>
                    <Badge variant="outline" className="gap-1 text-xs">
                      <Calendar className="w-3 h-3" />
                      Calendars.Read
                    </Badge>
                  </div>
                </div>
                <AlertDialog>
                  <AlertDialogTrigger asChild>
                    <Button variant="ghost" size="sm" className="text-destructive gap-1.5 shrink-0" data-testid={`button-disconnect-account-${account.id}`}>
                      <LogOut className="w-3.5 h-3.5" />
                      Disconnect
                    </Button>
                  </AlertDialogTrigger>
                  <AlertDialogContent>
                    <AlertDialogHeader>
                      <AlertDialogTitle>Disconnect account?</AlertDialogTitle>
                      <AlertDialogDescription>
                        This will remove the connection to {account.email}. You can
                        reconnect anytime.
                      </AlertDialogDescription>
                    </AlertDialogHeader>
                    <AlertDialogFooter>
                      <AlertDialogCancel data-testid="button-cancel-disconnect">Cancel</AlertDialogCancel>
                      <AlertDialogAction
                        onClick={() => handleDisconnect(account.id)}
                        data-testid="button-confirm-disconnect"
                      >
                        Disconnect
                      </AlertDialogAction>
                    </AlertDialogFooter>
                  </AlertDialogContent>
                </AlertDialog>
              </div>
            </Card>
          ))}
        </div>
      ) : (
        <Card className="p-8 text-center">
          <div className="flex flex-col items-center">
            <div className="w-14 h-14 rounded-full bg-primary/10 flex items-center justify-center mb-4">
              <Mail className="w-7 h-7 text-primary" />
            </div>
            <h3 className="font-medium mb-1">No accounts connected</h3>
            <p className="text-sm text-muted-foreground max-w-sm mb-4">
              Connect your Microsoft work or personal account to start pulling
              emails and calendar events for Claude.ai.
            </p>
            <a href="/api/auth/microsoft" data-testid="button-connect-from-empty">
              <Button className="gap-2">
                <Plus className="w-4 h-4" />
                Connect Microsoft Account
              </Button>
            </a>
          </div>
        </Card>
      )}

      <Separator className="my-6" />

      <Card className="p-4">
        <h3 className="text-sm font-medium mb-2">How it works</h3>
        <div className="space-y-3 text-sm text-muted-foreground">
          <div className="flex items-start gap-3">
            <div className="w-6 h-6 rounded-full bg-primary/10 flex items-center justify-center shrink-0 text-xs font-medium text-primary">
              1
            </div>
            <p>Connect your Microsoft work or personal Outlook account using OAuth</p>
          </div>
          <div className="flex items-start gap-3">
            <div className="w-6 h-6 rounded-full bg-primary/10 flex items-center justify-center shrink-0 text-xs font-medium text-primary">
              2
            </div>
            <p>View your emails and calendar events in one place</p>
          </div>
          <div className="flex items-start gap-3">
            <div className="w-6 h-6 rounded-full bg-primary/10 flex items-center justify-center shrink-0 text-xs font-medium text-primary">
              3
            </div>
            <p>Copy formatted data to clipboard and paste directly into Claude.ai</p>
          </div>
        </div>
      </Card>
    </div>
  );
}
