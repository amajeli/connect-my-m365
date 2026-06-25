import { Mail, Calendar, Settings, CreditCard, Plus, LogOut, LayoutGrid, SlidersHorizontal } from "lucide-react";
import { useLocation, Link } from "wouter";
import { useQuery } from "@tanstack/react-query";
import { useAuth } from "@/lib/auth";
import type { AccountInfo } from "@shared/schema";
import {
  Sidebar,
  SidebarContent,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarHeader,
  SidebarFooter,
  SidebarSeparator,
} from "@/components/ui/sidebar";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";

const navItems = [
  { title: "Inbox", url: "/inbox", icon: Mail },
  { title: "Calendar", url: "/calendar", icon: Calendar },
  { title: "Accounts", url: "/accounts", icon: Settings },
  { title: "Connector Settings", url: "/settings", icon: SlidersHorizontal },
  { title: "Billing", url: "/billing", icon: CreditCard },
];

export function AppSidebar() {
  const [location] = useLocation();
  const { toast } = useToast();
  const { user, logout } = useAuth();

  const { data: accounts, isLoading: accountsLoading } = useQuery<AccountInfo[]>({
    queryKey: ["/api/accounts"],
  });

  const handleDisconnect = async (accountId: string) => {
    try {
      await apiRequest("DELETE", `/api/accounts/${accountId}`);
      queryClient.invalidateQueries({ queryKey: ["/api/accounts"] });
      queryClient.invalidateQueries({ queryKey: ["/api/emails"] });
      queryClient.invalidateQueries({ queryKey: ["/api/calendar"] });
      toast({ title: "Account disconnected" });
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
    <Sidebar>
      <SidebarHeader className="p-4">
        <div className="flex items-center gap-2">
          <LayoutGrid className="w-6 h-6 text-[#0078D4]" />
          <div>
            <h2 className="text-sm font-semibold">Microsoft 365 Connector</h2>
            <p className="text-xs text-muted-foreground">for Claude.ai</p>
          </div>
        </div>
      </SidebarHeader>

      <SidebarSeparator />

      <SidebarContent>
        <SidebarGroup>
          <SidebarGroupLabel>Navigation</SidebarGroupLabel>
          <SidebarGroupContent>
            <SidebarMenu>
              {navItems.map((item) => (
                <SidebarMenuItem key={item.title}>
                  <SidebarMenuButton
                    asChild
                    isActive={location.startsWith(item.url)}
                  >
                    <Link href={item.url} data-testid={`link-nav-${item.title.toLowerCase()}`}>
                      <item.icon className="w-4 h-4" />
                      <span>{item.title}</span>
                    </Link>
                  </SidebarMenuButton>
                </SidebarMenuItem>
              ))}
            </SidebarMenu>
          </SidebarGroupContent>
        </SidebarGroup>

        <SidebarSeparator />

        <SidebarGroup>
          <SidebarGroupLabel>Connected Accounts</SidebarGroupLabel>
          <SidebarGroupContent>
            {accountsLoading ? (
              <div className="space-y-2 px-2">
                <Skeleton className="h-10 w-full" />
                <Skeleton className="h-10 w-full" />
              </div>
            ) : accounts && accounts.length > 0 ? (
              <SidebarMenu>
                {accounts.map((account) => (
                  <SidebarMenuItem key={account.id}>
                    <div className="flex items-center gap-2 px-2 py-1.5 w-full group">
                      <Avatar className="h-7 w-7">
                        <AvatarFallback className="text-[10px] bg-primary/10 text-primary">
                          {getInitials(account.displayName)}
                        </AvatarFallback>
                      </Avatar>
                      <div className="flex-1 min-w-0">
                        <p className="text-xs font-medium truncate">{account.displayName}</p>
                        <p className="text-[10px] text-muted-foreground truncate">{account.email}</p>
                      </div>
                      <Badge variant="secondary" className="text-[9px] shrink-0">
                        {account.accountType}
                      </Badge>
                      <Button
                        size="icon"
                        variant="ghost"
                        className="h-6 w-6 invisible group-hover:visible shrink-0"
                        onClick={() => handleDisconnect(account.id)}
                        data-testid={`button-disconnect-${account.id}`}
                      >
                        <LogOut className="w-3 h-3" />
                      </Button>
                    </div>
                  </SidebarMenuItem>
                ))}
              </SidebarMenu>
            ) : (
              <p className="text-xs text-muted-foreground px-2 py-1">No accounts connected</p>
            )}
          </SidebarGroupContent>
        </SidebarGroup>
      </SidebarContent>

      <SidebarFooter className="p-2 space-y-2">
        <a href="/api/auth/microsoft?connect=true" data-testid="button-add-account">
          <Button variant="outline" className="w-full gap-2">
            <Plus className="w-4 h-4" />
            Add Account
          </Button>
        </a>
        {user && (
          <div className="flex items-center gap-2 px-2 py-1.5">
            <Avatar className="h-7 w-7">
              <AvatarFallback className="text-[10px] bg-primary/10 text-primary">
                {getInitials(user.displayName)}
              </AvatarFallback>
            </Avatar>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-medium truncate">{user.displayName}</p>
              <p className="text-[10px] text-muted-foreground truncate">{user.email}</p>
            </div>
            <Button
              size="icon"
              variant="ghost"
              onClick={logout}
              data-testid="button-logout"
            >
              <LogOut className="w-3.5 h-3.5" />
            </Button>
          </div>
        )}
      </SidebarFooter>
    </Sidebar>
  );
}
