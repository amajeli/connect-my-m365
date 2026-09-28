import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { useAuth } from "@/lib/auth";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { Input } from "@/components/ui/input";
import { Separator } from "@/components/ui/separator";
import {
  CreditCard,
  Key,
  Copy,
  Check,
  ExternalLink,
  Zap,
} from "lucide-react";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";
import { Link } from "wouter";

export default function BillingPage() {
  const { user } = useAuth();
  const { toast } = useToast();
  const [copiedUrl, setCopiedUrl] = useState(false);
  const [copiedUrlV2, setCopiedUrlV2] = useState(false);
  const [copiedUrlV3, setCopiedUrlV3] = useState(false);

  const { data: productsData } = useQuery<{ data: any[] }>({
    queryKey: ["/api/products-with-prices"],
  });

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get("success") === "true") {
      apiRequest("POST", "/api/subscription/sync").then(() => {
        queryClient.invalidateQueries({ queryKey: ["/api/auth/me"] });
      });
      toast({ title: "Subscription activated", description: "Your MCP connector is now ready to use." });
      window.history.replaceState({}, "", "/billing");
    }
    if (params.get("canceled") === "true") {
      toast({ title: "Checkout canceled", description: "No charges were made." });
      window.history.replaceState({}, "", "/billing");
    }
  }, []);

  const product = productsData?.data?.[0];
  const monthlyPrice = product?.prices?.find((p: any) => p.recurring?.interval === "month");
  const yearlyPrice = product?.prices?.find((p: any) => p.recurring?.interval === "year");

  const isSubscribed = user?.subscriptionStatus === "active" || user?.subscriptionStatus === "trialing";
  const isFree = user?.subscriptionStatus === "free";
  const isAdminTier = user?.subscriptionStatus === "admin";
  const hasAccess = isSubscribed || isFree || isAdminTier;

  const handleCheckout = async (priceId: string) => {
    try {
      const res = await apiRequest("POST", "/api/checkout", { priceId });
      const data = await res.json();
      if (data.url) {
        window.location.href = data.url;
      }
    } catch (err: any) {
      toast({ title: "Checkout failed", description: err.message, variant: "destructive" });
    }
  };

  const handleActivateFree = async () => {
    try {
      await apiRequest("POST", "/api/subscription/free");
      queryClient.invalidateQueries({ queryKey: ["/api/auth/me"] });
      toast({ title: "Free tier activated", description: "Email access active — up to 20 emails/day via MCP." });
    } catch (err: any) {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    }
  };

  const handleActivateAdmin = async () => {
    try {
      await apiRequest("POST", "/api/subscription/admin");
      queryClient.invalidateQueries({ queryKey: ["/api/auth/me"] });
      toast({ title: "Admin tier activated", description: "Full access enabled with no expiry." });
    } catch (err: any) {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    }
  };

  const handleManageBilling = async () => {
    try {
      const res = await apiRequest("POST", "/api/billing/portal");
      const data = await res.json();
      if (data.url) {
        window.location.href = data.url;
      }
    } catch (err: any) {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    }
  };

  const handleCopyUrl = async () => {
    const url = `${window.location.origin}/mcp`;
    try {
      await navigator.clipboard.writeText(url);
      setCopiedUrl(true);
      setTimeout(() => setCopiedUrl(false), 2000);
      toast({ title: "MCP URL copied" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  const handleCopyUrlV2 = async () => {
    const url = `${window.location.origin}/mcp/v2`;
    try {
      await navigator.clipboard.writeText(url);
      setCopiedUrlV2(true);
      setTimeout(() => setCopiedUrlV2(false), 2000);
      toast({ title: "MCP v2 URL copied" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  const handleCopyUrlV3 = async () => {
    const url = `${window.location.origin}/mcp/v3`;
    try {
      await navigator.clipboard.writeText(url);
      setCopiedUrlV3(true);
      setTimeout(() => setCopiedUrlV3(false), 2000);
      toast({ title: "MCP v3 URL copied" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
    }
  };

  if (!user) {
    return (
      <div className="p-6 max-w-2xl mx-auto space-y-4">
        <Skeleton className="h-8 w-48" />
        <Skeleton className="h-40 w-full" />
      </div>
    );
  }

  return (
    <div className="p-4 md:p-6 max-w-2xl mx-auto space-y-6 overflow-y-auto h-full">
      <div>
        <h1 className="text-xl font-semibold" data-testid="text-billing-title">
          Billing & MCP Setup
        </h1>
        <p className="text-sm text-muted-foreground mt-0.5">
          Manage your subscription and Claude.ai connector
        </p>
      </div>

      <Card className="p-4" data-testid="card-subscription">
        <div className="flex items-center gap-2 mb-3">
          <CreditCard className="w-5 h-5 text-muted-foreground" />
          <h2 className="font-medium">Subscription</h2>
          <Badge variant={hasAccess ? "default" : "secondary"}>
            {isAdminTier ? "Admin" : isSubscribed ? "Active" : isFree ? "Free" : user.subscriptionStatus === "past_due" ? "Past Due" : "Inactive"}
          </Badge>
        </div>

        {isAdminTier ? (
          <div className="space-y-3">
            <p className="text-sm text-muted-foreground">
              You have <strong>Admin access</strong> — full functionality with no expiry. All MCP tools are available.
            </p>
          </div>
        ) : isSubscribed ? (
          <div className="space-y-3">
            <p className="text-sm text-muted-foreground">
              Your MCP connector is active. Claude.ai can access your full Microsoft 365 data.
            </p>
            <Button variant="outline" onClick={handleManageBilling} className="gap-2" data-testid="button-manage-billing">
              <ExternalLink className="w-4 h-4" />
              Manage Billing
            </Button>
          </div>
        ) : isFree ? (
          <div className="space-y-3">
            <p className="text-sm text-muted-foreground">
              You're on the <strong>Free tier</strong> — email access only, up to 20 emails per day. Calendar, OneDrive, and Teams require a paid subscription.
            </p>
            <div className="flex gap-3 flex-wrap">
              {monthlyPrice && (
                <Button
                  onClick={() => handleCheckout(monthlyPrice.id)}
                  className="gap-2"
                  data-testid="button-upgrade-monthly"
                >
                  <Zap className="w-4 h-4" />
                  Upgrade — ${(monthlyPrice.unit_amount / 100).toFixed(2)}/month
                </Button>
              )}
              {yearlyPrice && (
                <Button
                  variant="outline"
                  onClick={() => handleCheckout(yearlyPrice.id)}
                  className="gap-2"
                  data-testid="button-upgrade-yearly"
                >
                  ${(yearlyPrice.unit_amount / 100).toFixed(2)}/year (save ~33%)
                </Button>
              )}
            </div>
          </div>
        ) : (
          <div className="space-y-4">
            <p className="text-sm text-muted-foreground">
              Start a free 10-day trial for full access to email, calendar, OneDrive, and Teams. No charge during the trial.
            </p>
            <div className="flex gap-3 flex-wrap">
              {monthlyPrice && (
                <Button
                  onClick={() => handleCheckout(monthlyPrice.id)}
                  className="gap-2"
                  data-testid="button-subscribe-monthly"
                >
                  <Zap className="w-4 h-4" />
                  Start 10-Day Trial — ${(monthlyPrice.unit_amount / 100).toFixed(2)}/month after
                </Button>
              )}
              {yearlyPrice && (
                <Button
                  variant="outline"
                  onClick={() => handleCheckout(yearlyPrice.id)}
                  className="gap-2"
                  data-testid="button-subscribe-yearly"
                >
                  ${(yearlyPrice.unit_amount / 100).toFixed(2)}/year (save ~33%)
                </Button>
              )}
              <Button
                variant="ghost"
                onClick={handleActivateFree}
                className="gap-2 text-muted-foreground"
                data-testid="button-activate-free"
              >
                Continue with Free tier (email only, 20/day)
              </Button>
            </div>
            {user?.isAdmin && (
              <div className="pt-2 border-t">
                <p className="text-xs text-muted-foreground mb-2">Admin</p>
                <Button
                  variant="outline"
                  onClick={handleActivateAdmin}
                  className="gap-2"
                  data-testid="button-activate-admin"
                >
                  <Zap className="w-4 h-4" />
                  Activate Admin Tier
                </Button>
              </div>
            )}
          </div>
        )}
        {(isFree || isSubscribed) && user?.isAdmin && !isAdminTier && (
          <div className="mt-4 pt-3 border-t">
            <p className="text-xs text-muted-foreground mb-2">Admin</p>
            <Button
              variant="outline"
              size="sm"
              onClick={handleActivateAdmin}
              className="gap-2"
              data-testid="button-activate-admin-existing"
            >
              <Zap className="w-4 h-4" />
              Switch to Admin Tier
            </Button>
          </div>
        )}
      </Card>

      <Card className="p-4" data-testid="card-mcp-setup">
        <div className="flex items-center gap-2 mb-3">
          <Key className="w-5 h-5 text-muted-foreground" />
          <h2 className="font-medium">Claude.ai MCP Connector Setup</h2>
        </div>

        <div className="space-y-4">
          <div>
            <label className="text-xs font-medium text-muted-foreground mb-1 block">
              MCP Server URL
            </label>
            <div className="flex gap-2">
              <Input
                readOnly
                value={`${window.location.origin}/mcp`}
                className="font-mono text-xs"
                data-testid="input-mcp-url"
              />
              <Button size="icon" variant="ghost" onClick={handleCopyUrl} data-testid="button-copy-mcp-url">
                {copiedUrl ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
              </Button>
            </div>
          </div>

          <div>
            <label className="text-xs font-medium text-muted-foreground mb-1 block">
              MCP Server URL (v2 — token-efficient)
            </label>
            <div className="flex gap-2">
              <Input
                readOnly
                value={`${window.location.origin}/mcp/v2`}
                className="font-mono text-xs"
                data-testid="input-mcp-url-v2"
              />
              <Button size="icon" variant="ghost" onClick={handleCopyUrlV2} data-testid="button-copy-mcp-url-v2">
                {copiedUrlV2 ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
              </Button>
            </div>
            <p className="text-xs text-muted-foreground mt-1">
              v2 returns lean email metadata (fetching full bodies only on demand), supports date/sender filters, honors your{" "}
              <Link href="/settings" className="underline">Connector Settings</Link>, and can create drafts (never sends).
              Creating drafts requires re-consenting your Microsoft account with the Mail.ReadWrite permission.
            </p>
          </div>

          <div>
            <label className="text-xs font-medium text-muted-foreground mb-1 block">
              MCP Server URL (v3 — leanest, recommended)
            </label>
            <div className="flex gap-2">
              <Input
                readOnly
                value={`${window.location.origin}/mcp/v3`}
                className="font-mono text-xs"
                data-testid="input-mcp-url-v3"
              />
              <Button size="icon" variant="ghost" onClick={handleCopyUrlV3} data-testid="button-copy-mcp-url-v3">
                {copiedUrlV3 ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
              </Button>
            </div>
            <p className="text-xs text-muted-foreground mt-1">
              v3 keeps everything in v2 and uses fewer tokens: compact rows instead of JSON, short ids, and email
              reads that skip the quoted thread. It also finds double bookings across all your calendars, drafts
              replies from any account (never sends), and can send Teams chat messages.
            </p>
          </div>

          <Separator />

          <div>
            <h3 className="text-sm font-medium mb-2">Setup Instructions</h3>
            <ol className="text-sm text-muted-foreground space-y-2 list-decimal list-inside">
              <li>Go to Claude.ai Settings and find Connectors</li>
              <li>Click "Add custom connector"</li>
              <li>Paste the MCP Server URL above</li>
              <li>Claude will redirect you to sign in with your Microsoft account</li>
              <li>After signing in, the connector is ready to use in Claude conversations</li>
            </ol>
            <p className="text-xs text-muted-foreground mt-3">
              Authentication is handled automatically via OAuth. No API key or client secret is needed.
            </p>
          </div>
        </div>
      </Card>
    </div>
  );
}
