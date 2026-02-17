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

export default function BillingPage() {
  const { user } = useAuth();
  const { toast } = useToast();
  const [copiedUrl, setCopiedUrl] = useState(false);

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
          <Badge variant={isSubscribed ? "default" : "secondary"}>
            {isSubscribed ? "Active" : user.subscriptionStatus === "past_due" ? "Past Due" : "Inactive"}
          </Badge>
        </div>

        {isSubscribed ? (
          <div className="space-y-3">
            <p className="text-sm text-muted-foreground">
              Your MCP connector is active. Claude.ai can access your Outlook data.
            </p>
            <Button variant="outline" onClick={handleManageBilling} className="gap-2" data-testid="button-manage-billing">
              <ExternalLink className="w-4 h-4" />
              Manage Billing
            </Button>
          </div>
        ) : (
          <div className="space-y-4">
            <p className="text-sm text-muted-foreground">
              Start your free 5-day trial to activate the MCP connector and connect Claude.ai to your Outlook. No charge during the trial period.
            </p>
            <div className="flex gap-3 flex-wrap">
              {monthlyPrice && (
                <Button
                  onClick={() => handleCheckout(monthlyPrice.id)}
                  className="gap-2"
                  data-testid="button-subscribe-monthly"
                >
                  <Zap className="w-4 h-4" />
                  ${(monthlyPrice.unit_amount / 100).toFixed(2)}/month
                </Button>
              )}
              {yearlyPrice && (
                <Button
                  variant="outline"
                  onClick={() => handleCheckout(yearlyPrice.id)}
                  className="gap-2"
                  data-testid="button-subscribe-yearly"
                >
                  ${(yearlyPrice.unit_amount / 100).toFixed(2)}/year
                </Button>
              )}
            </div>
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
