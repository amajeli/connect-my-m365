import { useState, useEffect } from "react";
import { useQuery } from "@tanstack/react-query";
import { useAuth } from "@/lib/auth";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Skeleton } from "@/components/ui/skeleton";
import { Input } from "@/components/ui/input";
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
  CreditCard,
  Key,
  Copy,
  Check,
  RefreshCw,
  ExternalLink,
  Zap,
} from "lucide-react";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";

export default function BillingPage() {
  const { user } = useAuth();
  const { toast } = useToast();
  const [copiedKey, setCopiedKey] = useState(false);
  const [copiedUrl, setCopiedUrl] = useState(false);
  const [regenerating, setRegenerating] = useState(false);

  const { data: productsData } = useQuery<{ data: any[] }>({
    queryKey: ["/api/products-with-prices"],
  });

  const { data: mcpKeyData } = useQuery<{ mcpApiKey: string }>({
    queryKey: ["/api/mcp-key"],
    enabled: !!user,
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

  const handleCopyKey = async () => {
    const key = mcpKeyData?.mcpApiKey;
    if (!key) return;
    try {
      await navigator.clipboard.writeText(key);
      setCopiedKey(true);
      setTimeout(() => setCopiedKey(false), 2000);
      toast({ title: "API key copied" });
    } catch {
      toast({ title: "Failed to copy", variant: "destructive" });
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

  const handleRegenerateKey = async () => {
    setRegenerating(true);
    try {
      await apiRequest("POST", "/api/mcp-key/regenerate");
      queryClient.invalidateQueries({ queryKey: ["/api/mcp-key"] });
      toast({ title: "API key regenerated", description: "Update your Claude.ai connector settings with the new key." });
    } catch (err: any) {
      toast({ title: "Error", description: err.message, variant: "destructive" });
    } finally {
      setRegenerating(false);
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

          <div>
            <label className="text-xs font-medium text-muted-foreground mb-1 block">
              Your MCP API Key (Bearer Token)
            </label>
            <div className="flex gap-2">
              <Input
                readOnly
                value={mcpKeyData?.mcpApiKey || "Loading..."}
                className="font-mono text-xs"
                type="password"
                data-testid="input-mcp-api-key"
              />
              <Button size="icon" variant="ghost" onClick={handleCopyKey} data-testid="button-copy-api-key">
                {copiedKey ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
              </Button>
              <AlertDialog>
                <AlertDialogTrigger asChild>
                  <Button size="icon" variant="ghost" disabled={regenerating} data-testid="button-regenerate-key">
                    <RefreshCw className={`w-4 h-4 ${regenerating ? "animate-spin" : ""}`} />
                  </Button>
                </AlertDialogTrigger>
                <AlertDialogContent>
                  <AlertDialogHeader>
                    <AlertDialogTitle>Regenerate API Key?</AlertDialogTitle>
                    <AlertDialogDescription>
                      Your current key will stop working. You will need to update
                      the key in your Claude.ai connector settings.
                    </AlertDialogDescription>
                  </AlertDialogHeader>
                  <AlertDialogFooter>
                    <AlertDialogCancel data-testid="button-cancel-regenerate">Cancel</AlertDialogCancel>
                    <AlertDialogAction onClick={handleRegenerateKey} data-testid="button-confirm-regenerate">
                      Regenerate
                    </AlertDialogAction>
                  </AlertDialogFooter>
                </AlertDialogContent>
              </AlertDialog>
            </div>
          </div>

          <Separator />

          <div>
            <h3 className="text-sm font-medium mb-2">Setup Instructions</h3>
            <ol className="text-sm text-muted-foreground space-y-2 list-decimal list-inside">
              <li>Go to Claude.ai Settings and find Connectors</li>
              <li>Click &quot;Add custom connector&quot;</li>
              <li>Paste the MCP Server URL above</li>
              <li>In Advanced settings, add your API key as a Bearer token</li>
              <li>Save and start using Outlook tools in Claude conversations</li>
            </ol>
          </div>
        </div>
      </Card>
    </div>
  );
}
