import { useQuery } from "@tanstack/react-query";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { ThemeToggle } from "@/components/theme-toggle";
import {
  Mail,
  Calendar,
  HardDrive,
  MessageSquare,
  Zap,
  Shield,
  Check,
  ArrowRight,
  LogIn,
  LayoutGrid,
} from "lucide-react";

export default function LandingPage() {
  const { data: productsData } = useQuery<{ data: any[] }>({
    queryKey: ["/api/products-with-prices"],
  });

  const product = productsData?.data?.[0];
  const monthlyPrice = product?.prices?.find((p: any) => p.recurring?.interval === "month");
  const yearlyPrice = product?.prices?.find((p: any) => p.recurring?.interval === "year");

  return (
    <div className="min-h-screen bg-background">
      <header className="flex items-center justify-between gap-2 p-4 md:px-8 border-b sticky top-0 z-50 bg-background">
        <div className="flex items-center gap-2">
          <LayoutGrid className="w-5 h-5 text-[#0078D4]" />
          <span className="font-semibold text-sm">Microsoft 365 Connector</span>
        </div>
        <div className="flex items-center gap-2">
          <ThemeToggle />
          <a href="/api/auth/microsoft" data-testid="button-sign-in">
            <Button variant="default" className="gap-2">
              <LogIn className="w-4 h-4" />
              Sign in with Microsoft
            </Button>
          </a>
        </div>
      </header>

      <section className="py-16 md:py-24 px-4 md:px-8 max-w-4xl mx-auto text-center">
        <Badge variant="secondary" className="mb-4">MCP Custom Connector</Badge>
        <h1 className="text-3xl md:text-5xl font-bold tracking-tight mb-4">
          Your Microsoft 365 in Claude.ai
        </h1>
        <p className="text-lg text-muted-foreground max-w-2xl mx-auto mb-8">
          Connect your personal and work Microsoft email, calendar, OneDrive, and Teams directly to Claude.
          Read, search, and summarize your work data in natural language.
        </p>
        <a href="/api/auth/microsoft" data-testid="button-hero-cta">
          <Button size="lg" className="gap-2">
            Get Started
            <ArrowRight className="w-4 h-4" />
          </Button>
        </a>
      </section>

      <section className="py-12 px-4 md:px-8 max-w-5xl mx-auto">
        <div className="grid sm:grid-cols-2 md:grid-cols-4 gap-4">
          <Card className="p-6">
            <Mail className="w-8 h-8 text-[#0078D4] mb-3" />
            <h3 className="font-semibold mb-1">Email</h3>
            <p className="text-sm text-muted-foreground">
              Read, search, and summarize your inbox directly from Claude conversations.
            </p>
          </Card>
          <Card className="p-6">
            <Calendar className="w-8 h-8 text-[#0078D4] mb-3" />
            <h3 className="font-semibold mb-1">Calendar</h3>
            <p className="text-sm text-muted-foreground">
              View upcoming meetings, check availability, and get event details.
            </p>
          </Card>
          <Card className="p-6">
            <HardDrive className="w-8 h-8 text-[#0078D4] mb-3" />
            <h3 className="font-semibold mb-1">OneDrive</h3>
            <p className="text-sm text-muted-foreground">
              Browse, search, and read your OneDrive files across personal and work accounts.
            </p>
          </Card>
          <Card className="p-6">
            <MessageSquare className="w-8 h-8 text-[#0078D4] mb-3" />
            <h3 className="font-semibold mb-1">Teams</h3>
            <p className="text-sm text-muted-foreground">
              Search and read your Teams chat conversations from work accounts.
            </p>
          </Card>
        </div>
      </section>

      <section className="py-12 px-4 md:px-8 max-w-3xl mx-auto" id="pricing" data-testid="section-pricing">
        <h2 className="text-2xl font-bold text-center mb-2">Simple Pricing</h2>
        <p className="text-center text-muted-foreground mb-8">
          One plan, full access to all features.
        </p>

        <Card className="p-6 md:p-8 max-w-md mx-auto">
          <h3 className="text-xl font-bold mb-1">{product?.name || "Pro Plan"}</h3>
          <p className="text-sm text-muted-foreground mb-4">
            {product?.description || "Full access to Microsoft 365 in Claude.ai"}
          </p>

          <div className="flex items-baseline gap-1 mb-1">
            <span className="text-3xl font-bold">
              ${monthlyPrice ? (monthlyPrice.unit_amount / 100).toFixed(2) : "9.99"}
            </span>
            <span className="text-muted-foreground">/month</span>
          </div>
          {yearlyPrice && (
            <p className="text-xs text-muted-foreground mb-4">
              or ${(yearlyPrice.unit_amount / 100).toFixed(2)}/year (save ~33%)
            </p>
          )}

          <ul className="space-y-2 mb-6">
            {[
              "Connect multiple Microsoft accounts (work + personal)",
              "Email inbox access via Claude.ai",
              "Calendar events in Claude conversations",
              "OneDrive file browsing and search",
              "Teams chat reading (work accounts)",
              "Personal MCP API key",
            ].map((feature) => (
              <li key={feature} className="flex items-center gap-2 text-sm">
                <Check className="w-4 h-4 text-green-600 dark:text-green-400 shrink-0" />
                {feature}
              </li>
            ))}
          </ul>

          <a href="/api/auth/microsoft" data-testid="button-pricing-cta">
            <Button className="w-full gap-2">
              <LogIn className="w-4 h-4" />
              Start Free 10-Day Trial
            </Button>
          </a>
          <p className="text-xs text-center text-muted-foreground mt-2">
            No charge during trial. Cancel anytime.
          </p>
          <p className="text-xs text-center text-muted-foreground mt-1">
            After trial: free tier available with email-only access (20 emails/day).
          </p>
        </Card>
      </section>

      <section className="py-12 px-4 md:px-8 max-w-3xl mx-auto">
        <h2 className="text-2xl font-bold text-center mb-8">How it Works</h2>
        <div className="space-y-6">
          {[
            {
              step: "1",
              title: "Sign in with Microsoft",
              desc: "Use your work or personal Microsoft account. Connect multiple accounts to access both work and personal data.",
            },
            {
              step: "2",
              title: "Subscribe",
              desc: "Choose a monthly or yearly plan to activate your MCP connector.",
            },
            {
              step: "3",
              title: "Add to Claude.ai",
              desc: "Paste your MCP Server URL in Claude.ai Settings → Connectors. OAuth handles authentication automatically.",
            },
            {
              step: "4",
              title: "Start chatting",
              desc: "Ask Claude about your emails, meetings, files, and Teams chats — all in natural language.",
            },
          ].map(({ step, title, desc }) => (
            <div key={step} className="flex gap-4">
              <div className="w-8 h-8 rounded-full bg-primary/10 flex items-center justify-center shrink-0 text-sm font-semibold text-primary">
                {step}
              </div>
              <div>
                <h3 className="font-medium">{title}</h3>
                <p className="text-sm text-muted-foreground">{desc}</p>
              </div>
            </div>
          ))}
        </div>
      </section>

      <section className="py-12 px-4 md:px-8 max-w-3xl mx-auto">
        <Card className="p-6">
          <div className="flex items-start gap-3">
            <Shield className="w-6 h-6 text-[#0078D4] shrink-0 mt-0.5" />
            <div>
              <h3 className="font-semibold mb-1">Security & Privacy</h3>
              <p className="text-sm text-muted-foreground">
                We use Microsoft OAuth2 with delegated permissions. We never store your email
                content or file data. Your access tokens are encrypted and only used to make API
                calls on your behalf. Each user gets a unique MCP API key for isolated, secure access.
              </p>
            </div>
          </div>
        </Card>
      </section>

      <footer className="border-t py-6 px-4 text-center text-xs text-muted-foreground space-y-1">
        <p>Microsoft 365 Connector for Claude.ai</p>
        <p>Not affiliated with or endorsed by Microsoft Corporation. Microsoft, Microsoft 365, Outlook, OneDrive, and Teams are trademarks of Microsoft Corporation.</p>
      </footer>
    </div>
  );
}
