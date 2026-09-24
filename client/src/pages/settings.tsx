import { useEffect, useState } from "react";
import { useQuery } from "@tanstack/react-query";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Separator } from "@/components/ui/separator";
import { Skeleton } from "@/components/ui/skeleton";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Settings as SettingsIcon, X, Plus, FolderPlus, Mail, HardDrive } from "lucide-react";
import { apiRequest, queryClient } from "@/lib/queryClient";
import { useToast } from "@/hooks/use-toast";

interface ConnectorSettingsData {
  focusedOnly: boolean;
  approvedSenders: string[];
  allowedOnedriveFolders: string[];
}

interface ConnectedAccount {
  id: string;
  email: string;
  displayName: string;
  accountType: string;
}

export default function SettingsPage() {
  const { toast } = useToast();

  const { data: settings, isLoading } = useQuery<ConnectorSettingsData>({
    queryKey: ["/api/settings"],
  });

  const { data: accounts } = useQuery<ConnectedAccount[]>({
    queryKey: ["/api/accounts"],
  });

  const [focusedOnly, setFocusedOnly] = useState(false);
  const [approvedSenders, setApprovedSenders] = useState<string[]>([]);
  const [allowedFolders, setAllowedFolders] = useState<string[]>([]);
  const [senderInput, setSenderInput] = useState("");
  const [folderInput, setFolderInput] = useState("");
  const [browseAccount, setBrowseAccount] = useState<string>("");
  const [browsePath, setBrowsePath] = useState<string>("root");
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (settings) {
      setFocusedOnly(settings.focusedOnly);
      setApprovedSenders(settings.approvedSenders || []);
      setAllowedFolders(settings.allowedOnedriveFolders || []);
    }
  }, [settings]);

  useEffect(() => {
    if (accounts && accounts.length > 0 && !browseAccount) {
      setBrowseAccount(accounts[0].email);
    }
  }, [accounts, browseAccount]);

  const { data: browseData, isLoading: isBrowsing } = useQuery<{ folders: { id: string; name: string; path: string }[] }>({
    queryKey: ["/api/onedrive/browse", browseAccount, browsePath],
    enabled: Boolean(browseAccount),
  });

  const addSender = () => {
    const value = senderInput.trim().toLowerCase();
    if (!value) return;
    if (approvedSenders.includes(value)) {
      setSenderInput("");
      return;
    }
    setApprovedSenders([...approvedSenders, value]);
    setSenderInput("");
  };

  const removeSender = (value: string) => {
    setApprovedSenders(approvedSenders.filter((s) => s !== value));
  };

  const addFolder = (value: string) => {
    const v = value.trim().replace(/^\/+/, "").replace(/\/+$/, "");
    if (!v) return;
    if (allowedFolders.includes(v)) return;
    setAllowedFolders([...allowedFolders, v]);
  };

  const removeFolder = (value: string) => {
    setAllowedFolders(allowedFolders.filter((f) => f !== value));
  };

  const handleSave = async () => {
    setIsSaving(true);
    try {
      await apiRequest("PUT", "/api/settings", {
        focusedOnly,
        approvedSenders,
        allowedOnedriveFolders: allowedFolders,
      });
      queryClient.invalidateQueries({ queryKey: ["/api/settings"] });
      toast({ title: "Settings saved", description: "Your v2 connector will use these rules immediately." });
    } catch (err: any) {
      toast({ title: "Failed to save", description: err.message, variant: "destructive" });
    } finally {
      setIsSaving(false);
    }
  };

  if (isLoading) {
    return (
      <div className="p-6 max-w-3xl mx-auto space-y-4">
        <Skeleton className="h-8 w-64" />
        <Skeleton className="h-40 w-full" />
        <Skeleton className="h-40 w-full" />
      </div>
    );
  }

  return (
    <div className="p-6 max-w-3xl mx-auto space-y-6 overflow-y-auto h-full" data-testid="page-settings">
      <div className="flex items-center gap-2">
        <SettingsIcon className="w-6 h-6 text-muted-foreground" />
        <div>
          <h1 className="text-xl font-semibold">Connector Settings</h1>
          <p className="text-sm text-muted-foreground">
            These allow-lists apply to the v2 and v3 MCP connectors (<code className="text-xs">/mcp/v2</code>, <code className="text-xs">/mcp/v3</code>) to reduce noise and token usage.
          </p>
        </div>
      </div>

      <Card className="p-4 space-y-4" data-testid="card-email-settings">
        <div className="flex items-center gap-2">
          <Mail className="w-5 h-5 text-muted-foreground" />
          <h2 className="font-medium">Email Filtering</h2>
        </div>

        <div className="flex items-center justify-between gap-4">
          <div>
            <Label htmlFor="focused-toggle" className="text-sm font-medium">Focused inbox only</Label>
            <p className="text-xs text-muted-foreground">When on, the connector only returns emails Outlook classifies as "Focused".</p>
          </div>
          <Switch
            id="focused-toggle"
            checked={focusedOnly}
            onCheckedChange={setFocusedOnly}
            data-testid="switch-focused-only"
          />
        </div>

        <Separator />

        <div className="space-y-2">
          <Label className="text-sm font-medium">Approved senders</Label>
          <p className="text-xs text-muted-foreground">
            If you add any, the connector only returns emails from these senders. Use a full address
            (<code className="text-xs">jane@acme.com</code>) or a domain (<code className="text-xs">@acme.com</code>). Leave empty to allow all.
          </p>
          <div className="flex gap-2">
            <Input
              placeholder="jane@acme.com or @acme.com"
              value={senderInput}
              onChange={(e) => setSenderInput(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addSender();
                }
              }}
              data-testid="input-sender"
            />
            <Button type="button" variant="secondary" onClick={addSender} data-testid="button-add-sender">
              <Plus className="w-4 h-4 mr-1" /> Add
            </Button>
          </div>
          {approvedSenders.length > 0 ? (
            <div className="flex flex-wrap gap-2 pt-1">
              {approvedSenders.map((s) => (
                <Badge key={s} variant="secondary" className="gap-1" data-testid={`badge-sender-${s}`}>
                  {s}
                  <button onClick={() => removeSender(s)} className="ml-1" data-testid={`button-remove-sender-${s}`}>
                    <X className="w-3 h-3" />
                  </button>
                </Badge>
              ))}
            </div>
          ) : (
            <p className="text-xs text-muted-foreground italic pt-1">No approved senders — all senders allowed.</p>
          )}
        </div>
      </Card>

      <Card className="p-4 space-y-4" data-testid="card-onedrive-settings">
        <div className="flex items-center gap-2">
          <HardDrive className="w-5 h-5 text-muted-foreground" />
          <h2 className="font-medium">OneDrive Folders</h2>
        </div>
        <p className="text-xs text-muted-foreground">
          If you designate any folders, the connector can only read files within those folders (and their subfolders). Leave empty to allow all of OneDrive.
        </p>

        {allowedFolders.length > 0 ? (
          <div className="flex flex-wrap gap-2">
            {allowedFolders.map((f) => (
              <Badge key={f} variant="secondary" className="gap-1" data-testid={`badge-folder-${f}`}>
                {f || "/ (root)"}
                <button onClick={() => removeFolder(f)} className="ml-1" data-testid={`button-remove-folder-${f}`}>
                  <X className="w-3 h-3" />
                </button>
              </Badge>
            ))}
          </div>
        ) : (
          <p className="text-xs text-muted-foreground italic">No folders designated — all of OneDrive is accessible.</p>
        )}

        <div className="flex gap-2">
          <Input
            placeholder="Documents/Projects"
            value={folderInput}
            onChange={(e) => setFolderInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                addFolder(folderInput);
                setFolderInput("");
              }
            }}
            data-testid="input-folder"
          />
          <Button
            type="button"
            variant="secondary"
            onClick={() => {
              addFolder(folderInput);
              setFolderInput("");
            }}
            data-testid="button-add-folder"
          >
            <Plus className="w-4 h-4 mr-1" /> Add
          </Button>
        </div>

        <Separator />

        <div className="space-y-2">
          <Label className="text-sm font-medium">Browse OneDrive</Label>
          {accounts && accounts.length > 0 ? (
            <div className="flex flex-wrap gap-2 items-center">
              <Select value={browseAccount} onValueChange={(v) => { setBrowseAccount(v); setBrowsePath("root"); }}>
                <SelectTrigger className="w-64" data-testid="select-browse-account">
                  <SelectValue placeholder="Select account" />
                </SelectTrigger>
                <SelectContent>
                  {accounts.map((a) => (
                    <SelectItem key={a.id} value={a.email}>{a.email}</SelectItem>
                  ))}
                </SelectContent>
              </Select>
              {browsePath !== "root" && (
                <Button type="button" variant="ghost" size="sm" onClick={() => setBrowsePath("root")} data-testid="button-browse-root">
                  Back to root
                </Button>
              )}
            </div>
          ) : (
            <p className="text-xs text-muted-foreground">Connect a Microsoft account to browse folders.</p>
          )}

          <p className="text-xs text-muted-foreground">
            Current path: <code className="text-xs">{browsePath === "root" ? "/ (root)" : browsePath}</code>
          </p>

          {isBrowsing ? (
            <Skeleton className="h-24 w-full" />
          ) : browseData && browseData.folders.length > 0 ? (
            <div className="space-y-1">
              {browseData.folders.map((folder) => (
                <div key={folder.id} className="flex items-center justify-between gap-2 rounded-md border p-2" data-testid={`row-folder-${folder.path}`}>
                  <button
                    className="text-sm text-left hover:underline flex-1 truncate"
                    onClick={() => setBrowsePath(folder.path)}
                    data-testid={`button-open-folder-${folder.path}`}
                  >
                    📁 {folder.name}
                  </button>
                  <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    onClick={() => addFolder(folder.path)}
                    data-testid={`button-allow-folder-${folder.path}`}
                  >
                    <FolderPlus className="w-4 h-4 mr-1" /> Allow
                  </Button>
                </div>
              ))}
            </div>
          ) : (
            <p className="text-xs text-muted-foreground italic">No subfolders here.</p>
          )}
        </div>
      </Card>

      <div className="flex justify-end">
        <Button onClick={handleSave} disabled={isSaving} data-testid="button-save-settings">
          {isSaving ? "Saving..." : "Save Settings"}
        </Button>
      </div>
    </div>
  );
}
