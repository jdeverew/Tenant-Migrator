import { useParams, Link } from "wouter";
import { useProject, useProjectStats, useUpdateProject } from "@/hooks/use-projects";
import { useMigrationItems, useCreateMigrationItem, useUpdateMigrationItem, useDeleteMigrationItem } from "@/hooks/use-items";
import { Sidebar } from "@/components/Sidebar";
import { StatusBadge } from "@/components/StatusBadge";
import { Loader2, ArrowLeft, Mail, Cloud, Users, Plus, Trash2, RotateCw, Eye, EyeOff, CheckCircle2, XCircle, Shield, ExternalLink, Play, PlayCircle, FileText, Globe, KeyRound, Search, UserCheck, MapPin, Zap, AlertTriangle, Import, Boxes, Server, Download, Terminal, Wand2, Copy, Sparkles, HardDrive, RefreshCw, AtSign, Inbox, Building2 } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip } from "recharts";
import { Progress } from "@/components/ui/progress";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { useForm, Controller } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { useToast } from "@/hooks/use-toast";
import { useState, useEffect } from "react";
import { type MigrationItem, type MappingRule } from "@shared/schema";
import { format } from "date-fns";
import { apiRequest } from "@/lib/queryClient";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import { api, buildUrl } from "@shared/routes";

const itemSchema = z.object({
  sourceIdentity: z.string().min(1, "Source identity is required"),
  targetIdentity: z.string().optional().or(z.literal("")),
  itemType: z.enum(["mailbox", "onedrive", "sharepoint", "teams", "user", "distributiongroup", "sharedmailbox", "m365group", "powerplatform"]),
});

type ItemFormData = z.infer<typeof itemSchema>;

type ViewType = 'overview' | 'exchange' | 'sharepoint' | 'onedrive' | 'teams' | 'users' | 'powerplatform' | 'entra_ad' | 'distributiongroups' | 'sharedmailboxes' | 'm365groups' | 'discovery' | 'mapping' | 'tenant_config';

const MIGRATION_SERVICES = [
  { key: 'exchange',           label: 'Exchange Online',         icon: Mail,       itemTypes: ['mailbox'],            description: 'Email, calendars & mailboxes' },
  { key: 'sharepoint',         label: 'SharePoint Online',       icon: Globe,      itemTypes: ['sharepoint'],         description: 'Sites & document libraries' },
  { key: 'onedrive',           label: 'OneDrive',                icon: Cloud,      itemTypes: ['onedrive'],           description: 'Personal files & folders' },
  { key: 'teams',              label: 'Microsoft Teams',         icon: Users,      itemTypes: ['teams'],              description: 'Teams, channels & chats' },
  { key: 'users',              label: 'User Accounts',           icon: UserCheck,  itemTypes: ['user'],               description: 'User identities & accounts' },
  { key: 'distributiongroups', label: 'Distribution Groups',     icon: AtSign,     itemTypes: ['distributiongroup'],  description: 'Mail-enabled distribution lists' },
  { key: 'sharedmailboxes',    label: 'Shared Mailboxes',        icon: Inbox,      itemTypes: ['sharedmailbox'],      description: 'Shared mailboxes & delegates' },
  { key: 'm365groups',         label: 'Microsoft 365 Groups',    icon: Building2,  itemTypes: ['m365group'],          description: 'M365 Groups with members & owners' },
  { key: 'powerplatform',      label: 'Power Platform',          icon: Zap,        itemTypes: ['powerplatform'],      description: 'Apps, flows & environments' },
  { key: 'entra_ad',           label: 'Entra ID → On-Prem',     icon: Server,     itemTypes: ['entra_to_ad'],        description: 'Cloud to on-premises AD' },
] as const;

function formatBytes(bytes: number | null | undefined): string {
  if (!bytes) return '0 B';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

function ItemTypeIcon({ type }: { type: string }) {
  switch (type) {
    case 'mailbox': return <Mail className="w-4 h-4 text-blue-500" />;
    case 'onedrive': return <Cloud className="w-4 h-4 text-sky-500" />;
    case 'sharepoint': return <Globe className="w-4 h-4 text-teal-500" />;
    case 'teams': return <Users className="w-4 h-4 text-indigo-500" />;
    case 'user': return <UserCheck className="w-4 h-4 text-violet-500" />;
    case 'distributiongroup': return <AtSign className="w-4 h-4 text-orange-500" />;
    case 'sharedmailbox': return <Inbox className="w-4 h-4 text-rose-500" />;
    case 'm365group': return <Building2 className="w-4 h-4 text-cyan-500" />;
    case 'powerplatform': return <Zap className="w-4 h-4 text-amber-500" />;
    default: return null;
  }
}

export default function ProjectDetails() {
  const params = useParams();
  const id = Number(params.id);
  const queryClient = useQueryClient();
  const { data: project, isLoading: projectLoading } = useProject(id);
  const { data: stats, isLoading: statsLoading } = useProjectStats(id);
  const { data: items, isLoading: itemsLoading } = useMigrationItems(id);
  const { mutateAsync: updateProject } = useUpdateProject();
  const { mutateAsync: createItem } = useCreateMigrationItem();
  const { mutateAsync: updateItem } = useUpdateMigrationItem();
  const { mutateAsync: deleteItem } = useDeleteMigrationItem();

  const [currentView, setCurrentView] = useState<ViewType>('overview');
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [logsDialogItem, setLogsDialogItem] = useState<MigrationItem | null>(null);
  const [itemLogs, setItemLogs] = useState<string[]>([]);
  const [logsLoading, setLogsLoading] = useState(false);
  const [selectedServiceItems, setSelectedServiceItems] = useState<Set<number>>(new Set());
  const [isDeletingSelected, setIsDeletingSelected] = useState(false);
  const { toast } = useToast();

  // Continuous sync
  const { data: syncStatus, refetch: refetchSyncStatus } = useQuery<{
    syncEnabled: boolean;
    syncIntervalMinutes: number;
    completedItems: number;
    items: { id: number; itemType: string; sourceIdentity: string; targetIdentity: string | null; lastSyncedAt: string | null; nextSyncAt: string | null }[];
  }>({ queryKey: ['/api/projects', id, 'sync-status'], queryFn: async () => {
    const res = await fetch(`/api/projects/${id}/sync-status`, { credentials: 'include' });
    if (!res.ok) throw new Error('Failed to load sync status');
    return res.json();
  }, enabled: !!id });

  const [syncSaving, setSyncSaving] = useState(false);
  const [syncRunning, setSyncRunning] = useState(false);

  const handleToggleSync = async (enabled: boolean) => {
    setSyncSaving(true);
    try {
      await apiRequest('PATCH', `/api/projects/${id}/sync-settings`, { syncEnabled: enabled });
      await refetchSyncStatus();
      queryClient.invalidateQueries({ queryKey: ['/api/projects', id] });
      toast({ title: enabled ? 'Continuous sync enabled' : 'Continuous sync disabled', description: enabled ? 'New emails and files will be automatically synced.' : 'Automatic sync has been paused.' });
    } catch (e: any) {
      toast({ title: 'Error', description: e.message, variant: 'destructive' });
    } finally {
      setSyncSaving(false);
    }
  };

  const handleSyncIntervalChange = async (minutes: string) => {
    setSyncSaving(true);
    try {
      await apiRequest('PATCH', `/api/projects/${id}/sync-settings`, { syncIntervalMinutes: Number(minutes) });
      await refetchSyncStatus();
    } catch (e: any) {
      toast({ title: 'Error', description: e.message, variant: 'destructive' });
    } finally {
      setSyncSaving(false);
    }
  };

  const handleSyncNow = async () => {
    setSyncRunning(true);
    try {
      await apiRequest('POST', `/api/projects/${id}/sync-now`, {});
      toast({ title: 'Sync triggered', description: 'Catching up on new emails and files — check item logs for details.' });
      setTimeout(() => refetchSyncStatus(), 5000);
    } catch (e: any) {
      toast({ title: 'Error', description: e.message, variant: 'destructive' });
    } finally {
      setSyncRunning(false);
    }
  };

  const handleDeleteSelected = async (svcItemIds: number[]) => {
    const toDelete = svcItemIds.filter(id => selectedServiceItems.has(id));
    if (toDelete.length === 0) return;
    if (!confirm(`Delete ${toDelete.length} item(s)? This cannot be undone.`)) return;
    setIsDeletingSelected(true);
    try {
      await Promise.all(toDelete.map(itemId => deleteItem({ id: itemId, projectId: id })));
      setSelectedServiceItems(prev => { const next = new Set(prev); toDelete.forEach(i => next.delete(i)); return next; });
      toast({ title: "Deleted", description: `${toDelete.length} item(s) removed.` });
    } catch {
      toast({ title: "Error", description: "Failed to delete some items.", variant: "destructive" });
    } finally {
      setIsDeletingSelected(false);
    }
  };

  // Show toast after OAuth redirect back from Microsoft
  useEffect(() => {
    const url = new URL(window.location.href);
    const oauthSuccess = url.searchParams.get('oauth_success');
    const oauthError = url.searchParams.get('oauth_error');
    const appName = url.searchParams.get('app');
    if (oauthSuccess) {
      toast({
        title: "Tenant connected successfully!",
        description: `${appName ? `"${appName}" app was created and credentials saved` : 'App registration created and credentials saved'} for the ${oauthSuccess} tenant.`,
      });
      url.searchParams.delete('oauth_success');
      url.searchParams.delete('app');
      window.history.replaceState({}, '', url.toString());
      queryClient.invalidateQueries({ queryKey: [api.projects.get.path, id] });
    } else if (oauthError) {
      toast({
        title: "Connection failed",
        description: decodeURIComponent(oauthError),
        variant: "destructive",
      });
      url.searchParams.delete('oauth_error');
      window.history.replaceState({}, '', url.toString());
    }
  }, []);

  const hasInProgress = items?.some(i => i.status === 'in_progress');
  useEffect(() => {
    if (!hasInProgress) return;
    const interval = setInterval(() => {
      queryClient.invalidateQueries({ queryKey: [api.items.list.path, id] });
      queryClient.invalidateQueries({ queryKey: [api.projects.stats.path, id] });
    }, 3000);
    return () => clearInterval(interval);
  }, [hasInProgress, id, queryClient]);

  const form = useForm<ItemFormData>({
    resolver: zodResolver(itemSchema),
    defaultValues: {
      sourceIdentity: "",
      targetIdentity: "",
      itemType: "mailbox",
    },
  });

  const onSubmit = async (data: ItemFormData) => {
    try {
      await createItem({
        projectId: id,
        sourceIdentity: data.sourceIdentity,
        targetIdentity: data.targetIdentity || undefined,
        itemType: data.itemType,
        status: "pending",
      });
      setIsAddOpen(false);
      form.reset();
      toast({ title: "Success", description: "Migration item added" });
    } catch (error) {
      toast({ title: "Error", description: "Failed to add item", variant: "destructive" });
    }
  };

  const handleRetry = async (itemId: number) => {
    try {
      await updateItem({ id: itemId, projectId: id, status: "pending", errorDetails: null });
      toast({ title: "Retrying", description: "Item status reset to pending" });
    } catch (error) {
      toast({ title: "Error", description: "Failed to retry item", variant: "destructive" });
    }
  };

  const handleDelete = async (itemId: number) => {
    if (!confirm("Are you sure you want to delete this item?")) return;
    try {
      await deleteItem({ id: itemId, projectId: id });
      toast({ title: "Deleted", description: "Item removed from project" });
    } catch (error) {
      toast({ title: "Error", description: "Failed to delete item", variant: "destructive" });
    }
  };

  const handleMigrateItem = async (itemId: number) => {
    try {
      await apiRequest('POST', `/api/projects/${id}/items/${itemId}/migrate`);
      toast({ title: "Migration Started", description: "The migration has been started. Status will update automatically." });
      queryClient.invalidateQueries({ queryKey: [api.items.list.path, id] });
      queryClient.invalidateQueries({ queryKey: [api.projects.stats.path, id] });
    } catch (err: any) {
      const message = err.message?.includes('409') ? 'Migration already in progress' : 'Failed to start migration';
      toast({ title: "Error", description: message, variant: "destructive" });
    }
  };

  const handleExportLogs = async (format: 'txt' | 'csv' = 'txt') => {
    try {
      const res = await fetch(`/api/projects/${id}/export-logs?format=${format}`, { credentials: 'include' });
      if (!res.ok) throw new Error(`Export failed: ${res.statusText}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `migration-logs-${id}-${Date.now()}.${format}`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err: any) {
      toast({ title: 'Export failed', description: err.message, variant: 'destructive' });
    }
  };

  const handleMigrateAll = async () => {
    try {
      const res = await apiRequest('POST', `/api/projects/${id}/migrate-all`);
      const data = await res.json();
      toast({
        title: "Batch Migration Started",
        description: data.message,
      });
      queryClient.invalidateQueries({ queryKey: [api.items.list.path, id] });
      queryClient.invalidateQueries({ queryKey: [api.projects.stats.path, id] });
    } catch {
      toast({ title: "Error", description: "Failed to start batch migration", variant: "destructive" });
    }
  };

  const handleViewLogs = async (item: MigrationItem) => {
    setLogsDialogItem(item);
    setLogsLoading(true);
    try {
      const res = await apiRequest('GET', `/api/items/${item.id}/logs`);
      const data = await res.json();
      setItemLogs(data.logs || []);
    } catch {
      setItemLogs([]);
    } finally {
      setLogsLoading(false);
    }
  };

  if (projectLoading || statsLoading) {
    return (
      <div className="flex h-screen w-full items-center justify-center bg-background">
        <Loader2 className="w-10 h-10 animate-spin text-primary" />
      </div>
    );
  }

  if (!project) return <div>Project not found</div>;

  const chartData = stats ? [
    { name: 'Completed', value: stats.completed, color: '#10b981' },
    { name: 'In Progress', value: stats.inProgress, color: '#3b82f6' },
    { name: 'Failed', value: stats.failed, color: '#ef4444' },
    { name: 'Pending', value: stats.pending, color: '#94a3b8' },
  ].filter(d => d.value > 0) : [];

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 text-foreground font-sans">
      <Sidebar />

      <main className="flex-1 overflow-y-auto">
        <div className="container max-w-7xl mx-auto px-8 py-8">

          <div className="mb-6">
            <Link href="/projects" className="text-sm text-muted-foreground hover:text-foreground flex items-center gap-1 mb-2 transition-colors">
              <ArrowLeft className="w-4 h-4" /> Back to Projects
            </Link>
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <h1 className="text-3xl font-bold tracking-tight" data-testid="text-project-name">{project.name}</h1>
                <StatusBadge status={project.status} />
              </div>
              <div className="flex gap-2">
                 {project.status === 'draft' && (
                    <Button onClick={() => updateProject({ id, status: 'active' })} data-testid="button-start-migration">
                        Start Migration
                    </Button>
                 )}
                 {project.status === 'active' && (
                    <Button variant="outline" onClick={() => updateProject({ id, status: 'completed' })} data-testid="button-mark-complete">
                        Mark Complete
                    </Button>
                 )}
              </div>
            </div>
            <p className="text-muted-foreground mt-2">{project.description}</p>
          </div>

          {/* Service sidebar + content */}
          <div className="flex gap-0 border border-border/50 rounded-xl overflow-hidden bg-card shadow-sm min-h-[calc(100vh-14rem)]">

            {/* Left service navigation */}
            <div className="w-56 flex-shrink-0 border-r border-border/50 bg-muted/10 flex flex-col">
              {/* Overview */}
              <div className="p-3 border-b border-border/50">
                <button
                  onClick={() => setCurrentView('overview')}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm font-medium transition-colors ${currentView === 'overview' ? 'bg-primary text-primary-foreground' : 'hover:bg-accent text-muted-foreground hover:text-foreground'}`}
                  data-testid="nav-overview"
                >
                  <Globe className="w-4 h-4 flex-shrink-0" />
                  Overview
                </button>
              </div>

              {/* Migration services */}
              <div className="p-3 flex-1 space-y-4">
                <div>
                  <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider px-3 mb-1.5">Migration</p>
                  <div className="space-y-0.5">
                    {MIGRATION_SERVICES.map(({ key, label, icon: Icon, itemTypes }) => {
                      const count = items?.filter(i => (itemTypes as readonly string[]).includes(i.itemType)).length || 0;
                      const pending = items?.filter(i => (itemTypes as readonly string[]).includes(i.itemType) && (i.status === 'pending' || i.status === 'failed' || i.status === 'needs_action')).length || 0;
                      const active = currentView === key;
                      return (
                        <button
                          key={key}
                          onClick={() => setCurrentView(key as ViewType)}
                          className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors ${active ? 'bg-primary text-primary-foreground' : 'hover:bg-accent text-muted-foreground hover:text-foreground'}`}
                          data-testid={`nav-service-${key}`}
                        >
                          <Icon className="w-4 h-4 flex-shrink-0" />
                          <span className="flex-1 text-left truncate">{label}</span>
                          {count > 0 && (
                            <span className={`text-xs rounded-full px-1.5 py-0.5 font-medium min-w-[1.25rem] text-center ${active ? 'bg-primary-foreground/20 text-primary-foreground' : pending > 0 ? 'bg-amber-100 dark:bg-amber-900/40 text-amber-700 dark:text-amber-400' : 'bg-muted text-muted-foreground'}`}>
                              {count}
                            </span>
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div>
                  <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider px-3 mb-1.5">Tools</p>
                  <div className="space-y-0.5">
                    {([
                      { key: 'discovery', label: 'Discovery', icon: Search },
                      { key: 'mapping', label: 'Mapping Rules', icon: MapPin },
                      { key: 'tenant_config', label: 'Tenant Config', icon: Shield },
                    ] as const).map(({ key, label, icon: Icon }) => (
                      <button
                        key={key}
                        onClick={() => setCurrentView(key)}
                        className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors ${currentView === key ? 'bg-primary text-primary-foreground' : 'hover:bg-accent text-muted-foreground hover:text-foreground'}`}
                        data-testid={`nav-tool-${key}`}
                      >
                        <Icon className="w-4 h-4 flex-shrink-0" />
                        {label}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              {/* Bottom: overall progress */}
              <div className="p-3 border-t border-border/50 space-y-1.5">
                <div className="flex justify-between text-xs text-muted-foreground">
                  <span>{stats?.total || 0} total</span>
                  <span className="text-emerald-600 dark:text-emerald-400 font-medium">{stats?.completed || 0} done</span>
                </div>
                <Progress value={stats?.total ? Math.round(((stats.completed || 0) / stats.total) * 100) : 0} className="h-1.5" />
              </div>
            </div>

            {/* Main content */}
            <div className="flex-1 overflow-y-auto p-6 min-w-0">

              {/* Overview */}
              {currentView === 'overview' && (
                <div className="space-y-6">
                  {/* Export toolbar */}
                  {items && items.length > 0 && (
                    <div className="flex items-center gap-2 p-3 rounded-lg bg-muted/50 border border-border">
                      <Download className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                      <span className="text-sm font-medium flex-1">Export migration logs</span>
                      <Button variant="outline" size="sm" onClick={() => handleExportLogs('txt')} data-testid="button-export-logs-txt">
                        <Download className="w-3.5 h-3.5 mr-1.5" /> Plain text (.txt)
                      </Button>
                      <Button variant="outline" size="sm" onClick={() => handleExportLogs('csv')} data-testid="button-export-logs-csv">
                        <Download className="w-3.5 h-3.5 mr-1.5" /> Spreadsheet (.csv)
                      </Button>
                    </div>
                  )}
                  <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                    {[
                      { label: 'Total Items', value: stats?.total || 0, color: '' },
                      { label: 'Completed', value: stats?.completed || 0, color: 'text-emerald-600' },
                      { label: 'In Progress', value: stats?.inProgress || 0, color: 'text-blue-600' },
                      { label: 'Failed', value: stats?.failed || 0, color: 'text-red-600' },
                    ].map(({ label, value, color }) => (
                      <Card key={label} className="shadow-sm">
                        <CardHeader className="pb-2"><CardTitle className="text-sm font-medium text-muted-foreground">{label}</CardTitle></CardHeader>
                        <CardContent><div className={`text-3xl font-bold ${color}`}>{value}</div></CardContent>
                      </Card>
                    ))}
                  </div>

                  {/* Continuous Sync Panel */}
                  <Card className="shadow-sm">
                    <CardHeader className="pb-3">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-2">
                          <RefreshCw className="w-4 h-4 text-primary" />
                          <CardTitle className="text-base">Continuous Sync</CardTitle>
                          {syncStatus?.syncEnabled && (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 text-xs font-medium rounded-full bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400">
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                              Active
                            </span>
                          )}
                        </div>
                        <div className="flex items-center gap-3">
                          {syncSaving && <Loader2 className="w-4 h-4 animate-spin text-muted-foreground" />}
                          <Switch
                            checked={syncStatus?.syncEnabled ?? false}
                            onCheckedChange={handleToggleSync}
                            disabled={syncSaving}
                            data-testid="switch-sync-enabled"
                          />
                        </div>
                      </div>
                      <CardDescription className="text-xs mt-1 space-y-1">
                        <span className="block">Automatically copies new emails and file changes from source to target after migration completes. Covers: mailboxes, shared mailboxes, OneDrive, SharePoint.</span>
                        <span className="block text-amber-600 dark:text-amber-400">Requires <code className="bg-amber-50 dark:bg-amber-950/30 px-1 rounded">Mail.ReadWrite</code> and <code className="bg-amber-50 dark:bg-amber-950/30 px-1 rounded">Files.ReadWrite.All</code> Application permissions on both tenant app registrations.</span>
                      </CardDescription>
                    </CardHeader>
                    <CardContent className="space-y-4">
                      <div className="flex flex-wrap items-center gap-4">
                        <div className="flex items-center gap-2">
                          <Label className="text-sm text-muted-foreground whitespace-nowrap">Check every</Label>
                          <Select
                            value={String(syncStatus?.syncIntervalMinutes ?? 60)}
                            onValueChange={handleSyncIntervalChange}
                            disabled={syncSaving}
                          >
                            <SelectTrigger className="w-36 h-8 text-sm" data-testid="select-sync-interval">
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="15">15 minutes</SelectItem>
                              <SelectItem value="30">30 minutes</SelectItem>
                              <SelectItem value="60">1 hour</SelectItem>
                              <SelectItem value="240">4 hours</SelectItem>
                              <SelectItem value="720">12 hours</SelectItem>
                              <SelectItem value="1440">24 hours</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={handleSyncNow}
                          disabled={syncRunning || !syncStatus?.completedItems}
                          data-testid="button-sync-now"
                        >
                          {syncRunning
                            ? <><Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" /> Syncing…</>
                            : <><RefreshCw className="w-3.5 h-3.5 mr-1.5" /> Sync now</>
                          }
                        </Button>
                      </div>

                      {/* Per-item sync state */}
                      {(syncStatus as any)?.needsActionCount > 0 && (
                        <div className="flex items-start gap-2 rounded-lg border border-amber-200 dark:border-amber-800 bg-amber-50 dark:bg-amber-950/20 px-3 py-2 text-xs text-amber-700 dark:text-amber-400">
                          <AlertTriangle className="w-3.5 h-3.5 flex-shrink-0 mt-0.5" />
                          <span>{(syncStatus as any).needsActionCount} item(s) have status "Needs Action" and won't sync until fully migrated. Go to the Shared Mailboxes tab, complete the conversion, then re-run migration.</span>
                        </div>
                      )}

                      {syncStatus && syncStatus.items.length > 0 ? (
                        <div className="space-y-1.5">
                          <p className="text-xs font-medium text-muted-foreground uppercase tracking-wide">Syncing {syncStatus.items.length} completed item(s)</p>
                          <div className="rounded-lg border border-border divide-y divide-border overflow-hidden">
                            {syncStatus.items.map(item => (
                              <div key={item.id} className="flex items-center gap-3 px-3 py-2 text-sm bg-background hover:bg-muted/30">
                                <ItemTypeIcon type={item.itemType} />
                                <span className="flex-1 truncate font-medium">{item.sourceIdentity}</span>
                                {item.targetIdentity && item.targetIdentity !== item.sourceIdentity && (
                                  <span className="text-xs text-muted-foreground truncate max-w-[140px]">→ {item.targetIdentity}</span>
                                )}
                                <div className="text-right shrink-0">
                                  {item.lastSyncedAt ? (
                                    <p className="text-xs text-muted-foreground">Last synced {format(new Date(item.lastSyncedAt), 'MMM d, HH:mm')}</p>
                                  ) : (
                                    <p className="text-xs text-muted-foreground italic">Not yet synced</p>
                                  )}
                                  {item.nextSyncAt && syncStatus.syncEnabled && (
                                    <p className="text-xs text-emerald-600 dark:text-emerald-400">Next: {format(new Date(item.nextSyncAt), 'MMM d, HH:mm')}</p>
                                  )}
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      ) : (
                        <p className="text-xs text-muted-foreground italic">
                          {syncStatus?.completedItems === 0
                            ? 'Complete initial migrations first — sync only runs on finished items (mailbox, OneDrive, SharePoint).'
                            : 'No syncable completed items yet.'}
                        </p>
                      )}
                    </CardContent>
                  </Card>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {/* Service breakdown */}
                    <Card>
                      <CardHeader><CardTitle className="text-base">By Service</CardTitle></CardHeader>
                      <CardContent className="space-y-3">
                        {MIGRATION_SERVICES.map(({ key, label, icon: Icon, itemTypes }) => {
                          const svcItems = items?.filter(i => (itemTypes as readonly string[]).includes(i.itemType)) || [];
                          const done = svcItems.filter(i => i.status === 'completed').length;
                          const total = svcItems.length;
                          if (total === 0) return null;
                          return (
                            <button key={key} onClick={() => setCurrentView(key as ViewType)} className="w-full flex items-center gap-3 hover:bg-muted/50 rounded-lg p-2 -mx-2 transition-colors text-left">
                              <Icon className="w-4 h-4 text-muted-foreground flex-shrink-0" />
                              <span className="flex-1 text-sm font-medium">{label}</span>
                              <span className="text-xs text-muted-foreground">{done}/{total}</span>
                              <div className="w-20"><Progress value={total ? Math.round((done / total) * 100) : 0} className="h-1.5" /></div>
                            </button>
                          );
                        })}
                        {(!items || items.length === 0) && <p className="text-sm text-muted-foreground">No items yet. Open a service to add items.</p>}
                      </CardContent>
                    </Card>

                    {/* Chart + tenant details */}
                    <div className="space-y-4">
                      <Card className="h-[200px]">
                        <CardContent className="h-full pt-4">
                          {chartData.length > 0 ? (
                            <ResponsiveContainer width="100%" height="100%">
                              <PieChart>
                                <Pie data={chartData} cx="50%" cy="50%" innerRadius={50} outerRadius={70} paddingAngle={4} dataKey="value">
                                  {chartData.map((entry, i) => <Cell key={i} fill={entry.color} />)}
                                </Pie>
                                <Tooltip />
                              </PieChart>
                            </ResponsiveContainer>
                          ) : (
                            <div className="flex h-full items-center justify-center text-muted-foreground text-sm">No data yet</div>
                          )}
                        </CardContent>
                      </Card>
                      <Card>
                        <CardHeader><CardTitle className="text-base">Tenant Details</CardTitle></CardHeader>
                        <CardContent className="space-y-2 text-sm">
                          <div className="flex justify-between py-1.5 border-b border-border/50">
                            <span className="text-muted-foreground">Source</span>
                            <span className="font-mono text-xs" data-testid="text-source-tenant">{project.sourceTenantId}</span>
                          </div>
                          <div className="flex justify-between py-1.5 border-b border-border/50">
                            <span className="text-muted-foreground">Target</span>
                            <span className="font-mono text-xs" data-testid="text-target-tenant">{project.targetTenantId}</span>
                          </div>
                          <div className="flex justify-between py-1.5">
                            <span className="text-muted-foreground">Created</span>
                            <span>{project.createdAt ? format(new Date(project.createdAt), 'PP') : '-'}</span>
                          </div>
                        </CardContent>
                      </Card>
                    </div>
                  </div>
                </div>
              )}

              {/* Service dashboards */}
              {MIGRATION_SERVICES.map(({ key, label, icon: Icon, itemTypes, description }) => {
                if (currentView !== key) return null;
                if (key === 'entra_ad') {
                  return <EntraToAdTab key={key} projectId={id} project={project} />;
                }
                const svcItems = (items || []).filter(i => (itemTypes as readonly string[]).includes(i.itemType));
                const pending = svcItems.filter(i => i.status === 'pending' || i.status === 'failed' || i.status === 'needs_action').length;
                const svcItemType = itemTypes[0] as string;
                return (
                  <div key={key} className="space-y-5">
                    {/* Service header */}
                    <div className="flex items-start justify-between gap-4">
                      <div className="flex items-center gap-3">
                        <div className="p-2 rounded-lg bg-primary/10">
                          <Icon className="w-5 h-5 text-primary" />
                        </div>
                        <div>
                          <h2 className="text-xl font-bold">{label}</h2>
                          <p className="text-sm text-muted-foreground">{description}</p>
                        </div>
                      </div>
                      <div className="flex gap-2 flex-shrink-0 flex-wrap">
                        {selectedServiceItems.size > 0 && svcItems.some(i => selectedServiceItems.has(i.id)) && (
                          <Button variant="destructive" size="sm" disabled={isDeletingSelected} onClick={() => handleDeleteSelected(svcItems.map(i => i.id))} data-testid={`button-delete-selected-${key}`}>
                            {isDeletingSelected ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : <Trash2 className="w-4 h-4 mr-2" />}
                            Delete {svcItems.filter(i => selectedServiceItems.has(i.id)).length} selected
                          </Button>
                        )}
                        {pending > 0 && (
                          <Button onClick={async () => {
                            try {
                              await Promise.all(
                                svcItems.filter(i => i.status === 'pending' || i.status === 'failed' || i.status === 'needs_action')
                                  .map(i => apiRequest('POST', `/api/projects/${id}/items/${i.id}/migrate`))
                              );
                              toast({ title: "Migration started", description: `Running ${pending} ${label} item(s).` });
                              queryClient.invalidateQueries({ queryKey: [api.items.list.path, id] });
                            } catch { toast({ title: "Error", description: "Failed to start migration", variant: "destructive" }); }
                          }} data-testid={`button-migrate-all-${key}`}>
                            <PlayCircle className="w-4 h-4 mr-2" /> Run All ({pending})
                          </Button>
                        )}
                        <Button variant="outline" onClick={() => {
                          form.setValue('itemType', svcItemType);
                          form.setValue('sourceIdentity', '');
                          form.setValue('targetIdentity', '');
                          setIsAddOpen(true);
                        }} data-testid={`button-add-${key}`}>
                          <Plus className="w-4 h-4 mr-2" /> Add Item
                        </Button>
                        {svcItems.length > 0 && (
                          <>
                            <Button variant="ghost" size="sm" onClick={() => handleExportLogs('txt')} title="Export logs as plain text" data-testid={`button-export-txt-${key}`}>
                              <Download className="w-4 h-4 mr-1" /> .txt
                            </Button>
                            <Button variant="ghost" size="sm" onClick={() => handleExportLogs('csv')} title="Export logs as CSV (Excel)" data-testid={`button-export-csv-${key}`}>
                              <Download className="w-4 h-4 mr-1" /> .csv
                            </Button>
                          </>
                        )}
                      </div>
                    </div>

                    {/* Service stats */}
                    <div className="grid grid-cols-5 gap-3">
                      {[
                        { label: 'Total', value: svcItems.length, color: '' },
                        { label: 'Pending', value: svcItems.filter(i => i.status === 'pending').length, color: 'text-slate-600' },
                        { label: 'Completed', value: svcItems.filter(i => i.status === 'completed').length, color: 'text-emerald-600' },
                        { label: 'Needs Action', value: svcItems.filter(i => i.status === 'needs_action').length, color: 'text-amber-600' },
                        { label: 'Failed', value: svcItems.filter(i => i.status === 'failed').length, color: 'text-red-600' },
                      ].map(({ label: sl, value, color }) => (
                        <Card key={sl} className="shadow-sm">
                          <CardHeader className="pb-1 pt-4 px-4"><CardTitle className="text-xs font-medium text-muted-foreground">{sl}</CardTitle></CardHeader>
                          <CardContent className="px-4 pb-4"><div className={`text-2xl font-bold ${color}`}>{value}</div></CardContent>
                        </Card>
                      ))}
                    </div>

                    {/* Distribution Groups: API limitation warning */}
                    {key === 'distributiongroups' && (
                      <div className="flex gap-3 p-3.5 rounded-lg border border-amber-200 dark:border-amber-800 bg-amber-50 dark:bg-amber-950/30 text-sm">
                        <AlertTriangle className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0 mt-0.5" />
                        <div className="space-y-1">
                          <p className="font-medium text-amber-800 dark:text-amber-300">Graph API cannot create classic distribution lists</p>
                          <p className="text-amber-700 dark:text-amber-400 text-xs leading-relaxed">
                            Microsoft's Graph API does not support creating DLs or mail-enabled security groups directly.
                            By default, this tool attempts MESG creation first and <strong>automatically falls back to an M365 Unified Group</strong> if the API rejects it —
                            so migrations will always complete.
                            Toggle the <Building2 className="w-3 h-3 inline" /> button on any item to <strong>skip the MESG attempt</strong> and go straight to M365 Unified Group.
                          </p>
                        </div>
                      </div>
                    )}

                    {/* Items table */}
                    <div className="bg-background rounded-lg border border-border/60 shadow-sm overflow-hidden">
                      {itemsLoading ? (
                        <div className="p-8 flex justify-center"><Loader2 className="animate-spin" /></div>
                      ) : svcItems.length > 0 ? (
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="bg-muted/30 border-b border-border/60">
                              <th className="px-3 py-3 w-10">
                                <input
                                  type="checkbox"
                                  className="rounded border-border"
                                  data-testid={`checkbox-select-all-${key}`}
                                  checked={svcItems.length > 0 && svcItems.filter(i => i.status !== 'in_progress').every(i => selectedServiceItems.has(i.id))}
                                  onChange={e => {
                                    const deletable = svcItems.filter(i => i.status !== 'in_progress').map(i => i.id);
                                    setSelectedServiceItems(prev => {
                                      const next = new Set(prev);
                                      if (e.target.checked) deletable.forEach(i => next.add(i));
                                      else deletable.forEach(i => next.delete(i));
                                      return next;
                                    });
                                  }}
                                />
                              </th>
                              <th className="px-5 py-3 text-left font-semibold text-muted-foreground">Source</th>
                              <th className="px-5 py-3 text-left font-semibold text-muted-foreground">Target</th>
                              <th className="px-5 py-3 text-left font-semibold text-muted-foreground w-52">Status</th>
                              <th className="px-5 py-3 text-left font-semibold text-muted-foreground">Progress</th>
                              <th className="px-5 py-3 text-right font-semibold text-muted-foreground">Actions</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-border/40">
                            {svcItems.map((item: MigrationItem) => {
                              const hasBytesData = item.bytesTotal != null && item.bytesTotal > 0;
                              const pct = item.progressPercent ?? 0;
                              const isSelected = selectedServiceItems.has(item.id);
                              return (
                              <tr key={item.id} className={`hover:bg-muted/20 ${item.status === 'in_progress' ? 'bg-blue-50/40 dark:bg-blue-950/20' : ''} ${item.status === 'needs_action' ? 'bg-amber-50/50 dark:bg-amber-950/20' : ''} ${isSelected ? 'bg-red-50/30 dark:bg-red-950/20' : ''}`} data-testid={`row-item-${item.id}`}>
                                <td className="px-3 py-3.5">
                                  <input
                                    type="checkbox"
                                    className="rounded border-border"
                                    disabled={item.status === 'in_progress'}
                                    checked={isSelected}
                                    data-testid={`checkbox-item-${item.id}`}
                                    onChange={e => setSelectedServiceItems(prev => {
                                      const next = new Set(prev);
                                      if (e.target.checked) next.add(item.id); else next.delete(item.id);
                                      return next;
                                    })}
                                  />
                                </td>
                                <td className="px-5 py-3.5 font-medium" data-testid={`text-source-${item.id}`}>{item.sourceIdentity}</td>
                                <td className="px-5 py-3.5 text-muted-foreground text-sm" data-testid={`text-target-${item.id}`}>{item.targetIdentity || 'Auto-mapped'}</td>
                                <td className="px-5 py-3.5">
                                  <div className="flex items-center gap-2">
                                    <StatusBadge status={item.status} />
                                    {item.status === 'in_progress' && <Loader2 className="w-3.5 h-3.5 animate-spin text-blue-500 flex-shrink-0" />}
                                  </div>
                                  {item.status === 'failed' && item.errorDetails && (
                                    <div className="text-xs text-red-500 mt-1 max-w-[180px] truncate" title={item.errorDetails} data-testid={`text-error-${item.id}`}>{item.errorDetails}</div>
                                  )}
                                  {item.status === 'needs_action' && (
                                    <div className="text-xs text-amber-600 dark:text-amber-400 mt-1 max-w-[200px] truncate" title={item.errorDetails || 'Manual action required — open logs for details'} data-testid={`text-needs-action-${item.id}`}>
                                      ⚠ Action required — open logs
                                    </div>
                                  )}
                                </td>
                                <td className="px-5 py-3.5 min-w-[220px]" data-testid={`progress-cell-${item.id}`}>
                                  {item.status === 'in_progress' && (
                                    <div className="space-y-1.5">
                                      {hasBytesData ? (
                                        <>
                                          <div className="flex items-center justify-between text-xs font-medium">
                                            <span className="text-blue-600 dark:text-blue-400 tabular-nums">{formatBytes(item.bytesMigrated)} / {formatBytes(item.bytesTotal)}</span>
                                            <span className="text-blue-700 dark:text-blue-300 font-bold tabular-nums ml-3" data-testid={`text-pct-${item.id}`}>{pct}%</span>
                                          </div>
                                          <Progress value={pct} className="h-2" data-testid={`progress-bar-${item.id}`} />
                                          <p className="text-[11px] text-muted-foreground">Transferring data…</p>
                                        </>
                                      ) : (
                                        <>
                                          <div className="h-2 rounded-full bg-blue-100 dark:bg-blue-900/40 overflow-hidden">
                                            <div className="h-full bg-blue-500 rounded-full animate-pulse w-1/2" />
                                          </div>
                                          <p className="text-[11px] text-muted-foreground">Working…</p>
                                        </>
                                      )}
                                    </div>
                                  )}
                                  {item.status === 'completed' && (
                                    <div className="space-y-1">
                                      <Progress value={100} className="h-2" />
                                      <p className="text-[11px] text-emerald-600 dark:text-emerald-400 font-medium" data-testid={`bytes-${item.id}`}>
                                        {hasBytesData ? `${formatBytes(item.bytesMigrated)} migrated` : 'Completed'}
                                      </p>
                                    </div>
                                  )}
                                  {item.status === 'pending' && (
                                    <p className="text-[11px] text-muted-foreground">Queued</p>
                                  )}
                                  {item.status === 'failed' && (
                                    <p className="text-[11px] text-red-500">Migration failed</p>
                                  )}
                                </td>
                                <td className="px-5 py-3 text-right">
                                  <div className="flex items-center justify-end gap-1">
                                    {/* Distribution group: toggle M365 Group mode (skip MESG attempt) */}
                                    {item.itemType === 'distributiongroup' && item.status !== 'in_progress' && (() => {
                                      const m365On = !!(item.options as any)?.allowM365Upgrade;
                                      return (
                                        <Button
                                          size="sm"
                                          variant={m365On ? 'secondary' : 'ghost'}
                                          title={m365On ? 'M365 Group mode ON — will skip MESG attempt and create directly as M365 Unified Group. Click to switch back to auto mode.' : 'Auto mode: attempts MESG, falls back to M365 Unified Group if rejected. Click to force M365 Group directly.'}
                                          className={`gap-1.5 text-xs ${m365On ? 'text-cyan-700 dark:text-cyan-300 font-medium' : 'text-muted-foreground/60'}`}
                                          data-testid={`button-m365upgrade-${item.id}`}
                                          onClick={() => updateItem({ id: item.id, options: { ...(item.options as any || {}), allowM365Upgrade: !m365On } })}
                                        >
                                          <Building2 className="w-3.5 h-3.5" />
                                          {m365On ? 'M365' : 'Auto'}
                                        </Button>
                                      );
                                    })()}
                                    {(item.status === 'pending' || item.status === 'failed' || item.status === 'needs_action') && (
                                      <Button size="sm" variant="ghost" onClick={() => handleMigrateItem(item.id)} title="Start / Retry" data-testid={`button-migrate-${item.id}`}>
                                        <Play className="w-4 h-4 text-emerald-600" />
                                      </Button>
                                    )}
                                    {(item.status === 'failed' || item.status === 'needs_action') && (
                                      <Button size="sm" variant="ghost" onClick={() => handleRetry(item.id)} title="Reset to pending" data-testid={`button-retry-${item.id}`}>
                                        <RotateCw className="w-4 h-4 text-blue-600" />
                                      </Button>
                                    )}
                                    {(item.status === 'in_progress' || item.status === 'completed' || item.status === 'failed' || item.status === 'needs_action') && (
                                      <Button size="sm" variant="ghost" onClick={() => handleViewLogs(item)} title="Logs" data-testid={`button-logs-${item.id}`}>
                                        <FileText className="w-4 h-4 text-slate-500" />
                                      </Button>
                                    )}
                                    {item.status !== 'in_progress' && (
                                      <Button size="sm" variant="ghost" onClick={() => handleDelete(item.id)} title="Delete" data-testid={`button-delete-${item.id}`}>
                                        <Trash2 className="w-4 h-4 text-red-500" />
                                      </Button>
                                    )}
                                  </div>
                                </td>
                              </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      ) : (
                        <div className="p-10 text-center space-y-3">
                          <Icon className="w-10 h-10 text-muted-foreground/30 mx-auto" />
                          <p className="text-muted-foreground text-sm">No {label} items yet.</p>
                          <Button size="sm" variant="outline" onClick={() => {
                            form.setValue('itemType', svcItemType);
                            form.setValue('sourceIdentity', '');
                            form.setValue('targetIdentity', '');
                            setIsAddOpen(true);
                          }}>
                            <Plus className="w-4 h-4 mr-2" /> Add first item
                          </Button>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}

              {/* Tools */}
              {currentView === 'discovery' && (
                <DiscoveryTab projectId={id} existingItems={items || []} onImport={(newItems) => {
                  newItems.forEach(item => createItem(item).catch(() => {}));
                  queryClient.invalidateQueries({ queryKey: [api.items.list.path, id] });
                  const users = newItems.filter((i: any) => i.itemType === 'user').length;
                  const mailboxes = newItems.filter((i: any) => i.itemType === 'mailbox').length;
                  const onedrives = newItems.filter((i: any) => i.itemType === 'onedrive').length;
                  const other = newItems.length - users - mailboxes - onedrives;
                  const parts = [
                    users > 0 && `${users} user${users > 1 ? 's' : ''}`,
                    mailboxes > 0 && `${mailboxes} mailbox${mailboxes > 1 ? 'es' : ''}`,
                    onedrives > 0 && `${onedrives} OneDrive${onedrives > 1 ? 's' : ''}`,
                    other > 0 && `${other} other item${other > 1 ? 's' : ''}`,
                  ].filter(Boolean).join(', ');
                  toast({ title: "Imported", description: `Added to migration queue: ${parts || `${newItems.length} item(s)`}.` });
                }} />
              )}
              {currentView === 'mapping' && <MappingRulesTab projectId={id} />}
              {currentView === 'tenant_config' && <TenantConfigTab projectId={id} project={project} />}
            </div>
          </div>

          {/* Add item dialog (shared, type pre-set by service) */}
          <Dialog open={isAddOpen} onOpenChange={setIsAddOpen}>
            <DialogContent>
              <DialogHeader><DialogTitle>Add Migration Item</DialogTitle></DialogHeader>
              <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 py-4">
                <div className="space-y-2">
                  <Label>Item Type</Label>
                  <Controller control={form.control} name="itemType" render={({ field }) => (
                    <Select onValueChange={field.onChange} value={field.value}>
                      <SelectTrigger data-testid="select-item-type"><SelectValue placeholder="Select type" /></SelectTrigger>
                      <SelectContent>
                        <SelectItem value="mailbox">Mailbox (Exchange Online)</SelectItem>
                        <SelectItem value="sharepoint">SharePoint Online</SelectItem>
                        <SelectItem value="onedrive">OneDrive</SelectItem>
                        <SelectItem value="teams">Microsoft Teams</SelectItem>
                        <SelectItem value="user">User Account</SelectItem>
                        <SelectItem value="distributiongroup">Distribution Group</SelectItem>
                        <SelectItem value="sharedmailbox">Shared Mailbox</SelectItem>
                        <SelectItem value="m365group">Microsoft 365 Group</SelectItem>
                        <SelectItem value="powerplatform">Power Platform</SelectItem>
                      </SelectContent>
                    </Select>
                  )} />
                </div>
                <div className="space-y-2">
                  <Label>Source Identity</Label>
                  <Input {...form.register("sourceIdentity")} placeholder={form.watch("itemType") === "sharepoint" ? "contoso.sharepoint.com:/sites/Team" : "user@source.com"} data-testid="input-source-identity" />
                  {form.formState.errors.sourceIdentity && <p className="text-xs text-red-500">{form.formState.errors.sourceIdentity.message}</p>}
                </div>
                <div className="space-y-2">
                  <Label>Target Identity</Label>
                  <Input {...form.register("targetIdentity")} placeholder={form.watch("itemType") === "sharepoint" ? "fabrikam.sharepoint.com:/sites/Team" : "user@target.com"} data-testid="input-target-identity" />
                  {form.formState.errors.targetIdentity && <p className="text-xs text-red-500">{form.formState.errors.targetIdentity.message}</p>}
                </div>
                <div className="flex justify-end pt-2">
                  <Button type="submit" data-testid="button-submit-item">Add Item</Button>
                </div>
              </form>
            </DialogContent>
          </Dialog>
        </div>
      </main>

      <Dialog open={!!logsDialogItem} onOpenChange={(open) => { if (!open) setLogsDialogItem(null); }}>
        <DialogContent className="max-w-2xl max-h-[80vh]">
          <DialogHeader>
            <DialogTitle>
              Migration Logs — {logsDialogItem?.sourceIdentity}
              {logsDialogItem?.targetIdentity && ` → ${logsDialogItem.targetIdentity}`}
            </DialogTitle>
          </DialogHeader>
          <div className="overflow-y-auto max-h-[60vh] bg-slate-950 text-slate-200 rounded-lg p-4 font-mono text-xs leading-relaxed" data-testid="container-logs">
            {logsLoading ? (
              <div className="flex justify-center py-8"><Loader2 className="animate-spin text-slate-400" /></div>
            ) : itemLogs.length > 0 ? (
              itemLogs.map((log, i) => (
                <div key={i} className={`py-0.5 ${log.includes('failed') || log.includes('Failed') || log.includes('Error') ? 'text-red-400' : log.includes('ACTION REQUIRED') || log.includes('Needs Action') || log.includes('⚠') ? 'text-amber-300' : log.includes('complete') || log.includes('Complete') || log.includes('success') || log.includes('✓') ? 'text-emerald-400' : ''}`}>
                  {log}
                </div>
              ))
            ) : (
              <div className="text-slate-500 text-center py-8">No logs available yet.</div>
            )}
          </div>
          {logsDialogItem?.status === 'in_progress' && (
            <div className="flex items-center gap-2 text-sm text-blue-500">
              <Loader2 className="w-3 h-3 animate-spin" />
              Migration in progress — logs will update when you reopen this dialog.
            </div>
          )}
        </DialogContent>
      </Dialog>
    </div>
  );
}

const SERVICE_ICONS: Record<string, any> = { Mail, Cloud, Users, UserCheck };

const SERVICES = [
  { key: 'exchange', label: 'Exchange Online', description: 'Mailbox & email migration', icon: 'Mail' },
  { key: 'sharepoint', label: 'SharePoint & OneDrive', description: 'Sites & files migration', icon: 'Cloud' },
  { key: 'teams', label: 'Microsoft Teams', description: 'Teams & channels migration', icon: 'Users' },
  { key: 'users', label: 'Users & Directory', description: 'User accounts & directory', icon: 'UserCheck' },
] as const;

function TenantCredentialForm({
  label,
  tenantType,
  projectId,
  tenantId,
  clientId,
  clientSecret,
  consentedServices,
}: {
  label: string;
  tenantType: 'source' | 'target';
  projectId: number;
  tenantId: string;
  clientId: string | null;
  clientSecret: string | null;
  consentedServices?: string | null;
}) {
  const { toast } = useToast();
  const queryClient = useQueryClient();
  const { mutateAsync: updateProject, isPending: isSaving } = useUpdateProject();
  const [showSecret, setShowSecret] = useState(false);
  const [localTenantId, setLocalTenantId] = useState(tenantId || '');
  const [localClientId, setLocalClientId] = useState(clientId || '');
  const [localClientSecret, setLocalClientSecret] = useState('');
  const hasExistingSecret = !!clientSecret;
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [isTesting, setIsTesting] = useState(false);
  const [isRegranting, setIsRegranting] = useState(false);

  // Initialise from DB-persisted consent, keep local optimistic updates in sync
  const persistedServices: string[] = (() => {
    try { return JSON.parse(consentedServices || '[]'); } catch { return []; }
  })();
  const [grantedServices, setGrantedServices] = useState<Set<string>>(new Set(persistedServices));

  // Keep local state in sync when the project data refreshes from server (e.g. after postMessage invalidation)
  const [prevConsentedServices, setPrevConsentedServices] = useState(consentedServices);
  if (consentedServices !== prevConsentedServices) {
    setPrevConsentedServices(consentedServices);
    setGrantedServices(prev => {
      const fresh = new Set<string>(persistedServices);
      prev.forEach(s => fresh.add(s)); // keep any locally granted services too
      return fresh;
    });
  }

  const isConnected = !!(clientId && clientSecret);
  const allServiceKeys = SERVICES.map(s => s.key);
  const allGranted = allServiceKeys.every(k => grantedServices.has(k));

  // Listen for postMessage from the consent popup — update local state & refresh project
  const onConsentMessage = (projectId: number) => (e: MessageEvent) => {
    if (e.origin !== window.location.origin) return;
    if (e.data?.type === 'consent_success') {
      const all = new Set(allServiceKeys);
      setGrantedServices(all);
      // Refresh project from server so persisted consent is loaded (key matches useProject hook)
      queryClient.invalidateQueries({ queryKey: ['/api/projects/:id', projectId] });
    } else if (e.data?.type === 'consent_error') {
      toast({ title: "Consent error", description: e.data.error, variant: "destructive" });
    }
  };

  const openConsentPopup = (url: string, serviceKey: string) => {
    const handler = onConsentMessage(projectId);
    window.addEventListener('message', handler, { once: true });
    const popup = window.open(url, `consent_${serviceKey}`, 'width=700,height=700,scrollbars=yes');
    if (!popup) {
      window.removeEventListener('message', handler);
      window.open(url, '_blank');
      toast({ title: "Popup blocked", description: "Consent page opened in a new tab. Grant permissions there, then return here." });
    } else {
      toast({ title: "Select your Global Admin account", description: "Choose your Global Admin account in the popup, then click Accept." });
    }
  };

  const handleConnect = () => {
    if (!localTenantId?.trim()) {
      toast({ title: "Tenant ID required", description: "Enter a Tenant ID before connecting.", variant: "destructive" });
      return;
    }
    const params = new URLSearchParams({
      projectId: String(projectId),
      tenantType,
      tenantId: localTenantId,
      appName: `Tenant Migration Tool - ${label}`,
    });
    window.location.href = `/api/oauth/connect?${params}`;
  };

  // Re-grant using the EXISTING app registration — no new OAuth flow
  const handleRegrantAll = async () => {
    setIsRegranting(true);
    try {
      const res = await fetch(`/api/oauth/regrant-url?projectId=${projectId}&tenantType=${tenantType}`);
      const data = await res.json() as any;
      if (data.url) {
        openConsentPopup(data.url, 'regrant');
      } else {
        toast({ title: "Error", description: data.message || "Could not build re-grant URL.", variant: "destructive" });
      }
    } catch {
      toast({ title: "Error", description: "Could not build re-grant URL.", variant: "destructive" });
    } finally {
      setIsRegranting(false);
    }
  };

  const handleGrantService = async (serviceKey: string) => {
    if (!localTenantId || !localClientId) return;
    try {
      const params = new URLSearchParams({
        tenantId: localTenantId,
        clientId: localClientId,
        service: serviceKey,
        projectId: String(projectId),
        tenantType,
      });
      const res = await fetch(`/api/oauth/consent-url?${params}`);
      const data = await res.json() as any;
      if (data.url) {
        openConsentPopup(data.url, serviceKey);
      }
    } catch {
      toast({ title: "Error", description: "Could not build consent URL.", variant: "destructive" });
    }
  };

  const handleSave = async () => {
    const secretUpdate = localClientSecret ? (tenantType === 'source' ? { sourceClientSecret: localClientSecret } : { targetClientSecret: localClientSecret }) : {};
    const updates = tenantType === 'source'
      ? { sourceTenantId: localTenantId, sourceClientId: localClientId, ...secretUpdate }
      : { targetTenantId: localTenantId, targetClientId: localClientId, ...secretUpdate };
    try {
      await updateProject({ id: projectId, ...updates });
      toast({ title: "Saved", description: `${label} credentials updated.` });
      setLocalClientSecret('');
      setTestResult(null);
    } catch {
      toast({ title: "Error", description: "Failed to save credentials.", variant: "destructive" });
    }
  };

  const handleTestConnection = async () => {
    setIsTesting(true);
    setTestResult(null);
    try {
      const res = await apiRequest('POST', `/api/projects/${projectId}/test-connection`, { tenant: tenantType });
      const data = await res.json();
      setTestResult(data);
    } catch {
      setTestResult({ success: false, message: 'Failed to reach server.' });
    } finally {
      setIsTesting(false);
    }
  };

  const hasCredentials = localClientId && (localClientSecret || hasExistingSecret);

  return (
    <Card className="shadow-sm">
      <CardHeader>
        <div className="flex items-center justify-between flex-wrap gap-2">
          <div className="flex items-center gap-2">
            <Shield className="w-5 h-5 text-primary" />
            <CardTitle className="text-lg">{label}</CardTitle>
            {isConnected && (
              <span className="inline-flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full bg-emerald-100 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800">
                <CheckCircle2 className="w-3 h-3" /> Connected
              </span>
            )}
            {isConnected && allGranted && (
              <span className="inline-flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full bg-blue-100 dark:bg-blue-950/50 text-blue-700 dark:text-blue-400 border border-blue-200 dark:border-blue-800">
                <CheckCircle2 className="w-3 h-3" /> All permissions granted
              </span>
            )}
          </div>
          <div className="flex items-center gap-2">
            {isConnected && (
              <Button
                onClick={handleRegrantAll}
                size="sm"
                variant="outline"
                className="gap-2"
                disabled={isRegranting}
                data-testid={`button-regrant-${tenantType}`}
              >
                <KeyRound className="w-4 h-4" />
                {isRegranting ? 'Opening…' : 'Re-grant All Permissions'}
              </Button>
            )}
            <Button
              onClick={handleConnect}
              size="sm"
              className="gap-2"
              data-testid={`button-connect-tenant-${tenantType}`}
            >
              <Wand2 className="w-4 h-4" />
              {isConnected ? 'Reconnect (new app reg)' : 'Connect with Microsoft'}
            </Button>
          </div>
        </div>
        <CardDescription>
          {isConnected
            ? allGranted
              ? 'All permissions are granted. Use "Re-grant All Permissions" to renew consent with the same app registration if it ever expires.'
              : 'This tenant is connected. Grant permissions below for each migration service you need, or click "Re-grant All Permissions" to consent to everything at once.'
            : 'Click "Connect with Microsoft" to sign in as a Global Admin and automatically set up the app registration.'}
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* Tenant ID — always visible and editable */}
        <div className="space-y-1.5">
          <Label className="text-sm font-medium">Directory (Tenant) ID</Label>
          <Input
            value={localTenantId}
            onChange={(e) => setLocalTenantId(e.target.value)}
            placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
            className="font-mono"
            autoComplete="off"
            data-testid={`input-${tenantType}-tenant-id-main`}
          />
          <p className="text-xs text-muted-foreground">Found in Microsoft Entra ID &rarr; Overview &rarr; Directory (tenant) ID.</p>
        </div>

        {/* Per-service grant permissions */}
        {isConnected && (
          <div className="space-y-3">
            <Label className="text-sm font-semibold">Grant Permissions by Service</Label>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {SERVICES.map(({ key, label: svcLabel, description, icon }) => {
                const Icon = SERVICE_ICONS[icon];
                const granted = grantedServices.has(key);
                return (
                  <div
                    key={key}
                    className={`flex items-start gap-3 p-3 rounded-lg border transition-colors ${
                      granted
                        ? 'border-emerald-200 dark:border-emerald-800 bg-emerald-50 dark:bg-emerald-950/30'
                        : 'border-border bg-muted/20 hover:bg-muted/40'
                    }`}
                    data-testid={`service-card-${tenantType}-${key}`}
                  >
                    <div className={`mt-0.5 ${granted ? 'text-emerald-600 dark:text-emerald-400' : 'text-muted-foreground'}`}>
                      <Icon className="w-4 h-4" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium leading-tight">{svcLabel}</p>
                      <p className="text-xs text-muted-foreground mt-0.5">{description}</p>
                    </div>
                    <Button
                      size="sm"
                      variant={granted ? 'outline' : 'default'}
                      className="shrink-0 text-xs h-7 px-2"
                      onClick={() => handleGrantService(key)}
                      data-testid={`button-grant-${tenantType}-${key}`}
                    >
                      {granted ? <><CheckCircle2 className="w-3 h-3 mr-1" />Granted</> : <><KeyRound className="w-3 h-3 mr-1" />Grant</>}
                    </Button>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Divider + manual section */}
        <details className="group">
          <summary className="cursor-pointer text-xs text-muted-foreground hover:text-foreground select-none flex items-center gap-1 list-none">
            <span className="border rounded px-1.5 py-0.5 group-open:hidden">▸</span>
            <span className="border rounded px-1.5 py-0.5 hidden group-open:inline">▾</span>
            Manual credential entry
          </summary>
          <div className="mt-4 space-y-4 pl-1">
            <div className="space-y-2">
              <Label className="text-sm font-medium">Application (Client) ID</Label>
              <Input
                value={localClientId}
                onChange={(e) => setLocalClientId(e.target.value)}
                placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                className="font-mono"
                data-testid={`input-${tenantType}-client-id`}
              />
            </div>
            <div className="space-y-2">
              <Label className="text-sm font-medium">Client Secret</Label>
              <div className="relative">
                <Input
                  type={showSecret ? 'text' : 'password'}
                  value={localClientSecret}
                  onChange={(e) => setLocalClientSecret(e.target.value)}
                  placeholder={hasExistingSecret ? "Secret saved — enter new value to replace" : "Enter client secret value"}
                  className="font-mono pr-10"
                  data-testid={`input-${tenantType}-client-secret`}
                />
                <button
                  type="button"
                  onClick={() => setShowSecret(!showSecret)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground transition-colors"
                  data-testid={`button-toggle-${tenantType}-secret`}
                >
                  {showSecret ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
            </div>
            <div className="flex flex-wrap gap-2">
              <Button size="sm" onClick={handleSave} disabled={isSaving} data-testid={`button-save-${tenantType}-credentials`}>
                {isSaving ? <Loader2 className="w-3 h-3 animate-spin mr-1" /> : null}
                Save
              </Button>
              <Button size="sm" variant="outline" onClick={handleTestConnection} disabled={isTesting || !hasCredentials} data-testid={`button-test-${tenantType}-connection`}>
                {isTesting ? <Loader2 className="w-3 h-3 animate-spin mr-1" /> : null}
                Test Connection
              </Button>
            </div>
          </div>
        </details>

        {testResult && (
          <div
            className={`flex items-start gap-2 p-3 rounded-lg text-sm ${
              testResult.success
                ? 'bg-emerald-50 dark:bg-emerald-950/30 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800'
                : 'bg-red-50 dark:bg-red-950/30 text-red-800 dark:text-red-300 border border-red-200 dark:border-red-800'
            }`}
            data-testid={`status-${tenantType}-connection-result`}
          >
            {testResult.success ? <CheckCircle2 className="w-4 h-4 mt-0.5 flex-shrink-0" /> : <XCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />}
            <span>{testResult.message}</span>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

// ======================== DISCOVERY TAB ========================

type DiscoveryType = 'users' | 'onedrive' | 'sharepoint' | 'teams' | 'powerplatform' | 'distributiongroups' | 'sharedmailboxes' | 'm365groups';

interface DiscoveryTabProps {
  projectId: number;
  existingItems: any[];
  onImport: (items: any[]) => void;
}

function DiscoveryTab({ projectId, existingItems, onImport }: DiscoveryTabProps) {
  const [activeType, setActiveType] = useState<DiscoveryType>('users');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [results, setResults] = useState<any[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [targetSuffix, setTargetSuffix] = useState('');
  const [lastRunAt, setLastRunAt] = useState<Date | null>(null);
  const { toast } = useToast();

  const discoveryTypes: { id: DiscoveryType; label: string; icon: any; description: string }[] = [
    { id: 'users',              label: 'Users',                 icon: UserCheck,  description: 'Discover all licensed users with mailbox or OneDrive' },
    { id: 'onedrive',           label: 'OneDrive',              icon: HardDrive,  description: 'Discover all provisioned OneDrive accounts with storage usage' },
    { id: 'sharepoint',         label: 'SharePoint Sites',      icon: Globe,      description: 'Discover all SharePoint sites with storage details' },
    { id: 'teams',              label: 'Microsoft Teams',       icon: Users,      description: 'Discover all Teams with member and channel counts' },
    { id: 'distributiongroups', label: 'Distribution Groups',   icon: AtSign,     description: 'Discover mail-enabled distribution lists and mail-enabled security groups' },
    { id: 'sharedmailboxes',    label: 'Shared Mailboxes',      icon: Inbox,      description: 'Discover all shared mailboxes in the source tenant' },
    { id: 'm365groups',         label: 'M365 Groups',           icon: Building2,  description: 'Discover all Microsoft 365 Groups with members and owners' },
    { id: 'powerplatform',      label: 'Power Platform',        icon: Zap,        description: 'Discover Power Apps and Power Automate flows' },
  ];

  const handleDiscover = async () => {
    setLoading(true);
    setError(null);
    setResults([]);
    setSelected(new Set());
    try {
      const res = await apiRequest('GET', `/api/projects/${projectId}/discover/${activeType}`);
      const data = await res.json();
      setResults(data.data || []);
      setLastRunAt(new Date());
    } catch (err: any) {
      setError(err.message || 'Discovery failed. Check source tenant credentials.');
    } finally {
      setLoading(false);
    }
  };

  const toggleAll = () => {
    if (selected.size === results.length) {
      setSelected(new Set());
    } else {
      setSelected(new Set(results.map(r => r.id)));
    }
  };

  const toggle = (id: string) => {
    setSelected(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const handleImportSelected = async () => {
    const selectedItems = results.filter(r => selected.has(r.id));
    if (selectedItems.length === 0) return;

    setLoading(true);
    try {
      // Build source identities for mapping — must match what we store as sourceIdentity later
      const sourceIdentities = selectedItems.map(r => {
        if (activeType === 'users' || activeType === 'onedrive' || activeType === 'sharedmailboxes') return r.userPrincipalName;
        if (activeType === 'sharepoint') return r.webUrl;
        if (activeType === 'distributiongroups' || activeType === 'm365groups') return r.mail || r.displayName;
        return r.id;
      });

      // Apply configured mapping rules from the server
      let mappedTargets: Record<string, string> = {};
      try {
        const res = await apiRequest('POST', `/api/projects/${projectId}/apply-mapping`, { identities: sourceIdentities });
        const mappingResults: { source: string; target: string }[] = await res.json();
        mappingResults.forEach(({ source, target }) => { mappedTargets[source] = target; });
      } catch {
        // If mapping API fails, continue without mapped targets
      }

      const items = selectedItems.map(r => {
        const itemType = activeType === 'users' ? 'user'
          : activeType === 'onedrive' ? 'onedrive'
          : activeType === 'sharepoint' ? 'sharepoint'
          : activeType === 'teams' ? 'teams'
          : activeType === 'distributiongroups' ? 'distributiongroup'
          : activeType === 'sharedmailboxes' ? 'sharedmailbox'
          : activeType === 'm365groups' ? 'm365group'
          : 'powerplatform';
        let sourceIdentity = '';
        let targetIdentity = '';

        if (activeType === 'users') {
          sourceIdentity = r.userPrincipalName;
          const mapped = mappedTargets[sourceIdentity];
          if (mapped && mapped !== sourceIdentity) {
            targetIdentity = mapped;
          } else if (targetSuffix) {
            targetIdentity = sourceIdentity.replace(/@.*/, `@${targetSuffix}`);
          }
        } else if (activeType === 'onedrive') {
          sourceIdentity = r.userPrincipalName;
          const mapped = mappedTargets[sourceIdentity];
          if (mapped && mapped !== sourceIdentity) {
            targetIdentity = mapped;
          } else if (targetSuffix) {
            targetIdentity = sourceIdentity.replace(/@.*/, `@${targetSuffix}`);
          }
        } else if (activeType === 'sharepoint') {
          sourceIdentity = r.webUrl;
          const mapped = mappedTargets[sourceIdentity];
          if (mapped && mapped !== sourceIdentity) {
            targetIdentity = mapped;
          } else {
            targetIdentity = r.displayName || r.webUrl.split('/sites/').pop()?.split('/')[0] || '';
          }
        } else if (activeType === 'teams') {
          sourceIdentity = r.id;
          targetIdentity = r.displayName;
        } else if (activeType === 'distributiongroups' || activeType === 'm365groups') {
          // Use mail address as identity; target defaults to same displayName (engine will match/create by name)
          sourceIdentity = r.mail || r.displayName;
          targetIdentity = r.displayName;
        } else if (activeType === 'sharedmailboxes') {
          sourceIdentity = r.userPrincipalName;
          const mapped = mappedTargets[sourceIdentity];
          if (mapped && mapped !== sourceIdentity) {
            targetIdentity = mapped;
          } else if (targetSuffix) {
            targetIdentity = sourceIdentity.replace(/@.*/, `@${targetSuffix}`);
          }
        } else {
          sourceIdentity = r.id;
          targetIdentity = '';
        }

        return { projectId, sourceIdentity, targetIdentity: targetIdentity || undefined, itemType, status: 'pending' };
      });

      // For users — also auto-add mailbox and OneDrive items for licensed users
      const extraItems: any[] = [];
      if (activeType === 'users') {
        // Build a set of existing (itemType, sourceIdentity) pairs to avoid duplicates
        const existingSet = new Set(
          existingItems.map((i: any) => `${i.itemType}::${i.sourceIdentity}`)
        );

        for (const r of selectedItems) {
          const sourceIdentity = r.userPrincipalName as string;
          const targetIdentity = items.find((i: any) => i.sourceIdentity === sourceIdentity)?.targetIdentity;

          if (r.hasMailbox && !existingSet.has(`mailbox::${sourceIdentity}`)) {
            extraItems.push({ projectId, sourceIdentity, targetIdentity: targetIdentity || undefined, itemType: 'mailbox', status: 'pending' });
          }
          if (r.hasOneDrive && !existingSet.has(`onedrive::${sourceIdentity}`)) {
            extraItems.push({ projectId, sourceIdentity, targetIdentity: targetIdentity || undefined, itemType: 'onedrive', status: 'pending' });
          }
        }
      }

      const allItems = [...items, ...extraItems];
      onImport(allItems);
      setSelected(new Set());
    } finally {
      setLoading(false);
    }
  };

  const activeTypeConfig = discoveryTypes.find(t => t.id === activeType)!;

  return (
    <div className="space-y-6">
      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><Search className="w-5 h-5" /> Source Tenant Discovery</CardTitle>
          <CardDescription>Scan the source tenant to find users, sites, and teams, then bulk-import them into the migration queue.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {discoveryTypes.map(t => {
              const Icon = t.icon;
              return (
                <button
                  key={t.id}
                  data-testid={`button-discover-${t.id}`}
                  onClick={() => { setActiveType(t.id); setResults([]); setError(null); setSelected(new Set()); setLastRunAt(null); }}
                  className={`flex flex-col items-center gap-2 p-4 rounded-lg border-2 transition-all text-sm font-medium ${
                    activeType === t.id
                      ? 'border-primary bg-primary/5 text-primary'
                      : 'border-border hover:border-primary/50 text-muted-foreground hover:text-foreground'
                  }`}
                >
                  <Icon className="w-5 h-5" />
                  {t.label}
                </button>
              );
            })}
          </div>

          <div className="text-sm text-muted-foreground">{activeTypeConfig.description}</div>

          <div className="flex flex-wrap gap-3 items-end">
            {(activeType === 'users' || activeType === 'onedrive' || activeType === 'sharedmailboxes') && (
              <div className="space-y-1 flex-1 min-w-[200px]">
                <Label>Target domain <span className="text-muted-foreground font-normal">(verified domain in target tenant — required)</span></Label>
                <Input
                  placeholder="targetcompany.com"
                  value={targetSuffix}
                  onChange={e => setTargetSuffix(e.target.value)}
                  data-testid="input-target-domain"
                />
              </div>
            )}
            <div className="flex flex-col items-start gap-1">
              <Button onClick={handleDiscover} disabled={loading} data-testid="button-run-discovery"
                variant={lastRunAt ? 'outline' : 'default'}>
                {loading
                  ? <><Loader2 className="w-4 h-4 animate-spin mr-2" />Discovering...</>
                  : lastRunAt
                    ? <><RefreshCw className="w-4 h-4 mr-2" />Rerun Discovery</>
                    : <><Search className="w-4 h-4 mr-2" />Discover {activeTypeConfig.label}</>
                }
              </Button>
              {lastRunAt && !loading && (
                <span className="text-xs text-muted-foreground">
                  Last run: {lastRunAt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                </span>
              )}
            </div>
          </div>

          {error && (
            <div className="flex items-start gap-2 p-3 rounded-lg bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-800 text-red-800 dark:text-red-300 text-sm" data-testid="status-discovery-error">
              <AlertTriangle className="w-4 h-4 mt-0.5 flex-shrink-0" />
              <span>{error}</span>
            </div>
          )}
        </CardContent>
      </Card>

      {results.length > 0 && (
        <Card className="shadow-sm">
          <CardHeader>
            <div className="flex items-center justify-between">
              <CardTitle>{results.length} {activeTypeConfig.label} Found</CardTitle>
              <div className="flex gap-2">
                <Button variant="outline" size="sm" onClick={toggleAll} data-testid="button-select-all">
                  {selected.size === results.length ? 'Deselect All' : 'Select All'}
                </Button>
                <Button size="sm" disabled={selected.size === 0} onClick={handleImportSelected} data-testid="button-import-selected">
                  <Boxes className="w-4 h-4 mr-2" /> Import Selected ({selected.size})
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent>
            <div className="border rounded-lg divide-y divide-border overflow-hidden">
              {results.map((r) => (
                <DiscoveryResultRow
                  key={r.id}
                  item={r}
                  type={activeType}
                  selected={selected.has(r.id)}
                  onToggle={() => toggle(r.id)}
                />
              ))}
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function DiscoveryResultRow({ item, type, selected, onToggle }: { item: any; type: DiscoveryType; selected: boolean; onToggle: () => void }) {
  if (type === 'onedrive') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-onedrive-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <HardDrive className="w-4 h-4 flex-shrink-0 text-blue-500" />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.userPrincipalName}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          {item.storageUsedBytes != null && (
            <div>{formatBytes(item.storageUsedBytes)} used</div>
          )}
          {item.storageAllocatedBytes != null && (
            <div className="text-muted-foreground/60">of {formatBytes(item.storageAllocatedBytes)}</div>
          )}
          {item.storageUsedBytes == null && <div>Storage unknown</div>}
        </div>
      </div>
    );
  }

  if (type === 'users') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-user-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <UserCheck className={`w-4 h-4 flex-shrink-0 ${item.accountEnabled ? 'text-emerald-500' : 'text-slate-400'}`} />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.userPrincipalName}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          <div>{item.department || item.jobTitle || ''}</div>
          <div className="flex gap-1 justify-end mt-0.5">
            {item.hasMailbox && <span className="px-1.5 py-0.5 rounded bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300">Mailbox</span>}
            {item.hasOneDrive && <span className="px-1.5 py-0.5 rounded bg-sky-100 dark:bg-sky-900/30 text-sky-700 dark:text-sky-300">OneDrive</span>}
          </div>
        </div>
      </div>
    );
  }

  if (type === 'sharepoint') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-site-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <Globe className="w-4 h-4 flex-shrink-0 text-teal-500" />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.webUrl}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          <div className="capitalize">{item.siteType} site</div>
          {item.storageUsedBytes && <div>{formatBytes(item.storageUsedBytes)} used</div>}
        </div>
      </div>
    );
  }

  if (type === 'teams') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-team-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <Users className="w-4 h-4 flex-shrink-0 text-indigo-500" />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.description || 'No description'}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          <div>{item.memberCount} members</div>
          <div>{item.channelCount} channels</div>
        </div>
      </div>
    );
  }

  if (type === 'distributiongroups') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-dg-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <AtSign className="w-4 h-4 flex-shrink-0 text-orange-500" />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.mail || 'No email address'}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          <div>{item.memberCount} members</div>
          <div>{item.ownerCount} owners</div>
        </div>
      </div>
    );
  }

  if (type === 'sharedmailboxes') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-sm-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <Inbox className="w-4 h-4 flex-shrink-0 text-rose-500" />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.userPrincipalName}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          <div>{item.mail || 'No mail'}</div>
          {item.mailboxType && <div className="capitalize text-muted-foreground/70">{item.mailboxType} mailbox</div>}
        </div>
      </div>
    );
  }

  if (type === 'm365groups') {
    return (
      <div className={`flex items-center gap-3 px-4 py-3 hover:bg-muted/30 transition-colors cursor-pointer ${selected ? 'bg-primary/5' : ''}`} onClick={onToggle} data-testid={`row-m365g-${item.id}`}>
        <input type="checkbox" checked={selected} onChange={() => {}} className="rounded" />
        <Building2 className="w-4 h-4 flex-shrink-0 text-cyan-500" />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm truncate">{item.displayName}</div>
          <div className="text-xs text-muted-foreground truncate">{item.mail || 'No email'}</div>
        </div>
        <div className="text-xs text-muted-foreground text-right">
          <div>{item.memberCount} members · {item.ownerCount} owners</div>
          <div className="capitalize text-muted-foreground/70">{item.visibility || 'Private'}</div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex items-center gap-3 px-4 py-4 text-sm text-amber-700 dark:text-amber-300 bg-amber-50 dark:bg-amber-950/30" data-testid={`row-powerplatform-${item.id}`}>
      <AlertTriangle className="w-4 h-4 flex-shrink-0" />
      <span>{item.note}</span>
    </div>
  );
}

// ======================== MAPPING RULES TAB ========================

function MappingRulesTab({ projectId }: { projectId: number }) {
  const { toast } = useToast();
  const queryClient = useQueryClient();
  const [ruleType, setRuleType] = useState('domain');
  const [sourcePattern, setSourcePattern] = useState('');
  const [targetPattern, setTargetPattern] = useState('');
  const [description, setDescription] = useState('');
  const [testInput, setTestInput] = useState('');
  const [testOutput, setTestOutput] = useState('');
  const [testLoading, setTestLoading] = useState(false);
  const [saving, setSaving] = useState(false);

  const { data: rules, isLoading } = useQuery<MappingRule[]>({
    queryKey: ['/api/projects/mapping-rules', projectId],
    queryFn: async () => {
      const res = await apiRequest('GET', `/api/projects/${projectId}/mapping-rules`);
      return res.json();
    },
  });

  const handleAddRule = async () => {
    if (!sourcePattern || !targetPattern) {
      toast({ title: 'Validation Error', description: 'Source and target patterns are required', variant: 'destructive' });
      return;
    }
    setSaving(true);
    try {
      await apiRequest('POST', `/api/projects/${projectId}/mapping-rules`, { ruleType, sourcePattern, targetPattern, description });
      queryClient.invalidateQueries({ queryKey: ['/api/projects/mapping-rules', projectId] });
      setSourcePattern(''); setTargetPattern(''); setDescription('');
      toast({ title: 'Rule Added', description: 'Mapping rule saved successfully' });
    } catch (err: any) {
      toast({ title: 'Error', description: err.message, variant: 'destructive' });
    } finally {
      setSaving(false);
    }
  };

  const handleDeleteRule = async (id: number) => {
    if (!confirm('Delete this mapping rule?')) return;
    try {
      await apiRequest('DELETE', `/api/projects/${projectId}/mapping-rules/${id}`);
      queryClient.invalidateQueries({ queryKey: ['/api/projects/mapping-rules', projectId] });
      toast({ title: 'Deleted', description: 'Rule removed' });
    } catch (err: any) {
      toast({ title: 'Error', description: err.message, variant: 'destructive' });
    }
  };

  const handleTest = async () => {
    if (!testInput) return;
    setTestLoading(true);
    try {
      const res = await apiRequest('POST', `/api/projects/${projectId}/apply-mapping`, { identities: [testInput] });
      const data = await res.json();
      setTestOutput(data[0]?.target || testInput);
    } catch {
      setTestOutput('Error applying rules');
    } finally {
      setTestLoading(false);
    }
  };

  const ruleTypeDescriptions: Record<string, string> = {
    domain: 'Replace the email domain: @sourcedomain.com → @targetdomain.com',
    prefix: 'Replace username prefix: old.user@domain → new.user@domain',
    suffix: 'Replace username suffix: user.old@domain → user.new@domain',
    upn_prefix: 'Replace entire username part: john.smith@domain → jsmith@domain',
  };

  return (
    <div className="space-y-6">
      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><MapPin className="w-5 h-5" /> Auto-Mapping Rules</CardTitle>
          <CardDescription>Configure rules to automatically transform source identities (UPNs, email addresses) to their target equivalents.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <Label>Rule Type</Label>
              <Select value={ruleType} onValueChange={setRuleType}>
                <SelectTrigger data-testid="select-rule-type">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="domain">Domain Replacement</SelectItem>
                  <SelectItem value="prefix">Username Prefix</SelectItem>
                  <SelectItem value="suffix">Username Suffix</SelectItem>
                  <SelectItem value="upn_prefix">Full UPN Prefix</SelectItem>
                </SelectContent>
              </Select>
              <p className="text-xs text-muted-foreground">{ruleTypeDescriptions[ruleType]}</p>
            </div>

            <div className="space-y-2">
              <Label>Description (optional)</Label>
              <Input placeholder="e.g., Replace old domain" value={description} onChange={e => setDescription(e.target.value)} data-testid="input-rule-description" />
            </div>

            <div className="space-y-2">
              <Label>Source Pattern</Label>
              <Input
                placeholder={ruleType === 'domain' ? 'acme.com' : 'old.prefix'}
                value={sourcePattern}
                onChange={e => setSourcePattern(e.target.value)}
                data-testid="input-source-pattern"
              />
            </div>

            <div className="space-y-2">
              <Label>Target Pattern</Label>
              <Input
                placeholder={ruleType === 'domain' ? 'contoso.com' : 'new.prefix'}
                value={targetPattern}
                onChange={e => setTargetPattern(e.target.value)}
                data-testid="input-target-pattern"
              />
            </div>
          </div>

          <Button onClick={handleAddRule} disabled={saving} data-testid="button-add-rule">
            {saving ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : <Plus className="w-4 h-4 mr-2" />}
            Add Rule
          </Button>
        </CardContent>
      </Card>

      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle>Existing Rules</CardTitle>
        </CardHeader>
        <CardContent>
          {isLoading ? (
            <div className="flex justify-center py-6"><Loader2 className="w-6 h-6 animate-spin text-muted-foreground" /></div>
          ) : !rules || rules.length === 0 ? (
            <div className="text-center py-8 text-muted-foreground">No mapping rules configured yet.</div>
          ) : (
            <div className="border rounded-lg divide-y divide-border overflow-hidden">
              {rules.map(rule => (
                <div key={rule.id} className="flex items-center gap-3 px-4 py-3" data-testid={`row-rule-${rule.id}`}>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 text-sm font-medium">
                      <span className="px-2 py-0.5 rounded bg-muted text-xs capitalize">{rule.ruleType}</span>
                      <span className="font-mono text-red-600 dark:text-red-400">{rule.sourcePattern}</span>
                      <span className="text-muted-foreground">→</span>
                      <span className="font-mono text-emerald-600 dark:text-emerald-400">{rule.targetPattern}</span>
                    </div>
                    {rule.description && <div className="text-xs text-muted-foreground mt-0.5">{rule.description}</div>}
                  </div>
                  <Button variant="ghost" size="sm" onClick={() => handleDeleteRule(rule.id)} data-testid={`button-delete-rule-${rule.id}`}>
                    <Trash2 className="w-4 h-4 text-red-500" />
                  </Button>
                </div>
              ))}
            </div>
          )}
        </CardContent>
      </Card>

      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle>Test Rules</CardTitle>
          <CardDescription>Enter a source UPN or email to preview how the rules transform it.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="flex gap-3 items-end">
            <div className="flex-1 space-y-1">
              <Label>Test Input</Label>
              <Input
                placeholder="john.smith@acme.com"
                value={testInput}
                onChange={e => setTestInput(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleTest()}
                data-testid="input-test-identity"
              />
            </div>
            <Button variant="outline" onClick={handleTest} disabled={testLoading || !testInput} data-testid="button-test-mapping">
              {testLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : 'Preview'}
            </Button>
          </div>
          {testOutput && (
            <div className="flex items-center gap-3 p-3 rounded-lg bg-muted/50">
              <span className="font-mono text-sm text-muted-foreground">{testInput}</span>
              <span className="text-muted-foreground">→</span>
              <span className="font-mono text-sm font-semibold text-emerald-600 dark:text-emerald-400" data-testid="text-mapping-result">{testOutput}</span>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

// ======================== ENTRA → AD TAB ========================

interface EntraCloudUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  givenName: string | null;
  surname: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
  mobilePhone: string | null;
  accountEnabled: boolean;
}

interface AdMigrationResult {
  userPrincipalName: string;
  success: boolean;
  created: boolean;
  message: string;
  tempPassword?: string;
}

function EntraToAdTab({ projectId, project }: { projectId: number; project: any }) {
  const { toast } = useToast();
  const queryClient = useQueryClient();

  // AD connection form state — initialised from project
  const [dcHostname, setDcHostname] = useState(project.adDcHostname || '');
  const [ldapPort, setLdapPort] = useState(String(project.adLdapPort || '389'));
  const [bindDn, setBindDn] = useState(project.adBindDn || '');
  const [bindPassword, setBindPassword] = useState('');
  const [baseDn, setBaseDn] = useState(project.adBaseDn || '');
  const [useSsl, setUseSsl] = useState(project.adUseSsl || false);
  const [targetOu, setTargetOu] = useState(project.adTargetOu || '');
  const [showAdPassword, setShowAdPassword] = useState(false);
  const [savingSettings, setSavingSettings] = useState(false);
  const [testingConn, setTestingConn] = useState(false);
  const [connResult, setConnResult] = useState<{ success: boolean; message: string } | null>(null);

  // Discovery state
  const [discovering, setDiscovering] = useState(false);
  const [discoveredUsers, setDiscoveredUsers] = useState<EntraCloudUser[]>([]);
  const [discoveryError, setDiscoveryError] = useState<string | null>(null);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [targetUpns, setTargetUpns] = useState<Record<string, string>>({});
  const [filterText, setFilterText] = useState('');

  // Migration state
  const [migrating, setMigrating] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [migrationResults, setMigrationResults] = useState<AdMigrationResult[]>([]);

  const hasAdSettings = !!(project.adDcHostname && project.adBindDn && project.adBaseDn);

  const handleSaveSettings = async () => {
    setSavingSettings(true);
    setConnResult(null);
    try {
      await apiRequest('POST', `/api/projects/${projectId}/ad-settings`, {
        adDcHostname: dcHostname || null,
        adLdapPort: parseInt(ldapPort) || 389,
        adBindDn: bindDn || null,
        adBindPassword: bindPassword || undefined,
        adBaseDn: baseDn || null,
        adUseSsl: useSsl,
        adTargetOu: targetOu || null,
      });
      queryClient.invalidateQueries({ queryKey: ['/api/projects', projectId] });
      toast({ title: 'Saved', description: 'Active Directory settings saved.' });
      setBindPassword('');
    } catch (err: any) {
      toast({ title: 'Error', description: err.message, variant: 'destructive' });
    } finally {
      setSavingSettings(false);
    }
  };

  const handleTestConnection = async () => {
    setTestingConn(true);
    setConnResult(null);
    try {
      const res = await apiRequest('POST', `/api/projects/${projectId}/ad-test-connection`, {
        adDcHostname: dcHostname,
        adLdapPort: parseInt(ldapPort) || 389,
        adBindDn: bindDn,
        adBindPassword: bindPassword || undefined,
        adBaseDn: baseDn,
        adUseSsl: useSsl,
      });
      const data = await res.json();
      setConnResult(data);
    } catch (err: any) {
      setConnResult({ success: false, message: err.message || 'Connection failed' });
    } finally {
      setTestingConn(false);
    }
  };

  const handleDiscover = async () => {
    setDiscovering(true);
    setDiscoveryError(null);
    setDiscoveredUsers([]);
    setSelected(new Set());
    setMigrationResults([]);
    try {
      const res = await apiRequest('GET', `/api/projects/${projectId}/entra-ad/discover`);
      const data = await res.json();
      const users: EntraCloudUser[] = data.users || [];
      setDiscoveredUsers(users);
      // Pre-fill target UPNs
      const upns: Record<string, string> = {};
      users.forEach(u => { upns[u.userPrincipalName] = u.userPrincipalName; });
      setTargetUpns(upns);
    } catch (err: any) {
      setDiscoveryError(err.message || 'Discovery failed');
    } finally {
      setDiscovering(false);
    }
  };

  const toggleAll = () => {
    const filtered = filteredUsers.map(u => u.userPrincipalName);
    const allSelected = filtered.every(upn => selected.has(upn));
    const next = new Set(selected);
    if (allSelected) filtered.forEach(upn => next.delete(upn));
    else filtered.forEach(upn => next.add(upn));
    setSelected(next);
  };

  const getSelectedUsers = () =>
    discoveredUsers.filter(u => selected.has(u.userPrincipalName)).map(u => ({
      upn: u.userPrincipalName,
      targetUpn: targetUpns[u.userPrincipalName] || u.userPrincipalName,
      displayName: u.displayName,
      givenName: u.givenName || undefined,
      surname: u.surname || undefined,
      jobTitle: u.jobTitle || undefined,
      department: u.department || undefined,
      officeLocation: u.officeLocation || undefined,
      mobilePhone: u.mobilePhone || undefined,
      mail: u.mail || undefined,
    }));

  const handleMigrate = async () => {
    const users = getSelectedUsers();
    if (users.length === 0) return;
    setMigrating(true);
    setMigrationResults([]);
    try {
      const res = await apiRequest('POST', `/api/projects/${projectId}/entra-ad/migrate`, { users });
      const data = await res.json();
      setMigrationResults(data.results || []);
      queryClient.invalidateQueries({ queryKey: ['/api/projects', projectId] });
      const success = (data.results || []).filter((r: AdMigrationResult) => r.success).length;
      const failed = (data.results || []).filter((r: AdMigrationResult) => !r.success).length;
      toast({
        title: `Migration Complete`,
        description: `${success} created/skipped, ${failed} failed`,
        variant: failed > 0 ? 'destructive' : 'default',
      });
    } catch (err: any) {
      toast({ title: 'Migration Error', description: err.message, variant: 'destructive' });
    } finally {
      setMigrating(false);
    }
  };

  const handleExportPs = async () => {
    const users = getSelectedUsers();
    if (users.length === 0) return;
    setExporting(true);
    try {
      const res = await apiRequest('POST', `/api/projects/${projectId}/entra-ad/export-ps`, { users });
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `migrate-entra-to-ad-${Date.now()}.ps1`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast({ title: 'Downloaded', description: 'PowerShell script downloaded. Run it on a domain controller as an admin.' });
    } catch (err: any) {
      toast({ title: 'Export Error', description: err.message, variant: 'destructive' });
    } finally {
      setExporting(false);
    }
  };

  const filteredUsers = discoveredUsers.filter(u =>
    !filterText ||
    u.displayName.toLowerCase().includes(filterText.toLowerCase()) ||
    u.userPrincipalName.toLowerCase().includes(filterText.toLowerCase()) ||
    (u.department || '').toLowerCase().includes(filterText.toLowerCase())
  );

  return (
    <div className="space-y-6">
      {/* Info banner */}
      <Card className="border-violet-200 dark:border-violet-900 bg-violet-50/50 dark:bg-violet-950/20 shadow-sm">
        <CardContent className="pt-6">
          <div className="flex gap-3">
            <Server className="w-5 h-5 text-violet-600 dark:text-violet-400 mt-0.5 flex-shrink-0" />
            <div className="space-y-1 text-sm">
              <p className="font-medium text-violet-900 dark:text-violet-200">Entra ID (Cloud-Only) → On-Premises Active Directory</p>
              <p className="text-violet-700 dark:text-violet-300">
                This tool discovers users that exist <em>only</em> in Entra ID (Azure AD) with no on-premises sync, and provisions them in your on-premises Active Directory via LDAP — or generates a PowerShell script to run on your domain controller.
              </p>
              <ul className="list-disc pl-5 space-y-0.5 text-violet-600 dark:text-violet-400">
                <li><strong>Direct LDAP migration</strong> requires network access to the DC and LDAPS (port 636) for password setting.</li>
                <li><strong>PowerShell export</strong> runs on any machine with RSAT and requires Domain Admin rights.</li>
                <li>After migration, configure Azure AD Connect to sync the AD accounts back to Entra ID for a hybrid identity.</li>
              </ul>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* AD Connection Settings */}
      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><Server className="w-5 h-5" /> On-Premises AD Connection</CardTitle>
          <CardDescription>Configure the LDAP connection to your on-premises Domain Controller.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <Label>Domain Controller Hostname / IP</Label>
              <Input
                placeholder="dc01.corp.com or 192.168.1.10"
                value={dcHostname}
                onChange={e => setDcHostname(e.target.value)}
                data-testid="input-ad-dc-hostname"
              />
            </div>
            <div className="space-y-2">
              <Label>LDAP Port</Label>
              <Input
                placeholder="389 (LDAP) or 636 (LDAPS)"
                value={ldapPort}
                onChange={e => setLdapPort(e.target.value)}
                data-testid="input-ad-ldap-port"
              />
            </div>
            <div className="space-y-2">
              <Label>Bind DN (Admin Account)</Label>
              <Input
                placeholder="CN=Administrator,CN=Users,DC=corp,DC=com"
                value={bindDn}
                onChange={e => setBindDn(e.target.value)}
                className="font-mono text-sm"
                data-testid="input-ad-bind-dn"
              />
            </div>
            <div className="space-y-2">
              <Label>Bind Password</Label>
              <div className="relative">
                <Input
                  type={showAdPassword ? 'text' : 'password'}
                  placeholder={project.adBindDn ? '(saved — enter new to change)' : 'Admin password'}
                  value={bindPassword}
                  onChange={e => setBindPassword(e.target.value)}
                  className="pr-10"
                  data-testid="input-ad-bind-password"
                />
                <button
                  type="button"
                  onClick={() => setShowAdPassword(!showAdPassword)}
                  className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
                >
                  {showAdPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                </button>
              </div>
            </div>
            <div className="space-y-2">
              <Label>Base DN</Label>
              <Input
                placeholder="DC=corp,DC=com"
                value={baseDn}
                onChange={e => setBaseDn(e.target.value)}
                className="font-mono text-sm"
                data-testid="input-ad-base-dn"
              />
            </div>
            <div className="space-y-2">
              <Label>Target OU (where users will be created)</Label>
              <Input
                placeholder="OU=MigratedUsers,DC=corp,DC=com"
                value={targetOu}
                onChange={e => setTargetOu(e.target.value)}
                className="font-mono text-sm"
                data-testid="input-ad-target-ou"
              />
              <p className="text-xs text-muted-foreground">Leave empty to use Base DN</p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <input
              type="checkbox"
              id="ad-use-ssl"
              checked={useSsl}
              onChange={e => setUseSsl(e.target.checked)}
              data-testid="checkbox-ad-use-ssl"
              className="rounded"
            />
            <label htmlFor="ad-use-ssl" className="text-sm cursor-pointer">
              Use LDAPS (SSL, port 636) — required for password setting in production
            </label>
          </div>

          <div className="flex flex-wrap gap-3">
            <Button onClick={handleSaveSettings} disabled={savingSettings} data-testid="button-save-ad-settings">
              {savingSettings ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : null}
              Save Settings
            </Button>
            <Button
              variant="outline"
              onClick={handleTestConnection}
              disabled={testingConn || !dcHostname || !bindDn || !baseDn}
              data-testid="button-test-ad-connection"
            >
              {testingConn ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : null}
              Test Connection
            </Button>
          </div>

          {connResult && (
            <div className={`flex items-start gap-2 p-3 rounded-lg text-sm border ${
              connResult.success
                ? 'bg-emerald-50 dark:bg-emerald-950/30 border-emerald-200 dark:border-emerald-800 text-emerald-800 dark:text-emerald-300'
                : 'bg-red-50 dark:bg-red-950/30 border-red-200 dark:border-red-800 text-red-800 dark:text-red-300'
            }`} data-testid="status-ad-connection-result">
              {connResult.success ? <CheckCircle2 className="w-4 h-4 mt-0.5 flex-shrink-0" /> : <XCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />}
              <span>{connResult.message}</span>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Discovery */}
      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><Search className="w-5 h-5" /> Discover Cloud-Only Users</CardTitle>
          <CardDescription>Scan the source Entra ID tenant to find users that have no on-premises Active Directory counterpart.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-4">
          <Button onClick={handleDiscover} disabled={discovering} data-testid="button-discover-cloud-users">
            {discovering ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : <Search className="w-4 h-4 mr-2" />}
            {discovering ? 'Scanning Entra ID...' : 'Discover Cloud-Only Users'}
          </Button>
          {discoveryError && (
            <div className="flex items-start gap-2 p-3 rounded-lg bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-800 text-red-800 dark:text-red-300 text-sm" data-testid="status-entra-discovery-error">
              <AlertTriangle className="w-4 h-4 mt-0.5 flex-shrink-0" />
              <span>{discoveryError}</span>
            </div>
          )}
        </CardContent>
      </Card>

      {/* User Table */}
      {discoveredUsers.length > 0 && (
        <Card className="shadow-sm">
          <CardHeader>
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div>
                <CardTitle>{discoveredUsers.length} Cloud-Only Users Found</CardTitle>
                <CardDescription className="mt-1">{selected.size} selected</CardDescription>
              </div>
              <div className="flex flex-wrap gap-2">
                <Button variant="outline" size="sm" onClick={toggleAll} data-testid="button-select-all-entra">
                  {filteredUsers.every(u => selected.has(u.userPrincipalName)) ? 'Deselect All' : 'Select All'}
                </Button>
                <Button
                  size="sm"
                  disabled={selected.size === 0 || migrating || !hasAdSettings}
                  onClick={handleMigrate}
                  data-testid="button-migrate-to-ad"
                >
                  {migrating ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : <Server className="w-4 h-4 mr-2" />}
                  Migrate via LDAP ({selected.size})
                </Button>
                <Button
                  variant="outline"
                  size="sm"
                  disabled={selected.size === 0 || exporting}
                  onClick={handleExportPs}
                  data-testid="button-export-powershell"
                >
                  {exporting ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : <Terminal className="w-4 h-4 mr-2" />}
                  Export PowerShell (.ps1)
                </Button>
              </div>
            </div>
            {!hasAdSettings && (
              <p className="text-xs text-amber-600 dark:text-amber-400 mt-2 flex items-center gap-1">
                <AlertTriangle className="w-3 h-3" /> Save AD connection settings above to enable direct LDAP migration.
              </p>
            )}
          </CardHeader>
          <CardContent className="space-y-3">
            <Input
              placeholder="Filter by name, UPN, or department..."
              value={filterText}
              onChange={e => setFilterText(e.target.value)}
              data-testid="input-filter-users"
            />
            <div className="border rounded-lg overflow-hidden">
              <div className="grid grid-cols-[auto_2fr_2fr_1fr_1fr] gap-0 bg-muted/40 px-4 py-2 text-xs font-medium text-muted-foreground border-b">
                <div className="w-6"></div>
                <div>User</div>
                <div>Target UPN in AD</div>
                <div>Department</div>
                <div>Status</div>
              </div>
              <div className="divide-y divide-border max-h-[500px] overflow-y-auto">
                {filteredUsers.map(user => {
                  const result = migrationResults.find(r => r.userPrincipalName === (targetUpns[user.userPrincipalName] || user.userPrincipalName));
                  return (
                    <div
                      key={user.userPrincipalName}
                      className={`grid grid-cols-[auto_2fr_2fr_1fr_1fr] gap-3 items-center px-4 py-3 hover:bg-muted/20 transition-colors ${selected.has(user.userPrincipalName) ? 'bg-primary/5' : ''}`}
                      data-testid={`row-entra-user-${user.id}`}
                    >
                      <input
                        type="checkbox"
                        checked={selected.has(user.userPrincipalName)}
                        onChange={() => {
                          const next = new Set(selected);
                          if (next.has(user.userPrincipalName)) next.delete(user.userPrincipalName);
                          else next.add(user.userPrincipalName);
                          setSelected(next);
                        }}
                        className="rounded"
                      />
                      <div className="min-w-0">
                        <div className="font-medium text-sm truncate">{user.displayName}</div>
                        <div className="text-xs text-muted-foreground truncate">{user.userPrincipalName}</div>
                        {user.jobTitle && <div className="text-xs text-muted-foreground truncate">{user.jobTitle}</div>}
                      </div>
                      <Input
                        className="text-xs font-mono h-7 px-2"
                        value={targetUpns[user.userPrincipalName] || user.userPrincipalName}
                        onChange={e => setTargetUpns(prev => ({ ...prev, [user.userPrincipalName]: e.target.value }))}
                        data-testid={`input-target-upn-${user.id}`}
                      />
                      <div className="text-xs text-muted-foreground truncate">{user.department || '—'}</div>
                      <div className="text-xs">
                        {!result ? (
                          <span className="text-muted-foreground">{user.accountEnabled ? 'Enabled' : 'Disabled'}</span>
                        ) : result.success ? (
                          <span className="text-emerald-600 dark:text-emerald-400 flex items-center gap-1">
                            <CheckCircle2 className="w-3 h-3" />
                            {result.created ? 'Created' : 'Exists'}
                          </span>
                        ) : (
                          <span className="text-red-600 dark:text-red-400 flex items-center gap-1" title={result.message}>
                            <XCircle className="w-3 h-3" /> Failed
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Migration Results */}
      {migrationResults.length > 0 && (
        <Card className="shadow-sm">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <CheckCircle2 className="w-5 h-5 text-emerald-500" /> Migration Results
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="border rounded-lg divide-y divide-border overflow-hidden">
              {migrationResults.map((r, i) => (
                <div key={i} className="flex items-start gap-3 px-4 py-3" data-testid={`result-ad-migration-${i}`}>
                  {r.success ? (
                    <CheckCircle2 className="w-4 h-4 text-emerald-500 mt-0.5 flex-shrink-0" />
                  ) : (
                    <XCircle className="w-4 h-4 text-red-500 mt-0.5 flex-shrink-0" />
                  )}
                  <div className="flex-1 min-w-0">
                    <div className="text-sm font-mono truncate">{r.userPrincipalName}</div>
                    <div className="text-xs text-muted-foreground">{r.message}</div>
                    {r.tempPassword && (
                      <div className="text-xs mt-1 font-mono bg-muted px-2 py-1 rounded inline-block">
                        Temp password: <span className="font-semibold">{r.tempPassword}</span>
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
            <p className="text-xs text-muted-foreground mt-3">
              Migration records have been saved to the Migration Items tab. Users must change their temporary passwords on first login.
              Remember to assign AD licences and configure Azure AD Connect for hybrid identity sync.
            </p>
          </CardContent>
        </Card>
      )}
    </div>
  );
}

function ExoPowerShellPanel({ projectId }: { projectId: number }) {
  const { toast } = useToast();
  const [sourceCertPath, setSourceCertPath] = useState('');
  const [sourceCertPassword, setSourceCertPassword] = useState('');
  const [sourceOrg, setSourceOrg] = useState('');
  const [targetCertPath, setTargetCertPath] = useState('');
  const [targetCertPassword, setTargetCertPassword] = useState('');
  const [targetOrg, setTargetOrg] = useState('');
  const [autoDelegate, setAutoDelegate] = useState(true);
  const [saving, setSaving] = useState(false);
  const [installing, setInstalling] = useState(false);
  const [testingSource, setTestingSource] = useState(false);
  const [testingTarget, setTestingTarget] = useState(false);
  const [testResult, setTestResult] = useState<{ tenant: string; success: boolean; message: string } | null>(null);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    fetch(`/api/projects/${projectId}/exo-settings`, { credentials: 'include' })
      .then(r => r.json()).then(d => {
        setSourceCertPath(d.sourceCertPath || '');
        setSourceCertPassword(d.sourceCertPassword || '');
        setSourceOrg(d.sourceOrg || '');
        setTargetCertPath(d.targetCertPath || '');
        setTargetCertPassword(d.targetCertPassword || '');
        setTargetOrg(d.targetOrg || '');
        setAutoDelegate(d.autoDelegate !== false);
        setLoaded(true);
      }).catch(() => setLoaded(true));
  }, [projectId]);

  const handleSave = async () => {
    setSaving(true);
    try {
      await apiRequest('PATCH', `/api/projects/${projectId}/exo-settings`, {
        sourceCertPath, sourceCertPassword, sourceOrg,
        targetCertPath, targetCertPassword, targetOrg, autoDelegate,
      });
      toast({ title: 'EXO settings saved', description: 'Exchange Online PowerShell configuration updated.' });
    } catch (e: any) {
      toast({ title: 'Error', description: e.message, variant: 'destructive' });
    } finally {
      setSaving(false);
    }
  };

  const handleInstall = async () => {
    setInstalling(true);
    toast({ title: 'Installing module…', description: 'This may take a minute.' });
    try {
      const r = await apiRequest('POST', `/api/projects/${projectId}/exo-install-module`, {});
      const d = await r.json();
      toast({ title: d.ok ? 'Module ready' : 'Install failed', description: d.message, variant: d.ok ? 'default' : 'destructive' });
    } catch (e: any) {
      toast({ title: 'Error', description: e.message, variant: 'destructive' });
    } finally {
      setInstalling(false);
    }
  };

  const handleTest = async (tenant: 'source' | 'target') => {
    if (tenant === 'source') setTestingSource(true); else setTestingTarget(true);
    setTestResult(null);
    try {
      const r = await apiRequest('POST', `/api/projects/${projectId}/exo-test`, { tenant });
      const d = await r.json();
      setTestResult({ tenant, success: d.success, message: d.success ? d.output?.join(' ') || 'Connected' : d.errors?.join('; ') || d.message || 'Failed' });
    } catch (e: any) {
      setTestResult({ tenant, success: false, message: e.message });
    } finally {
      if (tenant === 'source') setTestingSource(false); else setTestingTarget(false);
    }
  };

  if (!loaded) return <div className="flex justify-center p-8"><Loader2 className="animate-spin" /></div>;

  return (
    <Card className="shadow-sm">
      <CardHeader className="pb-3">
        <div className="flex items-center gap-2">
          <Terminal className="w-4 h-4 text-primary" />
          <CardTitle className="text-base">Exchange Online PowerShell</CardTitle>
          {targetCertPath && targetOrg && (
            <span className="inline-flex items-center gap-1 px-2 py-0.5 text-xs font-medium rounded-full bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-400">
              <CheckCircle2 className="w-3 h-3" /> Configured
            </span>
          )}
        </div>
        <CardDescription className="text-xs leading-relaxed mt-1">
          Enables automatic migration of delegate permissions (FullAccess, SendAs, SendOnBehalf) that Microsoft Graph API cannot access.
          Requires a certificate uploaded to each tenant's app registration and the <code className="bg-muted px-1 rounded">Exchange.ManageAsApp</code> API permission.
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-5">
        {/* Setup instructions */}
        <div className="rounded-lg border border-amber-200 dark:border-amber-800 bg-amber-50 dark:bg-amber-950/20 p-3 text-xs space-y-1.5">
          <p className="font-semibold text-amber-800 dark:text-amber-300">One-time setup per tenant app registration:</p>
          <ol className="list-decimal pl-4 space-y-1 text-amber-700 dark:text-amber-400">
            <li>In Entra ID → App registrations → your app → <strong>API permissions</strong>, add <code className="bg-amber-100 dark:bg-amber-900/50 px-1 rounded">Exchange.ManageAsApp</code> (under "Office 365 Exchange Online") and grant admin consent.</li>
            <li>Go to <strong>Certificates & secrets → Certificates → Upload certificate</strong>. Generate a self-signed cert (PowerShell: <code className="bg-amber-100 dark:bg-amber-900/50 px-1 rounded">New-SelfSignedCertificate</code>) and upload the <code>.cer</code> public key.</li>
            <li>Save the <code>.pfx</code> private key file on this machine and enter the path below.</li>
          </ol>
        </div>

        {/* Install module button */}
        <div className="flex items-center gap-3">
          <Button variant="outline" size="sm" onClick={handleInstall} disabled={installing} data-testid="button-exo-install-module">
            {installing ? <Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" /> : <Download className="w-3.5 h-3.5 mr-1.5" />}
            Install / verify ExchangeOnlineManagement module
          </Button>
          <span className="text-xs text-muted-foreground">Run once to ensure the PowerShell module is available on this machine.</span>
        </div>

        {/* Auto-delegate toggle */}
        <div className="flex items-center justify-between rounded-lg border border-border p-3">
          <div>
            <p className="text-sm font-medium">Automatically migrate delegates during shared mailbox migration</p>
            <p className="text-xs text-muted-foreground mt-0.5">When enabled, delegate permissions are read from source and applied to target automatically. Requires both source and target certificates configured below.</p>
          </div>
          <Switch checked={autoDelegate} onCheckedChange={setAutoDelegate} data-testid="switch-exo-auto-delegate" />
        </div>

        {/* Source EXO config */}
        <div className="space-y-3">
          <h4 className="text-sm font-semibold flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-blue-500" /> Source Tenant EXO</h4>
          <div className="grid grid-cols-1 gap-2">
            <div className="grid grid-cols-2 gap-2">
              <div className="space-y-1">
                <Label className="text-xs">Certificate PFX path</Label>
                <Input value={sourceCertPath} onChange={e => setSourceCertPath(e.target.value)} placeholder="C:\certs\source-migration.pfx" className="h-8 text-sm font-mono" data-testid="input-exo-source-cert-path" />
              </div>
              <div className="space-y-1">
                <Label className="text-xs">PFX password (if any)</Label>
                <Input type="password" value={sourceCertPassword} onChange={e => setSourceCertPassword(e.target.value)} placeholder="Leave blank if no password" className="h-8 text-sm" data-testid="input-exo-source-cert-password" />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-2 items-end">
              <div className="space-y-1">
                <Label className="text-xs">Organization domain</Label>
                <Input value={sourceOrg} onChange={e => setSourceOrg(e.target.value)} placeholder="sourcecorp.onmicrosoft.com" className="h-8 text-sm font-mono" data-testid="input-exo-source-org" />
              </div>
              <Button variant="outline" size="sm" className="h-8" onClick={() => handleTest('source')} disabled={testingSource || !sourceCertPath || !sourceOrg} data-testid="button-exo-test-source">
                {testingSource ? <Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" /> : <Zap className="w-3.5 h-3.5 mr-1.5" />}
                Test connection
              </Button>
            </div>
          </div>
        </div>

        {/* Target EXO config */}
        <div className="space-y-3">
          <h4 className="text-sm font-semibold flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-emerald-500" /> Target Tenant EXO</h4>
          <div className="grid grid-cols-1 gap-2">
            <div className="grid grid-cols-2 gap-2">
              <div className="space-y-1">
                <Label className="text-xs">Certificate PFX path</Label>
                <Input value={targetCertPath} onChange={e => setTargetCertPath(e.target.value)} placeholder="C:\certs\target-migration.pfx" className="h-8 text-sm font-mono" data-testid="input-exo-target-cert-path" />
              </div>
              <div className="space-y-1">
                <Label className="text-xs">PFX password (if any)</Label>
                <Input type="password" value={targetCertPassword} onChange={e => setTargetCertPassword(e.target.value)} placeholder="Leave blank if no password" className="h-8 text-sm" data-testid="input-exo-target-cert-password" />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-2 items-end">
              <div className="space-y-1">
                <Label className="text-xs">Organization domain</Label>
                <Input value={targetOrg} onChange={e => setTargetOrg(e.target.value)} placeholder="targetcorp.onmicrosoft.com" className="h-8 text-sm font-mono" data-testid="input-exo-target-org" />
              </div>
              <Button variant="outline" size="sm" className="h-8" onClick={() => handleTest('target')} disabled={testingTarget || !targetCertPath || !targetOrg} data-testid="button-exo-test-target">
                {testingTarget ? <Loader2 className="w-3.5 h-3.5 mr-1.5 animate-spin" /> : <Zap className="w-3.5 h-3.5 mr-1.5" />}
                Test connection
              </Button>
            </div>
          </div>
        </div>

        {/* Test result */}
        {testResult && (
          <div className={`flex items-start gap-2 p-3 rounded-lg text-sm ${testResult.success ? 'bg-emerald-50 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-800' : 'bg-red-50 dark:bg-red-950/20 border border-red-200 dark:border-red-800'}`}>
            {testResult.success
              ? <CheckCircle2 className="w-4 h-4 text-emerald-600 mt-0.5" />
              : <XCircle className="w-4 h-4 text-red-600 mt-0.5" />}
            <div>
              <span className="font-medium">{testResult.tenant === 'source' ? 'Source' : 'Target'} tenant:</span>{' '}
              <span className={testResult.success ? 'text-emerald-700 dark:text-emerald-400' : 'text-red-600'}>{testResult.message}</span>
            </div>
          </div>
        )}

        <div className="flex justify-end">
          <Button onClick={handleSave} disabled={saving} data-testid="button-exo-save">
            {saving ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : null}
            Save EXO settings
          </Button>
        </div>
      </CardContent>
    </Card>
  );
}

function TenantConfigTab({ projectId, project }: { projectId: number; project: any }) {
  return (
    <div className="space-y-6">
      <Card className="border-blue-200 dark:border-blue-900 bg-blue-50/50 dark:bg-blue-950/20 shadow-sm">
        <CardContent className="pt-6">
          <div className="flex gap-3">
            <Shield className="w-5 h-5 text-blue-600 dark:text-blue-400 mt-0.5 flex-shrink-0" />
            <div className="space-y-2 text-sm">
              <p className="font-medium text-blue-900 dark:text-blue-200">Microsoft Entra ID App Registration Setup</p>
              <p className="text-blue-800 dark:text-blue-300">
                To connect to each tenant via Microsoft Graph API, you need an App Registration in Microsoft Entra ID (Azure AD) for both your source and target tenants.
              </p>
              <ol className="list-decimal pl-5 space-y-1 text-blue-700 dark:text-blue-400">
                <li>Go to <a href="https://entra.microsoft.com" target="_blank" rel="noopener noreferrer" className="underline font-medium inline-flex items-center gap-1">entra.microsoft.com <ExternalLink className="w-3 h-3" /></a> and sign in as an admin.</li>
                <li>Navigate to <strong>Identity &rarr; Applications &rarr; App registrations &rarr; New registration</strong>.</li>
                <li>Name your app (e.g., "Migration Tool"), select "Accounts in this organizational directory only", and register.</li>
                <li>Copy the <strong>Application (Client) ID</strong> and <strong>Directory (Tenant) ID</strong>.</li>
                <li>Go to <strong>Certificates & secrets &rarr; New client secret</strong> and copy the secret <strong>Value</strong> (not the Secret ID).</li>
                <li>Go to <strong>API permissions &rarr; Add a permission &rarr; Microsoft Graph &rarr; Application permissions</strong> and add:
                  <code className="bg-blue-100 dark:bg-blue-900/50 px-1.5 py-0.5 rounded text-xs ml-1">User.Read.All</code>
                  <code className="bg-blue-100 dark:bg-blue-900/50 px-1.5 py-0.5 rounded text-xs ml-1">Mail.ReadWrite</code>
                  <code className="bg-blue-100 dark:bg-blue-900/50 px-1.5 py-0.5 rounded text-xs ml-1">Files.ReadWrite.All</code>
                  <code className="bg-blue-100 dark:bg-blue-900/50 px-1.5 py-0.5 rounded text-xs ml-1">Sites.ReadWrite.All</code>
                </li>
                <li>Click <strong>Grant admin consent</strong> for your organization.</li>
              </ol>
              <p className="text-blue-600 dark:text-blue-500 text-xs mt-2">Repeat these steps for both source and target tenants.</p>
            </div>
          </div>
        </CardContent>
      </Card>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <TenantCredentialForm
          label="Source Tenant"
          tenantType="source"
          projectId={projectId}
          tenantId={project.sourceTenantId}
          clientId={project.sourceClientId}
          clientSecret={project.sourceClientSecret}
          consentedServices={project.sourceConsentedServices}
        />
        <TenantCredentialForm
          label="Target Tenant"
          tenantType="target"
          projectId={projectId}
          tenantId={project.targetTenantId}
          clientId={project.targetClientId}
          clientSecret={project.targetClientSecret}
          consentedServices={project.targetConsentedServices}
        />
      </div>

      <ExoPowerShellPanel projectId={projectId} />
    </div>
  );
}
