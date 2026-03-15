import { useParams, Link } from "wouter";
import { useProject, useProjectStats, useUpdateProject } from "@/hooks/use-projects";
import { useMigrationItems, useCreateMigrationItem, useUpdateMigrationItem, useDeleteMigrationItem } from "@/hooks/use-items";
import { Sidebar } from "@/components/Sidebar";
import { StatusBadge } from "@/components/StatusBadge";
import { Loader2, ArrowLeft, Mail, Cloud, Users, Plus, Trash2, RotateCw, Eye, EyeOff, CheckCircle2, XCircle, Shield, ExternalLink, Play, PlayCircle, FileText, Globe, KeyRound, Search, UserCheck, MapPin, Zap, AlertTriangle, Import, Boxes } from "lucide-react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip } from "recharts";
import { Progress } from "@/components/ui/progress";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
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
  itemType: z.enum(["mailbox", "onedrive", "sharepoint", "teams", "user", "powerplatform"]),
});

type ItemFormData = z.infer<typeof itemSchema>;

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

  const [isAddOpen, setIsAddOpen] = useState(false);
  const [logsDialogItem, setLogsDialogItem] = useState<MigrationItem | null>(null);
  const [itemLogs, setItemLogs] = useState<string[]>([]);
  const [logsLoading, setLogsLoading] = useState(false);
  const { toast } = useToast();

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

  const pendingOrFailedCount = items?.filter(i => i.status === 'pending' || i.status === 'failed').length || 0;

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

          <Tabs defaultValue="overview" className="space-y-6">
            <TabsList className="bg-white dark:bg-slate-900 border border-border/50 p-1 rounded-lg flex-wrap">
              <TabsTrigger value="overview" data-testid="tab-overview">Overview</TabsTrigger>
              <TabsTrigger value="items" data-testid="tab-items">Migration Items</TabsTrigger>
              <TabsTrigger value="discovery" data-testid="tab-discovery">Discovery</TabsTrigger>
              <TabsTrigger value="mapping" data-testid="tab-mapping">Auto-Mapping Rules</TabsTrigger>
              <TabsTrigger value="tenant-config" data-testid="tab-tenant-config">Tenant Configuration</TabsTrigger>
            </TabsList>

            <TabsContent value="overview" className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Total Items</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold" data-testid="stat-total">{stats?.total || 0}</div>
                  </CardContent>
                </Card>
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Completed</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-emerald-600" data-testid="stat-completed">{stats?.completed || 0}</div>
                  </CardContent>
                </Card>
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">In Progress</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-blue-600" data-testid="stat-in-progress">{stats?.inProgress || 0}</div>
                  </CardContent>
                </Card>
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Failed</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-red-600" data-testid="stat-failed">{stats?.failed || 0}</div>
                  </CardContent>
                </Card>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <Card className="h-[300px]">
                  <CardHeader>
                    <CardTitle>Migration Progress</CardTitle>
                  </CardHeader>
                  <CardContent className="h-[220px]">
                    {chartData.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie
                            data={chartData}
                            cx="50%"
                            cy="50%"
                            innerRadius={60}
                            outerRadius={80}
                            paddingAngle={5}
                            dataKey="value"
                          >
                            {chartData.map((entry, index) => (
                              <Cell key={`cell-${index}`} fill={entry.color} />
                            ))}
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    ) : (
                        <div className="flex h-full items-center justify-center text-muted-foreground">
                            No data available
                        </div>
                    )}
                  </CardContent>
                </Card>

                <Card>
                    <CardHeader>
                        <CardTitle>Tenant Details</CardTitle>
                    </CardHeader>
                    <CardContent className="space-y-4">
                        <div className="flex justify-between py-2 border-b border-border/50">
                            <span className="text-muted-foreground">Source Tenant ID</span>
                            <span className="font-mono text-sm" data-testid="text-source-tenant">{project.sourceTenantId}</span>
                        </div>
                        <div className="flex justify-between py-2 border-b border-border/50">
                            <span className="text-muted-foreground">Target Tenant ID</span>
                            <span className="font-mono text-sm" data-testid="text-target-tenant">{project.targetTenantId}</span>
                        </div>
                        <div className="flex justify-between py-2">
                            <span className="text-muted-foreground">Created At</span>
                            <span>{project.createdAt ? format(new Date(project.createdAt), 'PPP') : '-'}</span>
                        </div>
                    </CardContent>
                </Card>
              </div>
            </TabsContent>

            <TabsContent value="items" className="space-y-6">
              <div className="flex justify-between items-center">
                <h2 className="text-xl font-bold">Mapped Resources</h2>
                <div className="flex gap-2">
                  {pendingOrFailedCount > 0 && (
                    <Button variant="default" onClick={handleMigrateAll} data-testid="button-migrate-all">
                      <PlayCircle className="w-4 h-4 mr-2" /> Run All ({pendingOrFailedCount})
                    </Button>
                  )}
                  <Dialog open={isAddOpen} onOpenChange={setIsAddOpen}>
                    <DialogTrigger asChild>
                      <Button variant="outline" data-testid="button-add-item">
                        <Plus className="w-4 h-4 mr-2" /> Add Item
                      </Button>
                    </DialogTrigger>
                    <DialogContent>
                      <DialogHeader>
                        <DialogTitle>Add Migration Item</DialogTitle>
                      </DialogHeader>
                      <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 py-4">
                        <div className="space-y-2">
                          <Label>Item Type</Label>
                          <Controller
                            control={form.control}
                            name="itemType"
                            render={({ field }) => (
                              <Select onValueChange={field.onChange} defaultValue={field.value}>
                                <SelectTrigger data-testid="select-item-type">
                                  <SelectValue placeholder="Select type" />
                                </SelectTrigger>
                                <SelectContent>
                                  <SelectItem value="mailbox">Mailbox (Email)</SelectItem>
                                  <SelectItem value="onedrive">OneDrive</SelectItem>
                                  <SelectItem value="sharepoint">SharePoint</SelectItem>
                                  <SelectItem value="teams">Teams</SelectItem>
                                  <SelectItem value="user">User Account</SelectItem>
                                  <SelectItem value="powerplatform">Power Platform</SelectItem>
                                </SelectContent>
                              </Select>
                            )}
                          />
                        </div>
                        <div className="space-y-2">
                          <Label>Source Identity</Label>
                          <Input {...form.register("sourceIdentity")} placeholder={form.watch("itemType") === "sharepoint" ? "contoso.sharepoint.com:/sites/TeamSite" : "user@source.com"} data-testid="input-source-identity" />
                          {form.formState.errors.sourceIdentity && <p className="text-xs text-red-500">{form.formState.errors.sourceIdentity.message}</p>}
                          {form.watch("itemType") === "sharepoint" && (
                            <p className="text-xs text-muted-foreground">For SharePoint, use the site hostname and path (e.g., contoso.sharepoint.com:/sites/TeamSite) or site display name for search.</p>
                          )}
                        </div>
                        <div className="space-y-2">
                          <Label>Target Identity</Label>
                          <Input {...form.register("targetIdentity")} placeholder={form.watch("itemType") === "sharepoint" ? "fabrikam.sharepoint.com:/sites/TeamSite" : "user@target.com"} data-testid="input-target-identity" />
                          {form.formState.errors.targetIdentity && <p className="text-xs text-red-500">{form.formState.errors.targetIdentity.message}</p>}
                        </div>
                        <div className="flex justify-end pt-4">
                          <Button type="submit" data-testid="button-submit-item">Add Item</Button>
                        </div>
                      </form>
                    </DialogContent>
                  </Dialog>
                </div>
              </div>

              <div className="bg-card rounded-xl border border-border/60 shadow-sm overflow-hidden">
                {itemsLoading ? (
                    <div className="p-8 flex justify-center"><Loader2 className="animate-spin" /></div>
                ) : items && items.length > 0 ? (
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-muted/30 border-b border-border/60">
                        <th className="px-6 py-4 text-left font-semibold text-muted-foreground w-12">Type</th>
                        <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Source Identity</th>
                        <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Target Identity</th>
                        <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Status</th>
                        <th className="px-6 py-4 text-right font-semibold text-muted-foreground">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-border/40">
                      {items.map((item: MigrationItem) => (
                        <tr key={item.id} className="hover:bg-slate-50 dark:hover:bg-slate-900/50" data-testid={`row-item-${item.id}`}>
                          <td className="px-6 py-4">
                            <ItemTypeIcon type={item.itemType} />
                          </td>
                          <td className="px-6 py-4 font-medium" data-testid={`text-source-${item.id}`}>{item.sourceIdentity}</td>
                          <td className="px-6 py-4 text-muted-foreground" data-testid={`text-target-${item.id}`}>{item.targetIdentity || "Auto-mapped"}</td>
                          <td className="px-6 py-4 min-w-[200px]">
                            <div className="flex items-center gap-2">
                              <StatusBadge status={item.status} />
                              {item.status === 'in_progress' && <Loader2 className="w-3 h-3 animate-spin text-blue-500" />}
                            </div>
                            {item.status === 'in_progress' && item.bytesTotal != null && item.bytesTotal > 0 && (
                              <div className="mt-2 space-y-1" data-testid={`progress-${item.id}`}>
                                <Progress value={item.progressPercent ?? 0} className="h-1.5" />
                                <div className="text-xs text-muted-foreground flex justify-between">
                                  <span>{formatBytes(item.bytesMigrated)} / {formatBytes(item.bytesTotal)}</span>
                                  <span className="font-medium text-blue-500">{item.progressPercent ?? 0}%</span>
                                </div>
                              </div>
                            )}
                            {(item.status === 'completed') && item.bytesMigrated != null && item.bytesMigrated > 0 && (
                              <div className="text-xs text-muted-foreground mt-1" data-testid={`bytes-${item.id}`}>
                                {formatBytes(item.bytesMigrated)} migrated
                              </div>
                            )}
                            {item.status === 'failed' && item.errorDetails && (
                              <div className="text-xs text-red-500 mt-1 max-w-[200px] truncate" title={item.errorDetails} data-testid={`text-error-${item.id}`}>
                                {item.errorDetails}
                              </div>
                            )}
                          </td>
                          <td className="px-6 py-4 text-right">
                            <div className="flex items-center justify-end gap-1">
                              {(item.status === 'pending' || item.status === 'failed') && (
                                <Button size="sm" variant="ghost" onClick={() => handleMigrateItem(item.id)} title="Start Migration" data-testid={`button-migrate-${item.id}`}>
                                  <Play className="w-4 h-4 text-emerald-600" />
                                </Button>
                              )}
                              {item.status === 'failed' && (
                                <Button size="sm" variant="ghost" onClick={() => handleRetry(item.id)} title="Reset to Pending" data-testid={`button-retry-${item.id}`}>
                                  <RotateCw className="w-4 h-4 text-blue-600" />
                                </Button>
                              )}
                              {(item.logs && (item.logs as string[]).length > 0) || item.status === 'in_progress' || item.status === 'completed' || item.status === 'failed' ? (
                                <Button size="sm" variant="ghost" onClick={() => handleViewLogs(item)} title="View Logs" data-testid={`button-logs-${item.id}`}>
                                  <FileText className="w-4 h-4 text-slate-500" />
                                </Button>
                              ) : null}
                              {item.status !== 'in_progress' && (
                                <Button size="sm" variant="ghost" onClick={() => handleDelete(item.id)} title="Delete" data-testid={`button-delete-${item.id}`}>
                                  <Trash2 className="w-4 h-4 text-red-500" />
                                </Button>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <div className="p-12 text-center text-muted-foreground">
                    No items mapped yet. Click "Add Item" to start.
                  </div>
                )}
              </div>
            </TabsContent>

            <TabsContent value="discovery">
              <DiscoveryTab projectId={id} onImport={(newItems) => {
                newItems.forEach(item => createItem(item).catch(() => {}));
                queryClient.invalidateQueries({ queryKey: [api.items.list.path, id] });
                toast({ title: "Imported", description: `${newItems.length} item(s) added to migration queue.` });
              }} />
            </TabsContent>

            <TabsContent value="mapping">
              <MappingRulesTab projectId={id} />
            </TabsContent>

            <TabsContent value="tenant-config">
                <TenantConfigTab projectId={id} project={project} />
            </TabsContent>
          </Tabs>
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
                <div key={i} className={`py-0.5 ${log.includes('failed') || log.includes('Failed') || log.includes('Error') ? 'text-red-400' : log.includes('complete') || log.includes('Complete') || log.includes('success') ? 'text-emerald-400' : ''}`}>
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

function TenantCredentialForm({
  label,
  tenantType,
  projectId,
  tenantId,
  clientId,
  clientSecret,
}: {
  label: string;
  tenantType: 'source' | 'target';
  projectId: number;
  tenantId: string;
  clientId: string | null;
  clientSecret: string | null;
}) {
  const { toast } = useToast();
  const { mutateAsync: updateProject, isPending: isSaving } = useUpdateProject();
  const [showSecret, setShowSecret] = useState(false);
  const [localClientId, setLocalClientId] = useState(clientId || '');
  const [localClientSecret, setLocalClientSecret] = useState('');
  const hasExistingSecret = !!clientSecret;
  const [testResult, setTestResult] = useState<{ success: boolean; message: string } | null>(null);
  const [isTesting, setIsTesting] = useState(false);

  const handleSave = async () => {
    const secretUpdate = localClientSecret ? (tenantType === 'source' ? { sourceClientSecret: localClientSecret } : { targetClientSecret: localClientSecret }) : {};
    const updates = tenantType === 'source'
      ? { sourceClientId: localClientId, ...secretUpdate }
      : { targetClientId: localClientId, ...secretUpdate };

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
        <div className="flex items-center gap-2">
          <Shield className="w-5 h-5 text-primary" />
          <CardTitle className="text-lg">{label}</CardTitle>
        </div>
        <CardDescription>
          Microsoft Entra ID App Registration credentials for the {tenantType} tenant.
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-5">
        <div className="space-y-2">
          <Label className="text-sm font-medium">Tenant ID</Label>
          <Input
            value={tenantId}
            disabled
            className="font-mono bg-muted/50"
            data-testid={`input-${tenantType}-tenant-id`}
          />
          <p className="text-xs text-muted-foreground">Set when creating the project. Edit in project settings to change.</p>
        </div>

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
          <div className="flex gap-2">
            <div className="relative flex-1">
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
        </div>

        <div className="flex flex-wrap items-center gap-3 pt-2">
          <Button onClick={handleSave} disabled={isSaving} data-testid={`button-save-${tenantType}-credentials`}>
            {isSaving ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : null}
            Save Credentials
          </Button>
          <Button
            variant="outline"
            onClick={handleTestConnection}
            disabled={isTesting || !hasCredentials}
            data-testid={`button-test-${tenantType}-connection`}
          >
            {isTesting ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : null}
            Test Connection
          </Button>
          <Button
            variant="outline"
            disabled={!tenantId || !localClientId}
            onClick={() => {
              const redirectUri = encodeURIComponent(window.location.origin);
              const url = `https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${localClientId}&redirect_uri=${redirectUri}`;
              window.open(url, '_blank');
            }}
            data-testid={`button-consent-${tenantType}`}
          >
            <KeyRound className="w-4 h-4 mr-2" />
            Grant App Permissions
          </Button>
        </div>

        {testResult && (
          <div
            className={`flex items-start gap-2 p-3 rounded-lg text-sm ${
              testResult.success
                ? 'bg-emerald-50 dark:bg-emerald-950/30 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800'
                : 'bg-red-50 dark:bg-red-950/30 text-red-800 dark:text-red-300 border border-red-200 dark:border-red-800'
            }`}
            data-testid={`status-${tenantType}-connection-result`}
          >
            {testResult.success ? (
              <CheckCircle2 className="w-4 h-4 mt-0.5 flex-shrink-0" />
            ) : (
              <XCircle className="w-4 h-4 mt-0.5 flex-shrink-0" />
            )}
            <span>{testResult.message}</span>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

// ======================== DISCOVERY TAB ========================

type DiscoveryType = 'users' | 'sharepoint' | 'teams' | 'powerplatform';

interface DiscoveryTabProps {
  projectId: number;
  onImport: (items: any[]) => void;
}

function DiscoveryTab({ projectId, onImport }: DiscoveryTabProps) {
  const [activeType, setActiveType] = useState<DiscoveryType>('users');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [results, setResults] = useState<any[]>([]);
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [targetSuffix, setTargetSuffix] = useState('');
  const { toast } = useToast();

  const discoveryTypes: { id: DiscoveryType; label: string; icon: any; description: string }[] = [
    { id: 'users', label: 'Users', icon: UserCheck, description: 'Discover all licensed users with mailbox or OneDrive' },
    { id: 'sharepoint', label: 'SharePoint Sites', icon: Globe, description: 'Discover all SharePoint sites with storage details' },
    { id: 'teams', label: 'Microsoft Teams', icon: Users, description: 'Discover all Teams with member and channel counts' },
    { id: 'powerplatform', label: 'Power Platform', icon: Zap, description: 'Discover Power Apps and Power Automate flows' },
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

  const handleImportSelected = () => {
    const selectedItems = results.filter(r => selected.has(r.id));
    const items = selectedItems.map(r => {
      let sourceIdentity = '';
      let targetIdentity = '';
      let itemType = activeType === 'users' ? 'user' : activeType === 'sharepoint' ? 'sharepoint' : activeType === 'teams' ? 'teams' : 'powerplatform';

      if (activeType === 'users') {
        sourceIdentity = r.userPrincipalName;
        targetIdentity = targetSuffix ? r.userPrincipalName.replace(/@.*/, `@${targetSuffix}`) : '';
      } else if (activeType === 'sharepoint') {
        sourceIdentity = r.webUrl;
        targetIdentity = targetSuffix || '';
      } else if (activeType === 'teams') {
        sourceIdentity = r.id;
        targetIdentity = r.displayName;
      } else {
        sourceIdentity = r.id;
        targetIdentity = '';
      }

      return {
        projectId,
        sourceIdentity,
        targetIdentity: targetIdentity || undefined,
        itemType,
        status: 'pending',
      };
    });
    onImport(items);
    setSelected(new Set());
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
                  onClick={() => { setActiveType(t.id); setResults([]); setError(null); setSelected(new Set()); }}
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
            {(activeType === 'users') && (
              <div className="space-y-1 flex-1 min-w-[200px]">
                <Label>Target domain suffix (optional)</Label>
                <Input
                  placeholder="contoso.com"
                  value={targetSuffix}
                  onChange={e => setTargetSuffix(e.target.value)}
                  data-testid="input-target-domain"
                />
              </div>
            )}
            <Button onClick={handleDiscover} disabled={loading} data-testid="button-run-discovery">
              {loading ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : <Search className="w-4 h-4 mr-2" />}
              {loading ? 'Discovering...' : `Discover ${activeTypeConfig.label}`}
            </Button>
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
        />
        <TenantCredentialForm
          label="Target Tenant"
          tenantType="target"
          projectId={projectId}
          tenantId={project.targetTenantId}
          clientId={project.targetClientId}
          clientSecret={project.targetClientSecret}
        />
      </div>
    </div>
  );
}
