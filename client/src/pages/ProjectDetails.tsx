import { useParams, Link } from "wouter";
import { useProject, useProjectStats, useUpdateProject } from "@/hooks/use-projects";
import { useMigrationItems, useCreateMigrationItem, useUpdateMigrationItem, useDeleteMigrationItem } from "@/hooks/use-items";
import { Sidebar } from "@/components/Sidebar";
import { StatusBadge } from "@/components/StatusBadge";
import { Loader2, ArrowLeft, Mail, Cloud, Users, Plus, Trash2, RotateCw, Eye, EyeOff, CheckCircle2, XCircle, Shield, ExternalLink } from "lucide-react";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip } from "recharts";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useForm, Controller } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { z } from "zod";
import { useToast } from "@/hooks/use-toast";
import { useState } from "react";
import { type MigrationItem } from "@shared/schema";
import { format } from "date-fns";
import { apiRequest } from "@/lib/queryClient";

const itemSchema = z.object({
  sourceIdentity: z.string().email("Invalid email"),
  targetIdentity: z.string().email("Invalid email").optional().or(z.literal("")),
  itemType: z.enum(["mailbox", "onedrive", "teams"]),
});

type ItemFormData = z.infer<typeof itemSchema>;

export default function ProjectDetails() {
  const params = useParams();
  const id = Number(params.id);
  const { data: project, isLoading: projectLoading } = useProject(id);
  const { data: stats, isLoading: statsLoading } = useProjectStats(id);
  const { data: items, isLoading: itemsLoading } = useMigrationItems(id);
  const { mutateAsync: updateProject } = useUpdateProject();
  const { mutateAsync: createItem } = useCreateMigrationItem();
  const { mutateAsync: updateItem } = useUpdateMigrationItem();
  const { mutateAsync: deleteItem } = useDeleteMigrationItem();
  
  const [isAddOpen, setIsAddOpen] = useState(false);
  const { toast } = useToast();

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
      await updateItem({ id: itemId, projectId: id, status: "pending" });
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

  if (projectLoading || statsLoading) {
    return (
      <div className="flex h-screen w-full items-center justify-center bg-background">
        <Loader2 className="w-10 h-10 animate-spin text-primary" />
      </div>
    );
  }

  if (!project) return <div>Project not found</div>;

  const chartData = stats ? [
    { name: 'Completed', value: stats.completed, color: '#10b981' }, // emerald-500
    { name: 'In Progress', value: stats.inProgress, color: '#3b82f6' }, // blue-500
    { name: 'Failed', value: stats.failed, color: '#ef4444' }, // red-500
    { name: 'Pending', value: stats.pending, color: '#94a3b8' }, // slate-400
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
                <h1 className="text-3xl font-bold tracking-tight">{project.name}</h1>
                <StatusBadge status={project.status} />
              </div>
              <div className="flex gap-2">
                 {project.status === 'draft' && (
                    <Button onClick={() => updateProject({ id, status: 'active' })}>
                        Start Migration
                    </Button>
                 )}
                 {project.status === 'active' && (
                    <Button variant="outline" onClick={() => updateProject({ id, status: 'completed' })}>
                        Mark Complete
                    </Button>
                 )}
              </div>
            </div>
            <p className="text-muted-foreground mt-2">{project.description}</p>
          </div>

          <Tabs defaultValue="overview" className="space-y-6">
            <TabsList className="bg-white dark:bg-slate-900 border border-border/50 p-1 rounded-lg">
              <TabsTrigger value="overview">Overview</TabsTrigger>
              <TabsTrigger value="items">Migration Items</TabsTrigger>
              <TabsTrigger value="tenant-config" data-testid="tab-tenant-config">Tenant Configuration</TabsTrigger>
            </TabsList>

            <TabsContent value="overview" className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Total Items</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold">{stats?.total || 0}</div>
                  </CardContent>
                </Card>
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Completed</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-emerald-600">{stats?.completed || 0}</div>
                  </CardContent>
                </Card>
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">In Progress</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-blue-600">{stats?.inProgress || 0}</div>
                  </CardContent>
                </Card>
                <Card className="shadow-sm">
                  <CardHeader className="pb-2">
                    <CardTitle className="text-sm font-medium text-muted-foreground">Failed</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-3xl font-bold text-red-600">{stats?.failed || 0}</div>
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
                            <span className="font-mono">{project.sourceTenantId}</span>
                        </div>
                        <div className="flex justify-between py-2 border-b border-border/50">
                            <span className="text-muted-foreground">Target Tenant ID</span>
                            <span className="font-mono">{project.targetTenantId}</span>
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
                <Dialog open={isAddOpen} onOpenChange={setIsAddOpen}>
                  <DialogTrigger asChild>
                    <Button>
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
                              <SelectTrigger>
                                <SelectValue placeholder="Select type" />
                              </SelectTrigger>
                              <SelectContent>
                                <SelectItem value="mailbox">Mailbox</SelectItem>
                                <SelectItem value="onedrive">OneDrive</SelectItem>
                                <SelectItem value="teams">Teams</SelectItem>
                              </SelectContent>
                            </Select>
                          )}
                        />
                      </div>
                      <div className="space-y-2">
                        <Label>Source Identity</Label>
                        <Input {...form.register("sourceIdentity")} placeholder="user@source.com" />
                        {form.formState.errors.sourceIdentity && <p className="text-xs text-red-500">{form.formState.errors.sourceIdentity.message}</p>}
                      </div>
                      <div className="space-y-2">
                        <Label>Target Identity (Optional)</Label>
                        <Input {...form.register("targetIdentity")} placeholder="user@target.com" />
                        {form.formState.errors.targetIdentity && <p className="text-xs text-red-500">{form.formState.errors.targetIdentity.message}</p>}
                      </div>
                      <div className="flex justify-end pt-4">
                        <Button type="submit">Add Item</Button>
                      </div>
                    </form>
                  </DialogContent>
                </Dialog>
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
                        <tr key={item.id} className="hover:bg-slate-50 dark:hover:bg-slate-900/50">
                          <td className="px-6 py-4">
                            {item.itemType === 'mailbox' && <Mail className="w-4 h-4 text-blue-500" />}
                            {item.itemType === 'onedrive' && <Cloud className="w-4 h-4 text-sky-500" />}
                            {item.itemType === 'teams' && <Users className="w-4 h-4 text-indigo-500" />}
                          </td>
                          <td className="px-6 py-4 font-medium">{item.sourceIdentity}</td>
                          <td className="px-6 py-4 text-muted-foreground">{item.targetIdentity || "Auto-mapped"}</td>
                          <td className="px-6 py-4">
                            <StatusBadge status={item.status} />
                            {item.status === 'failed' && item.errorDetails && (
                                <div className="text-xs text-red-500 mt-1 max-w-[200px] truncate" title={item.errorDetails}>
                                    {item.errorDetails}
                                </div>
                            )}
                          </td>
                          <td className="px-6 py-4 text-right space-x-2">
                             {item.status === 'failed' && (
                                <Button size="sm" variant="ghost" onClick={() => handleRetry(item.id)} title="Retry">
                                    <RotateCw className="w-4 h-4 text-blue-600" />
                                </Button>
                             )}
                             <Button size="sm" variant="ghost" onClick={() => handleDelete(item.id)} title="Delete">
                                <Trash2 className="w-4 h-4 text-red-500" />
                             </Button>
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

            <TabsContent value="tenant-config">
                <TenantConfigTab projectId={id} project={project} />
            </TabsContent>
          </Tabs>
        </div>
      </main>
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

        <div className="flex items-center gap-3 pt-2">
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
                <li>Navigate to <strong>Identity → Applications → App registrations → New registration</strong>.</li>
                <li>Name your app (e.g., "Migration Tool"), select "Accounts in this organizational directory only", and register.</li>
                <li>Copy the <strong>Application (Client) ID</strong> and <strong>Directory (Tenant) ID</strong>.</li>
                <li>Go to <strong>Certificates & secrets → New client secret</strong> and copy the secret <strong>Value</strong> (not the Secret ID).</li>
                <li>Go to <strong>API permissions → Add a permission → Microsoft Graph → Application permissions</strong> and add:
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
