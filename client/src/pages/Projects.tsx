import { useState } from "react";
import { useProjects, useCreateProject } from "@/hooks/use-projects";
import { Sidebar } from "@/components/Sidebar";
import { StatusBadge } from "@/components/StatusBadge";
import { Link } from "wouter";
import { Plus, Search, Loader2 } from "lucide-react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { useToast } from "@/hooks/use-toast";
import { useForm } from "react-hook-form";
import { z } from "zod";
import { zodResolver } from "@hookform/resolvers/zod";
import { motion } from "framer-motion";

const projectSchema = z.object({
  name: z.string().min(1, "Name is required"),
  sourceTenantId: z.string().min(1, "Source Tenant ID is required"),
  targetTenantId: z.string().min(1, "Target Tenant ID is required"),
  description: z.string().optional(),
});

type ProjectFormData = z.infer<typeof projectSchema>;

export default function Projects() {
  const { data: projects, isLoading } = useProjects();
  const { mutateAsync: createProject, isPending: isCreating } = useCreateProject();
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const { toast } = useToast();

  const form = useForm<ProjectFormData>({
    resolver: zodResolver(projectSchema),
    defaultValues: {
      name: "",
      sourceTenantId: "",
      targetTenantId: "",
      description: "",
    },
  });

  const onSubmit = async (data: ProjectFormData) => {
    try {
      await createProject({ ...data, status: "draft" });
      setIsOpen(false);
      form.reset();
      toast({
        title: "Success",
        description: "Project created successfully",
      });
    } catch (error) {
      toast({
        title: "Error",
        description: error instanceof Error ? error.message : "Failed to create project",
        variant: "destructive",
      });
    }
  };

  const filteredProjects = projects?.filter(p => 
    p.name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    p.sourceTenantId.toLowerCase().includes(searchTerm.toLowerCase())
  );

  if (isLoading) {
    return (
      <div className="flex h-screen w-full items-center justify-center bg-background">
        <Loader2 className="w-10 h-10 animate-spin text-primary" />
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 text-foreground font-sans">
      <Sidebar />
      
      <main className="flex-1 overflow-y-auto">
        <div className="container max-w-7xl mx-auto px-8 py-8">
          
          <div className="flex items-center justify-between mb-8">
            <div>
              <h1 className="text-3xl font-bold tracking-tight">Projects</h1>
              <p className="text-muted-foreground mt-2">Manage all your tenant migration projects.</p>
            </div>
            
            <Dialog open={isOpen} onOpenChange={(open) => { setIsOpen(open); if (!open) form.reset(); }}>
              <DialogTrigger asChild>
                <Button className="bg-primary hover:bg-primary/90 shadow-lg shadow-primary/20 transition-all">
                  <Plus className="w-4 h-4 mr-2" />
                  New Project
                </Button>
              </DialogTrigger>
              <DialogContent className="sm:max-w-[500px]">
                <DialogHeader>
                  <DialogTitle>Create New Project</DialogTitle>
                </DialogHeader>
                <form onSubmit={form.handleSubmit(onSubmit)} className="space-y-4 py-4">
                  <div className="space-y-2">
                    <Label htmlFor="name">Project Name</Label>
                    <Input id="name" {...form.register("name")} placeholder="e.g. Acme Corp Migration" />
                    {form.formState.errors.name && <p className="text-xs text-red-500">{form.formState.errors.name.message}</p>}
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-2">
                      <Label htmlFor="source">Source Tenant ID</Label>
                      <Input id="source" {...form.register("sourceTenantId")} placeholder="source.onmicrosoft.com" autoComplete="off" />
                      {form.formState.errors.sourceTenantId && <p className="text-xs text-red-500">{form.formState.errors.sourceTenantId.message}</p>}
                    </div>
                    <div className="space-y-2">
                      <Label htmlFor="target">Target Tenant ID</Label>
                      <Input id="target" {...form.register("targetTenantId")} placeholder="target.onmicrosoft.com" autoComplete="off" />
                      {form.formState.errors.targetTenantId && <p className="text-xs text-red-500">{form.formState.errors.targetTenantId.message}</p>}
                    </div>
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="description">Description</Label>
                    <Textarea id="description" {...form.register("description")} placeholder="Migration scope and details..." />
                  </div>
                  <div className="flex justify-end pt-4">
                    <Button type="button" variant="outline" onClick={() => setIsOpen(false)} className="mr-2">Cancel</Button>
                    <Button type="submit" disabled={isCreating}>
                      {isCreating ? <Loader2 className="w-4 h-4 animate-spin mr-2" /> : null}
                      Create Project
                    </Button>
                  </div>
                </form>
              </DialogContent>
            </Dialog>
          </div>

          <div className="flex items-center gap-4 mb-6">
            <div className="relative flex-1 max-w-sm">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <Input 
                className="pl-10 bg-white dark:bg-slate-900 border-border/60" 
                placeholder="Search projects..." 
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {filteredProjects?.map((project, idx) => (
              <motion.div
                key={project.id}
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: idx * 0.05 }}
              >
                <Link href={`/projects/${project.id}`} className="block h-full">
                  <div className="bg-card rounded-xl border border-border/60 p-6 shadow-sm hover:shadow-md hover:border-primary/50 transition-all duration-200 h-full flex flex-col group">
                    <div className="flex items-start justify-between mb-4">
                      <div className="w-10 h-10 rounded-lg bg-slate-100 dark:bg-slate-800 flex items-center justify-center group-hover:bg-primary/10 transition-colors">
                        <span className="text-lg font-bold text-slate-500 group-hover:text-primary">{project.name.substring(0, 1).toUpperCase()}</span>
                      </div>
                      <StatusBadge status={project.status} />
                    </div>
                    
                    <h3 className="font-bold text-lg mb-2 group-hover:text-primary transition-colors">{project.name}</h3>
                    <p className="text-sm text-muted-foreground line-clamp-2 mb-6 flex-1">
                      {project.description || "No description provided."}
                    </p>
                    
                    <div className="space-y-2 text-xs text-muted-foreground pt-4 border-t border-border/50">
                      <div className="flex justify-between">
                        <span>Source:</span>
                        <span className="font-medium text-foreground">{project.sourceTenantId}</span>
                      </div>
                      <div className="flex justify-between">
                        <span>Target:</span>
                        <span className="font-medium text-foreground">{project.targetTenantId}</span>
                      </div>
                    </div>
                  </div>
                </Link>
              </motion.div>
            ))}
          </div>

          {filteredProjects?.length === 0 && (
            <div className="text-center py-20 text-muted-foreground">
              No projects found matching your search.
            </div>
          )}

        </div>
      </main>
    </div>
  );
}
