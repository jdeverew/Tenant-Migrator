import { useProjects } from "@/hooks/use-projects";
import { Sidebar } from "@/components/Sidebar";
import { StatusBadge } from "@/components/StatusBadge";
import { Link } from "wouter";
import { ArrowUpRight, Clock, CheckCircle2, AlertCircle, Loader2 } from "lucide-react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { motion } from "framer-motion";

export default function Home() {
  const { data: projects, isLoading } = useProjects();

  const totalProjects = projects?.length || 0;
  const activeProjects = projects?.filter(p => p.status === 'active').length || 0;
  const completedProjects = projects?.filter(p => p.status === 'completed').length || 0;

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
          
          <header className="mb-8">
            <h1 className="text-3xl font-bold tracking-tight">Dashboard Overview</h1>
            <p className="text-muted-foreground mt-2">Welcome back. Here's what's happening with your migrations.</p>
          </header>

          {/* Stats Grid */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-10">
            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.1 }}>
              <Card className="border-border/60 shadow-sm hover:shadow-md transition-all">
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Total Projects</CardTitle>
                  <Clock className="h-4 w-4 text-muted-foreground" />
                </CardHeader>
                <CardContent>
                  <div className="text-3xl font-bold">{totalProjects}</div>
                  <p className="text-xs text-muted-foreground mt-1">
                    All migration projects
                  </p>
                </CardContent>
              </Card>
            </motion.div>

            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.2 }}>
              <Card className="border-border/60 shadow-sm hover:shadow-md transition-all">
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Active Migrations</CardTitle>
                  <ArrowUpRight className="h-4 w-4 text-blue-500" />
                </CardHeader>
                <CardContent>
                  <div className="text-3xl font-bold text-blue-600 dark:text-blue-400">{activeProjects}</div>
                  <p className="text-xs text-muted-foreground mt-1">
                    Currently in progress
                  </p>
                </CardContent>
              </Card>
            </motion.div>

            <motion.div initial={{ opacity: 0, y: 20 }} animate={{ opacity: 1, y: 0 }} transition={{ delay: 0.3 }}>
              <Card className="border-border/60 shadow-sm hover:shadow-md transition-all">
                <CardHeader className="flex flex-row items-center justify-between space-y-0 pb-2">
                  <CardTitle className="text-sm font-medium">Completed</CardTitle>
                  <CheckCircle2 className="h-4 w-4 text-emerald-500" />
                </CardHeader>
                <CardContent>
                  <div className="text-3xl font-bold text-emerald-600 dark:text-emerald-400">{completedProjects}</div>
                  <p className="text-xs text-muted-foreground mt-1">
                    Successfully finished
                  </p>
                </CardContent>
              </Card>
            </motion.div>
          </div>

          <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
            Recent Projects
            <Link href="/projects" className="text-sm font-medium text-primary hover:underline ml-auto">View All</Link>
          </h2>

          <motion.div 
            initial={{ opacity: 0 }} 
            animate={{ opacity: 1 }} 
            transition={{ delay: 0.4 }}
            className="bg-card rounded-xl border border-border/60 shadow-sm overflow-hidden"
          >
            {projects && projects.length > 0 ? (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-muted/30 border-b border-border/60">
                    <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Project Name</th>
                    <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Source Tenant</th>
                    <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Target Tenant</th>
                    <th className="px-6 py-4 text-left font-semibold text-muted-foreground">Status</th>
                    <th className="px-6 py-4 text-right font-semibold text-muted-foreground">Action</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border/40">
                  {projects.slice(0, 5).map((project) => (
                    <tr key={project.id} className="hover:bg-slate-50 dark:hover:bg-slate-900/50 transition-colors">
                      <td className="px-6 py-4 font-medium">{project.name}</td>
                      <td className="px-6 py-4 text-muted-foreground">{project.sourceTenantId}</td>
                      <td className="px-6 py-4 text-muted-foreground">{project.targetTenantId}</td>
                      <td className="px-6 py-4">
                        <StatusBadge status={project.status} />
                      </td>
                      <td className="px-6 py-4 text-right">
                        <Link href={`/projects/${project.id}`} className="text-primary hover:text-primary/80 font-medium">
                          Manage
                        </Link>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div className="p-12 text-center">
                <div className="inline-flex h-12 w-12 items-center justify-center rounded-full bg-slate-100 dark:bg-slate-800 mb-4">
                  <AlertCircle className="h-6 w-6 text-muted-foreground" />
                </div>
                <h3 className="text-lg font-medium">No projects yet</h3>
                <p className="text-muted-foreground mb-4">Create your first migration project to get started.</p>
                <Link href="/projects">
                  <span className="inline-flex items-center justify-center rounded-lg bg-primary px-4 py-2 text-sm font-medium text-primary-foreground shadow transition-colors hover:bg-primary/90 focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-ring disabled:pointer-events-none disabled:opacity-50">
                    Create Project
                  </span>
                </Link>
              </div>
            )}
          </motion.div>

        </div>
      </main>
    </div>
  );
}
