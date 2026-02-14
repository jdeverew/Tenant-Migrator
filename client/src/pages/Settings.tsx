import { Sidebar } from "@/components/Sidebar";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { useAuth } from "@/hooks/use-auth";

export default function Settings() {
    const { user } = useAuth();

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 text-foreground font-sans">
      <Sidebar />
      
      <main className="flex-1 overflow-y-auto">
        <div className="container max-w-4xl mx-auto px-8 py-8">
            <h1 className="text-3xl font-bold tracking-tight mb-8">Settings</h1>
            
            <div className="space-y-6">
                <Card>
                    <CardHeader>
                        <CardTitle>Account Information</CardTitle>
                    </CardHeader>
                    <CardContent className="space-y-4">
                        <div className="grid grid-cols-2 gap-4">
                             <div>
                                <label className="text-sm font-medium text-muted-foreground">Username</label>
                                <p className="text-base">{user?.username || 'Loading...'}</p>
                             </div>
                             <div>
                                <label className="text-sm font-medium text-muted-foreground">Role</label>
                                <p className="text-base">{user?.isAdmin ? 'Administrator' : 'User'}</p>
                             </div>
                        </div>
                    </CardContent>
                </Card>
                
                <Card>
                    <CardHeader>
                        <CardTitle>App Configuration</CardTitle>
                    </CardHeader>
                    <CardContent>
                        <p className="text-muted-foreground">Global application settings placeholder.</p>
                    </CardContent>
                </Card>
            </div>
        </div>
      </main>
    </div>
  );
}
