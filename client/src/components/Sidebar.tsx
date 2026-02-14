import { Link, useLocation } from "wouter";
import { LayoutDashboard, Settings, LogOut, Cloud, ArrowLeftRight } from "lucide-react";
import { cn } from "@/lib/utils";
import { useAuth } from "@/hooks/use-auth";

export function Sidebar() {
  const [location] = useLocation();
  const { user, logout } = useAuth();

  const menuItems = [
    { icon: LayoutDashboard, label: "Dashboard", href: "/" },
    { icon: ArrowLeftRight, label: "Projects", href: "/projects" },
    { icon: Settings, label: "Settings", href: "/settings" },
  ];

  return (
    <div className="w-64 border-r border-border/50 h-screen flex flex-col bg-slate-50/50 dark:bg-slate-900/50 backdrop-blur-xl sticky top-0">
      <div className="p-6">
        <div className="flex items-center gap-3 mb-8">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-primary to-accent flex items-center justify-center shadow-lg shadow-primary/20">
            <Cloud className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="font-bold text-lg leading-tight tracking-tight">CloudMigrate</h1>
            <p className="text-xs text-muted-foreground font-medium">M365 Manager</p>
          </div>
        </div>

        <nav className="space-y-1">
          {menuItems.map((item) => {
            const Icon = item.icon;
            const isActive = location === item.href || (item.href !== "/" && location.startsWith(item.href));
            
            return (
              <Link key={item.href} href={item.href}>
                <div 
                  className={cn(
                    "flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-all duration-200 cursor-pointer group",
                    isActive 
                      ? "bg-primary/10 text-primary shadow-sm ring-1 ring-primary/20" 
                      : "text-muted-foreground hover:bg-white hover:text-foreground hover:shadow-sm"
                  )}
                >
                  <Icon className={cn("w-5 h-5 transition-colors", isActive ? "text-primary" : "text-muted-foreground group-hover:text-foreground")} />
                  {item.label}
                </div>
              </Link>
            );
          })}
        </nav>
      </div>

      <div className="mt-auto p-6 border-t border-border/50">
        <div className="flex items-center gap-3 mb-4">
          <div className="w-10 h-10 rounded-full bg-slate-200 dark:bg-slate-800 overflow-hidden border-2 border-white dark:border-slate-700 shadow-sm">
            {user?.profileImageUrl ? (
               <img src={user.profileImageUrl} alt="User" className="w-full h-full object-cover" />
            ) : (
              <div className="w-full h-full flex items-center justify-center text-xs font-bold text-muted-foreground">
                {user?.username?.substring(0, 2).toUpperCase() || "US"}
              </div>
            )}
          </div>
          <div className="overflow-hidden">
            <p className="text-sm font-semibold truncate">{user?.username || "Guest User"}</p>
            <p className="text-xs text-muted-foreground truncate">{user?.username || "Guest User"}</p>
          </div>
        </div>
        
        <button 
          onClick={() => logout()}
          className="w-full flex items-center justify-center gap-2 px-4 py-2 rounded-lg text-xs font-medium bg-white dark:bg-slate-800 border border-border hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors shadow-sm"
        >
          <LogOut className="w-3.5 h-3.5" />
          Sign Out
        </button>
      </div>
    </div>
  );
}
