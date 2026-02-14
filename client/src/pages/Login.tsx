import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Cloud, Lock } from "lucide-react";

export default function Login() {
  const handleLogin = () => {
    window.location.href = "/api/login";
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-slate-50 dark:bg-slate-950 p-4">
      <div className="w-full max-w-md">
        <div className="flex justify-center mb-8">
            <div className="w-16 h-16 rounded-2xl bg-gradient-to-tr from-primary to-accent flex items-center justify-center shadow-lg shadow-primary/20">
                <Cloud className="w-8 h-8 text-white" />
            </div>
        </div>
        
        <Card className="border-border/60 shadow-xl">
          <CardHeader className="text-center">
            <CardTitle className="text-2xl font-bold">Welcome Back</CardTitle>
            <CardDescription>Sign in to manage your M365 migrations</CardDescription>
          </CardHeader>
          <CardContent>
            <Button 
                onClick={handleLogin}
                className="w-full h-12 text-base font-medium shadow-lg shadow-primary/20 transition-all hover:scale-[1.02]"
            >
                <Lock className="w-4 h-4 mr-2" />
                Login with Replit
            </Button>
            
            <div className="mt-6 text-center text-xs text-muted-foreground">
                <p>Secure authentication powered by Replit.</p>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
