import { cn } from "@/lib/utils";

type Status = "draft" | "active" | "completed" | "archived" | "pending" | "in_progress" | "failed";

interface StatusBadgeProps {
  status: Status | string;
  className?: string;
}

export function StatusBadge({ status, className }: StatusBadgeProps) {
  const styles: Record<string, string> = {
    draft: "bg-slate-100 text-slate-600 border-slate-200",
    pending: "bg-slate-100 text-slate-600 border-slate-200",
    active: "bg-blue-50 text-blue-700 border-blue-200",
    in_progress: "bg-blue-50 text-blue-700 border-blue-200",
    completed: "bg-emerald-50 text-emerald-700 border-emerald-200",
    failed: "bg-red-50 text-red-700 border-red-200",
    archived: "bg-amber-50 text-amber-700 border-amber-200",
  };

  const normalizedStatus = status.toLowerCase();
  const style = styles[normalizedStatus] || styles.draft;
  
  const label = normalizedStatus.replace('_', ' ');

  return (
    <span 
      className={cn(
        "inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold border capitalize shadow-sm",
        style,
        className
      )}
    >
      <span className={cn(
        "w-1.5 h-1.5 rounded-full mr-1.5", 
        normalizedStatus === "completed" ? "bg-emerald-500" :
        normalizedStatus === "failed" ? "bg-red-500" :
        normalizedStatus === "in_progress" || normalizedStatus === "active" ? "bg-blue-500" :
        "bg-slate-400"
      )} />
      {label}
    </span>
  );
}
