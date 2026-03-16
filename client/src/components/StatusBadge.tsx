import { cn } from "@/lib/utils";

type Status = "draft" | "active" | "completed" | "archived" | "pending" | "in_progress" | "failed" | "needs_action" | "cancelled" | "reverting" | "reverted" | "revert_failed";

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
    needs_action: "bg-amber-50 text-amber-700 border-amber-300 dark:bg-amber-950/30 dark:text-amber-400 dark:border-amber-700",
    cancelled: "bg-slate-100 text-slate-500 border-slate-300 dark:bg-slate-800/40 dark:text-slate-400 dark:border-slate-600",
    reverting: "bg-violet-50 text-violet-700 border-violet-200 dark:bg-violet-950/30 dark:text-violet-400 dark:border-violet-700",
    reverted: "bg-orange-50 text-orange-700 border-orange-200 dark:bg-orange-950/30 dark:text-orange-400 dark:border-orange-700",
    revert_failed: "bg-red-50 text-red-800 border-red-300 dark:bg-red-950/30 dark:text-red-400 dark:border-red-700",
  };

  const dotStyles: Record<string, string> = {
    completed: "bg-emerald-500",
    failed: "bg-red-500",
    needs_action: "bg-amber-500",
    in_progress: "bg-blue-500",
    active: "bg-blue-500",
    cancelled: "bg-slate-400",
    reverting: "bg-violet-500 animate-pulse",
    reverted: "bg-orange-500",
    revert_failed: "bg-red-600",
  };

  const labelOverrides: Record<string, string> = {
    needs_action: "Needs Action",
    in_progress: "In Progress",
    revert_failed: "Revert Failed",
  };

  const normalizedStatus = status.toLowerCase();
  const style = styles[normalizedStatus] || styles.draft;
  const dotStyle = dotStyles[normalizedStatus] || "bg-slate-400";
  const label = labelOverrides[normalizedStatus] ?? normalizedStatus.replace(/_/g, ' ');

  return (
    <span
      className={cn(
        "inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold border capitalize shadow-sm",
        style,
        className
      )}
    >
      <span className={cn("w-1.5 h-1.5 rounded-full mr-1.5", dotStyle)} />
      {label}
    </span>
  );
}
