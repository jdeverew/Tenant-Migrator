import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api, buildUrl } from "@shared/routes";
import { type InsertMigrationItem, type MigrationItem } from "@shared/schema";

export function useMigrationItems(projectId: number) {
  return useQuery({
    queryKey: [api.items.list.path, projectId],
    queryFn: async () => {
      const url = buildUrl(api.items.list.path, { projectId });
      const res = await fetch(url, { credentials: "include" });
      if (!res.ok) throw new Error("Failed to fetch migration items");
      return api.items.list.responses[200].parse(await res.json());
    },
    enabled: !isNaN(projectId),
    refetchInterval: (query) => {
      const data = query.state.data as MigrationItem[] | undefined;
      if (!data) return false;
      return data.some((i) => i.status === 'in_progress') ? 2000 : false;
    },
  });
}

export function useCreateMigrationItem() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ projectId, ...data }: InsertMigrationItem) => {
      const validated = api.items.create.input.parse(data);
      const url = buildUrl(api.items.create.path, { projectId });
      
      const res = await fetch(url, {
        method: api.items.create.method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(validated),
        credentials: "include",
      });
      
      if (!res.ok) {
        if (res.status === 400) {
          const error = api.items.create.responses[400].parse(await res.json());
          throw new Error(error.message);
        }
        throw new Error("Failed to create migration item");
      }
      return api.items.create.responses[201].parse(await res.json());
    },
    onSuccess: (data, variables) => {
      queryClient.invalidateQueries({ queryKey: [api.items.list.path, variables.projectId] });
      queryClient.invalidateQueries({ queryKey: [api.projects.stats.path, variables.projectId] });
    },
  });
}

export function useUpdateMigrationItem() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, projectId, ...updates }: { id: number; projectId: number } & Partial<InsertMigrationItem>) => {
      const validated = api.items.update.input.parse(updates);
      const url = buildUrl(api.items.update.path, { id });
      
      const res = await fetch(url, {
        method: api.items.update.method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(validated),
        credentials: "include",
      });
      
      if (!res.ok) throw new Error("Failed to update migration item");
      return api.items.update.responses[200].parse(await res.json());
    },
    onSuccess: (data, variables) => {
      queryClient.invalidateQueries({ queryKey: [api.items.list.path, variables.projectId] });
      queryClient.invalidateQueries({ queryKey: [api.projects.stats.path, variables.projectId] });
    },
  });
}

export function useDeleteMigrationItem() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, projectId }: { id: number; projectId: number }) => {
      const url = buildUrl(api.items.delete.path, { id });
      const res = await fetch(url, { method: api.items.delete.method, credentials: "include" });
      if (!res.ok) throw new Error("Failed to delete migration item");
    },
    onSuccess: (data, variables) => {
      queryClient.invalidateQueries({ queryKey: [api.items.list.path, variables.projectId] });
      queryClient.invalidateQueries({ queryKey: [api.projects.stats.path, variables.projectId] });
    },
  });
}
