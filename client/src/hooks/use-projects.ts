import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { api, buildUrl } from "@shared/routes";
import { type InsertProject, type Project, type ProjectStats } from "@shared/schema";

export function useProjects() {
  return useQuery({
    queryKey: [api.projects.list.path],
    queryFn: async () => {
      const res = await fetch(api.projects.list.path, { credentials: "include" });
      if (!res.ok) throw new Error("Failed to fetch projects");
      return api.projects.list.responses[200].parse(await res.json());
    },
  });
}

export function useProject(id: number) {
  return useQuery({
    queryKey: [api.projects.get.path, id],
    queryFn: async () => {
      const url = buildUrl(api.projects.get.path, { id });
      const res = await fetch(url, { credentials: "include" });
      if (res.status === 404) return null;
      if (!res.ok) throw new Error("Failed to fetch project");
      return api.projects.get.responses[200].parse(await res.json());
    },
    enabled: !isNaN(id),
  });
}

export function useProjectStats(id: number) {
  return useQuery({
    queryKey: [api.projects.stats.path, id],
    queryFn: async () => {
      const url = buildUrl(api.projects.stats.path, { id });
      const res = await fetch(url, { credentials: "include" });
      if (!res.ok) throw new Error("Failed to fetch project stats");
      return api.projects.stats.responses[200].parse(await res.json());
    },
    enabled: !isNaN(id),
  });
}

export function useCreateProject() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (data: InsertProject) => {
      const validated = api.projects.create.input.parse(data);
      const res = await fetch(api.projects.create.path, {
        method: api.projects.create.method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(validated),
        credentials: "include",
      });
      if (!res.ok) {
        if (res.status === 400) {
          const error = api.projects.create.responses[400].parse(await res.json());
          throw new Error(error.message);
        }
        throw new Error("Failed to create project");
      }
      return api.projects.create.responses[201].parse(await res.json());
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [api.projects.list.path] });
    },
  });
}

export function useUpdateProject() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, ...updates }: { id: number } & Partial<InsertProject>) => {
      const validated = api.projects.update.input.parse(updates);
      const url = buildUrl(api.projects.update.path, { id });
      const res = await fetch(url, {
        method: api.projects.update.method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(validated),
        credentials: "include",
      });
      if (!res.ok) throw new Error("Failed to update project");
      return api.projects.update.responses[200].parse(await res.json());
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: [api.projects.list.path] });
      queryClient.invalidateQueries({ queryKey: [api.projects.get.path, data.id] });
    },
  });
}

export function useDeleteProject() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: number) => {
      const url = buildUrl(api.projects.delete.path, { id });
      const res = await fetch(url, { method: api.projects.delete.method, credentials: "include" });
      if (!res.ok) throw new Error("Failed to delete project");
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [api.projects.list.path] });
    },
  });
}
