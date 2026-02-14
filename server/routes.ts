import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { setupAuth, registerAuthRoutes } from "./replit_integrations/auth";
import { api } from "@shared/routes";
import { z } from "zod";

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  // Set up Replit Auth first
  await setupAuth(app);
  registerAuthRoutes(app);

  // === Projects ===
  app.get(api.projects.list.path, async (req, res) => {
    // Optional: Filter by user if we wanted to enforce ownership
    // const userId = req.user?.claims?.sub;
    const projects = await storage.getProjects();
    res.json(projects);
  });

  app.get(api.projects.get.path, async (req, res) => {
    const project = await storage.getProject(Number(req.params.id));
    if (!project) {
      return res.status(404).json({ message: 'Project not found' });
    }
    res.json(project);
  });

  app.post(api.projects.create.path, async (req, res) => {
    try {
      const input = api.projects.create.input.parse(req.body);
      // Optional: attach user ID from auth
      // const userId = req.user?.claims?.sub;
      // if (userId) input.userId = userId;
      
      const project = await storage.createProject(input);
      res.status(201).json(project);
    } catch (err) {
      if (err instanceof z.ZodError) {
        return res.status(400).json({
          message: err.errors[0].message,
          field: err.errors[0].path.join('.'),
        });
      }
      throw err;
    }
  });

  app.patch(api.projects.update.path, async (req, res) => {
    try {
      const input = api.projects.update.input.parse(req.body);
      const project = await storage.updateProject(Number(req.params.id), input);
      if (!project) return res.status(404).json({ message: 'Project not found' });
      res.json(project);
    } catch (err) {
      if (err instanceof z.ZodError) return res.status(400).json(err);
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.delete(api.projects.delete.path, async (req, res) => {
    await storage.deleteProject(Number(req.params.id));
    res.status(204).end();
  });

  app.get(api.projects.stats.path, async (req, res) => {
    const stats = await storage.getProjectStats(Number(req.params.id));
    res.json(stats);
  });

  // === Items ===
  app.get(api.items.list.path, async (req, res) => {
    const items = await storage.getItems(Number(req.params.projectId));
    res.json(items);
  });

  app.post(api.items.create.path, async (req, res) => {
    try {
      const input = api.items.create.input.parse(req.body);
      const item = await storage.createItem({
        ...input,
        projectId: Number(req.params.projectId)
      });
      res.status(201).json(item);
    } catch (err) {
      if (err instanceof z.ZodError) return res.status(400).json(err);
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.patch(api.items.update.path, async (req, res) => {
    try {
      const input = api.items.update.input.parse(req.body);
      const item = await storage.updateItem(Number(req.params.id), input);
      if (!item) return res.status(404).json({ message: 'Item not found' });
      res.json(item);
    } catch (err) {
      if (err instanceof z.ZodError) return res.status(400).json(err);
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.delete(api.items.delete.path, async (req, res) => {
    await storage.deleteItem(Number(req.params.id));
    res.status(204).end();
  });

  // Seed data if empty
  await seedDatabase();

  return httpServer;
}

async function seedDatabase() {
  const projects = await storage.getProjects();
  if (projects.length === 0) {
    const p1 = await storage.createProject({
      name: "Acme Corp to Contoso Migration",
      sourceTenantId: "acme-corp-id-123",
      targetTenantId: "contoso-id-456",
      status: "active",
      description: "Migrating 50 users from Acme to Contoso."
    });

    await storage.createItem({
      projectId: p1.id,
      sourceIdentity: "john.doe@acme.com",
      targetIdentity: "john.doe@contoso.com",
      itemType: "mailbox",
      status: "completed"
    });
    
    await storage.createItem({
      projectId: p1.id,
      sourceIdentity: "jane.smith@acme.com",
      targetIdentity: "jane.smith@contoso.com",
      itemType: "mailbox",
      status: "in_progress"
    });

    await storage.createItem({
      projectId: p1.id,
      sourceIdentity: "bob.jones@acme.com",
      targetIdentity: "bob.jones@contoso.com",
      itemType: "onedrive",
      status: "failed",
      errorDetails: "Permission denied on source drive"
    });
  }
}
