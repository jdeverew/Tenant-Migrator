import { db } from "./db";
import {
  migrationProjects,
  migrationItems,
  mappingRules,
  type Project,
  type InsertProject,
  type MigrationItem,
  type InsertMigrationItem,
  type UpdateProjectRequest,
  type UpdateItemRequest,
  type MappingRule,
  type InsertMappingRule,
} from "@shared/schema";
import { eq, desc, and, count } from "drizzle-orm";

export interface IStorage {
  // Projects
  getProjects(userId?: string): Promise<Project[]>;
  getProject(id: number): Promise<Project | undefined>;
  createProject(project: InsertProject): Promise<Project>;
  updateProject(id: number, updates: UpdateProjectRequest): Promise<Project>;
  deleteProject(id: number): Promise<void>;
  
  // Project Stats
  getProjectStats(projectId: number): Promise<{
    total: number;
    pending: number;
    inProgress: number;
    completed: number;
    failed: number;
  }>;

  // Items
  getItems(projectId: number): Promise<MigrationItem[]>;
  getItem(id: number): Promise<MigrationItem | undefined>;
  createItem(item: InsertMigrationItem): Promise<MigrationItem>;
  updateItem(id: number, updates: UpdateItemRequest): Promise<MigrationItem>;
  updateItemLogs(id: number, logs: string[]): Promise<void>;
  deleteItem(id: number): Promise<void>;

  // Mapping Rules
  getMappingRules(projectId: number): Promise<MappingRule[]>;
  createMappingRule(rule: InsertMappingRule): Promise<MappingRule>;
  deleteMappingRule(id: number): Promise<void>;
  applyMappingRules(projectId: number, sourceIdentity: string): Promise<string>;
}

export class DatabaseStorage implements IStorage {
  async getProjects(userId?: string): Promise<Project[]> {
    if (userId) {
      return await db.select().from(migrationProjects).where(eq(migrationProjects.userId, userId)).orderBy(desc(migrationProjects.createdAt));
    }
    return await db.select().from(migrationProjects).orderBy(desc(migrationProjects.createdAt));
  }

  async getProject(id: number): Promise<Project | undefined> {
    const [project] = await db.select().from(migrationProjects).where(eq(migrationProjects.id, id));
    return project;
  }

  async createProject(project: InsertProject): Promise<Project> {
    const [newProject] = await db.insert(migrationProjects).values(project).returning();
    return newProject;
  }

  async updateProject(id: number, updates: UpdateProjectRequest): Promise<Project> {
    const [updated] = await db.update(migrationProjects)
      .set(updates)
      .where(eq(migrationProjects.id, id))
      .returning();
    return updated;
  }

  async deleteProject(id: number): Promise<void> {
    // Delete items first (cascade simulation if needed, but ideally DB handles this via foreign keys if configured, 
    // strictly speaking we should delete items first to be safe or use ON DELETE CASCADE in schema)
    await db.delete(migrationItems).where(eq(migrationItems.projectId, id));
    await db.delete(migrationProjects).where(eq(migrationProjects.id, id));
  }

  async getProjectStats(projectId: number) {
    const items = await db.select({
      status: migrationItems.status,
      count: count(),
    })
    .from(migrationItems)
    .where(eq(migrationItems.projectId, projectId))
    .groupBy(migrationItems.status);

    const stats = {
      total: 0,
      pending: 0,
      inProgress: 0,
      completed: 0,
      failed: 0,
    };

    items.forEach(item => {
      stats.total += item.count;
      if (item.status === 'pending') stats.pending = item.count;
      else if (item.status === 'in_progress') stats.inProgress = item.count;
      else if (item.status === 'completed') stats.completed = item.count;
      else if (item.status === 'failed') stats.failed = item.count;
    });

    return stats;
  }

  async getItems(projectId: number): Promise<MigrationItem[]> {
    return await db.select().from(migrationItems).where(eq(migrationItems.projectId, projectId));
  }

  async getItem(id: number): Promise<MigrationItem | undefined> {
    const [item] = await db.select().from(migrationItems).where(eq(migrationItems.id, id));
    return item;
  }

  async createItem(item: InsertMigrationItem): Promise<MigrationItem> {
    const [newItem] = await db.insert(migrationItems).values(item).returning();
    return newItem;
  }

  async updateItem(id: number, updates: UpdateItemRequest): Promise<MigrationItem> {
    const [updated] = await db.update(migrationItems)
      .set({ ...updates, updatedAt: new Date() })
      .where(eq(migrationItems.id, id))
      .returning();
    return updated;
  }

  async updateItemLogs(id: number, logs: string[]): Promise<void> {
    await db.update(migrationItems)
      .set({ logs, updatedAt: new Date() })
      .where(eq(migrationItems.id, id));
  }

  async deleteItem(id: number): Promise<void> {
    await db.delete(migrationItems).where(eq(migrationItems.id, id));
  }

  async getMappingRules(projectId: number): Promise<MappingRule[]> {
    return await db.select().from(mappingRules).where(eq(mappingRules.projectId, projectId));
  }

  async createMappingRule(rule: InsertMappingRule): Promise<MappingRule> {
    const [created] = await db.insert(mappingRules).values(rule).returning();
    return created;
  }

  async deleteMappingRule(id: number): Promise<void> {
    await db.delete(mappingRules).where(eq(mappingRules.id, id));
  }

  async applyMappingRules(projectId: number, sourceIdentity: string): Promise<string> {
    const rules = await this.getMappingRules(projectId);
    let result = sourceIdentity;

    for (const rule of rules) {
      if (rule.ruleType === 'domain') {
        // Replace @sourcedomain.com with @targetdomain.com
        result = result.replace(new RegExp(`@${escapeRegex(rule.sourcePattern)}`, 'gi'), `@${rule.targetPattern}`);
      } else if (rule.ruleType === 'prefix') {
        // Add prefix to username part (before @)
        const atIndex = result.indexOf('@');
        if (atIndex > -1) {
          const username = result.substring(0, atIndex);
          const domain = result.substring(atIndex);
          if (username.startsWith(rule.sourcePattern)) {
            result = rule.targetPattern + username.substring(rule.sourcePattern.length) + domain;
          }
        }
      } else if (rule.ruleType === 'suffix') {
        // Add suffix to username part (before @)
        const atIndex = result.indexOf('@');
        if (atIndex > -1) {
          const username = result.substring(0, atIndex);
          const domain = result.substring(atIndex);
          result = username.replace(new RegExp(`${escapeRegex(rule.sourcePattern)}$`, 'i'), rule.targetPattern) + domain;
        }
      } else if (rule.ruleType === 'upn_prefix') {
        // Replace full UPN prefix (everything before first dot in username)
        const atIndex = result.indexOf('@');
        if (atIndex > -1) {
          const username = result.substring(0, atIndex);
          const domain = result.substring(atIndex);
          result = username.replace(rule.sourcePattern, rule.targetPattern) + domain;
        }
      }
    }

    return result;
  }
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export const storage = new DatabaseStorage();
