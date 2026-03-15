import { pgTable, text, serial, integer, bigint, timestamp, jsonb } from "drizzle-orm/pg-core";
import { createInsertSchema } from "drizzle-zod";
import { z } from "zod";
import { users } from "./models/auth";

// Export Auth Models
export * from "./models/auth";

// === MIGRATION PROJECTS ===
export const migrationProjects = pgTable("migration_projects", {
  id: serial("id").primaryKey(),
  name: text("name").notNull(),
  sourceTenantId: text("source_tenant_id").notNull(),
  targetTenantId: text("target_tenant_id").notNull(),
  sourceClientId: text("source_client_id"),
  sourceClientSecret: text("source_client_secret"),
  targetClientId: text("target_client_id"),
  targetClientSecret: text("target_client_secret"),
  status: text("status").default("draft").notNull(), // draft, active, completed, archived
  description: text("description"),
  userId: text("user_id").references(() => users.id), // Owner of the project
  createdAt: timestamp("created_at").defaultNow(),
});

// === MIGRATION ITEMS (Users/Resources to migrate) ===
export const migrationItems = pgTable("migration_items", {
  id: serial("id").primaryKey(),
  projectId: integer("project_id").notNull().references(() => migrationProjects.id),
  sourceIdentity: text("source_identity").notNull(), // e.g. user@source.com
  targetIdentity: text("target_identity"), // e.g. user@target.com
  itemType: text("item_type").default("mailbox").notNull(), // mailbox, onedrive, sharepoint, teams
  status: text("status").default("pending").notNull(), // pending, in_progress, completed, failed
  errorDetails: text("error_details"),
  logs: jsonb("logs").$type<string[]>(), // Array of log strings
  bytesTotal: bigint("bytes_total", { mode: "number" }),
  bytesMigrated: bigint("bytes_migrated", { mode: "number" }),
  progressPercent: integer("progress_percent"),
  updatedAt: timestamp("updated_at").defaultNow(),
});

// === MAPPING RULES ===
export const mappingRules = pgTable("mapping_rules", {
  id: serial("id").primaryKey(),
  projectId: integer("project_id").notNull().references(() => migrationProjects.id),
  ruleType: text("rule_type").notNull(), // 'domain', 'prefix', 'suffix', 'upn_prefix'
  sourcePattern: text("source_pattern").notNull(),
  targetPattern: text("target_pattern").notNull(),
  description: text("description"),
  createdAt: timestamp("created_at").defaultNow(),
});

export const insertMappingRuleSchema = createInsertSchema(mappingRules).omit({ id: true, createdAt: true });
export type MappingRule = typeof mappingRules.$inferSelect;
export type InsertMappingRule = z.infer<typeof insertMappingRuleSchema>;

// === SCHEMAS ===
export const insertProjectSchema = createInsertSchema(migrationProjects).omit({ id: true, createdAt: true });
export const insertItemSchema = createInsertSchema(migrationItems).omit({ id: true, updatedAt: true, logs: true });

// === TYPES ===
export type Project = typeof migrationProjects.$inferSelect;
export type InsertProject = z.infer<typeof insertProjectSchema>;

export type MigrationItem = typeof migrationItems.$inferSelect;
export type InsertMigrationItem = z.infer<typeof insertItemSchema>;

// API Request/Response Types
export type CreateProjectRequest = InsertProject;
export type UpdateProjectRequest = Partial<InsertProject>;

export type CreateItemRequest = InsertMigrationItem;
export type UpdateItemRequest = Partial<InsertMigrationItem>;
