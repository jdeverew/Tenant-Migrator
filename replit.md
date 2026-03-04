# Microsoft 365 Tenant-to-Tenant Migration Manager

## Overview
Web application for managing Microsoft 365 tenant-to-tenant migrations. Users can create migration projects, configure Graph API credentials for source/target tenants, map users/resources between tenants, and track migration status.

## Architecture
- **Frontend**: React + TypeScript, Vite, Wouter routing, TanStack Query, shadcn/ui, Tailwind CSS
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL via Drizzle ORM
- **Auth**: Replit Auth (OIDC — Google/GitHub/email)

## Key Files
- `shared/schema.ts` — Drizzle models: `migrationProjects` (with Graph API credentials), `migrationItems`
- `shared/routes.ts` — API contract definitions with Zod validation
- `server/routes.ts` — Express route handlers including `/api/projects/:id/test-connection`
- `server/storage.ts` — DatabaseStorage class implementing IStorage interface
- `client/src/pages/ProjectDetails.tsx` — Project detail page with Overview, Migration Items, and Tenant Configuration tabs
- `client/src/hooks/use-projects.ts` — React Query hooks for project CRUD
- `client/src/hooks/use-items.ts` — React Query hooks for migration item CRUD

## Data Model
- **migrationProjects**: id, name, sourceTenantId, targetTenantId, sourceClientId, sourceClientSecret, targetClientId, targetClientSecret, status, description, userId, createdAt
- **migrationItems**: id, projectId, sourceIdentity, targetIdentity, itemType (mailbox/onedrive/teams), status (pending/in_progress/completed/failed), errorDetails, logs, updatedAt

## Features
- Project CRUD with status workflow (draft → active → completed/archived)
- Tenant configuration with Microsoft Entra ID App Registration credentials (Client ID, Client Secret)
- Test Connection button that validates credentials against Microsoft Graph API (OAuth2 client_credentials flow)
- Migration item mapping (source ↔ target identity)
- Dashboard with pie chart progress visualization
- Seed data for demo purposes

## Running
- `npm run dev` starts both Express backend and Vite frontend on port 5000
- `npm run db:push` syncs Drizzle schema to PostgreSQL
