# Microsoft 365 Tenant-to-Tenant Migration Manager

## Overview
Web application for managing Microsoft 365 tenant-to-tenant migrations and Entra ID to on-premises Active Directory migrations. Users can create migration projects, configure Graph API credentials, discover source resources, auto-map identities, and run or script migrations.

## Architecture
- **Frontend**: React + TypeScript, Vite, Wouter routing, TanStack Query, shadcn/ui, Tailwind CSS
- **Backend**: Express.js + TypeScript
- **Database**: PostgreSQL via Drizzle ORM
- **Auth**: Local username/password (bcryptjs + express-session). Default: admin/admin. Override with `ADMIN_USERNAME`/`ADMIN_PASSWORD` env vars.

## Key Files
- `shared/schema.ts` — Drizzle models: `migrationProjects`, `migrationItems`, `mappingRules`
- `server/routes.ts` — All Express route handlers
- `server/storage.ts` — DatabaseStorage (IStorage interface)
- `server/auth.ts` — Session auth, `isAuthenticated` middleware
- `server/services/migration-engine.ts` — Core migration logic (mailbox, onedrive, sharepoint, teams, user, entra_to_ad)
- `server/services/graph-client.ts` — Microsoft Graph API client with token caching
- `server/services/discovery-service.ts` — Discover users/sites/teams/power platform from Entra
- `server/services/entra-ad-service.ts` — Entra→AD LDAP migration + PowerShell script generation
- `client/src/pages/ProjectDetails.tsx` — Main project page (all 6 tabs)

## Data Model
- **migrationProjects**: id, name, sourceTenantId/targetTenantId, sourceClientId/Secret, targetClientId/Secret, status, description, userId, createdAt, adDcHostname, adLdapPort, adBindDn, adBindPassword, adBaseDn, adUseSsl, adTargetOu
- **migrationItems**: id, projectId, sourceIdentity, targetIdentity, itemType (mailbox/onedrive/sharepoint/teams/user/powerplatform/entra_to_ad), status, errorDetails, logs, bytesTotal, bytesMigrated, progressPercent, updatedAt
- **mappingRules**: id, projectId, ruleType (domain/prefix/suffix/upn_prefix), sourcePattern, targetPattern, description, createdAt

## Project Tabs
1. **Overview** — Stats, pie chart, tenant details
2. **Migration Items** — List, add, run, retry, view logs per item
3. **Discovery** — Scan source tenant (users/sites/teams/power platform), bulk import to queue
4. **Auto-Mapping Rules** — Configure UPN/domain transform rules with live preview tester
5. **Entra → AD** — Discover cloud-only Entra users, LDAP-migrate to on-premises AD, or export PowerShell script
6. **Tenant Configuration** — Graph API app registration credentials, Test Connection, Grant App Permissions

## Migration Types
- **mailbox** — Exchange Online via Graph API (copy messages, folders, attachments)
- **onedrive** — OneDrive files via Graph copy API
- **sharepoint** — SharePoint document libraries (all drives, recursive)
- **teams** — Microsoft Teams (create team/channels, migrate channel file libraries)
- **user** — Entra ID user account creation in target tenant
- **powerplatform** — Informational only (requires Power Platform CoE Starter Kit)
- **entra_to_ad** — Cloud-only Entra user → on-premises AD via LDAP (ldapjs) or PowerShell

## Local Windows Setup
- Run: `start.bat` (or `npm run dev`)
- Push DB schema: `dbpush.bat` (run after schema changes)
- Env vars: `DATABASE_URL`, `SESSION_SECRET`, `ADMIN_USERNAME`, `ADMIN_PASSWORD`
- See `LOCAL_SETUP.md` for full instructions

## Dependencies
- ldapjs — LDAP client for on-premises AD writes
- bcryptjs — Password hashing for local auth
- drizzle-orm + drizzle-kit — ORM + schema management
- @tanstack/react-query — Data fetching
