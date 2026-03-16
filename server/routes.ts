import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { setupSession, registerAuthRoutes, isAuthenticated as requireAuth } from "./auth";
import { api } from "@shared/routes";
import { z } from "zod";
import type { Project } from "@shared/schema";
import { migrateItem, migrateAllPending } from "./services/migration-engine";
import { requestCancellation } from "./services/cancellation";
import { revertMigrationItem } from "./services/revert-engine";
import { GraphClient } from "./services/graph-client";
import { discoverUsers, discoverSharePointSites, discoverTeams, discoverPowerPlatform, discoverOneDrives, discoverDistributionGroups, discoverSharedMailboxes, discoverM365Groups } from "./services/discovery-service";
import { discoverCloudOnlyUsers, testAdConnection, migrateUserToAd, generatePowerShellScript, type AdConnectionConfig } from "./services/entra-ad-service";
import { buildConnectUrl, handleOAuthCallback, buildConsentUrl, buildRegrantConsentUrl, decodeConsentState, SERVICE_PERMISSION_GROUPS, type ServiceKey } from "./services/oauth-tenant-service";

function maskSecret(value: string | null): string | null {
  if (!value) return null;
  if (value.length <= 8) return '••••••••';
  return value.slice(0, 4) + '••••••••' + value.slice(-4);
}

function sanitizeProject(project: Project) {
  return {
    ...project,
    sourceClientSecret: maskSecret(project.sourceClientSecret),
    targetClientSecret: maskSecret(project.targetClientSecret),
    adBindPassword: maskSecret(project.adBindPassword),
  };
}

export async function registerRoutes(
  httpServer: Server,
  app: Express
): Promise<Server> {
  // Set up session and local auth
  setupSession(app);
  registerAuthRoutes(app);

  const getSessionUserId = (req: any): string => (req.session as any)?.user?.id;

  // === Projects ===
  app.get(api.projects.list.path, requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const projects = await storage.getProjects(userId);
    res.json(projects.map(sanitizeProject));
  });

  app.get(api.projects.get.path, requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) {
      return res.status(404).json({ message: 'Project not found' });
    }
    res.json(sanitizeProject(project));
  });

  app.post(api.projects.create.path, requireAuth, async (req, res) => {
    try {
      const input = api.projects.create.input.parse(req.body);
      const userId = getSessionUserId(req);
      const project = await storage.createProject({ ...input, userId });
      res.status(201).json(sanitizeProject(project));
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

  app.patch(api.projects.update.path, requireAuth, async (req, res) => {
    try {
      const input = api.projects.update.input.parse(req.body);
      const userId = getSessionUserId(req);
      const project = await storage.updateProject(Number(req.params.id), input, userId);
      if (!project) return res.status(404).json({ message: 'Project not found' });
      res.json(sanitizeProject(project));
    } catch (err) {
      if (err instanceof z.ZodError) return res.status(400).json(err);
      res.status(500).json({ message: "Internal server error" });
    }
  });

  app.delete(api.projects.delete.path, requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    await storage.deleteProject(Number(req.params.id), userId);
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

  // === Migration Endpoints ===
  app.post('/api/projects/:projectId/items/:itemId/migrate', async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);
      const itemId = Number(req.params.itemId);

      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const item = await storage.getItem(itemId);
      if (!item || item.projectId !== projectId) return res.status(404).json({ message: 'Item not found' });

      if (item.status === 'in_progress') {
        return res.status(409).json({ message: 'Migration already in progress for this item' });
      }

      migrateItem(projectId, itemId).catch(err => {
        console.error(`[migration] item ${itemId} (${item.itemType}) FAILED:`, err.message);
      });

      res.json({ message: 'Migration started', itemId });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Internal server error' });
    }
  });

  app.post('/api/projects/:projectId/migrate-all', async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);

      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const result = await migrateAllPending(projectId);
      res.json({
        message: `Started migration for ${result.started} items`,
        started: result.started,
        errors: result.errors,
      });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Internal server error' });
    }
  });

  // === Cancel a running migration ===
  app.post('/api/projects/:projectId/items/:itemId/cancel', async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);
      const itemId = Number(req.params.itemId);

      const item = await storage.getItem(itemId);
      if (!item || item.projectId !== projectId) return res.status(404).json({ message: 'Item not found' });

      if (item.status !== 'in_progress') {
        return res.status(409).json({ message: `Cannot cancel — item is currently "${item.status}", not "in_progress"` });
      }

      requestCancellation(itemId);
      res.json({ message: 'Cancellation requested — migration will stop at the next safe checkpoint.' });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Internal server error' });
    }
  });

  // === Revert a completed migration item ===
  app.post('/api/projects/:projectId/items/:itemId/revert', async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);
      const itemId = Number(req.params.itemId);

      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const item = await storage.getItem(itemId);
      if (!item || item.projectId !== projectId) return res.status(404).json({ message: 'Item not found' });

      const revertableStatuses = ['completed', 'reverted', 'revert_failed', 'cancelled'];
      if (!revertableStatuses.includes(item.status)) {
        return res.status(409).json({ message: `Cannot revert an item with status "${item.status}". Only completed items can be reverted.` });
      }

      // Run revert in background — caller can poll for status change
      revertMigrationItem(item, project).catch(err => {
        console.error(`[revert] item ${itemId} FAILED:`, err.message);
      });

      res.json({ message: 'Revert started — check item logs for progress.', itemId });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Internal server error' });
    }
  });

  app.get('/api/items/:id/logs', async (req, res) => {
    const item = await storage.getItem(Number(req.params.id));
    if (!item) return res.status(404).json({ message: 'Item not found' });
    res.json({ logs: item.logs || [] });
  });

  // === Export migration logs for a project ===
  app.get('/api/projects/:projectId/export-logs', requireAuth, async (req, res) => {
    try {
      const userId = getSessionUserId(req);
      const projectId = Number(req.params.projectId);
      const project = await storage.getProject(projectId, userId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const items = await storage.getItems(projectId);
      const fmt = (req.query.format as string || 'txt').toLowerCase();

      if (fmt === 'csv') {
        const escape = (v: string) => `"${(v || '').replace(/"/g, '""')}"`;
        const rows = [
          ['Type', 'Source Identity', 'Target Identity', 'Status', 'Error', 'Updated At', 'Log Lines'].map(escape).join(','),
          ...items.map(i => [
            i.itemType,
            i.sourceIdentity,
            i.targetIdentity || '',
            i.status,
            i.errorDetails || '',
            i.updatedAt ? new Date(i.updatedAt).toISOString() : '',
            (i.logs || []).join(' | '),
          ].map(escape).join(',')),
        ];
        res.setHeader('Content-Type', 'text/csv');
        res.setHeader('Content-Disposition', `attachment; filename="migration-logs-${projectId}-${Date.now()}.csv"`);
        return res.send(rows.join('\r\n'));
      }

      // Default: plain text
      const sep = '='.repeat(80);
      const thin = '-'.repeat(80);
      const lines: string[] = [
        sep,
        `M365 TENANT MIGRATION LOG EXPORT`,
        `Project : ${project.name}`,
        `Exported: ${new Date().toISOString()}`,
        `Items   : ${items.length} total, ${items.filter(i => i.status === 'failed').length} failed, ${items.filter(i => i.status === 'completed').length} completed`,
        sep,
        '',
      ];

      const typeOrder = ['mailbox','onedrive','sharepoint','teams','user','distributiongroup','sharedmailbox','m365group','powerplatform','entriad'];
      const sorted = [...items].sort((a, b) => typeOrder.indexOf(a.itemType) - typeOrder.indexOf(b.itemType));

      for (const item of sorted) {
        lines.push(sep);
        lines.push(`[${item.itemType.toUpperCase()}] ${item.sourceIdentity}${item.targetIdentity ? ` → ${item.targetIdentity}` : ''}`);
        lines.push(`Status : ${item.status.toUpperCase()}`);
        if (item.updatedAt) lines.push(`Updated: ${new Date(item.updatedAt).toISOString()}`);
        if (item.errorDetails) {
          lines.push(`Error  : ${item.errorDetails}`);
        }
        if ((item.logs || []).length > 0) {
          lines.push(thin);
          lines.push('LOGS:');
          for (const entry of item.logs!) {
            lines.push(`  ${entry}`);
          }
        } else {
          lines.push(`(no logs recorded)`);
        }
        lines.push('');
      }

      lines.push(sep);
      lines.push(`END OF REPORT`);
      lines.push(sep);

      res.setHeader('Content-Type', 'text/plain; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="migration-logs-${projectId}-${Date.now()}.txt"`);
      return res.send(lines.join('\n'));
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Export failed' });
    }
  });

  // === Test Connection ===
  app.post('/api/projects/:id/test-connection', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });

    const { tenant } = req.body as { tenant: 'source' | 'target' };
    if (!tenant || !['source', 'target'].includes(tenant)) {
      return res.status(400).json({ message: 'Invalid tenant parameter. Must be "source" or "target".' });
    }

    const tenantId = tenant === 'source' ? project.sourceTenantId : project.targetTenantId;
    const clientId = tenant === 'source' ? project.sourceClientId : project.targetClientId;
    const clientSecret = tenant === 'source' ? project.sourceClientSecret : project.targetClientSecret;

    if (!tenantId || !clientId || !clientSecret) {
      return res.json({
        success: false,
        message: `Missing credentials for ${tenant} tenant. Please configure Tenant ID, Client ID, and Client Secret.`,
      });
    }

    try {
      const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
      const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      });

      const tokenRes = await fetch(tokenUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: body.toString(),
      });

      if (!tokenRes.ok) {
        const errorData = await tokenRes.json().catch(() => ({}));
        return res.json({
          success: false,
          message: `Authentication failed: ${(errorData as any).error_description || (errorData as any).error || 'Unknown error'}`,
        });
      }

      const tokenData = await tokenRes.json() as { access_token: string };

      const graphRes = await fetch('https://graph.microsoft.com/v1.0/organization', {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      });

      if (graphRes.ok) {
        const orgData = await graphRes.json() as { value: Array<{ displayName: string }> };
        const orgName = orgData.value?.[0]?.displayName || 'Unknown';
        return res.json({
          success: true,
          message: `Connected successfully to "${orgName}" tenant.`,
        });
      } else {
        return res.json({
          success: true,
          message: 'Authentication successful, but could not read organization details. Check API permissions.',
        });
      }
    } catch (err: any) {
      return res.json({
        success: false,
        message: `Connection error: ${err.message || 'Network error'}`,
      });
    }
  });

  // ===================== DISCOVERY ROUTES =====================

  app.get('/api/projects/:id/discover/:type', requireAuth, async (req, res) => {
    try {
      const projectId = parseInt(req.params['id'] as string);
      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const missingCreds: string[] = [];
      if (!project.sourceTenantId) missingCreds.push('Source Tenant ID');
      if (!project.sourceClientId) missingCreds.push('Source Client ID');
      if (!project.sourceClientSecret) missingCreds.push('Source Client Secret');
      if (missingCreds.length) return res.status(400).json({ message: `Missing source credentials: ${missingCreds.join(', ')}` });

      const source = new GraphClient(project.sourceTenantId!, project.sourceClientId!, project.sourceClientSecret!);
      const type = req.params['type'] as string;

      let data: any;
      switch (type) {
        case 'users': data = await discoverUsers(source); break;
        case 'onedrive': data = await discoverOneDrives(source); break;
        case 'sharepoint': data = await discoverSharePointSites(source); break;
        case 'teams': data = await discoverTeams(source); break;
        case 'powerplatform': data = await discoverPowerPlatform(source); break;
        case 'distributiongroups': data = await discoverDistributionGroups(source); break;
        case 'sharedmailboxes': data = await discoverSharedMailboxes(source); break;
        case 'm365groups': data = await discoverM365Groups(source); break;
        default: return res.status(400).json({ message: `Unknown discovery type: ${type}` });
      }

      res.json({ type, data });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Discovery failed' });
    }
  });

  // ===================== MAPPING RULES ROUTES =====================

  app.get('/api/projects/:id/mapping-rules', requireAuth, async (req, res) => {
    try {
      const rules = await storage.getMappingRules(parseInt(req.params['id'] as string));
      res.json(rules);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post('/api/projects/:id/mapping-rules', requireAuth, async (req, res) => {
    try {
      const { ruleType, sourcePattern, targetPattern, description } = req.body;
      if (!ruleType || !sourcePattern || !targetPattern) {
        return res.status(400).json({ message: 'ruleType, sourcePattern, and targetPattern are required' });
      }
      const rule = await storage.createMappingRule({
        projectId: parseInt(req.params['id'] as string),
        ruleType,
        sourcePattern,
        targetPattern,
        description: description || null,
      });
      res.status(201).json(rule);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.delete('/api/projects/:id/mapping-rules/:ruleId', requireAuth, async (req, res) => {
    try {
      await storage.deleteMappingRule(parseInt(req.params['ruleId'] as string));
      res.json({ success: true });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  app.post('/api/projects/:id/apply-mapping', requireAuth, async (req, res) => {
    try {
      const { identities } = req.body as { identities: string[] };
      if (!Array.isArray(identities)) return res.status(400).json({ message: 'identities must be an array' });
      const pid = parseInt(req.params['id'] as string);
      const results = await Promise.all(
        identities.map(async (identity) => ({
          source: identity,
          target: await storage.applyMappingRules(pid, identity),
        }))
      );
      res.json(results);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  // ===================== ENTRA → AD ROUTES =====================

  // Save AD connection settings for a project
  app.post('/api/projects/:id/ad-settings', requireAuth, async (req, res) => {
    try {
      const projectId = parseInt(req.params['id'] as string);
      const { adDcHostname, adLdapPort, adBindDn, adBindPassword, adBaseDn, adUseSsl, adTargetOu } = req.body;
      const updated = await storage.updateProjectInternal(projectId, {
        adDcHostname: adDcHostname || null,
        adLdapPort: adLdapPort ? parseInt(adLdapPort) : 389,
        adBindDn: adBindDn || null,
        adBindPassword: adBindPassword || null,
        adBaseDn: adBaseDn || null,
        adUseSsl: !!adUseSsl,
        adTargetOu: adTargetOu || null,
      });
      res.json({
        ...updated,
        adBindPassword: updated.adBindPassword ? '••••••••' : null,
      });
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  // Test AD connection
  app.post('/api/projects/:id/ad-test-connection', requireAuth, async (req, res) => {
    try {
      const projectId = parseInt(req.params['id'] as string);
      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const { adDcHostname, adLdapPort, adBindDn, adBindPassword, adBaseDn, adUseSsl } = req.body;
      const dcHostname = adDcHostname || project.adDcHostname;
      const ldapPort = adLdapPort ? parseInt(adLdapPort) : (project.adLdapPort || 389);
      const bindDn = adBindDn || project.adBindDn;
      const bindPassword = adBindPassword || project.adBindPassword;
      const baseDn = adBaseDn || project.adBaseDn;
      const useSsl = adUseSsl !== undefined ? !!adUseSsl : (project.adUseSsl || false);

      if (!dcHostname || !bindDn || !bindPassword || !baseDn) {
        return res.status(400).json({ message: 'DC hostname, Bind DN, password, and Base DN are required to test the connection.' });
      }

      const result = await testAdConnection({ dcHostname, ldapPort, bindDn, bindPassword, baseDn, useSsl });
      res.json(result);
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Connection test failed' });
    }
  });

  // Discover cloud-only Entra users
  app.get('/api/projects/:id/entra-ad/discover', requireAuth, async (req, res) => {
    try {
      const projectId = parseInt(req.params['id'] as string);
      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      if (!project.sourceTenantId || !project.sourceClientId || !project.sourceClientSecret) {
        return res.status(400).json({ message: 'Source tenant credentials are required. Configure them in the Tenant Configuration tab.' });
      }

      const client = new GraphClient(project.sourceTenantId, project.sourceClientId, project.sourceClientSecret);
      const users = await discoverCloudOnlyUsers(client);
      res.json({ users });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Discovery failed' });
    }
  });

  // Migrate selected users to on-premises AD
  app.post('/api/projects/:id/entra-ad/migrate', requireAuth, async (req, res) => {
    try {
      const projectId = parseInt(req.params['id'] as string);
      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const { users } = req.body as { users: Array<{ upn: string; targetUpn: string; displayName: string; givenName?: string; surname?: string; jobTitle?: string; department?: string; officeLocation?: string; mobilePhone?: string; mail?: string }> };
      if (!Array.isArray(users) || users.length === 0) {
        return res.status(400).json({ message: 'No users provided for migration' });
      }

      if (!project.adDcHostname || !project.adBindDn || !project.adBindPassword || !project.adBaseDn) {
        return res.status(400).json({ message: 'Active Directory connection settings are not configured for this project.' });
      }

      const config: AdConnectionConfig = {
        dcHostname: project.adDcHostname,
        ldapPort: project.adLdapPort || 389,
        bindDn: project.adBindDn,
        bindPassword: project.adBindPassword,
        baseDn: project.adBaseDn,
        useSsl: project.adUseSsl || false,
        targetOu: project.adTargetOu,
      };

      const results = [];
      for (const u of users) {
        const entraUser = {
          id: u.upn,
          displayName: u.displayName,
          userPrincipalName: u.upn,
          mail: u.mail || null,
          givenName: u.givenName || null,
          surname: u.surname || null,
          jobTitle: u.jobTitle || null,
          department: u.department || null,
          officeLocation: u.officeLocation || null,
          mobilePhone: u.mobilePhone || null,
          accountEnabled: true,
          usageLocation: null,
        };
        const result = await migrateUserToAd(config, entraUser, u.targetUpn || u.upn);
        results.push(result);

        // Create migration item record
        await storage.createItem({
          projectId,
          sourceIdentity: u.upn,
          targetIdentity: u.targetUpn || u.upn,
          itemType: 'entra_to_ad',
          status: result.success ? 'completed' : 'failed',
          errorDetails: result.success ? null : result.message,
          logs: [result.message, ...(result.tempPassword ? [`Temporary password: ${result.tempPassword}`] : [])],
        } as any);
      }

      res.json({ results });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Migration failed' });
    }
  });

  // Generate PowerShell export script
  app.post('/api/projects/:id/entra-ad/export-ps', requireAuth, async (req, res) => {
    try {
      const projectId = parseInt(req.params['id'] as string);
      const project = await storage.getProjectInternal(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const { users } = req.body as { users: Array<{ upn: string; targetUpn: string; displayName: string; givenName?: string; surname?: string; jobTitle?: string; department?: string; officeLocation?: string; mobilePhone?: string; mail?: string }> };
      if (!Array.isArray(users) || users.length === 0) {
        return res.status(400).json({ message: 'No users provided' });
      }

      const entraUsers = users.map(u => ({
        id: u.upn,
        displayName: u.displayName,
        userPrincipalName: u.upn,
        mail: u.mail || null,
        givenName: u.givenName || null,
        surname: u.surname || null,
        jobTitle: u.jobTitle || null,
        department: u.department || null,
        officeLocation: u.officeLocation || null,
        mobilePhone: u.mobilePhone || null,
        accountEnabled: true,
        usageLocation: null,
      }));

      const targetUpns = users.map(u => u.targetUpn || u.upn);
      const script = generatePowerShellScript(entraUsers, targetUpns, {
        baseDn: project.adBaseDn || 'DC=corp,DC=com',
        targetOu: project.adTargetOu,
      });

      res.setHeader('Content-Type', 'text/plain');
      res.setHeader('Content-Disposition', `attachment; filename="migrate-entra-to-ad-${Date.now()}.ps1"`);
      res.send(script);
    } catch (err: any) {
      res.status(500).json({ message: err.message });
    }
  });

  // ── OAuth2 tenant connection (auth code + PKCE) ───────────────────────────

  // Step 1: Redirect browser to Microsoft login
  app.get('/api/oauth/connect', requireAuth, (req, res) => {
    try {
      const { projectId, tenantType, tenantId, appName } = req.query as Record<string, string>;
      if (!projectId || !tenantType || !tenantId) {
        return res.status(400).send('Missing projectId, tenantType, or tenantId');
      }
      const url = buildConnectUrl(
        tenantId,
        parseInt(projectId),
        tenantType as 'source' | 'target',
        appName || 'Tenant Migration Tool',
      );
      res.redirect(url);
    } catch (err: any) {
      res.redirect(`/?oauth_error=${encodeURIComponent(err.message)}`);
    }
  });

  // Step 2: Microsoft redirects back to http://localhost:5000 (root) with code.
  // We intercept at GET / before Vite serves the React app.
  // The bare http://localhost URI is registered on the Microsoft public client
  // and Azure AD accepts any port per RFC 8252.
  app.get('/', async (req, res, next) => {
    const { code, state, error, error_description } = req.query as Record<string, string>;

    // Not an OAuth callback — let Vite serve the React SPA
    if (!code && !error) return next();

    if (error) {
      const msg = error_description || error || 'Authentication failed';
      return res.redirect(`/?oauth_error=${encodeURIComponent(msg)}`);
    }
    if (!code || !state) {
      return res.redirect('/?oauth_error=Missing+code+or+state');
    }

    try {
      const result = await handleOAuthCallback(code, state);
      const updates = result.tenantType === 'source'
        ? { sourceTenantId: result.tenantId, sourceClientId: result.clientId, sourceClientSecret: result.clientSecret }
        : { targetTenantId: result.tenantId, targetClientId: result.clientId, targetClientSecret: result.clientSecret };
      await storage.updateProjectInternal(result.projectId, updates);
      res.redirect(`/projects/${result.projectId}?oauth_success=${result.tenantType}&app=${encodeURIComponent(result.displayName)}`);
    } catch (err: any) {
      console.error('[OAuth callback error]', err.message);
      // Try to extract projectId from the state so we can redirect back to the project page
      // where the error toast will be visible — otherwise the user lands on the dashboard silently
      let projectId: number | null = null;
      try {
        const dot = state?.lastIndexOf('.');
        if (dot > 0) {
          const decoded = JSON.parse(Buffer.from(state.slice(0, dot), 'base64url').toString());
          if (decoded?.projectId) projectId = decoded.projectId;
        }
      } catch { /* ignore */ }
      const errParam = `oauth_error=${encodeURIComponent(err.message)}`;
      res.redirect(projectId ? `/projects/${projectId}?${errParam}` : `/?${errParam}`);
    }
  });

  // Consent callback — close popup and notify parent window
  app.get('/oauth/consent-complete', async (req, res) => {
    const { error, error_description, admin_consent, state } = req.query as Record<string, string>;
    if (error) {
      const msg = error_description || error;
      if (msg?.includes('500113') || msg?.includes('reply address')) {
        const fixMsg = 'The app registration is missing a reply URL. Please click "Connect with Microsoft" again to re-register this tenant, then retry granting permissions.';
        res.send(`<html><body style="font-family:sans-serif;padding:2rem"><script>window.opener?.postMessage({type:'consent_error',error:${JSON.stringify(fixMsg)}},location.origin);window.close();</script><h3>Setup Required</h3><p>${fixMsg}</p></body></html>`);
      } else {
        res.send(`<html><body style="font-family:sans-serif;padding:2rem"><script>window.opener?.postMessage({type:'consent_error',error:${JSON.stringify(msg)}},location.origin);window.close();</script><h3>Error granting consent</h3><p>${msg}</p></body></html>`);
      }
      return;
    }

    // Persist consent to DB when admin_consent=True and we have a valid state token
    if ((admin_consent === 'True' || admin_consent === 'true') && state) {
      try {
        const cs = decodeConsentState(state);
        if (cs) {
          const project = await storage.getProjectInternal(cs.projectId);
          if (project) {
            const allServiceKeys = Object.keys(SERVICE_PERMISSION_GROUPS);
            const field = cs.tenantType === 'source' ? 'sourceConsentedServices' : 'targetConsentedServices';
            await storage.updateProjectInternal(cs.projectId, { [field]: JSON.stringify(allServiceKeys) });
          }
        }
      } catch { /* non-fatal — consent was still granted, just not persisted */ }
    }

    if (admin_consent === 'True' || admin_consent === 'true') {
      res.send(`<html><body style="font-family:sans-serif;padding:2rem"><script>window.opener?.postMessage({type:'consent_success'},location.origin);window.close();</script><h3>✓ Permissions granted</h3><p>Admin consent was successfully granted. You can close this window.</p></body></html>`);
    } else {
      res.send(`<html><body style="font-family:sans-serif;padding:2rem"><script>window.opener?.postMessage({type:'consent_success'},location.origin);window.close();</script><h3>Done</h3><p>You can close this window.</p></body></html>`);
    }
  });

  // Return service permission groups so the frontend can render grant buttons
  app.get('/api/oauth/services', requireAuth, (req, res) => {
    res.json(SERVICE_PERMISSION_GROUPS);
  });

  // Build consent URL for a specific service (includes projectId+tenantType in state for DB persistence)
  app.get('/api/oauth/consent-url', requireAuth, (req, res) => {
    const { tenantId, clientId, service, projectId, tenantType } = req.query as Record<string, string>;
    if (!tenantId || !clientId || !service) {
      return res.status(400).json({ message: 'tenantId, clientId, and service are required' });
    }
    if (!(service in SERVICE_PERMISSION_GROUPS)) {
      return res.status(400).json({ message: 'Invalid service key' });
    }
    const pid = projectId ? Number(projectId) : undefined;
    const tt = tenantType === 'source' || tenantType === 'target' ? tenantType : undefined;
    const url = buildConsentUrl(tenantId, clientId, service as ServiceKey, pid, tt);
    res.json({ url });
  });

  // Build a re-grant consent URL using the existing app registration (no new OAuth flow)
  app.get('/api/oauth/regrant-url', requireAuth, async (req, res) => {
    const { projectId, tenantType } = req.query as Record<string, string>;
    if (!projectId || !tenantType) {
      return res.status(400).json({ message: 'projectId and tenantType are required' });
    }
    if (tenantType !== 'source' && tenantType !== 'target') {
      return res.status(400).json({ message: 'tenantType must be source or target' });
    }
    const project = await storage.getProject(Number(projectId), (req as any).user?.id);
    if (!project) return res.status(404).json({ message: 'Project not found' });
    const tenantId = tenantType === 'source' ? project.sourceTenantId : project.targetTenantId;
    const clientId = tenantType === 'source' ? project.sourceClientId : project.targetClientId;
    if (!clientId) return res.status(400).json({ message: 'No app registration found for this tenant. Connect with Microsoft first.' });
    const url = buildRegrantConsentUrl(tenantId, clientId, Number(projectId), tenantType);
    res.json({ url });
  });

  // === Exchange Online PowerShell Routes ===

  // GET /api/projects/:id/exo-settings
  app.get('/api/projects/:id/exo-settings', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });
    const exo = (project.exoSettings as any) || {};
    // Never return passwords — mask them
    return res.json({
      sourceCertPath: exo.sourceCertPath || '',
      sourceCertPassword: exo.sourceCertPassword ? '••••••••' : '',
      sourceOrg: exo.sourceOrg || '',
      targetCertPath: exo.targetCertPath || '',
      targetCertPassword: exo.targetCertPassword ? '••••••••' : '',
      targetOrg: exo.targetOrg || '',
      autoDelegate: exo.autoDelegate !== false,
      configured: !!(exo.targetCertPath && exo.targetOrg),
    });
  });

  // PATCH /api/projects/:id/exo-settings
  app.patch('/api/projects/:id/exo-settings', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });

    const existing = (project.exoSettings as any) || {};
    const body = req.body as {
      sourceCertPath?: string;
      sourceCertPassword?: string;
      sourceOrg?: string;
      targetCertPath?: string;
      targetCertPassword?: string;
      targetOrg?: string;
      autoDelegate?: boolean;
    };

    const merged = {
      sourceCertPath: body.sourceCertPath ?? existing.sourceCertPath,
      // Keep old password if masked value sent back
      sourceCertPassword: (body.sourceCertPassword && !body.sourceCertPassword.includes('•'))
        ? body.sourceCertPassword : existing.sourceCertPassword,
      sourceOrg: body.sourceOrg ?? existing.sourceOrg,
      targetCertPath: body.targetCertPath ?? existing.targetCertPath,
      targetCertPassword: (body.targetCertPassword && !body.targetCertPassword.includes('•'))
        ? body.targetCertPassword : existing.targetCertPassword,
      targetOrg: body.targetOrg ?? existing.targetOrg,
      autoDelegate: body.autoDelegate ?? existing.autoDelegate ?? true,
    };

    await storage.updateProject(project.id, { exoSettings: merged } as any, userId);
    return res.json({ ok: true });
  });

  // POST /api/projects/:id/exo-test — test EXO connection
  app.post('/api/projects/:id/exo-test', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });

    const { tenant } = req.body as { tenant: 'source' | 'target' };
    const exo = (project.exoSettings as any) || {};
    const certPath = tenant === 'source' ? exo.sourceCertPath : exo.targetCertPath;
    const certPassword = tenant === 'source' ? exo.sourceCertPassword : exo.targetCertPassword;
    const org = tenant === 'source' ? exo.sourceOrg : exo.targetOrg;
    const clientId = tenant === 'source' ? project.sourceClientId : project.targetClientId;

    if (!certPath || !org || !clientId) {
      return res.status(400).json({ success: false, message: 'Certificate path, organization domain, and app client ID are required.' });
    }

    const { testExoConnection } = await import('./services/exo-runner');
    const result = await testExoConnection({ clientId, certPath, certPassword: certPassword || '', organization: org });
    return res.json({ success: result.success, output: result.output, errors: result.errors });
  });

  // POST /api/projects/:id/exo-install-module — install ExchangeOnlineManagement module
  app.post('/api/projects/:id/exo-install-module', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });
    const { ensureExoModuleInstalled } = await import('./services/exo-runner');
    const result = await ensureExoModuleInstalled();
    return res.json(result);
  });

  // === Continuous Sync Routes ===

  // GET /api/projects/:id/sync-status  — returns sync settings + per-item sync state
  app.get('/api/projects/:id/sync-status', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });
    const items = await storage.getItems(project.id);
    const syncableTypes = ['mailbox', 'onedrive', 'sharepoint', 'sharedmailbox'];
    const syncableItems = items.filter(i => i.status === 'completed' && syncableTypes.includes(i.itemType));
    const needsActionItems = items.filter(i => i.status === 'needs_action' && syncableTypes.includes(i.itemType));
    return res.json({
      syncEnabled: project.syncEnabled ?? false,
      syncIntervalMinutes: project.syncIntervalMinutes ?? 60,
      completedItems: syncableItems.length,
      needsActionCount: needsActionItems.length,
      items: syncableItems.map(i => ({
        id: i.id,
        itemType: i.itemType,
        sourceIdentity: i.sourceIdentity,
        targetIdentity: i.targetIdentity,
        lastSyncedAt: i.lastSyncedAt,
        nextSyncAt: i.nextSyncAt,
      })),
    });
  });

  // PATCH /api/projects/:id/sync-settings — enable/disable sync and set interval
  app.patch('/api/projects/:id/sync-settings', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });
    const { syncEnabled, syncIntervalMinutes } = req.body as {
      syncEnabled?: boolean;
      syncIntervalMinutes?: number;
    };
    const updates: Record<string, any> = {};
    if (syncEnabled !== undefined) updates.syncEnabled = syncEnabled;
    if (syncIntervalMinutes !== undefined) updates.syncIntervalMinutes = syncIntervalMinutes;
    const updated = await storage.updateProject(project.id, updates, userId);
    return res.json({ syncEnabled: updated.syncEnabled, syncIntervalMinutes: updated.syncIntervalMinutes });
  });

  // POST /api/projects/:id/sync-now — trigger an immediate sync for this project
  app.post('/api/projects/:id/sync-now', requireAuth, async (req, res) => {
    const userId = getSessionUserId(req);
    const project = await storage.getProject(Number(req.params.id), userId);
    if (!project) return res.status(404).json({ message: 'Project not found' });
    // Import here to avoid circular deps at top-level
    const { runProjectSync } = await import('./services/sync-engine');
    // Run async — return immediately
    runProjectSync(project.id).catch(e => console.error('[sync-now] error:', e.message));
    return res.json({ message: 'Sync started — check item logs for progress' });
  });

  return httpServer;
}
