import type { Express } from "express";
import { createServer, type Server } from "http";
import { storage } from "./storage";
import { setupSession, registerAuthRoutes, isAuthenticated as requireAuth } from "./auth";
import { api } from "@shared/routes";
import { z } from "zod";
import type { Project } from "@shared/schema";
import { migrateItem, migrateAllPending } from "./services/migration-engine";
import { GraphClient } from "./services/graph-client";
import { discoverUsers, discoverSharePointSites, discoverTeams, discoverPowerPlatform } from "./services/discovery-service";
import { discoverCloudOnlyUsers, testAdConnection, migrateUserToAd, generatePowerShellScript, type AdConnectionConfig } from "./services/entra-ad-service";
import { buildConnectUrl, handleOAuthCallback, buildConsentUrl, SERVICE_PERMISSION_GROUPS, type ServiceKey } from "./services/oauth-tenant-service";

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

  // === Projects ===
  app.get(api.projects.list.path, async (req, res) => {
    const projects = await storage.getProjects();
    res.json(projects.map(sanitizeProject));
  });

  app.get(api.projects.get.path, async (req, res) => {
    const project = await storage.getProject(Number(req.params.id));
    if (!project) {
      return res.status(404).json({ message: 'Project not found' });
    }
    res.json(sanitizeProject(project));
  });

  app.post(api.projects.create.path, async (req, res) => {
    try {
      const input = api.projects.create.input.parse(req.body);
      // Optional: attach user ID from auth
      // const userId = req.user?.claims?.sub;
      // if (userId) input.userId = userId;
      
      const project = await storage.createProject(input);
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

  app.patch(api.projects.update.path, async (req, res) => {
    try {
      const input = api.projects.update.input.parse(req.body);
      const project = await storage.updateProject(Number(req.params.id), input);
      if (!project) return res.status(404).json({ message: 'Project not found' });
      res.json(sanitizeProject(project));
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

  // === Migration Endpoints ===
  app.post('/api/projects/:projectId/items/:itemId/migrate', async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);
      const itemId = Number(req.params.itemId);

      const project = await storage.getProject(projectId);
      if (!project) return res.status(404).json({ message: 'Project not found' });

      const item = await storage.getItem(itemId);
      if (!item || item.projectId !== projectId) return res.status(404).json({ message: 'Item not found' });

      if (item.status === 'in_progress') {
        return res.status(409).json({ message: 'Migration already in progress for this item' });
      }

      migrateItem(projectId, itemId).catch(err => {
        console.error(`Background migration failed for item ${itemId}:`, err.message);
      });

      res.json({ message: 'Migration started', itemId });
    } catch (err: any) {
      res.status(500).json({ message: err.message || 'Internal server error' });
    }
  });

  app.post('/api/projects/:projectId/migrate-all', async (req, res) => {
    try {
      const projectId = Number(req.params.projectId);

      const project = await storage.getProject(projectId);
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

  app.get('/api/items/:id/logs', async (req, res) => {
    const item = await storage.getItem(Number(req.params.id));
    if (!item) return res.status(404).json({ message: 'Item not found' });
    res.json({ logs: item.logs || [] });
  });

  // === Test Connection ===
  app.post('/api/projects/:id/test-connection', async (req, res) => {
    const project = await storage.getProject(Number(req.params.id));
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
      const project = await storage.getProject(projectId);
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
        case 'sharepoint': data = await discoverSharePointSites(source); break;
        case 'teams': data = await discoverTeams(source); break;
        case 'powerplatform': data = await discoverPowerPlatform(source); break;
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
      const updated = await storage.updateProject(projectId, {
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
      const project = await storage.getProject(projectId);
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
      const project = await storage.getProject(projectId);
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
      const project = await storage.getProject(projectId);
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
      const project = await storage.getProject(projectId);
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

  // Step 2: Microsoft redirects back here with code
  // Note: no /api prefix — must match registered redirect URI exactly
  app.get('/oauth/callback', async (req, res) => {
    const { code, state, error, error_description } = req.query as Record<string, string>;

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
        ? { sourceClientId: result.clientId, sourceClientSecret: result.clientSecret }
        : { targetClientId: result.clientId, targetClientSecret: result.clientSecret };
      await storage.updateProject(result.projectId, updates);
      res.redirect(`/projects/${result.projectId}?oauth_success=${result.tenantType}&app=${encodeURIComponent(result.displayName)}`);
    } catch (err: any) {
      res.redirect(`/?oauth_error=${encodeURIComponent(err.message)}`);
    }
  });

  // Consent callback — just close/redirect back
  app.get('/oauth/consent-complete', (req, res) => {
    const { error, error_description } = req.query as Record<string, string>;
    if (error) {
      res.send(`<html><body><script>window.opener?.postMessage({type:'consent_error',error:${JSON.stringify(error_description||error)}},location.origin);window.close();</script><p>Error: ${error_description||error}</p></body></html>`);
    } else {
      res.send(`<html><body><script>window.opener?.postMessage({type:'consent_success'},location.origin);window.close();</script><p>Permissions granted. You can close this window.</p></body></html>`);
    }
  });

  // Return service permission groups so the frontend can render grant buttons
  app.get('/api/oauth/services', requireAuth, (req, res) => {
    res.json(SERVICE_PERMISSION_GROUPS);
  });

  // Build consent URL for a specific service
  app.get('/api/oauth/consent-url', requireAuth, (req, res) => {
    const { tenantId, clientId, service } = req.query as Record<string, string>;
    if (!tenantId || !clientId || !service) {
      return res.status(400).json({ message: 'tenantId, clientId, and service are required' });
    }
    if (!(service in SERVICE_PERMISSION_GROUPS)) {
      return res.status(400).json({ message: 'Invalid service key' });
    }
    const url = buildConsentUrl(tenantId, clientId, service as ServiceKey);
    res.json({ url });
  });

  return httpServer;
}
