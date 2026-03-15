// Uses Microsoft Graph Command Line Tools (public client) for device code flow.
// This lets an admin sign in and we use their delegated token to create an app registration.

const GRAPH_CLI_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const MICROSOFT_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

// All Graph API application permissions the migration tool needs
const REQUIRED_PERMISSIONS = [
  { id: 'df021288-bdef-4463-88db-98f22de89214', name: 'User.Read.All' },
  { id: 'e2a3a72e-5f79-4c64-b1b1-878b674786c9', name: 'Mail.ReadWrite' },
  { id: '6931b611-5b96-4af8-87a9-f151f27e3be1', name: 'MailboxSettings.ReadWrite' },
  { id: '75359482-378d-4052-8f01-80520e7db3cd', name: 'Files.ReadWrite.All' },
  { id: '9492366f-7969-46a4-8d15-ed1a20078fff', name: 'Sites.ReadWrite.All' },
  { id: '62a82d76-70ea-41e2-9197-370581804d09', name: 'Group.ReadWrite.All' },
  { id: 'bdd80a03-d9bc-451d-b7c4-ce7c63fe3c8f', name: 'TeamSettings.ReadWrite.All' },
  { id: '243cded2-bd16-4fd6-a953-ff8177894c3d', name: 'Channel.ReadWrite.All' },
  { id: 'cc7e7635-2586-41d6-adaa-a8d3bcad5ee5', name: 'TeamMember.ReadWrite.All' },
  { id: '5b567255-7703-4780-807c-7be8301ae99b', name: 'Group.Read.All' },
  { id: '741f803b-c850-494e-b5df-cde7c675a1ca', name: 'User.ReadWrite.All' },
];

interface PendingFlow {
  deviceCode: string;
  tenantId: string;
  interval: number;
  expiresAt: number;
}

const pendingFlows = new Map<string, PendingFlow>();

// ── Device code flow ──────────────────────────────────────────────────────────

export async function startDeviceCodeFlow(tenantId: string): Promise<{
  requestId: string;
  userCode: string;
  verificationUri: string;
  expiresIn: number;
  message: string;
}> {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`;
  const body = new URLSearchParams({
    client_id: GRAPH_CLI_CLIENT_ID,
    scope: [
      'https://graph.microsoft.com/Application.ReadWrite.All',
      'https://graph.microsoft.com/Directory.ReadWrite.All',
      'https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All',
      'offline_access',
    ].join(' '),
  });

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  const data = await res.json() as any;

  if (!res.ok) {
    throw new Error(data.error_description || data.error || `Failed to start sign-in flow (${res.status})`);
  }

  const requestId = `flow_${Date.now()}_${Math.random().toString(36).slice(-6)}`;
  pendingFlows.set(requestId, {
    deviceCode: data.device_code,
    tenantId,
    interval: data.interval || 5,
    expiresAt: Date.now() + (data.expires_in ?? 900) * 1000,
  });

  // Clean up expired flows after 20 minutes
  setTimeout(() => pendingFlows.delete(requestId), 20 * 60 * 1000);

  return {
    requestId,
    userCode: data.user_code,
    verificationUri: data.verification_uri || 'https://microsoft.com/devicelogin',
    expiresIn: data.expires_in ?? 900,
    message: data.message || `Go to ${data.verification_uri} and enter the code.`,
  };
}

// ── Token polling ─────────────────────────────────────────────────────────────

async function pollForToken(flow: PendingFlow): Promise<{ access_token: string } | null | 'expired' | 'declined'> {
  if (Date.now() > flow.expiresAt) return 'expired';

  const url = `https://login.microsoftonline.com/${flow.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
    client_id: GRAPH_CLI_CLIENT_ID,
    device_code: flow.deviceCode,
  });

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  const data = await res.json() as any;

  if (data.error === 'authorization_pending' || data.error === 'slow_down') return null;
  if (data.error === 'authorization_declined') return 'declined';
  if (data.error === 'expired_token') return 'expired';
  if (data.access_token) return { access_token: data.access_token };
  return null;
}

// ── Graph API helper ──────────────────────────────────────────────────────────

async function graph(token: string, method: string, path: string, body?: any) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });

  const text = await res.text();
  const json = text ? JSON.parse(text) : {};

  if (!res.ok) {
    const msg = json.error?.message || json.message || `Graph error ${res.status}`;
    throw new Error(msg);
  }
  return json;
}

// ── Poll + create app ─────────────────────────────────────────────────────────

export interface AppCreationResult {
  status: 'pending' | 'completed' | 'expired' | 'declined' | 'failed';
  clientId?: string;
  clientSecret?: string;
  tenantId?: string;
  appId?: string; // object ID of the application
  displayName?: string;
  consentUrl?: string;
  consentGranted?: boolean;
  permissions?: string[];
  error?: string;
}

export async function pollAndCreateApp(requestId: string, appName: string): Promise<AppCreationResult> {
  const flow = pendingFlows.get(requestId);
  if (!flow) return { status: 'expired' };

  const tokenResult = await pollForToken(flow);

  if (tokenResult === 'expired') {
    pendingFlows.delete(requestId);
    return { status: 'expired' };
  }
  if (tokenResult === 'declined') {
    pendingFlows.delete(requestId);
    return { status: 'declined', error: 'The sign-in was cancelled or declined.' };
  }
  if (tokenResult === null) {
    return { status: 'pending' };
  }

  // Token acquired — remove flow and proceed
  pendingFlows.delete(requestId);
  const token = tokenResult.access_token;

  try {
    // 1. Create the application registration
    const app = await graph(token, 'POST', '/applications', {
      displayName: appName,
      signInAudience: 'AzureADMyOrg',
      requiredResourceAccess: [{
        resourceAppId: MICROSOFT_GRAPH_APP_ID,
        resourceAccess: REQUIRED_PERMISSIONS.map(p => ({ id: p.id, type: 'Role' })),
      }],
    });

    // 2. Add a client secret (2-year expiry)
    const secretResp = await graph(token, 'POST', `/applications/${app.id}/addPassword`, {
      passwordCredential: {
        displayName: 'Tenant Migration Tool',
        endDateTime: new Date(Date.now() + 2 * 365 * 24 * 60 * 60 * 1000).toISOString(),
      },
    });

    // 3. Create a service principal (needed for admin consent)
    let spId: string | null = null;
    try {
      const sp = await graph(token, 'POST', '/servicePrincipals', { appId: app.appId });
      spId = sp.id;
    } catch {
      // May already exist — look it up
      try {
        const existing = await graph(token, 'GET', `/servicePrincipals?$filter=appId eq '${app.appId}'&$select=id`);
        spId = existing.value?.[0]?.id ?? null;
      } catch { /* ignore */ }
    }

    // 4. Try to grant admin consent automatically via appRoleAssignments
    let consentGranted = false;
    if (spId) {
      try {
        const graphSpResp = await graph(token, 'GET', `/servicePrincipals?$filter=appId eq '${MICROSOFT_GRAPH_APP_ID}'&$select=id`);
        const graphSpId = graphSpResp.value?.[0]?.id;

        if (graphSpId) {
          let allOk = true;
          for (const perm of REQUIRED_PERMISSIONS) {
            try {
              await graph(token, 'POST', `/servicePrincipals/${graphSpId}/appRoleAssignedTo`, {
                principalId: spId,
                resourceId: graphSpId,
                appRoleId: perm.id,
              });
            } catch { allOk = false; }
          }
          consentGranted = allOk;
        }
      } catch { /* non-fatal — fall back to consent URL */ }
    }

    const consentUrl = `https://login.microsoftonline.com/${flow.tenantId}/adminconsent?client_id=${app.appId}&redirect_uri=${encodeURIComponent('https://portal.azure.com')}`;

    return {
      status: 'completed',
      clientId: app.appId,
      clientSecret: secretResp.secretText,
      tenantId: flow.tenantId,
      appId: app.id,
      displayName: app.displayName,
      consentUrl,
      consentGranted,
      permissions: REQUIRED_PERMISSIONS.map(p => p.name),
    };
  } catch (err: any) {
    return { status: 'failed', error: err.message };
  }
}

export function getFlowStatus(requestId: string): 'active' | 'not_found' {
  return pendingFlows.has(requestId) ? 'active' : 'not_found';
}
