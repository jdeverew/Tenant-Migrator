// OAuth2 authorization code flow with PKCE for tenant connection.
// Uses Microsoft Graph Command Line Tools as a public client (no secret needed).
// Microsoft allows http://localhost redirect URIs for public clients per RFC 8252.

import { createHash, randomBytes } from 'crypto';

const PUBLIC_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Command Line Tools

// Microsoft allows any port for the bare http://localhost URI (RFC 8252).
// We use the root path so we don't need to register a custom path on Microsoft's public client app.
const REDIRECT_URI = 'http://localhost:5000';
const CONSENT_REDIRECT_URI = 'http://localhost:5000/oauth/consent-complete';

// ── Per-service permission groups (application permissions) ─────────────────
export const SERVICE_PERMISSION_GROUPS = {
  exchange: {
    label: 'Exchange Online',
    description: 'Required for mailbox and email migration',
    icon: 'Mail',
    scopes: [
      'https://graph.microsoft.com/Mail.ReadWrite',
      'https://graph.microsoft.com/MailboxSettings.ReadWrite',
      'https://graph.microsoft.com/Calendars.ReadWrite',
    ],
    appRoleIds: [
      'e2a3a72e-5f79-4c64-b1b1-878b674786c9', // Mail.ReadWrite
      '6931b611-5b96-4af8-87a9-f151f27e3be1', // MailboxSettings.ReadWrite
      '6e98f277-4fea-4a57-a96d-153778c26628', // Calendars.ReadWrite
    ],
  },
  sharepoint: {
    label: 'SharePoint & OneDrive',
    description: 'Required for SharePoint sites and OneDrive file migration',
    icon: 'Cloud',
    scopes: [
      'https://graph.microsoft.com/Sites.ReadWrite.All',
      'https://graph.microsoft.com/Files.ReadWrite.All',
    ],
    appRoleIds: [
      '9492366f-7969-46a4-8d15-ed1a20078fff', // Sites.ReadWrite.All
      '75359482-378d-4052-8f01-80520e7db3cd', // Files.ReadWrite.All
    ],
  },
  teams: {
    label: 'Microsoft Teams',
    description: 'Required for Teams and channel migration',
    icon: 'Users',
    scopes: [
      'https://graph.microsoft.com/Team.ReadWrite.All',
      'https://graph.microsoft.com/Channel.ReadWrite.All',
      'https://graph.microsoft.com/TeamMember.ReadWrite.All',
      'https://graph.microsoft.com/Group.ReadWrite.All',
    ],
    appRoleIds: [
      'bdd80a03-d9bc-451d-b7c4-ce7c63fe3c8f', // TeamSettings.ReadWrite.All
      '243cded2-bd16-4fd6-a953-ff8177894c3d', // Channel.ReadWrite.All
      'cc7e7635-2586-41d6-adaa-a8d3bcad5ee5', // TeamMember.ReadWrite.All
      '62a82d76-70ea-41e2-9197-370581804d09', // Group.ReadWrite.All
    ],
  },
  users: {
    label: 'Users & Directory',
    description: 'Required for user account and Entra ID migrations',
    icon: 'UserCheck',
    scopes: [
      'https://graph.microsoft.com/User.ReadWrite.All',
      'https://graph.microsoft.com/Directory.ReadWrite.All',
    ],
    appRoleIds: [
      '741f803b-c850-494e-b5df-cde7c675a1ca', // User.ReadWrite.All
      '19dbc75e-c2e2-444c-a770-ec69d8559fc7', // Directory.ReadWrite.All
    ],
  },
};

export type ServiceKey = keyof typeof SERVICE_PERMISSION_GROUPS;

// ── PKCE helpers ─────────────────────────────────────────────────────────────

function generateCodeVerifier(): string {
  return randomBytes(32).toString('base64url');
}

function generateCodeChallenge(verifier: string): string {
  return createHash('sha256').update(verifier).digest('base64url');
}

// ── In-memory state store ─────────────────────────────────────────────────────

interface PendingState {
  tenantId: string;
  projectId: number;
  tenantType: 'source' | 'target';
  codeVerifier: string;
  appName: string;
  createdAt: number;
}

const pendingStates = new Map<string, PendingState>();

// Clean up states older than 15 minutes
setInterval(() => {
  const cutoff = Date.now() - 15 * 60 * 1000;
  for (const [key, val] of pendingStates) {
    if (val.createdAt < cutoff) pendingStates.delete(key);
  }
}, 5 * 60 * 1000);

// ── Start OAuth2 flow ─────────────────────────────────────────────────────────

export function buildConnectUrl(
  tenantId: string,
  projectId: number,
  tenantType: 'source' | 'target',
  appName: string,
): string {
  const state = randomBytes(16).toString('base64url');
  const codeVerifier = generateCodeVerifier();
  const codeChallenge = generateCodeChallenge(codeVerifier);

  pendingStates.set(state, {
    tenantId,
    projectId,
    tenantType,
    codeVerifier,
    appName,
    createdAt: Date.now(),
  });

  const params = new URLSearchParams({
    client_id: PUBLIC_CLIENT_ID,
    response_type: 'code',
    redirect_uri: REDIRECT_URI,
    scope: 'openid offline_access https://graph.microsoft.com/Application.ReadWrite.All https://graph.microsoft.com/Directory.ReadWrite.All https://graph.microsoft.com/AppRoleAssignment.ReadWrite.All',
    state,
    code_challenge: codeChallenge,
    code_challenge_method: 'S256',
    prompt: 'select_account',
  });

  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?${params}`;
}

// ── Handle OAuth callback ─────────────────────────────────────────────────────

const MICROSOFT_GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000';

async function graphPost(token: string, path: string, body: any) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  const json = await res.json() as any;
  if (!res.ok) throw new Error(json.error?.message || `Graph error ${res.status}`);
  return json;
}

async function graphGet(token: string, path: string) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const json = await res.json() as any;
  if (!res.ok) throw new Error(json.error?.message || `Graph error ${res.status}`);
  return json;
}

export interface OAuthCallbackResult {
  projectId: number;
  tenantType: 'source' | 'target';
  clientId: string;
  clientSecret: string;
  tenantId: string;
  displayName: string;
  error?: string;
}

export async function handleOAuthCallback(code: string, state: string): Promise<OAuthCallbackResult> {
  const pending = pendingStates.get(state);
  if (!pending) throw new Error('Invalid or expired OAuth state. Please try connecting again.');
  pendingStates.delete(state);

  // 1. Exchange code for token
  const tokenRes = await fetch(`https://login.microsoftonline.com/${pending.tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'authorization_code',
      client_id: PUBLIC_CLIENT_ID,
      code,
      redirect_uri: REDIRECT_URI,
      code_verifier: pending.codeVerifier,
    }).toString(),
  });

  const tokenData = await tokenRes.json() as any;
  if (!tokenRes.ok || !tokenData.access_token) {
    throw new Error(tokenData.error_description || tokenData.error || 'Token exchange failed');
  }

  const token = tokenData.access_token;
  const allAppRoles = Object.values(SERVICE_PERMISSION_GROUPS).flatMap(g => g.appRoleIds);

  // 2. Create app registration with all required permissions and redirect URIs
  const app = await graphPost(token, '/applications', {
    displayName: pending.appName,
    signInAudience: 'AzureADMyOrg',
    web: {
      redirectUris: [CONSENT_REDIRECT_URI],
    },
    requiredResourceAccess: [{
      resourceAppId: MICROSOFT_GRAPH_APP_ID,
      resourceAccess: allAppRoles.map(id => ({ id, type: 'Role' })),
    }],
  });

  // 3. Create client secret
  const secretResp = await graphPost(token, `/applications/${app.id}/addPassword`, {
    passwordCredential: {
      displayName: 'Tenant Migration Tool',
      endDateTime: new Date(Date.now() + 2 * 365 * 24 * 60 * 60 * 1000).toISOString(),
    },
  });

  // 4. Create service principal
  let spId: string | null = null;
  try {
    const sp = await graphPost(token, '/servicePrincipals', { appId: app.appId });
    spId = sp.id;
  } catch {
    try {
      const existing = await graphGet(token, `/servicePrincipals?$filter=appId eq '${app.appId}'&$select=id`);
      spId = existing.value?.[0]?.id ?? null;
    } catch { /* ignore */ }
  }

  // 5. Attempt auto-consent for all permissions
  if (spId) {
    try {
      const graphSpResp = await graphGet(token, `/servicePrincipals?$filter=appId eq '${MICROSOFT_GRAPH_APP_ID}'&$select=id`);
      const graphSpId = graphSpResp.value?.[0]?.id;
      if (graphSpId) {
        for (const roleId of allAppRoles) {
          await graphPost(token, `/servicePrincipals/${graphSpId}/appRoleAssignedTo`, {
            principalId: spId,
            resourceId: graphSpId,
            appRoleId: roleId,
          }).catch(() => {});
        }
      }
    } catch { /* non-fatal */ }
  }

  return {
    projectId: pending.projectId,
    tenantType: pending.tenantType,
    clientId: app.appId,
    clientSecret: secretResp.secretText,
    tenantId: pending.tenantId,
    displayName: app.displayName,
  };
}

// ── Build per-service admin consent URL ──────────────────────────────────────

export function buildConsentUrl(tenantId: string, clientId: string, _service: ServiceKey): string {
  // Use the v1 adminconsent endpoint with NO redirect_uri and NO scope.
  // The v2.0/adminconsent endpoint requires redirect_uri to be pre-registered on the app,
  // which causes AADSTS5000224 for any app the user didn't create through our OAuth flow.
  // The v1 endpoint only needs client_id — Azure shows a built-in success/error page,
  // and our UI marks the service as granted optimistically after the popup is opened.
  const params = new URLSearchParams({ client_id: clientId });
  return `https://login.microsoftonline.com/${tenantId}/adminconsent?${params}`;
}
