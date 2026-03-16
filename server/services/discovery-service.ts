import { GraphClient } from "./graph-client";

export interface DiscoveredUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  jobTitle: string | null;
  department: string | null;
  hasMailbox: boolean;
  hasOneDrive: boolean;
  accountEnabled: boolean;
}

export interface DiscoveredSite {
  id: string;
  displayName: string;
  webUrl: string;
  name: string;
  siteType: string; // 'group', 'communication', 'classic'
  storageUsedBytes: number | null;
  storageAllocatedBytes: number | null;
  groupId: string | null;
}

export interface DiscoveredTeam {
  id: string;
  displayName: string;
  description: string | null;
  visibility: string;
  memberCount: number;
  channelCount: number;
}

export interface DiscoveredPowerPlatformItem {
  id: string;
  displayName: string;
  itemType: 'app' | 'flow' | 'environment';
  environment: string | null;
  note: string;
}

export async function discoverUsers(client: GraphClient): Promise<DiscoveredUser[]> {
  const users = await client.getAllPages<any>(
    `/users?$select=id,displayName,userPrincipalName,mail,jobTitle,department,accountEnabled,assignedLicenses&$top=999`
  );

  return users
    .filter((u: any) => u.userPrincipalName && !u.userPrincipalName.includes('#EXT#'))
    .map((u: any) => ({
      id: u.id,
      displayName: u.displayName || u.userPrincipalName,
      userPrincipalName: u.userPrincipalName,
      mail: u.mail || null,
      jobTitle: u.jobTitle || null,
      department: u.department || null,
      hasMailbox: (u.assignedLicenses || []).length > 0,
      hasOneDrive: (u.assignedLicenses || []).length > 0,
      accountEnabled: u.accountEnabled !== false,
    }));
}

export async function discoverSharePointSites(client: GraphClient): Promise<DiscoveredSite[]> {
  const results: DiscoveredSite[] = [];

  // Get all sites via search
  const sitesData = await client.getAllPages<any>(`/sites?search=*&$select=id,displayName,webUrl,name,sharepointIds`);

  for (const site of sitesData) {
    // Skip personal OneDrive sites
    if (site.webUrl && site.webUrl.includes('-my.sharepoint.com')) continue;
    if (!site.name || site.name === 'root') continue;

    let storageUsed: number | null = null;
    let storageAllocated: number | null = null;
    let siteType = 'classic';
    let groupId: string | null = null;

    try {
      const siteDetails = await client.get(`/sites/${site.id}?$expand=drive`);
      if (siteDetails.drive?.quota) {
        storageUsed = siteDetails.drive.quota.used || null;
        storageAllocated = siteDetails.drive.quota.total || null;
      }
    } catch {
      // quota not available for all sites
    }

    // Determine site type from sharepointIds or URL
    if (site.sharepointIds?.siteUrl) {
      if (site.webUrl.match(/\/sites\//)) {
        siteType = 'communication';
      }
    }

    // Check if group-connected
    try {
      const groups = await client.get(`/groups?$filter=sharepointSiteUrl eq '${encodeURIComponent(site.webUrl)}'&$select=id`);
      if (groups.value && groups.value.length > 0) {
        groupId = groups.value[0].id;
        siteType = 'group';
      }
    } catch {
      // not a group site
    }

    results.push({
      id: site.id,
      displayName: site.displayName || site.name,
      webUrl: site.webUrl,
      name: site.name,
      siteType,
      storageUsedBytes: storageUsed,
      storageAllocatedBytes: storageAllocated,
      groupId,
    });
  }

  return results;
}

export async function discoverTeams(client: GraphClient): Promise<DiscoveredTeam[]> {
  const teams = await client.getAllPages<any>(
    `/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description,visibility,members`
  );

  const results: DiscoveredTeam[] = [];

  for (const team of teams) {
    let memberCount = 0;
    let channelCount = 0;

    try {
      const members = await client.get(`/teams/${team.id}/members?$top=1`);
      memberCount = members['@odata.count'] || 0;
    } catch { }

    try {
      const channels = await client.get(`/teams/${team.id}/channels`);
      channelCount = (channels.value || []).length;
    } catch { }

    results.push({
      id: team.id,
      displayName: team.displayName,
      description: team.description || null,
      visibility: team.visibility || 'Private',
      memberCount,
      channelCount,
    });
  }

  return results;
}

export interface DiscoveredOneDrive {
  id: string;
  displayName: string;
  userPrincipalName: string;
  storageUsedBytes: number | null;
  storageAllocatedBytes: number | null;
  webUrl: string | null;
  lastModified: string | null;
}

export interface DiscoveredGroup {
  id: string;
  displayName: string;
  mail: string | null;
  mailNickname: string | null;
  description: string | null;
  visibility: string | null;
  memberCount: number;
  ownerCount: number;
}

export interface DiscoveredSharedMailbox {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  mailboxType: string; // 'shared' | 'user' | '' etc.
  memberCount: number;
}

export async function discoverDistributionGroups(client: GraphClient): Promise<DiscoveredGroup[]> {
  const selectFields = 'id,displayName,mail,mailNickname,description,visibility,groupTypes,securityEnabled';
  // Prefer server-side filter that excludes M365 Unified groups; fall back if NOT() is unsupported.
  let groups: any[] = [];
  try {
    // NOT() on a collection property requires ConsistencyLevel: eventual
    groups = await client.getAllPagesAdvanced<any>(
      `/groups?$filter=mailEnabled eq true and NOT(groupTypes/any(c:c eq 'Unified'))&$count=true&$select=${selectFields}&$top=100`
    );
  } catch {
    // Fallback: simple property filters that don't need ConsistencyLevel
    const [dl, mesg] = await Promise.all([
      client.getAllPages<any>(`/groups?$filter=mailEnabled eq true and securityEnabled eq false&$select=${selectFields}&$top=100`).catch(() => [] as any[]),
      client.getAllPages<any>(`/groups?$filter=mailEnabled eq true and securityEnabled eq true&$select=${selectFields}&$top=100`).catch(() => [] as any[]),
    ]);
    groups = [...dl, ...mesg];
  }
  const results: DiscoveredGroup[] = [];
  const seen = new Set<string>();
  for (const g of groups) {
    // Always exclude M365 Unified groups — they belong to the M365 Groups discovery tab
    if (Array.isArray(g.groupTypes) && g.groupTypes.includes('Unified')) continue;
    if (seen.has(g.id)) continue;
    seen.add(g.id);
    let memberCount = 0, ownerCount = 0;
    try { const r = await client.get(`/groups/${g.id}/members?$select=id&$top=999`); memberCount = (r.value || []).length; } catch { }
    try { const r = await client.get(`/groups/${g.id}/owners?$select=id&$top=999`); ownerCount = (r.value || []).length; } catch { }
    results.push({ id: g.id, displayName: g.displayName, mail: g.mail || null, mailNickname: g.mailNickname || null, description: g.description || null, visibility: g.visibility || null, memberCount, ownerCount });
  }
  return results;
}

export async function discoverSharedMailboxes(client: GraphClient): Promise<DiscoveredSharedMailbox[]> {
  // Strategy: fetch all users with license + accountEnabled info, filter to candidates,
  // then confirm each via mailboxSettings.userPurpose.
  // Candidates = unlicensed OR sign-in disabled (accountEnabled=false) accounts with a mail address.
  // Cannot use $count-based OData filters without ConsistencyLevel header, so we filter client-side.
  const allUsers = await client.getAllPages<any>(
    `/users?$select=id,displayName,userPrincipalName,mail,assignedLicenses,accountEnabled&$top=999`
  );

  // Phase 1: broad client-side filter.
  // Include accounts that are either:
  //   a) Unlicensed AND have a mail address (most shared mailboxes)
  //   b) Sign-in disabled (accountEnabled=false) AND have a mail address
  //      — Exchange sets accountEnabled=false for shared mailboxes
  const candidates = allUsers.filter((u: any) => {
    if (!u.mail) return false;
    if (u.userPrincipalName?.includes('#EXT#')) return false; // skip guests
    const isUnlicensed = (u.assignedLicenses || []).length === 0;
    const isDisabled = u.accountEnabled === false;
    return isUnlicensed || isDisabled;
  });

  // Phase 2: confirm via mailboxSettings.userPurpose
  // userPurpose values: 'shared', 'user', 'linked', 'room', 'equipment', 'others', or absent
  // Accept 'shared' (explicit), blank (field not returned by tenant), or 'user' on a disabled account
  // (Exchange sometimes reports 'user' for shared mailboxes — disabled sign-in is the real indicator)
  // Skip 'room', 'equipment', 'others'
  const results: DiscoveredSharedMailbox[] = [];
  for (const u of candidates) {
    try {
      const settings = await client.get(`/users/${u.id}/mailboxSettings`);
      const purpose: string = settings?.userPurpose || '';
      const isShared = purpose === 'shared';
      const isPurposeBlank = purpose === '';
      const isDisabledWithUserPurpose = u.accountEnabled === false && purpose === 'user';
      if (isShared || isPurposeBlank || isDisabledWithUserPurpose) {
        results.push({
          id: u.id,
          displayName: u.displayName || u.userPrincipalName,
          userPrincipalName: u.userPrincipalName,
          mail: u.mail,
          mailboxType: 'shared',
          memberCount: 0,
        });
      }
    } catch { /* no mailbox or insufficient permissions — skip */ }
  }
  return results;
}

export async function discoverM365Groups(client: GraphClient): Promise<DiscoveredGroup[]> {
  // any() lambda on groupTypes is an advanced query — requires ConsistencyLevel: eventual
  let groups: any[] = [];
  try {
    groups = await client.getAllPagesAdvanced<any>(
      `/groups?$filter=groupTypes/any(c:c eq 'Unified')&$count=true&$select=id,displayName,mail,mailNickname,description,visibility,groupTypes&$top=100`
    );
  } catch {
    groups = await client.getAllPages<any>(
      `/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName,mail,mailNickname,description,visibility,groupTypes&$top=100`
    ).catch(() => []);
  }
  const results: DiscoveredGroup[] = [];
  for (const g of groups) {
    let memberCount = 0, ownerCount = 0;
    try { const r = await client.get(`/groups/${g.id}/members?$select=id&$top=999`); memberCount = (r.value || []).length; } catch { }
    try { const r = await client.get(`/groups/${g.id}/owners?$select=id&$top=999`); ownerCount = (r.value || []).length; } catch { }
    results.push({ id: g.id, displayName: g.displayName, mail: g.mail || null, mailNickname: g.mailNickname || null, description: g.description || null, visibility: g.visibility || 'Private', memberCount, ownerCount });
  }
  return results;
}

export async function discoverOneDrives(client: GraphClient): Promise<DiscoveredOneDrive[]> {
  const users = await client.getAllPages<any>(
    `/users?$select=id,displayName,userPrincipalName,assignedLicenses&$top=999`
  );

  const licensedUsers = users.filter((u: any) =>
    u.userPrincipalName &&
    !u.userPrincipalName.includes('#EXT#') &&
    (u.assignedLicenses || []).length > 0
  );

  const results: DiscoveredOneDrive[] = [];

  for (const user of licensedUsers) {
    try {
      const drive = await client.get(`/users/${user.id}/drive`);
      if (drive?.id) {
        results.push({
          id: user.id,
          displayName: user.displayName || user.userPrincipalName,
          userPrincipalName: user.userPrincipalName,
          storageUsedBytes: drive.quota?.used ?? null,
          storageAllocatedBytes: drive.quota?.total ?? null,
          webUrl: drive.webUrl ?? null,
          lastModified: drive.lastModifiedDateTime ?? null,
        });
      }
    } catch {
      // User has no OneDrive provisioned — skip
    }
  }

  return results;
}

export async function discoverPowerPlatform(_client: GraphClient): Promise<DiscoveredPowerPlatformItem[]> {
  // Power Platform uses separate APIs (api.powerapps.com, api.flow.microsoft.com)
  // These are not accessible via standard Microsoft Graph API with client credentials
  // Return an informational entry
  return [
    {
      id: 'powerplatform-info',
      displayName: 'Power Platform Migration',
      itemType: 'environment',
      environment: null,
      note: 'Power Platform (Power Apps, Power Automate) migration requires separate Power Platform admin API credentials. Use the Power Platform CoE Starter Kit or PowerShell module (Microsoft.PowerApps.Administration.PowerShell) to export and import apps and flows between tenants.',
    },
  ];
}
