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
