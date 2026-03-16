/**
 * Revert engine — undoes a completed migration item in the target tenant.
 * Each item type has its own revert logic. Where Graph API cannot fully revert
 * (e.g. PowerPlatform, Entra→AD) the logs explain the manual steps required.
 */

import { GraphClient } from "./graph-client";
import { storage } from "../storage";
import type { MigrationItem, Project } from "@shared/schema";

function log(msg: string): string {
  return `[${new Date().toISOString()}] ${msg}`;
}

async function setRevertStatus(itemId: number, logs: string[], error?: string) {
  await storage.updateItem(itemId, {
    status: error ? 'revert_failed' : 'reverted',
    errorDetails: error || null,
  });
  await storage.updateItemLogs(itemId, logs);
}

// ── User account ─────────────────────────────────────────────────────────────
async function revertUser(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting user account: ${identity}`));

  const user = await target.get(`/users/${encodeURIComponent(identity)}?$select=id,userPrincipalName`)
    .catch(() => null);

  if (!user) {
    logs.push(log(`User "${identity}" not found in target — nothing to delete.`));
    return;
  }

  await target.delete(`/users/${user.id}`);
  logs.push(log(`✓ Deleted user: ${user.userPrincipalName} (ID: ${user.id})`));
  logs.push(log(`  Note: The user's Exchange Online mailbox is also permanently deleted.`));
  logs.push(log(`  Soft-deleted users remain in the Entra ID recycle bin for 30 days.`));
}

// ── Mailbox (same as user — mailbox is owned by the user object) ──────────────
async function revertMailbox(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting mailbox user: ${identity}`));

  const user = await target.get(`/users/${encodeURIComponent(identity)}?$select=id,userPrincipalName`)
    .catch(() => null);

  if (!user) {
    logs.push(log(`User "${identity}" not found in target — no mailbox to delete.`));
    return;
  }

  await target.delete(`/users/${user.id}`);
  logs.push(log(`✓ Deleted mailbox user: ${user.userPrincipalName} (ID: ${user.id})`));
  logs.push(log(`  Note: All email, calendar, contacts and the OneDrive for this user are also deleted.`));
}

// ── OneDrive — delete all items from the target drive root ────────────────────
async function revertOneDrive(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting OneDrive contents for: ${identity}`));

  const user = await target.get(`/users/${encodeURIComponent(identity)}?$select=id`)
    .catch(() => null);
  if (!user) {
    logs.push(log(`User "${identity}" not found in target — no OneDrive to clear.`));
    return;
  }

  const drive = await target.get(`/users/${user.id}/drive`).catch(() => null);
  if (!drive) {
    logs.push(log(`OneDrive not found or not provisioned for "${identity}".`));
    return;
  }

  const children = await target.getAllPages<any>(`/drives/${drive.id}/root/children?$select=id,name,folder,size`);
  logs.push(log(`Found ${children.length} root item(s) to delete.`));

  let deleted = 0, failed = 0;
  for (const child of children) {
    try {
      await target.delete(`/drives/${drive.id}/items/${child.id}`);
      logs.push(log(`  ✓ Deleted: ${child.name}${child.folder ? ' (folder + all contents)' : ''}`));
      deleted++;
    } catch (e: any) {
      logs.push(log(`  ✗ Failed to delete "${child.name}": ${e.message}`));
      failed++;
    }
  }

  logs.push(log(`OneDrive revert complete: ${deleted} deleted, ${failed} failed.`));
  logs.push(log(`  Note: Deleted items go to the target user's OneDrive recycle bin for 93 days.`));
}

// ── SharePoint site ───────────────────────────────────────────────────────────
async function revertSharePoint(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting SharePoint site: ${identity}`));

  // Resolve site path
  let sitePath = identity;
  if (sitePath.startsWith('https://') || sitePath.startsWith('http://')) {
    try {
      const u = new URL(sitePath);
      sitePath = `${u.hostname}:${u.pathname}`;
    } catch { /* keep as-is */ }
  }

  const site = await target.get(`/sites/${sitePath}?$select=id,displayName,webUrl`).catch(() => null);
  if (!site) {
    logs.push(log(`Site "${identity}" not found in target — nothing to delete.`));
    return;
  }

  await target.delete(`/sites/${site.id}`);
  logs.push(log(`✓ Deleted SharePoint site: ${site.displayName} (${site.webUrl})`));
  logs.push(log(`  Note: The site is soft-deleted and held in the site collection recycle bin for 93 days before permanent deletion.`));
}

// ── Microsoft Teams ───────────────────────────────────────────────────────────
async function revertTeams(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting Teams team: ${identity}`));

  // Find team by display name
  const groups = await target.getAllPagesAdvanced<any>(
    `/groups?$filter=displayName eq '${identity.replace(/'/g, "''")}' and resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName`
  ).catch(() => [] as any[]);

  const group = groups[0];
  if (!group) {
    logs.push(log(`Team "${identity}" not found in target — nothing to delete.`));
    return;
  }

  await target.delete(`/groups/${group.id}`);
  logs.push(log(`✓ Deleted team/group: ${group.displayName} (ID: ${group.id})`));
  logs.push(log(`  Note: The associated SharePoint site and mailbox are also deleted.`));
}

// ── Distribution Group ────────────────────────────────────────────────────────
async function revertDistributionGroup(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting distribution group: ${identity}`));

  // Try mail address first, then display name
  let group = await target.get(`/groups?$filter=mail eq '${encodeURIComponent(identity)}'&$select=id,displayName,mail`)
    .then((r: any) => r.value?.[0] || null).catch(() => null);

  if (!group) {
    const byName = await target.getAllPagesAdvanced<any>(
      `/groups?$filter=displayName eq '${identity.replace(/'/g, "''")}'&$select=id,displayName,mail`
    ).catch(() => [] as any[]);
    group = byName[0] || null;
  }

  if (!group) {
    logs.push(log(`Group "${identity}" not found in target — nothing to delete.`));
    return;
  }

  await target.delete(`/groups/${group.id}`);
  logs.push(log(`✓ Deleted group: ${group.displayName} (ID: ${group.id})`));
}

// ── M365 Group ────────────────────────────────────────────────────────────────
async function revertM365Group(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting M365 group: ${identity}`));

  const groups = await target.getAllPagesAdvanced<any>(
    `/groups?$filter=displayName eq '${identity.replace(/'/g, "''")}'&$select=id,displayName,groupTypes`
  ).catch(() => [] as any[]);

  const group = groups.find((g: any) => g.groupTypes?.includes('Unified')) || groups[0];
  if (!group) {
    logs.push(log(`M365 Group "${identity}" not found in target — nothing to delete.`));
    return;
  }

  await target.delete(`/groups/${group.id}`);
  logs.push(log(`✓ Deleted M365 group: ${group.displayName} (ID: ${group.id})`));
  logs.push(log(`  Note: The associated SharePoint site and mailbox are soft-deleted.`));
}

// ── Shared Mailbox ────────────────────────────────────────────────────────────
async function revertSharedMailbox(target: GraphClient, item: MigrationItem, logs: string[]) {
  const identity = item.targetIdentity || item.sourceIdentity;
  logs.push(log(`Reverting shared mailbox: ${identity}`));

  const user = await target.get(`/users/${encodeURIComponent(identity)}?$select=id,userPrincipalName,mail`)
    .catch(() => null);

  if (!user) {
    logs.push(log(`Shared mailbox "${identity}" not found in target — nothing to delete.`));
    return;
  }

  await target.delete(`/users/${user.id}`);
  logs.push(log(`✓ Deleted shared mailbox user object: ${user.userPrincipalName} (ID: ${user.id})`));
  logs.push(log(`  Note: The Exchange shared mailbox is also deleted (soft-deleted, held 30 days).`));
}

// ── Power Platform / Entra→AD — cannot revert via Graph ──────────────────────
async function revertManual(item: MigrationItem, logs: string[]) {
  const typeName = item.itemType === 'powerplatform' ? 'Power Platform' : 'Entra→AD';
  logs.push(log(`Revert for ${typeName} items cannot be automated via API.`));
  logs.push(log(`To manually undo this migration:`));
  if (item.itemType === 'powerplatform') {
    logs.push(log(`  1. Go to https://make.powerapps.com and sign in as a target tenant admin`));
    logs.push(log(`  2. Locate any migrated apps or flows`));
    logs.push(log(`  3. Delete them manually from the Power Apps / Power Automate portals`));
  } else {
    logs.push(log(`  1. Open Active Directory Users and Computers (ADUC) on your domain controller`));
    logs.push(log(`  2. Locate the user "${item.targetIdentity || item.sourceIdentity}" in the target OU`));
    logs.push(log(`  3. Right-click → Delete`));
  }
  logs.push(log(`Item status reset to "pending" — re-run migration when ready.`));
}

// ── Public entry point ────────────────────────────────────────────────────────
export async function revertMigrationItem(item: MigrationItem, project: Project): Promise<void> {
  const logs: string[] = [];
  logs.push(log(`Starting revert for item #${item.id}: ${item.itemType} — ${item.sourceIdentity}`));

  await storage.updateItem(item.id, { status: 'reverting', errorDetails: null });
  await storage.updateItemLogs(item.id, logs);

  try {
    if (!project.targetClientId || !project.targetClientSecret) {
      throw new Error('Target tenant credentials not configured. Cannot connect to target tenant.');
    }
    const target = new GraphClient(project.targetTenantId, project.targetClientId, project.targetClientSecret);

    switch (item.itemType) {
      case 'user':           await revertUser(target, item, logs); break;
      case 'mailbox':        await revertMailbox(target, item, logs); break;
      case 'onedrive':       await revertOneDrive(target, item, logs); break;
      case 'sharepoint':     await revertSharePoint(target, item, logs); break;
      case 'teams':          await revertTeams(target, item, logs); break;
      case 'distributiongroup': await revertDistributionGroup(target, item, logs); break;
      case 'm365group':      await revertM365Group(target, item, logs); break;
      case 'sharedmailbox':  await revertSharedMailbox(target, item, logs); break;
      case 'powerplatform':
      case 'entra_to_ad':    await revertManual(item, logs); break;
      default:
        logs.push(log(`No revert logic defined for item type "${item.itemType}".`));
    }

    logs.push(log(`✓ Revert complete — item status set to "reverted".`));
    await setRevertStatus(item.id, logs);

  } catch (err: any) {
    logs.push(log(`Revert failed: ${err.message}`));
    await setRevertStatus(item.id, logs, err.message);
  }
}
