import { GraphClient } from "./graph-client";
import { storage } from "../storage";
import type { Project, MigrationItem } from "@shared/schema";
import { db } from "../db";
import { migrationProjects, migrationItems } from "@shared/schema";
import { eq, and, lte, isNotNull } from "drizzle-orm";

function syncLog(message: string): string {
  return `[${new Date().toISOString()}] [SYNC] ${message}`;
}

function getGraphClients(project: Project): { source: GraphClient; target: GraphClient } {
  if (!project.sourceClientId || !project.sourceClientSecret) throw new Error("Source credentials not configured");
  if (!project.targetClientId || !project.targetClientSecret) throw new Error("Target credentials not configured");
  return {
    source: new GraphClient(project.sourceTenantId, project.sourceClientId, project.sourceClientSecret),
    target: new GraphClient(project.targetTenantId, project.targetClientId, project.targetClientSecret),
  };
}

// ── Mailbox delta sync ─────────────────────────────────────────────────────
// Finds new messages received in source since lastSyncedAt and copies them to target.
async function syncMailboxDelta(
  source: GraphClient,
  target: GraphClient,
  sourceUpn: string,
  targetUpn: string,
  lastSyncedAt: Date | null
): Promise<{ newMessages: number; logs: string[]; lastSuccessfulReceivedAt: Date | null }> {
  const logs: string[] = [];
  let newMessages = 0;

  // Resolve source user
  const sourceUser = await source.get(`/users/${encodeURIComponent(sourceUpn)}?$select=id`).catch(() => null);
  if (!sourceUser) {
    logs.push(syncLog(`Source user ${sourceUpn} not found — skipping mailbox sync`));
    return { newMessages, logs, lastSuccessfulReceivedAt: null };
  }

  // Resolve target user
  const targetUser = await target.get(`/users/${encodeURIComponent(targetUpn)}?$select=id`).catch(() => null);
  if (!targetUser) {
    logs.push(syncLog(`Target user ${targetUpn} not found — skipping mailbox sync`));
    return { newMessages, logs, lastSuccessfulReceivedAt: null };
  }

  // Get messages received since lastSyncedAt (or last 24h if no previous sync)
  const since = lastSyncedAt
    ? lastSyncedAt.toISOString()
    : new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();

  logs.push(syncLog(`Fetching messages received since ${since}`));

  let messages: any[] = [];
  try {
    messages = await source.getAllPages<any>(
      `/users/${sourceUser.id}/messages?$filter=receivedDateTime gt ${since}&$select=id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,body,importance,isRead,hasAttachments,internetMessageId,parentFolderId&$orderby=receivedDateTime asc&$top=50`
    );
  } catch (e: any) {
    logs.push(syncLog(`Failed to fetch messages: ${e.message}`));
    return { newMessages, logs, lastSuccessfulReceivedAt: null };
  }

  logs.push(syncLog(`Found ${messages.length} new message(s) since ${since}`));

  // Track the receivedDateTime of the last successfully copied message so we advance
  // lastSyncedAt accurately (avoiding re-fetching already-copied messages or losing failures)
  let lastSuccessfulReceivedAt: Date | null = null;

  for (const msg of messages) {
    try {
      // Resolve target folder — match by well-known name or display name; fall back to inbox
      let targetFolderId: string | null = null;
      try {
        const srcFolder = await source.get(`/users/${sourceUser.id}/mailFolders/${msg.parentFolderId}?$select=displayName,wellKnownName`).catch(() => null);
        if (srcFolder) {
          const wkn = srcFolder.wellKnownName;
          if (wkn) {
            const tgt = await target.get(`/users/${targetUser.id}/mailFolders/${wkn}`).catch(() => null);
            if (tgt) targetFolderId = tgt.id;
          }
          if (!targetFolderId) {
            const dispName = (srcFolder.displayName || '').replace(/'/g, "''");
            const existing = await target.get(`/users/${targetUser.id}/mailFolders?$filter=displayName eq '${dispName}'`).catch(() => null);
            if (existing?.value?.length) targetFolderId = existing.value[0].id;
          }
        }
      } catch { /* fall back to inbox */ }

      // Build payload with only writable properties.
      // receivedDateTime and internetMessageId are read-only in Graph API and must NOT be sent
      // — including them causes HTTP 400 and silently prevents all message copies.
      const payload: any = {
        subject: msg.subject || '(No subject)',
        body: msg.body || { contentType: 'text', content: '' },
        toRecipients: msg.toRecipients || [],
        ccRecipients: msg.ccRecipients || [],
        bccRecipients: msg.bccRecipients || [],
        importance: msg.importance || 'normal',
        isRead: msg.isRead ?? false,
        isDraft: false,  // Create as received message (not draft) so it appears in inbox properly
      };
      // from is settable with Mail.ReadWrite Application permission
      if (msg.from) payload.from = msg.from;

      // Always target inbox (or matched folder) — posting to /messages creates a draft in Drafts
      const endpoint = targetFolderId
        ? `/users/${targetUser.id}/mailFolders/${targetFolderId}/messages`
        : `/users/${targetUser.id}/mailFolders/inbox/messages`;

      const created = await target.post(endpoint, payload);
      newMessages++;
      lastSuccessfulReceivedAt = msg.receivedDateTime ? new Date(msg.receivedDateTime) : now;

      // Copy attachments (best-effort — size limit 4 MB per attachment)
      if (msg.hasAttachments && created?.id) {
        try {
          const attachments = await source.getAllPages<any>(`/users/${sourceUser.id}/messages/${msg.id}/attachments`);
          for (const att of attachments) {
            if ((att.size ?? 0) > 4 * 1024 * 1024) continue;
            await target.post(`/users/${targetUser.id}/messages/${created.id}/attachments`, {
              '@odata.type': att['@odata.type'] || '#microsoft.graph.fileAttachment',
              name: att.name,
              contentType: att.contentType,
              contentBytes: att.contentBytes,
            }).catch(() => {});
          }
        } catch { /* attachment errors are non-fatal */ }
      }
    } catch (e: any) {
      logs.push(syncLog(`  Failed to copy message "${msg.subject}": ${e.message}`));
    }
  }

  logs.push(syncLog(`Copied ${newMessages}/${messages.length} new message(s) to target mailbox`));
  return { newMessages, logs, lastSuccessfulReceivedAt };
}

// ── OneDrive / SharePoint delta sync ──────────────────────────────────────
// Uses Graph delta API to efficiently find new/modified files since last token.
async function syncDriveDelta(
  source: GraphClient,
  target: GraphClient,
  sourceDriveId: string,
  targetDriveId: string,
  deltaToken: string | null
): Promise<{ filesSync: number; newDeltaToken: string | null; logs: string[] }> {
  const logs: string[] = [];
  let filesSync = 0;
  let newDeltaToken: string | null = null;

  // Build delta URL — use stored token for incremental, or get initial state
  const deltaUrl = deltaToken
    ? deltaToken  // deltaToken IS the full nextLink / deltaLink URL from the previous run
    : `/drives/${sourceDriveId}/root/delta?$select=id,name,file,folder,parentReference,lastModifiedDateTime,size,webUrl`;

  logs.push(syncLog(`Running drive delta sync (${deltaToken ? 'incremental' : 'initial'})`));

  let page = deltaUrl;
  const changedItems: any[] = [];

  while (page) {
    let data: any;
    try {
      data = await source.get(page);
    } catch (e: any) {
      logs.push(syncLog(`Delta fetch failed: ${e.message}`));
      break;
    }
    if (data.value) changedItems.push(...data.value.filter((i: any) => i.file)); // only files, not folders
    newDeltaToken = data['@odata.deltaLink'] || null;
    page = data['@odata.nextLink'] || null;
  }

  logs.push(syncLog(`Delta: ${changedItems.length} changed file(s) to sync`));

  for (const item of changedItems) {
    try {
      // Build target path mirroring source
      const parentPath = item.parentReference?.path?.replace(`/drives/${sourceDriveId}/root`, '') || '/';
      const targetPath = `${parentPath}/${item.name}`.replace('//', '/');

      // Check size — use upload session for large files
      if ((item.size || 0) <= 4 * 1024 * 1024) {
        const content = await source.getBuffer(`/drives/${sourceDriveId}/items/${item.id}/content`);
        await target.put(`/drives/${targetDriveId}/root:${targetPath}:/content`, content, 'application/octet-stream');
      } else {
        const sessionRes = await target.post(`/drives/${targetDriveId}/root:${targetPath}:/createUploadSession`, {
          item: { '@microsoft.graph.conflictBehavior': 'replace', name: item.name },
        });
        if (sessionRes?.uploadUrl) {
          const content = await source.getBuffer(`/drives/${sourceDriveId}/items/${item.id}/content`);
          const chunkSize = 10 * 1024 * 1024;
          for (let offset = 0; offset < content.length; offset += chunkSize) {
            const chunk = content.slice(offset, Math.min(offset + chunkSize, content.length));
            await fetch(sessionRes.uploadUrl, {
              method: 'PUT',
              headers: {
                'Content-Length': chunk.length.toString(),
                'Content-Range': `bytes ${offset}-${offset + chunk.length - 1}/${content.length}`,
              },
              body: chunk,
            });
          }
        }
      }
      filesSync++;
    } catch (e: any) {
      logs.push(syncLog(`  Failed to sync file "${item.name}": ${e.message}`));
    }
  }

  logs.push(syncLog(`Synced ${filesSync}/${changedItems.length} file(s)`));
  return { filesSync, newDeltaToken, logs };
}

// ── Per-item sync dispatcher ───────────────────────────────────────────────
async function syncItem(project: Project, item: MigrationItem, intervalMinutes: number): Promise<void> {
  const { source, target } = getGraphClients(project);
  const targetIdentity = item.targetIdentity || item.sourceIdentity;
  const now = new Date();

  const appendLogs = async (newLogs: string[]) => {
    const current = await storage.getItem(item.id);
    const existing = current?.logs || [];
    // Keep last 500 log lines to avoid unbounded growth
    const combined = [...existing, ...newLogs].slice(-500);
    await storage.updateItemLogs(item.id, combined);
  };

  try {
    if (item.itemType === 'mailbox') {
      const { newMessages, logs, lastSuccessfulReceivedAt } = await syncMailboxDelta(source, target, item.sourceIdentity, targetIdentity, item.lastSyncedAt ?? null);
      await appendLogs(logs);
      // Advance lastSyncedAt only to the last successfully copied message time so that
      // any messages that failed to copy are retried on the next sync run.
      // If no new messages were found at all, advance to now (normal interval tick).
      const advanceTo = newMessages > 0 && lastSuccessfulReceivedAt ? lastSuccessfulReceivedAt : now;
      await db.update(migrationItems).set({
        lastSyncedAt: advanceTo,
        nextSyncAt: new Date(now.getTime() + intervalMinutes * 60 * 1000),
        ...(newMessages > 0 ? { updatedAt: now } : {}),
      }).where(eq(migrationItems.id, item.id));

    } else if (item.itemType === 'onedrive') {
      // Get source drive ID
      const sourceLookup = await source.get(`/users/${encodeURIComponent(item.sourceIdentity)}/drive?$select=id`).catch(() => null);
      const targetLookup = await target.get(`/users/${encodeURIComponent(targetIdentity)}/drive?$select=id`).catch(() => null);
      if (!sourceLookup?.id || !targetLookup?.id) {
        await appendLogs([syncLog(`OneDrive sync: could not resolve drive IDs for ${item.sourceIdentity}`)]);
        return;
      }
      const { filesSync, newDeltaToken, logs } = await syncDriveDelta(source, target, sourceLookup.id, targetLookup.id, item.deltaToken ?? null);
      await appendLogs(logs);
      await db.update(migrationItems).set({
        deltaToken: newDeltaToken ?? item.deltaToken,
        lastSyncedAt: now,
        nextSyncAt: new Date(now.getTime() + intervalMinutes * 60 * 1000),
        updatedAt: filesSync > 0 ? now : item.updatedAt,
      }).where(eq(migrationItems.id, item.id));

    } else if (item.itemType === 'sharepoint') {
      // Resolve site drives
      const siteId = item.sourceIdentity; // webUrl or siteId
      let sourceSite: any = null;
      try { sourceSite = await source.get(`/sites/${siteId}?$select=id`); } catch { }
      if (!sourceSite) {
        try {
          const url = new URL(siteId.startsWith('http') ? siteId : `https://${siteId}`);
          const host = url.hostname;
          const sitePath = url.pathname.replace('/sites/', '');
          sourceSite = await source.get(`/sites/${host}:/sites/${sitePath}?$select=id`).catch(() => null);
        } catch { }
      }
      if (!sourceSite?.id) {
        await appendLogs([syncLog(`SharePoint sync: could not resolve source site for ${item.sourceIdentity}`)]);
        return;
      }
      const sourceDrives = await source.getAllPages<any>(`/sites/${sourceSite.id}/drives?$select=id,name`);
      const primaryDrive = sourceDrives[0];
      if (!primaryDrive) {
        await appendLogs([syncLog(`SharePoint sync: no drives found on source site`)]);
        return;
      }

      // For target, we stored targetIdentity as the display name — find by name
      const targetSites = await target.getAllPages<any>(`/sites?search=${encodeURIComponent(item.targetIdentity || '')}&$select=id,displayName`).catch(() => []);
      const targetSite = targetSites[0];
      const targetDrives = targetSite ? await target.getAllPages<any>(`/sites/${targetSite.id}/drives?$select=id,name`).catch(() => []) : [];
      const targetDrive = targetDrives[0];

      if (!targetDrive) {
        await appendLogs([syncLog(`SharePoint sync: could not resolve target site/drive`)]);
        return;
      }

      const { filesSync, newDeltaToken, logs } = await syncDriveDelta(source, target, primaryDrive.id, targetDrive.id, item.deltaToken ?? null);
      await appendLogs(logs);
      await db.update(migrationItems).set({
        deltaToken: newDeltaToken ?? item.deltaToken,
        lastSyncedAt: now,
        nextSyncAt: new Date(now.getTime() + intervalMinutes * 60 * 1000),
      }).where(eq(migrationItems.id, item.id));

    } else if (item.itemType === 'sharedmailbox') {
      // Same sync logic as mailbox — new messages from source are copied to target
      const { newMessages, logs, lastSuccessfulReceivedAt } = await syncMailboxDelta(source, target, item.sourceIdentity, targetIdentity, item.lastSyncedAt ?? null);
      await appendLogs(logs);
      const advanceTo = newMessages > 0 && lastSuccessfulReceivedAt ? lastSuccessfulReceivedAt : now;
      await db.update(migrationItems).set({
        lastSyncedAt: advanceTo,
        nextSyncAt: new Date(now.getTime() + intervalMinutes * 60 * 1000),
        ...(newMessages > 0 ? { updatedAt: now } : {}),
      }).where(eq(migrationItems.id, item.id));

    } else {
      // Groups, users, teams — not continuously syncable
      await db.update(migrationItems).set({
        lastSyncedAt: now,
        nextSyncAt: new Date(now.getTime() + intervalMinutes * 60 * 1000),
      }).where(eq(migrationItems.id, item.id));
    }
  } catch (e: any) {
    console.error(`[sync] item ${item.id} (${item.itemType}) sync error:`, e.message);
    await appendLogs([syncLog(`Sync error: ${e.message}`)]);
    await db.update(migrationItems).set({
      nextSyncAt: new Date(now.getTime() + intervalMinutes * 60 * 1000),
    }).where(eq(migrationItems.id, item.id));
  }
}

// ── Project-level sync ─────────────────────────────────────────────────────
export async function runProjectSync(projectId: number): Promise<{ synced: number; errors: number }> {
  const project = await storage.getProjectInternal(projectId);
  if (!project || !project.syncEnabled) return { synced: 0, errors: 0 };

  const intervalMinutes = project.syncIntervalMinutes || 60;
  const items = await storage.getItems(projectId);

  // Only sync completed items (initial migration must have finished first)
  const syncable = items.filter(i => i.status === 'completed');

  console.log(`[sync] project ${projectId}: running sync for ${syncable.length} completed item(s)`);

  let synced = 0;
  let errors = 0;

  for (const item of syncable) {
    try {
      await syncItem(project, item, intervalMinutes);
      synced++;
    } catch (e: any) {
      errors++;
      console.error(`[sync] item ${item.id} failed:`, e.message);
    }
  }

  return { synced, errors };
}

// ── Background scheduler ───────────────────────────────────────────────────
// Runs every 2 minutes; only triggers projects that are due for a sync.
let schedulerRunning = false;

async function schedulerTick() {
  if (schedulerRunning) return;
  schedulerRunning = true;
  try {
    const now = new Date();

    // Find all projects with sync enabled
    const projects = await db.select().from(migrationProjects).where(eq(migrationProjects.syncEnabled, true));

    for (const project of projects) {
      try {
        const intervalMinutes = project.syncIntervalMinutes || 60;
        const items = await storage.getItems(project.id);
        const completedItems = items.filter(i => i.status === 'completed');

        if (completedItems.length === 0) continue;

        // Check if any item is due (nextSyncAt <= now, or never synced)
        const isDue = completedItems.some(i => !i.nextSyncAt || new Date(i.nextSyncAt) <= now);
        if (!isDue) continue;

        console.log(`[sync] scheduler: triggering sync for project ${project.id} (${project.name})`);
        // Run in background without blocking the scheduler
        runProjectSync(project.id).catch(e => console.error(`[sync] project ${project.id} sync failed:`, e.message));
      } catch (e: any) {
        console.error(`[sync] scheduler error for project ${project.id}:`, e.message);
      }
    }
  } finally {
    schedulerRunning = false;
  }
}

export function startSyncScheduler(): void {
  console.log('[sync] continuous sync scheduler started (checks every 2 minutes)');
  setInterval(schedulerTick, 2 * 60 * 1000);
  // Also run immediately so first sync happens quickly after enabling
  setTimeout(schedulerTick, 10 * 1000);
}
