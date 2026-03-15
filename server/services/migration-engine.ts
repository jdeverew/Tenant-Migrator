import { GraphClient } from "./graph-client";
import { storage } from "../storage";
import type { Project, MigrationItem } from "@shared/schema";

function logEntry(message: string): string {
  return `[${new Date().toISOString()}] ${message}`;
}

function getGraphClients(project: Project): { source: GraphClient; target: GraphClient } {
  if (!project.sourceClientId || !project.sourceClientSecret) {
    throw new Error("Source tenant credentials not configured");
  }
  if (!project.targetClientId || !project.targetClientSecret) {
    throw new Error("Target tenant credentials not configured");
  }
  return {
    source: new GraphClient(project.sourceTenantId, project.sourceClientId, project.sourceClientSecret),
    target: new GraphClient(project.targetTenantId, project.targetClientId, project.targetClientSecret),
  };
}

async function updateItemProgress(
  itemId: number,
  status: string,
  logs: string[],
  errorDetails?: string,
  bytes?: { bytesMigrated: number; bytesTotal: number }
) {
  const progressPercent = bytes && bytes.bytesTotal > 0
    ? Math.min(100, Math.round((bytes.bytesMigrated / bytes.bytesTotal) * 100))
    : undefined;

  await storage.updateItem(itemId, {
    status,
    errorDetails: errorDetails || null,
    ...(bytes !== undefined ? {
      bytesMigrated: bytes.bytesMigrated,
      bytesTotal: bytes.bytesTotal,
      progressPercent: progressPercent ?? 0,
    } : {}),
  });
  await storage.updateItemLogs(itemId, logs);
}

async function migrateAttachments(
  source: GraphClient,
  target: GraphClient,
  sourceUser: string,
  targetUser: string,
  sourceMessageId: string,
  targetMessageId: string,
): Promise<{ migrated: number; failed: number }> {
  let migrated = 0;
  let failed = 0;

  try {
    const attachments = await source.getAllPages<any>(
      `/users/${sourceUser}/messages/${sourceMessageId}/attachments`
    );

    for (const attachment of attachments) {
      try {
        await target.post(
          `/users/${targetUser}/messages/${targetMessageId}/attachments`,
          {
            "@odata.type": attachment["@odata.type"] || "#microsoft.graph.fileAttachment",
            name: attachment.name,
            contentType: attachment.contentType,
            contentBytes: attachment.contentBytes,
            size: attachment.size,
            isInline: attachment.isInline || false,
            contentId: attachment.contentId,
          }
        );
        migrated++;
      } catch {
        failed++;
      }
    }
  } catch {
    // if we can't read attachments, count as failure
  }

  return { migrated, failed };
}

async function migrateMailbox(
  source: GraphClient,
  target: GraphClient,
  sourceUser: string,
  targetUser: string,
  itemId: number
): Promise<void> {
  const logs: string[] = [];

  logs.push(logEntry(`Starting mailbox migration: ${sourceUser} → ${targetUser}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    logs.push(logEntry("Fetching mail folders from source..."));
    await updateItemProgress(itemId, 'in_progress', logs);

    const folders = await source.getAllPages<any>(`/users/${sourceUser}/mailFolders?$top=100`);
    logs.push(logEntry(`Found ${folders.length} mail folders`));
    await updateItemProgress(itemId, 'in_progress', logs);

    let totalMessages = 0;
    let migratedMessages = 0;
    let failedMessages = 0;
    let attachmentsMigrated = 0;
    let attachmentsFailed = 0;

    for (const folder of folders) {
      const folderName = folder.displayName;
      logs.push(logEntry(`Processing folder: ${folderName}`));
      await updateItemProgress(itemId, 'in_progress', logs);

      let targetFolderId: string;
      try {
        const targetFolders = await target.get(`/users/${targetUser}/mailFolders?$filter=displayName eq '${folderName}'`);
        if (targetFolders.value && targetFolders.value.length > 0) {
          targetFolderId = targetFolders.value[0].id;
        } else {
          const created = await target.post(`/users/${targetUser}/mailFolders`, {
            displayName: folderName,
          });
          targetFolderId = created.id;
          logs.push(logEntry(`Created folder "${folderName}" in target`));
        }
      } catch (err: any) {
        logs.push(logEntry(`Using Inbox for folder "${folderName}": ${err.message}`));
        try {
          const inbox = await target.get(`/users/${targetUser}/mailFolders/Inbox`);
          targetFolderId = inbox.id;
        } catch {
          logs.push(logEntry(`Cannot access Inbox for target user, skipping folder "${folderName}"`));
          continue;
        }
      }

      try {
        const messages = await source.getAllPages<any>(
          `/users/${sourceUser}/mailFolders/${folder.id}/messages?$top=50&$select=id,subject,body,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,importance,isRead,hasAttachments`
        );
        totalMessages += messages.length;
        logs.push(logEntry(`Found ${messages.length} messages in "${folderName}"`));
        await updateItemProgress(itemId, 'in_progress', logs);

        for (const msg of messages) {
          try {
            const newMessage: any = {
              subject: msg.subject || "(No Subject)",
              body: msg.body,
              from: msg.from,
              toRecipients: msg.toRecipients || [],
              ccRecipients: msg.ccRecipients || [],
              bccRecipients: msg.bccRecipients || [],
              receivedDateTime: msg.receivedDateTime,
              importance: msg.importance || "normal",
              isRead: msg.isRead !== undefined ? msg.isRead : true,
            };

            const createdMsg = await target.post(`/users/${targetUser}/mailFolders/${targetFolderId}/messages`, newMessage);

            if (msg.hasAttachments && createdMsg?.id) {
              const attResult = await migrateAttachments(source, target, sourceUser, targetUser, msg.id, createdMsg.id);
              attachmentsMigrated += attResult.migrated;
              attachmentsFailed += attResult.failed;
            }

            migratedMessages++;

            if (migratedMessages % 25 === 0) {
              logs.push(logEntry(`Migrated ${migratedMessages}/${totalMessages} messages so far...`));
              await updateItemProgress(itemId, 'in_progress', logs);
            }
          } catch (msgErr: any) {
            failedMessages++;
            logs.push(logEntry(`Failed to migrate message "${msg.subject}": ${msgErr.message}`));
          }
        }
      } catch (folderErr: any) {
        logs.push(logEntry(`Failed to read messages from "${folderName}": ${folderErr.message}`));
      }
    }

    const attachmentSummary = attachmentsMigrated + attachmentsFailed > 0
      ? `, ${attachmentsMigrated} attachments migrated (${attachmentsFailed} failed)`
      : '';
    logs.push(logEntry(`Mailbox migration complete: ${migratedMessages} messages migrated, ${failedMessages} failed out of ${totalMessages} total${attachmentSummary}`));

    if (failedMessages > 0 && migratedMessages === 0) {
      await updateItemProgress(itemId, 'failed', logs, `All ${failedMessages} messages failed to migrate`);
    } else if (failedMessages > 0) {
      await updateItemProgress(itemId, 'completed', logs, `${failedMessages} messages failed`);
    } else {
      await updateItemProgress(itemId, 'completed', logs);
    }
  } catch (err: any) {
    logs.push(logEntry(`Mailbox migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  return `${(bytes / 1024 / 1024 / 1024).toFixed(2)} GB`;
}

async function migrateDriveItemsRecursive(
  source: GraphClient,
  target: GraphClient,
  sourceDriveId: string,
  targetDriveId: string,
  sourceParentItemId: string | null,
  targetParentPath: string,
  itemId: number,
  logs: string[],
  counters: { migrated: number; failed: number; total: number; bytesMigrated: number; bytesTotal: number }
): Promise<void> {
  const listPath = sourceParentItemId
    ? `/drives/${sourceDriveId}/items/${sourceParentItemId}/children?$top=200`
    : `/drives/${sourceDriveId}/root/children?$top=200`;

  const items = await source.getAllPages<any>(listPath);

  for (const item of items) {
    counters.total++;

    if (item.folder) {
      logs.push(logEntry(`Processing folder: ${item.name}`));
      await updateItemProgress(itemId, 'in_progress', logs, undefined, counters);

      try {
        try {
          await target.post(`/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:/children`, {
            name: ".",
            folder: {},
            "@microsoft.graph.conflictBehavior": "replace",
          });
        } catch {
          // may fail if folder exists — that's ok
        }

        await migrateDriveItemsRecursive(
          source, target,
          sourceDriveId, targetDriveId,
          item.id,
          `${targetParentPath}/${item.name}`,
          itemId, logs, counters
        );
        counters.migrated++;
      } catch (err: any) {
        counters.failed++;
        logs.push(logEntry(`Failed to process folder "${item.name}": ${err.message}`));
      }
    } else if (item.file) {
      try {
        const fileSize = item.size || 0;
        const targetPath = `/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:/content`;

        if (fileSize > 4 * 1024 * 1024) {
          logs.push(logEntry(`Uploading large file: ${item.name} (${formatBytes(fileSize)})`));
          await updateItemProgress(itemId, 'in_progress', logs, undefined, counters);

          const uploadSession = await target.post(
            `/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:/createUploadSession`,
            {
              item: {
                "@microsoft.graph.conflictBehavior": "rename",
                name: item.name,
              },
            }
          );

          const fileBuffer = await source.getBuffer(`/drives/${sourceDriveId}/items/${item.id}/content`);

          const chunkSize = 3276800; // 3.125 MB (must be multiple of 320 KiB)
          let offset = 0;
          while (offset < fileBuffer.length) {
            const end = Math.min(offset + chunkSize, fileBuffer.length);
            const chunk = fileBuffer.subarray(offset, end);

            const uploadRes = await fetch(uploadSession.uploadUrl, {
              method: 'PUT',
              headers: {
                'Content-Range': `bytes ${offset}-${end - 1}/${fileBuffer.length}`,
                'Content-Type': 'application/octet-stream',
              },
              body: chunk,
            });

            if (!uploadRes.ok && uploadRes.status !== 202) {
              throw new Error(`Upload chunk failed: ${uploadRes.status}`);
            }
            offset = end;
          }
        } else {
          const fileBuffer = await source.getBuffer(`/drives/${sourceDriveId}/items/${item.id}/content`);
          await target.put(
            targetPath,
            fileBuffer,
            item.file.mimeType || 'application/octet-stream'
          );
        }

        counters.migrated++;
        counters.bytesMigrated += fileSize;

        const pct = counters.bytesTotal > 0
          ? ` (${Math.round(counters.bytesMigrated / counters.bytesTotal * 100)}%)`
          : '';
        logs.push(logEntry(`✓ ${item.name} — ${formatBytes(fileSize)} | Total: ${formatBytes(counters.bytesMigrated)} / ${formatBytes(counters.bytesTotal)}${pct}`));
        await updateItemProgress(itemId, 'in_progress', logs, undefined, counters);

      } catch (err: any) {
        counters.failed++;
        logs.push(logEntry(`Failed to migrate file "${item.name}": ${err.message}`));
        await updateItemProgress(itemId, 'in_progress', logs, undefined, counters);
      }
    }
  }
}

async function resolveUserDrive(client: GraphClient, userId: string, displayName: string): Promise<any> {
  const errors: string[] = [];

  // Method 1: standard /users/{id}/drive
  try {
    const drive = await client.get(`/users/${userId}/drive`);
    if (drive?.id) return drive;
  } catch (e: any) {
    errors.push(`Method 1 (/users/{id}/drive): ${e.message}`);
  }

  // Method 2: list all drives for the user
  try {
    const drivesRes = await client.get(`/users/${userId}/drives`);
    const drives: any[] = drivesRes.value || [];
    const biz = drives.find((d: any) => d.driveType === 'business') || drives[0];
    if (biz) return biz;
  } catch (e: any) {
    errors.push(`Method 2 (/users/{id}/drives): ${e.message}`);
  }

  // Method 3: resolve via the user's mySite personal SharePoint URL
  try {
    const userProfile = await client.get(`/users/${userId}?$select=mySite`);
    const mySite: string = userProfile?.mySite;
    if (mySite) {
      const url = new URL(mySite);
      const hostname = url.hostname;
      const path = url.pathname;
      const drive = await client.get(`/sites/${hostname}:${path}:/drive`);
      if (drive?.id) return drive;
    }
  } catch (e: any) {
    errors.push(`Method 3 (mySite): ${e.message}`);
  }

  throw new Error(
    `Could not access OneDrive for "${displayName}" after trying 3 methods.\n` +
    `Diagnostic log:\n${errors.join('\n')}\n\n` +
    `Most common fixes:\n` +
    `- Run this PowerShell as a Global Admin to force-provision the drive:\n` +
    `  Request-SPOPersonalSite -UserEmails @("${displayName}")\n` +
    `- Or have the user sign into https://onedrive.com once to initialise their drive.`
  );
}

async function migrateOneDrive(
  source: GraphClient,
  target: GraphClient,
  sourceUser: string,
  targetUser: string,
  itemId: number
): Promise<void> {
  const logs: string[] = [];

  logs.push(logEntry(`Starting OneDrive migration: ${sourceUser} → ${targetUser}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    logs.push(logEntry("Resolving source user..."));
    const sourceUserObj = await source.get(`/users/${sourceUser}`).catch(() => null);
    if (!sourceUserObj) {
      throw new Error(`Source user "${sourceUser}" not found. Check the email address is correct and the app has User.Read.All permission.`);
    }
    const sourceUserId = sourceUserObj.id;

    logs.push(logEntry("Resolving target user..."));
    const targetUserObj = await target.get(`/users/${targetUser}`).catch(() => null);
    if (!targetUserObj) {
      throw new Error(`Target user "${targetUser}" not found. Check the email address is correct and the app has User.Read.All permission.`);
    }
    const targetUserId = targetUserObj.id;

    logs.push(logEntry(`Source user ID: ${sourceUserId}, Target user ID: ${targetUserId}`));
    logs.push(logEntry("Reading source user's OneDrive..."));

    const sourceDrive = await resolveUserDrive(source, sourceUserId, sourceUser);
    const targetDrive = await resolveUserDrive(target, targetUserId, targetUser);

    const bytesTotal = sourceDrive.quota?.used || 0;
    logs.push(logEntry(`Source drive: ${sourceDrive.id} | Size: ${formatBytes(bytesTotal)}`));
    logs.push(logEntry(`Target drive: ${targetDrive.id}`));
    await updateItemProgress(itemId, 'in_progress', logs, undefined, { bytesMigrated: 0, bytesTotal });

    const counters = { migrated: 0, failed: 0, total: 0, bytesMigrated: 0, bytesTotal };

    await migrateDriveItemsRecursive(
      source, target,
      sourceDrive.id, targetDrive.id,
      null, "",
      itemId, logs, counters
    );

    logs.push(logEntry(`OneDrive migration complete: ${counters.migrated} files migrated (${formatBytes(counters.bytesMigrated)}), ${counters.failed} failed out of ${counters.total} total items`));

    const finalBytes = { bytesMigrated: counters.bytesMigrated, bytesTotal: counters.bytesTotal };
    if (counters.failed > 0 && counters.migrated === 0) {
      await updateItemProgress(itemId, 'failed', logs, `All ${counters.failed} items failed`, finalBytes);
    } else if (counters.failed > 0) {
      await updateItemProgress(itemId, 'completed', logs, `${counters.failed} items failed`, finalBytes);
    } else {
      await updateItemProgress(itemId, 'completed', logs, undefined, finalBytes);
    }
  } catch (err: any) {
    logs.push(logEntry(`OneDrive migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

async function migrateSharePoint(
  source: GraphClient,
  target: GraphClient,
  sourceIdentity: string,
  targetIdentity: string,
  itemId: number
): Promise<void> {
  const logs: string[] = [];

  logs.push(logEntry(`Starting SharePoint migration: ${sourceIdentity} → ${targetIdentity}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    logs.push(logEntry("Resolving source SharePoint site..."));
    await updateItemProgress(itemId, 'in_progress', logs);

    let sourceSite: any;
    try {
      sourceSite = await source.get(`/sites/${sourceIdentity}`);
    } catch {
      const searchResult = await source.get(`/sites?search=${encodeURIComponent(sourceIdentity)}`);
      if (!searchResult.value || searchResult.value.length === 0) {
        throw new Error(`Source SharePoint site "${sourceIdentity}" not found`);
      }
      sourceSite = searchResult.value[0];
    }
    logs.push(logEntry(`Found source site: ${sourceSite.displayName} (${sourceSite.id})`));

    logs.push(logEntry("Resolving target SharePoint site..."));
    let targetSite: any;
    try {
      targetSite = await target.get(`/sites/${targetIdentity}`);
    } catch {
      const searchResult = await target.get(`/sites?search=${encodeURIComponent(targetIdentity)}`);
      if (!searchResult.value || searchResult.value.length === 0) {
        throw new Error(`Target SharePoint site "${targetIdentity}" not found`);
      }
      targetSite = searchResult.value[0];
    }
    logs.push(logEntry(`Found target site: ${targetSite.displayName} (${targetSite.id})`));
    await updateItemProgress(itemId, 'in_progress', logs);

    const sourceDrives = await source.getAllPages<any>(`/sites/${sourceSite.id}/drives`);
    logs.push(logEntry(`Found ${sourceDrives.length} document libraries in source site`));

    const targetDrives = await target.getAllPages<any>(`/sites/${targetSite.id}/drives`);

    const counters = { migrated: 0, failed: 0, total: 0 };

    for (const drive of sourceDrives) {
      logs.push(logEntry(`Processing document library: ${drive.name}`));
      await updateItemProgress(itemId, 'in_progress', logs);

      const matchingDrive = targetDrives.find((d: any) => d.name === drive.name);
      const targetDriveId = matchingDrive?.id || targetDrives[0]?.id;

      if (!targetDriveId) {
        logs.push(logEntry(`No document libraries found in target site, skipping`));
        continue;
      }

      if (!matchingDrive) {
        logs.push(logEntry(`No matching library "${drive.name}" in target — using default library`));
      }

      try {
        await migrateDriveItemsRecursive(
          source, target,
          drive.id, targetDriveId,
          null, "",
          itemId, logs, counters
        );
      } catch (err: any) {
        logs.push(logEntry(`Failed to migrate library "${drive.name}": ${err.message}`));
      }
    }

    logs.push(logEntry(`SharePoint migration complete: ${counters.migrated} migrated, ${counters.failed} failed out of ${counters.total} total`));

    if (counters.failed > 0 && counters.migrated === 0) {
      await updateItemProgress(itemId, 'failed', logs, `All ${counters.failed} items failed`);
    } else if (counters.failed > 0) {
      await updateItemProgress(itemId, 'completed', logs, `${counters.failed} items failed`);
    } else {
      await updateItemProgress(itemId, 'completed', logs);
    }
  } catch (err: any) {
    logs.push(logEntry(`SharePoint migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

export async function migrateItem(projectId: number, itemId: number): Promise<void> {
  const project = await storage.getProject(projectId);
  if (!project) throw new Error("Project not found");

  const item = await storage.getItem(itemId);
  if (!item || item.projectId !== projectId) throw new Error("Migration item not found");

  if (item.status === 'in_progress') {
    throw new Error("Migration already in progress for this item");
  }

  await storage.updateItem(itemId, { status: 'in_progress', errorDetails: null });

  try {
    const { source, target } = getGraphClients(project);
    const targetIdentity = item.targetIdentity || item.sourceIdentity;

    switch (item.itemType) {
      case 'mailbox':
        await migrateMailbox(source, target, item.sourceIdentity, targetIdentity, itemId);
        break;
      case 'onedrive':
        await migrateOneDrive(source, target, item.sourceIdentity, targetIdentity, itemId);
        break;
      case 'sharepoint':
        await migrateSharePoint(source, target, item.sourceIdentity, targetIdentity, itemId);
        break;
      case 'teams':
        await storage.updateItem(itemId, {
          status: 'failed',
          errorDetails: 'Teams migration is not yet supported. Microsoft Graph API does not provide endpoints to migrate Teams channel messages and data between tenants.',
        });
        await storage.updateItemLogs(itemId, [
          logEntry('Teams migration is not supported at this time.'),
          logEntry('Microsoft Graph API does not expose endpoints for copying Teams messages, channel data, or team structures between tenants.'),
          logEntry('Consider using Microsoft\'s native Teams migration tools or third-party solutions for Teams data.'),
        ]);
        break;
      default:
        await storage.updateItem(itemId, {
          status: 'failed',
          errorDetails: `Unsupported item type: ${item.itemType}`,
        });
        break;
    }
  } catch (err: any) {
    await storage.updateItem(itemId, {
      status: 'failed',
      errorDetails: err.message,
    });
    await storage.updateItemLogs(itemId, [logEntry(`Migration failed: ${err.message}`)]);
  }
}

export async function migrateAllPending(projectId: number): Promise<{ started: number; skipped: number; errors: string[] }> {
  const project = await storage.getProject(projectId);
  if (!project) throw new Error("Project not found");

  getGraphClients(project);

  const items = await storage.getItems(projectId);
  const pending = items.filter(i => i.status === 'pending' || i.status === 'failed');

  const errors: string[] = [];
  let started = 0;
  let skipped = 0;

  for (const item of pending) {
    if (item.status === 'in_progress') {
      skipped++;
      continue;
    }

    try {
      migrateItem(projectId, item.id).catch(err => {
        console.error(`Migration failed for item ${item.id}:`, err.message);
      });
      started++;
    } catch (err: any) {
      errors.push(`Item ${item.id}: ${err.message}`);
    }
  }

  return { started, skipped, errors };
}
