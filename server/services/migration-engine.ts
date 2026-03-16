import { GraphClient } from "./graph-client";
import { storage } from "../storage";
import type { Project, MigrationItem } from "@shared/schema";
import { readMailboxDelegates, applyMailboxDelegates, createSharedMailboxDirect, type ExoConfig } from "./exo-runner";

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
    // ── Verify target user exists before doing anything ───────────────────
    logs.push(logEntry(`Verifying target user exists: ${targetUser}`));
    await updateItemProgress(itemId, 'in_progress', logs);
    const targetUserObj = await target.get(`/users/${encodeURIComponent(targetUser)}?$select=id,userPrincipalName,assignedLicenses`).catch(() => null);
    if (!targetUserObj) {
      throw new Error(
        `Target user "${targetUser}" does not exist in the target tenant. ` +
        `Create the user account first (via the Users migration tab or manually in the target tenant's admin centre), then re-run this mailbox migration. ` +
        `Alternatively, verify the "Target identity" field on this item is set to the correct UPN.`
      );
    }
    const hasLicense = Array.isArray(targetUserObj.assignedLicenses) && targetUserObj.assignedLicenses.length > 0;
    if (!hasLicense) {
      logs.push(logEntry(`⚠ Target user has no Microsoft 365 license assigned. Mailbox migration requires a license that includes Exchange Online (e.g. M365 Business Basic/Standard/Premium, E1/E3/E5). Assign a license in the target admin centre, wait ~5 minutes for the mailbox to provision, then re-run.`));
    }
    logs.push(logEntry(`✓ Target user found: ${targetUserObj.userPrincipalName}`));

    // ── Fetch source folders ──────────────────────────────────────────────
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
        const targetFolders = await target.get(`/users/${targetUser}/mailFolders?$filter=displayName eq '${folderName.replace(/'/g, "''")}'`);
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
        // Fall back to well-known Inbox — note the error but continue
        logs.push(logEntry(`Could not access/create folder "${folderName}" in target — placing messages in Inbox. Error: ${err.message}`));
        try {
          const inbox = await target.get(`/users/${targetUser}/mailFolders/Inbox`);
          targetFolderId = inbox.id;
        } catch {
          logs.push(logEntry(`Cannot access target mailbox at all — skipping folder "${folderName}". Check Mail.ReadWrite Application permission on the TARGET app registration.`));
          continue;
        }
      }

      try {
        const messages = await source.getAllPages<any>(
          `/users/${sourceUser}/mailFolders/${folder.id}/messages?$top=50&$select=id,subject,body,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,importance,isRead,hasAttachments,flag`
        );
        totalMessages += messages.length;
        logs.push(logEntry(`Found ${messages.length} messages in "${folderName}"`));
        await updateItemProgress(itemId, 'in_progress', logs);

        for (const msg of messages) {
          try {
            const newMessage: any = {
              subject: msg.subject || "(No Subject)",
              body: msg.body,
              toRecipients: msg.toRecipients || [],
              ccRecipients: msg.ccRecipients || [],
              bccRecipients: msg.bccRecipients || [],
              importance: msg.importance || "normal",
              isRead: msg.isRead !== undefined ? msg.isRead : true,
              isDraft: false,
              flag: msg.flag || { flagStatus: 'notFlagged' },
            };
            // Note: receivedDateTime and internetMessageId are read-only in Graph API
            // and cannot be set when creating messages. The copy will show today's date.

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
    logs.push(logEntry(`Messages complete: ${migratedMessages} migrated, ${failedMessages} failed out of ${totalMessages}${attachmentSummary}`));
    await updateItemProgress(itemId, 'in_progress', logs);

    // ── CALENDAR EVENTS ──────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating calendar events...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const calendars = await source.getAllPages<any>(`/users/${sourceUser}/calendars`);
      let calMigrated = 0, calFailed = 0, calTotal = 0;
      for (const cal of calendars) {
        let targetCalId: string;
        try {
          const tCals = await target.get(`/users/${targetUser}/calendars?$filter=name eq '${cal.name.replace(/'/g, "''")}'`);
          targetCalId = tCals.value?.length > 0
            ? tCals.value[0].id
            : (await target.post(`/users/${targetUser}/calendars`, { name: cal.name })).id;
        } catch {
          targetCalId = (await target.get(`/users/${targetUser}/calendar`)).id;
        }
        const events = await source.getAllPages<any>(`/users/${sourceUser}/calendars/${cal.id}/events?$top=50`);
        calTotal += events.length;
        for (const ev of events) {
          try {
            const payload: any = {
              subject: ev.subject, body: ev.body,
              start: ev.start, end: ev.end,
              location: ev.location, isAllDay: ev.isAllDay || false,
              showAs: ev.showAs, sensitivity: ev.sensitivity,
              importance: ev.importance, isReminderOn: ev.isReminderOn,
              reminderMinutesBeforeStart: ev.reminderMinutesBeforeStart,
            };
            if (ev.attendees?.length) payload.attendees = ev.attendees;
            if (ev.recurrence) payload.recurrence = ev.recurrence;
            await target.post(`/users/${targetUser}/calendars/${targetCalId}/events`, payload);
            calMigrated++;
          } catch { calFailed++; }
        }
      }
      logs.push(logEntry(`Calendar: ${calMigrated} events migrated, ${calFailed} failed out of ${calTotal}`));
    } catch (e: any) {
      const is403 = e.message?.includes('403') || e.message?.toLowerCase().includes('access is denied');
      logs.push(logEntry(`Calendar migration skipped: ${e.message}${is403 ? ' — Grant Calendars.ReadWrite (Application) on BOTH tenant app registrations and re-run admin consent.' : ''}`));
    }

    // ── CONTACTS ─────────────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating contacts...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const contacts = await source.getAllPages<any>(`/users/${sourceUser}/contacts?$top=100`);
      let cntMigrated = 0, cntFailed = 0;
      for (const c of contacts) {
        try {
          const payload: any = { displayName: c.displayName };
          if (c.givenName) payload.givenName = c.givenName;
          if (c.surname) payload.surname = c.surname;
          if (c.emailAddresses?.length) payload.emailAddresses = c.emailAddresses;
          if (c.businessPhones?.length) payload.businessPhones = c.businessPhones;
          if (c.mobilePhone) payload.mobilePhone = c.mobilePhone;
          if (c.jobTitle) payload.jobTitle = c.jobTitle;
          if (c.companyName) payload.companyName = c.companyName;
          if (c.department) payload.department = c.department;
          if (c.officeLocation) payload.officeLocation = c.officeLocation;
          if (c.businessAddress) payload.businessAddress = c.businessAddress;
          if (c.birthday) payload.birthday = c.birthday;
          if (c.personalNotes) payload.personalNotes = c.personalNotes;
          await target.post(`/users/${targetUser}/contacts`, payload);
          cntMigrated++;
        } catch { cntFailed++; }
      }
      logs.push(logEntry(`Contacts: ${cntMigrated} migrated, ${cntFailed} failed out of ${contacts.length}`));
    } catch (e: any) {
      const is403 = e.message?.includes('403') || e.message?.toLowerCase().includes('access is denied');
      logs.push(logEntry(`Contacts migration skipped: ${e.message}${is403 ? ' — Grant Contacts.ReadWrite (Application) on BOTH tenant app registrations and re-run admin consent.' : ''}`));
    }

    // ── TASKS (Microsoft To Do) ───────────────────────────────────────────────
    logs.push(logEntry('── Migrating tasks (Microsoft To Do)...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const taskLists = await source.getAllPages<any>(`/users/${sourceUser}/todo/lists`);
      let taskMigrated = 0, taskFailed = 0, taskTotal = 0;
      for (const list of taskLists) {
        let targetListId: string;
        try {
          const escaped = list.displayName.replace(/'/g, "''");
          const tLists = await target.get(`/users/${targetUser}/todo/lists?$filter=displayName eq '${escaped}'`);
          targetListId = tLists.value?.length > 0
            ? tLists.value[0].id
            : (await target.post(`/users/${targetUser}/todo/lists`, { displayName: list.displayName })).id;
        } catch { continue; }
        const tasks = await source.getAllPages<any>(`/users/${sourceUser}/todo/lists/${list.id}/tasks`);
        taskTotal += tasks.length;
        for (const task of tasks) {
          try {
            const payload: any = {
              title: task.title,
              importance: task.importance || 'normal',
              status: task.status || 'notStarted',
            };
            if (task.body) payload.body = task.body;
            if (task.dueDateTime) payload.dueDateTime = task.dueDateTime;
            if (task.reminderDateTime) payload.reminderDateTime = task.reminderDateTime;
            if (task.completedDateTime) payload.completedDateTime = task.completedDateTime;
            await target.post(`/users/${targetUser}/todo/lists/${targetListId}/tasks`, payload);
            taskMigrated++;
          } catch { taskFailed++; }
        }
      }
      logs.push(logEntry(`Tasks: ${taskMigrated} migrated, ${taskFailed} failed out of ${taskTotal}`));
    } catch (e: any) {
      const is401or403 = e.message?.includes('401') || e.message?.includes('403') || e.message?.toLowerCase().includes('access is denied');
      logs.push(logEntry(`Tasks migration skipped: ${e.message}${is401or403 ? ' — Grant Tasks.ReadWrite (Application) on BOTH tenant app registrations and re-run admin consent.' : ''}`));
    }

    // ── MAILBOX RULES ─────────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating mailbox rules...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const rulesRes = await source.get(`/users/${sourceUser}/mailFolders/inbox/messageRules`);
      let rulesMigrated = 0, rulesFailed = 0;
      for (const rule of rulesRes.value || []) {
        try {
          await target.post(`/users/${targetUser}/mailFolders/inbox/messageRules`, {
            displayName: rule.displayName,
            sequence: rule.sequence,
            isEnabled: rule.isEnabled,
            conditions: rule.conditions,
            actions: rule.actions,
            exceptions: rule.exceptions,
          });
          rulesMigrated++;
        } catch { rulesFailed++; }
      }
      logs.push(logEntry(`Mailbox rules: ${rulesMigrated} migrated, ${rulesFailed} failed`));
    } catch (e: any) {
      const is403 = e.message?.includes('403') || e.message?.toLowerCase().includes('access is denied');
      logs.push(logEntry(`Mailbox rules migration skipped: ${e.message}${is403 ? ' — Grant MailboxSettings.ReadWrite (Application) on BOTH tenant app registrations and re-run admin consent.' : ''}`));
    }

    // ── OUT-OF-OFFICE / MAILBOX SETTINGS ────────────────────────────────────
    logs.push(logEntry('── Migrating mailbox settings (OOO, timezone, working hours)...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const mbSettings = await source.get(`/users/${sourceUser}/mailboxSettings`);
      const settingsPayload: any = {};
      if (mbSettings.automaticRepliesSetting) settingsPayload.automaticRepliesSetting = mbSettings.automaticRepliesSetting;
      if (mbSettings.timeZone) settingsPayload.timeZone = mbSettings.timeZone;
      if (mbSettings.workingHours) settingsPayload.workingHours = mbSettings.workingHours;
      if (mbSettings.language) settingsPayload.language = mbSettings.language;
      if (mbSettings.dateFormat) settingsPayload.dateFormat = mbSettings.dateFormat;
      if (mbSettings.timeFormat) settingsPayload.timeFormat = mbSettings.timeFormat;
      if (Object.keys(settingsPayload).length > 0) {
        await target.patch(`/users/${targetUser}/mailboxSettings`, settingsPayload);
        logs.push(logEntry('Mailbox settings migrated (auto-reply, timezone, working hours, language)'));
      }
    } catch (e: any) {
      const is403 = e.message?.includes('403') || e.message?.toLowerCase().includes('access is denied');
      logs.push(logEntry(`Mailbox settings migration skipped: ${e.message}${is403 ? ' — Grant MailboxSettings.ReadWrite (Application) on BOTH tenant app registrations and re-run admin consent.' : ''}`));
    }

    // ── CATEGORIES ───────────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating Outlook categories...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const cats = await source.get(`/users/${sourceUser}/outlook/masterCategories`);
      let catMigrated = 0;
      for (const cat of cats.value || []) {
        try {
          await target.post(`/users/${targetUser}/outlook/masterCategories`, {
            displayName: cat.displayName,
            color: cat.color,
          });
          catMigrated++;
        } catch { /* may already exist in target */ }
      }
      if (catMigrated > 0) logs.push(logEntry(`Categories: ${catMigrated} migrated`));
      else logs.push(logEntry('Categories: none to migrate or already exist'));
    } catch (e: any) {
      const is403 = e.message?.includes('403') || e.message?.toLowerCase().includes('access is denied');
      logs.push(logEntry(`Categories migration skipped: ${e.message}${is403 ? ' — Grant MailboxSettings.ReadWrite (Application) on BOTH tenant app registrations and re-run admin consent.' : ''}`));
    }

    // ── API LIMITATIONS NOTE ────────────────────────────────────────────────
    logs.push(logEntry('Note: Email signatures and full mailbox delegate permissions cannot be migrated — Microsoft Graph API does not expose these endpoints.'));

    logs.push(logEntry('✓ Mailbox migration complete'));
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

        // ── Preserve original file timestamps ────────────────────────────
        try {
          if (item.fileSystemInfo?.lastModifiedDateTime || item.fileSystemInfo?.createdDateTime) {
            const fsPayload: any = { fileSystemInfo: {} };
            if (item.fileSystemInfo.createdDateTime)      fsPayload.fileSystemInfo.createdDateTime      = item.fileSystemInfo.createdDateTime;
            if (item.fileSystemInfo.lastModifiedDateTime) fsPayload.fileSystemInfo.lastModifiedDateTime = item.fileSystemInfo.lastModifiedDateTime;
            await target.patch(`/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:`, fsPayload);
          }
        } catch { /* non-critical — timestamps are best-effort */ }

        // ── Copy item-level permissions ────────────────────────────────────
        try {
          const permRes = await source.get(`/drives/${sourceDriveId}/items/${item.id}/permissions`);
          for (const perm of permRes.value || []) {
            if (perm.inheritedFrom) continue; // skip inherited — parent folder handles them
            try {
              if (perm.link) {
                // Recreate sharing link with same role/scope
                await target.post(`/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:/createLink`, {
                  type: perm.link.type,
                  scope: perm.link.scope,
                });
              } else if (perm.grantedToV2?.user?.email) {
                // Direct user permission
                await target.post(`/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:/invite`, {
                  requireSignIn: true,
                  sendInvitation: false,
                  roles: perm.roles || ['read'],
                  recipients: [{ email: perm.grantedToV2.user.email }],
                });
              } else if (perm.grantedToIdentitiesV2?.length) {
                // Group/multiple identities
                for (const identity of perm.grantedToIdentitiesV2) {
                  if (identity.user?.email) {
                    await target.post(`/drives/${targetDriveId}/root:${targetParentPath}/${item.name}:/invite`, {
                      requireSignIn: true,
                      sendInvitation: false,
                      roles: perm.roles || ['read'],
                      recipients: [{ email: identity.user.email }],
                    });
                  }
                }
              }
            } catch { /* individual permission copy is best-effort */ }
          }
        } catch { /* permissions are best-effort */ }

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

  // Convert a full https:// URL or a bare path to the Graph API site path format:
  // "https://tenant.sharepoint.com/sites/Team" → "tenant.sharepoint.com:/sites/Team"
  // "tenant.sharepoint.com:/sites/Team"        → unchanged
  function toGraphSitePath(identity: string): string {
    try {
      if (identity.startsWith('http://') || identity.startsWith('https://')) {
        const u = new URL(identity);
        return `${u.hostname}:${u.pathname}`;
      }
    } catch { /* fall through to returning as-is */ }
    return identity;
  }

  async function resolveSite(client: GraphClient, identity: string): Promise<any | null> {
    // Try 1: hostname:/path format (canonical Graph API format for full URLs)
    const graphPath = toGraphSitePath(identity);
    if (graphPath.includes(':')) {
      try { return await client.get(`/sites/${graphPath}`); } catch { /* try next */ }
    }

    // Try 2: search by keyword/display name
    try {
      const keyword = graphPath.split('/').pop()?.split(':').pop() || identity;
      const searchResult = await client.get(`/sites?search=${encodeURIComponent(keyword)}`);
      if (searchResult.value?.length > 0) return searchResult.value[0];
    } catch { /* try next */ }

    return null;
  }

  async function resolveOrCreateTargetSite(displayName: string): Promise<any> {
    // Try finding existing site by display name
    const existing = await resolveSite(target, displayName);
    if (existing) return existing;

    logs.push(logEntry(`Target site "${displayName}" not found — creating new Team site in target tenant...`));
    await updateItemProgress(itemId, 'in_progress', logs);

    // Create via M365 Group, which automatically provisions a SharePoint Team site
    const mailNickname = displayName.toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 59) || 'migrated';
    const group = await target.post('/groups', {
      displayName,
      mailNickname,
      groupTypes: ['Unified'],
      mailEnabled: true,
      securityEnabled: false,
      visibility: 'Private',
    });

    logs.push(logEntry(`M365 Group created (id: ${group.id}) — waiting for SharePoint site to provision...`));
    await updateItemProgress(itemId, 'in_progress', logs);

    // Poll for site to be provisioned (can take up to 30s)
    for (let i = 0; i < 12; i++) {
      await new Promise(r => setTimeout(r, 5000));
      try {
        const site = await target.get(`/groups/${group.id}/sites/root`);
        if (site?.id) {
          logs.push(logEntry(`✓ Target site provisioned: ${site.webUrl}`));
          return site;
        }
      } catch { /* not ready yet */ }
    }
    throw new Error(`Timed out waiting for SharePoint site to provision for group "${displayName}"`);
  }

  try {
    logs.push(logEntry("Resolving source SharePoint site..."));
    await updateItemProgress(itemId, 'in_progress', logs);

    const sourceSite = await resolveSite(source, sourceIdentity);
    if (!sourceSite) {
      throw new Error(
        `Source SharePoint site "${sourceIdentity}" not found. ` +
        `Ensure the site exists and the app has Sites.ReadWrite.All permission.`
      );
    }
    logs.push(logEntry(`Found source site: ${sourceSite.displayName} (${sourceSite.id})`));

    logs.push(logEntry("Resolving target SharePoint site..."));
    // targetIdentity is the display name; fall back to source display name if blank
    const targetName = targetIdentity || sourceSite.displayName;
    const targetSite = await resolveOrCreateTargetSite(targetName);
    logs.push(logEntry(`Found target site: ${targetSite.displayName} (${targetSite.id})`));
    await updateItemProgress(itemId, 'in_progress', logs);

    const sourceDrives = await source.getAllPages<any>(`/sites/${sourceSite.id}/drives`);
    logs.push(logEntry(`Found ${sourceDrives.length} document libraries in source site`));

    const targetDrives = await target.getAllPages<any>(`/sites/${targetSite.id}/drives`);

    const counters = { migrated: 0, failed: 0, total: 0, bytesMigrated: 0, bytesTotal: 0 };

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

    logs.push(logEntry(`Document libraries: ${counters.migrated} files migrated, ${counters.failed} failed out of ${counters.total} total`));
    await updateItemProgress(itemId, 'in_progress', logs);

    // ── NON-LIBRARY LISTS ────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating SharePoint lists (non-library)...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const allLists = await source.getAllPages<any>(`/sites/${sourceSite.id}/lists?$expand=columns`);
      const customLists = allLists.filter((l: any) =>
        l.list?.template !== 'documentLibrary' &&
        l.list?.template !== 'webPageLibrary' &&
        !l.name?.startsWith('_') &&
        !['appdata', 'appfiles', 'Composed Looks', 'Master Page Gallery', 'Solution Gallery', 'Theme Gallery', 'User Information List', 'Web Part Gallery'].includes(l.name)
      );
      let listsMigrated = 0, listsFailed = 0;
      for (const list of customLists) {
        try {
          // Create target list
          let targetList: any;
          try {
            targetList = await target.post(`/sites/${targetSite.id}/lists`, {
              displayName: list.displayName,
              list: { template: list.list?.template || 'genericList' },
            });
          } catch {
            // May already exist — try to find it
            const existing = await target.get(`/sites/${targetSite.id}/lists?$filter=displayName eq '${list.displayName.replace(/'/g, "''")}'`);
            if (existing.value?.length > 0) targetList = existing.value[0];
            else throw new Error(`Could not create or find list "${list.displayName}"`);
          }

          // Create custom columns (skip built-in ones)
          const builtIn = new Set(['Title', 'ID', 'Created', 'Modified', 'Author', 'Editor', 'ContentType', 'Attachments', '_UIVersionString', 'Edit', 'DocIcon', 'LinkTitleNoMenu', 'LinkTitle', 'ItemChildCount', 'FolderChildCount', 'AppAuthor', 'AppEditor']);
          for (const col of (list.columns || [])) {
            if (builtIn.has(col.name) || col.readOnly || col.hidden) continue;
            try {
              await target.post(`/sites/${targetSite.id}/lists/${targetList.id}/columns`, {
                name: col.name,
                displayName: col.displayName,
                description: col.description || '',
                [col.text ? 'text' : col.number ? 'number' : col.boolean ? 'boolean' : col.dateTime ? 'dateTime' : col.choice ? 'choice' : 'text']: col.text || col.number || col.boolean || col.dateTime || col.choice || {},
              });
            } catch { /* column may already exist or be invalid */ }
          }

          // Copy list items
          const items = await source.getAllPages<any>(`/sites/${sourceSite.id}/lists/${list.id}/items?$expand=fields`);
          let itemsMigrated = 0;
          for (const item of items) {
            try {
              const fields: any = {};
              for (const [key, val] of Object.entries(item.fields || {})) {
                if (['@odata.etag', 'id', 'ID', 'Created', 'Modified', 'AuthorLookupId', 'EditorLookupId', 'ContentType', 'Attachments', '_UIVersionString', 'Edit', 'DocIcon', 'LinkTitleNoMenu', 'LinkTitle', 'ItemChildCount', 'FolderChildCount', 'AppAuthor', 'AppEditor'].includes(key)) continue;
                if (!key.startsWith('@') && !key.startsWith('_') && val !== null && val !== undefined) {
                  fields[key] = val;
                }
              }
              await target.post(`/sites/${targetSite.id}/lists/${targetList.id}/items`, { fields });
              itemsMigrated++;
            } catch { /* item copy is best-effort */ }
          }
          logs.push(logEntry(`✓ List "${list.displayName}": ${itemsMigrated}/${items.length} items migrated`));
          listsMigrated++;
        } catch (e: any) {
          listsFailed++;
          logs.push(logEntry(`List "${list.displayName}" failed: ${e.message}`));
        }
      }
      logs.push(logEntry(`Lists: ${listsMigrated} lists migrated, ${listsFailed} failed`));
    } catch (e: any) {
      logs.push(logEntry(`Lists migration skipped: ${e.message}`));
    }

    // ── SITE PAGES ───────────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating site pages...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      // Pages API is in beta
      const pagesRes = await source.get(`https://graph.microsoft.com/beta/sites/${sourceSite.id}/pages`);
      const pages = pagesRes.value || [];
      let pagesMigrated = 0, pagesFailed = 0;
      for (const page of pages) {
        try {
          // Fetch full page content
          const fullPage = await source.get(`https://graph.microsoft.com/beta/sites/${sourceSite.id}/pages/${page.id}?$expand=canvasLayout`);
          const pagePayload: any = {
            name: fullPage.name,
            title: fullPage.title,
            pageLayout: fullPage.pageLayout || 'article',
            showComments: fullPage.showComments ?? true,
            showRecommendedPages: fullPage.showRecommendedPages ?? false,
          };
          if (fullPage.canvasLayout) pagePayload.canvasLayout = fullPage.canvasLayout;
          if (fullPage.titleArea) pagePayload.titleArea = fullPage.titleArea;
          await target.post(`https://graph.microsoft.com/beta/sites/${targetSite.id}/pages`, pagePayload);
          pagesMigrated++;
        } catch { pagesFailed++; }
      }
      logs.push(logEntry(`Pages: ${pagesMigrated} migrated, ${pagesFailed} failed out of ${pages.length}`));
    } catch (e: any) {
      logs.push(logEntry(`Pages migration skipped: ${e.message}`));
    }

    // ── SITE PERMISSIONS ─────────────────────────────────────────────────────
    logs.push(logEntry('── Migrating site permissions...'));
    await updateItemProgress(itemId, 'in_progress', logs);
    try {
      const permsRes = await source.getAllPages<any>(`/sites/${sourceSite.id}/permissions`);
      let permsMigrated = 0, permsFailed = 0;
      for (const perm of permsRes) {
        try {
          if (perm.grantedToIdentities?.length) {
            await target.post(`/sites/${targetSite.id}/permissions`, {
              roles: perm.roles,
              grantedToIdentities: perm.grantedToIdentities,
            });
            permsMigrated++;
          }
        } catch { permsFailed++; }
      }
      logs.push(logEntry(`Site permissions: ${permsMigrated} migrated, ${permsFailed} failed`));
    } catch (e: any) {
      logs.push(logEntry(`Site permissions migration skipped: ${e.message}`));
    }

    // ── API LIMITATIONS NOTE ────────────────────────────────────────────────
    logs.push(logEntry('Note: SharePoint site branding, navigation, and classic/Power Automate workflows cannot be migrated via Microsoft Graph API.'));
    logs.push(logEntry('Note: File version history cannot be migrated — Microsoft Graph API does not support creating historical versions.'));

    logs.push(logEntry('✓ SharePoint migration complete'));
    if (counters.failed > 0 && counters.migrated === 0) {
      await updateItemProgress(itemId, 'failed', logs, `All ${counters.failed} file items failed`);
    } else if (counters.failed > 0) {
      await updateItemProgress(itemId, 'completed', logs, `${counters.failed} file items failed`);
    } else {
      await updateItemProgress(itemId, 'completed', logs);
    }
  } catch (err: any) {
    logs.push(logEntry(`SharePoint migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

async function migrateUser(
  source: GraphClient,
  target: GraphClient,
  sourceIdentity: string,
  targetIdentity: string,
  itemId: number
): Promise<void> {
  const logs: string[] = [];
  logs.push(logEntry(`Starting user migration: ${sourceIdentity} → ${targetIdentity}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    logs.push(logEntry(`Looking up source user: ${sourceIdentity}`));
    const sourceUser = await source.get(
      `/users/${sourceIdentity}?$select=id,displayName,givenName,surname,jobTitle,department,usageLocation,mobilePhone,businessPhones,officeLocation`
    );

    const mailNickname = targetIdentity.split('@')[0].replace(/[^a-zA-Z0-9]/g, '');
    const tempPassword = `Migr@tion${Math.random().toString(36).slice(-6).toUpperCase()}1!`;

    logs.push(logEntry(`Creating user in target tenant: ${targetIdentity}`));

    // Graph API rejects empty string for givenName/surname — omit the field entirely if blank
    const userPayload: Record<string, any> = {
      accountEnabled: true,
      displayName: sourceUser.displayName || targetIdentity.split('@')[0],
      mailNickname,
      userPrincipalName: targetIdentity,
      usageLocation: sourceUser.usageLocation || 'US',
      passwordProfile: {
        forceChangePasswordNextSignIn: true,
        password: tempPassword,
      },
    };
    if (sourceUser.givenName) userPayload.givenName = sourceUser.givenName;
    if (sourceUser.surname) userPayload.surname = sourceUser.surname;
    if (sourceUser.jobTitle) userPayload.jobTitle = sourceUser.jobTitle;
    if (sourceUser.department) userPayload.department = sourceUser.department;
    if (sourceUser.mobilePhone) userPayload.mobilePhone = sourceUser.mobilePhone;
    if (sourceUser.officeLocation) userPayload.officeLocation = sourceUser.officeLocation;

    const newUser = await target.post('/users', userPayload);

    logs.push(logEntry(`✓ User created: ${newUser.userPrincipalName} (ID: ${newUser.id})`));
    logs.push(logEntry(`Temporary password: ${tempPassword} (user must change on first login)`));
    logs.push(logEntry(`User migration complete. Remember to assign licences in the target tenant.`));
    await updateItemProgress(itemId, 'completed', logs);
  } catch (err: any) {
    const msg = err.message || String(err);
    if (msg.includes('409') || msg.toLowerCase().includes('already exists')) {
      logs.push(logEntry(`User ${targetIdentity} already exists in target tenant — skipping creation.`));
      await updateItemProgress(itemId, 'completed', logs);
    } else {
      logs.push(logEntry(`User migration failed: ${msg}`));
      await updateItemProgress(itemId, 'failed', logs, msg);
    }
  }
}

async function pollTeamProvisioning(target: GraphClient, operationUrl: string, logs: string[], maxWaitMs = 120000): Promise<string> {
  const start = Date.now();
  while (Date.now() - start < maxWaitMs) {
    await new Promise(r => setTimeout(r, 5000));
    try {
      const status = await target.get(operationUrl);
      if (status.status === 'succeeded' && status.targetResourceId) {
        return status.targetResourceId;
      }
      if (status.status === 'failed') {
        throw new Error(`Team provisioning failed: ${status.error?.message || 'Unknown error'}`);
      }
      logs.push(logEntry(`Team provisioning status: ${status.status}...`));
    } catch (e: any) {
      if (e.message?.includes('provisioning failed')) throw e;
      // ignore transient errors
    }
  }
  throw new Error('Team provisioning timed out after 2 minutes');
}

async function migrateTeam(
  source: GraphClient,
  target: GraphClient,
  sourceIdentity: string,
  targetIdentity: string,
  itemId: number
): Promise<void> {
  const logs: string[] = [];
  logs.push(logEntry(`Starting Teams migration: ${sourceIdentity} → ${targetIdentity}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    // Resolve source team by display name or ID
    logs.push(logEntry('Resolving source team...'));
    let sourceTeam: any;
    try {
      sourceTeam = await source.get(`/teams/${sourceIdentity}`);
    } catch {
      const search = await source.get(`/groups?$filter=displayName eq '${sourceIdentity}' and resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description,visibility`);
      if (!search.value || search.value.length === 0) throw new Error(`Team "${sourceIdentity}" not found in source tenant`);
      sourceTeam = search.value[0];
    }

    const teamName = targetIdentity || sourceTeam.displayName;
    logs.push(logEntry(`Found source team: ${sourceTeam.displayName} (${sourceTeam.id})`));

    // Check if target team already exists
    let targetTeamId: string | null = null;
    logs.push(logEntry(`Checking if team "${teamName}" exists in target...`));
    try {
      const existing = await target.get(`/groups?$filter=displayName eq '${teamName}' and resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id`);
      if (existing.value && existing.value.length > 0) {
        targetTeamId = existing.value[0].id;
        logs.push(logEntry(`Team already exists in target (ID: ${targetTeamId}) — will migrate content into it.`));
      }
    } catch { }

    if (!targetTeamId) {
      logs.push(logEntry(`Creating team "${teamName}" in target tenant...`));
      const res = await target.request('/teams', {
        method: 'POST',
        body: JSON.stringify({
          'template@odata.bind': "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
          displayName: teamName,
          description: sourceTeam.description || '',
          visibility: sourceTeam.visibility || 'Private',
        }),
      });

      if (res.status === 202) {
        const location = res.headers.get('Location') || '';
        const operationPath = location.replace('https://graph.microsoft.com/v1.0', '');
        logs.push(logEntry('Team creation accepted — waiting for provisioning (up to 2 min)...'));
        await updateItemProgress(itemId, 'in_progress', logs);
        targetTeamId = await pollTeamProvisioning(target, operationPath, logs);
        logs.push(logEntry(`✓ Team provisioned (ID: ${targetTeamId})`));
      } else if (res.ok) {
        const data = await res.json();
        targetTeamId = data.id;
      } else {
        const err = await res.text();
        throw new Error(`Failed to create team: ${err}`);
      }
    }

    // Get source channels
    logs.push(logEntry('Reading source channels...'));
    const sourceChannels = await source.getAllPages<any>(`/teams/${sourceTeam.id}/channels`);
    logs.push(logEntry(`Found ${sourceChannels.length} channels in source team`));

    const targetChannels = await target.getAllPages<any>(`/teams/${targetTeamId}/channels`);
    const existingChannelNames = new Set(targetChannels.map((c: any) => c.displayName.toLowerCase()));

    const counters = { migrated: 0, failed: 0, total: 0, bytesMigrated: 0, bytesTotal: 0 };

    for (const channel of sourceChannels) {
      if (channel.membershipType === 'standard' && channel.displayName.toLowerCase() === 'general') {
        logs.push(logEntry(`Skipping "General" channel creation (auto-exists).`));
      } else if (!existingChannelNames.has(channel.displayName.toLowerCase())) {
        try {
          await target.post(`/teams/${targetTeamId}/channels`, {
            displayName: channel.displayName,
            description: channel.description || '',
            membershipType: 'standard',
          });
          logs.push(logEntry(`✓ Created channel: ${channel.displayName}`));
        } catch (err: any) {
          logs.push(logEntry(`Could not create channel "${channel.displayName}": ${err.message}`));
        }
      }

      // Migrate channel files via SharePoint
      try {
        const filesFolder = await source.get(`/teams/${sourceTeam.id}/channels/${channel.id}/filesFolder`);
        const targetFilesFolder = await target.get(`/teams/${targetTeamId}/channels/${channel.id}/filesFolder`).catch(() => null);

        if (filesFolder?.parentReference?.driveId && targetFilesFolder?.parentReference?.driveId) {
          logs.push(logEntry(`Migrating files for channel: ${channel.displayName}`));
          await updateItemProgress(itemId, 'in_progress', logs, undefined, counters);
          await migrateDriveItemsRecursive(
            source, target,
            filesFolder.parentReference.driveId, targetFilesFolder.parentReference.driveId,
            filesFolder.id, `/${channel.displayName}`,
            itemId, logs, counters
          );
        }
      } catch (err: any) {
        logs.push(logEntry(`Could not migrate files for "${channel.displayName}": ${err.message}`));
      }
    }

    logs.push(logEntry(`Teams migration complete: ${counters.migrated} files migrated (${formatBytes(counters.bytesMigrated)}), ${counters.failed} failed`));
    await updateItemProgress(itemId, 'completed', logs, undefined, counters);
  } catch (err: any) {
    logs.push(logEntry(`Teams migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

// ── Helper: resolve a user in a tenant by UPN, email, or alias (cross-tenant aware) ──
async function resolveUserInTenant(client: GraphClient, upn: string): Promise<string | null> {
  if (!upn) return null;

  // 1. Try exact UPN / object ID lookup
  try {
    const u = await client.get(`/users/${encodeURIComponent(upn)}?$select=id`);
    if (u?.id) return u.id;
  } catch { }

  // 2. Try filter by mail address (handles alias@differentdomain.com)
  try {
    const res = await client.get(`/users?$filter=mail eq '${upn.replace(/'/g, "''")}'&$select=id&$top=1`);
    if (res.value?.length) return res.value[0].id;
  } catch { }

  // 3. Try filter by mailNickname (prefix before @) — key for cross-tenant migrations
  //    e.g. "john.smith@sourcetenant.com" → find "john.smith" in target
  const alias = upn.split('@')[0];
  if (alias) {
    try {
      const res = await client.get(`/users?$filter=mailNickname eq '${alias.replace(/'/g, "''")}'&$select=id&$top=1`);
      if (res.value?.length) return res.value[0].id;
    } catch { }

    // 4. Also try proxyAddresses contains the alias (catches smtp: aliases)
    try {
      const res = await client.get(`/users?$filter=startsWith(userPrincipalName,'${alias.replace(/'/g, "''")}@')&$select=id&$top=1`);
      if (res.value?.length) return res.value[0].id;
    } catch { }
  }

  return null;
}

// ── Helper: resolve a group in a tenant by mail/displayName, return group object ──
// Throws a descriptive error on permission failures; returns null only if genuinely not found.
async function resolveGroupInTenant(client: GraphClient, identity: string): Promise<any | null> {
  const sel = `$select=id,displayName,mail,mailNickname,description,visibility,groupTypes`;

  // 1. Try direct ID lookup (works only for GUIDs)
  try { const g = await client.get(`/groups/${identity}?${sel}`); if (g?.id) return g; } catch { /* not a GUID or not found */ }

  const safeId = identity.replace(/'/g, "''");

  // 2. Try by mail address (basic filter — no ConsistencyLevel needed)
  try {
    const res = await client.get(`/groups?$filter=mail eq '${safeId}'&${sel}`);
    if (res.value?.length) return res.value[0];
  } catch (e: any) {
    if (e.message?.includes('403') || e.message?.toLowerCase().includes('insufficient') || e.message?.toLowerCase().includes('authorization')) {
      throw new Error(`Permission denied reading groups: ${e.message}. Ensure Group.Read.All or Group.ReadWrite.All admin consent has been granted.`);
    }
    // Otherwise (400/404/etc.) fall through and try next method
  }

  // 3. Try by displayName — MUST use ConsistencyLevel:eventual (advanced query requirement)
  try {
    const res = await client.getAdvanced(`/groups?$filter=displayName eq '${safeId}'&$count=true&${sel}`);
    if (res.value?.length) return res.value[0];
  } catch (e: any) {
    if (e.message?.includes('403') || e.message?.toLowerCase().includes('insufficient') || e.message?.toLowerCase().includes('authorization')) {
      throw new Error(`Permission denied reading groups: ${e.message}. Ensure Group.Read.All or Group.ReadWrite.All admin consent has been granted.`);
    }
    // Log but don't throw — return null below so caller can surface the error
  }

  return null;
}

async function migrateDistributionGroup(
  source: GraphClient,
  target: GraphClient,
  sourceIdentity: string,
  targetIdentity: string,
  itemId: number,
  allowM365Upgrade: boolean = false
): Promise<void> {
  const logs: string[] = [];
  logs.push(logEntry(`Starting distribution group migration: ${sourceIdentity} → ${targetIdentity || 'same name'}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    // Resolve source group
    const sourceGroup = await resolveGroupInTenant(source, sourceIdentity);
    if (!sourceGroup) throw new Error(
      `Source distribution group "${sourceIdentity}" not found. ` +
      `Searched by mail address and display name. ` +
      `Confirm the group exists in the source tenant and that the app registration has Group.Read.All consent.`
    );
    logs.push(logEntry(`Found source group: ${sourceGroup.displayName} (${sourceGroup.mail || 'no mail'})`));

    // Get members and owners
    const members = await source.getAllPages<any>(`/groups/${sourceGroup.id}/members?$select=id,userPrincipalName,mail,displayName`);
    const owners  = await source.getAllPages<any>(`/groups/${sourceGroup.id}/owners?$select=id,userPrincipalName,mail,displayName`);
    logs.push(logEntry(`Source: ${members.length} members, ${owners.length} owners`));

    // Determine target name and mail alias
    // If targetIdentity looks like an email, use its alias; otherwise use the source group's mailNickname
    const targetName = targetIdentity || sourceGroup.displayName;
    const targetNickname = (targetIdentity?.includes('@')
      ? targetIdentity.split('@')[0]
      : sourceGroup.mailNickname || targetIdentity || '')
      .replace(/[^a-zA-Z0-9]/g, '').slice(0, 59) || `dl${Date.now()}`;

    // Create or find group in target.
    // Graph API v1.0 CANNOT create classic distribution lists or mail-enabled security groups (MESG).
    // Behaviour:
    //   allowM365Upgrade=false (default): attempt MESG first; if the API rejects it (400), auto-fall back
    //     to M365 Unified Group — the only mail-enabled group type Graph API can actually create.
    //   allowM365Upgrade=true: skip the MESG attempt entirely and go straight to M365 Unified Group.
    // Either way the migration will complete; the log records which group type was used.
    // If mailNickname is already taken, retry once with a unique suffix.
    let targetGroup = await resolveGroupInTenant(target, targetName);
    if (!targetGroup) {
      const tryCreate = async (nickname: string) => {
        const base = {
          displayName: sourceGroup.displayName,
          mailNickname: nickname,
          description: (sourceGroup.description || '').slice(0, 1024),
        };

        if (!allowM365Upgrade) {
          // Attempt 1: Mail-Enabled Security Group
          logs.push(logEntry(`Creating distribution group "${targetName}" in target (attempting Mail-Enabled Security Group)...`));
          try {
            const g = await target.post('/groups', { ...base, mailEnabled: true, securityEnabled: true, groupTypes: [] });
            logs.push(logEntry(`✓ Mail-enabled security group created (ID: ${g.id})`));
            return g;
          } catch (e1: any) {
            // Graph API returns 400 when MESG/DL creation is not supported — this is expected.
            logs.push(logEntry(`  MESG creation rejected by Graph API (${e1.message?.slice(0, 120)})`));
            logs.push(logEntry(`  Graph API does not support creating classic distribution lists. Auto-falling back to M365 Unified Group…`));
          }
        } else {
          logs.push(logEntry(`Creating distribution group "${targetName}" as M365 Unified Group (MESG skipped)...`));
        }

        // Fallback (or primary when allowM365Upgrade=true): M365 Unified Group
        const g2 = await target.post('/groups', { ...base, mailEnabled: true, securityEnabled: false, groupTypes: ['Unified'], visibility: 'Private' });
        logs.push(logEntry(`✓ Created as M365 Unified Group (ID: ${g2.id}) — Note: Microsoft 365 Groups are the only mail-enabled group type creatable via Graph API`));
        return g2;
      };

      try {
        targetGroup = await tryCreate(targetNickname);
      } catch (err: any) {
        // Nickname may be taken — retry once with a unique suffix
        if (err.message?.includes('409') || err.message?.toLowerCase().includes('conflict') || err.message?.includes('already exists') || err.message?.includes('ObjectConflict')) {
          const uniqueNickname = `${targetNickname.slice(0, 52)}${Date.now().toString(36)}`;
          logs.push(logEntry(`  Nickname conflict — retrying with unique suffix: ${uniqueNickname}`));
          targetGroup = await tryCreate(uniqueNickname);
        } else {
          throw new Error(`Failed to create group in target: ${err.message}`);
        }
      }
    } else {
      logs.push(logEntry(`Group already exists in target (ID: ${targetGroup.id})`));
    }

    // Add owners first (must be users in target tenant)
    let ownersMigrated = 0;
    for (const owner of owners) {
      const email = owner.userPrincipalName || owner.mail;
      if (!email) continue;
      const targetUserId = await resolveUserInTenant(target, email);
      if (!targetUserId) {
        logs.push(logEntry(`  Owner not found in target: ${email} (alias '${email.split('@')[0]}' — ensure user exists in target tenant)`));
        continue;
      }
      try {
        await target.post(`/groups/${targetGroup.id}/owners/$ref`, {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${targetUserId}`,
        });
        logs.push(logEntry(`  ✓ Owner added: ${email}`));
        ownersMigrated++;
      } catch (e: any) { logs.push(logEntry(`  Owner already exists or failed: ${email}`)); }
    }

    // Add members
    let membersMigrated = 0;
    for (const member of members) {
      const email = member.userPrincipalName || member.mail;
      if (!email) continue;
      const targetUserId = await resolveUserInTenant(target, email);
      if (!targetUserId) {
        logs.push(logEntry(`  Member not found in target: ${email} (alias '${email.split('@')[0]}' — ensure user exists in target tenant)`));
        continue;
      }
      try {
        await target.post(`/groups/${targetGroup.id}/members/$ref`, {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${targetUserId}`,
        });
        logs.push(logEntry(`  ✓ Member added: ${email}`));
        membersMigrated++;
      } catch (e: any) { logs.push(logEntry(`  Member already exists or failed: ${email}`)); }
    }

    logs.push(logEntry(`✓ Owners added: ${ownersMigrated}/${owners.length}`));
    logs.push(logEntry(`✓ Members added: ${membersMigrated}/${members.length}`));
    if (membersMigrated < members.length) {
      logs.push(logEntry(`  Note: ${members.length - membersMigrated} member(s) not found in target. Members must exist in target tenant. Cross-tenant UPNs are matched by alias (prefix before @).`));
    }
    logs.push(logEntry('Distribution group migration complete'));
    await updateItemProgress(itemId, 'completed', logs);
  } catch (err: any) {
    console.error(`[migration] distributiongroup ${itemId} FAILED:`, err.message);
    logs.push(logEntry(`Distribution group migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

async function migrateSharedMailbox(
  source: GraphClient,
  target: GraphClient,
  sourceIdentity: string,
  targetIdentity: string,
  itemId: number,
  project?: Project
): Promise<void> {
  const logs: string[] = [];
  logs.push(logEntry(`Starting shared mailbox migration: ${sourceIdentity} → ${targetIdentity || '(will derive from target domain)'}`));
  console.log(`[migration] sharedmailbox ${itemId}: starting — sourceIdentity="${sourceIdentity}" targetIdentity="${targetIdentity}"`);
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    // ── Step 1: Read source shared mailbox properties via Graph ──────────
    logs.push(logEntry(`Reading source shared mailbox: ${sourceIdentity}`));
    await updateItemProgress(itemId, 'in_progress', logs);

    let sourceLookupError: string | null = null;
    const sourceUser = await source
      .get(`/users/${encodeURIComponent(sourceIdentity)}?$select=id,displayName,mail,userPrincipalName,mailNickname`)
      .catch((e: any) => { sourceLookupError = e.message; return null; });

    if (!sourceUser) {
      const isPermission = sourceLookupError?.includes('403') || sourceLookupError?.toLowerCase().includes('insufficient');
      throw new Error(
        `Source shared mailbox "${sourceIdentity}" not found${isPermission
          ? ' — Permission denied. Ensure User.Read.All is granted via admin consent on the SOURCE app registration.'
          : sourceLookupError ? ` — ${sourceLookupError}` : '.'}`
      );
    }
    logs.push(logEntry(`✓ Source: ${sourceUser.displayName} <${sourceUser.mail || sourceUser.userPrincipalName}>`));

    const alias = (sourceUser.mailNickname || sourceIdentity.split('@')[0]).replace(/[^a-zA-Z0-9._-]/g, '');
    const displayName = sourceUser.displayName || alias;

    // ── Step 2: Determine target SMTP address ────────────────────────────
    let targetSmtp = (targetIdentity || '').trim();
    if (!targetSmtp || targetSmtp === sourceIdentity) {
      logs.push(logEntry(`No explicit target address — looking up target tenant's default domain...`));
      await updateItemProgress(itemId, 'in_progress', logs);
      try {
        const domainsRes = await target.get(`/domains?$select=id,isDefault,isVerified`);
        const domains: any[] = domainsRes.value || [];
        const defaultDomain =
          domains.find((d: any) => d.isDefault && d.isVerified)?.id ||
          domains.find((d: any) => d.isVerified && !d.id.endsWith('.onmicrosoft.com'))?.id ||
          domains.find((d: any) => d.isVerified)?.id;
        if (defaultDomain) {
          targetSmtp = `${alias}@${defaultDomain}`;
          logs.push(logEntry(`  Target address: ${targetSmtp} (target tenant default domain: ${defaultDomain})`));
        } else {
          throw new Error('Could not determine a verified domain in the target tenant. Set an explicit target address in the Migration tab.');
        }
      } catch (e: any) {
        throw new Error(`Domain lookup failed: ${e.message}`);
      }
    }
    logs.push(logEntry(`Target SMTP: ${targetSmtp}`));

    // ── Step 3: Create shared mailbox directly via EXO (or show manual cmd) ──
    // Shared mailboxes CANNOT be created via Graph API and do NOT need a license.
    // The correct approach is New-Mailbox -Shared which creates the Exchange object
    // and corresponding AAD user simultaneously without requiring a license.
    const exo = project?.exoSettings as any;
    const hasSourceExo = !!(exo?.sourceCertPath && exo?.sourceOrg);
    const hasTargetExo = !!(exo?.targetCertPath && exo?.targetOrg);
    const autoDelegate = exo?.autoDelegate !== false;

    let mailboxReady = false;  // true when shared mailbox exists and is confirmed ready

    if (hasTargetExo) {
      logs.push(logEntry(''));
      logs.push(logEntry('════ Creating shared mailbox via Exchange Online PowerShell ════'));
      logs.push(logEntry(`Running: New-Mailbox -Shared -Name "${displayName}" -Alias "${alias}" -PrimarySmtpAddress "${targetSmtp}"`));
      await updateItemProgress(itemId, 'in_progress', logs);

      const targetCfg: ExoConfig = {
        clientId: project!.targetClientId!,
        certPath: exo.targetCertPath,
        certPassword: exo.targetCertPassword || '',
        organization: exo.targetOrg,
      };

      const createResult = await createSharedMailboxDirect(targetCfg, displayName, alias, targetSmtp);
      for (const line of createResult.output) {
        logs.push(logEntry(`  ${line}`));
      }

      if (createResult.success || createResult.alreadyExists) {
        mailboxReady = true;
        logs.push(logEntry(createResult.alreadyExists
          ? `✓ Shared mailbox already exists: ${targetSmtp}`
          : `✓ Shared mailbox created: ${targetSmtp} (no license required)`));
      } else {
        logs.push(logEntry(`⚠ Shared mailbox creation errors: ${createResult.errors.slice(0, 3).join('; ')}`));
        logs.push(logEntry(`  You can re-run this migration or create it manually:`));
        logs.push(logEntry(`    New-Mailbox -Shared -Name "${displayName}" -Alias "${alias}" -PrimarySmtpAddress "${targetSmtp}"`));
      }
    } else {
      // EXO not configured — cannot create shared mailbox via Graph
      logs.push(logEntry(''));
      logs.push(logEntry('════ ACTION REQUIRED: Create shared mailbox manually ════'));
      logs.push(logEntry('Shared mailboxes cannot be created via Microsoft Graph API.'));
      logs.push(logEntry('Configure Exchange Online PowerShell in the Tenant Config tab to automate this,'));
      logs.push(logEntry('or run the following command in Exchange Online PowerShell:'));
      logs.push(logEntry(''));
      logs.push(logEntry(`  Connect-ExchangeOnline`));
      logs.push(logEntry(`  New-Mailbox -Shared -Name "${displayName}" -Alias "${alias}" -PrimarySmtpAddress "${targetSmtp}"`));
      logs.push(logEntry(''));
      logs.push(logEntry('No license is required. The command creates the shared mailbox and its AAD object automatically.'));
      logs.push(logEntry('After creating it, re-run this migration to migrate delegates.'));
      logs.push(logEntry('══════════════════════════════════════════════════════'));
    }

    // ── Step 4: Delegate permissions via EXO ─────────────────────────────
    logs.push(logEntry(''));
    logs.push(logEntry('════ Delegate Permissions (FullAccess / SendAs / SendOnBehalf) ════'));

    if (hasSourceExo && hasTargetExo && autoDelegate && mailboxReady) {
      logs.push(logEntry('Reading delegates from source mailbox...'));
      await updateItemProgress(itemId, 'in_progress', logs);

      const sourceCfg: ExoConfig = {
        clientId: project!.sourceClientId!,
        certPath: exo.sourceCertPath,
        certPassword: exo.sourceCertPassword || '',
        organization: exo.sourceOrg,
      };
      const targetCfg: ExoConfig = {
        clientId: project!.targetClientId!,
        certPath: exo.targetCertPath,
        certPassword: exo.targetCertPassword || '',
        organization: exo.targetOrg,
      };

      const { delegates, errors: readErrors } = await readMailboxDelegates(sourceCfg, sourceIdentity);
      if (readErrors.length) logs.push(logEntry(`  Source EXO warnings: ${readErrors.slice(0, 3).join('; ')}`));

      if (delegates.length === 0) {
        logs.push(logEntry('  No delegates found on source mailbox (FullAccess/SendAs/SendOnBehalf).'));
      } else {
        logs.push(logEntry(`  Found ${delegates.length} delegate(s): ${delegates.map(d => d.user).join(', ')}`));
        const applyResult = await applyMailboxDelegates(targetCfg, targetSmtp, delegates);
        for (const line of applyResult.output) logs.push(logEntry(`  ${line}`));
        if (applyResult.errors.length) logs.push(logEntry(`  EXO errors: ${applyResult.errors.slice(0, 5).join('; ')}`));
        logs.push(logEntry(applyResult.success
          ? `✓ Delegates applied (${delegates.length})`
          : '⚠ Some delegates may not have been applied — check errors above'));
      }
    } else if (!hasSourceExo || !hasTargetExo) {
      logs.push(logEntry(mailboxReady
        ? 'EXO not fully configured — delegates must be applied manually after creation:'
        : 'Configure EXO in Tenant Config to auto-apply delegates. Manual commands:'));
      logs.push(logEntry(`  Add-MailboxPermission -Identity "${targetSmtp}" -User "<delegate>" -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false`));
      logs.push(logEntry(`  Add-RecipientPermission -Identity "${targetSmtp}" -Trustee "<delegate>" -AccessRights SendAs -Confirm:$false`));
      logs.push(logEntry(`  Set-Mailbox -Identity "${targetSmtp}" -GrantSendOnBehalfTo @{add="<delegate>"} -Confirm:$false`));
    }
    logs.push(logEntry('══════════════════════════════════════════════════════════'));

    // ── Final status ──────────────────────────────────────────────────────
    if (mailboxReady) {
      logs.push(logEntry('✓ Shared mailbox migration complete'));
      await updateItemProgress(itemId, 'completed', logs);
    } else {
      logs.push(logEntry('⚠ STATUS: Needs Action — configure Exchange Online PowerShell in Tenant Config to create this shared mailbox automatically.'));
      await updateItemProgress(
        itemId,
        'needs_action',
        logs,
        `Shared mailbox not yet created. Run: New-Mailbox -Shared -Name "${displayName}" -Alias "${alias}" -PrimarySmtpAddress "${targetSmtp}" — or configure EXO PowerShell in Tenant Config.`
      );
    }
  } catch (err: any) {
    console.error(`[migration] sharedmailbox ${itemId} FAILED:`, err.message);
    logs.push(logEntry(`Shared mailbox migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

async function migrateM365Group(
  source: GraphClient,
  target: GraphClient,
  sourceIdentity: string,
  targetIdentity: string,
  itemId: number
): Promise<void> {
  const logs: string[] = [];
  logs.push(logEntry(`Starting M365 Group migration: ${sourceIdentity} → ${targetIdentity || 'same name'}`));
  await updateItemProgress(itemId, 'in_progress', logs);

  try {
    const sourceGroup = await resolveGroupInTenant(source, sourceIdentity);
    if (!sourceGroup) throw new Error(
      `Source M365 Group "${sourceIdentity}" not found. ` +
      `Searched by mail address and display name. ` +
      `Confirm the group exists in the source tenant and that the app registration has Group.Read.All consent.`
    );
    logs.push(logEntry(`Found source group: ${sourceGroup.displayName} (${sourceGroup.mail || 'no mail'})`));

    const members = await source.getAllPages<any>(`/groups/${sourceGroup.id}/members?$select=id,userPrincipalName,mail,displayName`);
    const owners  = await source.getAllPages<any>(`/groups/${sourceGroup.id}/owners?$select=id,userPrincipalName,mail,displayName`);
    logs.push(logEntry(`Source: ${members.length} members, ${owners.length} owners`));

    const targetName     = targetIdentity || sourceGroup.displayName;
    const targetNickname = (targetIdentity?.includes('@')
      ? targetIdentity.split('@')[0]
      : sourceGroup.mailNickname || targetIdentity || '')
      .replace(/[^a-zA-Z0-9]/g, '').slice(0, 59) || `m365g${Date.now()}`;

    // Create or find group in target
    let targetGroup = await resolveGroupInTenant(target, targetName);
    if (!targetGroup) {
      logs.push(logEntry(`Creating M365 Group "${targetName}" in target...`));
      const createM365 = async (nickname: string) => {
        const g = await target.post('/groups', {
          displayName: sourceGroup.displayName,
          mailNickname: nickname,
          mailEnabled: true,
          securityEnabled: false,
          groupTypes: ['Unified'],
          visibility: sourceGroup.visibility || 'Private',
          description: (sourceGroup.description || '').slice(0, 1024),
        });
        return g;
      };
      try {
        targetGroup = await createM365(targetNickname);
      } catch (err: any) {
        if (err.message?.includes('409') || err.message?.toLowerCase().includes('conflict') || err.message?.includes('ObjectConflict')) {
          const uniqueNickname = `${targetNickname.slice(0, 52)}${Date.now().toString(36)}`;
          logs.push(logEntry(`  Nickname conflict — retrying with suffix: ${uniqueNickname}`));
          targetGroup = await createM365(uniqueNickname);
        } else {
          throw new Error(`Failed to create M365 Group: ${err.message}`);
        }
      }
      logs.push(logEntry(`✓ M365 Group created (ID: ${targetGroup.id})`));
      // Wait briefly for group to fully provision
      await new Promise(r => setTimeout(r, 3000));
    } else {
      logs.push(logEntry(`M365 Group already exists in target (ID: ${targetGroup.id})`));
    }

    // Add owners
    let ownersMigrated = 0;
    for (const owner of owners) {
      const email = owner.userPrincipalName || owner.mail;
      if (!email) continue;
      const targetUserId = await resolveUserInTenant(target, email);
      if (!targetUserId) {
        logs.push(logEntry(`  Owner not found in target: ${email} (alias '${email.split('@')[0]}' — ensure user exists in target tenant)`));
        continue;
      }
      try {
        await target.post(`/groups/${targetGroup.id}/owners/$ref`, {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${targetUserId}`,
        });
        logs.push(logEntry(`  ✓ Owner added: ${email}`));
        ownersMigrated++;
      } catch (e: any) { logs.push(logEntry(`  Owner already exists or failed: ${email}`)); }
    }

    // Add members
    let membersMigrated = 0;
    for (const member of members) {
      const email = member.userPrincipalName || member.mail;
      if (!email) continue;
      const targetUserId = await resolveUserInTenant(target, email);
      if (!targetUserId) {
        logs.push(logEntry(`  Member not found in target: ${email} (alias '${email.split('@')[0]}' — ensure user exists in target tenant)`));
        continue;
      }
      try {
        await target.post(`/groups/${targetGroup.id}/members/$ref`, {
          '@odata.id': `https://graph.microsoft.com/v1.0/users/${targetUserId}`,
        });
        logs.push(logEntry(`  ✓ Member added: ${email}`));
        membersMigrated++;
      } catch (e: any) { logs.push(logEntry(`  Member already exists or failed: ${email}`)); }
    }

    logs.push(logEntry(`✓ Owners added: ${ownersMigrated}/${owners.length}`));
    logs.push(logEntry(`✓ Members added: ${membersMigrated}/${members.length}`));
    if (membersMigrated < members.length) {
      logs.push(logEntry(`  Note: ${members.length - membersMigrated} member(s) not found in target. Members must exist in target tenant. Cross-tenant UPNs are matched by alias (prefix before @).`));
    }
    logs.push(logEntry('✓ M365 Group migration complete'));
    await updateItemProgress(itemId, 'completed', logs);
  } catch (err: any) {
    console.error(`[migration] m365group ${itemId} FAILED:`, err.message);
    logs.push(logEntry(`M365 Group migration failed: ${err.message}`));
    await updateItemProgress(itemId, 'failed', logs, err.message);
  }
}

export async function migrateItem(projectId: number, itemId: number): Promise<void> {
  const project = await storage.getProjectInternal(projectId);
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
        await migrateTeam(source, target, item.sourceIdentity, targetIdentity, itemId);
        break;
      case 'user':
        await migrateUser(source, target, item.sourceIdentity, targetIdentity, itemId);
        break;
      case 'distributiongroup':
        await migrateDistributionGroup(source, target, item.sourceIdentity, targetIdentity, itemId, !!(item.options as any)?.allowM365Upgrade);
        break;
      case 'sharedmailbox':
        await migrateSharedMailbox(source, target, item.sourceIdentity, targetIdentity, itemId, project);
        break;
      case 'm365group':
        await migrateM365Group(source, target, item.sourceIdentity, targetIdentity, itemId);
        break;
      case 'powerplatform':
        await storage.updateItem(itemId, { status: 'failed', errorDetails: 'Power Platform migration requires the Power Platform Admin PowerShell module or CoE Starter Kit. Export/import via Graph API is not supported.' });
        await storage.updateItemLogs(itemId, [logEntry('Power Platform migration is not supported via Graph API. Use the Power Platform CoE Starter Kit or PowerShell module to export and import Power Apps and Power Automate flows.')]);
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
  const project = await storage.getProjectInternal(projectId);
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
