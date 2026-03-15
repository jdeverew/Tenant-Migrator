// ldapjs is loaded lazily so the server starts even if the package is not yet installed.
// Run `npm install` (or `npm install ldapjs`) to enable direct LDAP migration.
import type ldapTypes from "ldapjs";
import type { GraphClient } from "./graph-client";

export interface EntraCloudOnlyUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  givenName: string | null;
  surname: string | null;
  jobTitle: string | null;
  department: string | null;
  officeLocation: string | null;
  mobilePhone: string | null;
  accountEnabled: boolean;
  usageLocation: string | null;
}

export interface AdConnectionConfig {
  dcHostname: string;
  ldapPort: number;
  bindDn: string;
  bindPassword: string;
  baseDn: string;
  useSsl: boolean;
  targetOu?: string | null;
}

// ── Lazy loader ──────────────────────────────────────────────────────────────
let _ldap: typeof ldapTypes | null = null;

async function getLdap(): Promise<typeof ldapTypes> {
  if (_ldap) return _ldap;
  try {
    const mod = await import("ldapjs");
    _ldap = (mod.default ?? mod) as typeof ldapTypes;
    return _ldap;
  } catch {
    throw new Error(
      "The ldapjs package is not installed. Run `npm install` in your project folder and restart the server."
    );
  }
}

// ── Entra discovery ──────────────────────────────────────────────────────────

export async function discoverCloudOnlyUsers(client: GraphClient): Promise<EntraCloudOnlyUser[]> {
  const users = await client.getAllPages<any>(
    `/users?$select=id,displayName,userPrincipalName,mail,givenName,surname,jobTitle,department,officeLocation,mobilePhone,accountEnabled,usageLocation,onPremisesSyncEnabled,onPremisesDistinguishedName,assignedLicenses&$top=999`
  );

  return users
    .filter((u: any) => {
      if (u.userPrincipalName?.includes('#EXT#')) return false;
      const isSynced = u.onPremisesSyncEnabled === true || !!u.onPremisesDistinguishedName;
      return !isSynced;
    })
    .map((u: any) => ({
      id: u.id,
      displayName: u.displayName || u.userPrincipalName,
      userPrincipalName: u.userPrincipalName,
      mail: u.mail || null,
      givenName: u.givenName || null,
      surname: u.surname || null,
      jobTitle: u.jobTitle || null,
      department: u.department || null,
      officeLocation: u.officeLocation || null,
      mobilePhone: u.mobilePhone || null,
      accountEnabled: u.accountEnabled !== false,
      usageLocation: u.usageLocation || null,
    }));
}

// ── LDAP helpers ─────────────────────────────────────────────────────────────

function escapeLdapFilter(val: string): string {
  return val
    .replace(/\\/g, '\\5c')
    .replace(/\*/g, '\\2a')
    .replace(/\(/g, '\\28')
    .replace(/\)/g, '\\29')
    .replace(/\0/g, '\\00');
}

async function createLdapClient(config: AdConnectionConfig): Promise<ldapTypes.Client> {
  const ldap = await getLdap();
  return new Promise((resolve, reject) => {
    const url = `${config.useSsl ? 'ldaps' : 'ldap'}://${config.dcHostname}:${config.ldapPort}`;
    const client = ldap.createClient({
      url,
      tlsOptions: config.useSsl ? { rejectUnauthorized: false } : undefined,
      timeout: 10000,
      connectTimeout: 10000,
    });

    client.on('error', (err: Error) => {
      reject(new Error(`LDAP connection error: ${err.message}`));
    });

    client.bind(config.bindDn, config.bindPassword, (err) => {
      if (err) {
        client.destroy();
        reject(new Error(`LDAP bind failed: ${err.message || 'Invalid credentials or DN'}`));
      } else {
        resolve(client);
      }
    });
  });
}

// ── Public API ────────────────────────────────────────────────────────────────

export async function testAdConnection(config: AdConnectionConfig): Promise<{ success: boolean; message: string }> {
  let client: ldapTypes.Client | null = null;
  try {
    client = await createLdapClient(config);

    await new Promise<void>((resolve, reject) => {
      client!.search(
        config.baseDn,
        { scope: 'base', filter: '(objectClass=*)' },
        (err, res) => {
          if (err) return reject(new Error(`Base DN search failed: ${err.message}`));
          res.on('error', (e: Error) => reject(new Error(`Search error: ${e.message}`)));
          res.on('end', (result: any) => {
            if (result && result.status !== 0) {
              reject(new Error(`Base DN not found or not accessible (status ${result.status})`));
            } else {
              resolve();
            }
          });
          res.on('searchEntry', () => {});
        }
      );
    });

    return {
      success: true,
      message: `Connected successfully to ${config.dcHostname}:${config.ldapPort}. Base DN "${config.baseDn}" is accessible.`,
    };
  } catch (err: any) {
    return { success: false, message: err.message || 'Connection failed' };
  } finally {
    if (client) client.destroy();
  }
}

function checkUserExists(client: ldapTypes.Client, baseDn: string, upn: string): Promise<boolean> {
  return new Promise((resolve) => {
    client.search(
      baseDn,
      { scope: 'sub', filter: `(userPrincipalName=${escapeLdapFilter(upn)})`, attributes: ['dn'] },
      (err, res) => {
        if (err) return resolve(false);
        let found = false;
        res.on('searchEntry', () => { found = true; });
        res.on('error', () => resolve(false));
        res.on('end', () => resolve(found));
      }
    );
  });
}

function encodePassword(password: string): Buffer {
  const quoted = `"${password}"`;
  const buf = Buffer.alloc(quoted.length * 2);
  for (let i = 0; i < quoted.length; i++) {
    buf.writeUInt16LE(quoted.charCodeAt(i), i * 2);
  }
  return buf;
}

export interface AdMigrationResult {
  userPrincipalName: string;
  success: boolean;
  created: boolean;
  message: string;
  tempPassword?: string;
}

export async function migrateUserToAd(
  config: AdConnectionConfig,
  user: EntraCloudOnlyUser,
  targetUpn: string,
): Promise<AdMigrationResult> {
  const ldap = await getLdap();
  let client: ldapTypes.Client | null = null;
  try {
    client = await createLdapClient(config);

    const samAccountName = (targetUpn.split('@')[0] || user.userPrincipalName.split('@')[0])
      .replace(/[^a-zA-Z0-9._-]/g, '')
      .substring(0, 20);

    const exists = await checkUserExists(client, config.baseDn, targetUpn);
    if (exists) {
      client.destroy();
      return {
        userPrincipalName: targetUpn,
        success: true,
        created: false,
        message: `User ${targetUpn} already exists in Active Directory — skipped.`,
      };
    }

    const ouDn = config.targetOu || config.baseDn;
    const cn = (user.displayName || `${user.givenName || ''} ${user.surname || ''}`.trim() || samAccountName);
    const userDn = `CN=${cn},${ouDn}`;
    const tempPassword = `Migr@tion${Math.random().toString(36).slice(-6).toUpperCase()}1!`;

    const entry: Record<string, any> = {
      objectClass: ['top', 'person', 'organizationalPerson', 'user'],
      cn,
      sAMAccountName: samAccountName,
      userPrincipalName: targetUpn,
      userAccountControl: '514', // disabled initially
    };

    if (user.givenName) entry.givenName = user.givenName;
    if (user.surname) entry.sn = user.surname;
    if (user.mail || targetUpn) entry.mail = user.mail || targetUpn;
    if (user.jobTitle) entry.title = user.jobTitle;
    if (user.department) entry.department = user.department;
    if (user.officeLocation) entry.physicalDeliveryOfficeName = user.officeLocation;
    if (user.mobilePhone) entry.mobile = user.mobilePhone;
    if (user.displayName) entry.displayName = user.displayName;
    entry.description = `Migrated from Entra ID (${user.userPrincipalName})`;

    // Step 1: Create user object
    await new Promise<void>((resolve, reject) => {
      client!.add(userDn, entry, (err) => {
        if (err) reject(new Error(`Failed to create AD user object: ${err.message}`));
        else resolve();
      });
    });

    // Step 2: Set password (requires LDAPS in production)
    const passwordBuf = encodePassword(tempPassword);
    await new Promise<void>((resolve) => {
      const change = new ldap.Change({
        operation: 'replace',
        modification: new ldap.Attribute({ type: 'unicodePwd', values: [passwordBuf] }),
      });
      client!.modify(userDn, change, () => resolve()); // non-fatal if fails over plain LDAP
    });

    // Step 3: Enable account
    await new Promise<void>((resolve) => {
      const change = new ldap.Change({
        operation: 'replace',
        modification: new ldap.Attribute({ type: 'userAccountControl', values: ['512'] }),
      });
      client!.modify(userDn, change, () => resolve());
    });

    // Step 4: Force password change at next logon
    await new Promise<void>((resolve) => {
      const change = new ldap.Change({
        operation: 'replace',
        modification: new ldap.Attribute({ type: 'pwdLastSet', values: ['0'] }),
      });
      client!.modify(userDn, change, () => resolve());
    });

    client.destroy();
    return {
      userPrincipalName: targetUpn,
      success: true,
      created: true,
      message: `User created: ${userDn}`,
      tempPassword,
    };
  } catch (err: any) {
    if (client) client.destroy();
    return {
      userPrincipalName: targetUpn,
      success: false,
      created: false,
      message: err.message || 'Unknown error',
    };
  }
}

// ── PowerShell export ─────────────────────────────────────────────────────────

export function generatePowerShellScript(
  users: EntraCloudOnlyUser[],
  targetUpns: string[],
  config: Pick<AdConnectionConfig, 'baseDn' | 'targetOu'>
): string {
  const ou = config.targetOu || config.baseDn;
  const date = new Date().toISOString().split('T')[0];

  const lines: string[] = [
    `# Entra ID to Active Directory Migration Script`,
    `# Generated: ${date}`,
    `# Source: Entra ID cloud-only accounts`,
    `# Target OU: ${ou}`,
    `#`,
    `# Requirements:`,
    `#   - Run on a Domain Controller or machine with RSAT (AD PowerShell module)`,
    `#   - Run as Domain Admin or account with 'Create User' rights in the target OU`,
    ``,
    `Import-Module ActiveDirectory -ErrorAction Stop`,
    ``,
    `$TargetOU = "${ou}"`,
    `$ErrorLog = @()`,
    `$SuccessLog = @()`,
    ``,
  ];

  users.forEach((user, i) => {
    const targetUpn = targetUpns[i] || user.userPrincipalName;
    const samAccount = targetUpn.split('@')[0].replace(/[^a-zA-Z0-9._-]/g, '').substring(0, 20);
    const tempPass = `Migr@tion${Math.random().toString(36).slice(-6).toUpperCase()}1!`;
    const cn = (user.displayName || `${user.givenName || ''} ${user.surname || ''}`.trim() || samAccount).replace(/'/g, "''");

    lines.push(`# --- User: ${user.displayName} (${user.userPrincipalName}) ---`);
    lines.push(`try {`);
    lines.push(`  $existing = Get-ADUser -Filter { UserPrincipalName -eq '${targetUpn}' } -ErrorAction SilentlyContinue`);
    lines.push(`  if ($existing) {`);
    lines.push(`    Write-Host "SKIP: ${targetUpn} already exists" -ForegroundColor Yellow`);
    lines.push(`    $ErrorLog += "SKIP: ${targetUpn} already exists"`);
    lines.push(`  } else {`);
    lines.push(`    $SecurePass = ConvertTo-SecureString '${tempPass}' -AsPlainText -Force`);
    lines.push(`    $params = @{`);
    lines.push(`      Name              = '${cn}'`);
    lines.push(`      DisplayName       = '${cn}'`);
    if (user.givenName) lines.push(`      GivenName         = '${user.givenName.replace(/'/g, "''")}'`);
    if (user.surname) lines.push(`      Surname           = '${user.surname.replace(/'/g, "''")}'`);
    lines.push(`      UserPrincipalName = '${targetUpn}'`);
    lines.push(`      SamAccountName    = '${samAccount}'`);
    lines.push(`      EmailAddress      = '${user.mail || targetUpn}'`);
    if (user.jobTitle) lines.push(`      Title             = '${user.jobTitle.replace(/'/g, "''")}'`);
    if (user.department) lines.push(`      Department        = '${user.department.replace(/'/g, "''")}'`);
    if (user.officeLocation) lines.push(`      Office            = '${user.officeLocation.replace(/'/g, "''")}'`);
    if (user.mobilePhone) lines.push(`      MobilePhone       = '${user.mobilePhone}'`);
    lines.push(`      Path              = $TargetOU`);
    lines.push(`      AccountPassword   = $SecurePass`);
    lines.push(`      Enabled           = $true`);
    lines.push(`      ChangePasswordAtLogon = $true`);
    lines.push(`      Description       = 'Migrated from Entra ID (${user.userPrincipalName})'`);
    lines.push(`    }`);
    lines.push(`    New-ADUser @params`);
    lines.push(`    Write-Host "CREATED: ${targetUpn} | Temp password: ${tempPass}" -ForegroundColor Green`);
    lines.push(`    $SuccessLog += "CREATED: ${targetUpn} | TempPass: ${tempPass}"`);
    lines.push(`  }`);
    lines.push(`} catch {`);
    lines.push(`  Write-Host "ERROR: ${targetUpn} - $($_.Exception.Message)" -ForegroundColor Red`);
    lines.push(`  $ErrorLog += "ERROR: ${targetUpn} - $($_.Exception.Message)"`);
    lines.push(`}`);
    lines.push(``);
  });

  lines.push(`# --- Summary ---`);
  lines.push(`Write-Host ""`);
  lines.push(`Write-Host "=== Migration Complete ===" -ForegroundColor Cyan`);
  lines.push(`Write-Host "Created: $($SuccessLog.Count)"`);
  lines.push(`Write-Host "Skipped/Errors: $($ErrorLog.Count)"`);
  lines.push(`if ($SuccessLog.Count -gt 0) {`);
  lines.push(`  Write-Host ""`);
  lines.push(`  Write-Host "--- Created Users + Temp Passwords ---" -ForegroundColor Green`);
  lines.push(`  $SuccessLog | ForEach-Object { Write-Host $_ }`);
  lines.push(`}`);
  lines.push(`if ($ErrorLog.Count -gt 0) {`);
  lines.push(`  Write-Host ""`);
  lines.push(`  Write-Host "--- Issues ---" -ForegroundColor Yellow`);
  lines.push(`  $ErrorLog | ForEach-Object { Write-Host $_ }`);
  lines.push(`}`);

  return lines.join('\n');
}
