/**
 * Exchange Online PowerShell runner.
 *
 * Exchange Online delegate permissions (FullAccess, SendAs, SendOnBehalf) are NOT
 * accessible via Microsoft Graph API v1.0 — they are Exchange-level permissions.
 * The only way to read and write them programmatically without an interactive session
 * is via the ExchangeOnlineManagement PowerShell module with app-only certificate auth.
 *
 * Auth requirements:
 *  - Azure AD app registration must have the "Exchange.ManageAsApp" API permission
 *    (under "APIs my organization uses" → "Office 365 Exchange Online")
 *    with admin consent granted.
 *  - A certificate must be uploaded to the app registration (Graph API / Azure Portal).
 *  - The certificate's private key (PFX) must be accessible on this machine.
 *
 * This runner spawns powershell.exe (Windows) or pwsh (cross-platform) and executes
 * commands against Exchange Online. On non-Windows environments it gracefully skips.
 */

import { spawn } from 'child_process';
import * as path from 'path';
import * as fs from 'fs';

export interface ExoConfig {
  clientId: string;
  certPath: string;         // Absolute path to PFX file
  certPassword: string;     // PFX password (may be empty)
  organization: string;     // e.g. "contoso.onmicrosoft.com"
}

export interface MailboxDelegate {
  user: string;             // UPN of the delegate
  fullAccess: boolean;
  sendAs: boolean;
  sendOnBehalf: boolean;
}

export interface ExoResult {
  success: boolean;
  output: string[];
  errors: string[];
}

// Find the PowerShell executable — powershell.exe on Windows, pwsh on Linux/macOS
function getPowerShell(): string | null {
  const candidates = ['powershell.exe', 'pwsh'];
  for (const c of candidates) {
    try {
      const { execSync } = require('child_process');
      execSync(`${c} -Command "$null"`, { stdio: 'ignore', timeout: 5000 });
      return c;
    } catch { /* not found or failed */ }
  }
  return null;
}

let _pwsh: string | null | undefined = undefined;
function findPowerShell(): string | null {
  if (_pwsh === undefined) _pwsh = getPowerShell();
  return _pwsh;
}

async function runPowerShellScript(script: string): Promise<ExoResult> {
  const pwsh = findPowerShell();
  if (!pwsh) {
    return {
      success: false,
      output: [],
      errors: ['PowerShell not found on this system. Exchange Online PowerShell requires Windows (powershell.exe) or PowerShell 7+ (pwsh).'],
    };
  }

  return new Promise((resolve) => {
    const outputLines: string[] = [];
    const errorLines: string[] = [];

    const proc = spawn(pwsh, ['-NoProfile', '-NonInteractive', '-Command', script], {
      timeout: 120_000,
      windowsHide: true,
    });

    proc.stdout.on('data', (d: Buffer) => {
      d.toString().split('\n').filter(Boolean).forEach(l => outputLines.push(l.trim()));
    });
    proc.stderr.on('data', (d: Buffer) => {
      d.toString().split('\n').filter(Boolean).forEach(l => {
        const line = l.trim();
        if (line && !line.startsWith('WARNING:')) errorLines.push(line);
      });
    });

    proc.on('close', (code) => {
      resolve({
        success: code === 0 && errorLines.length === 0,
        output: outputLines,
        errors: errorLines,
      });
    });

    proc.on('error', (err) => {
      resolve({ success: false, output: [], errors: [err.message] });
    });
  });
}

// Install the ExchangeOnlineManagement module if not present (requires internet)
export async function ensureExoModuleInstalled(): Promise<{ ok: boolean; message: string }> {
  const script = `
$ErrorActionPreference = 'Stop'
$mod = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
if (-not $mod) {
  Write-Host "Installing ExchangeOnlineManagement module..."
  Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber
  Write-Host "Module installed."
} else {
  Write-Host "ExchangeOnlineManagement $($mod.Version) already installed."
}
`;
  const result = await runPowerShellScript(script);
  const message = result.errors.length ? result.errors.join('; ') : result.output.join('; ');
  return { ok: result.success, message };
}

// Build the connection block for a given tenant
function buildConnectBlock(cfg: ExoConfig): string {
  // Escape paths and passwords for PowerShell single-quoted strings
  const escapedPath = cfg.certPath.replace(/'/g, "''");
  const escapedPass = cfg.certPassword.replace(/'/g, "''");
  const escapedOrg  = cfg.organization.replace(/'/g, "''");
  const escapedId   = cfg.clientId.replace(/'/g, "''");

  return `
Import-Module ExchangeOnlineManagement -MinimumVersion 3.0.0
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import('${escapedPath}', '${escapedPass}', [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet)
Connect-ExchangeOnline -AppId '${escapedId}' -Certificate $cert -Organization '${escapedOrg}' -ShowBanner:$false -ErrorAction Stop
`;
}

// Read all FullAccess and SendAs delegates from a mailbox
export async function readMailboxDelegates(cfg: ExoConfig, mailboxIdentity: string): Promise<{ delegates: MailboxDelegate[]; errors: string[] }> {
  const escaped = mailboxIdentity.replace(/'/g, "''");
  const script = `
$ErrorActionPreference = 'Stop'
${buildConnectBlock(cfg)}

$results = @()

# Full Access permissions
try {
  $fa = Get-MailboxPermission -Identity '${escaped}' | Where-Object { $_.User -notlike 'NT AUTHORITY*' -and $_.User -notlike 'S-1-5*' -and $_.IsInherited -eq $false }
  foreach ($p in $fa) {
    $results += "FULLACCESS:$($p.User)"
  }
} catch { Write-Warning "FullAccess read failed: $_" }

# Send As permissions
try {
  $sa = Get-RecipientPermission -Identity '${escaped}' | Where-Object { $_.Trustee -notlike 'NT AUTHORITY*' -and $_.IsInherited -eq $false }
  foreach ($p in $sa) {
    $results += "SENDAS:$($p.Trustee)"
  }
} catch { Write-Warning "SendAs read failed: $_" }

# Send on Behalf
try {
  $mb = Get-Mailbox -Identity '${escaped}' | Select-Object -ExpandProperty GrantSendOnBehalfTo
  foreach ($d in $mb) {
    $resolved = Get-Recipient $d -ErrorAction SilentlyContinue
    if ($resolved) { $results += "SENDONBEHALF:$($resolved.PrimarySmtpAddress)" }
  }
} catch { Write-Warning "SendOnBehalf read failed: $_" }

$results | ForEach-Object { Write-Host $_ }
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
`;

  const result = await runPowerShellScript(script);
  const delegates: MailboxDelegate[] = [];
  const delegateMap = new Map<string, MailboxDelegate>();

  for (const line of result.output) {
    const [type, user] = line.split(':');
    if (!user) continue;
    const key = user.toLowerCase();
    if (!delegateMap.has(key)) {
      delegateMap.set(key, { user, fullAccess: false, sendAs: false, sendOnBehalf: false });
    }
    const d = delegateMap.get(key)!;
    if (type === 'FULLACCESS')    d.fullAccess    = true;
    if (type === 'SENDAS')        d.sendAs        = true;
    if (type === 'SENDONBEHALF')  d.sendOnBehalf  = true;
  }

  return { delegates: Array.from(delegateMap.values()), errors: result.errors };
}

// Apply delegate permissions to a target mailbox
export async function applyMailboxDelegates(
  cfg: ExoConfig,
  targetIdentity: string,
  delegates: MailboxDelegate[]
): Promise<ExoResult> {
  if (delegates.length === 0) {
    return { success: true, output: ['No delegates to apply.'], errors: [] };
  }

  const escapedTarget = targetIdentity.replace(/'/g, "''");
  const commands: string[] = [];

  for (const d of delegates) {
    const escapedUser = d.user.replace(/'/g, "''");
    if (d.fullAccess) {
      commands.push(`Add-MailboxPermission -Identity '${escapedTarget}' -User '${escapedUser}' -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false`);
    }
    if (d.sendAs) {
      commands.push(`Add-RecipientPermission -Identity '${escapedTarget}' -Trustee '${escapedUser}' -AccessRights SendAs -Confirm:$false`);
    }
    if (d.sendOnBehalf) {
      commands.push(`Set-Mailbox -Identity '${escapedTarget}' -GrantSendOnBehalfTo @{add='${escapedUser}'} -Confirm:$false`);
    }
  }

  const script = `
$ErrorActionPreference = 'Stop'
${buildConnectBlock(cfg)}
${commands.map(cmd => `
try { ${cmd}; Write-Host "OK: ${cmd.slice(0, 60)}" } catch { Write-Host "FAIL: $($_.Exception.Message)" }`).join('\n')}
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
`;

  return runPowerShellScript(script);
}

// Quick connectivity test — verifies cert auth works
export async function testExoConnection(cfg: ExoConfig): Promise<ExoResult> {
  const script = `
$ErrorActionPreference = 'Stop'
${buildConnectBlock(cfg)}
$org = Get-OrganizationConfig | Select-Object -ExpandProperty DisplayName
Write-Host "Connected to: $org"
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
`;
  return runPowerShellScript(script);
}
