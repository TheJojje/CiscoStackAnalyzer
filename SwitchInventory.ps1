#Requires -Version 5.1
<#
Cisco Switch Inventory v2.2
Requires: PuTTY Plink and ImportExcel (bundled in Modules\ImportExcel or installed for user).
CSV: Name,IPAddress,Cabinet
Core-list: coreswitches.txt (en switch per rad, # for kommentarer)
Interactive credentials per scan, no credentials.xml. Automatic TOFU for new CSV/IP targets.
Pins observed SHA256 fingerprints in hostkeys.json. Changed keys are never auto-accepted.
Credential is only in process memory; Plink receives it in a short-lived ACL-protected -pwfile.
#>
[CmdletBinding()]
param(
    [string]$CsvPath = (Join-Path $PSScriptRoot 'switches.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'Reports'),
    [string]$PlinkPath = (Join-Path $PSScriptRoot 'putty\plink.exe'),
    [string]$HostKeyDatabase = (Join-Path $PSScriptRoot 'hostkeys.json'),
    [string]$CoreSwitchPath = (Join-Path $PSScriptRoot 'coreswitches.txt')
)

$ErrorActionPreference = 'Stop'
# Core-switchar laddas fran coreswitches.txt. Ingen lista behovs i scriptet.
# Distribution devices are optional. All others default to Access.
$DistributionSwitches = @()
$script:CoreNames = @()
$script:Credentials = $null
$script:ShowProgress = $true

function Initialize-ImportExcel {
    # Load bundled ImportExcel from the script folder before searching the user/system modules.
    # Supports both Modules\ImportExcel\7.x.x\ImportExcel.psd1 and
    # Modules\ImportExcel\ImportExcel.psd1.
    if ((Get-Command Export-Excel -ErrorAction SilentlyContinue) -and
        (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) { return }

    $moduleRoot = Join-Path $PSScriptRoot 'Modules\ImportExcel'
    $localManifests = @()
    if (Test-Path -LiteralPath $moduleRoot -PathType Container) {
        $localManifests = @(Get-ChildItem -LiteralPath $moduleRoot -Filter 'ImportExcel.psd1' -File -Recurse -ErrorAction Stop |
            Sort-Object @{ Expression = {
                $folderName = $_.Directory.Name
                $parsedVersion = [version]'0.0'
                if ([version]::TryParse($folderName, [ref]$parsedVersion)) { return $parsedVersion }
                return [version]'0.0'
            }; Descending = $true })
    }

    $localFailure = $null
    foreach ($manifest in $localManifests) {
        try {
            Import-Module -Name $manifest.FullName -Force -ErrorAction Stop
            if ((Get-Command Export-Excel -ErrorAction SilentlyContinue) -and
                (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) {
                Write-Host "[OK] ImportExcel (lokal): $($manifest.DirectoryName)" -ForegroundColor Green
                return
            }
            throw 'ImportExcel laddades, men Export-Excel eller Open-ExcelPackage saknas.'
        }
        catch {
            $localFailure = $_.Exception.Message
            Write-Warning "Kunde inte ladda lokal ImportExcel: $localFailure"
        }
    }

    try {
        Import-Module -Name ImportExcel -ErrorAction Stop
        if ((Get-Command Export-Excel -ErrorAction SilentlyContinue) -and
            (Get-Command Open-ExcelPackage -ErrorAction SilentlyContinue)) {
            Write-Host '[OK] ImportExcel laddad fran installerade moduler.' -ForegroundColor Green
            return
        }
        throw 'Export-Excel eller Open-ExcelPackage saknas.'
    }
    catch {
        $detail = $_.Exception.Message
        if ($localFailure) { $detail += " | Lokalt fel: $localFailure" }
        throw "ImportExcel kan inte laddas. Placera hela modulen i '$moduleRoot' (eventuellt med versionsundermapp). Detaljer: $detail"
    }
}

function Update-CoreNames {
    if (-not (Test-Path -LiteralPath $CoreSwitchPath -PathType Leaf)) {
        throw "Core-listan saknas: $CoreSwitchPath. Skapa filen med ett hostname per rad."
    }
    $names = @(
        Get-Content -LiteralPath $CoreSwitchPath -ErrorAction Stop |
            ForEach-Object {
                $line = $_.Trim()
                if (-not $line -or $line.StartsWith('#')) { return }
                $short = (($line -split '\.')[0]).Trim().ToUpperInvariant()
                if ($short) { $short }
            } |
            Select-Object -Unique
    )
    if ($names.Count -eq 0) {
        throw "Core-listan ar tom: $CoreSwitchPath. Lagg in minst en core-switch."
    }
    $script:CoreNames = $names
}
function Edit-CoreSwitches {
    if (-not (Test-Path -LiteralPath $CoreSwitchPath -PathType Leaf)) {
        New-Item -ItemType File -Path $CoreSwitchPath -Force | Out-Null
    }
    Write-Host "Redigera filen, spara och stang Notepad: $CoreSwitchPath" -ForegroundColor Cyan
    $editor = Start-Process -FilePath notepad.exe -ArgumentList ('"{0}"' -f $CoreSwitchPath) -PassThru -Wait
    Update-CoreNames
    Write-Host ('[OK] Sparade core-switchar: ' + ($script:CoreNames -join ', ')) -ForegroundColor Green
}
function Get-ShortName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    return (($Name.Trim() -split '\.')[0]).ToUpperInvariant()
}
function Invoke-WithScanCredentials([scriptblock]$Action) {
    # Credential object is held only during this menu action. Never exported to disk.
    $sessionCredential = Get-Credential -Message 'Ange ditt Cisco/AAA-konto for denna korning (sparas inte)'
    if (-not $sessionCredential) { Write-Warning 'Inloggningen avbrots.'; return }
    if ([string]::IsNullOrWhiteSpace($sessionCredential.UserName)) { throw 'Anvandarnamn saknas.' }
    try {
        $script:Credentials = $sessionCredential
        & $Action
    }
    finally {
        $script:Credentials = $null
        $sessionCredential = $null
        [GC]::Collect()
    }
}
function Resolve-Plink {
    $candidates = @($PlinkPath, (Join-Path $PSScriptRoot 'plink.exe'), (Join-Path $PSScriptRoot 'putty\plink.exe'))
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }
    $cmd = Get-Command 'plink.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    throw "Plink hittades inte. Kontrollerade: $($candidates -join ', ')"
}
function New-ProtectedPasswordFile([System.Management.Automation.PSCredential]$Credential) {
    # Plink -pwfile avoids exposing the password in the command line.
    # A cleartext file exists briefly; do not place TEMP on shared/network storage.
    # A unique temporary folder is ACL-restricted BEFORE the password is written.
    $folder = Join-Path ([System.IO.Path]::GetTempPath()) ('CiscoInventory_' + [guid]::NewGuid().ToString('N'))
    $created = $false
    try {
        $sid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        $security = New-Object System.Security.AccessControl.DirectorySecurity
        $security.SetAccessRuleProtection($true, $false)
        $rule = [System.Security.AccessControl.FileSystemAccessRule]::new(
            $sid,
            [System.Security.AccessControl.FileSystemRights]::FullControl,
            [System.Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit',
            [System.Security.AccessControl.PropagationFlags]::None,
            [System.Security.AccessControl.AccessControlType]::Allow
        )
        $security.AddAccessRule($rule)
        [System.IO.Directory]::CreateDirectory($folder, $security) | Out-Null
        $created = $true
        $file = Join-Path $folder 'password.txt'
        $password = $Credential.GetNetworkCredential().Password
        if ([string]::IsNullOrEmpty($password)) { throw 'Losenordet ar tomt.' }
        if ($password.Contains("`r") -or $password.Contains("`n")) { throw 'Plink -pwfile stoder inte radbrytningar i losenord.' }
        [System.IO.File]::WriteAllText($file, ($password + "`n"), ([System.Text.UTF8Encoding]::new($false)))
        return [pscustomobject]@{ File=$file; Folder=$folder }
    } catch {
        if ($created) { Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction SilentlyContinue }
        throw
    } finally {
        $password = $null
    }
}
function Invoke-CiscoCommands([string]$IP, [string]$HostKey = '') {
    $plink = Resolve-Plink
    if (-not $script:Credentials) { throw 'Inga sessionscredentials. Starta scanningen via menyn.' }
    $secretFile = $null
    $p = $null
    try {
        $secretFile = New-ProtectedPasswordFile $script:Credentials
        # Initial Enter handles Cisco's 'Access granted. Press Return to begin session.'
        # Two more Enters help synchronize the shell before the commands.
        $cmds = "`n`nterminal length 0`nterminal width 511`nshow version`nshow switch`nshow inventory`nshow interfaces status`nshow cdp neighbors detail`nexit`n"
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName = $plink
        # ProcessStartInfo.Arguments quotes only safe paths/user/IP; secret is never an argument.
        $user = $script:Credentials.UserName
        if ($user -match '["\r\n]' -or $IP -notmatch '^[a-zA-Z0-9.:-]+$') { throw 'Ogiltigt anvandarnamn eller IP-adress.' }
        $hostKeyArg = ''
        if ($HostKey) {
            if ($HostKey -notmatch '^SHA256:[A-Za-z0-9+/]{43}$') { throw 'Ogiltigt SHA256-hostfingerprint.' }
            $hostKeyArg = ' -hostkey "' + $HostKey + '"'
        }
        $psi.Arguments = ('-ssh -batch -t -l "{0}" -pwfile "{1}"{2} "{3}"' -f $user, $secretFile.File, $hostKeyArg, $IP)
        $psi.UseShellExecute = $false
        $psi.RedirectStandardInput = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        $p = [System.Diagnostics.Process]::new()
        $p.StartInfo = $psi
        if (-not $p.Start()) { throw 'Kunde inte starta Plink.' }
        $outTask = $p.StandardOutput.ReadToEndAsync()
        $errTask = $p.StandardError.ReadToEndAsync()
        $p.StandardInput.Write($cmds)
        $p.StandardInput.Close()
        if (-not $p.WaitForExit(90000)) {
            try { $p.Kill() } catch {}
            throw 'SSH timeout (90 sekunder).'
        }
        $stdout = $outTask.GetAwaiter().GetResult()
        $stderr = $errTask.GetAwaiter().GetResult()
        if ($p.ExitCode -ne 0) {
            if ($stderr -match '(?i)(unknown option|unrecognised option|unknown argument).*pwfile') {
                throw 'Denna Plink-version saknar -pwfile. Uppdatera plink.exe till en nyare PuTTY-version.'
            }
            throw ("Plink exit {0}: {1}" -f $p.ExitCode, $stderr.Trim())
        }
        if ($stdout -match '(?im)^% ?(Invalid input|Error|Authorization failed)') { Write-Warning "$IP returnerade CLI-fel; kontrollera resultatet." }
        return $stdout.Replace("`r",'')
    } finally {
        if ($p) {
            if (-not $p.HasExited) { try { $p.Kill() } catch {} }
            $p.Dispose()
        }
        if ($secretFile) { Remove-Item -LiteralPath $secretFile.Folder -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
function Get-Section([string]$Raw, [string]$Command, [string]$NextCommand) {
    # Command echo appears as HOST#show ... (prompt may contain punctuation).
    $head = '(?im)^\s*[^\r\n]*?[#>]\s*' + [regex]::Escape($Command) + '\s*$'
    $next = '(?im)^\s*[^\r\n]*?[#>]\s*' + [regex]::Escape($NextCommand) + '\s*$'
    $m = [regex]::Match($Raw,$head)
    if (-not $m.Success) { return '' }
    # Find the next command echo after the current section.
    $tail = $Raw.Substring($m.Index + $m.Length)
    $n2 = [regex]::Match($tail,$next)
    if ($n2.Success) { return $tail.Substring(0,$n2.Index) }
    return $tail
}
function Parse-SwitchMembers([string]$Text) {
    $members = @{}
    foreach ($line in ($Text -split "`n")) {
        if ($line -match '^\s*\*?\s*(\d+)\s+(Active|Standby|Member|Ready|Provisioned)\s+([0-9a-fA-F\.:-]{12,})\s+(\d+)\s+(\S+)\s+(\S+)') {
            $id = [int]$Matches[1]
            $members[$id] = [pscustomobject]@{Number=$id;Role=$Matches[2];State=$Matches[6]}
        }
    }
    return $members
}
function Parse-Inventory([string]$Text) {
    $inventory = @{}
    # 2960X / IE-3000 / Rockwell: "1"; 9300: "Switch 1"; 9500: "Switch 1 Chassis".
    # Exact names exclude 9500 supervisor, power supply, optics and other FRU entries.
    $pattern = '(?im)^\s*NAME:\s*"(?:Switch\s+)?(\d+)(?:\s+Chassis)?"[^\r\n]*\r?\n\s*PID:\s*([^,\r\n]+),\s*VID:\s*[^,\r\n]*,\s*SN:\s*([^\s,\r\n]+)'
    foreach ($m in [regex]::Matches($Text,$pattern)) {
        $id = [int]$m.Groups[1].Value
        $inventory[$id] = [pscustomobject]@{Model=$m.Groups[2].Value.Trim(); Serial=$m.Groups[3].Value.Trim()}
    }
    return $inventory
}
function Get-AccessPortCount([object[]]$MemberPorts, [string]$Model) {
    # Industrial Ethernet devices are standalone; expansion cards do NOT create new
    # physical switches.  Count Fa access ports across all slots for IE-3000 / 1783.
    # IE-3300-8T2S has eight Gi copper access ports plus two Gi SFP uplinks.
    if ($Model -match '(?i)^IE-3300-8T2S(?:$|-)') {
        return @($MemberPorts | Where-Object { $_.Name -match '^(?:Gi|GigabitEthernet)1/[1-8]$' }).Count
    }
    if ($Model -match '(?i)^(?:IE-3000|1783-)') {
        $fastPorts = @($MemberPorts | Where-Object { $_.Name -match '^(?:Fa|FastEthernet)\d+/\d+$' })
        if ($fastPorts.Count -gt 0) { return $fastPorts.Count }
    }
    # Fixed-access platform capacities.  Do not confuse integrated uplink slots with
    # front-panel access ports (e.g. 2960X Gi1/0/49..52).
    if ($Model -match '(?i)^WS-C2960X-(8|12|24|48)(?:[A-Z0-9-]*)$') {
        return [int]$Matches[1]
    }
    if ($Model -match '(?i)^C9[23]00[A-Z0-9-]*?-(8|12|24|48)(?=[A-Z-])') {
        return [int]$Matches[1]
    }
    if ($Model -match '(?i)^WS-C(?:3560|3850)[A-Z0-9-]*?-(8|12|24|48)(?=[A-Z-])') {
        return [int]$Matches[1]
    }
    if ($Model -match '(?i)^C9500-(8|12|24|32|48)(?=[A-Z-])') {
        return [int]$Matches[1]
    }
    # Unknown models: use physical port observations, not any 'Ports' total in show version.
    $access = @($MemberPorts | Where-Object { $_.Module -eq 0 -and ($_.Port -le 48) })
    if ($access.Count -gt 0) { return $access.Count }
    return @($MemberPorts).Count
}
function Get-UplinkPortCount([object[]]$MemberPorts, [string]$Model, [int]$AccessCount) {
    # For platform inventory, uplinks are the remaining physical Ethernet interfaces.
    return [Math]::Max(0, (@($MemberPorts).Count - $AccessCount))
}
function Get-PhysicalPort([string]$Name) {
    if ($Name -match '^(Gi|GigabitEthernet|Fa|FastEthernet|Te|TenGigabitEthernet|Twe|TwentyFiveGigE|Fo|FortyGigabitEthernet|Hu|HundredGigE|Eth|Ethernet)(\d+)/(\d+)/(\d+)$') {
        return [pscustomobject]@{Member=[int]$Matches[2];Module=[int]$Matches[3];Port=[int]$Matches[4];Name=$Name}
    }
    # Two-component names are slots/modules on standalone IE-3000/IE-3300/1783,
    # NOT stack members: Fa2/1 belongs to the same chassis as Fa1/1.
    if ($Name -match '^(Gi|GigabitEthernet|Fa|FastEthernet|Te|TenGigabitEthernet)(\d+)/(\d+)$') {
        return [pscustomobject]@{Member=1;Module=[int]$Matches[2];Port=[int]$Matches[3];Name=$Name}
    }
    return $null
}
function Normalize-Port([string]$Port) {
    $aliases = @{
        'GigabitEthernet'='Gi'; 'FastEthernet'='Fa'; 'TenGigabitEthernet'='Te'
        'TwentyFiveGigE'='Twe'; 'FortyGigabitEthernet'='Fo'; 'HundredGigE'='Hu'
        'Ethernet'='Eth'
    }
    foreach ($key in ($aliases.Keys | Sort-Object Length -Descending)) {
        if ($Port -match ('^' + [regex]::Escape($key) + '(\d.*)$')) { return $aliases[$key] + $Matches[1] }
    }
    return $Port
}
function Parse-Ports([string]$Text) {
    $ports = @()
    foreach ($line in ($Text -split "`n")) {
        # First token is interface. The Name column can contain spaces, so do not split by fixed whitespace positions.
        if ($line -notmatch '^\s*(\S+)\s+') { continue }
        $name = $Matches[1]
        $info = Get-PhysicalPort $name
        if (-not $info) { continue }
        # Cisco status columns are fixed width, but interface descriptions may contain spaces.
        # Match the status word rather than assuming fixed whitespace tokens.
        $media = ''
        $status = ''
        if ($line -match '(?i)\s+(connected|notconnect|disabled|err-disabled|inactive|sfpAbsent|suspended|monitoring|faulty|notpresent|xcvrAbsen|up|down)\s+(\S+)\s+(\S+)\s+(\S+)\s+(.+?)\s*$') {
            $status = $Matches[1]
            $media = $Matches[5].Trim()
        }
        $ports += [pscustomobject]@{Name=$name;Member=$info.Member;Module=$info.Module;Port=$info.Port;Status=$status;Media=$media}
    }
    return @($ports | Sort-Object Name -Unique)
}
function Parse-CoreLinks([string]$Text) {
    $links = @()
    $blocks = [regex]::Split($Text,'(?m)^-{8,}\s*$')
    foreach ($block in $blocks) {
        if ($block -notmatch '(?im)^Device ID:\s*(\S+)') { continue }
        $device = $Matches[1]
        if ((Get-ShortName $device) -notin $script:CoreNames) { continue }
        if ($block -notmatch '(?im)^Interface:\s*([^,\r\n]+),\s*Port ID \(outgoing port\):\s*([^\r\n]+)') { continue }
        $local = $Matches[1].Trim(); $remote = $Matches[2].Trim()
        $physical = Get-PhysicalPort $local
        if (-not $physical) { continue }
        $address = ''
        if ($block -match '(?im)^\s*IP address:\s*(\d{1,3}(?:\.\d{1,3}){3})\s*$') { $address = $Matches[1] }
        $links += [pscustomobject]@{Core=(Get-ShortName $device);CoreIP=$address;LocalPort=$local;RemotePort=$remote;Member=$physical.Member}
    }
    return @($links | Sort-Object Core,LocalPort -Unique)
}
function Get-SwitchList {
    if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) { throw "CSV saknas: $CsvPath" }
    $items = @(Import-Csv -LiteralPath $CsvPath)
    if ($items.Count -eq 0) { throw 'CSV-filen ar tom.' }
    foreach ($field in @('Name','IPAddress','Cabinet')) {
        if ($field -notin $items[0].PSObject.Properties.Name) { throw "CSV saknar kolumnen $field." }
    }
    return $items
}
function Test-PuttyCachedHostKey([string]$HostIP, [int]$Port=22) {
    # PuTTY stores SSH host keys in HKCU, so no local admin privileges are needed.
    $regPath = 'Registry::HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\SshHostKeys'
    if (-not (Test-Path -LiteralPath $regPath)) { return $false }
    $reg = Get-Item -LiteralPath $regPath -ErrorAction Stop
    try {
        $suffix = '@{0}:{1}' -f $Port,$HostIP
        foreach ($keyName in $reg.GetValueNames()) {
            if ($keyName.EndsWith($suffix, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
        }
        return $false
    } finally {
        $reg.Close()
    }
}
# TOFU: trust a previously unseen host key on first connection, then pin it.
# A saved/changed key must NEVER be automatically replaced.
function Read-HostKeyDatabase {
    $map = @{}
    if (-not (Test-Path -LiteralPath $HostKeyDatabase -PathType Leaf)) { return $map }
    try {
        $raw = Get-Content -LiteralPath $HostKeyDatabase -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { throw 'Filen ar tom.' }
        $parsed = ConvertFrom-Json -InputObject $raw -ErrorAction Stop
        foreach ($prop in $parsed.PSObject.Properties) {
            if ($prop.Value -notmatch '^SHA256:[A-Za-z0-9+/]{43}$') {
                throw "Ogiltigt fingerprint for $($prop.Name)"
            }
            $map[[string]$prop.Name] = [string]$prop.Value
        }
        return $map
    } catch {
        throw "Kan inte lasa hostkeys.json. Ingen automatisk nyckelregistrering sker: $($_.Exception.Message)"
    }
}
function Write-HostKeyDatabase([hashtable]$Map) {
    $parent = Split-Path -Parent $HostKeyDatabase
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $json = ConvertTo-Json -InputObject $Map -Depth 3
    $tmp = "$HostKeyDatabase.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [System.IO.File]::WriteAllText($tmp, $json, [System.Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $tmp -Destination $HostKeyDatabase -Force
    } finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } }
}
function Invoke-CiscoWithAutoHostKey([string]$IP) {
    $map = Read-HostKeyDatabase
    $identity = "22@$IP"
    if ($map.ContainsKey($identity)) {
        # Even if PuTTY cached a different key, only our pinned value is accepted.
        return Invoke-CiscoCommands -IP $IP -HostKey $map[$identity]
    }
    try {
        # Cached PuTTY keys still work; no need to register those again.
        return Invoke-CiscoCommands -IP $IP
    } catch {
        $errorText = $_.Exception.Message
        # Strictly handle ONLY the explicit 'not cached' message. A changed
        # cached host key produces a different error and must abort.
        if ($errorText -notmatch '(?i)host key is not cached for this server') { throw }
        if (Test-PuttyCachedHostKey $IP) { throw "Nyckelproblem for $IP trots tidigare PuTTY-cache. Stoppar. $errorText" }
        $m = [regex]::Match($errorText, '(?i)ssh-(?:rsa|ed25519|dss)\s+\d+\s+(SHA256:[A-Za-z0-9+/]{43})')
        if (-not $m.Success) { throw "Kunde inte tolka hostfingerprint for $IP. Stoppar utan godkannande. $errorText" }
        $observed = $m.Groups[1].Value
        Write-Host "  [TOFU] Ny hostnyckel for $IP : $observed" -ForegroundColor Yellow
        # Plink -hostkey validates the SAME presented fingerprint while authenticating.
        $result = Invoke-CiscoCommands -IP $IP -HostKey $observed
        # Preserve the first pinned key if another process updated the database.
        $latest = Read-HostKeyDatabase
        if ($latest.ContainsKey($identity) -and $latest[$identity] -cne $observed) {
            throw "Hostfingerprint andrades under anslutning for $IP; sparar inte nyckeln."
        }
        $latest[$identity] = $observed
        Write-HostKeyDatabase $latest
        Write-Host "  [PINNED] $IP sparad i hostkeys.json" -ForegroundColor Green
        return $result
    }
}

function Register-HostKeys {
    $devices = @(Get-SwitchList)
    $plink = Resolve-Plink
    if (-not $script:Credentials) { throw 'Inga sessionscredentials. Starta scanningen via menyn.' }
    $user = $script:Credentials.UserName
    if ($user -match '["\r\n]') { throw 'Ogiltigt anvandarnamn.' }
    $pending = [System.Collections.Generic.List[object]]::new()
    foreach ($dev in $devices) {
        $hostIP = ([string]$dev.IPAddress).Trim()
        if ($hostIP -notmatch '^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$') {
            Write-Warning "Hoppar over ogiltig IP-adress: $hostIP"
            continue
        }
        if (-not (Test-PuttyCachedHostKey $hostIP)) { $pending.Add($dev) }
    }
    Write-Host ("`nPuTTY host keys: {0} registrerade, {1} saknas (av {2})." -f ($devices.Count-$pending.Count),$pending.Count,$devices.Count) -ForegroundColor Cyan
    if ($pending.Count -eq 0) { return }
    Write-Host 'Varje ny enhet maste verifieras mot ett BETROTT fingerprint (t.ex. via konsol/administrativ dokumentation).' -ForegroundColor Yellow
    Write-Host 'Plink visar fingeravtrycket. Ange y i Plinks prompt ENDAST om det matchar.' -ForegroundColor Yellow
    Write-Host 'Tryck Enter i Plinks prompt utan y om du inte kan verifiera nyckeln.' -ForegroundColor Yellow
    $answer = Read-Host 'Starta interaktiv registrering av saknade host keys? (J/N)'
    if ($answer -notmatch '^(?i:j|ja|y|yes)$') { return }
    $secretFile = $null
    $added = 0; $skipped = 0
    try {
        $secretFile = New-ProtectedPasswordFile $script:Credentials
        $index = 0
        foreach ($dev in $pending) {
            $index++
            $hostIP = ([string]$dev.IPAddress).Trim()
            $devName = [string]$dev.Name
            Write-Host ("`n[{0}/{1}] {2} ({3})" -f $index,$pending.Count,$devName,$hostIP) -ForegroundColor Cyan
            $action = (Read-Host 'V=verifiera via Plink, H=hoppa over, A=avbryt').Trim().ToUpperInvariant()
            if ($action -eq 'A') { break }
            if ($action -ne 'V') { $skipped++; continue }
            Write-Host 'Plink startar nu. Kontrollera fingerprintet innan du eventuellt godkanner (y).' -ForegroundColor Yellow
            # Plink is deliberately NOT run with -batch: only the operator can accept a host key.
            # A remote command is provided so that successful sessions exit on their own.
            # Some IOS versions reject remote exec, but the host key may still be cached.
            & $plink -ssh -l $user -pwfile $secretFile.File $hostIP 'exit'
            if (Test-PuttyCachedHostKey $hostIP) {
                $added++
                Write-Host "  Registrerad: $devName ($hostIP)" -ForegroundColor Green
            } else {
                Write-Warning "Nyckeln finns inte i PuTTY-cache for $hostIP. Inget har registrerats."
                $skipped++
            }
        }
    } finally {
        if ($secretFile) { Remove-Item -LiteralPath $secretFile.Folder -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Write-Host ("`nFardigt: {0} registrerade, {1} hoppades over/ej registrerade." -f $added,$skipped) -ForegroundColor Cyan
}
function Scan-Devices([object[]]$SelectedDevices = @()) {
    Update-CoreNames
    if (-not $script:Credentials) { throw 'Inga sessionscredentials. Starta scanningen via menyn.' }
    Initialize-ImportExcel
    $devices = if ($SelectedDevices.Count -gt 0) { @($SelectedDevices) } else { @(Get-SwitchList) }
    $report = [System.Collections.Generic.List[object]]::new()
    $overview = [System.Collections.Generic.List[object]]::new()
    $membersOut = [System.Collections.Generic.List[object]]::new()
    $coreOut = [System.Collections.Generic.List[object]]::new()
    $logsOut = [System.Collections.Generic.List[object]]::new()
    $i=0
    foreach ($dev in $devices) {
        $i++
        $name = [string]$dev.Name; $ip=[string]$dev.IPAddress; $cabinet=[string]$dev.Cabinet
        Write-Host ("[{0}/{1}] {2} ({3})" -f $i,$devices.Count,$name,$ip) -ForegroundColor Cyan
        try {
            if ([string]::IsNullOrWhiteSpace($ip)) { throw 'IP-adress saknas.' }
            $raw = Invoke-CiscoWithAutoHostKey $ip
            $versionText = Get-Section $raw 'show version' 'show switch'
            $switchText = Get-Section $raw 'show switch' 'show inventory'
            $invText = Get-Section $raw 'show inventory' 'show interfaces status'
            $portText = Get-Section $raw 'show interfaces status' 'show cdp neighbors detail'
            $cdpText = Get-Section $raw 'show cdp neighbors detail' 'exit'
            if (-not $invText -or -not $portText) { throw 'CLI-output kunde inte avgransas (show inventory/interfaces status). Kontrollera SSH-prompt eller kommandobehorighet.' }
            $stack = Parse-SwitchMembers $switchText
            $inventory = Parse-Inventory $invText
            # Standalone 3560/3850 may use a different inventory chassis name.
            # Only fall back when no member entries were found, never invent stack members.
            if ($inventory.Count -eq 0) {
                $chassis = [regex]::Match($invText, '(?im)^\s*NAME:\s*"(?:Chassis|[^"]*3560[^"]*|[^"]*3850[^"]*)"[^\r\n]*\r?\n\s*PID:\s*([^,\r\n]+),\s*VID:\s*[^,\r\n]*,\s*SN:\s*([^\s,\r\n]+)')
                if ($chassis.Success -and $chassis.Groups[1].Value -match '^(?:WS-C3560|WS-C3850|C3560|C3850)') {
                    $inventory[1] = [pscustomobject]@{ Model=$chassis.Groups[1].Value.Trim(); Serial=$chassis.Groups[2].Value.Trim() }
                }
            }
            # Standalone industrial chassis: IE-3300 uses NAME: "Chassis" and a
            # duplicate "Supervisor" entry with the same serial number.
            # Take Chassis only.  The 1783 expansion module is not a stack member.
            if ($inventory.Count -eq 0) {
                $industrial = [regex]::Match($invText, '(?im)^\s*NAME:\s*"Chassis"[^\r\n]*\r?\n\s*PID:\s*([^,\r\n]+),\s*VID:\s*[^,\r\n]*,\s*SN:\s*([^\s,\r\n]+)')
                if ($industrial.Success -and $industrial.Groups[1].Value.Trim() -match '(?i)^IE-') {
                    $inventory[1] = [pscustomobject]@{ Model=$industrial.Groups[1].Value.Trim(); Serial=$industrial.Groups[2].Value.Trim() }
                }
            }
            if ($inventory.Count -eq 0 -and $versionText -match '(?im)^\s*(?:Model number|Model Number)\s*:\s*(\S+)') {
                $modelFallback = $Matches[1]
                if ($modelFallback -match '(?i)^(?:WS-C3560|WS-C3850|C3560|C3850|IE-3000|IE-3300|1783-)' -and $versionText -match '(?im)^\s*(?:System serial number|System Serial Number)\s*:\s*(\S+)') {
                    $inventory[1] = [pscustomobject]@{ Model=$modelFallback; Serial=$Matches[1] }
                }
            }
            $ports = @(Parse-Ports $portText)
            $links = @(Parse-CoreLinks $cdpText)
            if ($inventory.Count -eq 0) { throw 'Hittade inga individuella switchar i show inventory.' }
            $memberIds = @($inventory.Keys | Sort-Object)
            $total = 0; $accessTotal = 0; $uplinkTotal = 0
            foreach ($id in $memberIds) {
                $all = @($ports | Where-Object Member -eq $id)
                $access = Get-AccessPortCount $all $inventory[$id].Model
                $uplink = Get-UplinkPortCount $all $inventory[$id].Model $access
                $total += $all.Count; $accessTotal += $access; $uplinkTotal += $uplink
                $role = ''; $state = ''
                if ($stack.ContainsKey([int]$id)) { $role=$stack[[int]$id].Role; $state=$stack[[int]$id].State }
                $membersOut.Add([pscustomobject]@{
                    Name=$name;IPAddress=$ip;Cabinet=$cabinet;Member=[int]$id;Role=$role;State=$state
                    Model=$inventory[$id].Model;SerialNumber=$inventory[$id].Serial
                    AccessPorts=$access;UplinkPorts=$uplink;TotalPorts=$all.Count
                })
            }
            foreach ($id in $memberIds) {
                $all = @($ports | Where-Object Member -eq $id)
                $accessCount = Get-AccessPortCount $all $inventory[$id].Model
                $roleName = 'Access'
                $shortName = Get-ShortName $name
                if ($shortName -in $script:CoreNames) { $roleName = 'Core' }
                elseif ($shortName -in @($DistributionSwitches | ForEach-Object { Get-ShortName $_ })) { $roleName = 'Distribution' }
                $memberLinks = @($links | Where-Object Member -eq $id | Sort-Object Core,LocalPort)
                $coreNames = @($memberLinks | ForEach-Object { $_.Core })
                $coreIPs = @($memberLinks | ForEach-Object { $_.CoreIP })
                $uplinkPorts = @($memberLinks | ForEach-Object { (Normalize-Port $_.LocalPort) })
                $mediaTypes = @()
                foreach ($link in $memberLinks) {
                    $normalized = Normalize-Port $link.LocalPort
                    $portRec = $all | Where-Object { (Normalize-Port $_.Name) -eq $normalized } | Select-Object -First 1
                    $mediaTypes += $(if ($portRec -and $portRec.Media) { $portRec.Media } else { 'Unknown' })
                }
                $report.Add([pscustomobject][ordered]@{
                    'Serial number'=$inventory[$id].Serial
                    'Device Name'=$name
                    'IP Address'=$ip
                    'Model'=$inventory[$id].Model
                    'Stack Switch'=[int]$id
                    'Port Count'="${accessCount} Port"
                    'Switch Role'=$roleName
                    'Upstream Connected Switch Name Details'=($coreNames -join "`n")
                    'Upstream IP Address'=($coreIPs -join "`n")
                    'Uplink Port #s'=($uplinkPorts -join "`n")
                    'SFP used'=($mediaTypes -join "`n")
                })
            }
            foreach ($link in $links) {
                $coreOut.Add([pscustomobject]@{Name=$name;IPAddress=$ip;Cabinet=$cabinet;Member=$link.Member;LocalPort=$link.LocalPort;CoreSwitch=$link.Core;CoreIP=$link.CoreIP;CorePort=$link.RemotePort})
            }
            $overview.Add([pscustomobject]@{Name=$name;IPAddress=$ip;Cabinet=$cabinet;StackMembers=$memberIds.Count;AccessPorts=$accessTotal;UplinkPorts=$uplinkTotal;TotalPorts=$total;CoreLinks=$links.Count;CoreSwitches=(@($links.Core | Sort-Object -Unique) -join ', ');Result='OK'})
            $logsOut.Add([pscustomobject]@{Time=(Get-Date);Name=$name;IPAddress=$ip;Result='OK';Detail="Medlemmar: $($memberIds.Count); portar: $total; core: $($links.Count)"})
            Write-Host "  OK - $($memberIds.Count) medlemmar, $total portar, $($links.Count) core-uplinks" -ForegroundColor Green
        } catch {
            $msg = $_.Exception.Message
            $overview.Add([pscustomobject]@{Name=$name;IPAddress=$ip;Cabinet=$cabinet;StackMembers=0;AccessPorts=0;UplinkPorts=0;TotalPorts=0;CoreLinks=0;CoreSwitches='';Result='ERROR'})
            $logsOut.Add([pscustomobject]@{Time=(Get-Date);Name=$name;IPAddress=$ip;Result='ERROR';Detail=$msg})
            Write-Warning "$name : $msg"
        }
    }
    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
    $path = Join-Path $OutputFolder ('SwitchInventory_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    # Ensure even empty sheets have column headers.
    if ($membersOut.Count -eq 0) { $membersOut.Add([pscustomobject]@{Name='';IPAddress='';Cabinet='';Member='';Role='';State='';Model='';SerialNumber='';AccessPorts='';UplinkPorts='';TotalPorts=''}) }
    if ($coreOut.Count -eq 0) { $coreOut.Add([pscustomobject]@{Name='';IPAddress='';Cabinet='';Member='';LocalPort='';CoreSwitch='';CoreIP='';CorePort=''}) }
    if ($report.Count -eq 0) {
        $report.Add([pscustomobject][ordered]@{'Serial number'='';'Device Name'='';'IP Address'='';'Model'='';'Stack Switch'='';'Port Count'='';'Switch Role'='';'Upstream Connected Switch Name Details'='';'Upstream IP Address'='';'Uplink Port #s'='';'SFP used'=''})
    }
    @($report.ToArray()) | Export-Excel -Path $path -WorksheetName 'Inventory' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    # Set practical widths and wrap multiline uplink details while preserving row alignment.
    $excel = Open-ExcelPackage -Path $path
    try {
        $ws = $excel.Workbook.Worksheets['Inventory']
        $widths = @(22,24,18,24,15,14,19,27,18,18,25)
        for ($c=1; $c -le 11; $c++) {
            $ws.Column($c).Width = $widths[$c-1]
            $ws.Column($c).Style.WrapText = $true
            $ws.Column($c).Style.VerticalAlignment = 'Center'
        }
        for ($r=2; $r -le ($report.Count + 1); $r++) {
            $lineCount = [Math]::Max(1, ($ws.Cells[$r,8].Text -split "`n").Count)
            $ws.Row($r).Height = [Math]::Max(20, $lineCount * 17)
        }
        Close-ExcelPackage $excel
    } catch {
        Close-ExcelPackage $excel -NoSave
        throw
    }
    @($overview.ToArray()) | Export-Excel -Path $path -WorksheetName 'Overview' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Append
    @($membersOut.ToArray()) | Export-Excel -Path $path -WorksheetName 'StackMembers' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Append
    @($coreOut.ToArray()) | Export-Excel -Path $path -WorksheetName 'CoreUplinks' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Append
    @($logsOut.ToArray()) | Export-Excel -Path $path -WorksheetName 'ScanLog' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Append
    Write-Host "`nExcelrapport: $path" -ForegroundColor Green
}
function Scan-One {
    $selection = (Read-Host 'Ange namn fran CSV eller IP-adress').Trim()
    if ([string]::IsNullOrWhiteSpace($selection)) { Write-Warning 'Ingen switch angavs.'; return }
    $devices = @()
    if (Test-Path -LiteralPath $CsvPath -PathType Leaf) { $devices = @(Get-SwitchList) }
    $matched = @($devices | Where-Object {
        ([string]$_.Name).Trim().Equals($selection, [System.StringComparison]::OrdinalIgnoreCase) -or
        ([string]$_.IPAddress).Trim().Equals($selection, [System.StringComparison]::OrdinalIgnoreCase)
    })
    if ($matched.Count -gt 1) { Write-Warning 'Flera CSV-poster matchar valet.'; return }
    if ($matched.Count -eq 1) {
        Scan-Devices -SelectedDevices $matched
        return
    }
    $ipObj = $null
    if (-not [System.Net.IPAddress]::TryParse($selection, [ref]$ipObj)) {
        Write-Warning "Hittar inte '$selection' i switches.csv. Ange en IPv4/IPv6-adress for en okand enhet."
        return
    }
    $deviceName = (Read-Host 'Enhetsnamn (frivilligt, annars IP)').Trim()
    if (-not $deviceName) { $deviceName = $selection }
    $cabinet = (Read-Host 'Skap (frivilligt)').Trim()
    Scan-Devices -SelectedDevices @([pscustomobject]@{Name=$deviceName;IPAddress=$selection;Cabinet=$cabinet})
}
function Show-Menu {
    while ($true) {
        Write-Host "`n=== Cisco Switch Inventory ===" -ForegroundColor Cyan
        Write-Host '1. Registrera SSH-hostnycklar (valfri manuell verifiering)'
        Write-Host '2. Scanna alla switchar och exportera Excel (auto-TOFU)'
        Write-Host '3. Scanna en specifik switch (CSV eller IP, auto-TOFU)'
        Write-Host '4. Visa core-switchar'
        Write-Host '5. Redigera coreswitches.txt (sparas permanent)'
        Write-Host '6. Om autentisering och sakerhet'
        Write-Host '7. Avsluta'
        switch ((Read-Host 'Val').Trim()) {
            '1' { try { Invoke-WithScanCredentials { Register-HostKeys } } catch { Write-Warning $_.Exception.Message } }
            '2' { try { Invoke-WithScanCredentials { Scan-Devices } } catch { Write-Warning $_.Exception.Message } }
            '3' { try { Invoke-WithScanCredentials { Scan-One } } catch { Write-Warning $_.Exception.Message } }
            '4' { try { Update-CoreNames; Write-Host ('Core: ' + ($script:CoreNames -join ', ')) } catch { Write-Warning $_.Exception.Message } }
            '5' { try { Edit-CoreSwitches } catch { Write-Warning $_.Exception.Message } }
            '6' { Write-Host 'Kontot anges vid varje korning. credentials.xml anvands inte. Plink -pwfile skapar en temporar klartextfil med restriktiv ACL per anslutning och tar bort den i finally. Filrensning ar ingen garanti mot aterstallning ur disk/minne.' -ForegroundColor Yellow }
            '7' { return }
            default { Write-Warning 'Ogiltigt val.' }
        }
    }
}
Show-Menu
