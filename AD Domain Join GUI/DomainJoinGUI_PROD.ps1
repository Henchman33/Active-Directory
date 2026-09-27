<#
Domain Join GUITool.ps1 - Domain Join & Provisioning GUI
Requires -Version 5.1
Requires -RunAsAdministrator
#>

<#
.SYNOPSIS
    Domain Join & Provisioning Tool - GUI for joining devices to AD domains,
    installing software, configuring networking, and launching admin tools.

.DESCRIPTION
    A Windows Forms GUI tool that provides:
    - Multi-domain selection dropdown
    - Credential entry for domain join
    - Computer rename and domain join/change
    - Software installation from C:\IT with browse support
    - IPv4 network configuration with NIC chooser
    - Quick-launch buttons for CMD (admin), Notepad, Computer Management
    - Configurable log file location (Domain Join GUI.log)
    - Silent operation with detailed logging

.NOTES
    Author : Steve McKee
    Created: 2026-09-26
    Requirements: PowerShell 5.1+, Run as Administrator, RSAT AD tools optional
#>

# ============================================================================
# REGION: Assembly Loading
# ============================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName PresentationFramework

# Enable modern Windows 10/11 visual styles
try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
} catch {}

# ============================================================================
# REGION: Global Variables & Configuration
# ============================================================================

# --- EDIT THIS LIST to add your domain names ---
# Add or remove domains as needed - these populate the dropdown
$script:DomainList = @(
    'MYIGT.COM'
    'IGTSAP.AD.IGT.COM'
    'AD.IGT.COM'
    'IS.AD.IGT.COM'
    'HQ.GLOBALCASHACCESS.US'
    'EVERI.GLOBALCASHACCESS.US'
    'COPPERSYNC.COM'
    'CCR.COPPERSYNC.COM'
    'TITANIUM.COM'
)

# --- EDIT THIS LIST to add your Domain Controllers ---
# Map each DC hostname to a friendly location label.
# The dropdown shows: "Location - hostname"
# Leave the list empty to auto-resolve DCs via DNS.
$script:DomainControllerList = @(
    [PSCustomObject]@{ Label = 'Reno - L-IGT';        Host = 'USRNOPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Graz';                Host = 'ATUPSPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Austria';             Host = 'AUSYDPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Azure';               Host = 'AZSIWADIGT01P.myigt.com' }
    [PSCustomObject]@{ Label = 'Moncton';             Host = 'CAMNTPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Beijing';             Host = 'CNBJSPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Las Vegas - L-Everi'; Host = 'LASWDCIGT01P.myigt.com' }
    [PSCustomObject]@{ Label = 'Amsterdam';           Host = 'NLAMSPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Belgrade';            Host = 'RSBEGPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Las Vegas - L-IGT';   Host = 'USLASPADDSI01.myigt.com' }
    [PSCustomObject]@{ Label = 'Reno';                Host = 'USRNOPADDSI01-N.myigt.com' }
    [PSCustomObject]@{ Label = 'Reno';                Host = 'USRNOPADDSI02-N.myigt.com' }
    [PSCustomObject]@{ Label = 'Reno';                Host = 'USRNOPADDSI03-N.myigt.com' }
)

# Default software source folder
$script:DefaultSoftwarePath = 'C:\IT'

# Log file name
$script:LogFileName = 'DomainJoinGUI.log'

# Global log path - set by user via GUI
$script:LogPath = $null

# NIC list cache
$script:NicList = @()

# Software package definitions - display names map to patterns for detection
# SilentSwitches: default per-package switches used for .exe installers.
# Leave empty to use the generic "/S /norestart" default.
# For .msi files, /qn /norestart is always used (standard MSI behavior).
# Users can override these per-package in the GUI DataGridView.
$script:SoftwarePackages = @(
    [PSCustomObject]@{ Name = 'CrowdStrike';       Pattern = 'CrowdStrike*';     SilentSwitches = '/install /quiet /norestart' }
    [PSCustomObject]@{ Name = 'Rapid7';            Pattern = 'Rapid7*';          SilentSwitches = '/q /norestart' }
    [PSCustomObject]@{ Name = 'Cribl';             Pattern = 'Cribl*';           SilentSwitches = '/S /norestart' }
    [PSCustomObject]@{ Name = 'MEMCM/SCCM Client'; Pattern = 'ccmsetup*';        SilentSwitches = '' }
    [PSCustomObject]@{ Name = 'CMTrace';           Pattern = 'cmtrace*';          SilentSwitches = '' }
)

# --- Environment-specific configuration for IGT / Legacy-Everi ---
# CrowdStrike Falcon Sensor CID
$script:CrowdStrikeCID = '1B406277AB3D424D9A4C536A1128F9C3-BD'

# SCCM site configurations
# Tier 1/2 L-Everi: SMSSITECODE=TIT, MP=LASAPPSCCM01P.titanium.com
# Tier 3/4 L-Everi: SMSSITECODE=GHQ, MP=LASAPPSCM01P.hq.globalcashaccess.us
$script:SCCMConfigs = @(
    [PSCustomObject]@{ Tier = 'Tier 1/2 L-Everi'; SiteCode = 'TIT'; MP = 'LASAPPSCCM01P.titanium.com'; Share = '\\LASAPPSCCM01P.titanium.com\SMS_TIT\Client' }
    [PSCustomObject]@{ Tier = 'Tier 3/4 L-Everi'; SiteCode = 'GHQ'; MP = 'LASAPPSCM01P.hq.globalcashaccess.us'; Share = '\\LASAPPSCM01P.hq.globalcashaccess.us\SMS_GHQ\Client' }
)

# CRIBL configuration scripts per environment
$script:CriblScripts = @(
    [PSCustomObject]@{ Env = 'Production L-Everi';  Script = 'Production_L-Everi.ps1' }
    [PSCustomObject]@{ Env = 'Non-Production L-Everi'; Script = 'Non-Production_L-Everi.ps1' }
    [PSCustomObject]@{ Env = 'IGT';                Script = 'IGT.ps1' }
)

# Admin accounts to set password non-expire
$script:AdminAccounts = @('Administrator', 'sysengadm')

# IISCrypto template path
$script:IISCryptoTemplate = 'C:\IT\Everi_24June2025.ictpl'

# Rapid7 ScanAssist certificate (embedded for automation)
$script:Rapid7Cert = '-----BEGIN CERTIFICATE----- MIIBXzCB5gIIPCHRQf/4Sq8wCgYIKoZIzj0EAwIwGjEYMBYGA1UEAwwPUjdTY2Fu QXNzaXN0YW50MB4XDTIzMDYwMjIxMTA0MVoXDTI2MDYwMTIxMTA0MVowGjEYMBYG A1UEAwwPUjdTY2FuQXNzaXN0YW50MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEghWe 9pldxxBF2GAGoqYmfRUlN/QV5YLpO8HULNu38sHzopuJhGZ3mHBJ5UUQNGE2XV19 YDOB+6MJI2DwiGcKJCtB2pQcYWmweq3oGQSxsrzJSA30e3hZnPuSsM9X+LaDMAoG CCqGSM49BAMCA2gAMGUCMQC+ncMz7zuUBDxJ3JrcvIOFlrS/5TPu13uW4iOJ4G0V MsRIGwBqsMf+6wM6MdUNTKkCMCljztoLeMOvq1plvGjkihkRSL8WhdAoJazz7Kj9 Zhu77rY5q+h26NYSVYZMF0biFQ== -----END CERTIFICATE-----'

# BGInfo startup folder path
$script:BGInfoStartupPath = 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup'
$script:BGInfoFolder = 'C:\IT\BGInfo'

# CRIBL scripts folder
$script:CriblFolder = 'C:\IT\CRIBL'

# Rapid7 ScanAssist folder
$script:Rapid7Folder = 'C:\IT\Rapid7_manual'

# CrowdStrike folder
$script:CrowdStrikeFolder = 'C:\IT\CrowdStrike'

# IT folder base
$script:ITFolder = 'C:\IT'

# ============================================================================
# REGION: Logging Functions
# ============================================================================

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped entry to the Domain Join GUI.log file and optionally to console.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS','DEBUG')]
        [string]$Level = 'INFO'
    )

    if ([string]::IsNullOrWhiteSpace($script:LogPath)) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"

    try {
        $logDir = Split-Path -Path $script:LogPath -Parent
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        Add-Content -Path $script:LogPath -Value $logEntry -ErrorAction Stop
    } catch {
        # Silent fail on logging errors to avoid disrupting the workflow
    }
}

function Write-LogAndConsole {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS','DEBUG')]
        [string]$Level = 'INFO'
    )
    Write-Log -Message $Message -Level $Level
    $color = switch ($Level) {
        'INFO'    { 'White' }
        'WARNING' { 'Yellow' }
        'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }
        'DEBUG'   { 'Gray' }
    }
    Write-Host $Message -ForegroundColor $color
}

# ============================================================================
# REGION: Helper Functions
# ============================================================================

function Get-NetworkAdapters {
    <#
    .SYNOPSIS
        Returns a list of physical and virtual network adapters with IPv4 enabled.
    #>
    try {
        $adapters = Get-NetAdapter -ErrorAction Stop |
            Where-Object { $_.Status -eq 'Up' -or $_.Status -eq 'Disconnected' } |
            Select-Object Name, InterfaceDescription, Status, MacAddress, ifIndex |
            Sort-Object Name
        return $adapters
    } catch {
        Write-Log -Message "Failed to enumerate network adapters: $($_.Exception.Message)" -Level 'ERROR'
        return @()
    }
}

function Set-Ipv4Config {
    <#
    .SYNOPSIS
        Configures IPv4 settings on a specified network adapter.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$AdapterName,
        [string]$IPAddress,
        [string]$SubnetMask,
        [string]$Gateway,
        [string[]]$DNSServers
    )

    $errors = @()

    # Validate inputs
    if ([string]::IsNullOrWhiteSpace($IPAddress)) {
        $errors += "IP Address is required."
    }
    if ([string]::IsNullOrWhiteSpace($SubnetMask)) {
        $errors += "Subnet Mask is required."
    }
    if ([string]::IsNullOrWhiteSpace($Gateway)) {
        $errors += "Default Gateway is required."
    }

    # Validate IP format
    $ipRegex = '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    if ($IPAddress -and $IPAddress -notmatch $ipRegex) { $errors += "Invalid IP Address format." }
    if ($Gateway -and $Gateway -notmatch $ipRegex) { $errors += "Invalid Gateway format." }

    foreach ($dns in $DNSServers) {
        if ($dns -and $dns -notmatch $ipRegex) { $errors += "Invalid DNS server: $dns" }
    }

    if ($errors.Count -gt 0) {
        return [PSCustomObject]@{ Success = $false; Errors = $errors }
    }

    try {
        # Convert subnet mask to prefix length
        $prefix = $SubnetMask.Split('.') |
            ForEach-Object { [Convert]::ToString([int]$_, 2) } |
            ForEach-Object { $_.PadLeft(8, '0') } |
            Join-String
        $prefixLength = ($prefix -replace '0', '').Length

        Write-Log -Message "Configuring NIC '$AdapterName': IP=$IPAddress/$prefixLength, GW=$Gateway, DNS=$($DNSServers -join ', ')" -Level 'INFO'

        # Remove existing IP configuration
        $existingIPs = Get-NetIPAddress -InterfaceAlias $AdapterName -AddressFamily IPv4 -ErrorAction SilentlyContinue
        foreach ($ip in $existingIPs) {
            Remove-NetIPAddress -IPAddress $ip.IPAddress -InterfaceAlias $AdapterName -Confirm:$false -ErrorAction SilentlyContinue
        }

        # Set new IP and gateway
        $newIP = New-NetIPAddress -InterfaceAlias $AdapterName -IPAddress $IPAddress -PrefixLength $prefixLength -DefaultGateway $Gateway -ErrorAction Stop

        # Set DNS servers
        if ($DNSServers -and $DNSServers.Count -gt 0) {
            Set-DnsClientServerAddress -InterfaceAlias $AdapterName -ServerAddresses $DNSServers -ErrorAction Stop
        }

        Write-Log -Message "Successfully configured IPv4 on '$AdapterName'" -Level 'SUCCESS'
        return [PSCustomObject]@{ Success = $true; Errors = @() }
    } catch {
        Write-Log -Message "Failed to configure IPv4 on '$AdapterName': $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Errors = @($_.Exception.Message) }
    }
}

function Invoke-ComputerRename {
    <#
    .SYNOPSIS
        Renames the computer.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$NewName
    )

    if ([string]::IsNullOrWhiteSpace($NewName)) {
        return [PSCustomObject]@{ Success = $false; Message = "Computer name cannot be empty." }
    }

    # Validate computer name (NetBIOS rules: 15 chars max, alphanumeric and hyphens)
    if ($NewName.Length -gt 15) {
        return [PSCustomObject]@{ Success = $false; Message = "Computer name exceeds 15 characters (NetBIOS limit)." }
    }
    if ($NewName -notmatch '^[A-Za-z0-9][A-Za-z0-9-]*[A-Za-z0-9]$|^[A-Za-z0-9]$') {
        return [PSCustomObject]@{ Success = $false; Message = "Computer name contains invalid characters. Use letters, numbers, and hyphens only." }
    }

    try {
        Write-Log -Message "Renaming computer from '$env:COMPUTERNAME' to '$NewName'" -Level 'INFO'
        Rename-Computer -NewName $NewName -Force -PassThru -ErrorAction Stop | Out-Null
        Write-Log -Message "Computer rename pending. Will take effect after reboot." -Level 'SUCCESS'
        return [PSCustomObject]@{ Success = $true; Message = "Computer renamed to '$NewName'. Reboot required." }
    } catch {
        Write-Log -Message "Failed to rename computer: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-DomainJoin {
    <#
    .SYNOPSIS
        Joins the computer to an Active Directory domain using Add-Computer.
    .DESCRIPTION
        Supports targeting a specific Domain Controller via -Server to bypass
        site affinity and avoid fallback to provisioning joins.
        Optionally enables NetJoinLegacyAccountReuse registry key before joining.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DomainName,
        [Parameter(Mandatory)]
        [string]$Username,
        [Parameter(Mandatory)]
        [string]$Password,
        [string]$NewComputerName,
        [string]$OUPath,
        [string]$Server,
        [switch]$EnableLegacyReuse,
        [switch]$ChangeDomain
    )

    $errors = @()

    if ([string]::IsNullOrWhiteSpace($DomainName)) { $errors += "Domain name is required." }
    if ([string]::IsNullOrWhiteSpace($Username)) { $errors += "Username is required." }
    if ([string]::IsNullOrWhiteSpace($Password)) { $errors += "Password is required." }

    if ($errors.Count -gt 0) {
        return [PSCustomObject]@{ Success = $false; Errors = $errors; Message = ($errors -join "`n") }
    }

    try {
        # Build credential object
        $securePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $securePassword

        # Enable NetJoinLegacyAccountReuse if requested
        if ($EnableLegacyReuse) {
            Write-Log -Message 'Enabling NetJoinLegacyAccountReuse registry key...' -Level 'INFO'
            try {
                $regPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters'
                if (-not (Test-Path -Path $regPath)) {
                    New-Item -Path $regPath -Force | Out-Null
                }
                Set-ItemProperty -Path $regPath -Name 'NetJoinLegacyAccountReuse' -Value 1 -Type DWord -ErrorAction Stop
                Write-Log -Message 'NetJoinLegacyAccountReuse set to 1.' -Level 'SUCCESS'
            } catch {
                Write-Log -Message "Failed to set NetJoinLegacyAccountReuse: $($_.Exception.Message)" -Level 'WARNING'
            }
        }

        $params = @{
            DomainName  = $DomainName
            Credential  = $credential
            Force       = $true
            Options     = 'JoinWithNewName', 'AccountCreate'
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($NewComputerName)) {
            $params['NewName'] = $NewComputerName
        }
        if (-not [string]::IsNullOrWhiteSpace($OUPath)) {
            $params['OUPath'] = $OUPath
        }
        if (-not [string]::IsNullOrWhiteSpace($Server)) {
            $params['Server'] = $Server
            Write-Log -Message "Targeting specific Domain Controller: $Server" -Level 'INFO'
        }

        # If changing domain, need to unjoin current domain first
        if ($ChangeDomain) {
            Write-Log -Message "Unjoining current domain/workgroup before joining new domain..." -Level 'INFO'
            try {
                Remove-Computer -UnjoinDomainCredential $credential -Force -ErrorAction Stop
                Write-Log -Message "Successfully unjoined from current domain." -Level 'SUCCESS'
            } catch {
                Write-Log -Message "Note: Unjoin may have partially failed (this can be normal if not domain-joined): $($_.Exception.Message)" -Level 'WARNING'
            }
        }

        Write-Log -Message "Joining computer to domain '$DomainName' with account '$Username'..." -Level 'INFO'

        $result = Add-Computer @params -PassThru

        if ($result.HasSucceeded) {
            Write-Log -Message "Successfully joined domain '$DomainName'. Reboot required." -Level 'SUCCESS'
            return [PSCustomObject]@{ Success = $true; Errors = @(); Message = "Successfully joined domain '$DomainName'. Reboot required." }
        } else {
            $msg = "Domain join reported failure."
            Write-Log -Message $msg -Level 'ERROR'
            return [PSCustomObject]@{ Success = $false; Errors = @($msg); Message = $msg }
        }
    } catch {
        Write-Log -Message "Failed to join domain '$DomainName': $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Errors = @($_.Exception.Message); Message = $_.Exception.Message }
    }
}

function Install-SoftwarePackage {
    <#
    .SYNOPSIS
        Installs a software package silently.
    .DESCRIPTION
        Supports custom silent switches per package. For .msi files, /qn /norestart
        is always used. For .exe files, custom switches override the default /S /norestart.
        The SilentSwitches parameter accepts a full argument string (e.g. '/install /quiet /norestart').
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName,
        [Parameter(Mandatory)]
        [string]$InstallerPath,
        [string]$SilentSwitches
    )

    if (-not (Test-Path -Path $InstallerPath)) {
        Write-Log -Message "Installer not found for ${DisplayName}: $InstallerPath" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = "Installer not found: $InstallerPath" }
    }

    $extension = [System.IO.Path]::GetExtension($InstallerPath).ToLower()
    $processName = [System.IO.Path]::GetFileNameWithoutExtension($InstallerPath)

    Write-Log -Message "Installing ${DisplayName} from: $InstallerPath" -Level 'INFO'

    try {
        switch ($extension) {
            '.msi' {
                $installArgs = "/i `"$InstallerPath`" /qn /norestart /l*v `"$($script:LogPath).${DisplayName}.install.log`""
                Write-Log -Message "Running: msiexec.exe $installArgs" -Level 'DEBUG'
                $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
            }
            '.exe' {
                # Use custom silent switches if provided, otherwise default
                if (-not [string]::IsNullOrWhiteSpace($SilentSwitches)) {
                    $installArgs = $SilentSwitches
                } else {
                    $installArgs = "/S /norestart"
                }
                Write-Log -Message "Running: $InstallerPath $installArgs" -Level 'DEBUG'
                $proc = Start-Process -FilePath $InstallerPath -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
            }
            '.ps1' {
                Write-Log -Message "Running PowerShell script: $InstallerPath" -Level 'DEBUG'
                $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$InstallerPath`"" -Wait -PassThru -NoNewWindow -ErrorAction Stop
            }
            '.bat' {
                Write-Log -Message "Running batch file: $InstallerPath" -Level 'DEBUG'
                $proc = Start-Process -FilePath $InstallerPath -Wait -PassThru -NoNewWindow -ErrorAction Stop
            }
            default {
                Write-Log -Message "Running: $InstallerPath" -Level 'DEBUG'
                $proc = Start-Process -FilePath $InstallerPath -Wait -PassThru -NoNewWindow -ErrorAction Stop
            }
        }

        if ($proc.ExitCode -eq 0) {
            Write-Log -Message "${DisplayName} installed successfully (exit code: $($proc.ExitCode))." -Level 'SUCCESS'
            return [PSCustomObject]@{ Success = $true; Message = "${DisplayName} installed successfully." }
        } else {
            $msg = "${DisplayName} installation returned exit code: $($proc.ExitCode)"
            Write-Log -Message $msg -Level 'WARNING'
            # Some installers return non-zero on success (e.g., 1638 = already installed)
            if ($proc.ExitCode -eq 1638) {
                Write-Log -Message "${DisplayName} appears to already be installed (exit 1638)." -Level 'INFO'
                return [PSCustomObject]@{ Success = $true; Message = "${DisplayName} was already installed." }
            }
            return [PSCustomObject]@{ Success = $false; Message = $msg }
        }
    } catch {
        Write-Log -Message "Error installing ${DisplayName}: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Find-SoftwareInstallers {
    <#
    .SYNOPSIS
        Scans the software folder for installer files matching known package names.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    if (-not (Test-Path -Path $FolderPath)) {
        return @()
    }

    $found = @()
    $installerExts = @('*.msi', '*.exe', '*.ps1', '*.bat')

    foreach ($pkg in $script:SoftwarePackages) {
        $matchedFiles = @()
        foreach ($ext in $installerExts) {
            $files = Get-ChildItem -Path $FolderPath -Filter $ext -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match $pkg.Pattern }
            if ($files) {
                $matchedFiles += $files
            }
        }
        if ($matchedFiles.Count -gt 0) {
            # Prefer MSI, then EXE
            $best = $matchedFiles | Where-Object { $_.Extension -eq '.msi' } | Select-Object -First 1
            if (-not $best) {
                $best = $matchedFiles | Where-Object { $_.Extension -eq '.exe' } | Select-Object -First 1
            }
            if (-not $best) {
                $best = $matchedFiles | Select-Object -First 1
            }
            $found += [PSCustomObject]@{
                Package         = $pkg.Name
                Pattern         = $pkg.Pattern
                SilentSwitches  = $pkg.SilentSwitches
                Installer       = $best.FullName
                FileName        = $best.Name
            }
        }
    }

    return $found
}

function Get-DomainOUs {
    <#
    .SYNOPSIS
        Enumerates Organizational Units from a target domain using ADSI.
    .DESCRIPTION
        Uses System.DirectoryServices.DirectorySearcher to query the domain
        for all OUs. Returns distinguished names sorted alphabetically.
        Does not require the ActiveDirectory PowerShell module.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$DomainName,
        [string]$Server
    )

    $results = @()

    try {
        # Build the LDAP path - target the domain root or specific DC
        $domainParts = $DomainName.Split('.')
        $domainDN = ($domainParts | ForEach-Object { "DC=$_" }) -join ','

        if ($Server) {
            $ldapPath = "LDAP://$Server/$domainDN"
        } else {
            $ldapPath = "LDAP://$domainDN"
        }

        Write-Log -Message "Querying OUs from: $ldapPath" -Level 'DEBUG'

        $searchRoot = New-Object System.DirectoryServices.DirectoryEntry($ldapPath)
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
        $searcher.Filter = '(objectClass=organizationalUnit)'
        $searcher.PropertiesToLoad.Add('distinguishedName') | Out-Null
        $searcher.PropertiesToLoad.Add('name') | Out-Null
        $searcher.PageSize = 1000
        $searcher.SizeLimit = 5000

        $searchResults = $searcher.FindAll()

        foreach ($sr in $searchResults) {
            $dn = $sr.Properties['distinguishedName'][0]
            $name = $sr.Properties['name'][0]
            if ($dn) {
                $results += [PSCustomObject]@{
                    Name              = $name
                    DistinguishedName = $dn
                    Depth             = ($dn -split ',').Count
                }
            }
        }

        $searcher.Dispose()
        $searchRoot.Dispose()

        # Sort by depth (root OUs first), then alphabetically
        $results = $results | Sort-Object Depth, Name

        Write-Log -Message "Found $($results.Count) OUs in domain '$DomainName'." -Level 'INFO'
        return $results
    } catch {
        Write-Log -Message "Failed to enumerate OUs from '$DomainName': $($_.Exception.Message)" -Level 'ERROR'
        return @()
    }
}

# ============================================================================
# REGION: System Prep Functions (Step 3 - PrepSystemForDeployment)
# ============================================================================

function Invoke-BGInfoRemoval {
    <#
    .SYNOPSIS
        Removes BGInfo from the system (IGT/IGTSAP only).
    #>
    Write-Log -Message 'Removing BGInfo from machine...' -Level 'INFO'
    try {
        # Delete BGInfo shortcuts from Startup folder
        if (Test-Path -Path $script:BGInfoStartupPath) {
            $bgFiles = Get-ChildItem -Path $script:BGInfoStartupPath -Filter '*BGInfo*' -ErrorAction SilentlyContinue
            foreach ($f in $bgFiles) {
                Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                Write-Log -Message "Deleted: $($f.FullName)" -Level 'DEBUG'
            }
        }
        # Remove BGInfo folder
        if (Test-Path -Path $script:BGInfoFolder) {
            Remove-Item -Path $script:BGInfoFolder -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log -Message "Removed folder: $script:BGInfoFolder" -Level 'DEBUG'
        }
        Write-Log -Message 'BGInfo removed. Ensure computer is in proper Security Group (MYIGT) or OU (IGTSAP).' -Level 'SUCCESS'
        return [PSCustomObject]@{ Success = $true; Message = 'BGInfo removed from system.' }
    } catch {
        Write-Log -Message "Failed to remove BGInfo: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-CriblConfig {
    <#
    .SYNOPSIS
        Runs the appropriate CRIBL configuration script based on environment.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Environment  # 'Production L-Everi', 'Non-Production L-Everi', 'IGT'
    )

    $scriptInfo = $script:CriblScripts | Where-Object { $_.Env -eq $Environment } | Select-Object -First 1
    if (-not $scriptInfo) {
        Write-Log -Message "Unknown CRIBL environment: $Environment" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = "Unknown CRIBL environment: $Environment" }
    }

    $scriptPath = Join-Path $script:CriblFolder $scriptInfo.Script
    if (-not (Test-Path -Path $scriptPath)) {
        Write-Log -Message "CRIBL script not found: $scriptPath" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = "CRIBL script not found: $scriptPath" }
    }

    Write-Log -Message "Running CRIBL configuration for ${Environment}: $scriptPath" -Level 'INFO'
    try {
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Wait -PassThru -NoNewWindow -ErrorAction Stop
        if ($proc.ExitCode -eq 0) {
            Write-Log -Message "CRIBL configuration completed for $Environment." -Level 'SUCCESS'
            return [PSCustomObject]@{ Success = $true; Message = "CRIBL configured for $Environment." }
        } else {
            Write-Log -Message "CRIBL script returned exit code: $($proc.ExitCode)" -Level 'WARNING'
            return [PSCustomObject]@{ Success = $true; Message = "CRIBL script completed (exit $($proc.ExitCode))." }
        }
    } catch {
        Write-Log -Message "Failed to run CRIBL script: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-SCCMInstall {
    <#
    .SYNOPSIS
        Installs SCCM client with the correct site code and management point.
    .DESCRIPTION
        Tier 1/2 L-Everi: SMSSITECODE=TIT, MP=LASAPPSCCM01P.titanium.com
        Tier 3/4 L-Everi: SMSSITECODE=GHQ, MP=LASAPPSCM01P.hq.globalcashaccess.us
        IGT systems: skipped (no SCCM in this workflow)
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Tier  # 'Tier 1/2 L-Everi', 'Tier 3/4 L-Everi'
    )

    $sccmCfg = $script:SCCMConfigs | Where-Object { $_.Tier -eq $Tier } | Select-Object -First 1
    if (-not $sccmCfg) {
        Write-Log -Message "Unknown SCCM tier: $Tier" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = "Unknown SCCM tier: $Tier" }
    }

    Write-Log -Message "Installing SCCM client: Site=$($sccmCfg.SiteCode), MP=$($sccmCfg.MP)" -Level 'INFO'

    try {
        # Step 1: Local ccmsetup with ForceInstall
        $ccmSetup = Join-Path $script:ITFolder 'ccmsetup.exe'
        if (Test-Path -Path $ccmSetup) {
            Write-Log -Message "Running local ccmsetup: $ccmSetup SMSSITECODE=$($sccmCfg.SiteCode) /ForceInstall" -Level 'DEBUG'
            $proc = Start-Process -FilePath $ccmSetup -ArgumentList "SMSSITECODE=$($sccmCfg.SiteCode) /ForceInstall" -Wait -PassThru -NoNewWindow -ErrorAction Stop
            Write-Log -Message "Local ccmsetup exit code: $($proc.ExitCode)" -Level 'DEBUG'
        }

        # Step 2: Map network share and run ccmsetup from there
        $driveLetter = if ($Tier -eq 'Tier 1/2 L-Everi') { 'T:' } else { 'S:' }
        $sharePath = $sccmCfg.Share

        Write-Log -Message "Mapping $driveLetter to $sharePath..." -Level 'DEBUG'
        try {
            # Remove existing mapping if present
            net use $driveLetter /delete 2>$null
            # Map the share
            net use $driveLetter $sharePath /p:no 2>$null
        } catch {
            Write-Log -Message "NET USE mapping may require credentials. Error: $($_.Exception.Message)" -Level 'WARNING'
        }

        # Run ccmsetup from the mapped drive
        $remoteCcmSetup = "$driveLetter\ccmsetup.exe"
        if (Test-Path -Path $remoteCcmSetup) {
            Write-Log -Message "Running remote ccmsetup: $remoteCcmSetup /mp:$($sccmCfg.MP) SMSSITECODE=$($sccmCfg.SiteCode) /logon" -Level 'DEBUG'
            $proc2 = Start-Process -FilePath $remoteCcmSetup -ArgumentList "/mp:$($sccmCfg.MP) SMSSITECODE=$($sccmCfg.SiteCode) /logon" -Wait -PassThru -NoNewWindow -ErrorAction Stop
            Write-Log -Message "Remote ccmsetup exit code: $($proc2.ExitCode)" -Level 'DEBUG'
        } else {
            Write-Log -Message "Remote ccmsetup not found at: $remoteCcmSetup. NET USE may require manual credentials." -Level 'WARNING'
        }

        # Clean up the mapped drive
        try { net use $driveLetter /delete 2>$null } catch {}

        Write-Log -Message "SCCM client installation completed for $Tier." -Level 'SUCCESS'
        return [PSCustomObject]@{ Success = $true; Message = "SCCM client installed (Site: $($sccmCfg.SiteCode), MP: $($sccmCfg.MP))." }
    } catch {
        Write-Log -Message "Failed to install SCCM client: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-AdminPasswordNoExpire {
    <#
    .SYNOPSIS
        Sets Administrator and sysengadm account passwords to non-expiring.
    #>
    Write-Log -Message 'Setting admin account passwords to non-expire...' -Level 'INFO'
    $results = @()
    foreach ($account in $script:AdminAccounts) {
        try {
            $user = Get-CimInstance -ClassName Win32_UserAccount -Filter "Name='$account'" -ErrorAction Stop
            if ($user) {
                $user | Set-CimInstance -Property @{ PasswordExpires = $false } -ErrorAction Stop
                Write-Log -Message "Set '$account' password to non-expire." -Level 'SUCCESS'
                $results += [PSCustomObject]@{ Account = $account; Success = $true }
            } else {
                Write-Log -Message "Account '$account' not found - skipping." -Level 'WARNING'
                $results += [PSCustomObject]@{ Account = $account; Success = $false }
            }
        } catch {
            Write-Log -Message "Failed to set '$account' password non-expire: $($_.Exception.Message)" -Level 'WARNING'
            $results += [PSCustomObject]@{ Account = $account; Success = $false }
        }
    }
    $successCount = ($results | Where-Object { $_.Success }).Count
    return [PSCustomObject]@{ Success = ($successCount -gt 0); Message = "Admin password non-expire: $successCount of $($results.Count) accounts set." }
}

function Invoke-Rapid7ServiceConfig {
    <#
    .SYNOPSIS
        Sets the Rapid7 Insight Agent service (ir_agent) to Automatic startup.
    #>
    Write-Log -Message 'Setting Rapid7 Insight Agent service to Automatic...' -Level 'INFO'
    try {
        Set-Service -Name 'ir_agent' -StartupType Automatic -ErrorAction Stop
        Write-Log -Message 'Rapid7 ir_agent service set to Automatic.' -Level 'SUCCESS'
        return [PSCustomObject]@{ Success = $true; Message = 'Rapid7 ir_agent set to Automatic.' }
    } catch {
        Write-Log -Message "Failed to set ir_agent service: $($_.Exception.Message)" -Level 'WARNING'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-Rapid7ScanAssist {
    <#
    .SYNOPSIS
        Installs Rapid7 ScanAssist with embedded certificate.
    #>
    $msiPath = Join-Path $script:Rapid7Folder 'ScanAssistantInstaller.msi'
    if (-not (Test-Path -Path $msiPath)) {
        Write-Log -Message "Rapid7 ScanAssist MSI not found: $msiPath" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = "ScanAssist MSI not found: $msiPath" }
    }

    Write-Log -Message 'Installing Rapid7 ScanAssist...' -Level 'INFO'
    try {
        $installArgs = "/i `"$msiPath`" /qn /norestart /l*v `"$($script:LogPath).Rapid7ScanAssist.install.log`" CLIENT_CERTIFICATE=`"$($script:Rapid7Cert -replace '\s+', ' ')`""
        Write-Log -Message 'Running: msiexec.exe (Rapid7 ScanAssist)' -Level 'DEBUG'
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop

        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 1638) {
            Write-Log -Message "Rapid7 ScanAssist installed (exit code: $($proc.ExitCode))." -Level 'SUCCESS'
            return [PSCustomObject]@{ Success = $true; Message = 'Rapid7 ScanAssist installed.' }
        } else {
            Write-Log -Message "Rapid7 ScanAssist exit code: $($proc.ExitCode)" -Level 'WARNING'
            return [PSCustomObject]@{ Success = $false; Message = "ScanAssist exit code: $($proc.ExitCode)" }
        }
    } catch {
        Write-Log -Message "Failed to install Rapid7 ScanAssist: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-CrowdStrikeInstall {
    <#
    .SYNOPSIS
        Installs CrowdStrike Falcon Sensor with the configured CID.
    #>
    $sensorPath = Join-Path $script:CrowdStrikeFolder 'FalconSensor_Windows.exe'
    if (-not (Test-Path -Path $sensorPath)) {
        # Try to find it with a wildcard
        $found = Get-ChildItem -Path $script:CrowdStrikeFolder -Filter 'FalconSensor*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) {
            $sensorPath = $found.FullName
        } else {
            Write-Log -Message "CrowdStrike Falcon Sensor not found in $script:CrowdStrikeFolder" -Level 'ERROR'
            return [PSCustomObject]@{ Success = $false; Message = 'CrowdStrike sensor not found.' }
        }
    }

    Write-Log -Message "Installing CrowdStrike Falcon Sensor (CID: $($script:CrowdStrikeCID))..." -Level 'INFO'
    try {
        $installArgs = "/install /quiet /norestart CID=$($script:CrowdStrikeCID)"
        Write-Log -Message "Running: $sensorPath $installArgs" -Level 'DEBUG'
        $proc = Start-Process -FilePath $sensorPath -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop

        if ($proc.ExitCode -eq 0) {
            Write-Log -Message 'CrowdStrike Falcon Sensor installed successfully.' -Level 'SUCCESS'
            return [PSCustomObject]@{ Success = $true; Message = 'CrowdStrike installed.' }
        } else {
            Write-Log -Message "CrowdStrike install exit code: $($proc.ExitCode)" -Level 'WARNING'
            return [PSCustomObject]@{ Success = $false; Message = "CrowdStrike exit code: $($proc.ExitCode)" }
        }
    } catch {
        Write-Log -Message "Failed to install CrowdStrike: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-GpUpdateAndCleanup {
    <#
    .SYNOPSIS
        Runs gpupdate /force and cleans up stored credentials.
    #>
    Write-Log -Message 'Running gpupdate /force...' -Level 'INFO'
    try {
        $proc = Start-Process -FilePath 'gpupdate.exe' -ArgumentList '/force' -Wait -PassThru -NoNewWindow -ErrorAction Stop
        Write-Log -Message "gpupdate completed (exit: $($proc.ExitCode))." -Level 'DEBUG'
    } catch {
        Write-Log -Message "gpupdate failed: $($_.Exception.Message)" -Level 'WARNING'
    }

    Write-Log -Message 'Cleaning up stored credentials...' -Level 'INFO'
    try {
        $proc = Start-Process -FilePath 'cmdkey.exe' -ArgumentList '/delete:WindowsLive:target=virtualapp/didlogical' -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        $proc = Start-Process -FilePath 'cmdkey.exe' -ArgumentList '/delete:MicrosoftAccount:target=SSO_POP_Device' -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        Write-Log -Message 'Credential cleanup completed.' -Level 'SUCCESS'
    } catch {
        Write-Log -Message 'Credential cleanup had errors (non-critical).' -Level 'WARNING'
    }

    return [PSCustomObject]@{ Success = $true; Message = 'gpupdate and credential cleanup done.' }
}

function Invoke-IISCrypto {
    <#
    .SYNOPSIS
        Runs IISCryptoCli to disable weak ciphers and apply security template.
    #>
    $iisCryptoPath = Join-Path $script:ITFolder 'IISCryptoCli.exe'
    $templatePath = $script:IISCryptoTemplate

    if (-not (Test-Path -Path $iisCryptoPath)) {
        Write-Log -Message "IISCryptoCli.exe not found: $iisCryptoPath" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = 'IISCryptoCli.exe not found.' }
    }
    if (-not (Test-Path -Path $templatePath)) {
        Write-Log -Message "IISCrypto template not found: $templatePath" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = 'IISCrypto template not found.' }
    }

    Write-Log -Message "Running IISCryptoCli with template: $templatePath" -Level 'INFO'
    try {
        $installArgs = "/template `"$templatePath`""
        Write-Log -Message "Running: $iisCryptoPath $installArgs" -Level 'DEBUG'
        $proc = Start-Process -FilePath $iisCryptoPath -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop

        if ($proc.ExitCode -eq 0) {
            Write-Log -Message 'IISCryptoCli applied template successfully.' -Level 'SUCCESS'
            return [PSCustomObject]@{ Success = $true; Message = 'IISCrypto template applied. Reboot required.' }
        } else {
            Write-Log -Message "IISCryptoCli exit code: $($proc.ExitCode)" -Level 'WARNING'
            return [PSCustomObject]@{ Success = $false; Message = "IISCryptoCli exit code: $($proc.ExitCode)" }
        }
    } catch {
        Write-Log -Message "Failed to run IISCryptoCli: $($_.Exception.Message)" -Level 'ERROR'
        return [PSCustomObject]@{ Success = $false; Message = $_.Exception.Message }
    }
}

function Invoke-PrepSystem {
    <#
    .SYNOPSIS
        Orchestrates the full PrepSystemForDeployment workflow.
    .DESCRIPTION
        Runs the selected prep steps in the correct order:
        1. BGInfo removal (IGT/IGTSAP only)
        2. CRIBL configuration (environment-specific)
        3. SCCM client install (Everi tiers only)
        4. Admin password non-expire
        5. Rapid7 service config + ScanAssist install
        6. CrowdStrike install
        7. gpupdate /force + credential cleanup
        8. IISCryptoCli template + reboot
    #>
    param(
        [string]$SystemType,      # 'IGT/IGTSAP', 'Legacy-Everi Production', 'Legacy-Everi Non-Production', 'Legacy-IGT'
        [string]$SCCMTier,        # 'Tier 1/2 L-Everi', 'Tier 3/4 L-Everi', ''
        [bool]$RunBGInfo,
        [bool]$RunCribl,
        [bool]$RunSCCM,
        [bool]$RunAdminAccounts,
        [bool]$RunRapid7,
        [bool]$RunCrowdStrike,
        [bool]$RunGpUpdate,
        [bool]$RunIISCrypto
    )

    $stepResults = @()
    $totalSteps = 0
    $completedSteps = 0

    # Count steps
    if ($RunBGInfo) { $totalSteps++ }
    if ($RunCribl) { $totalSteps++ }
    if ($RunSCCM) { $totalSteps++ }
    if ($RunAdminAccounts) { $totalSteps++ }
    if ($RunRapid7) { $totalSteps++ }
    if ($RunCrowdStrike) { $totalSteps++ }
    if ($RunGpUpdate) { $totalSteps++ }
    if ($RunIISCrypto) { $totalSteps++ }

    Write-Log -Message "=== Starting System Prep ($totalSteps steps) ===" -Level 'INFO'
    Write-Log -Message "System Type: $SystemType" -Level 'INFO'
    if ($SCCMTier) { Write-Log -Message "SCCM Tier: $SCCMTier" -Level 'INFO' }

    # Step 1: BGInfo removal
    if ($RunBGInfo) {
        $stepResults += [PSCustomObject]@{ Step = 'BGInfo Removal'; Result = (Invoke-BGInfoRemoval) }
        $completedSteps++
    }

    # Step 2: CRIBL configuration
    if ($RunCribl) {
        $criblEnv = switch ($SystemType) {
            'Legacy-Everi Production'     { 'Production L-Everi' }
            'Legacy-Everi Non-Production' { 'Non-Production L-Everi' }
            'Legacy-IGT'                  { 'IGT' }
            'IGT/IGTSAP'                  { 'IGT' }
            default { 'IGT' }
        }
        $stepResults += [PSCustomObject]@{ Step = 'CRIBL Config'; Result = (Invoke-CriblConfig -Environment $criblEnv) }
        $completedSteps++
    }

    # Step 3: SCCM client install
    if ($RunSCCM -and $SCCMTier) {
        $stepResults += [PSCustomObject]@{ Step = 'SCCM Client'; Result = (Invoke-SCCMInstall -Tier $SCCMTier) }
        $completedSteps++
    }

    # Step 4: Admin password non-expire
    if ($RunAdminAccounts) {
        $stepResults += [PSCustomObject]@{ Step = 'Admin Passwords'; Result = (Invoke-AdminPasswordNoExpire) }
        $completedSteps++
    }

    # Step 5: Rapid7 service config + ScanAssist
    if ($RunRapid7) {
        $svcResult = Invoke-Rapid7ServiceConfig
        $scanResult = Invoke-Rapid7ScanAssist
        $stepResults += [PSCustomObject]@{ Step = 'Rapid7 Service'; Result = $svcResult }
        $stepResults += [PSCustomObject]@{ Step = 'Rapid7 ScanAssist'; Result = $scanResult }
        $completedSteps++
    }

    # Step 6: CrowdStrike
    if ($RunCrowdStrike) {
        $stepResults += [PSCustomObject]@{ Step = 'CrowdStrike'; Result = (Invoke-CrowdStrikeInstall) }
        $completedSteps++
    }

    # Step 7: gpupdate + credential cleanup
    if ($RunGpUpdate) {
        $stepResults += [PSCustomObject]@{ Step = 'GPUpdate + Cleanup'; Result = (Invoke-GpUpdateAndCleanup) }
        $completedSteps++
    }

    # Step 8: IISCryptoCli
    if ($RunIISCrypto) {
        $stepResults += [PSCustomObject]@{ Step = 'IISCrypto'; Result = (Invoke-IISCrypto) }
        $completedSteps++
    }

    $successCount = ($stepResults | Where-Object { $_.Result.Success }).Count
    $failCount = $stepResults.Count - $successCount

    Write-Log -Message "=== System Prep Complete: $successCount succeeded, $failCount failed ===" -Level 'INFO'

    return [PSCustomObject]@{
        Success = $failCount -eq 0
        StepResults = $stepResults
        SuccessCount = $successCount
        FailCount = $failCount
        Message = "Prep complete: $successCount succeeded, $failCount failed."
    }
}

# ============================================================================
# REGION: GUI Construction
# ============================================================================

# --- Main Form ---
$mainForm = New-Object System.Windows.Forms.Form -Property @{
    Text          = 'Domain Join & Provisioning Tool'
    Size          = New-Object System.Drawing.Size(960, 900)
    MinimumSize   = New-Object System.Drawing.Size(960, 600)
    StartPosition = 'CenterScreen'
    FormBorderStyle = 'Sizable'
    MaximizeBox   = $true
    MinimizeBox   = $true
    BackColor     = [System.Drawing.Color]::FromArgb(45, 45, 48)
    Font          = New-Object System.Drawing.Font('Segoe UI', 9)
    AutoScroll    = $true
    AutoScrollMinSize = New-Object System.Drawing.Size(940, 1230)
}

# --- Colors ---
$colorBg       = [System.Drawing.Color]::FromArgb(45, 45, 48)
$colorPanel    = [System.Drawing.Color]::FromArgb(62, 62, 66)
$colorPanelAlt = [System.Drawing.Color]::FromArgb(51, 51, 55)
$colorAccent   = [System.Drawing.Color]::FromArgb(0, 122, 204)
$colorAccent2  = [System.Drawing.Color]::FromArgb(38, 160, 218)
$colorText     = [System.Drawing.Color]::White
$colorTextDim  = [System.Drawing.Color]::FromArgb(200, 200, 200)
$colorSuccess  = [System.Drawing.Color]::FromArgb(76, 175, 80)
$colorWarning  = [System.Drawing.Color]::FromArgb(255, 152, 0)
$colorError    = [System.Drawing.Color]::FromArgb(244, 67, 54)
$colorInput    = [System.Drawing.Color]::FromArgb(69, 69, 73)

# --- Helper function to create styled controls ---
function New-StyledLabel {
    param($Text, $X, $Y, $Width = 200, $Height = 22)
    $lbl = New-Object System.Windows.Forms.Label -Property @{
        Text     = $Text
        Location = New-Object System.Drawing.Point($X, $Y)
        Size     = New-Object System.Drawing.Size($Width, $Height)
        ForeColor = $colorTextDim
        BackColor = [System.Drawing.Color]::Transparent
        Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    }
    return $lbl
}

function New-StyledTextBox {
    param($X, $Y, $Width = 250, $Height = 26, $IsPassword = $false)
    $tb = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object System.Drawing.Point($X, $Y)
        Size     = New-Object System.Drawing.Size($Width, $Height)
        BackColor = $colorInput
        ForeColor = $colorText
        BorderStyle = 'FixedSingle'
        Font      = New-Object System.Drawing.Font('Segoe UI', 9)
    }
    if ($IsPassword) {
        $tb.UseSystemPasswordChar = $true
    }
    return $tb
}

function New-StyledComboBox {
    param($X, $Y, $Width = 250, $Height = 26)
    $cb = New-Object System.Windows.Forms.ComboBox -Property @{
        Location = New-Object System.Drawing.Point($X, $Y)
        Size     = New-Object System.Drawing.Size($Width, $Height)
        BackColor = $colorInput
        ForeColor = $colorText
        FlatStyle = 'Flat'
        Font      = New-Object System.Drawing.Font('Segoe UI', 9)
        DropDownStyle = 'DropDownList'
    }
    return $cb
}

function New-StyledButton {
    param($Text, $X, $Y, $Width = 120, $Height = 32)
    $btn = New-Object System.Windows.Forms.Button -Property @{
        Text     = $Text
        Location = New-Object System.Drawing.Point($X, $Y)
        Size     = New-Object System.Drawing.Size($Width, $Height)
        BackColor = $colorAccent
        ForeColor = [System.Drawing.Color]::White
        FlatStyle = 'Flat'
        Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
        Cursor    = [System.Windows.Forms.Cursors]::Hand
    }
    $btn.FlatAppearance.BorderSize = 0
    return $btn
}

function New-StyledGroupBox {
    param($Text, $X, $Y, $Width, $Height)
    $gb = New-Object System.Windows.Forms.GroupBox -Property @{
        Text     = "  $Text"
        Location = New-Object System.Drawing.Point($X, $Y)
        Size     = New-Object System.Drawing.Size($Width, $Height)
        BackColor = $colorPanel
        ForeColor = $colorAccent2
        Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    }
    return $gb
}

# --- Title Label ---
$lblTitle = New-Object System.Windows.Forms.Label -Property @{
    Text     = 'Domain Join & Provisioning Tool'
    Location = New-Object System.Drawing.Point(20, 12)
    Size     = New-Object System.Drawing.Size(600, 32)
    ForeColor = [System.Drawing.Color]::White
    BackColor = [System.Drawing.Color]::Transparent
    Font      = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
}
$mainForm.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label -Property @{
    Text     = "Computer: $env:COMPUTERNAME  |  Logged in as: $env:USERNAME"
    Location = New-Object System.Drawing.Point(20, 44)
    Size     = New-Object System.Drawing.Size(600, 20)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font      = New-Object System.Drawing.Font('Segoe UI', 8)
}
$mainForm.Controls.Add($lblSubtitle)

# ============================================================================
# SECTION 1: Log File Location
# ============================================================================
$gbLog = New-StyledGroupBox -Text 'Log File Location' -X 15 -Y 72 -Width 920 -Height 60
$mainForm.Controls.Add($gbLog)

$lblLogPath = New-StyledLabel -Text 'DomainJoinGUI.log save location:' -X 15 -Y 25 -Width 180
$gbLog.Controls.Add($lblLogPath)

$txtLogPath = New-StyledTextBox -X 200 -Y 22 -Width 550
$gbLog.Controls.Add($txtLogPath)

$btnBrowseLog = New-StyledButton -Text 'Browse...' -X 760 -Y 20 -Width 80 -Height 28
$gbLog.Controls.Add($btnBrowseLog)

$btnBrowseLog.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = 'Select folder to save Domain Join GUI.log'
        ShowNewFolderButton = $true
        RootFolder = [System.Environment+SpecialFolder]::MyComputer
    }
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogPath.Text = Join-Path $folderDialog.SelectedPath $script:LogFileName
    }
})

$txtLogPath.Add_TextChanged({
    if (-not [string]::IsNullOrWhiteSpace($txtLogPath.Text)) {
        $script:LogPath = $txtLogPath.Text
        Write-Log -Message "Log file location set to: $script:LogPath" -Level 'INFO'
    }
})

# Set default log path
$defaultLogDir = 'C:\IT\Logs'
if (-not (Test-Path $defaultLogDir)) { $defaultLogDir = $env:TEMP }
$txtLogPath.Text = Join-Path $defaultLogDir $script:LogFileName
$script:LogPath = $txtLogPath.Text

# ============================================================================
# SECTION 2: Domain Join & Computer Rename
# ============================================================================
$gbDomain = New-StyledGroupBox -Text 'Domain Join & Computer Rename' -X 15 -Y 140 -Width 450 -Height 380
$mainForm.Controls.Add($gbDomain)

$lblDomain = New-StyledLabel -Text 'Select Domain:' -X 15 -Y 30 -Width 130
$gbDomain.Controls.Add($lblDomain)

$cmbDomain = New-StyledComboBox -X 150 -Y 27 -Width 280
$gbDomain.Controls.Add($cmbDomain)
foreach ($d in $script:DomainList) { [void]$cmbDomain.Items.Add($d) }
if ($cmbDomain.Items.Count -gt 0) { $cmbDomain.SelectedIndex = 0 }

$lblNewName = New-StyledLabel -Text 'New Computer Name:' -X 15 -Y 62 -Width 130
$gbDomain.Controls.Add($lblNewName)

$txtNewName = New-StyledTextBox -X 150 -Y 59 -Width 280
$gbDomain.Controls.Add($txtNewName)

$lblDomainUser = New-StyledLabel -Text 'Domain Username:' -X 15 -Y 94 -Width 130
$gbDomain.Controls.Add($lblDomainUser)

$txtDomainUser = New-StyledTextBox -X 150 -Y 91 -Width 280
$gbDomain.Controls.Add($txtDomainUser)

$lblDomainPass = New-StyledLabel -Text 'Domain Password:' -X 15 -Y 126 -Width 130
$gbDomain.Controls.Add($lblDomainPass)

$txtDomainPass = New-StyledTextBox -X 150 -Y 123 -Width 280 -IsPassword $true
$gbDomain.Controls.Add($txtDomainPass)

$lblOUPath = New-StyledLabel -Text 'OU Path (optional):' -X 15 -Y 158 -Width 130
$gbDomain.Controls.Add($lblOUPath)

$txtOUPath = New-StyledTextBox -X 150 -Y 155 -Width 200
$gbDomain.Controls.Add($txtOUPath)
$txtOUPath.Text = ''

$btnBrowseOU = New-StyledButton -Text 'Browse OUs...' -X 355 -Y 153 -Width 85 -Height 28
$btnBrowseOU.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$gbDomain.Controls.Add($btnBrowseOU)

$lblDomainController = New-StyledLabel -Text 'Domain Controller:' -X 15 -Y 190 -Width 130
$gbDomain.Controls.Add($lblDomainController)

$cmbDomainController = New-StyledComboBox -X 150 -Y 187 -Width 280
$gbDomain.Controls.Add($cmbDomainController)
# Populate DC list
foreach ($dc in $script:DomainControllerList) {
    [void]$cmbDomainController.Items.Add("$($dc.Label) - $($dc.Host)")
}
# Add an "Auto (DNS)" option at the top
[void]$cmbDomainController.Items.Insert(0, 'Auto (DNS resolve)')
if ($cmbDomainController.Items.Count -gt 0) { $cmbDomainController.SelectedIndex = 0 }

$lblDomainHint = New-StyledLabel -Text 'Format: domain\username (e.g. corp\user.name)' -X 150 -Y 215 -Width 280 -Height 16
$lblDomainHint.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Italic)
$lblDomainHint.ForeColor = $colorTextDim
$gbDomain.Controls.Add($lblDomainHint)

# NetJoinLegacyAccountReuse checkbox
$chkLegacyReuse = New-Object System.Windows.Forms.CheckBox -Property @{
    Text     = 'Enable NetJoinLegacyAccountReuse'
    Location = New-Object System.Drawing.Point(15, 240)
    Size     = New-Object System.Drawing.Size(280, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font      = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked   = $true
}
$gbDomain.Controls.Add($chkLegacyReuse)

$btnRenameOnly = New-StyledButton -Text 'Rename Only' -X 15 -Y 275 -Width 125 -Height 30
$gbDomain.Controls.Add($btnRenameOnly)

$btnJoinDomain = New-StyledButton -Text 'Join Domain' -X 150 -Y 275 -Width 125 -Height 30
$gbDomain.Controls.Add($btnJoinDomain)

$btnChangeDomain = New-StyledButton -Text 'Change Domain' -X 285 -Y 275 -Width 125 -Height 30
$gbDomain.Controls.Add($btnChangeDomain)

$lblDomainStatus = New-Object System.Windows.Forms.Label -Property @{
    Text     = 'Status: Ready'
    Location = New-Object System.Drawing.Point(15, 315)
    Size     = New-Object System.Drawing.Size(410, 50)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font      = New-Object System.Drawing.Font('Segoe UI', 8)
}
$gbDomain.Controls.Add($lblDomainStatus)

# --- OU Browse Button ---
$btnBrowseOU.Add_Click({
    $domain = $cmbDomain.SelectedItem
    if (-not $domain) {
        [System.Windows.Forms.MessageBox]::Show('Please select a domain first.', 'OU Browser', 'OK', 'Warning') | Out-Null
        return
    }

    # Determine the target DC from the dropdown
    $dcHost = ''
    $dcSelection = $cmbDomainController.SelectedItem
    if ($dcSelection -and $dcSelection -ne 'Auto (DNS resolve)') {
        # Extract hostname (after " - ")
        $parts = $dcSelection -split ' - '
        if ($parts.Count -ge 2) {
            $dcHost = $parts[-1].Trim()
        }
    }

    $lblDomainStatus.Text = 'Status: Querying OUs from domain... please wait.'
    $lblDomainStatus.ForeColor = $colorAccent2
    [System.Windows.Forms.Application]::DoEvents()

    $ous = Get-DomainOUs -DomainName $domain -Server $dcHost

    if ($ous.Count -eq 0) {
        $lblDomainStatus.Text = 'Status: No OUs found or unable to connect to domain.'
        $lblDomainStatus.ForeColor = $colorWarning
        [System.Windows.Forms.MessageBox]::Show(
            "Could not enumerate OUs from '$domain'.`n`nYou can still type the OU path manually.",
            'OU Browser',
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    # Create a popup dialog with a ListBox for OU selection
    $ouForm = New-Object System.Windows.Forms.Form -Property @{
        Text          = "Select OU from $domain"
        Size          = New-Object System.Drawing.Size(700, 500)
        StartPosition = 'CenterParent'
        FormBorderStyle = 'Sizable'
        BackColor     = $colorBg
        Font          = New-Object System.Drawing.Font('Segoe UI', 9)
    }

    $ouLabel = New-Object System.Windows.Forms.Label -Property @{
        Text     = 'Select an Organizational Unit:'
        Location = New-Object System.Drawing.Point(10, 10)
        Size     = New-Object System.Drawing.Size(660, 22)
        ForeColor = $colorAccent2
        BackColor = [System.Drawing.Color]::Transparent
        Font      = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    }
    $ouForm.Controls.Add($ouLabel)

    $ouListBox = New-Object System.Windows.Forms.ListBox -Property @{
        Location = New-Object System.Drawing.Point(10, 38)
        Size     = New-Object System.Drawing.Size(660, 390)
        BackColor = $colorInput
        ForeColor = $colorText
        BorderStyle = 'FixedSingle'
        Font      = New-Object System.Drawing.Font('Consolas', 9)
        ScrollAlwaysVisible = $true
    }

    # Build display strings with indentation based on depth
    $maxDepth = ($ous | Measure-Object -Property Depth -Maximum).Maximum
    foreach ($ou in $ous) {
        $indent = '    ' * ($ou.Depth - 1)
        $displayText = "$indent$($ou.DistinguishedName)"
        [void]$ouListBox.Items.Add($displayText)
    }
    $ouForm.Controls.Add($ouListBox)

    $btnSelectOU = New-StyledButton -Text 'Select' -X 480 -Y 435 -Width 80 -Height 30
    $ouForm.Controls.Add($btnSelectOU)

    $btnCancelOU = New-StyledButton -Text 'Cancel' -X 570 -Y 435 -Width 80 -Height 30
    $btnCancelOU.BackColor = $colorPanelAlt
    $ouForm.Controls.Add($btnCancelOU)

    $btnSelectOU.Add_Click({
        if ($ouListBox.SelectedIndex -ge 0) {
            $selectedText = $ouListBox.SelectedItem.ToString().TrimStart()
            $txtOUPath.Text = $selectedText
            Write-Log -Message "OU selected: $selectedText" -Level 'INFO'
        }
        $ouForm.Close()
    })

    $btnCancelOU.Add_Click({
        $ouForm.Close()
    })

    $ouListBox.Add_DoubleClick({
        if ($ouListBox.SelectedIndex -ge 0) {
            $selectedText = $ouListBox.SelectedItem.ToString().TrimStart()
            $txtOUPath.Text = $selectedText
            Write-Log -Message "OU selected: $selectedText" -Level 'INFO'
        }
        $ouForm.Close()
    })

    $lblDomainStatus.Text = "Status: Found $($ous.Count) OUs. Select from popup."
    $lblDomainStatus.ForeColor = $colorSuccess

    [void]$ouForm.ShowDialog($mainForm)
})

# --- Rename Only Button ---
$btnRenameOnly.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $newName = $txtNewName.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($newName)) {
        $lblDomainStatus.Text = 'Status: Enter a new computer name.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }

    $result = Invoke-ComputerRename -NewName $newName
    if ($result.Success) {
        $lblDomainStatus.Text = "Status: $($result.Message)"
        $lblDomainStatus.ForeColor = $colorSuccess
        [System.Windows.Forms.MessageBox]::Show($result.Message, 'Rename Computer', 'OK', 'Information') | Out-Null
    } else {
        $lblDomainStatus.Text = "Status: $($result.Message)"
        $lblDomainStatus.ForeColor = $colorError
        [System.Windows.Forms.MessageBox]::Show($result.Message, 'Error', 'OK', 'Error') | Out-Null
    }
})

# --- Join Domain Button ---
$btnJoinDomain.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $domain = $cmbDomain.SelectedItem
    $username = $txtDomainUser.Text.Trim()
    $password = $txtDomainPass.Text
    $newName = $txtNewName.Text.Trim()
    $ouPath = $txtOUPath.Text.Trim()

    # Resolve DC hostname from dropdown
    $dcHost = ''
    $dcSelection = $cmbDomainController.SelectedItem
    if ($dcSelection -and $dcSelection -ne 'Auto (DNS resolve)') {
        $parts = $dcSelection -split ' - '
        if ($parts.Count -ge 2) {
            $dcHost = $parts[-1].Trim()
        }
    }
    $legacyReuse = $chkLegacyReuse.Checked

    if (-not $domain) {
        $lblDomainStatus.Text = 'Status: Select a domain.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }
    if ([string]::IsNullOrWhiteSpace($username)) {
        $lblDomainStatus.Text = 'Status: Enter domain username.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }
    if ([string]::IsNullOrWhiteSpace($password)) {
        $lblDomainStatus.Text = 'Status: Enter domain password.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }

    $lblDomainStatus.Text = 'Status: Joining domain... please wait.'
    $lblDomainStatus.ForeColor = $colorAccent2
    [System.Windows.Forms.Application]::DoEvents()

    $result = Invoke-DomainJoin -DomainName $domain -Username $username -Password $password -NewComputerName $newName -OUPath $ouPath -Server $dcHost -EnableLegacyReuse:$legacyReuse

    if ($result.Success) {
        $lblDomainStatus.Text = "Status: $($result.Message)"
        $lblDomainStatus.ForeColor = $colorSuccess
        $reboot = [System.Windows.Forms.MessageBox]::Show("$($result.Message)`n`nWould you like to reboot now?", 'Domain Join', 'YesNo', 'Question')
        if ($reboot -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log -Message 'Rebooting computer as requested by user.' -Level 'INFO'
            Restart-Computer -Force
        }
    } else {
        $lblDomainStatus.Text = "Status: $($result.Message)"
        $lblDomainStatus.ForeColor = $colorError
        [System.Windows.Forms.MessageBox]::Show($result.Message, 'Domain Join Failed', 'OK', 'Error') | Out-Null
    }
})

# --- Change Domain Button ---
$btnChangeDomain.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $domain = $cmbDomain.SelectedItem
    $username = $txtDomainUser.Text.Trim()
    $password = $txtDomainPass.Text
    $newName = $txtNewName.Text.Trim()
    $ouPath = $txtOUPath.Text.Trim()

    # Resolve DC hostname from dropdown
    $dcHost = ''
    $dcSelection = $cmbDomainController.SelectedItem
    if ($dcSelection -and $dcSelection -ne 'Auto (DNS resolve)') {
        $parts = $dcSelection -split ' - '
        if ($parts.Count -ge 2) {
            $dcHost = $parts[-1].Trim()
        }
    }
    $legacyReuse = $chkLegacyReuse.Checked

    if (-not $domain) {
        $lblDomainStatus.Text = 'Status: Select a domain.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }
    if ([string]::IsNullOrWhiteSpace($username)) {
        $lblDomainStatus.Text = 'Status: Enter domain username.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }
    if ([string]::IsNullOrWhiteSpace($password)) {
        $lblDomainStatus.Text = 'Status: Enter domain password.'
        $lblDomainStatus.ForeColor = $colorWarning
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "This will unjoin the current domain/workgroup and join '$domain'.`n`nComputer name will be set to: $newName`n`nDomain Controller: $(if($dcHost){$dcHost}else{'Auto (DNS)'})`n`nContinue?",
        'Confirm Domain Change',
        'YesNo',
        'Warning'
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $lblDomainStatus.Text = 'Status: Changing domain... please wait.'
    $lblDomainStatus.ForeColor = $colorAccent2
    [System.Windows.Forms.Application]::DoEvents()

    $result = Invoke-DomainJoin -DomainName $domain -Username $username -Password $password -NewComputerName $newName -OUPath $ouPath -Server $dcHost -EnableLegacyReuse:$legacyReuse -ChangeDomain

    if ($result.Success) {
        $lblDomainStatus.Text = "Status: $($result.Message)"
        $lblDomainStatus.ForeColor = $colorSuccess
        $reboot = [System.Windows.Forms.MessageBox]::Show("$($result.Message)`n`nWould you like to reboot now?", 'Domain Change', 'YesNo', 'Question')
        if ($reboot -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log -Message 'Rebooting computer as requested by user.' -Level 'INFO'
            Restart-Computer -Force
        }
    } else {
        $lblDomainStatus.Text = "Status: $($result.Message)"
        $lblDomainStatus.ForeColor = $colorError
        [System.Windows.Forms.MessageBox]::Show($result.Message, 'Domain Change Failed', 'OK', 'Error') | Out-Null
    }
})

# ============================================================================
# SECTION 3: Software Installation
# ============================================================================
$gbSoftware = New-StyledGroupBox -Text 'Software Installation' -X 480 -Y 140 -Width 455 -Height 380
$mainForm.Controls.Add($gbSoftware)

$lblSoftwarePath = New-StyledLabel -Text 'Software Folder:' -X 15 -Y 30 -Width 120
$gbSoftware.Controls.Add($lblSoftwarePath)

$txtSoftwarePath = New-StyledTextBox -X 135 -Y 27 -Width 215
$gbSoftware.Controls.Add($txtSoftwarePath)
$txtSoftwarePath.Text = $script:DefaultSoftwarePath

$btnBrowseSoftware = New-StyledButton -Text 'Browse...' -X 355 -Y 25 -Width 80 -Height 28
$gbSoftware.Controls.Add($btnBrowseSoftware)

$btnBrowseSoftware.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = 'Select software installation folder'
        ShowNewFolderButton = $false
        RootFolder = [System.Environment+SpecialFolder]::MyComputer
    }
    # Try to open at current path
    if ($txtSoftwarePath.Text -and (Test-Path $txtSoftwarePath.Text)) {
        $folderDialog.SelectedPath = $txtSoftwarePath.Text
    }
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtSoftwarePath.Text = $folderDialog.SelectedPath
    }
})

# --- DataGridView for software packages ---
$dgvSoftware = New-Object System.Windows.Forms.DataGridView -Property @{
    Location = New-Object System.Drawing.Point(15, 60)
    Size     = New-Object System.Drawing.Size(420, 200)
    BackgroundColor = $colorPanelAlt
    BorderStyle = 'None'
    AllowUserToAddRows = $false
    AllowUserToDeleteRows = $false
    AllowUserToResizeRows = $false
    RowHeadersVisible = $false
    AutoSizeColumnsMode = 'Fill'
    SelectionMode = 'FullRowSelect'
    MultiSelect = $false
    ReadOnly = $false
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    DefaultCellStyle = @{
        BackColor = $colorPanelAlt
        ForeColor = $colorText
        SelectionBackColor = $colorAccent
        SelectionForeColor = [System.Drawing.Color]::White
    }
    ColumnHeadersDefaultCellStyle = @{
        BackColor = $colorAccent
        ForeColor = [System.Drawing.Color]::White
    }
    EnableHeadersVisualStyles = $false
}
$dgvSoftware.ColumnHeadersHeight = 28
$gbSoftware.Controls.Add($dgvSoftware)

# --- Refresh Software List Button ---
$btnRefreshSoftware = New-StyledButton -Text 'Scan for Installers' -X 15 -Y 270 -Width 140 -Height 30
$gbSoftware.Controls.Add($btnRefreshSoftware)

$btnInstallSelected = New-StyledButton -Text 'Install Selected' -X 165 -Y 270 -Width 120 -Height 30
$gbSoftware.Controls.Add($btnInstallSelected)

$btnInstallAll = New-StyledButton -Text 'Install All' -X 295 -Y 270 -Width 140 -Height 30
$gbSoftware.Controls.Add($btnInstallAll)

$lblSoftwareStatus = New-Object System.Windows.Forms.Label -Property @{
    Text     = 'Click "Scan for Installers" to detect software.'
    Location = New-Object System.Drawing.Point(15, 305)
    Size     = New-Object System.Drawing.Size(420, 20)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font      = New-Object System.Drawing.Font('Segoe UI', 8)
}
$gbSoftware.Controls.Add($lblSoftwareStatus)

# --- Initialize DataGridView columns ---
$colInstall = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn -Property @{
    Name = 'Install'
    HeaderText = 'Install'
    Width = 50
    ReadOnly = $false
}
$dgvSoftware.Columns.Add($colInstall)

$colPackage = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
    Name = 'Package'
    HeaderText = 'Package'
    ReadOnly = $true
    Width = 100
}
$dgvSoftware.Columns.Add($colPackage)

$colFile = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
    Name = 'Installer'
    HeaderText = 'Installer File'
    ReadOnly = $true
    Width = 130
}
$dgvSoftware.Columns.Add($colFile)

$colSwitches = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
    Name = 'Switches'
    HeaderText = 'Silent Switches'
    ReadOnly = $false
    Width = 150
}
$dgvSoftware.Columns.Add($colSwitches)

$colPath = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
    Name = 'Path'
    HeaderText = 'Full Path'
    ReadOnly = $true
    Visible = $false
}
$dgvSoftware.Columns.Add($colPath)

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{
    Name = 'Status'
    HeaderText = 'Status'
    ReadOnly = $true
    Width = 60
}
$dgvSoftware.Columns.Add($colStatus)

# --- Refresh Software Button ---
$btnRefreshSoftware.Add_Click({
    $dgvSoftware.Rows.Clear()
    $folder = $txtSoftwarePath.Text.Trim()

    if (-not (Test-Path -Path $folder)) {
        $lblSoftwareStatus.Text = "Folder not found: $folder"
        $lblSoftwareStatus.ForeColor = $colorError
        return
    }

    Write-Log -Message "Scanning for software installers in: $folder" -Level 'INFO'
    $found = Find-SoftwareInstallers -FolderPath $folder

    if ($found.Count -eq 0) {
        $lblSoftwareStatus.Text = 'No matching installers found in folder.'
        $lblSoftwareStatus.ForeColor = $colorWarning
        # Still show the expected packages with "Not Found" status
        foreach ($pkg in $script:SoftwarePackages) {
            $rowIdx = $dgvSoftware.Rows.Add()
            $dgvSoftware.Rows[$rowIdx].Cells['Install'].Value = $false
            $dgvSoftware.Rows[$rowIdx].Cells['Package'].Value = $pkg.Name
            $dgvSoftware.Rows[$rowIdx].Cells['Installer'].Value = 'Not Found'
            $dgvSoftware.Rows[$rowIdx].Cells['Switches'].Value = $pkg.SilentSwitches
            $dgvSoftware.Rows[$rowIdx].Cells['Path'].Value = ''
            $dgvSoftware.Rows[$rowIdx].Cells['Status'].Value = 'Missing'
            $dgvSoftware.Rows[$rowIdx].DefaultCellStyle.ForeColor = $colorError
        }
        return
    }

    # Add found packages
    $foundPackages = @()
    foreach ($item in $found) {
        $foundPackages += $item.Package
        $rowIdx = $dgvSoftware.Rows.Add()
        $dgvSoftware.Rows[$rowIdx].Cells['Install'].Value = $true
        $dgvSoftware.Rows[$rowIdx].Cells['Package'].Value = $item.Package
        $dgvSoftware.Rows[$rowIdx].Cells['Installer'].Value = $item.FileName
        $dgvSoftware.Rows[$rowIdx].Cells['Switches'].Value = $item.SilentSwitches
        $dgvSoftware.Rows[$rowIdx].Cells['Path'].Value = $item.Installer
        $dgvSoftware.Rows[$rowIdx].Cells['Status'].Value = 'Ready'
    }

    # Add missing packages
    foreach ($pkg in $script:SoftwarePackages) {
        if ($pkg.Name -notin $foundPackages) {
            $rowIdx = $dgvSoftware.Rows.Add()
            $dgvSoftware.Rows[$rowIdx].Cells['Install'].Value = $false
            $dgvSoftware.Rows[$rowIdx].Cells['Package'].Value = $pkg.Name
            $dgvSoftware.Rows[$rowIdx].Cells['Installer'].Value = 'Not Found'
            $dgvSoftware.Rows[$rowIdx].Cells['Switches'].Value = $pkg.SilentSwitches
            $dgvSoftware.Rows[$rowIdx].Cells['Path'].Value = ''
            $dgvSoftware.Rows[$rowIdx].Cells['Status'].Value = 'Missing'
            $dgvSoftware.Rows[$rowIdx].DefaultCellStyle.ForeColor = $colorError
        }
    }

    $lblSoftwareStatus.Text = "Found $($found.Count) of $($script:SoftwarePackages.Count) packages."
    $lblSoftwareStatus.ForeColor = $colorSuccess
    Write-Log -Message "Software scan complete: found $($found.Count) of $($script:SoftwarePackages.Count) packages." -Level 'SUCCESS'
})

# --- Install Selected Button ---
$btnInstallSelected.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $packagesToInstall = @()
    foreach ($row in $dgvSoftware.Rows) {
        if ($row.Cells['Install'].Value -eq $true -and $row.Cells['Path'].Value -and $row.Cells['Path'].Value -ne '') {
            $packagesToInstall += [PSCustomObject]@{
                Package = $row.Cells['Package'].Value
                Path    = $row.Cells['Path'].Value
                Switches = $row.Cells['Switches'].Value
                RowIndex = $row.Index
            }
        }
    }

    if ($packagesToInstall.Count -eq 0) {
        $lblSoftwareStatus.Text = 'No packages selected for installation.'
        $lblSoftwareStatus.ForeColor = $colorWarning
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Install $($packagesToInstall.Count) package(s)?`n`n$(($packagesToInstall | ForEach-Object { "$($_.Package): $($_.Switches)" }) -join "`n")",
        'Confirm Installation',
        'YesNo',
        'Question'
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $successCount = 0
    $failCount = 0

    foreach ($pkg in $packagesToInstall) {
        $dgvSoftware.Rows[$pkg.RowIndex].Cells['Status'].Value = 'Installing...'
        $dgvSoftware.Rows[$pkg.RowIndex].DefaultCellStyle.ForeColor = $colorAccent2
        [System.Windows.Forms.Application]::DoEvents()

        $result = Install-SoftwarePackage -DisplayName $pkg.Package -InstallerPath $pkg.Path -SilentSwitches $pkg.Switches

        if ($result.Success) {
            $dgvSoftware.Rows[$pkg.RowIndex].Cells['Status'].Value = 'Installed'
            $dgvSoftware.Rows[$pkg.RowIndex].DefaultCellStyle.ForeColor = $colorSuccess
            $dgvSoftware.Rows[$pkg.RowIndex].Cells['Install'].Value = $false
            $successCount++
        } else {
            $dgvSoftware.Rows[$pkg.RowIndex].Cells['Status'].Value = 'Failed'
            $dgvSoftware.Rows[$pkg.RowIndex].DefaultCellStyle.ForeColor = $colorError
            $failCount++
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblSoftwareStatus.Text = "Installation complete: $successCount succeeded, $failCount failed."
    $lblSoftwareStatus.ForeColor = if ($failCount -gt 0) { $colorWarning } else { $colorSuccess }
})

# --- Install All Button ---
$btnInstallAll.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $packagesToInstall = @()
    foreach ($row in $dgvSoftware.Rows) {
        if ($row.Cells['Path'].Value -and $row.Cells['Path'].Value -ne '') {
            $packagesToInstall += [PSCustomObject]@{
                Package = $row.Cells['Package'].Value
                Path    = $row.Cells['Path'].Value
                Switches = $row.Cells['Switches'].Value
                RowIndex = $row.Index
            }
        }
    }

    if ($packagesToInstall.Count -eq 0) {
        $lblSoftwareStatus.Text = 'No installers found. Click "Scan for Installers" first.'
        $lblSoftwareStatus.ForeColor = $colorWarning
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Install all $($packagesToInstall.Count) package(s)?`n`n$(($packagesToInstall | ForEach-Object { "$($_.Package): $($_.Switches)" }) -join "`n")",
        'Confirm Installation',
        'YesNo',
        'Question'
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $successCount = 0
    $failCount = 0

    foreach ($pkg in $packagesToInstall) {
        $dgvSoftware.Rows[$pkg.RowIndex].Cells['Status'].Value = 'Installing...'
        $dgvSoftware.Rows[$pkg.RowIndex].DefaultCellStyle.ForeColor = $colorAccent2
        [System.Windows.Forms.Application]::DoEvents()

        $result = Install-SoftwarePackage -DisplayName $pkg.Package -InstallerPath $pkg.Path -SilentSwitches $pkg.Switches

        if ($result.Success) {
            $dgvSoftware.Rows[$pkg.RowIndex].Cells['Status'].Value = 'Installed'
            $dgvSoftware.Rows[$pkg.RowIndex].DefaultCellStyle.ForeColor = $colorSuccess
            $dgvSoftware.Rows[$pkg.RowIndex].Cells['Install'].Value = $false
            $successCount++
        } else {
            $dgvSoftware.Rows[$pkg.RowIndex].Cells['Status'].Value = 'Failed'
            $dgvSoftware.Rows[$pkg.RowIndex].DefaultCellStyle.ForeColor = $colorError
            $failCount++
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblSoftwareStatus.Text = "Installation complete: $successCount succeeded, $failCount failed."
    $lblSoftwareStatus.ForeColor = if ($failCount -gt 0) { $colorWarning } else { $colorSuccess }
})

# ============================================================================
# SECTION 4: Network Configuration (IPv4)
# ============================================================================
$gbNetwork = New-StyledGroupBox -Text 'IPv4 Network Configuration' -X 15 -Y 530 -Width 450 -Height 180
$mainForm.Controls.Add($gbNetwork)

$lblNic = New-StyledLabel -Text 'Network Adapter:' -X 15 -Y 30 -Width 130
$gbNetwork.Controls.Add($lblNic)

$cmbNic = New-StyledComboBox -X 150 -Y 27 -Width 280
$gbNetwork.Controls.Add($cmbNic)

$btnRefreshNic = New-StyledButton -Text 'Refresh' -X 150 -Y 55 -Width 80 -Height 24
$btnRefreshNic.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$gbNetwork.Controls.Add($btnRefreshNic)

$lblIP = New-StyledLabel -Text 'IP Address:' -X 15 -Y 85 -Width 130
$gbNetwork.Controls.Add($lblIP)

$txtIP = New-StyledTextBox -X 150 -Y 82 -Width 130
$gbNetwork.Controls.Add($txtIP)

$lblMask = New-StyledLabel -Text 'Subnet Mask:' -X 290 -Y 85 -Width 80
$gbNetwork.Controls.Add($lblMask)

$txtMask = New-StyledTextBox -X 370 -Y 82 -Width 60
$gbNetwork.Controls.Add($txtMask)

$lblGateway = New-StyledLabel -Text 'Default Gateway:' -X 15 -Y 115 -Width 130
$gbNetwork.Controls.Add($lblGateway)

$txtGateway = New-StyledTextBox -X 150 -Y 112 -Width 130
$gbNetwork.Controls.Add($txtGateway)

$lblDNS = New-StyledLabel -Text 'DNS Servers:' -X 15 -Y 145 -Width 130
$gbNetwork.Controls.Add($lblDNS)

$txtDNS1 = New-StyledTextBox -X 150 -Y 142 -Width 130
$gbNetwork.Controls.Add($txtDNS1)

$txtDNS2 = New-StyledTextBox -X 290 -Y 142 -Width 140
$gbNetwork.Controls.Add($txtDNS2)

$lblDnsHint = New-StyledLabel -Text '(DNS 2 optional)' -X 290 -Y 145 -Width 140 -Height 14
$lblDnsHint.Font = New-Object System.Drawing.Font('Segoe UI', 7, [System.Drawing.FontStyle]::Italic)
$gbNetwork.Controls.Add($lblDnsHint)

# --- Apply Network Config Button ---
$btnApplyNetwork = New-StyledButton -Text 'Apply Network Config' -X 150 -Y 55 -Width 0 -Height 0
# Reposition
$btnApplyNetwork.Location = New-Object System.Drawing.Point(290, 55)
$btnApplyNetwork.Size = New-Object System.Drawing.Size(140, 24)
$btnApplyNetwork.Text = 'Apply Network Config'
$btnApplyNetwork.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Bold)
$gbNetwork.Controls.Add($btnApplyNetwork)

$lblNetworkStatus = New-Object System.Windows.Forms.Label -Property @{
    Text     = ''
    Location = New-Object System.Drawing.Point(15, 165)
    Size     = New-Object System.Drawing.Size(420, 12)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font      = New-Object System.Drawing.Font('Segoe UI', 8)
}
$gbNetwork.Controls.Add($lblNetworkStatus)

# --- Populate NIC list ---
function Update-NicList {
    $cmbNic.Items.Clear()
    $script:NicList = Get-NetworkAdapters
    foreach ($nic in $script:NicList) {
        $displayText = "$($nic.Name) [$($nic.Status)]"
        [void]$cmbNic.Items.Add($displayText)
    }
    if ($cmbNic.Items.Count -gt 0) {
        $cmbNic.SelectedIndex = 0
    }
}

$btnRefreshNic.Add_Click({
    Update-NicList
    $lblNetworkStatus.Text = "Found $($cmbNic.Items.Count) network adapter(s)."
    $lblNetworkStatus.ForeColor = $colorSuccess
})

# --- Apply Network Config ---
$btnApplyNetwork.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $selectedNic = $cmbNic.SelectedItem
    if (-not $selectedNic) {
        $lblNetworkStatus.Text = 'Select a network adapter.'
        $lblNetworkStatus.ForeColor = $colorWarning
        return
    }

    # Extract adapter name (before the [Status] part)
    $adapterName = ($selectedNic -split ' \[')[0]

    $ipAddress = $txtIP.Text.Trim()
    $subnetMask = $txtMask.Text.Trim()
    $gateway = $txtGateway.Text.Trim()
    $dns = @()
    if ($txtDNS1.Text.Trim()) { $dns += $txtDNS1.Text.Trim() }
    if ($txtDNS2.Text.Trim()) { $dns += $txtDNS2.Text.Trim() }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Apply the following network configuration?`n`nAdapter: $adapterName`nIP: $ipAddress`nMask: $subnetMask`nGateway: $gateway`nDNS: $($dns -join ', ')",
        'Confirm Network Configuration',
        'YesNo',
        'Question'
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $result = Set-Ipv4Config -AdapterName $adapterName -IPAddress $ipAddress -SubnetMask $subnetMask -Gateway $gateway -DNSServers $dns

    if ($result.Success) {
        $lblNetworkStatus.Text = "Network configuration applied to '$adapterName'."
        $lblNetworkStatus.ForeColor = $colorSuccess
        [System.Windows.Forms.MessageBox]::Show("Network configuration applied successfully to '$adapterName'.", 'Network Config', 'OK', 'Information') | Out-Null
    } else {
        $lblNetworkStatus.Text = "Error: $($result.Errors -join ', ')"
        $lblNetworkStatus.ForeColor = $colorError
        [System.Windows.Forms.MessageBox]::Show("Failed to apply network configuration:`n`n$($result.Errors -join "`n")", 'Network Config Error', 'OK', 'Error') | Out-Null
    }
})

# Initialize NIC list on form load
$mainForm.Add_Load({
    Update-NicList
})

# ============================================================================
# SECTION 5: Quick Launch Tools
# ============================================================================
$gbTools = New-StyledGroupBox -Text 'Quick Launch Tools' -X 480 -Y 530 -Width 455 -Height 230
$mainForm.Controls.Add($gbTools)

$btnLaunchCmd = New-StyledButton -Text 'CMD (Admin)' -X 15 -Y 30 -Width 130 -Height 40
$gbTools.Controls.Add($btnLaunchCmd)

$btnLaunchNotepad = New-StyledButton -Text 'Notepad' -X 155 -Y 30 -Width 130 -Height 40
$gbTools.Controls.Add($btnLaunchNotepad)

$btnLaunchCompMgmt = New-StyledButton -Text 'Computer Mgmt (Admin)' -X 295 -Y 30 -Width 140 -Height 40
$gbTools.Controls.Add($btnLaunchCompMgmt)

$btnLaunchPwrShell = New-StyledButton -Text 'PowerShell (Admin)' -X 15 -Y 80 -Width 130 -Height 40
$gbTools.Controls.Add($btnLaunchPwrShell)

$btnLaunchPwrShellISE = New-StyledButton -Text 'PS ISE (Admin)' -X 155 -Y 80 -Width 130 -Height 40
$gbTools.Controls.Add($btnLaunchPwrShellISE)

$btnLaunchRegEdit = New-StyledButton -Text 'Reg Editor (Admin)' -X 295 -Y 80 -Width 140 -Height 40
$gbTools.Controls.Add($btnLaunchRegEdit)

$btnLaunchEventVwr = New-StyledButton -Text 'Event Viewer (Admin)' -X 15 -Y 130 -Width 140 -Height 40
$gbTools.Controls.Add($btnLaunchEventVwr)

$btnOpenLog = New-StyledButton -Text 'Open Domain Join GUI.log' -X 165 -Y 130 -Width 130 -Height 40
$gbTools.Controls.Add($btnOpenLog)

$btnOpenAdminFolder = New-StyledButton -Text 'Open C:\IT' -X 305 -Y 130 -Width 130 -Height 40
$gbTools.Controls.Add($btnOpenAdminFolder)

$btnReboot = New-StyledButton -Text 'Reboot' -X 15 -Y 180 -Width 140 -Height 40
$btnReboot.BackColor = $colorError
$gbTools.Controls.Add($btnReboot)

# --- Quick Launch Button Handlers ---
$btnLaunchCmd.Add_Click({
    try {
        Start-Process -FilePath 'cmd.exe' -Verb RunAs -ErrorAction Stop
        Write-Log -Message 'Launched CMD as administrator.' -Level 'INFO'
    } catch {
        Start-Process -FilePath 'cmd.exe' -ErrorAction SilentlyContinue
        Write-Log -Message 'Launched CMD (could not elevate).' -Level 'WARNING'
    }
})

$btnLaunchNotepad.Add_Click({
    Start-Process -FilePath 'notepad.exe' -ErrorAction SilentlyContinue
    Write-Log -Message 'Launched Notepad.' -Level 'INFO'
})

$btnLaunchCompMgmt.Add_Click({
    try {
        Start-Process -FilePath 'compmgmt.msc' -Verb RunAs -ErrorAction Stop
        Write-Log -Message 'Launched Computer Management as administrator.' -Level 'INFO'
    } catch {
        Start-Process -FilePath 'compmgmt.msc' -ErrorAction SilentlyContinue
        Write-Log -Message 'Launched Computer Management (could not elevate).' -Level 'WARNING'
    }
})

$btnLaunchPwrShell.Add_Click({
    try {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ErrorAction Stop
        Write-Log -Message 'Launched PowerShell as administrator.' -Level 'INFO'
    } catch {
        Start-Process -FilePath 'powershell.exe' -ErrorAction SilentlyContinue
        Write-Log -Message 'Launched PowerShell (could not elevate).' -Level 'WARNING'
    }
})

$btnLaunchPwrShellISE.Add_Click({
    try {
        Start-Process -FilePath 'powershell_ise.exe' -Verb RunAs -ErrorAction Stop
        Write-Log -Message 'Launched PowerShell ISE as administrator.' -Level 'INFO'
    } catch {
        Start-Process -FilePath 'powershell_ise.exe' -ErrorAction SilentlyContinue
        Write-Log -Message 'Launched PowerShell ISE (could not elevate).' -Level 'WARNING'
    }
})

$btnLaunchRegEdit.Add_Click({
    try {
        Start-Process -FilePath 'regedit.exe' -Verb RunAs -ErrorAction Stop
        Write-Log -Message 'Launched Registry Editor as administrator.' -Level 'INFO'
    } catch {
        Start-Process -FilePath 'regedit.exe' -ErrorAction SilentlyContinue
        Write-Log -Message 'Launched Registry Editor (could not elevate).' -Level 'WARNING'
    }
})

$btnLaunchEventVwr.Add_Click({
    try {
        Start-Process -FilePath 'eventvwr.msc' -Verb RunAs -ErrorAction Stop
        Write-Log -Message 'Launched Event Viewer as administrator.' -Level 'INFO'
    } catch {
        Start-Process -FilePath 'eventvwr.msc' -ErrorAction SilentlyContinue
        Write-Log -Message 'Launched Event Viewer (could not elevate).' -Level 'WARNING'
    }
})

$btnOpenLog.Add_Click({
    if ($script:LogPath -and (Test-Path -Path $script:LogPath)) {
        Start-Process -FilePath 'notepad.exe' -ArgumentList $script:LogPath -ErrorAction SilentlyContinue
        Write-Log -Message "Opened log file: $script:LogPath" -Level 'INFO'
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log file not found at: $script:LogPath", 'File Not Found', 'OK', 'Warning') | Out-Null
    }
})

$btnOpenAdminFolder.Add_Click({
    $adminPath = $txtSoftwarePath.Text.Trim()
    if (-not $adminPath) { $adminPath = $script:DefaultSoftwarePath }
    if (Test-Path -Path $adminPath) {
        Start-Process -FilePath 'explorer.exe' -ArgumentList $adminPath -ErrorAction SilentlyContinue
        Write-Log -Message "Opened folder: $adminPath" -Level 'INFO'
    } else {
        [System.Windows.Forms.MessageBox]::Show("Folder not found: $adminPath", 'Not Found', 'OK', 'Warning') | Out-Null
    }
})

$btnReboot.Add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        'Are you sure you want to reboot this computer NOW?',
        'Confirm Reboot',
        'YesNo',
        'Warning'
    )
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log -Message 'User confirmed reboot.' -Level 'INFO'
        Restart-Computer -Force
    }
})

# ============================================================================
# SECTION 6: System Prep (Step 3 - PrepSystemForDeployment)
# ============================================================================
$gbPrep = New-StyledGroupBox -Text 'System Prep (Step 3)' -X 15 -Y 770 -Width 920 -Height 300
$mainForm.Controls.Add($gbPrep)

$lblSystemType = New-StyledLabel -Text 'System Type:' -X 15 -Y 30 -Width 120
$gbPrep.Controls.Add($lblSystemType)

$cmbSystemType = New-StyledComboBox -X 135 -Y 27 -Width 200
$gbPrep.Controls.Add($cmbSystemType)
[void]$cmbSystemType.Items.Add('IGT/IGTSAP')
[void]$cmbSystemType.Items.Add('Legacy-Everi Production')
[void]$cmbSystemType.Items.Add('Legacy-Everi Non-Production')
[void]$cmbSystemType.Items.Add('Legacy-IGT')
$cmbSystemType.SelectedIndex = 0

$lblSCCMTier = New-StyledLabel -Text 'SCCM Tier:' -X 345 -Y 30 -Width 90
$gbPrep.Controls.Add($lblSCCMTier)

$cmbSCCMTier = New-StyledComboBox -X 435 -Y 27 -Width 180
$gbPrep.Controls.Add($cmbSCCMTier)
[void]$cmbSCCMTier.Items.Add('Tier 1/2 L-Everi')
[void]$cmbSCCMTier.Items.Add('Tier 3/4 L-Everi')
[void]$cmbSCCMTier.Items.Add('(Skip - IGT)')
$cmbSCCMTier.SelectedIndex = 0

# Prep step checkboxes
$chkPrepBGInfo = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'Remove BGInfo'
    Location = New-Object System.Drawing.Point(15, 65)
    Size = New-Object System.Drawing.Size(130, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepBGInfo)

$chkPrepCribl = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'Install CRIBL'
    Location = New-Object System.Drawing.Point(155, 65)
    Size = New-Object System.Drawing.Size(130, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepCribl)

$chkPrepSCCM = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'Install SCCM Client'
    Location = New-Object System.Drawing.Point(295, 65)
    Size = New-Object System.Drawing.Size(150, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepSCCM)

$chkPrepAdmin = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'Admin Password Non-Expire'
    Location = New-Object System.Drawing.Point(455, 65)
    Size = New-Object System.Drawing.Size(180, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepAdmin)

$chkPrepRapid7 = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'Rapid7 Service + ScanAssist'
    Location = New-Object System.Drawing.Point(645, 65)
    Size = New-Object System.Drawing.Size(180, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepRapid7)

$chkPrepCrowdStrike = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'Install CrowdStrike'
    Location = New-Object System.Drawing.Point(15, 95)
    Size = New-Object System.Drawing.Size(140, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepCrowdStrike)

$chkPrepGpUpdate = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'GPUpdate + Credential Cleanup'
    Location = New-Object System.Drawing.Point(165, 95)
    Size = New-Object System.Drawing.Size(200, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepGpUpdate)

$chkPrepIISCrypto = New-Object System.Windows.Forms.CheckBox -Property @{
    Text = 'IISCrypto (Disable Weak Ciphers)'
    Location = New-Object System.Drawing.Point(375, 95)
    Size = New-Object System.Drawing.Size(220, 24)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
    Checked = $true
}
$gbPrep.Controls.Add($chkPrepIISCrypto)

# Run Prep button
$btnRunPrep = New-StyledButton -Text 'Run System Prep' -X 15 -Y 130 -Width 160 -Height 36
$btnRunPrep.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)
$gbPrep.Controls.Add($btnRunPrep)

# Select All / Deselect All buttons
$btnSelectAllPrep = New-StyledButton -Text 'Select All' -X 185 -Y 130 -Width 100 -Height 36
$btnSelectAllPrep.BackColor = $colorPanelAlt
$gbPrep.Controls.Add($btnSelectAllPrep)

$btnDeselectAllPrep = New-StyledButton -Text 'Deselect All' -X 295 -Y 130 -Width 100 -Height 36
$btnDeselectAllPrep.BackColor = $colorPanelAlt
$gbPrep.Controls.Add($btnDeselectAllPrep)

# Prep status label
$lblPrepStatus = New-Object System.Windows.Forms.Label -Property @{
    Text = 'Select steps and click Run System Prep.'
    Location = New-Object System.Drawing.Point(15, 175)
    Size = New-Object System.Drawing.Size(890, 20)
    ForeColor = $colorTextDim
    BackColor = [System.Drawing.Color]::Transparent
    Font = New-Object System.Drawing.Font('Segoe UI', 8)
}
$gbPrep.Controls.Add($lblPrepStatus)

# Prep results text box
$txtPrepResults = New-Object System.Windows.Forms.TextBox -Property @{
    Location = New-Object System.Drawing.Point(15, 200)
    Size = New-Object System.Drawing.Size(890, 85)
    BackColor = $colorPanelAlt
    ForeColor = $colorText
    BorderStyle = 'FixedSingle'
    Font = New-Object System.Drawing.Font('Consolas', 8)
    ReadOnly = $true
    Multiline = $true
    ScrollBars = 'Vertical'
    Text = 'Prep results will appear here...'
}
$gbPrep.Controls.Add($txtPrepResults)

# Select All handler
$btnSelectAllPrep.Add_Click({
    $chkPrepBGInfo.Checked = $true
    $chkPrepCribl.Checked = $true
    $chkPrepSCCM.Checked = $true
    $chkPrepAdmin.Checked = $true
    $chkPrepRapid7.Checked = $true
    $chkPrepCrowdStrike.Checked = $true
    $chkPrepGpUpdate.Checked = $true
    $chkPrepIISCrypto.Checked = $true
})

# Deselect All handler
$btnDeselectAllPrep.Add_Click({
    $chkPrepBGInfo.Checked = $false
    $chkPrepCribl.Checked = $false
    $chkPrepSCCM.Checked = $false
    $chkPrepAdmin.Checked = $false
    $chkPrepRapid7.Checked = $false
    $chkPrepCrowdStrike.Checked = $false
    $chkPrepGpUpdate.Checked = $false
    $chkPrepIISCrypto.Checked = $false
})

# System Type change handler - auto-adjust SCCM tier
$cmbSystemType.Add_SelectedIndexChanged({
    $sysType = $cmbSystemType.SelectedItem
    if ($sysType -eq 'IGT/IGTSAP' -or $sysType -eq 'Legacy-IGT') {
        # IGT systems skip SCCM
        $cmbSCCMTier.SelectedIndex = $cmbSCCMTier.Items.IndexOf('(Skip - IGT)')
        $chkPrepSCCM.Checked = $false
        $chkPrepBGInfo.Checked = $true
    } elseif ($sysType -eq 'Legacy-Everi Production' -or $sysType -eq 'Legacy-Everi Non-Production') {
        # Everi systems get SCCM
        if ($cmbSCCMTier.SelectedIndex -eq $cmbSCCMTier.Items.IndexOf('(Skip - IGT)')) {
            $cmbSCCMTier.SelectedIndex = 0
        }
        $chkPrepSCCM.Checked = $true
        $chkPrepBGInfo.Checked = $false
    }
})

# Run System Prep handler
$btnRunPrep.Add_Click({
    if (-not $script:LogPath) {
        [System.Windows.Forms.MessageBox]::Show('Please set a log file location first.', 'Warning', 'OK', 'Warning') | Out-Null
        return
    }

    $systemType = $cmbSystemType.SelectedItem
    $sccmTier = $cmbSCCMTier.SelectedItem
    $sccmTierParam = if ($sccmTier -and $sccmTier -ne '(Skip - IGT)') { $sccmTier } else { '' }

    # Build summary for confirmation
    $steps = @()
    if ($chkPrepBGInfo.Checked) { $steps += 'BGInfo Removal' }
    if ($chkPrepCribl.Checked) { $steps += 'CRIBL Config' }
    if ($chkPrepSCCM.Checked -and $sccmTierParam) { $steps += "SCCM ($sccmTierParam)" }
    if ($chkPrepAdmin.Checked) { $steps += 'Admin Passwords' }
    if ($chkPrepRapid7.Checked) { $steps += 'Rapid7 + ScanAssist' }
    if ($chkPrepCrowdStrike.Checked) { $steps += 'CrowdStrike' }
    if ($chkPrepGpUpdate.Checked) { $steps += 'GPUpdate + Cleanup' }
    if ($chkPrepIISCrypto.Checked) { $steps += 'IISCrypto' }

    if ($steps.Count -eq 0) {
        $lblPrepStatus.Text = 'No steps selected. Check at least one step.'
        $lblPrepStatus.ForeColor = $colorWarning
        return
    }

    $confirmMsg = "Run System Prep with the following steps?`n`n"
    $confirmMsg += "System Type: $systemType`n"
    if ($sccmTierParam) { $confirmMsg += "SCCM Tier: $sccmTierParam`n" }
    $confirmMsg += "`nSteps:`n"
    $confirmMsg += ($steps | ForEach-Object { "  - $_" }) -join "`n"
    $confirmMsg += "`n`nThis may take 15+ minutes. Continue?"

    $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, 'Confirm System Prep', 'YesNo', 'Question')
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $lblPrepStatus.Text = 'Running System Prep... please wait.'
    $lblPrepStatus.ForeColor = $colorAccent2
    $txtPrepResults.Text = 'Running...'
    [System.Windows.Forms.Application]::DoEvents()

    $result = Invoke-PrepSystem `
        -SystemType $systemType `
        -SCCMTier $sccmTierParam `
        -RunBGInfo $chkPrepBGInfo.Checked `
        -RunCribl $chkPrepCribl.Checked `
        -RunSCCM $chkPrepSCCM.Checked `
        -RunAdminAccounts $chkPrepAdmin.Checked `
        -RunRapid7 $chkPrepRapid7.Checked `
        -RunCrowdStrike $chkPrepCrowdStrike.Checked `
        -RunGpUpdate $chkPrepGpUpdate.Checked `
        -RunIISCrypto $chkPrepIISCrypto.Checked

    # Build results display
    $resultsText = "System Prep Results:`r`n"
    $resultsText += "System Type: $systemType`r`n"
    $resultsText += "========================================`r`n"
    foreach ($sr in $result.StepResults) {
        $status = if ($sr.Result.Success) { 'OK' } else { 'FAIL' }
        $resultsText += "[$status] $($sr.Step): $($sr.Result.Message)`r`n"
    }
    $resultsText += "========================================`r`n"
    $resultsText += "Total: $($result.SuccessCount) succeeded, $($result.FailCount) failed`r`n"
    if ($chkPrepIISCrypto.Checked) {
        $resultsText += "`nNOTE: IISCrypto requires a reboot to take effect."
    }

    $txtPrepResults.Text = $resultsText

    if ($result.Success) {
        $lblPrepStatus.Text = $result.Message
        $lblPrepStatus.ForeColor = $colorSuccess
    } else {
        $lblPrepStatus.Text = $result.Message
        $lblPrepStatus.ForeColor = $colorWarning
    }

    # Offer reboot if IISCrypto was run
    if ($chkPrepIISCrypto.Checked) {
        $reboot = [System.Windows.Forms.MessageBox]::Show(
            "System Prep complete. IISCrypto was applied and requires a reboot.`n`nWould you like to reboot now?",
            'Reboot Required',
            'YesNo',
            'Question'
        )
        if ($reboot -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log -Message 'Rebooting after System Prep (IISCrypto).' -Level 'INFO'
            Restart-Computer -Force
        }
    }
})

# ============================================================================
# SECTION 7: Status Bar
# ============================================================================
$statusStrip = New-Object System.Windows.Forms.StatusStrip -Property @{
    BackColor = $colorPanel
}
$mainForm.Controls.Add($statusStrip)

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel -Property @{
    Text = 'Ready.'
    ForeColor = $colorTextDim
}
$statusStrip.Items.Add($statusLabel)

# Status update timer
$statusTimer = New-Object System.Windows.Forms.Timer -Property @{
    Interval = 2000
}
$statusTimer.Add_Tick({
    $domainStatus = if ($env:USERDOMAIN -eq $env:COMPUTERNAME) { 'Workgroup' } else { $env:USERDOMAIN }
    $statusLabel.Text = "Computer: $env:COMPUTERNAME | Domain: $domainStatus | User: $env:USERNAME | Log: $script:LogPath"
})
$statusTimer.Start()

# ============================================================================
# REGION: Form Show
# ============================================================================

# Initialize log
Write-Log -Message '==========================================' -Level 'INFO'
Write-Log -Message 'Domain Join GUI Tool Started' -Level 'INFO'
Write-Log -Message "Computer: $env:COMPUTERNAME" -Level 'INFO'
Write-Log -Message "User: $env:USERNAME" -Level 'INFO'
Write-Log -Message "OS: $((Get-CimInstance Win32_OperatingSystem).Caption)" -Level 'INFO'
Write-Log -Message '==========================================' -Level 'INFO'

# Show the form
[void]$mainForm.ShowDialog()
