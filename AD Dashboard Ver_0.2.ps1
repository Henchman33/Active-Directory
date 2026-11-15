#requires -Version 5.1
<#
.SYNOPSIS
  Active Directory Security & Health Analyzer (ADSHA) v0.2 - read-only.
.DESCRIPTION
  Forest/domain discovery, user and computer inventory, privileged group review,
  DC health, DNS SRV checks, AD replication, SYSVOL checks, and AD CS assessment.
  Supports a separately hosted issuing CA (including Windows Server 2022) and
  legacy member-server inventory back to Windows Server 2003.
  Produces JSON, CSV, HTML, historical JSON, and a text log. No SQL is used.
.NOTES
  Windows PowerShell 5.1. Run from an AD management host or DC with the
  ActiveDirectory module. Remote CA/DC certificate checks use PowerShell remoting.
  The script is read-only. Scores are provisional risk indicators, not audit
  certifications. Failed/unavailable collection is reported separately.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = 'C:\ADSHA',
    [string]$ConfigPath = '',
    [switch]$SkipRemoteChecks,
    [switch]$NoOpenReport
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:Version = '0.2'
$script:RunId = [guid]::NewGuid().ToString()
$script:Started = Get-Date
$script:Now = Get-Date
$script:Findings = New-Object System.Collections.Generic.List[object]
$script:Errors = New-Object System.Collections.Generic.List[object]
$script:Inventory = New-Object System.Collections.Generic.List[object]
$script:DCInventory = New-Object System.Collections.Generic.List[object]
$script:CategoryScores = @{}
$script:LogPath = $null

function Write-Log {
    param([string]$Message,[ValidateSet('INFO','WARN','ERROR')][string]$Level='INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    if ($script:LogPath) { Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8 }
    if ($Level -eq 'ERROR') { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq 'WARN') { Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}

function Add-Finding {
    param(
        [string]$Category,
        [ValidateSet('Critical','High','Medium','Low','Informational')][string]$Severity,
        [string]$Title,[string]$Object,[string]$Details,[string]$Recommendation
    )
    $key = '{0}|{1}|{2}|{3}' -f $Category,$Title,$Object,$Severity
    if (@($script:Findings | Where-Object { $_.FindingKey -eq $key }).Count -gt 0) { return }
    $script:Findings.Add([pscustomobject]@{
        FindingKey=$key; Category=$Category; Severity=$Severity; Title=$Title
        Object=$Object; Details=$Details; Recommendation=$Recommendation
        DetectedAt=(Get-Date).ToString('s')
    }) | Out-Null
}

function Add-AssessmentError {
    param([string]$Category,[string]$Target,[string]$Message,[string]$Check='Collection')
    $script:Errors.Add([pscustomobject]@{
        Category=$Category; Target=$Target; Check=$Check; Message=$Message
        DetectedAt=(Get-Date).ToString('s')
    }) | Out-Null
    Write-Log "$Category / $Target / $Check : $Message" 'WARN'
}

function ConvertTo-HtmlSafe {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Test-TcpPort {
    param([string]$ComputerName,[int]$Port,[int]$TimeoutMs=1800)
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $client.BeginConnect($ComputerName,$Port,$null,$null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs,$false)) { return $false }
        $client.EndConnect($iar)
        return $client.Connected
    } catch { return $false }
    finally { $client.Close() }
}

function Get-OSClassification {
    param([string]$OperatingSystem,[string]$Version)
    $os = [string]$OperatingSystem
    if ([string]::IsNullOrWhiteSpace($os)) { return 'Unknown' }
    if ($os -match 'Domain Controller') { return 'Domain Controller' }
    if ($os -match 'Windows Server') { return 'Windows Server' }
    if ($os -match 'Windows 10|Windows 11|Windows 8|Windows 7|Windows Vista|Windows XP') { return 'Workstation' }
    return 'Other'
}

function Get-LegacyOSStatus {
    param([string]$OperatingSystem,[string]$Version)
    $os = [string]$OperatingSystem
    if ($os -match '2003|2008|2012|Windows 7|Windows 8(\s|$)|XP|Vista') { return 'Unsupported-Legacy' }
    if ($os -match '2016|2019|2022|2025') { return 'Review-Lifecycle' }
    if ([string]::IsNullOrWhiteSpace($os)) { return 'Unknown-OS' }
    return 'Review-Lifecycle'
}

function Get-Status {
    param([double]$Score)
    if ($Score -ge 95) { return 'Excellent' }
    if ($Score -ge 90) { return 'Healthy' }
    if ($Score -ge 80) { return 'Acceptable' }
    if ($Score -ge 65) { return 'Needs Attention' }
    if ($Score -ge 50) { return 'Poor' }
    return 'Critical'
}

function Get-CategoryScore {
    param([string]$Category,[int]$Denominator)
    if ($Denominator -le 0) { return $null }
    $points = @{Critical=30;High=15;Medium=6;Low=2;Informational=0}
    $risk = 0
    foreach ($f in @($script:Findings | Where-Object Category -eq $Category)) {
        $risk += $points[$f.Severity]
    }
    # Normalize by assessed population. Multiple distinct controls on one object
    # can exceed the denominator; the penalty is capped at 100.
    $penalty = [math]::Min(100,($risk / [double]$Denominator) * 5)
    return [math]::Round((100-$penalty),1)
}

function Get-WeightedScore {
    param([object[]]$Rows,[hashtable]$Weights)
    $numerator=0.0; $denominator=0.0
    foreach ($row in $Rows) {
        if ($null -ne $row.Score) {
            $w=[double]$Weights[$row.Category]
            $numerator += ([double]$row.Score*$w)
            $denominator += $w
        }
    }
    if ($denominator -eq 0) { return $null }
    return [math]::Round(($numerator/$denominator),1)
}

function Read-JsonConfig {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        $default = [ordered]@{
            InactiveUserDays=90
            InactiveComputerDays=90
            OldPasswordDays=365
            ReplicationWarningHours=0.5
            ReplicationHighHours=2
            ReplicationCriticalHours=24
            CertificateWarningDays=30
            IssuingCAComputerName=''
            IssuingCAName=''
            ApprovedAdmins=@()
            ApprovedExceptions=@()
            LegacyMigrationTracking=@()
            EnableRemoteDCChecks=$true
            EnableRemoteCAChecks=$true
            Notes='Populate IssuingCAComputerName if AD discovery does not identify the issuing CA host. ApprovedAdmins entries should be DOMAIN\\samAccountName.'
        }
        $default | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
        Write-Log "Created starter configuration: $Path" 'WARN'
    }
    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) }
    catch { throw "Configuration JSON is invalid: $Path. $($_.Exception.Message)" }
}

function Test-ApprovedException {
    param([string]$Object,[string]$FindingTitle)
    foreach ($entry in @($script:Config.ApprovedExceptions)) {
        if ($entry.Object -eq $Object -and $entry.FindingTitle -eq $FindingTitle) { return $true }
    }
    return $false
}

function Get-DirectoryRoot {
    param([string]$Server)
    return Get-ADRootDSE -Server $Server -ErrorAction Stop
}

try {
    foreach ($folder in @($OutputPath,(Join-Path $OutputPath 'Configuration'),(Join-Path $OutputPath 'Reports\Current'),(Join-Path $OutputPath 'Reports\Historical'),(Join-Path $OutputPath 'Logs'))) {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    }
    $script:LogPath = Join-Path $OutputPath ('Logs\ADSHA_{0}.log' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $OutputPath 'Configuration\ADSHA_Config.json' }
    $script:Config = Read-JsonConfig -Path $ConfigPath

    Import-Module ActiveDirectory -ErrorAction Stop
    $forest = Get-ADForest -ErrorAction Stop
    $domains = @($forest.Domains | Sort-Object -Unique)
    Write-Log "ADSHA v$script:Version started for forest $($forest.Name)"
    Write-Log "Domains discovered: $($domains -join ', ')"

    # Forest/site/DC discovery.
    $allDCs = New-Object System.Collections.Generic.List[object]
    $sites = New-Object System.Collections.Generic.List[string]
    foreach ($domainName in $domains) {
        try {
            foreach ($dc in @(Get-ADDomainController -Filter * -Server $domainName -ErrorAction Stop)) {
                $allDCs.Add($dc) | Out-Null
                if ($dc.Site -and -not $sites.Contains([string]$dc.Site)) { $sites.Add([string]$dc.Site) | Out-Null }
            }
        } catch { Add-AssessmentError 'Discovery' $domainName $_.Exception.Message }
    }
    if ($allDCs.Count -eq 0) { throw 'No domain controllers were discovered. Check AD module connectivity and credentials.' }

    # 1. USERS
    Write-Log '[1/7] User account assessment'
    $userCount=0
    foreach ($domainName in $domains) {
        try {
            $users=@(Get-ADUser -Filter * -Server $domainName -Properties Enabled,LastLogonDate,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,DoesNotRequirePreAuth,AllowReversiblePasswordEncryption,TrustedForDelegation,AdminCount,ServicePrincipalName,Description,DistinguishedName)
            $userCount += $users.Count
            foreach ($u in $users) {
                $id="$domainName\$($u.SamAccountName)"
                if ($u.Enabled -and $u.LastLogonDate -and $u.LastLogonDate -lt $script:Now.AddDays(-[int]$script:Config.InactiveUserDays) -and -not (Test-ApprovedException $id 'Enabled user inactive')) {
                    Add-Finding Users Medium 'Enabled user inactive' $id "Last logon: $($u.LastLogonDate)" 'Validate ownership and business need; process through the approved account lifecycle.'
                }
                if ($u.Enabled -and -not $u.LastLogonDate) { Add-Finding Users Low 'Enabled user has no recorded last logon' $id 'LastLogonDate is not populated.' 'Validate whether the account is new, unused, or a service identity.' }
                if ($u.Enabled -and $u.PasswordNotRequired) { Add-Finding Users Critical 'Password not required' $id 'PasswordNotRequired is enabled.' 'Remove the setting after validating dependencies and reset credentials.' }
                if ($u.Enabled -and $u.PasswordNeverExpires -and -not (Test-ApprovedException $id 'Password never expires')) { Add-Finding Users Medium 'Password never expires' $id 'PasswordNeverExpires is enabled.' 'Use a managed credential strategy or document an approved exception.' }
                if ($u.Enabled -and $u.PasswordLastSet -and $u.PasswordLastSet -lt $script:Now.AddDays(-[int]$script:Config.OldPasswordDays)) { Add-Finding Users Medium 'Password age exceeds configured threshold' $id "Password last set: $($u.PasswordLastSet)" 'Review account type and rotate credentials under policy.' }
                if ($u.Enabled -and $u.DoesNotRequirePreAuth) { Add-Finding Users High 'Kerberos preauthentication disabled' $id 'Account may be AS-REP roastable.' 'Enable preauthentication unless an approved exception exists.' }
                if ($u.Enabled -and $u.AllowReversiblePasswordEncryption) { Add-Finding Users Critical 'Reversible password encryption enabled' $id 'AllowReversiblePasswordEncryption is enabled.' 'Disable reversible password storage and reset the password.' }
                if ($u.Enabled -and $u.TrustedForDelegation) { Add-Finding Users High 'Unconstrained delegation enabled on user' $id 'TrustedForDelegation is enabled.' 'Review service requirements and replace with constrained delegation where possible.' }
            }
        } catch { Add-AssessmentError Users $domainName $_.Exception.Message }
    }

    # 2. COMPUTER INVENTORY - DCs are separated from member servers/workstations.
    Write-Log '[2/7] Computer and legacy member-server inventory'
    $computerCount=0; $serverCount=0; $legacyCount=0
    foreach ($domainName in $domains) {
        try {
            $computers=@(Get-ADComputer -Filter * -Server $domainName -Properties Enabled,LastLogonDate,OperatingSystem,OperatingSystemVersion,PasswordLastSet,TrustedForDelegation,PrimaryGroupID,IPv4Address,Description,DistinguishedName)
            $computerCount += $computers.Count
            $dcNames=@($allDCs | Where-Object { $_.Domain -eq $domainName } | ForEach-Object { $_.Name.ToLowerInvariant() })
            foreach ($c in $computers) {
                $isDC = $dcNames -contains $c.Name.ToLowerInvariant()
                $class=Get-OSClassification $c.OperatingSystem $c.OperatingSystemVersion
                if ($isDC) { $class='Domain Controller' }
                $legacy=if ($class -eq 'Windows Server' -and -not $isDC) { Get-LegacyOSStatus $c.OperatingSystem $c.OperatingSystemVersion } else { 'NotApplicable' }
                if ($class -eq 'Windows Server' -and -not $isDC) { $serverCount++ }
                if ($legacy -eq 'Unsupported-Legacy') { $legacyCount++ }
                $item=[pscustomobject]@{
                    Domain=$domainName; Name=$c.Name; DNSHostName=$c.DNSHostName
                    Enabled=$c.Enabled; ObjectClass=$class; OperatingSystem=$c.OperatingSystem
                    OperatingSystemVersion=$c.OperatingSystemVersion; LastLogonDate=$c.LastLogonDate
                    PasswordLastSet=$c.PasswordLastSet; IPv4Address=$c.IPv4Address
                    DistinguishedName=$c.DistinguishedName; IsDomainController=$isDC
                    LegacyStatus=$legacy; MigrationStatus='Untracked'
                }
                foreach ($tracking in @($script:Config.LegacyMigrationTracking)) {
                    if ($tracking.ComputerName -eq $c.Name -and $tracking.Domain -eq $domainName) { $item.MigrationStatus=$tracking.Status }
                }
                $script:Inventory.Add($item) | Out-Null
                $id="$domainName\$($c.Name)"
                if ($c.Enabled -and $c.LastLogonDate -and $c.LastLogonDate -lt $script:Now.AddDays(-[int]$script:Config.InactiveComputerDays)) { Add-Finding Computers Medium 'Enabled computer inactive' $id "Last logon: $($c.LastLogonDate); OS: $($c.OperatingSystem)" 'Validate ownership and retire or refresh stale systems through change control.' }
                if ($c.Enabled -and -not $c.LastLogonDate) { Add-Finding Computers Low 'Enabled computer has no recorded last logon' $id 'LastLogonDate is not populated.' 'Validate whether this is a new or orphaned computer object.' }
                if ($c.Enabled -and $c.TrustedForDelegation) { Add-Finding Computers High 'Computer has unconstrained delegation' $id 'TrustedForDelegation is enabled.' 'Review the requirement and migrate to constrained or resource-based constrained delegation.' }
                if ($legacy -eq 'Unsupported-Legacy') { Add-Finding Computers Critical 'Unsupported legacy member server OS' $id "Operating system: $($c.OperatingSystem); version: $($c.OperatingSystemVersion)" 'Prioritize migration, isolation, or retirement. This is a member-server finding, not a DC outage.' }
                elseif ($legacy -eq 'Review-Lifecycle' -and $class -eq 'Windows Server') { Add-Finding Computers Low 'Server OS lifecycle requires validation' $id "Operating system: $($c.OperatingSystem); version: $($c.OperatingSystemVersion)" 'Validate exact edition/build and Microsoft lifecycle/support status against organizational policy.' }
            }
        } catch { Add-AssessmentError Computers $domainName $_.Exception.Message }
    }

    # 3. PRIVILEGED GROUPS - no generic member-count thresholds.
    Write-Log '[3/7] Privileged group review'
    $privilegedNames=@('Domain Admins','Enterprise Admins','Schema Admins','Administrators','Account Operators','Backup Operators','Server Operators','Print Operators','DnsAdmins','Group Policy Creator Owners','Key Admins','Enterprise Key Admins')
    $privilegedMemberCount=0
    foreach ($domainName in $domains) {
        foreach ($groupName in $privilegedNames) {
            try {
                $group=Get-ADGroup -Identity $groupName -Server $domainName -ErrorAction Stop
                $members=@(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -Server $domainName -ErrorAction Stop)
                foreach ($m in $members) {
                    if ($m.objectClass -eq 'user') {
                        $u=Get-ADUser -Identity $m.DistinguishedName -Server $domainName -Properties Enabled,LastLogonDate,PasswordNeverExpires,DoesNotRequirePreAuth,SIDHistory
                        $privilegedMemberCount++
                        $id="$domainName\$($u.SamAccountName)"
                        if (-not $u.Enabled) { Add-Finding 'Privileged Groups' Medium 'Disabled account remains in privileged group' $id "Group: $groupName" 'Review and remove unnecessary privileged membership through change control.' }
                        if ($u.Enabled -and $u.LastLogonDate -and $u.LastLogonDate -lt $script:Now.AddDays(-30)) { Add-Finding 'Privileged Groups' High 'Inactive privileged account' $id "Group: $groupName; last logon: $($u.LastLogonDate)" 'Validate ownership and remove standing privilege when no longer needed.' }
                        if ($u.Enabled -and $u.PasswordNeverExpires) { Add-Finding 'Privileged Groups' High 'Privileged account password never expires' $id "Group: $groupName" 'Apply privileged credential policy or record an approved exception.' }
                        if ($u.Enabled -and $u.DoesNotRequirePreAuth) { Add-Finding 'Privileged Groups' Critical 'Privileged account lacks Kerberos preauthentication' $id "Group: $groupName" 'Enable preauthentication and investigate potential credential exposure.' }
                        if ($u.SIDHistory) { Add-Finding 'Privileged Groups' High 'Privileged account has SIDHistory' $id "Group: $groupName; SIDHistory: $($u.SIDHistory -join ',')" 'Validate migration provenance and remove unneeded SIDHistory under an approved plan.' }
                        if (@($script:Config.ApprovedAdmins).Count -gt 0 -and $groupName -in @('Domain Admins','Enterprise Admins','Schema Admins','Administrators')) {
                            if ($script:Config.ApprovedAdmins -notcontains $id) { Add-Finding 'Privileged Groups' Critical 'Privileged membership is not in approved administrator baseline' $id "Group: $groupName" 'Validate identity and authorization with the AD security owner; update the baseline only after approval.' }
                        }
                    }
                }
            } catch {
                if ($_.Exception.Message -notmatch 'cannot find|does not exist|cannot be found') { Add-AssessmentError 'Privileged Groups' "$domainName\$groupName" $_.Exception.Message }
            }
        }
    }

    # 4. DOMAIN CONTROLLERS - ICMP failure is informational only; validate service ports/shares.
    Write-Log '[4/7] Domain controller health'
    foreach ($dc in $allDCs) {
        $dcName=[string]$dc.HostName
        $dcRow=[ordered]@{Domain=$dc.Domain;HostName=$dcName;Site=$dc.Site;OS=$dc.OperatingSystem;IsGlobalCatalog=$dc.IsGlobalCatalog;ReachablePorts=@();FailedPorts=@();Sysvol=$null;Netlogon=$null;Services=@()}
        try {
            foreach ($port in @(53,88,135,389,445,3268)) {
                if (Test-TcpPort $dcName $port) { $dcRow.ReachablePorts += $port }
                else {
                    $dcRow.FailedPorts += $port
                    $sev=if ($port -in @(389,445)) {'High'} else {'Medium'}
                    Add-Finding 'Domain Controllers' $sev 'Expected DC service port did not respond' $dcName "TCP $port did not respond. A firewall or role design may explain this." 'Validate service role, firewall policy, and required connectivity; confirm applicability before remediation.'
                }
            }
            $dcRow.Sysvol=Test-Path "\\$dcName\SYSVOL"
            $dcRow.Netlogon=Test-Path "\\$dcName\NETLOGON"
            if (-not $dcRow.Sysvol) { Add-Finding 'Domain Controllers' Critical 'SYSVOL share unavailable' $dcName "\\$dcName\SYSVOL could not be accessed." 'Investigate DFSR/SYSVOL state, Netlogon, DNS, and replication.' }
            if (-not $dcRow.Netlogon) { Add-Finding 'Domain Controllers' High 'NETLOGON share unavailable' $dcName "\\$dcName\NETLOGON could not be accessed." 'Investigate Netlogon service and SYSVOL readiness.' }
            if (-not $SkipRemoteChecks -and $script:Config.EnableRemoteDCChecks) {
                try {
                    $services=Invoke-Command -ComputerName $dcName -ScriptBlock {
                        Get-Service -Name NTDS,DNS,KDC,Netlogon,DFSR,W32Time -ErrorAction SilentlyContinue |
                        Select-Object Name,Status,StartType
                    } -ErrorAction Stop
                    foreach ($svc in @($services)) {
                        if ($svc.Status -ne 'Running' -and $svc.Name -in @('NTDS','KDC','Netlogon','DFSR')) {
                            Add-Finding 'Domain Controllers' Critical 'Required DC service is not running' $dcName "$($svc.Name) status: $($svc.Status)" 'Validate service dependencies and DC health before attempting service recovery.'
                        }
                    }
                    $dcRow.Services=@($services)
                } catch { Add-AssessmentError 'Domain Controllers' $dcName $_.Exception.Message 'Remote service query' }
            }
        } catch { Add-AssessmentError 'Domain Controllers' $dcName $_.Exception.Message }
        $script:DCInventory.Add([pscustomobject]$dcRow) | Out-Null
    }

    # 5. DNS - SRV and hostname resolution. A DNS query failure is not inferred from ping.
    Write-Log '[5/7] DNS and DC locator records'
    $dnsChecks=0
    foreach ($dc in $allDCs) {
        $domainName=[string]$dc.Domain
        foreach ($query in @("_ldap._tcp.dc._msdcs.$domainName","_kerberos._tcp.$domainName","_ldap._tcp.gc._msdcs.$($forest.Name)")) {
            try {
                $answer=@(Resolve-DnsName -Name $query -Type SRV -Server $dc.HostName -ErrorAction Stop)
                $dnsChecks++
                if ($answer.Count -eq 0) { Add-Finding DNS High 'AD DNS SRV lookup returned no records' $dc.HostName "Query: $query" 'Validate AD-integrated zones and SRV registration.' }
            } catch { Add-Finding DNS High 'AD DNS SRV lookup failed' $dc.HostName "Query: $query; $($_.Exception.Message)" 'Validate DNS service, zone replication, SRV records, and DNS client configuration.' }
        }
        try { Resolve-DnsName -Name $dc.HostName -Server $dc.HostName -ErrorAction Stop | Out-Null; $dnsChecks++ }
        catch { Add-Finding DNS High 'DC hostname does not resolve using its configured DNS host' $dc.HostName $_.Exception.Message 'Validate the host A/AAAA record and DNS zone health.' }
    }

    # 6. REPLICATION and SYSVOL DFSR backlog/health.
    Write-Log '[6/7] AD replication and SYSVOL'
    $replicationChecks=0
    foreach ($domainName in $domains) {
        try {
            $failures=@(Get-ADReplicationFailure -Target $domainName -Scope Domain -ErrorAction Stop)
            foreach ($failure in $failures) {
                $ageHours=if ($failure.FirstFailureTime) { ($script:Now-$failure.FirstFailureTime).TotalHours } else { 0 }
                $sev=if ($ageHours -ge [double]$script:Config.ReplicationCriticalHours) {'Critical'} elseif ($ageHours -ge [double]$script:Config.ReplicationHighHours) {'High'} else {'Medium'}
                Add-Finding Replication $sev 'Active Directory replication failure' "$($failure.Server) <- $($failure.Partner)" "First failure: $($failure.FirstFailureTime); consecutive failures: $($failure.ConsecutiveFailureCount); error: $($failure.LastError)" 'Investigate DNS, RPC/firewall, time, secure channel, site links, and the reported replication error.'
                $replicationChecks++
            }
            $metadata=@(Get-ADReplicationPartnerMetadata -Target $domainName -Scope Domain -ErrorAction Stop)
            foreach ($m in $metadata) {
                $replicationChecks++
                if ($m.LastReplicationSuccess) {
                    $age=($script:Now-$m.LastReplicationSuccess).TotalHours
                    if ($age -ge [double]$script:Config.ReplicationCriticalHours) { Add-Finding Replication Critical 'Replication partner has not succeeded within critical threshold' "$($m.Server) <- $($m.Partner)" "Last success: $($m.LastReplicationSuccess); partition: $($m.Partition); age hours: $([math]::Round($age,1))" 'Investigate replication topology, connectivity, DNS, time, and site-link schedules.' }
                    elseif ($age -ge [double]$script:Config.ReplicationHighHours) { Add-Finding Replication High 'Replication partner is delayed' "$($m.Server) <- $($m.Partner)" "Last success: $($m.LastReplicationSuccess); partition: $($m.Partition); age hours: $([math]::Round($age,1))" 'Check site-link schedules and partner health.' }
                    elseif ($age -ge [double]$script:Config.ReplicationWarningHours) { Add-Finding Replication Medium 'Replication partner exceeds warning age' "$($m.Server) <- $($m.Partner)" "Last success: $($m.LastReplicationSuccess); age hours: $([math]::Round($age,1))" 'Confirm expected replication schedule and investigate unusual latency.' }
                }
            }
        } catch { Add-AssessmentError Replication $domainName $_.Exception.Message }
    }
    try {
        $dfsrState=@(Get-ADObject -LDAPFilter '(objectClass=msDFSR-Subscription)' -SearchBase (Get-ADRootDSE).defaultNamingContext -Properties msDFSR-Enabled,msDFSR-Options,distinguishedName -ErrorAction Stop)
        if ($dfsrState.Count -eq 0) { Add-AssessmentError Replication $forest.Name 'No DFSR SYSVOL subscription objects were returned; confirm SYSVOL migration/state.' 'DFSR directory check' }
    } catch { Add-AssessmentError Replication $forest.Name $_.Exception.Message 'DFSR directory check' }

    # 7. PKI - AD metadata and separately hosted CA remote query.
    Write-Log '[7/7] PKI and certificate assessment'
    $pkiDenominator=0
    $caHosts=New-Object System.Collections.Generic.List[string]
    try {
        $configNC=(Get-ADRootDSE).configurationNamingContext
        $templateBase="CN=Certificate Templates,CN=Public Key Services,CN=Services,$configNC"
        $caBase="CN=Enrollment Services,CN=Public Key Services,CN=Services,$configNC"
        $templates=@(Get-ADObject -SearchBase $templateBase -LDAPFilter '(objectClass=pKICertificateTemplate)' -Properties displayName,pKIExtendedKeyUsage,pKIExpirationPeriod,pKIOverlapPeriod,nTSecurityDescriptor -ErrorAction Stop)
        $published=@(Get-ADObject -SearchBase $caBase -LDAPFilter '(objectClass=pKIEnrollmentService)' -Properties dNSHostName,certificateTemplates,cACertificate -ErrorAction Stop)
        $pkiDenominator += [math]::Max(1,$templates.Count)
        foreach ($ca in $published) {
            if ($ca.dNSHostName -and -not $caHosts.Contains([string]$ca.dNSHostName)) { $caHosts.Add([string]$ca.dNSHostName) | Out-Null }
            if (-not $ca.dNSHostName) { Add-Finding PKI High 'Enterprise CA has no published DNS hostname' $ca.Name 'Enrollment service object lacks dNSHostName.' 'Validate AD CS enrollment service publication.' }
            foreach ($templateName in @($ca.certificateTemplates)) {
                if (-not ($templates | Where-Object { $_.Name -eq $templateName })) { Add-Finding PKI Medium 'CA publishes a template not found in directory template query' $ca.Name "Template: $templateName" 'Validate template publication and directory replication.' }
            }
        }
        if ($templates.Count -eq 0) { Add-AssessmentError PKI $forest.Name 'No certificate templates were returned from the Configuration partition.' 'Template discovery' }
    } catch { Add-AssessmentError PKI $forest.Name $_.Exception.Message 'AD CS metadata' }

    $configuredCA=[string]$script:Config.IssuingCAComputerName
    if (-not [string]::IsNullOrWhiteSpace($configuredCA)) {
        if (-not $caHosts.Contains($configuredCA)) { $caHosts.Add($configuredCA) | Out-Null }
    }
    foreach ($caHost in $caHosts) {
        $pkiDenominator++
        if ($SkipRemoteChecks -or -not $script:Config.EnableRemoteCAChecks) {
            Add-AssessmentError PKI $caHost 'Remote CA checks were skipped by configuration or command-line switch.' 'CA remote assessment'
            continue
        }
        try {
            $caData=Invoke-Command -ComputerName $caHost -ScriptBlock {
                $svc=Get-Service -Name CertSvc -ErrorAction Stop
                $certUtilOutput=& certutil.exe -getreg CA\CAName 2>&1
                $caCerts=@(Get-ChildItem Cert:\LocalMachine\My | Where-Object {
                    $bc = $_.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.19' }
                    ($bc -and $bc.Format($false) -match 'CA\s*=\s*TRUE')
                } | Select-Object Subject,Issuer,Thumbprint,NotAfter,NotBefore,SignatureAlgorithm)
                [pscustomobject]@{
                    ServiceStatus=$svc.Status.ToString()
                    CAName=($certUtilOutput -join ' ')
                    CACertificates=$caCerts
                    ComputerName=$env:COMPUTERNAME
                }
            } -ErrorAction Stop
            if ($caData.ServiceStatus -ne 'Running') { Add-Finding PKI Critical 'Certificate Services is not running' $caHost "CertSvc status: $($caData.ServiceStatus)" 'Validate CA service health and review CA event logs.' }
            foreach ($cert in @($caData.CACertificates)) {
                if ($cert.NotAfter -lt $script:Now) { Add-Finding PKI Critical 'Issuing CA signing certificate expired' $caHost "$($cert.Subject); expired $($cert.NotAfter); thumbprint $($cert.Thumbprint)" 'Follow the CA certificate renewal and hierarchy maintenance procedure.' }
                elseif ($cert.NotAfter -lt $script:Now.AddDays([int]$script:Config.CertificateWarningDays)) { Add-Finding PKI High 'Issuing CA signing certificate nearing expiration' $caHost "$($cert.Subject); expires $($cert.NotAfter); thumbprint $($cert.Thumbprint)" 'Plan CA certificate renewal and validate dependent certificate chains.' }
                if ($cert.SignatureAlgorithm -match 'sha1|md5') { Add-Finding PKI High 'Weak CA certificate signature algorithm' $caHost "$($cert.Subject); algorithm: $($cert.SignatureAlgorithm)" 'Review cryptographic policy and plan supported CA certificate renewal.' }
            }
        } catch { Add-AssessmentError PKI $caHost $_.Exception.Message 'Remote issuing CA assessment (WinRM)' }
    }
    if ($caHosts.Count -eq 0) { Add-AssessmentError PKI 'Issuing CA' 'No CA host was discovered from AD Enrollment Services and no IssuingCAComputerName was configured.' 'CA host discovery' }

    # Query machine certificate stores on DCs and issuing CA host; failure is explicitly incomplete.
    if (-not $SkipRemoteChecks) {
        foreach ($hostName in @($allDCs | ForEach-Object { $_.HostName })) {
            try {
                $certs=Invoke-Command -ComputerName $hostName -ScriptBlock {
                    Get-ChildItem Cert:\LocalMachine\My | Select-Object Subject,Issuer,Thumbprint,NotAfter,NotBefore,EnhancedKeyUsageList
                } -ErrorAction Stop
                foreach ($cert in @($certs)) {
                    if ($cert.NotAfter -lt $script:Now) { Add-Finding PKI Critical 'Expired certificate in DC machine store' $hostName "$($cert.Subject); expired $($cert.NotAfter); thumbprint $($cert.Thumbprint)" 'Identify the certificate purpose and renew/replace it; validate dependent services.' }
                    elseif ($cert.NotAfter -lt $script:Now.AddDays([int]$script:Config.CertificateWarningDays)) { Add-Finding PKI High 'DC machine certificate nearing expiration' $hostName "$($cert.Subject); expires $($cert.NotAfter); thumbprint $($cert.Thumbprint)" 'Renew and validate auto-enrollment, template permissions, and chain/revocation access.' }
                }
            } catch { Add-AssessmentError PKI $hostName $_.Exception.Message 'Remote DC certificate store' }
        }
    }

    # Score reporting: denominators are assessed population counts, not raw finding totals.
    # Category scores remain null when collection yielded no assessable population.
    $weights=[ordered]@{Users=.15;Computers=.15;'Privileged Groups'=.20;'Domain Controllers'=.15;DNS=.10;Replication=.10;PKI=.15}
    $denominators=@{
        Users=$userCount
        Computers=$computerCount
        'Privileged Groups'=[math]::Max(1,$privilegedMemberCount)
        'Domain Controllers'=$allDCs.Count
        DNS=[math]::Max(0,$dnsChecks)
        Replication=[math]::Max(0,$replicationChecks)
        PKI=$pkiDenominator
    }
    $categoryRows=New-Object System.Collections.Generic.List[object]
    foreach ($category in $weights.Keys) {
        $denom=[int]$denominators[$category]
        $score=Get-CategoryScore $category $denom
        $script:CategoryScores[$category]=$score
        $catErrors=@($script:Errors | Where-Object Category -eq $category).Count
        $status=if ($null -eq $score) {'Not Assessed'} else {Get-Status $score}
        $categoryRows.Add([pscustomobject]@{
            Category=$category;Score=$score;WeightPercent=[math]::Round($weights[$category]*100,0)
            Status=$status;AssessedPopulation=$denom
            FindingCount=@($script:Findings | Where-Object Category -eq $category).Count
            CollectionErrors=$catErrors
        }) | Out-Null
    }
    $overall=Get-WeightedScore @($categoryRows) $weights
    $healthWeights=@{ 'Domain Controllers'=.35;DNS=.25;Replication=.25;Computers=.15 }
    $securityWeights=@{ Users=.20;Computers=.15;'Privileged Groups'=.30;'Domain Controllers'=.10;DNS=.05;PKI=.20 }
    $healthRows=@($categoryRows | Where-Object Category -in @('Domain Controllers','DNS','Replication','Computers'))
    $securityRows=@($categoryRows | Where-Object Category -in @('Users','Computers','Privileged Groups','Domain Controllers','DNS','PKI'))
    $health=Get-WeightedScore $healthRows $healthWeights
    $security=Get-WeightedScore $securityRows $securityWeights

    $capReason=$null
    if (@($script:Findings | Where-Object { $_.Category -eq 'Privileged Groups' -and $_.Severity -eq 'Critical' }).Count -gt 0) {
        $overall=[math]::Min([double]$overall,49); $security=[math]::Min([double]$security,49); $capReason='Critical privileged access baseline finding'
    }
    if (@($script:Findings | Where-Object { $_.Category -eq 'Replication' -and $_.Severity -eq 'Critical' }).Count -ge 2) {
        $overall=[math]::Min([double]$overall,49); $health=[math]::Min([double]$health,49); $capReason='Multiple critical replication findings'
    }
    $summary=[pscustomobject]@{
        Version=$script:Version;RunId=$script:RunId;Started=$script:Started;Completed=Get-Date
        Forest=$forest.Name;RootDomain=$forest.RootDomain;Domains=$domains
        Sites=@($sites);DomainControllerCount=$allDCs.Count
        UserCount=$userCount;ComputerCount=$computerCount;MemberServerCount=$serverCount
        UnsupportedLegacyMemberServerCount=$legacyCount
        ADHealthScore=$health;ADHealthStatus=if($null -ne $health){Get-Status $health}else{'Not Assessed'}
        ADSecurityScore=$security;ADSecurityStatus=if($null -ne $security){Get-Status $security}else{'Not Assessed'}
        OverallScore=$overall;OverallStatus=if($null -ne $overall){Get-Status $overall}else{'Not Assessed'}
        ScoreCapReason=$capReason
        CriticalFindings=@($script:Findings | Where-Object Severity -eq 'Critical').Count
        HighFindings=@($script:Findings | Where-Object Severity -eq 'High').Count
        MediumFindings=@($script:Findings | Where-Object Severity -eq 'Medium').Count
        LowFindings=@($script:Findings | Where-Object Severity -eq 'Low').Count
        AssessmentErrors=$script:Errors.Count
    }

    # Output paths and historical snapshot.
    $stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
    $current=Join-Path $OutputPath 'Reports\Current'
    $history=Join-Path (Join-Path $OutputPath 'Reports\Historical') (Get-Date -Format 'yyyy-MM-dd')
    if (-not (Test-Path $history)) { New-Item -Path $history -ItemType Directory -Force | Out-Null }
    $jsonPath=Join-Path $current 'Assessment.json'
    $csvPath=Join-Path $current 'Findings.csv'
    $htmlPath=Join-Path $current 'Dashboard.html'
    $inventoryPath=Join-Path $current 'ComputerInventory.csv'
    $legacyPath=Join-Path $current 'LegacyServerTracking.csv'
    $report=[pscustomobject]@{
        Summary=$summary;Categories=@($categoryRows);Findings=@($script:Findings)
        AssessmentErrors=@($script:Errors);DomainControllers=@($script:DCInventory)
        ComputerInventory=@($script:Inventory);ConfigurationPath=$ConfigPath
    }
    $json=$report | ConvertTo-Json -Depth 12
    Set-Content -LiteralPath $jsonPath -Value $json -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $history "Assessment_$stamp.json") -Value $json -Encoding UTF8
    @($script:Findings) | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    @($script:Inventory) | Export-Csv -LiteralPath $inventoryPath -NoTypeInformation -Encoding UTF8
    @($script:Inventory | Where-Object { $_.LegacyStatus -eq 'Unsupported-Legacy' }) | Export-Csv -LiteralPath $legacyPath -NoTypeInformation -Encoding UTF8

    $categoryHtml=(@($categoryRows) | ForEach-Object {
        $scoreText=if($null -eq $_.Score){'N/A'}else{$_.Score}
        $class=if($null -eq $_.Score){'muted'}elseif($_.Score -ge 90){'good'}elseif($_.Score -ge 80){'warn'}else{'bad'}
        "<tr><td>$(ConvertTo-HtmlSafe $_.Category)</td><td>$scoreText</td><td>$($_.WeightPercent)%</td><td>$($_.AssessedPopulation)</td><td>$($_.FindingCount)</td><td>$($_.CollectionErrors)</td><td class='$class'>$(ConvertTo-HtmlSafe $_.Status)</td></tr>"
    }) -join "`n"
    $findingHtml=(@($script:Findings | Sort-Object @{Expression={switch($_.Severity){'Critical'{0};'High'{1};'Medium'{2};'Low'{3};default{4}}}},Category) | ForEach-Object {
        "<tr><td>$(ConvertTo-HtmlSafe $_.Severity)</td><td>$(ConvertTo-HtmlSafe $_.Category)</td><td>$(ConvertTo-HtmlSafe $_.Title)</td><td>$(ConvertTo-HtmlSafe $_.Object)</td><td>$(ConvertTo-HtmlSafe $_.Details)</td><td>$(ConvertTo-HtmlSafe $_.Recommendation)</td></tr>"
    }) -join "`n"
    $errorHtml=(@($script:Errors) | ForEach-Object {
        "<tr><td>$(ConvertTo-HtmlSafe $_.Category)</td><td>$(ConvertTo-HtmlSafe $_.Target)</td><td>$(ConvertTo-HtmlSafe $_.Check)</td><td>$(ConvertTo-HtmlSafe $_.Message)</td></tr>"
    }) -join "`n"
    $legacyHtml=(@($script:Inventory | Where-Object {$_.LegacyStatus -eq 'Unsupported-Legacy'} | ForEach-Object {
        "<tr><td>$(ConvertTo-HtmlSafe $_.Domain)</td><td>$(ConvertTo-HtmlSafe $_.Name)</td><td>$(ConvertTo-HtmlSafe $_.OperatingSystem)</td><td>$(ConvertTo-HtmlSafe $_.OperatingSystemVersion)</td><td>$(ConvertTo-HtmlSafe $_.MigrationStatus)</td></tr>"
    }) -join "`n"
    )
    $html=@"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ADSHA v$script:Version Dashboard</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#101827;color:#e5e7eb;margin:0;padding:26px}
h1,h2{color:#f8fafc}.sub{color:#9ca3af}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:13px;margin:22px 0}
.card{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:17px}.score{font-size:34px;font-weight:700}
.good{color:#4ade80}.warn{color:#facc15}.bad{color:#f87171}.muted{color:#94a3b8}
table{width:100%;border-collapse:collapse;background:#1e293b;margin:12px 0 26px;font-size:12px}
th,td{padding:9px;border-bottom:1px solid #334155;text-align:left;vertical-align:top;overflow-wrap:anywhere}th{background:#273449;color:#cbd5e1}
section{margin-top:28px}.note{padding:12px;background:#332b18;border-left:4px solid #eab308;color:#fde68a}
</style></head><body>
<h1>Active Directory Security &amp; Health Analyzer</h1>
<div class="sub">Version $script:Version | Forest: $(ConvertTo-HtmlSafe $forest.Name) | Run: $($script:RunId) | Started: $($script:Started.ToString('yyyy-MM-dd HH:mm:ss'))</div>
<div class="grid">
<div class="card"><div class="sub">Overall AD Score</div><div class="score $(if($null -ne $overall -and $overall -ge 80){'good'}elseif($null -ne $overall -and $overall -ge 65){'warn'}else{'bad'})">$(if($null -ne $overall){$overall}else{'N/A'})</div><div>$(if($null -ne $overall){Get-Status $overall}else{'Not Assessed'})</div></div>
<div class="card"><div class="sub">AD Health Score</div><div class="score">$(if($null -ne $health){$health}else{'N/A'})</div><div>$(if($null -ne $health){Get-Status $health}else{'Not Assessed'})</div></div>
<div class="card"><div class="sub">AD Security Score</div><div class="score">$(if($null -ne $security){$security}else{'N/A'})</div><div>$(if($null -ne $security){Get-Status $security}else{'Not Assessed'})</div></div>
<div class="card"><div class="sub">Domain Controllers</div><div class="score">$($allDCs.Count)</div></div>
<div class="card"><div class="sub">Member Servers</div><div class="score">$serverCount</div></div>
<div class="card"><div class="sub">Unsupported Legacy Servers</div><div class="score bad">$legacyCount</div></div>
<div class="card"><div class="sub">Critical Findings</div><div class="score bad">$($summary.CriticalFindings)</div></div>
<div class="card"><div class="sub">Collection Errors</div><div class="score">$($script:Errors.Count)</div></div>
</div>
<div class="note">Read-only assessment. Collection errors mean checks were not completed and must not be interpreted as healthy. Scores are provisional risk indicators and require validation against your approved security baseline.</div>
$(if($capReason){"<p class='bad'><strong>Score cap:</strong> $(ConvertTo-HtmlSafe $capReason)</p>"}else{''})
<section><h2>Category Scores</h2><table><tr><th>Category</th><th>Score</th><th>Weight</th><th>Population</th><th>Findings</th><th>Errors</th><th>Status</th></tr>$categoryHtml</table></section>
<section><h2>Security &amp; Health Findings</h2><table><tr><th>Severity</th><th>Category</th><th>Finding</th><th>Object</th><th>Details</th><th>Recommendation</th></tr>$findingHtml</table></section>
<section><h2>Unsupported Legacy Member Servers</h2><p class="sub">Legacy server inventory is separate from DC health. Migration status is Untracked unless explicitly entered in ADSHA_Config.json.</p><table><tr><th>Domain</th><th>Server</th><th>OS</th><th>Version</th><th>Migration Status</th></tr>$legacyHtml</table></section>
<section><h2>Assessment Errors / Incomplete Checks</h2><table><tr><th>Category</th><th>Target</th><th>Check</th><th>Message</th></tr>$errorHtml</table></section>
</body></html>
"@
    Set-Content -LiteralPath $htmlPath -Value $html -Encoding UTF8
    Write-Log "Completed. Findings=$($script:Findings.Count); Errors=$($script:Errors.Count); Overall=$overall"
    Write-Host "`nADSHA v$script:Version completed." -ForegroundColor Green
    Write-Host "Dashboard: $htmlPath"
    Write-Host "Findings:  $csvPath"
    Write-Host "Inventory: $inventoryPath"
    Write-Host "Legacy:    $legacyPath"
    Write-Host "JSON:      $jsonPath"
    Write-Host "Config:    $ConfigPath"
    Write-Host "Log:       $script:LogPath"
    if (-not $NoOpenReport) { Start-Process -FilePath $htmlPath }
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    throw
}
