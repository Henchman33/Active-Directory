<#
.SYNOPSIS
    GUI tool to ping a list of servers, retrieve hostname, IPv4, domain, OS type and distro,
    and export results to CSV, XLSX, or HTML.
.DESCRIPTION
    Paste a list of hostnames into the text box, optionally provide alternate domain credentials,
    then click "Run" to query each host. Results are shown in a grid and can be exported.
.NOTES
    Requires: ActiveDirectory module, ImportExcel module (for XLSX).
    Supports Windows Server 2003–2025 and Windows 7–11.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Helper: Get OS friendly name (all Server 2003–2025, Desktop 7–11) ─────────
function Get-OSFriendlyName {
    param([string]$Caption, [string]$Version)

    # Try to extract a friendly name directly from the Caption first
    if ($Caption -match 'Windows Server (\d{4}( R2)?)') { return $Matches[0] }
    if ($Caption -match 'Windows (7|8\.1|10|11)') { return $Matches[0] }

    # Normalise the version string to Major.Minor.Build
    $ver = $null
    if ($Version -match '(\d+\.\d+)\.(\d+)') {
        $ver = "$($Matches[1]).$($Matches[2])"
    } elseif ($Version -match '(\d+\.\d+) \((\d+)\)') {
        $ver = "$($Matches[1]).$($Matches[2])"
    } elseif ($Version -match '^(\d+\.\d+\.\d+)') {
        $ver = $Matches[1]
    }

    if (-not $ver) { return $Caption }

    # Determine if this is a Server OS based on Caption
    $isServer = $Caption -match 'Server'

    if ($isServer) {
        $map = @{
            '5.2.3790'   = 'Windows Server 2003'
            '6.0.6001'   = 'Windows Server 2008'
            '6.0.6002'   = 'Windows Server 2008'
            '6.1.7600'   = 'Windows Server 2008 R2'
            '6.1.7601'   = 'Windows Server 2008 R2'
            '6.2.9200'   = 'Windows Server 2012'
            '6.3.9600'   = 'Windows Server 2012 R2'
            '10.0.14393' = 'Windows Server 2016'
            '10.0.17763' = 'Windows Server 2019'
            '10.0.20348' = 'Windows Server 2022'
            '10.0.26100' = 'Windows Server 2025'
        }
    } else {
        $map = @{
            '6.1.7600'   = 'Windows 7'
            '6.1.7601'   = 'Windows 7'
            '6.2.9200'   = 'Windows 8'
            '6.3.9600'   = 'Windows 8.1'
            '10.0.10240' = 'Windows 10'
            '10.0.10586' = 'Windows 10'
            '10.0.14393' = 'Windows 10'
            '10.0.15063' = 'Windows 10'
            '10.0.16299' = 'Windows 10'
            '10.0.17134' = 'Windows 10'
            '10.0.17763' = 'Windows 10'
            '10.0.18362' = 'Windows 10'
            '10.0.18363' = 'Windows 10'
            '10.0.19041' = 'Windows 10'
            '10.0.19042' = 'Windows 10'
            '10.0.19043' = 'Windows 10'
            '10.0.19044' = 'Windows 10'
            '10.0.19045' = 'Windows 10'
            '10.0.22000' = 'Windows 11'
            '10.0.22621' = 'Windows 11'
            '10.0.22631' = 'Windows 11'
            '10.0.26100' = 'Windows 11'
        }
    }

    if ($map.ContainsKey($ver)) { return $map[$ver] }

    # Fallback
    if ($isServer) { return "Windows Server (build $ver)" }
    else           { return "Windows (build $ver)" }
}

# ── Helper: Query a single host ───────────────────────────────────────────────
function Get-HostInfo {
    param(
        [string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $result = [PSCustomObject]@{
        HostName      = $ComputerName
        IPv4Addresses = ''
        Domain        = ''
        OSType        = ''
        Distro        = ''
        Status        = 'Offline'
    }

    # Ping test
    $pingParams = @{ ComputerName = $ComputerName; Count = 1; Quiet = $true }
    if ($Credential) { $pingParams.Credential = $Credential }
    if (-not (Test-Connection @pingParams)) {
        return $result
    }

    # AD lookup for domain / OS (fallback to WMI if not in AD)
    try {
        $adComp = Get-ADComputer -Identity $ComputerName -Properties OperatingSystem, OperatingSystemVersion, DNSHostName -ErrorAction Stop
        $result.Domain = ($adComp.DNSHostName -split '\.', 2)[1]
        $result.OSType = $adComp.OperatingSystem
        $result.Distro = Get-OSFriendlyName -Caption $adComp.OperatingSystem -Version $adComp.OperatingSystemVersion
    } catch {
        # Not in AD or no permission – try WMI
        try {
            $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
            $result.OSType = $os.Caption
            $result.Distro = Get-OSFriendlyName -Caption $os.Caption -Version $os.Version
            $result.Domain = $os.Domain
        } catch {
            $result.OSType = 'Unknown'
            $result.Distro = 'Unknown'
        }
    }

    # IPv4 addresses via DNS / WMI
    try {
        $ip = (Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled = TRUE" -ErrorAction Stop).IPAddress |
              Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' }
        $result.IPv4Addresses = ($ip -join ', ')
    } catch {
        try {
            $ip = [System.Net.Dns]::GetHostAddresses($ComputerName) |
                  Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                  ForEach-Object { $_.IPAddressToString }
            $result.IPv4Addresses = ($ip -join ', ')
        } catch { }
    }

    $result.Status = 'Online'
    return $result
}

# ── Build GUI ─────────────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text = 'AD Server Ping & Inventory Tool'
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Hostname label
$lblHosts = New-Object System.Windows.Forms.Label
$lblHosts.Text = 'Paste hostnames (one per line):'
$lblHosts.Location = New-Object System.Drawing.Point(10, 10)
$lblHosts.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblHosts)

# Hostname text box
$txtHosts = New-Object System.Windows.Forms.TextBox
$txtHosts.Multiline = $true
$txtHosts.ScrollBars = 'Vertical'
$txtHosts.Location = New-Object System.Drawing.Point(10, 35)
$txtHosts.Size = New-Object System.Drawing.Size(400, 200)
$form.Controls.Add($txtHosts)

# Credential group
$grpCred = New-Object System.Windows.Forms.GroupBox
$grpCred.Text = 'Alternate Credentials (optional)'
$grpCred.Location = New-Object System.Drawing.Point(430, 10)
$grpCred.Size = New-Object System.Drawing.Size(440, 150)

$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = 'Domain:'
$lblDomain.Location = New-Object System.Drawing.Point(10, 25)
$lblDomain.Size = New-Object System.Drawing.Size(60, 20)
$grpCred.Controls.Add($lblDomain)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = New-Object System.Drawing.Point(80, 22)
$txtDomain.Size = New-Object System.Drawing.Size(200, 20)
$grpCred.Controls.Add($txtDomain)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = 'User:'
$lblUser.Location = New-Object System.Drawing.Point(10, 55)
$lblUser.Size = New-Object System.Drawing.Size(60, 20)
$grpCred.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(80, 52)
$txtUser.Size = New-Object System.Drawing.Size(200, 20)
$grpCred.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = 'Password:'
$lblPass.Location = New-Object System.Drawing.Point(10, 85)
$lblPass.Size = New-Object System.Drawing.Size(60, 20)
$grpCred.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(80, 82)
$txtPass.Size = New-Object System.Drawing.Size(200, 20)
$txtPass.UseSystemPasswordChar = $true
$grpCred.Controls.Add($txtPass)

$form.Controls.Add($grpCred)

# Run button
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(430, 170)
$btnRun.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnRun)

# Progress bar
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(10, 250)
$progress.Size = New-Object System.Drawing.Size(860, 20)
$form.Controls.Add($progress)

# Status label
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready.'
$lblStatus.Location = New-Object System.Drawing.Point(10, 275)
$lblStatus.Size = New-Object System.Drawing.Size(860, 20)
$form.Controls.Add($lblStatus)

# Results grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 300)
$grid.Size = New-Object System.Drawing.Size(860, 300)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AutoSizeColumnsMode = 'Fill'
$form.Controls.Add($grid)

# Export buttons
$btnCSV = New-Object System.Windows.Forms.Button
$btnCSV.Text = 'Export CSV'
$btnCSV.Location = New-Object System.Drawing.Point(10, 610)
$btnCSV.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnCSV)

$btnXLSX = New-Object System.Windows.Forms.Button
$btnXLSX.Text = 'Export XLSX'
$btnXLSX.Location = New-Object System.Drawing.Point(120, 610)
$btnXLSX.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnXLSX)

$btnHTML = New-Object System.Windows.Forms.Button
$btnHTML.Text = 'Export HTML'
$btnHTML.Location = New-Object System.Drawing.Point(230, 610)
$btnHTML.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnHTML)

# ── Global results store ──────────────────────────────────────────────────────
$script:Results = @()

# ── Run button logic ──────────────────────────────────────────────────────────
$btnRun.Add_Click({
    $hosts = $txtHosts.Text -split "`r?`n" | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
    if ($hosts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Please enter at least one hostname.', 'No Hosts', 'OK', 'Warning')
        return
    }

    # Build credential if provided
    $cred = $null
    if ($txtDomain.Text -and $txtUser.Text -and $txtPass.Text) {
        $secPass = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential("$($txtDomain.Text)\$($txtUser.Text)", $secPass)
    }

    $script:Results = @()
    $progress.Maximum = $hosts.Count
    $progress.Value = 0
    $grid.DataSource = $null

    for ($i = 0; $i -lt $hosts.Count; $i++) {
        $hostName = $hosts[$i]
        $lblStatus.Text = "Processing $hostName ($($i+1)/$($hosts.Count))..."
        $form.Refresh()

        $info = Get-HostInfo -ComputerName $hostName -Credential $cred
        $script:Results += $info

        $progress.Value = $i + 1
    }

    $grid.DataSource = $script:Results | Select-Object HostName, IPv4Addresses, Domain, OSType, Distro, Status
    $lblStatus.Text = "Completed. $($script:Results.Count) hosts processed."
})

# ── Export CSV ────────────────────────────────────────────────────────────────
$btnCSV.Add_Click({
    if ($script:Results.Count -eq 0) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = 'ServerInventory.csv'
    if ($sfd.ShowDialog() -eq 'OK') {
        $script:Results | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exported to $($sfd.FileName)", 'Export Complete', 'OK', 'Information')
    }
})

# ── Export XLSX (requires ImportExcel module) ─────────────────────────────────
$btnXLSX.Add_Click({
    if ($script:Results.Count -eq 0) { return }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        [System.Windows.Forms.MessageBox]::Show('ImportExcel module not found. Install with: Install-Module ImportExcel', 'Module Missing', 'OK', 'Error')
        return
    }
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Excel files (*.xlsx)|*.xlsx'
    $sfd.FileName = 'ServerInventory.xlsx'
    if ($sfd.ShowDialog() -eq 'OK') {
        $script:Results | Export-Excel -Path $sfd.FileName -AutoSize -AutoFilter -FreezeTopRow
        [System.Windows.Forms.MessageBox]::Show("Exported to $($sfd.FileName)", 'Export Complete', 'OK', 'Information')
    }
})

# ── Export HTML (built-in ConvertTo-Html) ─────────────────────────────────────
$btnHTML.Add_Click({
    if ($script:Results.Count -eq 0) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = 'ServerInventory.html'
    if ($sfd.ShowDialog() -eq 'OK') {
        $htmlHead = @"
<style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
    h1 { color: #2c3e50; }
    table { border-collapse: collapse; width: 100%; }
    th { background-color: #34495e; color: white; padding: 8px; text-align: left; }
    td { padding: 6px 8px; border-bottom: 1px solid #ddd; }
    tr:nth-child(even) { background-color: #f2f2f2; }
</style>
"@
        $script:Results |
            Select-Object HostName, IPv4Addresses, Domain, OSType, Distro, Status |
            ConvertTo-Html -Title 'Server Inventory' -Head $htmlHead -Body "<h1>Server Inventory Report – $(Get-Date -Format 'yyyy-MM-dd HH:mm')</h1>" |
            Out-File -FilePath $sfd.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exported to $($sfd.FileName)", 'Export Complete', 'OK', 'Information')
    }
})

# ── Show form ─────────────────────────────────────────────────────────────────
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
