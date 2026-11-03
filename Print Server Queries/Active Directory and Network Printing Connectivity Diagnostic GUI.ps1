
#requires -Version 5.1
<#
.SYNOPSIS
    Active Directory and Network Printing Connectivity Diagnostic GUI.

.DESCRIPTION
    Tests DNS, Active Directory TCP ports, Windows print server
    connectivity, SMB printer sharing, and configurable RPC ports.

    Compatible with Windows PowerShell 5.1 / ISE and PowerShell 7
    on Windows.

    No administrative elevation is required for outbound TCP tests.
    UDP connectivity is not conclusively tested by this tool.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------
# GLOBAL SETTINGS
# ------------------------------------------------------------

$script:Results = @()
$script:Worker = $null

$script:Desktop = [Environment]::GetFolderPath('Desktop')

$script:TimeoutMs = 900

# ------------------------------------------------------------
# TCP CONNECTION TEST FUNCTION
# ------------------------------------------------------------

function Test-TcpConnection {

    param(
        [string]$ComputerName,
        [int]$Port,
        [int]$Timeout = 900
    )

    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $Client = New-Object System.Net.Sockets.TcpClient

    try {

        $Async = $Client.BeginConnect(
            $ComputerName,
            $Port,
            $null,
            $null
        )

        $Connected = $Async.AsyncWaitHandle.WaitOne($Timeout, $false)

        if ($Connected -and $Client.Connected) {

            $Client.EndConnect($Async)
            $Stopwatch.Stop()

            return [PSCustomObject]@{
                Status = 'OPEN'
                ResponseMS = $Stopwatch.ElapsedMilliseconds
                Details = 'TCP connection successful'
            }
        }

        $Stopwatch.Stop()

        return [PSCustomObject]@{
            Status = 'FAILED'
            ResponseMS = $Stopwatch.ElapsedMilliseconds
            Details = 'Connection timed out or was refused'
        }

    }
    catch {

        $Stopwatch.Stop()

        return [PSCustomObject]@{
            Status = 'FAILED'
            ResponseMS = $Stopwatch.ElapsedMilliseconds
            Details = $_.Exception.Message
        }

    }
    finally {

        if ($Client) {
            $Client.Close()
        }

    }
}

# ------------------------------------------------------------
# RESULT CREATION
# ------------------------------------------------------------

function New-DiagnosticResult {

    param(
        [string]$Target,
        [string]$Category,
        [string]$Service,
        [string]$Protocol,
        [string]$Port,
        [string]$Status,
        [string]$Response,
        [string]$Details
    )

    [PSCustomObject]@{
        Timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Source     = $env:COMPUTERNAME
        Domain     = $env:USERDNSDOMAIN
        Target     = $Target
        Category   = $Category
        Service    = $Service
        Protocol   = $Protocol
        Port       = $Port
        Status     = $Status
        ResponseMS = $Response
        Details    = $Details
    }
}

# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form

$form.Text = 'Active Directory & Network Printing Diagnostic Tool'
$form.Size = New-Object System.Drawing.Size(1120, 790)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1000, 700)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$font = New-Object System.Drawing.Font('Segoe UI', 9)

# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------

$header = New-Object System.Windows.Forms.Label
$header.Text = 'ACTIVE DIRECTORY & NETWORK PRINTING DIAGNOSTICS'
$header.Location = New-Object System.Drawing.Point(18, 12)
$header.Size = New-Object System.Drawing.Size(1050, 30)
$header.Font = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
$header.ForeColor = [System.Drawing.Color]::FromArgb(35, 65, 105)

$form.Controls.Add($header)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'Network connectivity testing from multiple workstation locations'
$subtitle.Location = New-Object System.Drawing.Point(20, 43)
$subtitle.Size = New-Object System.Drawing.Size(900, 22)
$subtitle.Font = $font

$form.Controls.Add($subtitle)

# ------------------------------------------------------------
# TARGET INPUTS
# ------------------------------------------------------------

$groupTargets = New-Object System.Windows.Forms.GroupBox
$groupTargets.Text = 'Diagnostic Targets'
$groupTargets.Location = New-Object System.Drawing.Point(18, 75)
$groupTargets.Size = New-Object System.Drawing.Size(1065, 145)
$groupTargets.Font = $font

$form.Controls.Add($groupTargets)

function Add-TargetField {

    param(
        $Parent,
        [string]$LabelText,
        [int]$X,
        [int]$Y,
        [int]$Width = 280,
        [string]$Default = ''
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 20)
    $Parent.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point($X, ($Y + 23))
    $textbox.Size = New-Object System.Drawing.Size($Width, 25)
    $textbox.Text = $Default
    $Parent.Controls.Add($textbox)

    return $textbox
}

$txtDC = Add-TargetField $groupTargets 'Domain Controller (FQDN or IP)' 18 25 310

$txtPrint = Add-TargetField $groupTargets 'Print Server (FQDN or IP)' 350 25 310

$txtPrinter = Add-TargetField $groupTargets 'Optional Network Printer IP' 680 25 340

# ------------------------------------------------------------
# RPC OPTIONS
# ------------------------------------------------------------

$groupRPC = New-Object System.Windows.Forms.GroupBox
$groupRPC.Text = 'RPC Dynamic Port Testing'
$groupRPC.Location = New-Object System.Drawing.Point(18, 230)
$groupRPC.Size = New-Object System.Drawing.Size(1065, 75)
$groupRPC.Font = $font

$form.Controls.Add($groupRPC)

$chkRPC = New-Object System.Windows.Forms.CheckBox
$chkRPC.Text = 'Enable dynamic RPC port scan'
$chkRPC.Location = New-Object System.Drawing.Point(15, 30)
$chkRPC.Size = New-Object System.Drawing.Size(240, 25)
$chkRPC.Checked = $true
$groupRPC.Controls.Add($chkRPC)

$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = 'Start Port:'
$lblStart.Location = New-Object System.Drawing.Point(290, 32)
$lblStart.Size = New-Object System.Drawing.Size(75, 20)
$groupRPC.Controls.Add($lblStart)

$txtStart = New-Object System.Windows.Forms.TextBox
$txtStart.Text = '49152'
$txtStart.Location = New-Object System.Drawing.Point(370, 29)
$txtStart.Size = New-Object System.Drawing.Size(85, 24)
$groupRPC.Controls.Add($txtStart)

$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = 'End Port:'
$lblEnd.Location = New-Object System.Drawing.Point(480, 32)
$lblEnd.Size = New-Object System.Drawing.Size(75, 20)
$groupRPC.Controls.Add($lblEnd)

$txtEnd = New-Object System.Windows.Forms.TextBox
$txtEnd.Text = '49200'
$txtEnd.Location = New-Object System.Drawing.Point(560, 29)
$txtEnd.Size = New-Object System.Drawing.Size(85, 24)
$groupRPC.Controls.Add($txtEnd)

$lblRPCNote = New-Object System.Windows.Forms.Label
$lblRPCNote.Text = 'Full Windows Server dynamic range: 49152-65535'
$lblRPCNote.Location = New-Object System.Drawing.Point(680, 32)
$lblRPCNote.Size = New-Object System.Drawing.Size(350, 20)
$groupRPC.Controls.Add($lblRPCNote)

# ------------------------------------------------------------
# ACTION BUTTONS
# ------------------------------------------------------------

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'RUN DIAGNOSTICS'
$btnRun.Location = New-Object System.Drawing.Point(18, 318)
$btnRun.Size = New-Object System.Drawing.Size(165, 38)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(35, 105, 170)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = 'Flat'
$btnRun.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

$form.Controls.Add($btnRun)

$btnCSV = New-Object System.Windows.Forms.Button
$btnCSV.Text = 'EXPORT CSV'
$btnCSV.Location = New-Object System.Drawing.Point(195, 318)
$btnCSV.Size = New-Object System.Drawing.Size(135, 38)
$form.Controls.Add($btnCSV)

$btnHTML = New-Object System.Windows.Forms.Button
$btnHTML.Text = 'EXPORT HTML'
$btnHTML.Location = New-Object System.Drawing.Point(340, 318)
$btnHTML.Size = New-Object System.Drawing.Size(135, 38)
$form.Controls.Add($btnHTML)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'CLEAR RESULTS'
$btnClear.Location = New-Object System.Drawing.Point(485, 318)
$btnClear.Size = New-Object System.Drawing.Size(135, 38)
$form.Controls.Add($btnClear)

# ------------------------------------------------------------
# PROGRESS
# ------------------------------------------------------------

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(18, 367)
$progress.Size = New-Object System.Drawing.Size(1065, 18)
$progress.Style = 'Continuous'

$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready - Enter your diagnostic targets.'
$lblStatus.Location = New-Object System.Drawing.Point(18, 390)
$lblStatus.Size = New-Object System.Drawing.Size(1065, 22)

$form.Controls.Add($lblStatus)

# ------------------------------------------------------------
# RESULTS GRID
# ------------------------------------------------------------

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(18, 420)
$grid.Size = New-Object System.Drawing.Size(1065, 290)
$grid.Anchor = 'Top, Bottom, Left, Right'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
$grid.RowHeadersVisible = $false
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$columns = @(
    'Target',
    'Category',
    'Service',
    'Protocol',
    'Port',
    'Status',
    'ResponseMS',
    'Details'
)

foreach ($column in $columns) {

    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $column
    $col.HeaderText = $column
    $grid.Columns.Add($col) | Out-Null

}

$form.Controls.Add($grid)

# ------------------------------------------------------------
# BACKGROUND WORKER
# ------------------------------------------------------------

$worker = New-Object System.ComponentModel.BackgroundWorker
$worker.WorkerReportsProgress = $true
$worker.WorkerSupportsCancellation = $false

$worker.add_DoWork({

    param($sender, $e)

    $argsData = $e.Argument

    $DC = $argsData.DC
    $Print = $argsData.Print
    $Printer = $argsData.Printer
    $DoRPC = $argsData.DoRPC
    $RPCStart = $argsData.RPCStart
    $RPCEnd = $argsData.RPCEnd

    $allResults = New-Object System.Collections.Generic.List[object]

    # --------------------------------------------------------
    # DNS TEST
    # --------------------------------------------------------

    $targets = @(
        @{ Name = $DC; Category = 'Active Directory' },
        @{ Name = $Print; Category = 'Printing' }
    )

    if ($Printer) {
        $targets += @{ Name = $Printer; Category = 'Network Printer' }
    }

    foreach ($target in $targets) {

        try {

            $addresses = [System.Net.Dns]::GetHostAddresses($target.Name)

            $ipList = ($addresses | ForEach-Object { $_.IPAddressToString }) -join ', '

            $result = New-DiagnosticResult `
                -Target $target.Name `
                -Category $target.Category `
                -Service 'DNS Name Resolution' `
                -Protocol 'DNS' `
                -Port '53' `
                -Status 'OPEN' `
                -Response '-' `
                -Details "Resolved: $ipList"

        }
        catch {

            $result = New-DiagnosticResult `
                -Target $target.Name `
                -Category $target.Category `
                -Service 'DNS Name Resolution' `
                -Protocol 'DNS' `
                -Port '53' `
                -Status 'FAILED' `
                -Response '-' `
                -Details $_.Exception.Message

        }

        $allResults.Add($result)
        $sender.ReportProgress(0, $result)

    }

    # --------------------------------------------------------
    # DOMAIN CONTROLLER DISCOVERY
    # --------------------------------------------------------

    try {

        $domain = $env:USERDNSDOMAIN

        if ($domain) {

            $nltestOutput = & nltest.exe "/dsgetdc:$domain" 2>&1
            $nltestExit = $LASTEXITCODE

            $discoveryStatus = if ($nltestExit -eq 0) {
                'OPEN'
            } else {
                'FAILED'
            }

            $result = New-DiagnosticResult `
                -Target $DC `
                -Category 'Active Directory' `
                -Service 'Domain Controller Discovery' `
                -Protocol 'DNS/Netlogon' `
                -Port 'N/A' `
                -Status $discoveryStatus `
                -Response '-' `
                -Details (($nltestOutput | Out-String).Trim() -replace '\s+', ' ')

        }
        else {

            $result = New-DiagnosticResult `
                -Target $DC `
                -Category 'Active Directory' `
                -Service 'Domain Controller Discovery' `
                -Protocol 'DNS/Netlogon' `
                -Port 'N/A' `
                -Status 'FAILED' `
                -Response '-' `
                -Details 'Workstation domain could not be identified.'

        }

        $allResults.Add($result)
        $sender.ReportProgress(0, $result)

    }
    catch {

        $result = New-DiagnosticResult `
            -Target $DC `
            -Category 'Active Directory' `
            -Service 'Domain Controller Discovery' `
            -Protocol 'DNS/Netlogon' `
            -Port 'N/A' `
            -Status 'FAILED' `
            -Response '-' `
            -Details $_.Exception.Message

        $allResults.Add($result)
        $sender.ReportProgress(0, $result)

    }

    # --------------------------------------------------------
    # ACTIVE DIRECTORY PORTS
    # --------------------------------------------------------

    $adPorts = @(
        @{ Port = 53;   Service = 'DNS'; Protocol = 'TCP' },
        @{ Port = 88;   Service = 'Kerberos'; Protocol = 'TCP' },
        @{ Port = 135;  Service = 'RPC Endpoint Mapper'; Protocol = 'TCP' },
        @{ Port = 389;  Service = 'LDAP'; Protocol = 'TCP' },
        @{ Port = 445;  Service = 'SMB'; Protocol = 'TCP' },
        @{ Port = 464;  Service = 'Kerberos Password Change'; Protocol = 'TCP' },
        @{ Port = 636;  Service = 'LDAPS'; Protocol = 'TCP' },
        @{ Port = 3268; Service = 'Global Catalog'; Protocol = 'TCP' },
        @{ Port = 3269; Service = 'Secure Global Catalog'; Protocol = 'TCP' },
        @{ Port = 9389; Service = 'Active Directory Web Services'; Protocol = 'TCP' }
    )

    foreach ($item in $adPorts) {

        $test = Test-TcpConnection -ComputerName $DC -Port $item.Port

        $result = New-DiagnosticResult `
            -Target $DC `
            -Category 'Active Directory' `
            -Service $item.Service `
            -Protocol $item.Protocol `
            -Port ([string]$item.Port) `
            -Status $test.Status `
            -Response ([string]$test.ResponseMS) `
            -Details $test.Details

        $allResults.Add($result)
        $sender.ReportProgress(0, $result)

    }

    # --------------------------------------------------------
    # PRINT SERVER PORTS
    # --------------------------------------------------------

    $printPorts = @(
        @{ Port = 135; Service = 'RPC Endpoint Mapper' },
        @{ Port = 139; Service = 'NetBIOS Session' },
        @{ Port = 445; Service = 'SMB / Shared Printers' }
    )

    foreach ($item in $printPorts) {

        $test = Test-TcpConnection -ComputerName $Print -Port $item.Port

        $result = New-DiagnosticResult `
            -Target $Print `
            -Category 'Printing' `
            -Service $item.Service `
            -Protocol 'TCP' `
            -Port ([string]$item.Port) `
            -Status $test.Status `
            -Response ([string]$test.ResponseMS) `
            -Details $test.Details

        $allResults.Add($result)
        $sender.ReportProgress(0, $result)

    }

    # --------------------------------------------------------
    # PRINT$ SHARE ACCESS
    # --------------------------------------------------------

    try {

        $sharePath = "\\$Print\print$"

        $shareAccessible = Test-Path -LiteralPath $sharePath -ErrorAction Stop

        $shareStatus = if ($shareAccessible) {
            'OPEN'
        } else {
            'FAILED'
        }

        $shareDetails = if ($shareAccessible) {
            'Print$ share is accessible using current credentials.'
        } else {
            'Share not accessible. Check SMB, permissions, and share availability.'
        }

    }
    catch {

        $shareStatus = 'FAILED'
        $shareDetails = $_.Exception.Message

    }

    $result = New-DiagnosticResult `
        -Target $Print `
        -Category 'Printing' `
        -Service 'Print$ Administrative Share' `
        -Protocol 'SMB' `
        -Port '445' `
        -Status $shareStatus `
        -Response '-' `
        -Details $shareDetails

    $allResults.Add($result)
    $sender.ReportProgress(0, $result)

    # --------------------------------------------------------
    # DYNAMIC RPC RANGE
    # --------------------------------------------------------

    if ($DoRPC) {

        $rangeCount = $RPCEnd - $RPCStart + 1
        $counter = 0

        for ($port = $RPCStart; $port -le $RPCEnd; $port++) {

            $counter++

            $test = Test-TcpConnection `
                -ComputerName $Print `
                -Port $port `
                -Timeout 350

            if ($test.Status -eq 'OPEN') {

                $result = New-DiagnosticResult `
                    -Target $Print `
                    -Category 'RPC Dynamic Ports' `
                    -Service 'Dynamic RPC Port' `
                    -Protocol 'TCP' `
                    -Port ([string]$port) `
                    -Status 'OPEN' `
                    -Response ([string]$test.ResponseMS) `
                    -Details 'TCP connection successful. Not proof that this port is the assigned print RPC endpoint.'

                $allResults.Add($result)
                $sender.ReportProgress(0, $result)

            }

            if (($counter % 10) -eq 0) {

                $sender.ReportProgress(
                    0,
                    [PSCustomObject]@{
                        Progress = [int](($counter / $rangeCount) * 100)
                        Message = "Scanning RPC ports: $port of $RPCEnd"
                    }
                )

            }

        }

        $result = New-DiagnosticResult `
            -Target $Print `
            -Category 'RPC Dynamic Ports' `
            -Service 'RPC Range Scan Summary' `
            -Protocol 'TCP' `
            -Port "$RPCStart-$RPCEnd" `
            -Status 'COMPLETED' `
            -Response '-' `
            -Details "$counter ports scanned. Only successful connections are listed individually."

        $allResults.Add($result)
        $sender.ReportProgress(0, $result)

    }

    # --------------------------------------------------------
    # DIRECT NETWORK PRINTER PORTS
    # --------------------------------------------------------

    if ($Printer) {

        $printerPorts = @(
            @{ Port = 9100; Service = 'RAW Printing' },
            @{ Port = 515; Service = 'LPR / LPD' },
            @{ Port = 631; Service = 'Internet Printing Protocol' }
        )

        foreach ($item in $printerPorts) {

            $test = Test-TcpConnection `
                -ComputerName $Printer `
                -Port $item.Port

            $result = New-DiagnosticResult `
                -Target $Printer `
                -Category 'Network Printer' `
                -Service $item.Service `
                -Protocol 'TCP' `
                -Port ([string]$item.Port) `
                -Status $test.Status `
                -Response ([string]$test.ResponseMS) `
                -Details $test.Details

            $allResults.Add($result)
            $sender.ReportProgress(0, $result)

        }

    }

    $e.Result = $allResults.ToArray()

})

# ------------------------------------------------------------
# PROGRESS AND GRID UPDATE
# ------------------------------------------------------------

$worker.add_ProgressChanged({

    param($sender, $e)

    $item = $e.UserState

    if ($item.Progress -ne $null -and $item.Message) {

        $progress.Value = [Math]::Max(
            0,
            [Math]::Min(100, [int]$item.Progress)
        )

        $lblStatus.Text = $item.Message

        return
    }

    $grid.Rows.Add(
        $item.Target,
        $item.Category,
        $item.Service,
        $item.Protocol,
        $item.Port,
        $item.Status,
        $item.ResponseMS,
        $item.Details
    ) | Out-Null

    if ($item.Status -eq 'FAILED') {

        $row = $grid.Rows[$grid.Rows.Count - 1]
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose

    }
    elseif ($item.Status -eq 'OPEN') {

        $row = $grid.Rows[$grid.Rows.Count - 1]
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Honeydew

    }

    $lblStatus.Text = "Testing... Results collected: $($grid.Rows.Count)"

    $grid.FirstDisplayedScrollingRowIndex = $grid.Rows.Count - 1

})

# ------------------------------------------------------------
# COMPLETION
# ------------------------------------------------------------

$worker.add_RunWorkerCompleted({

    param($sender, $e)

    $btnRun.Enabled = $true
    $btnRun.Text = 'RUN DIAGNOSTICS'
    $progress.Value = 100

    if ($e.Error) {

        $lblStatus.Text = "ERROR: $($e.Error.Message)"

        [System.Windows.Forms.MessageBox]::Show(
            $e.Error.Message,
            'Diagnostic Error',
            'OK',
            'Error'
        )

        return
    }

    $script:Results = @($e.Result)

    $openCount = @($script:Results | Where-Object Status -eq 'OPEN').Count
    $failedCount = @($script:Results | Where-Object Status -eq 'FAILED').Count

    $lblStatus.Text = "Completed. Successful: $openCount | Failed: $failedCount | Total records: $($script:Results.Count)"

    [System.Windows.Forms.MessageBox]::Show(
        "Diagnostics completed.`r`n`r`nSuccessful: $openCount`r`nFailed: $failedCount`r`nTotal records: $($script:Results.Count)`r`n`r`nUse EXPORT CSV or EXPORT HTML to save the results.",
        'Diagnostic Complete',
        'OK',
        'Information'
    ) | Out-Null

})

# ------------------------------------------------------------
# RUN BUTTON
# ------------------------------------------------------------

$btnRun.Add_Click({

    $DC = $txtDC.Text.Trim()
    $Print = $txtPrint.Text.Trim()
    $Printer = $txtPrinter.Text.Trim()

    if (!$DC -or !$Print) {

        [System.Windows.Forms.MessageBox]::Show(
            'Please enter both a Domain Controller and Print Server.',
            'Missing Information'
        )

        return
    }

    if ($chkRPC.Checked) {

        $start = 0
        $end = 0

        $validStart = [int]::TryParse($txtStart.Text, [ref]$start)
        $validEnd = [int]::TryParse($txtEnd.Text, [ref]$end)

        if (
            !$validStart -or
            !$validEnd -or
            $start -lt 1 -or
            $end -gt 65535 -or
            $start -gt $end
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                'Enter a valid RPC port range between 1 and 65535.',
                'Invalid RPC Range'
            )

            return
        }

    }
    else {

        $start = 0
        $end = 0

    }

    $grid.Rows.Clear()
    $script:Results = @()
    $progress.Value = 0

    $btnRun.Enabled = $false
    $btnRun.Text = 'TESTING...'

    $lblStatus.Text = 'Starting network diagnostics...'

    $worker.RunWorkerAsync(@{
        DC = $DC
        Print = $Print
        Printer = $Printer
        DoRPC = $chkRPC.Checked
        RPCStart = $start
        RPCEnd = $end
    })

})

# ------------------------------------------------------------
# EXPORT CSV
# ------------------------------------------------------------

$btnCSV.Add_Click({

    if ($script:Results.Count -eq 0) {

        [System.Windows.Forms.MessageBox]::Show(
            'No diagnostic results available.',
            'Export'
        )

        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $path = Join-Path $script:Desktop "AD_Print_Network_Diagnostics_$timestamp.csv"

    try {

        $script:Results |
            Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "CSV report successfully exported:`r`n`r`n$path",
            'CSV Export Complete'
        )

    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'CSV Export Error'
        )

    }

})

# ------------------------------------------------------------
# EXPORT HTML
# ------------------------------------------------------------

$btnHTML.Add_Click({

    if ($script:Results.Count -eq 0) {

        [System.Windows.Forms.MessageBox]::Show(
            'No diagnostic results available.',
            'Export'
        )

        return
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $path = Join-Path $script:Desktop "AD_Print_Network_Diagnostics_$timestamp.html"

    $openCount = @($script:Results | Where-Object Status -eq 'OPEN').Count
    $failedCount = @($script:Results | Where-Object Status -eq 'FAILED').Count

    $table = $script:Results |
        ConvertTo-Html -Fragment -Property Timestamp,Source,Domain,Target,Category,Service,Protocol,Port,Status,ResponseMS,Details

    $style = @'
<style>
body {
    font-family: Segoe UI, Arial;
    margin: 25px;
    background: #f4f6f9;
    color: #263445;
}
h1 { color: #23466d; }
.summary {
    padding: 15px;
    background: white;
    border: 1px solid #ddd;
    margin-bottom: 20px;
}
table {
    border-collapse: collapse;
    width: 100%;
    background: white;
    font-size: 12px;
}
th {
    background: #23466d;
    color: white;
    padding: 10px;
    text-align: left;
}
td {
    border: 1px solid #ddd;
    padding: 7px;
    word-break: break-word;
}
tr:nth-child(even) { background: #f2f5f8; }
</style>
'@

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>AD and Network Printing Diagnostics</title>
$style
</head>
<body>
<h1>Active Directory & Network Printing Diagnostics</h1>
<div class="summary">
<strong>Source Workstation:</strong> $env:COMPUTERNAME<br>
<strong>Domain:</strong> $env:USERDNSDOMAIN<br>
<strong>Report Generated:</strong> $(Get-Date)<br>
<strong>Successful Tests:</strong> $openCount<br>
<strong>Failed Tests:</strong> $failedCount<br>
<strong>Total Records:</strong> $($script:Results.Count)
</div>
$table
</body>
</html>
"@

    try {

        $html | Out-File -FilePath $path -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "HTML report successfully exported:`r`n`r`n$path",
            'HTML Export Complete'
        )

    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'HTML Export Error'
        )

    }

})

# ------------------------------------------------------------
# CLEAR BUTTON
# ------------------------------------------------------------

$btnClear.Add_Click({

    if ($worker.IsBusy) {
        return
    }

    $grid.Rows.Clear()
    $script:Results = @()
    $progress.Value = 0
    $lblStatus.Text = 'Results cleared. Ready for another test.'

})

# ------------------------------------------------------------
# START GUI
# ------------------------------------------------------------

[void]$form.ShowDialog()
