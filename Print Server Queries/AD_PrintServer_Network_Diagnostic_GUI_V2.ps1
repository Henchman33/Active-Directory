
#requires -Version 5.1
<#
.SYNOPSIS
    Active Directory & Network Printing Diagnostic Tool V2.0

.DESCRIPTION
    Tests Active Directory, DNS, Windows printing, SMB, RPC,
    and optional physical printer connectivity.

    Designed for:
      - Windows PowerShell 5.1
      - PowerShell ISE 5.1
      - PowerShell 7 on Windows

    Uses synchronous PowerShell execution with Windows Forms
    event processing. No BackgroundWorker or shared Runspace.

    Exports CSV and HTML reports to the current user's Desktop.

.NOTES
    TCP connectivity tests only.
    UDP connectivity is not conclusively tested.
    A successful TCP connection does not prove application-level
    authentication or printing functionality.
#>

# ------------------------------------------------------------
# INITIALIZATION
# ------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Results = New-Object System.Collections.ArrayList
$script:IsRunning = $false
$script:CancelRequested = $false

$script:Desktop = [Environment]::GetFolderPath('Desktop')
$script:TimeoutMs = 900
$script:RpcTimeoutMs = 350

$script:Workstation = $env:COMPUTERNAME
$script:Domain = $env:USERDNSDOMAIN
$script:Username = $env:USERNAME

# ------------------------------------------------------------
# TCP CONNECTIVITY TEST
# ------------------------------------------------------------

function Test-TcpConnection {

    param(
        [string]$ComputerName,
        [int]$Port,
        [int]$Timeout = 900
    )

    $Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $Client = New-Object System.Net.Sockets.TcpClient
    $Async = $null

    try {

        $Async = $Client.BeginConnect(
            $ComputerName,
            $Port,
            $null,
            $null
        )

        $Connected = $Async.AsyncWaitHandle.WaitOne(
            $Timeout,
            $false
        )

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

        if ($Async -and $Async.AsyncWaitHandle) {
            $Async.AsyncWaitHandle.Close()
        }

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

    return [PSCustomObject]@{

        Timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Source     = $script:Workstation
        Domain     = $script:Domain
        Username   = $script:Username
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
# ADD RESULT TO GUI AND RESULTS COLLECTION
# ------------------------------------------------------------

function Add-DiagnosticResult {

    param(
        [PSCustomObject]$Result
    )

    [void]$script:Results.Add($Result)

    $Index = $grid.Rows.Add(
        $Result.Target,
        $Result.Category,
        $Result.Service,
        $Result.Protocol,
        $Result.Port,
        $Result.Status,
        $Result.ResponseMS,
        $Result.Details
    )

    $Row = $grid.Rows[$Index]

    switch ($Result.Status) {

        'OPEN' {
            $Row.DefaultCellStyle.BackColor =
                [System.Drawing.Color]::Honeydew

            $Row.DefaultCellStyle.ForeColor =
                [System.Drawing.Color]::DarkGreen
        }

        'FAILED' {
            $Row.DefaultCellStyle.BackColor =
                [System.Drawing.Color]::MistyRose

            $Row.DefaultCellStyle.ForeColor =
                [System.Drawing.Color]::DarkRed
        }

        'INFO' {
            $Row.DefaultCellStyle.BackColor =
                [System.Drawing.Color]::LightYellow
        }

        'COMPLETED' {
            $Row.DefaultCellStyle.BackColor =
                [System.Drawing.Color]::LightBlue
        }

    }

    $lblStatus.Text = "Results collected: $($script:Results.Count)"

    if ($grid.Rows.Count -gt 0) {

        $grid.FirstDisplayedScrollingRowIndex =
            $grid.Rows.Count - 1

    }

    # Process pending Windows Forms events.
    [System.Windows.Forms.Application]::DoEvents()

}

# ------------------------------------------------------------
# RUN A SINGLE TCP TEST
# ------------------------------------------------------------

function Invoke-PortTest {

    param(
        [string]$Target,
        [string]$Category,
        [string]$Service,
        [int]$Port,
        [int]$Timeout = 900
    )

    if ($script:CancelRequested) {
        return
    }

    $lblStatus.Text = "Testing $Target : $Port - $Service"

    [System.Windows.Forms.Application]::DoEvents()

    $Test = Test-TcpConnection `
        -ComputerName $Target `
        -Port $Port `
        -Timeout $Timeout

    $Result = New-DiagnosticResult `
        -Target $Target `
        -Category $Category `
        -Service $Service `
        -Protocol 'TCP' `
        -Port ([string]$Port) `
        -Status $Test.Status `
        -Response ([string]$Test.ResponseMS) `
        -Details $Test.Details

    Add-DiagnosticResult $Result

}

# ------------------------------------------------------------
# DNS RESOLUTION TEST
# ------------------------------------------------------------

function Invoke-DnsTest {

    param(
        [string]$Target,
        [string]$Category
    )

    if ($script:CancelRequested) {
        return
    }

    $lblStatus.Text = "Resolving DNS: $Target"
    [System.Windows.Forms.Application]::DoEvents()

    try {

        $Addresses = [System.Net.Dns]::GetHostAddresses($Target)

        $IPv4 = @(
            $Addresses |
                Where-Object {
                    $_.AddressFamily -eq
                    [System.Net.Sockets.AddressFamily]::InterNetwork
                } |
                ForEach-Object {
                    $_.IPAddressToString
                }
        )

        if ($IPv4.Count -gt 0) {

            $Status = 'OPEN'
            $Details = "IPv4: $($IPv4 -join ', ')"

        }
        else {

            $Status = 'INFO'
            $Details = 'DNS resolved, but no IPv4 address was returned.'

        }

    }
    catch {

        $Status = 'FAILED'
        $Details = $_.Exception.Message

    }

    $Result = New-DiagnosticResult `
        -Target $Target `
        -Category $Category `
        -Service 'DNS Name Resolution' `
        -Protocol 'DNS' `
        -Port '53' `
        -Status $Status `
        -Response '-' `
        -Details $Details

    Add-DiagnosticResult $Result

}

# ------------------------------------------------------------
# DOMAIN CONTROLLER DISCOVERY
# ------------------------------------------------------------

function Invoke-DCDiscovery {

    param(
        [string]$Target
    )

    if ($script:CancelRequested) {
        return
    }

    $lblStatus.Text = 'Discovering Active Directory Domain Controller...'
    [System.Windows.Forms.Application]::DoEvents()

    try {

        $DomainName = $script:Domain

        if (!$DomainName) {
            throw 'USERDNSDOMAIN is unavailable. The workstation may not be domain joined.'
        }

        $Output = & nltest.exe "/dsgetdc:$DomainName" 2>&1
        $ExitCode = $LASTEXITCODE

        if ($ExitCode -eq 0) {

            $Status = 'OPEN'

            $Details = (
                $Output | Out-String
            ).Trim() -replace '\s+', ' '

        }
        else {

            $Status = 'FAILED'

            $Details = (
                $Output | Out-String
            ).Trim() -replace '\s+', ' '

        }

    }
    catch {

        $Status = 'FAILED'
        $Details = $_.Exception.Message

    }

    $Result = New-DiagnosticResult `
        -Target $Target `
        -Category 'Active Directory' `
        -Service 'Domain Controller Discovery' `
        -Protocol 'DNS / Netlogon' `
        -Port 'N/A' `
        -Status $Status `
        -Response '-' `
        -Details $Details

    Add-DiagnosticResult $Result

}

# ------------------------------------------------------------
# PRINT$ SHARE TEST
# ------------------------------------------------------------

function Invoke-PrintShareTest {

    param(
        [string]$Target
    )

    if ($script:CancelRequested) {
        return
    }

    $lblStatus.Text = "Testing Print$ share on $Target"
    [System.Windows.Forms.Application]::DoEvents()

    try {

        $SharePath = "\\$Target\print$"

        $Accessible = Test-Path `
            -LiteralPath $SharePath `
            -ErrorAction Stop

        if ($Accessible) {

            $Status = 'OPEN'
            $Details = 'Print$ share is accessible using current credentials.'

        }
        else {

            $Status = 'FAILED'
            $Details = 'Print$ share not accessible. Check SMB, permissions, and share availability.'

        }

    }
    catch {

        $Status = 'FAILED'
        $Details = $_.Exception.Message

    }

    $Result = New-DiagnosticResult `
        -Target $Target `
        -Category 'Printing' `
        -Service 'Print$ Administrative Share' `
        -Protocol 'SMB' `
        -Port '445' `
        -Status $Status `
        -Response '-' `
        -Details $Details

    Add-DiagnosticResult $Result

}

# ------------------------------------------------------------
# MAIN DIAGNOSTIC ENGINE
# ------------------------------------------------------------

function Invoke-Diagnostics {

    if ($script:IsRunning) {
        return
    }

    $DC = $txtDC.Text.Trim()
    $Print = $txtPrint.Text.Trim()
    $Printer = $txtPrinter.Text.Trim()

    if (!$DC -or !$Print) {

        [System.Windows.Forms.MessageBox]::Show(
            'Please enter both a Domain Controller and Print Server.',
            'Missing Information',
            'OK',
            'Warning'
        ) | Out-Null

        return

    }

    # --------------------------------------------------------
    # VALIDATE RPC RANGE
    # --------------------------------------------------------

    $RPCStart = 0
    $RPCEnd = 0

    if ($chkRPC.Checked) {

        $ValidStart = [int]::TryParse(
            $txtStart.Text,
            [ref]$RPCStart
        )

        $ValidEnd = [int]::TryParse(
            $txtEnd.Text,
            [ref]$RPCEnd
        )

        if (
            !$ValidStart -or
            !$ValidEnd -or
            $RPCStart -lt 1 -or
            $RPCEnd -gt 65535 -or
            $RPCStart -gt $RPCEnd
        ) {

            [System.Windows.Forms.MessageBox]::Show(
                'Enter a valid RPC port range between 1 and 65535.',
                'Invalid RPC Range'
            ) | Out-Null

            return

        }

        if (($RPCEnd - $RPCStart) -gt 1000) {

            $Confirmation = [System.Windows.Forms.MessageBox]::Show(
                "You selected $($RPCEnd - $RPCStart + 1) RPC ports.`r`n`r`nA full scan of 49152-65535 can take considerable time.`r`n`r`nContinue?",
                'Large RPC Scan',
                'YesNo',
                'Warning'
            )

            if ($Confirmation -ne 'Yes') {
                return
            }

        }

    }

    # --------------------------------------------------------
    # RESET GUI
    # --------------------------------------------------------

    $script:Results.Clear()
    $script:CancelRequested = $false
    $script:IsRunning = $true

    $grid.Rows.Clear()

    $progress.Value = 0
    $progress.Maximum = 100

    $btnRun.Enabled = $false
    $btnCancel.Enabled = $true

    $lblStatus.Text = 'Starting diagnostics...'

    [System.Windows.Forms.Application]::DoEvents()

    try {

        # ----------------------------------------------------
        # TARGET DNS
        # ----------------------------------------------------

        Invoke-DnsTest -Target $DC -Category 'Active Directory'

        if ($script:CancelRequested) {
            return
        }

        Invoke-DnsTest -Target $Print -Category 'Printing'

        if ($script:CancelRequested) {
            return
        }

        if ($Printer) {

            Invoke-DnsTest `
                -Target $Printer `
                -Category 'Network Printer'

        }

        if ($script:CancelRequested) {
            return
        }

        # ----------------------------------------------------
        # DOMAIN CONTROLLER DISCOVERY
        # ----------------------------------------------------

        Invoke-DCDiscovery -Target $DC

        if ($script:CancelRequested) {
            return
        }

        # ----------------------------------------------------
        # ACTIVE DIRECTORY TCP PORTS
        # ----------------------------------------------------

        $ADPorts = @(
            @{ Port = 53; Service = 'DNS' },
            @{ Port = 88; Service = 'Kerberos Authentication' },
            @{ Port = 135; Service = 'RPC Endpoint Mapper' },
            @{ Port = 389; Service = 'LDAP' },
            @{ Port = 445; Service = 'SMB' },
            @{ Port = 464; Service = 'Kerberos Password Change' },
            @{ Port = 636; Service = 'LDAPS' },
            @{ Port = 3268; Service = 'Global Catalog' },
            @{ Port = 3269; Service = 'Secure Global Catalog' },
            @{ Port = 9389; Service = 'Active Directory Web Services' }
        )

        foreach ($Item in $ADPorts) {

            if ($script:CancelRequested) {
                break
            }

            Invoke-PortTest `
                -Target $DC `
                -Category 'Active Directory' `
                -Service $Item.Service `
                -Port $Item.Port

        }

        if ($script:CancelRequested) {
            return
        }

        # ----------------------------------------------------
        # PRINT SERVER TCP PORTS
        # ----------------------------------------------------

        $PrintPorts = @(
            @{ Port = 135; Service = 'RPC Endpoint Mapper' },
            @{ Port = 139; Service = 'NetBIOS Session' },
            @{ Port = 445; Service = 'SMB / Shared Printers' }
        )

        foreach ($Item in $PrintPorts) {

            if ($script:CancelRequested) {
                break
            }

            Invoke-PortTest `
                -Target $Print `
                -Category 'Printing' `
                -Service $Item.Service `
                -Port $Item.Port

        }

        if ($script:CancelRequested) {
            return
        }

        # ----------------------------------------------------
        # PRINT$ SHARE
        # ----------------------------------------------------

        Invoke-PrintShareTest -Target $Print

        if ($script:CancelRequested) {
            return
        }

        # ----------------------------------------------------
        # DYNAMIC RPC SCAN
        # ----------------------------------------------------

        if ($chkRPC.Checked) {

            $TotalPorts = $RPCEnd - $RPCStart + 1
            $Counter = 0
            $OpenRPC = 0

            for ($Port = $RPCStart; $Port -le $RPCEnd; $Port++) {

                if ($script:CancelRequested) {
                    break
                }

                $Counter++

                $lblStatus.Text = "RPC Scan: $Port of $RPCEnd"
                $progress.Value = [Math]::Min(
                    100,
                    [int](($Counter / $TotalPorts) * 100)
                )

                [System.Windows.Forms.Application]::DoEvents()

                $Test = Test-TcpConnection `
                    -ComputerName $Print `
                    -Port $Port `
                    -Timeout $script:RpcTimeoutMs

                if ($Test.Status -eq 'OPEN') {

                    $OpenRPC++

                    $Result = New-DiagnosticResult `
                        -Target $Print `
                        -Category 'RPC Dynamic Ports' `
                        -Service 'Dynamic RPC Port' `
                        -Protocol 'TCP' `
                        -Port ([string]$Port) `
                        -Status 'OPEN' `
                        -Response ([string]$Test.ResponseMS) `
                        -Details 'TCP connection successful. Not proof that this is the assigned print RPC endpoint.'

                    Add-DiagnosticResult $Result

                }

            }

            if (!$script:CancelRequested) {

                $Result = New-DiagnosticResult `
                    -Target $Print `
                    -Category 'RPC Dynamic Ports' `
                    -Service 'RPC Scan Summary' `
                    -Protocol 'TCP' `
                    -Port "$RPCStart-$RPCEnd" `
                    -Status 'COMPLETED' `
                    -Response '-' `
                    -Details "$Counter ports scanned. $OpenRPC TCP connections succeeded."

                Add-DiagnosticResult $Result

            }

        }

        if ($script:CancelRequested) {
            return
        }

        # ----------------------------------------------------
        # PHYSICAL PRINTER TESTS
        # ----------------------------------------------------

        if ($Printer) {

            $PrinterPorts = @(
                @{ Port = 9100; Service = 'RAW Printing' },
                @{ Port = 515; Service = 'LPR / LPD' },
                @{ Port = 631; Service = 'Internet Printing Protocol' }
            )

            foreach ($Item in $PrinterPorts) {

                if ($script:CancelRequested) {
                    break
                }

                Invoke-PortTest `
                    -Target $Printer `
                    -Category 'Network Printer' `
                    -Service $Item.Service `
                    -Port $Item.Port

            }

        }

    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'Diagnostic Error',
            'OK',
            'Error'
        ) | Out-Null

    }
    finally {

        $script:IsRunning = $false

        $btnRun.Enabled = $true
        $btnCancel.Enabled = $false

        if ($script:CancelRequested) {

            $lblStatus.Text = 'Diagnostics cancelled. Partial results are available for export.'

        }
        else {

            $progress.Value = 100

            $OpenCount = @(
                $script:Results |
                    Where-Object { $_.Status -eq 'OPEN' }
            ).Count

            $FailedCount = @(
                $script:Results |
                    Where-Object { $_.Status -eq 'FAILED' }
            ).Count

            $lblStatus.Text = "Completed | Successful: $OpenCount | Failed: $FailedCount | Records: $($script:Results.Count)"

        }

        [System.Windows.Forms.Application]::DoEvents()

    }

    if (!$script:CancelRequested) {

        $OpenCount = @(
            $script:Results |
                Where-Object { $_.Status -eq 'OPEN' }
        ).Count

        $FailedCount = @(
            $script:Results |
                Where-Object { $_.Status -eq 'FAILED' }
        ).Count

        [System.Windows.Forms.MessageBox]::Show(
            "Diagnostics completed.`r`n`r`nSuccessful: $OpenCount`r`nFailed: $FailedCount`r`nTotal records: $($script:Results.Count)`r`n`r`nExport your CSV and HTML reports using the buttons.",
            'Diagnostic Complete',
            'OK',
            'Information'
        ) | Out-Null

    }

}

# ------------------------------------------------------------
# GUI FORM
# ------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form

$form.Text = 'Active Directory & Network Printing Diagnostic Tool V2.0'
$form.Size = New-Object System.Drawing.Size(1280, 850)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$Font = New-Object System.Drawing.Font('Segoe UI', 9)

# ------------------------------------------------------------
# HEADER
# ------------------------------------------------------------

$header = New-Object System.Windows.Forms.Label

$header.Text = 'ACTIVE DIRECTORY & NETWORK PRINTING DIAGNOSTICS'

$header.Location = New-Object System.Drawing.Point(18, 12)
$header.Size = New-Object System.Drawing.Size(1150, 32)

$header.Font = New-Object System.Drawing.Font(
    'Segoe UI',
    15,
    [System.Drawing.FontStyle]::Bold
)

$header.ForeColor = [System.Drawing.Color]::FromArgb(35, 65, 105)

$form.Controls.Add($header)

$subtitle = New-Object System.Windows.Forms.Label

$subtitle.Text = "Network connectivity testing | Source: $script:Workstation | Domain: $script:Domain"

$subtitle.Location = New-Object System.Drawing.Point(20, 45)
$subtitle.Size = New-Object System.Drawing.Size(1150, 22)

$subtitle.Font = $Font

$form.Controls.Add($subtitle)

# ------------------------------------------------------------
# TARGET INPUT GROUP
# ------------------------------------------------------------

$groupTargets = New-Object System.Windows.Forms.GroupBox

$groupTargets.Text = 'Diagnostic Targets'
$groupTargets.Location = New-Object System.Drawing.Point(18, 78)
$groupTargets.Size = New-Object System.Drawing.Size(1225, 140)

$groupTargets.Font = $Font

$form.Controls.Add($groupTargets)

function Add-TargetField {

    param(
        $Parent,
        [string]$LabelText,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [string]$Default = ''
    )

    $Label = New-Object System.Windows.Forms.Label

    $Label.Text = $LabelText
    $Label.Location = New-Object System.Drawing.Point($X, $Y)
    $Label.Size = New-Object System.Drawing.Size($Width, 20)

    $Parent.Controls.Add($Label)

    $TextBox = New-Object System.Windows.Forms.TextBox

    $TextBox.Location = New-Object System.Drawing.Point($X, ($Y + 23))
    $TextBox.Size = New-Object System.Drawing.Size($Width, 25)
    $TextBox.Text = $Default

    $Parent.Controls.Add($TextBox)

    return $TextBox

}

$txtDC = Add-TargetField `
    $groupTargets `
    'Domain Controller (FQDN or IP)' `
    18 25 350

$txtPrint = Add-TargetField `
    $groupTargets `
    'Print Server (FQDN or IP)' `
    390 25 350

$txtPrinter = Add-TargetField `
    $groupTargets `
    'Optional Physical Network Printer IP / Hostname' `
    765 25 420

# ------------------------------------------------------------
# RPC GROUP
# ------------------------------------------------------------

$groupRPC = New-Object System.Windows.Forms.GroupBox

$groupRPC.Text = 'RPC Dynamic Port Testing'
$groupRPC.Location = New-Object System.Drawing.Point(18, 228)
$groupRPC.Size = New-Object System.Drawing.Size(1225, 85)

$groupRPC.Font = $Font

$form.Controls.Add($groupRPC)

$chkRPC = New-Object System.Windows.Forms.CheckBox

$chkRPC.Text = 'Enable dynamic RPC port scan'
$chkRPC.Location = New-Object System.Drawing.Point(15, 35)
$chkRPC.Size = New-Object System.Drawing.Size(250, 25)
$chkRPC.Checked = $true

$groupRPC.Controls.Add($chkRPC)

$lblStart = New-Object System.Windows.Forms.Label

$lblStart.Text = 'Start Port:'
$lblStart.Location = New-Object System.Drawing.Point(290, 38)
$lblStart.Size = New-Object System.Drawing.Size(75, 20)

$groupRPC.Controls.Add($lblStart)

$txtStart = New-Object System.Windows.Forms.TextBox

$txtStart.Text = '49152'
$txtStart.Location = New-Object System.Drawing.Point(370, 34)
$txtStart.Size = New-Object System.Drawing.Size(85, 24)

$groupRPC.Controls.Add($txtStart)

$lblEnd = New-Object System.Windows.Forms.Label

$lblEnd.Text = 'End Port:'
$lblEnd.Location = New-Object System.Drawing.Point(480, 38)
$lblEnd.Size = New-Object System.Drawing.Size(75, 20)

$groupRPC.Controls.Add($lblEnd)

$txtEnd = New-Object System.Windows.Forms.TextBox

$txtEnd.Text = '49200'
$txtEnd.Location = New-Object System.Drawing.Point(560, 34)
$txtEnd.Size = New-Object System.Drawing.Size(85, 24)

$groupRPC.Controls.Add($txtEnd)

$lblRPCNote = New-Object System.Windows.Forms.Label

$lblRPCNote.Text = 'Full Windows Server dynamic range: 49152-65535'

$lblRPCNote.Location = New-Object System.Drawing.Point(680, 38)
$lblRPCNote.Size = New-Object System.Drawing.Size(450, 20)

$groupRPC.Controls.Add($lblRPCNote)

# ------------------------------------------------------------
# BUTTONS
# ------------------------------------------------------------

$btnRun = New-Object System.Windows.Forms.Button

$btnRun.Text = 'RUN DIAGNOSTICS'
$btnRun.Location = New-Object System.Drawing.Point(18, 326)
$btnRun.Size = New-Object System.Drawing.Size(165, 40)

$btnRun.BackColor = [System.Drawing.Color]::FromArgb(35, 105, 170)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = 'Flat'

$form.Controls.Add($btnRun)

$btnCancel = New-Object System.Windows.Forms.Button

$btnCancel.Text = 'CANCEL TEST'
$btnCancel.Location = New-Object System.Drawing.Point(195, 326)
$btnCancel.Size = New-Object System.Drawing.Size(135, 40)

$btnCancel.Enabled = $false

$form.Controls.Add($btnCancel)

$btnCSV = New-Object System.Windows.Forms.Button

$btnCSV.Text = 'EXPORT CSV'
$btnCSV.Location = New-Object System.Drawing.Point(342, 326)
$btnCSV.Size = New-Object System.Drawing.Size(135, 40)

$form.Controls.Add($btnCSV)

$btnHTML = New-Object System.Windows.Forms.Button

$btnHTML.Text = 'EXPORT HTML'
$btnHTML.Location = New-Object System.Drawing.Point(489, 326)
$btnHTML.Size = New-Object System.Drawing.Size(135, 40)

$form.Controls.Add($btnHTML)

$btnClear = New-Object System.Windows.Forms.Button

$btnClear.Text = 'CLEAR RESULTS'
$btnClear.Location = New-Object System.Drawing.Point(636, 326)
$btnClear.Size = New-Object System.Drawing.Size(135, 40)

$form.Controls.Add($btnClear)

# ------------------------------------------------------------
# PROGRESS
# ------------------------------------------------------------

$progress = New-Object System.Windows.Forms.ProgressBar

$progress.Location = New-Object System.Drawing.Point(18, 378)
$progress.Size = New-Object System.Drawing.Size(1225, 20)

$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0

$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label

$lblStatus.Text = 'Ready - Enter diagnostic targets.'

$lblStatus.Location = New-Object System.Drawing.Point(18, 405)
$lblStatus.Size = New-Object System.Drawing.Size(1225, 22)

$form.Controls.Add($lblStatus)

# ------------------------------------------------------------
# RESULTS GRID
# ------------------------------------------------------------

$grid = New-Object System.Windows.Forms.DataGridView

$grid.Location = New-Object System.Drawing.Point(18, 435)
$grid.Size = New-Object System.Drawing.Size(1225, 330)

$grid.Anchor = 'Top, Bottom, Left, Right'

$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false

$grid.AutoSizeColumnsMode = 'Fill'
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.RowHeadersVisible = $false

$grid.BackgroundColor = [System.Drawing.Color]::White

$grid.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$Columns = @(
    'Target',
    'Category',
    'Service',
    'Protocol',
    'Port',
    'Status',
    'ResponseMS',
    'Details'
)

foreach ($Column in $Columns) {

    $Col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn

    $Col.Name = $Column
    $Col.HeaderText = $Column

    [void]$grid.Columns.Add($Col)

}

$form.Controls.Add($grid)

# ------------------------------------------------------------
# BUTTON EVENTS
# ------------------------------------------------------------

$btnRun.Add_Click({

    Invoke-Diagnostics

})

$btnCancel.Add_Click({

    $script:CancelRequested = $true

    $lblStatus.Text = 'Cancellation requested. Finishing current test...'

    [System.Windows.Forms.Application]::DoEvents()

})

$btnClear.Add_Click({

    if ($script:IsRunning) {
        return
    }

    $grid.Rows.Clear()
    $script:Results.Clear()

    $progress.Value = 0

    $lblStatus.Text = 'Results cleared. Ready for another diagnostic.'

})

# ------------------------------------------------------------
# CSV EXPORT
# ------------------------------------------------------------

$btnCSV.Add_Click({

    if ($script:Results.Count -eq 0) {

        [System.Windows.Forms.MessageBox]::Show(
            'No diagnostic results available.',
            'CSV Export'
        ) | Out-Null

        return

    }

    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $Path = Join-Path $script:Desktop (
        "AD_Print_Network_Diagnostics_$Timestamp.csv"
    )

    try {

        $script:Results |
            Export-Csv `
                -Path $Path `
                -NoTypeInformation `
                -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "CSV report exported successfully:`r`n`r`n$Path",
            'CSV Export Complete'
        ) | Out-Null

    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'CSV Export Error'
        ) | Out-Null

    }

})

# ------------------------------------------------------------
# HTML EXPORT
# ------------------------------------------------------------

$btnHTML.Add_Click({

    if ($script:Results.Count -eq 0) {

        [System.Windows.Forms.MessageBox]::Show(
            'No diagnostic results available.',
            'HTML Export'
        ) | Out-Null

        return

    }

    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $Path = Join-Path $script:Desktop (
        "AD_Print_Network_Diagnostics_$Timestamp.html"
    )

    $OpenCount = @(
        $script:Results |
            Where-Object { $_.Status -eq 'OPEN' }
    ).Count

    $FailedCount = @(
        $script:Results |
            Where-Object { $_.Status -eq 'FAILED' }
    ).Count

    $Table = (
        $script:Results |
            ConvertTo-Html -Fragment -Property `
                Timestamp,
                Source,
                Domain,
                Username,
                Target,
                Category,
                Service,
                Protocol,
                Port,
                Status,
                ResponseMS,
                Details
    ) -join "`r`n"

    $Style = @'
<style>
body {
    font-family: Segoe UI, Arial, sans-serif;
    margin: 25px;
    background: #f4f6f9;
    color: #263445;
}
h1 { color: #23466d; }
h2 { color: #23466d; }
.summary {
    padding: 18px;
    background: white;
    border: 1px solid #ddd;
    margin-bottom: 20px;
    line-height: 1.8;
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

    $Html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>AD and Network Printing Diagnostics</title>
$Style
</head>
<body>

<h1>Active Directory & Network Printing Diagnostics</h1>

<div class="summary">

<strong>Source Workstation:</strong> $script:Workstation<br>
<strong>Domain:</strong> $script:Domain<br>
<strong>User:</strong> $script:Username<br>
<strong>Report Generated:</strong> $(Get-Date)<br>
<strong>Successful Tests:</strong> $OpenCount<br>
<strong>Failed Tests:</strong> $FailedCount<br>
<strong>Total Records:</strong> $($script:Results.Count)<br>

</div>

<h2>Diagnostic Results</h2>

$Table

<p>Generated by AD Print Server Network Diagnostic Tool V2.0.</p>

<p>Note: TCP connectivity results do not establish UDP connectivity or successful application-level printing.</p>

</body>
</html>
"@

    try {

        $Html | Out-File `
            -FilePath $Path `
            -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "HTML report exported successfully:`r`n`r`n$Path",
            'HTML Export Complete'
        ) | Out-Null

    }
    catch {

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            'HTML Export Error'
        ) | Out-Null

    }

})

# ------------------------------------------------------------
# FORM CLOSING
# ------------------------------------------------------------

$form.Add_FormClosing({

    if ($script:IsRunning) {

        $Confirmation = [System.Windows.Forms.MessageBox]::Show(
            'A diagnostic is currently running. Do you want to cancel and close?',
            'Diagnostic Running',
            'YesNo',
            'Warning'
        )

        if ($Confirmation -ne 'Yes') {

            $_.Cancel = $true

        }
        else {

            $script:CancelRequested = $true

        }

    }

})

# ------------------------------------------------------------
# DISPLAY GUI
# ------------------------------------------------------------

[void]$form.ShowDialog()

# ------------------------------------------------------------
# END OF SCRIPT
# ------------------------------------------------------------
