#requires -Version 3.0
#requires -Modules ActiveDirectory

Set-StrictMode -Version 2.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ADCredential = $null
$script:SourceGroups = @()
$script:SourceUser = $null
$script:TargetUser = $null

# Colors and fonts
$script:Colors = @{
    Navy       = [System.Drawing.Color]::FromArgb(24, 35, 52)
    NavyLight  = [System.Drawing.Color]::FromArgb(35, 49, 70)
    Background = [System.Drawing.Color]::FromArgb(242, 245, 249)
    White      = [System.Drawing.Color]::White
    Text       = [System.Drawing.Color]::FromArgb(42, 52, 65)
    Muted      = [System.Drawing.Color]::FromArgb(111, 123, 138)
    Border     = [System.Drawing.Color]::FromArgb(218, 225, 233)
    Blue       = [System.Drawing.Color]::FromArgb(42, 111, 219)
    BlueHover  = [System.Drawing.Color]::FromArgb(31, 91, 184)
    Green      = [System.Drawing.Color]::FromArgb(37, 145, 98)
    Red        = [System.Drawing.Color]::FromArgb(190, 65, 65)
    LogBack    = [System.Drawing.Color]::FromArgb(250, 251, 253)
}
$script:FontRegular = New-Object System.Drawing.Font('Segoe UI', 9)
$script:FontSmall = New-Object System.Drawing.Font('Segoe UI', 8)
$script:FontBold = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$script:FontTitle = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
$script:FontHeader = New-Object System.Drawing.Font('Segoe UI Semibold', 10)

# Main window
$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = 'AD Group Membership Copier'
$script:Form.StartPosition = 'CenterScreen'
$script:Form.Size = New-Object System.Drawing.Size(1180, 830)
$script:Form.MinimumSize = New-Object System.Drawing.Size(980, 700)
$script:Form.BackColor = $script:Colors.Background
$script:Form.Font = $script:FontRegular
$script:Form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

# Header
$script:Header = New-Object System.Windows.Forms.Panel
$script:Header.Location = New-Object System.Drawing.Point(0, 0)
$script:Header.Size = New-Object System.Drawing.Size(1164, 82)
$script:Header.Anchor = 'Top,Left,Right'
$script:Header.BackColor = $script:Colors.Navy
$script:Form.Controls.Add($script:Header)

$script:TitleLabel = New-Object System.Windows.Forms.Label
$script:TitleLabel.Text = 'Active Directory'
$script:TitleLabel.Location = New-Object System.Drawing.Point(24, 13)
$script:TitleLabel.Size = New-Object System.Drawing.Size(350, 30)
$script:TitleLabel.ForeColor = $script:Colors.White
$script:TitleLabel.Font = $script:FontTitle
$script:Header.Controls.Add($script:TitleLabel)

$script:SubtitleLabel = New-Object System.Windows.Forms.Label
$script:SubtitleLabel.Text = 'GROUP MEMBERSHIP COPIER'
$script:SubtitleLabel.Location = New-Object System.Drawing.Point(27, 48)
$script:SubtitleLabel.Size = New-Object System.Drawing.Size(350, 20)
$script:SubtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(180, 195, 214)
$script:SubtitleLabel.Font = $script:FontSmall
$script:Header.Controls.Add($script:SubtitleLabel)

# Connection panel
$script:ConnectionPanel = New-Object System.Windows.Forms.Panel
$script:ConnectionPanel.Location = New-Object System.Drawing.Point(20, 98)
$script:ConnectionPanel.Size = New-Object System.Drawing.Size(1124, 92)
$script:ConnectionPanel.Anchor = 'Top,Left,Right'
$script:ConnectionPanel.BackColor = $script:Colors.White
$script:ConnectionPanel.BorderStyle = 'FixedSingle'
$script:Form.Controls.Add($script:ConnectionPanel)

$script:ConnectionHeading = New-Object System.Windows.Forms.Label
$script:ConnectionHeading.Text = 'CONNECTION'
$script:ConnectionHeading.Location = New-Object System.Drawing.Point(16, 9)
$script:ConnectionHeading.Size = New-Object System.Drawing.Size(180, 20)
$script:ConnectionHeading.Font = $script:FontHeader
$script:ConnectionHeading.ForeColor = $script:Colors.Text
$script:ConnectionPanel.Controls.Add($script:ConnectionHeading)

$script:DomainLabel = New-Object System.Windows.Forms.Label
$script:DomainLabel.Text = 'Domain'
$script:DomainLabel.Location = New-Object System.Drawing.Point(17, 43)
$script:DomainLabel.Size = New-Object System.Drawing.Size(55, 22)
$script:DomainLabel.ForeColor = $script:Colors.Muted
$script:ConnectionPanel.Controls.Add($script:DomainLabel)

$script:DomainText = New-Object System.Windows.Forms.TextBox
$script:DomainText.Location = New-Object System.Drawing.Point(75, 39)
$script:DomainText.Size = New-Object System.Drawing.Size(245, 25)
$script:DomainText.Text = 'hq.globalcashaccess.us'
$script:ConnectionPanel.Controls.Add($script:DomainText)

$script:DCLabel = New-Object System.Windows.Forms.Label
$script:DCLabel.Text = 'Domain controller'
$script:DCLabel.Location = New-Object System.Drawing.Point(342, 43)
$script:DCLabel.Size = New-Object System.Drawing.Size(105, 22)
$script:DCLabel.ForeColor = $script:Colors.Muted
$script:ConnectionPanel.Controls.Add($script:DCLabel)

$script:DCText = New-Object System.Windows.Forms.TextBox
$script:DCText.Location = New-Object System.Drawing.Point(451, 39)
$script:DCText.Size = New-Object System.Drawing.Size(220, 25)
$script:ConnectionPanel.Controls.Add($script:DCText)

$script:CredentialButton = New-Object System.Windows.Forms.Button
$script:CredentialButton.Text = 'Set credentials'
$script:CredentialButton.Location = New-Object System.Drawing.Point(695, 36)
$script:CredentialButton.Size = New-Object System.Drawing.Size(135, 31)
$script:CredentialButton.FlatStyle = 'Flat'
$script:CredentialButton.FlatAppearance.BorderColor = $script:Colors.Border
$script:CredentialButton.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$script:CredentialButton.ForeColor = $script:Colors.Text
$script:ConnectionPanel.Controls.Add($script:CredentialButton)

$script:CredentialStatus = New-Object System.Windows.Forms.Label
$script:CredentialStatus.Text = 'Using current session'
$script:CredentialStatus.Location = New-Object System.Drawing.Point(844, 43)
$script:CredentialStatus.Size = New-Object System.Drawing.Size(255, 22)
$script:CredentialStatus.ForeColor = $script:Colors.Muted
$script:ConnectionPanel.Controls.Add($script:CredentialStatus)

# User selection panel
$script:UserPanel = New-Object System.Windows.Forms.Panel
$script:UserPanel.Location = New-Object System.Drawing.Point(20, 202)
$script:UserPanel.Size = New-Object System.Drawing.Size(1124, 94)
$script:UserPanel.Anchor = 'Top,Left,Right'
$script:UserPanel.BackColor = $script:Colors.White
$script:UserPanel.BorderStyle = 'FixedSingle'
$script:Form.Controls.Add($script:UserPanel)

$script:UserHeading = New-Object System.Windows.Forms.Label
$script:UserHeading.Text = 'USER ACCOUNTS'
$script:UserHeading.Location = New-Object System.Drawing.Point(16, 9)
$script:UserHeading.Size = New-Object System.Drawing.Size(180, 20)
$script:UserHeading.Font = $script:FontHeader
$script:UserHeading.ForeColor = $script:Colors.Text
$script:UserPanel.Controls.Add($script:UserHeading)

$script:SourceLabel = New-Object System.Windows.Forms.Label
$script:SourceLabel.Text = 'Copy memberships from'
$script:SourceLabel.Location = New-Object System.Drawing.Point(17, 45)
$script:SourceLabel.Size = New-Object System.Drawing.Size(145, 22)
$script:SourceLabel.ForeColor = $script:Colors.Muted
$script:UserPanel.Controls.Add($script:SourceLabel)

$script:SourceText = New-Object System.Windows.Forms.TextBox
$script:SourceText.Location = New-Object System.Drawing.Point(165, 41)
$script:SourceText.Size = New-Object System.Drawing.Size(285, 25)
$script:UserPanel.Controls.Add($script:SourceText)

$script:TargetLabel = New-Object System.Windows.Forms.Label
$script:TargetLabel.Text = 'Add memberships to'
$script:TargetLabel.Location = New-Object System.Drawing.Point(480, 45)
$script:TargetLabel.Size = New-Object System.Drawing.Size(135, 22)
$script:TargetLabel.ForeColor = $script:Colors.Muted
$script:UserPanel.Controls.Add($script:TargetLabel)

$script:TargetText = New-Object System.Windows.Forms.TextBox
$script:TargetText.Location = New-Object System.Drawing.Point(620, 41)
$script:TargetText.Size = New-Object System.Drawing.Size(300, 25)
$script:UserPanel.Controls.Add($script:TargetText)

# Action bar
$script:ActionPanel = New-Object System.Windows.Forms.Panel
$script:ActionPanel.Location = New-Object System.Drawing.Point(20, 308)
$script:ActionPanel.Size = New-Object System.Drawing.Size(1124, 48)
$script:ActionPanel.Anchor = 'Top,Left,Right'
$script:ActionPanel.BackColor = $script:Colors.Background
$script:Form.Controls.Add($script:ActionPanel)

$script:LoadSourceButton = New-Object System.Windows.Forms.Button
$script:LoadSourceButton.Text = 'Load source groups'
$script:LoadSourceButton.Location = New-Object System.Drawing.Point(0, 5)
$script:LoadSourceButton.Size = New-Object System.Drawing.Size(155, 34)
$script:LoadSourceButton.FlatStyle = 'Flat'
$script:LoadSourceButton.FlatAppearance.BorderColor = $script:Colors.Border
$script:LoadSourceButton.BackColor = $script:Colors.White
$script:LoadSourceButton.ForeColor = $script:Colors.Text
$script:ActionPanel.Controls.Add($script:LoadSourceButton)

$script:LoadTargetButton = New-Object System.Windows.Forms.Button
$script:LoadTargetButton.Text = 'Load target groups'
$script:LoadTargetButton.Location = New-Object System.Drawing.Point(165, 5)
$script:LoadTargetButton.Size = New-Object System.Drawing.Size(155, 34)
$script:LoadTargetButton.FlatStyle = 'Flat'
$script:LoadTargetButton.FlatAppearance.BorderColor = $script:Colors.Border
$script:LoadTargetButton.BackColor = $script:Colors.White
$script:LoadTargetButton.ForeColor = $script:Colors.Text
$script:ActionPanel.Controls.Add($script:LoadTargetButton)

$script:SelectAllButton = New-Object System.Windows.Forms.Button
$script:SelectAllButton.Text = 'Select all'
$script:SelectAllButton.Location = New-Object System.Drawing.Point(350, 5)
$script:SelectAllButton.Size = New-Object System.Drawing.Size(105, 34)
$script:SelectAllButton.FlatStyle = 'Flat'
$script:SelectAllButton.FlatAppearance.BorderColor = $script:Colors.Border
$script:SelectAllButton.BackColor = $script:Colors.White
$script:SelectAllButton.ForeColor = $script:Colors.Text
$script:ActionPanel.Controls.Add($script:SelectAllButton)

$script:ClearButton = New-Object System.Windows.Forms.Button
$script:ClearButton.Text = 'Clear selection'
$script:ClearButton.Location = New-Object System.Drawing.Point(465, 5)
$script:ClearButton.Size = New-Object System.Drawing.Size(125, 34)
$script:ClearButton.FlatStyle = 'Flat'
$script:ClearButton.FlatAppearance.BorderColor = $script:Colors.Border
$script:ClearButton.BackColor = $script:Colors.White
$script:ClearButton.ForeColor = $script:Colors.Text
$script:ActionPanel.Controls.Add($script:ClearButton)

$script:CopyButton = New-Object System.Windows.Forms.Button
$script:CopyButton.Text = 'Add selected groups'
$script:CopyButton.Location = New-Object System.Drawing.Point(610, 5)
$script:CopyButton.Size = New-Object System.Drawing.Size(190, 34)
$script:CopyButton.FlatStyle = 'Flat'
$script:CopyButton.FlatAppearance.BorderSize = 0
$script:CopyButton.BackColor = $script:Colors.Blue
$script:CopyButton.ForeColor = $script:Colors.White
$script:CopyButton.Font = $script:FontBold
$script:ActionPanel.Controls.Add($script:CopyButton)

$script:SelectionCount = New-Object System.Windows.Forms.Label
$script:SelectionCount.Text = '0 selected'
$script:SelectionCount.Location = New-Object System.Drawing.Point(820, 12)
$script:SelectionCount.Size = New-Object System.Drawing.Size(170, 22)
$script:SelectionCount.ForeColor = $script:Colors.Muted
$script:ActionPanel.Controls.Add($script:SelectionCount)

# Group list panels
$script:SourcePanel = New-Object System.Windows.Forms.Panel
$script:SourcePanel.Location = New-Object System.Drawing.Point(20, 366)
$script:SourcePanel.Size = New-Object System.Drawing.Size(550, 260)
$script:SourcePanel.Anchor = 'Top,Bottom,Left'
$script:SourcePanel.BackColor = $script:Colors.White
$script:SourcePanel.BorderStyle = 'FixedSingle'
$script:Form.Controls.Add($script:SourcePanel)

$script:TargetPanel = New-Object System.Windows.Forms.Panel
$script:TargetPanel.Location = New-Object System.Drawing.Point(594, 366)
$script:TargetPanel.Size = New-Object System.Drawing.Size(550, 260)
$script:TargetPanel.Anchor = 'Top,Bottom,Left,Right'
$script:TargetPanel.BackColor = $script:Colors.White
$script:TargetPanel.BorderStyle = 'FixedSingle'
$script:Form.Controls.Add($script:TargetPanel)

$script:SourceHeading = New-Object System.Windows.Forms.Label
$script:SourceHeading.Text = 'SOURCE GROUPS'
$script:SourceHeading.Location = New-Object System.Drawing.Point(15, 10)
$script:SourceHeading.Size = New-Object System.Drawing.Size(240, 22)
$script:SourceHeading.Font = $script:FontHeader
$script:SourceHeading.ForeColor = $script:Colors.Text
$script:SourcePanel.Controls.Add($script:SourceHeading)

$script:SourceCount = New-Object System.Windows.Forms.Label
$script:SourceCount.Text = '0 groups'
$script:SourceCount.Location = New-Object System.Drawing.Point(420, 11)
$script:SourceCount.Size = New-Object System.Drawing.Size(110, 20)
$script:SourceCount.TextAlign = 'MiddleRight'
$script:SourceCount.ForeColor = $script:Colors.Muted
$script:SourcePanel.Controls.Add($script:SourceCount)

$script:SourceSearch = New-Object System.Windows.Forms.TextBox
$script:SourceSearch.Location = New-Object System.Drawing.Point(15, 40)
$script:SourceSearch.Size = New-Object System.Drawing.Size(515, 25)
$script:SourceSearch.Anchor = 'Top,Left,Right'
$script:SourceSearch.ForeColor = $script:Colors.Text
$script:SourcePanel.Controls.Add($script:SourceSearch)

$script:SourceList = New-Object System.Windows.Forms.ListBox
$script:SourceList.Location = New-Object System.Drawing.Point(15, 76)
$script:SourceList.Size = New-Object System.Drawing.Size(515, 168)
$script:SourceList.Anchor = 'Top,Bottom,Left,Right'
$script:SourceList.SelectionMode = 'MultiExtended'
$script:SourceList.BorderStyle = 'FixedSingle'
$script:SourceList.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
$script:SourceList.ForeColor = $script:Colors.Text
$script:SourcePanel.Controls.Add($script:SourceList)

$script:TargetHeading = New-Object System.Windows.Forms.Label
$script:TargetHeading.Text = 'TARGET GROUPS'
$script:TargetHeading.Location = New-Object System.Drawing.Point(15, 10)
$script:TargetHeading.Size = New-Object System.Drawing.Size(240, 22)
$script:TargetHeading.Font = $script:FontHeader
$script:TargetHeading.ForeColor = $script:Colors.Text
$script:TargetPanel.Controls.Add($script:TargetHeading)

$script:TargetCount = New-Object System.Windows.Forms.Label
$script:TargetCount.Text = '0 groups'
$script:TargetCount.Location = New-Object System.Drawing.Point(420, 11)
$script:TargetCount.Size = New-Object System.Drawing.Size(110, 20)
$script:TargetCount.TextAlign = 'MiddleRight'
$script:TargetCount.ForeColor = $script:Colors.Muted
$script:TargetPanel.Controls.Add($script:TargetCount)

$script:TargetSearch = New-Object System.Windows.Forms.TextBox
$script:TargetSearch.Location = New-Object System.Drawing.Point(15, 40)
$script:TargetSearch.Size = New-Object System.Drawing.Size(515, 25)
$script:TargetSearch.Anchor = 'Top,Left,Right'
$script:TargetSearch.ForeColor = $script:Colors.Text
$script:TargetPanel.Controls.Add($script:TargetSearch)

$script:TargetList = New-Object System.Windows.Forms.ListBox
$script:TargetList.Location = New-Object System.Drawing.Point(15, 76)
$script:TargetList.Size = New-Object System.Drawing.Size(515, 168)
$script:TargetList.Anchor = 'Top,Bottom,Left,Right'
$script:TargetList.SelectionMode = 'One'
$script:TargetList.BorderStyle = 'FixedSingle'
$script:TargetList.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
$script:TargetList.ForeColor = $script:Colors.Text
$script:TargetPanel.Controls.Add($script:TargetList)

# Activity log
$script:LogPanel = New-Object System.Windows.Forms.Panel
$script:LogPanel.Location = New-Object System.Drawing.Point(20, 638)
$script:LogPanel.Size = New-Object System.Drawing.Size(1124, 137)
$script:LogPanel.Anchor = 'Bottom,Left,Right'
$script:LogPanel.BackColor = $script:Colors.White
$script:LogPanel.BorderStyle = 'FixedSingle'
$script:Form.Controls.Add($script:LogPanel)

$script:LogHeading = New-Object System.Windows.Forms.Label
$script:LogHeading.Text = 'ACTIVITY'
$script:LogHeading.Location = New-Object System.Drawing.Point(15, 8)
$script:LogHeading.Size = New-Object System.Drawing.Size(150, 20)
$script:LogHeading.Font = $script:FontHeader
$script:LogHeading.ForeColor = $script:Colors.Text
$script:LogPanel.Controls.Add($script:LogHeading)

$script:ClearLogButton = New-Object System.Windows.Forms.Button
$script:ClearLogButton.Text = 'Clear log'
$script:ClearLogButton.Location = New-Object System.Drawing.Point(1035, 5)
$script:ClearLogButton.Size = New-Object System.Drawing.Size(75, 25)
$script:ClearLogButton.Anchor = 'Top,Right'
$script:ClearLogButton.FlatStyle = 'Flat'
$script:ClearLogButton.FlatAppearance.BorderColor = $script:Colors.Border
$script:ClearLogButton.BackColor = $script:Colors.White
$script:ClearLogButton.ForeColor = $script:Colors.Muted
$script:LogPanel.Controls.Add($script:ClearLogButton)

$script:LogText = New-Object System.Windows.Forms.RichTextBox
$script:LogText.Location = New-Object System.Drawing.Point(15, 34)
$script:LogText.Size = New-Object System.Drawing.Size(1095, 88)
$script:LogText.Anchor = 'Top,Bottom,Left,Right'
$script:LogText.ReadOnly = $true
$script:LogText.BorderStyle = 'FixedSingle'
$script:LogText.BackColor = $script:Colors.LogBack
$script:LogText.ForeColor = $script:Colors.Text
$script:LogText.Font = New-Object System.Drawing.Font('Consolas', 8)
$script:LogPanel.Controls.Add($script:LogText)

# Shared helper scriptblocks keep this a single-file script.
$script:WriteLog = {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = '{0}  [{1}]  {2}{3}' -f $stamp, $Level.ToUpperInvariant(), $Message,
        [Environment]::NewLine

    switch ($Level) {
        'Success' { $color = $script:Colors.Green }
        'Warning' { $color = [System.Drawing.Color]::FromArgb(170, 115, 30) }
        'Error'   { $color = $script:Colors.Red }
        default   { $color = $script:Colors.Text }
    }

    $script:LogText.SelectionStart = $script:LogText.TextLength
    $script:LogText.SelectionLength = 0
    $script:LogText.SelectionColor = $color
    $script:LogText.AppendText($line)
    $script:LogText.SelectionColor = $script:LogText.ForeColor
    $script:LogText.ScrollToCaret()
}

$script:GetADParams = {
    $parameters = @{}

    if (-not [string]::IsNullOrWhiteSpace($script:DCText.Text)) {
        $parameters['Server'] = $script:DCText.Text.Trim()
    }
    elseif (-not [string]::IsNullOrWhiteSpace($script:DomainText.Text)) {
        $parameters['Server'] = $script:DomainText.Text.Trim()
    }

    if ($null -ne $script:ADCredential) {
        $parameters['Credential'] = $script:ADCredential
    }

    return $parameters
}

$script:RefreshSourceList = {
    $script:SourceList.BeginUpdate()
    $script:SourceList.Items.Clear()

    $searchText = $script:SourceSearch.Text.Trim()
    foreach ($group in $script:SourceGroups) {
        if (($searchText.Length -eq 0) -or
            ($group.Name.IndexOf($searchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) {
            [void]$script:SourceList.Items.Add($group.Name)
        }
    }

    $script:SourceList.EndUpdate()
    $script:SourceCount.Text = '{0} groups' -f $script:SourceGroups.Count
    $script:SelectionCount.Text = '{0} selected' -f $script:SourceList.SelectedItems.Count
}

$script:RefreshTargetList = {
    $script:TargetList.BeginUpdate()
    $script:TargetList.Items.Clear()

    $searchText = $script:TargetSearch.Text.Trim()
    foreach ($groupName in $script:TargetGroupNames) {
        if (($searchText.Length -eq 0) -or
            ($groupName.IndexOf($searchText, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)) {
            [void]$script:TargetList.Items.Add($groupName)
        }
    }

    $script:TargetList.EndUpdate()
    $script:TargetCount.Text = '{0} groups' -f $script:TargetGroupNames.Count
}

$script:TargetGroupNames = @()

# Credentials
$script:CredentialButton.Add_Click({
    try {
        $credential = Get-Credential -Message 'Enter credentials for Active Directory'
        if ($null -ne $credential) {
            $script:ADCredential = $credential
            $script:CredentialStatus.Text = $credential.UserName
            $script:CredentialStatus.ForeColor = $script:Colors.Green
            & $script:WriteLog 'Alternate credentials loaded.' 'Success'
        }
    }
    catch {
        & $script:WriteLog $_.Exception.Message 'Error'
    }
})

# Load source memberships
$script:LoadSourceButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:SourceText.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Enter a source user account name.',
            'Source user required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $script:LoadSourceButton.Enabled = $false
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        $parameters = & $script:GetADParams
        $userParameters = @{
            Identity    = $script:SourceText.Text.Trim()
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $userParameters[$key] = $parameters[$key]
        }

        $user = Get-ADUser @userParameters
        $groupParameters = @{
            Identity    = $user
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $groupParameters[$key] = $parameters[$key]
        }

        $groups = @(Get-ADPrincipalGroupMembership @groupParameters | Sort-Object -Property Name)
        $script:SourceUser = $user
        $script:SourceGroups = $groups
        & $script:RefreshSourceList
        & $script:WriteLog ('Loaded {0} groups for {1}.' -f $groups.Count, $user.SamAccountName) 'Success'
    }
    catch {
        & $script:WriteLog ('Could not load source groups: {0}' -f $_.Exception.Message) 'Error'
    }
    finally {
        $script:LoadSourceButton.Enabled = $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Load target memberships
$script:LoadTargetButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($script:TargetText.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Enter a target user account name.',
            'Target user required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $script:LoadTargetButton.Enabled = $false
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        $parameters = & $script:GetADParams
        $userParameters = @{
            Identity    = $script:TargetText.Text.Trim()
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $userParameters[$key] = $parameters[$key]
        }

        $user = Get-ADUser @userParameters
        $groupParameters = @{
            Identity    = $user
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $groupParameters[$key] = $parameters[$key]
        }

        $groups = @(Get-ADPrincipalGroupMembership @groupParameters | Sort-Object -Property Name)
        $script:TargetUser = $user
        $script:TargetGroupNames = @($groups | ForEach-Object { $_.Name })
        & $script:RefreshTargetList
        & $script:WriteLog ('Loaded {0} groups for {1}.' -f $groups.Count, $user.SamAccountName) 'Success'
    }
    catch {
        & $script:WriteLog ('Could not load target groups: {0}' -f $_.Exception.Message) 'Error'
    }
    finally {
        $script:LoadTargetButton.Enabled = $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Group list filtering and selection
$script:SourceSearch.Add_TextChanged({
    & $script:RefreshSourceList
})
$script:TargetSearch.Add_TextChanged({
    & $script:RefreshTargetList
})
$script:SourceList.Add_SelectedIndexChanged({
    $script:SelectionCount.Text = '{0} selected' -f $script:SourceList.SelectedItems.Count
})

$script:SelectAllButton.Add_Click({
    for ($index = 0; $index -lt $script:SourceList.Items.Count; $index++) {
        $script:SourceList.SetSelected($index, $true)
    }
    $script:SelectionCount.Text = '{0} selected' -f $script:SourceList.SelectedItems.Count
})

$script:ClearButton.Add_Click({
    $script:SourceList.ClearSelected()
    $script:SelectionCount.Text = '0 selected'
})

$script:ClearLogButton.Add_Click({
    $script:LogText.Clear()
})

# Add selected memberships to target
$script:CopyButton.Add_Click({
    if ($null -eq $script:SourceUser -or $script:SourceGroups.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            'Load the source user groups before copying memberships.',
            'Source groups not loaded',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($script:TargetText.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Enter a target user account name.',
            'Target user required',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    if ($script:SourceList.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            'Select at least one source group.',
            'No groups selected',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    $selectedNames = New-Object 'System.Collections.Generic.List[string]'
    foreach ($selectedItem in $script:SourceList.SelectedItems) {
        $selectedNames.Add([string]$selectedItem)
    }

    $confirmationText = "Add $($selectedNames.Count) selected group membership(s) to '$($script:TargetText.Text.Trim())'?"
    $answer = [System.Windows.Forms.MessageBox]::Show(
        $confirmationText,
        'Confirm group membership changes',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    $script:CopyButton.Enabled = $false
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        $parameters = & $script:GetADParams
        $targetParameters = @{
            Identity    = $script:TargetText.Text.Trim()
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $targetParameters[$key] = $parameters[$key]
        }

        $targetUser = Get-ADUser @targetParameters
        $existingParameters = @{
            Identity    = $targetUser
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $existingParameters[$key] = $parameters[$key]
        }

        $existingGroups = @(Get-ADPrincipalGroupMembership @existingParameters)
        $existingDNs = @{}
        foreach ($existingGroup in $existingGroups) {
            $existingDNs[$existingGroup.DistinguishedName] = $true
        }

        $addedCount = 0
        $skippedCount = 0
        $failedCount = 0

        foreach ($group in $script:SourceGroups) {
            if (-not $selectedNames.Contains([string]$group.Name)) {
                continue
            }

            if ($group.Name -eq 'Domain Users') {
                $skippedCount++
                & $script:WriteLog 'Skipped Domain Users. This is the account primary group.' 'Warning'
                continue
            }

            if ($existingDNs.ContainsKey($group.DistinguishedName)) {
                $skippedCount++
                & $script:WriteLog ('{0} is already assigned.' -f $group.Name) 'Warning'
                continue
            }

            try {
                $addParameters = @{
                    Identity    = $group.DistinguishedName
                    Members     = $targetUser.DistinguishedName
                    ErrorAction = 'Stop'
                }
                foreach ($key in $parameters.Keys) {
                    $addParameters[$key] = $parameters[$key]
                }

                Add-ADGroupMember @addParameters
                $addedCount++
                $existingDNs[$group.DistinguishedName] = $true
                & $script:WriteLog ('Added {0}.' -f $group.Name) 'Success'
            }
            catch {
                $failedCount++
                & $script:WriteLog ('Failed to add {0}: {1}' -f $group.Name, $_.Exception.Message) 'Error'
            }
        }

        & $script:WriteLog ('Finished. Added: {0}; already assigned or skipped: {1}; failed: {2}.' -f
            $addedCount, $skippedCount, $failedCount) 'Info'

        # Refresh target memberships after the changes.
        $refreshParameters = @{
            Identity    = $targetUser
            ErrorAction = 'Stop'
        }
        foreach ($key in $parameters.Keys) {
            $refreshParameters[$key] = $parameters[$key]
        }
        $updatedGroups = @(Get-ADPrincipalGroupMembership @refreshParameters | Sort-Object -Property Name)
        $script:TargetUser = $targetUser
        $script:TargetGroupNames = @($updatedGroups | ForEach-Object { $_.Name })
        & $script:RefreshTargetList
    }
    catch {
        & $script:WriteLog ('Copy operation stopped: {0}' -f $_.Exception.Message) 'Error'
    }
    finally {
        $script:CopyButton.Enabled = $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:Form.Add_Shown({
    $script:Form.Activate()
    & $script:WriteLog 'Ready. Enter the source and target accounts to begin.' 'Info'
})

$script:Form.Add_FormClosed({
    $script:FontRegular.Dispose()
    $script:FontSmall.Dispose()
    $script:FontBold.Dispose()
    $script:FontTitle.Dispose()
    $script:FontHeader.Dispose()
})

[void]$script:Form.ShowDialog()
