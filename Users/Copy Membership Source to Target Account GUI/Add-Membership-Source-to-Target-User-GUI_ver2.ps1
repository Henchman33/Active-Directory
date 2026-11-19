# This uses current Domain, also gives an option to connect to another domain using credentials

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Import-Module ActiveDirectory

$Global:ADCred = $null
$Global:SourceGroups = @()

########################################################################
# FORM
########################################################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Group Membership Copier"
$form.Size = New-Object System.Drawing.Size(1200,800)
$form.StartPosition = "CenterScreen"

########################################################################
# DOMAIN SETTINGS
########################################################################

$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = "Domain"
$lblDomain.Location = "10,15"
$form.Controls.Add($lblDomain)

$txtDomain = New-Object System.Windows.Forms.TextBox
$txtDomain.Location = "80,10"
$txtDomain.Size = "250,25"
$txtDomain.Text = "hq.globalcashaccess.us"
$form.Controls.Add($txtDomain)

$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Text = "Domain Controller"
$lblDC.Location = "350,15"
$form.Controls.Add($lblDC)

$txtDC = New-Object System.Windows.Forms.TextBox
$txtDC.Location = "470,10"
$txtDC.Size = "250,25"
$txtDC.Text = ""
$form.Controls.Add($txtDC)

########################################################################
# CREDENTIALS
########################################################################

$btnCreds = New-Object System.Windows.Forms.Button
$btnCreds.Text = "Set Credentials"
$btnCreds.Location = "740,8"
$btnCreds.Size = "120,30"
$form.Controls.Add($btnCreds)

$lblCreds = New-Object System.Windows.Forms.Label
$lblCreds.Text = "Using Current Session"
$lblCreds.Location = "880,15"
$lblCreds.Size = "250,20"
$form.Controls.Add($lblCreds)

########################################################################
# SOURCE USER
########################################################################

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source User"
$lblSource.Location = "10,55"
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = "100,50"
$txtSource.Width = 200
$form.Controls.Add($txtSource)

########################################################################
# TARGET USER
########################################################################

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target User"
$lblTarget.Location = "330,55"
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = "420,50"
$txtTarget.Width = 200
$form.Controls.Add($txtTarget)

########################################################################
# BUTTONS
########################################################################

$btnLoadSource = New-Object System.Windows.Forms.Button
$btnLoadSource.Text = "Load Source Groups"
$btnLoadSource.Location = "10,95"
$btnLoadSource.Width = 180
$form.Controls.Add($btnLoadSource)

$btnLoadTarget = New-Object System.Windows.Forms.Button
$btnLoadTarget.Text = "Load Target Groups"
$btnLoadTarget.Location = "210,95"
$btnLoadTarget.Width = 180
$form.Controls.Add($btnLoadTarget)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = "420,95"
$form.Controls.Add($btnSelectAll)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear Selection"
$btnClear.Location = "520,95"
$form.Controls.Add($btnClear)

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Add Selected Groups"
$btnCopy.Location = "650,95"
$btnCopy.Width = 200
$form.Controls.Add($btnCopy)

########################################################################
# GROUP LISTS
########################################################################

$listSource = New-Object System.Windows.Forms.ListBox
$listSource.Location = "10,140"
$listSource.Size = "550,500"
$listSource.SelectionMode = "MultiExtended"
$form.Controls.Add($listSource)

$listTarget = New-Object System.Windows.Forms.ListBox
$listTarget.Location = "620,140"
$listTarget.Size = "550,500"
$form.Controls.Add($listTarget)

########################################################################
# LOG
########################################################################

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = "10,660"
$txtLog.Size = "1160,90"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

########################################################################
# HELPER
########################################################################

function Write-Log {
    param($Message)

    $txtLog.AppendText(
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message`r`n"
    )
}

function Get-ADParams {

    $Params = @{}

    if (![string]::IsNullOrWhiteSpace($txtDC.Text))
    {
        $Params["Server"] = $txtDC.Text
    }
    elseif (![string]::IsNullOrWhiteSpace($txtDomain.Text))
    {
        $Params["Server"] = $txtDomain.Text
    }

    if ($Global:ADCred)
    {
        $Params["Credential"] = $Global:ADCred
    }

    return $Params
}

########################################################################
# CREDENTIALS
########################################################################

$btnCreds.Add_Click({

    $Global:ADCred = Get-Credential

    if($Global:ADCred)
    {
        $lblCreds.Text = $Global:ADCred.UserName
        Write-Log "Credential loaded."
    }

})

########################################################################
# LOAD SOURCE GROUPS
########################################################################

$btnLoadSource.Add_Click({

    try {

        $listSource.Items.Clear()
        $Global:SourceGroups = @()

        $Params = Get-ADParams

        $User = Get-ADUser `
            $txtSource.Text `
            @Params `
            -ErrorAction Stop

        $Global:SourceGroups =
            Get-ADPrincipalGroupMembership $User `
            @Params |
            Sort-Object Name

        foreach($Group in $Global:SourceGroups)
        {
            [void]$listSource.Items.Add($Group.Name)
        }

        Write-Log "Loaded $($Global:SourceGroups.Count) source groups."

    }
    catch {

        Write-Log $_.Exception.Message

    }

})

########################################################################
# LOAD TARGET GROUPS
########################################################################

$btnLoadTarget.Add_Click({

    try {

        $listTarget.Items.Clear()

        $Params = Get-ADParams

        $Groups =
            Get-ADPrincipalGroupMembership `
            $txtTarget.Text `
            @Params |
            Sort-Object Name

        foreach($Group in $Groups)
        {
            [void]$listTarget.Items.Add($Group.Name)
        }

        Write-Log "Target groups loaded."

    }
    catch {

        Write-Log $_.Exception.Message

    }

})

########################################################################
# SELECT ALL
########################################################################

$btnSelectAll.Add_Click({

    for($i=0;$i -lt $listSource.Items.Count;$i++)
    {
        $listSource.SetSelected($i,$true)
    }

})

########################################################################
# CLEAR
########################################################################

$btnClear.Add_Click({

    for($i=0;$i -lt $listSource.Items.Count;$i++)
    {
        $listSource.SetSelected($i,$false)
    }

})

########################################################################
# COPY GROUPS
########################################################################

$btnCopy.Add_Click({

    try {

        $Params = Get-ADParams

        $TargetUser = Get-ADUser `
            $txtTarget.Text `
            @Params `
            -ErrorAction Stop

        $ExistingMemberships =
            Get-ADPrincipalGroupMembership `
                $TargetUser `
                @Params |
                Select-Object -ExpandProperty DistinguishedName

        foreach($SelectedGroup in $listSource.SelectedItems)
        {
            $Group =
                $Global:SourceGroups |
                Where-Object Name -eq $SelectedGroup |
                Select-Object -First 1

            if(!$Group)
            {
                continue
            }

            if($Group.Name -eq "Domain Users")
            {
                continue
            }

            if($Group.DistinguishedName -in $ExistingMemberships)
            {
                Write-Log "$($Group.Name) already assigned."
                continue
            }

            try {

                Add-ADGroupMember `
                    -Identity $Group.DistinguishedName `
                    -Members $TargetUser.DistinguishedName `
                    @Params `
                    -ErrorAction Stop

                Write-Log "Added $($Group.Name)"

            }
            catch {

                Write-Log "FAILED: $($Group.Name)"
                Write-Log $_.Exception.Message

            }
        }

        Write-Log "Copy complete."

        $btnLoadTarget.PerformClick()

    }
    catch {

        Write-Log $_.Exception.Message

    }

})

########################################################################

[void]$form.ShowDialog()
