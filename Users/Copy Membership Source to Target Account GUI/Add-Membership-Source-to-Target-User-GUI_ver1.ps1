Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Import-Module ActiveDirectory

###########################################################################
# DOMAIN CONFIGURATION
###########################################################################

# Edit Domain Map as needed 
$DomainMap = @{
    "HQ"         = "avenge.com"
    "CCR"        = "superfriends.com"
    "COPPERSYNC" = "skywalker.com"
    "MYIGT"      = "darthmaul.com"
}

$Global:ADCred = $null
$Global:LoadedGroups = @()

###########################################################################
# FORM
###########################################################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "Tier0 AD Group Membership Copier"
$form.Size = New-Object System.Drawing.Size(1200,750)
$form.StartPosition = "CenterScreen"

###########################################################################
# SOURCE
###########################################################################

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source User"
$lblSource.Location = "10,15"
$form.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = "100,10"
$txtSource.Width = 180
$form.Controls.Add($txtSource)

$lblSourceDomain = New-Object System.Windows.Forms.Label
$lblSourceDomain.Text = "Source Domain"
$lblSourceDomain.Location = "300,15"
$form.Controls.Add($lblSourceDomain)

$cboSourceDomain = New-Object System.Windows.Forms.ComboBox
$cboSourceDomain.Location = "400,10"
$cboSourceDomain.Width = 150
$cboSourceDomain.DropDownStyle = "DropDownList"
$cboSourceDomain.Items.AddRange(@(
    "HQ",
    "CCR",
    "COPPERSYNC",
    "MYIGT"
))
$cboSourceDomain.SelectedItem = "MYIGT"
$form.Controls.Add($cboSourceDomain)

###########################################################################
# TARGET
###########################################################################

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target User"
$lblTarget.Location = "10,50"
$form.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = "100,45"
$txtTarget.Width = 180
$form.Controls.Add($txtTarget)

$lblTargetDomain = New-Object System.Windows.Forms.Label
$lblTargetDomain.Text = "Target Domain"
$lblTargetDomain.Location = "300,50"
$form.Controls.Add($lblTargetDomain)

$cboTargetDomain = New-Object System.Windows.Forms.ComboBox
$cboTargetDomain.Location = "400,45"
$cboTargetDomain.Width = 150
$cboTargetDomain.DropDownStyle = "DropDownList"
$cboTargetDomain.Items.AddRange(@(
    "HQ",
    "CCR",
    "COPPERSYNC",
    "MYIGT"
))
$cboTargetDomain.SelectedItem = "MYIGT"
$form.Controls.Add($cboTargetDomain)

###########################################################################
# BUTTONS
###########################################################################

$btnCreds = New-Object System.Windows.Forms.Button
$btnCreds.Text = "Connect Using T0"
$btnCreds.Location = "600,10"
$btnCreds.Width = 150
$form.Controls.Add($btnCreds)

$btnLoadSource = New-Object System.Windows.Forms.Button
$btnLoadSource.Text = "Load Source Groups"
$btnLoadSource.Location = "770,10"
$btnLoadSource.Width = 150
$form.Controls.Add($btnLoadSource)

$btnLoadTarget = New-Object System.Windows.Forms.Button
$btnLoadTarget.Text = "Load Target Groups"
$btnLoadTarget.Location = "940,10"
$btnLoadTarget.Width = 150
$form.Controls.Add($btnLoadTarget)

###########################################################################
# LISTS
###########################################################################

$listSource = New-Object System.Windows.Forms.ListBox
$listSource.Location = "10,90"
$listSource.Size = "550,500"
$listSource.SelectionMode = "MultiExtended"
$form.Controls.Add($listSource)

$listTarget = New-Object System.Windows.Forms.ListBox
$listTarget.Location = "620,90"
$listTarget.Size = "550,500"
$form.Controls.Add($listTarget)

###########################################################################
# OPERATION BUTTONS
###########################################################################

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = "10,610"
$form.Controls.Add($btnSelectAll)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear Selection"
$btnClear.Location = "120,610"
$form.Controls.Add($btnClear)

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Add Selected Groups To Target"
$btnCopy.Location = "280,610"
$btnCopy.Width = 220
$form.Controls.Add($btnCopy)

###########################################################################
# LOG WINDOW
###########################################################################

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = "10,645"
$txtLog.Size = "1160,60"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$form.Controls.Add($txtLog)

###########################################################################
# CREDENTIALS
###########################################################################

$btnCreds.Add_Click({

    $Global:ADCred = Get-Credential `
        -Message "Enter MYIGT Tier0 credentials"

    $txtLog.AppendText(
        "Credential stored in memory.`r`n"
    )
})

###########################################################################
# LOAD SOURCE MEMBERSHIPS
###########################################################################

$btnLoadSource.Add_Click({

    $listSource.Items.Clear()

    $Server = $DomainMap[$cboSourceDomain.Text]

    try {

        $User = Get-ADUser `
            $txtSource.Text `
            -Server $Server `
            -Credential $Global:ADCred `
            -ErrorAction Stop

        $Global:LoadedGroups =
            Get-ADPrincipalGroupMembership `
                $User `
                -Server $Server `
                -Credential $Global:ADCred |
                Sort-Object Name

        foreach($Group in $Global:LoadedGroups)
        {
            [void]$listSource.Items.Add($Group.Name)
        }

        $txtLog.AppendText(
            "Loaded $($Global:LoadedGroups.Count) source groups.`r`n"
        )
    }
    catch {
        $txtLog.AppendText(
            "SOURCE LOAD FAILED: $($_.Exception.Message)`r`n"
        )
    }

})

###########################################################################
# LOAD TARGET MEMBERSHIPS
###########################################################################

$btnLoadTarget.Add_Click({

    $listTarget.Items.Clear()

    $Server = $DomainMap[$cboTargetDomain.Text]

    try {

        Get-ADPrincipalGroupMembership `
            $txtTarget.Text `
            -Server $Server `
            -Credential $Global:ADCred |
            Sort-Object Name |
            ForEach-Object {

                [void]$listTarget.Items.Add($_.Name)
            }

        $txtLog.AppendText(
            "Loaded target memberships.`r`n"
        )

    }
    catch {

        $txtLog.AppendText(
            "TARGET LOAD FAILED: $($_.Exception.Message)`r`n"
        )
    }

})

###########################################################################
# SELECT ALL
###########################################################################

$btnSelectAll.Add_Click({

    for($i = 0; $i -lt $listSource.Items.Count; $i++)
    {
        $listSource.SetSelected($i,$true)
    }

})

###########################################################################
# CLEAR ALL
###########################################################################

$btnClear.Add_Click({

    for($i = 0; $i -lt $listSource.Items.Count; $i++)
    {
        $listSource.SetSelected($i,$false)
    }

})

###########################################################################
# COPY
###########################################################################

$btnCopy.Add_Click({

    $Server = $DomainMap[$cboTargetDomain.Text]

    try {

        $TargetUser = Get-ADUser `
            $txtTarget.Text `
            -Server $Server `
            -Credential $Global:ADCred `
            -ErrorAction Stop

        $ExistingGroups =
            Get-ADPrincipalGroupMembership `
                $TargetUser `
                -Server $Server `
                -Credential $Global:ADCred |
                Select-Object -ExpandProperty DistinguishedName

        foreach($SelectedName in $listSource.SelectedItems)
        {
            $Group =
                $Global:LoadedGroups |
                Where-Object Name -eq $SelectedName |
                Select-Object -First 1

            if(!$Group)
            {
                continue
            }

            if($Group.Name -eq "Domain Users")
            {
                continue
            }

            if($Group.DistinguishedName -in $ExistingGroups)
            {
                $txtLog.AppendText(
                    "[SKIP] $($Group.Name) already assigned.`r`n"
                )

                continue
            }

            try {

                Add-ADGroupMember `
                    -Identity $Group.DistinguishedName `
                    -Members $TargetUser.DistinguishedName `
                    -Server $Server `
                    -Credential $Global:ADCred `
                    -ErrorAction Stop

                $txtLog.AppendText(
                    "[ADDED] $($Group.Name)`r`n"
                )

            }
            catch {

                $txtLog.AppendText(
                    "[FAILED] $($Group.Name)`r`n"
                )

                $txtLog.AppendText(
                    "$($_.Exception.Message)`r`n"
                )
            }
        }

        $txtLog.AppendText(
            "Operation completed.`r`n"
        )

    }
    catch {

        $txtLog.AppendText(
            "Target validation failed: $($_.Exception.Message)`r`n"
        )
    }

})

###########################################################################

[void]$form.ShowDialog()
