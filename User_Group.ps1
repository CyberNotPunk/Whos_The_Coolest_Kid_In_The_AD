Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


if (-not (Get-Module -Name ActiveDirectory)) {
    Import-Module ActiveDirectory -ErrorAction Stop
}

function Get-AllADUsers {
    return Get-ADUser -Filter * -Properties Name, sAMAccountName | Select-Object Name, sAMAccountName | Sort-Object Name
}


function Get-UserGroups {
    param([string]$sAMAccountName)
    try {
        $user = Get-ADUser $sAMAccountName -Properties MemberOf
        return $user.MemberOf | ForEach-Object {
            (Get-ADGroup $_).Name
        }
    }
    catch {
        Write-Output "Erreur : Impossible de trouver l'utilisateur ou de récupérer ses groupes."
        return @()
    }
}


function Add-GroupsToUser {
    param([string]$sAMAccountName, [array]$Groups)
    foreach ($group in $Groups) {
        try {
            Add-ADGroupMember -Identity $group -Members $sAMAccountName -ErrorAction Stop
            Write-Output "Ajout de $sAMAccountName au groupe $group : OK"
        }
        catch {
            Write-Output "Erreur lors de l'ajout de $sAMAccountName au groupe $group : $_"
        }
    }
}


$form = New-Object System.Windows.Forms.Form
$form.Text = "Transfert d'appartenances de groupes AD"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"


$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Location = New-Object System.Drawing.Point(20, 20)
$labelSource.Size = New-Object System.Drawing.Size(200, 20)
$labelSource.Text = "Utilisateur source :"
$form.Controls.Add($labelSource)

$comboSource = New-Object System.Windows.Forms.ComboBox
$comboSource.Location = New-Object System.Drawing.Point(20, 50)
$comboSource.Size = New-Object System.Drawing.Size(250, 20)
$comboSource.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboSource)


$labelTarget = New-Object System.Windows.Forms.Label
$labelTarget.Location = New-Object System.Drawing.Point(20, 90)
$labelTarget.Size = New-Object System.Drawing.Size(200, 20)
$labelTarget.Text = "Utilisateur cible :"
$form.Controls.Add($labelTarget)

$comboTarget = New-Object System.Windows.Forms.ComboBox
$comboTarget.Location = New-Object System.Drawing.Point(20, 120)
$comboTarget.Size = New-Object System.Drawing.Size(250, 20)
$comboTarget.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$form.Controls.Add($comboTarget)


$buttonShowGroups = New-Object System.Windows.Forms.Button
$buttonShowGroups.Location = New-Object System.Drawing.Point(20, 160)
$buttonShowGroups.Size = New-Object System.Drawing.Size(250, 30)
$buttonShowGroups.Text = "Afficher les groupes de l'utilisateur source"
$form.Controls.Add($buttonShowGroups)


$textBoxGroups = New-Object System.Windows.Forms.TextBox
$textBoxGroups.Location = New-Object System.Drawing.Point(20, 200)
$textBoxGroups.Size = New-Object System.Drawing.Size(550, 200)
$textBoxGroups.Multiline = $true
$textBoxGroups.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$form.Controls.Add($textBoxGroups)


$buttonTransfer = New-Object System.Windows.Forms.Button
$buttonTransfer.Location = New-Object System.Drawing.Point(20, 420)
$buttonTransfer.Size = New-Object System.Drawing.Size(250, 30)
$buttonTransfer.Text = "Transférer les groupes vers l'utilisateur cible"
$buttonTransfer.Enabled = $false
$form.Controls.Add($buttonTransfer)


$users = Get-AllADUsers
foreach ($user in $users) {
    $comboSource.Items.Add($user.Name)
    $comboTarget.Items.Add($user.Name)
}


$userMap = @{}
foreach ($user in $users) {
    $userMap[$user.Name] = $user.sAMAccountName
}


$buttonShowGroups.Add_Click({
    $sourceUserName = $comboSource.SelectedItem
    if ($sourceUserName) {
        $sAMAccountName = $userMap[$sourceUserName]
        $groups = Get-UserGroups $sAMAccountName
        $textBoxGroups.Text = $groups -join "`r`n"
        $buttonTransfer.Enabled = ($groups.Count -gt 0)
    }
})


$buttonTransfer.Add_Click({
    $sourceUserName = $comboSource.SelectedItem
    $targetUserName = $comboTarget.SelectedItem
    if ($sourceUserName -and $targetUserName) {
        $sourceSAM = $userMap[$sourceUserName]
        $targetSAM = $userMap[$targetUserName]
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Voulez-vous vraiment transférer les groupes de $sourceUserName ($sourceSAM) vers $targetUserName ($targetSAM) ?",
            "Confirmation",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -eq "Yes") {
            $groups = Get-UserGroups $sourceSAM
            Add-GroupsToUser $targetSAM $groups
            [System.Windows.Forms.MessageBox]::Show(
                "Transfert terminé !",
                "Succès",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
})

$form.ShowDialog()
