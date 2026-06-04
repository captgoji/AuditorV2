# Version : 1.1
# Derniere modif : Fix separateur _ dans nom CSV
$tempCsv    = $env:AUDIT_TEMP_CSV
$CompSys    = Get-CimInstance Win32_ComputerSystem
$DomainName = $CompSys.Domain
$ExportFile = "$tempCsv\Comptes\$env:COMPUTERNAME.$DomainName" + "_Audit_comptes_T0.csv"
$Results    = @()

# Creer le sous-dossier si absent
$null = New-Item -ItemType Directory -Force -Path (Split-Path $ExportFile)

$IsADDSInstalled = (Get-WindowsFeature -Name AD-Domain-Services).InstallState -eq 'Installed'

if ($IsADDSInstalled) {
    # FIX : ajout de userAccountControl pour fiabiliser la detection Enabled
    $Tier0Users = Get-ADUser -Filter { adminCount -eq 1 } `
                  -Properties userAccountControl, Description, PasswordLastSet, MemberOf, WhenCreated `
                  -ErrorAction SilentlyContinue

    foreach ($ADUser in $Tier0Users) {
        try {
            $GroupsAD = ($ADUser.MemberOf | ForEach-Object {
                ($_ -split ",")[0] -replace "CN=", ""
            }) -join " - "

            $DomServiceList = (Get-CimInstance Win32_Service |
                Where-Object { $_.StartName -match $ADUser.SamAccountName } |
                ForEach-Object { $_.Name }) -join " - "

            $DomTaskList = (Get-ScheduledTask |
                Where-Object { $_.Principal.UserId -match $ADUser.SamAccountName } |
                ForEach-Object { $_.TaskName }) -join " - "

            # FIX : lecture directe du bit ACCOUNTDISABLE (bit 2 de userAccountControl)
            $IsEnabled = -not [bool]($ADUser.userAccountControl -band 2)

            $Results += [PSCustomObject]@{
                Ordinateur      = "$env:COMPUTERNAME.$DomainName"
                Name            = $ADUser.SamAccountName
                Type            = "Compte domaine (Tier 0)"
                DateCreation    = $ADUser.WhenCreated
                Active          = $IsEnabled
                Description     = $ADUser.Description
                PasswordLastSet = $ADUser.PasswordLastSet
                Groupes         = $GroupsAD
                Services        = $DomServiceList
                ScheduledTask   = $DomTaskList
            }
        } catch {
            Write-Warning "Impossible de traiter le compte $($ADUser.SamAccountName) : $_"
        }
    }
} else {
    $Users = Get-LocalUser

    foreach ($User in $Users) {
        $Groups = (Get-LocalGroup | ForEach-Object {
            try {
                if (Get-LocalGroupMember $_.Name -Member $User.Name -ErrorAction SilentlyContinue) {
                    $_.Name
                }
            } catch {}
        }) -join " - "

        $ServiceList = (Get-CimInstance Win32_Service |
            Where-Object { $_.StartName -match $User.Name } |
            ForEach-Object { $_.Name }) -join " - "

        $TaskList = (Get-ScheduledTask |
            Where-Object { $_.Principal.UserId -match $User.Name } |
            ForEach-Object { $_.TaskName }) -join " - "

        $Results += [PSCustomObject]@{
            Ordinateur      = "$env:COMPUTERNAME.$DomainName"
            Name            = $User.Name
            Type            = "Compte local"
            Active          = $User.Enabled
            Description     = $User.Description
            PasswordLastSet = $User.PasswordLastSet
            Groupes         = $Groups
            Services        = $ServiceList
            ScheduledTask   = $TaskList
        }
    }
}

$Results | Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8
Write-Host "Rapport genere : $ExportFile"
