# Version : 1.1
# Derniere modif : Exit 0 propre si RSAT absent
$tempCsv    = $env:AUDIT_TEMP_CSV
$CompSys    = Get-CimInstance Win32_ComputerSystem
$DomainName = $CompSys.Domain
$ExportFile = "$tempCsv\Comptes\$env:COMPUTERNAME.$DomainName" + "Audit_GPO.csv"

$null = New-Item -ItemType Directory -Force -Path (Split-Path $ExportFile)

if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
    Write-Host "Module GroupPolicy introuvable - script ignore (necessite RSAT/GPMC sur un DC ou membre de domaine)." -ForegroundColor Yellow
    exit 0
}
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "Module ActiveDirectory introuvable - script ignore (necessite RSAT sur un DC ou membre de domaine)." -ForegroundColor Yellow
    exit 0
}
Import-Module GroupPolicy     -ErrorAction Stop
Import-Module ActiveDirectory -ErrorAction Stop

# Recupere le createur d'une GPO
# Methode 1 : XML report (owner du SecurityDescriptor)
# Methode 2 : fallback via ACL SYSVOL (owner du dossier GPO)
function Get-GPOCreator {
    param([string]$GpoId, [string]$Domain)
    try {
        $report = [xml](Get-GPOReport -Guid $GpoId -ReportType Xml -Domain $Domain -ErrorAction Stop)

        # Essayer les differents chemins selon la version de l'OS
        $owner = $report.GPO.SecurityDescriptor.Owner.Name.'#text'
        if (-not $owner) { $owner = $report.GPO.SecurityDescriptor.Owner.'#text' }
        if (-not $owner) { $owner = $report.GPO.SecurityDescriptor.Owner.Name }
        if ($owner) { return $owner }
    } catch {}

    # Fallback : lire l'owner du dossier GPO dans SYSVOL
    try {
        $dnsDomain = (Get-ADDomain -Identity $Domain).DNSRoot
        $gpoPath   = "\\$dnsDomain\SYSVOL\$dnsDomain\Policies\{$GpoId}"
        $owner     = (Get-Acl -Path $gpoPath -ErrorAction Stop).Owner
        if ($owner) {
            # Retourner uniquement le nom sans le domaine (DOMAIN\User -> User)
            return ($owner -split '\\')[-1]
        }
    } catch {}

    return ""
}

# Recupere les OUs liees a une GPO - retourne uniquement le nom court de l'OU
function Get-GPOLinkedOUs {
    param([string]$GpoId, [string]$Domain)

    $linked = @()
    try {
        $domainDN  = (Get-ADDomain -Identity $Domain).DistinguishedName
        $domainDNS = (Get-ADDomain -Identity $Domain).DNSRoot

        # Domaine racine
        $domObj = Get-ADObject -Identity $domainDN -Properties gpLink -ErrorAction SilentlyContinue
        if ($domObj.gpLink -match $GpoId) { $linked += "Domaine ($domainDNS)" }

        # OUs - extraire uniquement le nom court (premier composant du DN)
        Get-ADOrganizationalUnit -Filter * -Properties gpLink, Name -ErrorAction SilentlyContinue |
            Where-Object { $_.gpLink -match $GpoId } |
            ForEach-Object { $linked += $_.Name }

        # Sites
        Get-ADObject -Filter { ObjectClass -eq "site" } `
            -SearchBase "CN=Sites,CN=Configuration,$domainDN" `
            -Properties gpLink, Name -ErrorAction SilentlyContinue |
            Where-Object { $_.gpLink -match $GpoId } |
            ForEach-Object { $linked += "Site:$($_.Name)" }

    } catch {
        Write-Warning "Erreur liaisons GPO $GpoId : $_"
    }

    return $linked -join " - "
}

try {
    $AllGPOs = Get-GPO -All -Domain $DomainName -ErrorAction Stop
} catch {
    Write-Error "Impossible de recuperer les GPO du domaine $DomainName : $_"
    exit 1
}

Write-Host "$($AllGPOs.Count) GPO(s) trouvee(s) dans $DomainName"

$Results = @()

foreach ($GPO in $AllGPOs) {
    Write-Host "  Traitement : $($GPO.DisplayName)" -ForegroundColor DarkCyan

    $statutGPO = switch ($GPO.GpoStatus) {
        "AllSettingsEnabled"       { "Oui" }
        "AllSettingsDisabled"      { "Non" }
        "ComputerSettingsDisabled" { "Partiel" }
        "UserSettingsDisabled"     { "Partiel" }
        default                    { $GPO.GpoStatus }
    }

    $creator = Get-GPOCreator  -GpoId $GPO.Id.ToString() -Domain $DomainName
    $ouLiees = Get-GPOLinkedOUs -GpoId $GPO.Id.ToString() -Domain $DomainName

    $Results += [PSCustomObject]@{
        "Domaine"          = $DomainName
        "Nom"              = $GPO.DisplayName
        "Cree par"         = $creator
        "Date de creation" = $GPO.CreationTime
        "OUs liees"        = $ouLiees
        "Active"           = $statutGPO
        "Description"      = ""
        "Utilite"          = ""
    }
}

$Results | Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8
Write-Host "Rapport genere : $ExportFile ($($Results.Count) GPO(s))"