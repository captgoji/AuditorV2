# Usage : soit passer les chemins en parametre soit definir la variable d'environnement AUDIT_NTFS_PATHS (separateur ";")
#   $env:AUDIT_NTFS_PATHS = "C:\Partages;D:\Data"
# Version : 1.2
# Derniere modif : Try/catch Get-Acl, fix nom CSV, exit 0

param(
    [string[]]$FolderPath
)

$tempCsv        = $env:AUDIT_TEMP_CSV
$CompSys        = Get-CimInstance Win32_ComputerSystem
$IsDomainJoined = $CompSys.PartOfDomain
$DomainName     = if ($IsDomainJoined) { $CompSys.Domain } else { "WORKGROUP" }
$ExportFile     = "$tempCsv\$env:COMPUTERNAME.$DomainName" + "_Audit_NTFS.csv"

if (-not $FolderPath -or $FolderPath.Count -eq 0) {
    $envPaths = $env:AUDIT_NTFS_PATHS
    if ($envPaths) {
        $FolderPath = $envPaths -split ";" | Where-Object { $_ -ne "" }
    }
}

if (-not $FolderPath -or $FolderPath.Count -eq 0) {
    Write-Host "Aucun chemin NTFS defini - script ignore. Definir AUDIT_NTFS_PATHS dans main.ps1 pour activer cet audit."
    exit 0
}

$FolderItems = foreach ($p in $FolderPath) {
    $p = $p.Trim()
    if (Test-Path -LiteralPath $p -PathType Container) {
        Get-Item -LiteralPath $p
    } else {
        Write-Warning "Chemin introuvable ou non accessible : $p"
    }
}

if (-not $FolderItems) {
    Write-Warning "Aucun dossier valide trouve parmi les chemins specifies."
    exit 0
}

$Rapport = @()

function Get-Depth {
    param([string]$Path, [string]$RootPath)
    $rel = $Path.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
    if ($rel -eq "") { return 0 }
    return ($rel -split '\\').Count
}

foreach ($Folder in $FolderItems) {
    $RootPath = $Folder.FullName

    try {
        $Acl = Get-Acl -LiteralPath $Folder.FullName -ErrorAction Stop | Select-Object -ExpandProperty Access
        foreach ($Entry in $Acl) {
            $Rapport += [PSCustomObject]@{
                Folder            = $Folder.FullName
                Depth             = 0
                IdentityReference = $Entry.IdentityReference
                FileSystemRights  = $Entry.FileSystemRights
                IsInherited       = $Entry.IsInherited
            }
        }
    } catch { Write-Warning "ACL inaccessible : $($Folder.FullName) - $_" }

    $SubFolders = Get-ChildItem -LiteralPath $Folder.FullName -Directory -Recurse -ErrorAction SilentlyContinue
    foreach ($SubFolder in $SubFolders) {
        try {
            $SubFolderAcl = Get-Acl -LiteralPath $SubFolder.FullName -ErrorAction Stop | Select-Object -ExpandProperty Access
            $depth = Get-Depth -Path $SubFolder.FullName -RootPath $RootPath
            foreach ($Entry in $SubFolderAcl) {
                $Rapport += [PSCustomObject]@{
                    Folder            = $SubFolder.FullName
                    Depth             = $depth
                    IdentityReference = $Entry.IdentityReference
                    FileSystemRights  = $Entry.FileSystemRights
                    IsInherited       = $Entry.IsInherited
                }
            }
        } catch { Write-Warning "ACL inaccessible : $($SubFolder.FullName) - $_" }
    }
}

$Rapport | Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8
Write-Host "Rapport NTFS genere : $ExportFile ($($Rapport.Count) entrees)"