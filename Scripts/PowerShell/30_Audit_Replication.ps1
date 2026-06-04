# DEPENDANCE : ce script necessite que 20_Audit_VMs.ps1 ait ete execute au prealable.
# Version : 1.3
# Derniere modif : Fix _ CSV InputVMs, exit 0 si repadmin absent
# Il lit le CSV genere dans Temp_CSV\VM\*Audit_VMs.csv.
# Pour l'activer : retirer le prefixe "_" du nom de fichier.
 
$tempCsv    = $env:AUDIT_TEMP_CSV
$CompSys    = Get-CimInstance Win32_ComputerSystem
$DomainName = $CompSys.Domain
 
$InputVMs   = "$tempCsv\VM\$env:COMPUTERNAME.$DomainName" + "_Audit_VMs.csv"
$ExportFile = "$tempCsv\VM\$env:COMPUTERNAME.$DomainName" + "_Audit_Replication.csv"
 
Write-Host "Audit replication : lecture des VMs dans $InputVMs"
 
if (-not (Test-Path $InputVMs)) {
    Write-Host "Fichier introuvable : $InputVMs"
    exit 1
}
 
if (-not (Get-Command repadmin -ErrorAction SilentlyContinue)) {
    Write-Host "repadmin introuvable (RSAT ?)."
    exit 1
}
 
$rows = Import-Csv -Path $InputVMs
if (-not $rows -or $rows.Count -eq 0) {
    Write-Host "CSV VMs vide."
    exit 0
}
 
$Result = @()
 
foreach ($row in $rows) {
    $NomVM        = $row.'Nom VM'
    $rep          = "Non"
    $cibles       = ""
    $dernierStatut = "N/A"
    $nbErreurs    = 0
 
    if ($NomVM -and $NomVM.Trim() -ne "") {
        $reachable = Test-Connection -ComputerName $NomVM -Count 1 -Quiet -ErrorAction SilentlyContinue
 
        if ($reachable) {
            $raw = repadmin /showrepl $NomVM /csv 2>$null
 
            if ($raw) {
                $headerLine = ($raw | Select-String -SimpleMatch "," | Select-Object -First 1)
                if ($headerLine) {
                    $start    = $headerLine.LineNumber - 1
                    $csvBlock = $raw[$start..($raw.Count - 1)]
 
                    $tmp = [IO.Path]::GetTempFileName() + ".csv"
                    $csvBlock | Set-Content -Path $tmp -Encoding UTF8
 
                    try { $data = Import-Csv -Path $tmp } catch { $data = $null }
 
                    if ($data -and $data.Count -gt 0) {
                        $sample = $data[0]
                        function PickCol($obj, $cands) {
                            foreach ($c in $cands) {
                                if ($obj.PSObject.Properties.Name -contains $c) { return $c }
                            }
                            return $null
                        }
 
                        $colSrc    = PickCol $sample @('Source DSA','Source Server','Source','Src DSA')
                        $colDst    = PickCol $sample @('Destination DSA','Destination Server','Destination','Dst DSA')
                        $colResult = PickCol $sample @('Last Failure Status','Failure Status','Last Result','Status')
                        $colTime   = PickCol $sample @('Last Success Time','Last Attempt Time','Last Sync Time','Time')
 
                        if ($colSrc -and $colDst) {
                            $partners = New-Object System.Collections.Generic.HashSet[string]
                            foreach ($r in $data) {
                                $src = $r.$colSrc
                                $dst = $r.$colDst
                                if ([string]::IsNullOrWhiteSpace($src) -or [string]::IsNullOrWhiteSpace($dst)) { continue }
                                if ($src -notlike "*$NomVM*") { [void]$partners.Add($src) }
                                if ($dst -notlike "*$NomVM*") { [void]$partners.Add($dst) }
 
                                # Comptage erreurs (statut != 0 et non vide)
                                if ($colResult) {
                                    $statusVal = $r.$colResult
                                    if ($statusVal -and $statusVal -ne "0" -and $statusVal -ne "") {
                                        $nbErreurs++
                                    }
                                }
                            }
 
                            $arr = $partners.ToArray() | Sort-Object -Unique
                            if ($arr.Count -gt 0) {
                                $rep    = "Oui"
                                $cibles = ($arr -join " - ")
                            }
 
                            # Dernier statut : derniere ligne avec valeur de statut
                            if ($colResult) {
                                $lastRow = $data | Where-Object { $_.$colResult -ne "" } | Select-Object -Last 1
                                if ($lastRow) {
                                    $code = $lastRow.$colResult
                                    $dernierStatut = if ($code -eq "0") { "OK" } else { "Erreur ($code)" }
                                    if ($colTime -and $lastRow.$colTime) {
                                        $dernierStatut += " - $($lastRow.$colTime)"
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } else {
            $dernierStatut = "Injoignable"
        }
    }
 
    $Result += [PSCustomObject]@{
        "Nom VM"                   = $NomVM
        "Replication"              = $rep
        "Cible de la replication"  = $cibles
        "Statut derniere repl"     = $dernierStatut
        "Nb erreurs replication"   = $nbErreurs
    }
}
 
$Result | Export-Csv -Path $ExportFile -NoTypeInformation -Encoding UTF8
Write-Host "Export : $ExportFile"