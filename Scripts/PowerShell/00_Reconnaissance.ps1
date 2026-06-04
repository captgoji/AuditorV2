# 00_Reconnaissance.ps1
# Version : 2.8
# Derniere modif : Suppression --packet-trace
# Scan reseau par etapes Nmap → identification serveurs → test WinRM → validation utilisateur → lancement audit distant

$ErrorActionPreference = "Stop"

# Calcul du chemin racine : fonctionne en execution directe ET en dotsource depuis main.ps1
$_scriptPath = $MyInvocation.MyCommand.Path
if (-not $_scriptPath) { $_scriptPath = Join-Path $PSScriptRoot "00_Reconnaissance.ps1" }
$_scriptDir  = Split-Path $_scriptPath -Parent

# Structure attendue : NewAuditroV2\Scripts\PowerShell\
if ((Split-Path $_scriptDir -Leaf) -eq "PowerShell") {
    $root = Split-Path (Split-Path $_scriptDir -Parent) -Parent
} else {
    $root = $_scriptDir
}
$logsDir    = Join-Path $root "Logs"
$tempCsv    = Join-Path $root "Temp_CSV"
$runId      = Get-Date -Format "yyyyMMdd_HHmmss"
$recoDir    = Join-Path $logsDir "Reconnaissance_$runId"
$csvResults = Join-Path $tempCsv "Reconnaissance_$runId.csv"

New-Item -ItemType Directory -Force -Path $recoDir  | Out-Null
New-Item -ItemType Directory -Force -Path $tempCsv  | Out-Null

# Fichiers de travail inter-etapes
$fileStep1  = Join-Path $recoDir "step1_hosts_up.txt"
$fileStep2  = Join-Path $recoDir "step2_servers.xml"
$fileStep3  = Join-Path $recoDir "step3_winrm.xml"
$logStep1   = Join-Path $recoDir "step1_scan_updown.log"
$logStep2   = Join-Path $recoDir "step2_scan_os.log"
$logStep3   = Join-Path $recoDir "step3_scan_winrm.log"

# ─────────────────────────────────────────────────────────────
# FONCTIONS UTILITAIRES
# ─────────────────────────────────────────────────────────────

function Write-Step {
    param([string]$Msg, [string]$Color = "Cyan")
    Write-Host ""
    Write-Host ("-" * 60) -ForegroundColor $Color
    Write-Host "  $Msg" -ForegroundColor $Color
    Write-Host ("-" * 60) -ForegroundColor $Color
    Write-Host ""
}

function Write-Log {
    param([string]$File, [string]$Msg)
    $line = "[$(Get-Date -Format 'HH:mm:ss')] $Msg"
    Write-Host $line
    Add-Content -Path $File -Value $line -Encoding UTF8
}

function Confirm-Action {
    param([string]$Question, [string]$Color = "Yellow")
    Write-Host "$Question [O/N] : " -ForegroundColor $Color -NoNewline
    $r = Read-Host
    return ($r -match '^[OoYy]')
}

# ─────────────────────────────────────────────────────────────
# ETAPE 0 : INSTALLATION NMAP
# ─────────────────────────────────────────────────────────────

function Get-NmapExe {
    Write-Step "ETAPE 0 : Verification Nmap"

    # 1. Nmap portable dans Tools\Nmap\ (prioritaire)
    $portablePath = Join-Path $root "Tools\Nmap\nmap.exe"
    if (Test-Path $portablePath) {
        $ver = & $portablePath --version 2>&1 | Select-Object -First 1
        Write-Host "  Nmap portable detecte : $ver" -ForegroundColor Green

        # Npcap est gere par init.ps1 - verifier seulement que le service tourne
        $npcapService = Get-Service -Name npcap -ErrorAction SilentlyContinue
        if (-not $npcapService -or $npcapService.Status -ne "Running") {
            Write-Warning "Npcap non demarre. Lance init.ps1 d abord ou installe npcap-1.88.exe depuis Tools\Nmap\"
            exit 1
        }
        Write-Host "  Npcap OK." -ForegroundColor Green
        return $portablePath
    }

    # 2. Fallback : Nmap installe sur le systeme
    $nmapCmd = Get-Command nmap -ErrorAction SilentlyContinue
    $systemPaths = @(
        $(if ($nmapCmd) { $nmapCmd.Source } else { $null }),
        "$env:ProgramFiles\Nmap\nmap.exe",
        "${env:ProgramFiles(x86)}\Nmap\nmap.exe"
    ) | Where-Object { $_ -and (Test-Path $_) }

    if ($systemPaths) {
        $ver = & $systemPaths[0] --version 2>&1 | Select-Object -First 1
        Write-Host "  Nmap systeme detecte : $ver" -ForegroundColor Green
        return $systemPaths[0]
    }

    Write-Error "Nmap introuvable.`nPlace nmap.exe et ses DLLs dans : $root\Tools\Nmap\`nTelecharge la version zip depuis https://nmap.org/download.html"
    exit 1
}

# ─────────────────────────────────────────────────────────────
# ETAPE 1 : SCAN UP/DOWN
# ─────────────────────────────────────────────────────────────

function Invoke-Step1-HostDiscovery {
    param([string]$NmapExe, [string[]]$Ranges)

    Write-Step "ETAPE 1 : Decouverte des hotes actifs (ping sweep)"
    Write-Log $logStep1 "Debut Step1 - Plages : $($Ranges -join ', ')"

    $rangeArgs = $Ranges -join " "
    $cmd = "`"$NmapExe`" -sn -T4 --open $rangeArgs -oG `"$fileStep1`""
    Write-Log $logStep1 "Commande : nmap -sn -T4 --open $rangeArgs"

    Write-Host "  Ping sweep en cours (T2, peut prendre quelques minutes)..." -ForegroundColor DarkCyan
    $output = & $NmapExe -sn -T4 --open -v --stats-every 15s $rangeArgs 2>&1
    $output | ForEach-Object {
        Write-Log $logStep1 $_
        if ($_ -match "Nmap scan report|Stats:|hosts up|host up|Initiating") {
            Write-Host "  > $_" -ForegroundColor DarkCyan
        }
    }
    $output | Out-File $fileStep1 -Encoding UTF8

    # Parser les IPs up depuis la sortie nmap
    $hostsUp = $output | Where-Object { $_ -match "Nmap scan report for" } | ForEach-Object {
        if ($_ -match 'for\s+(\S+)\s+\((\d[\d.]+)\)') {
            [PSCustomObject]@{ Hostname = $Matches[1]; IP = $Matches[2] }
        } elseif ($_ -match 'for\s+(\d[\d.]+)') {
            [PSCustomObject]@{ Hostname = $Matches[1]; IP = $Matches[1] }
        }
    }

    Write-Log $logStep1 "Hotes actifs detectes : $($hostsUp.Count)"
    Write-Host ""
    Write-Host "  Hotes actifs trouves : $($hostsUp.Count)" -ForegroundColor Green
    $hostsUp | ForEach-Object { Write-Host "    - $($_.IP)  $($_.Hostname)" -ForegroundColor DarkCyan }

    return $hostsUp
}

# ─────────────────────────────────────────────────────────────
# ETAPE 2 : IDENTIFICATION SERVEURS (OS fingerprint)
# ─────────────────────────────────────────────────────────────

function Invoke-Step2-ServerIdentification {
    param([string]$NmapExe, [object[]]$Hosts)

    Write-Step "ETAPE 2 : Identification des serveurs (OS fingerprinting)"
    Write-Log $logStep2 "Debut Step2 - $($Hosts.Count) hotes a analyser"

    $ipList = @($Hosts | Select-Object -ExpandProperty IP)
    $ips = $ipList -join " "
    Write-Log $logStep2 "Commande : nmap -O --osscan-guess -T4 $ips"

    Write-Host "  OS fingerprinting en cours..." -ForegroundColor DarkCyan
    $output = & $NmapExe -O --osscan-guess -T4 -v --stats-every 15s @ipList 2>&1
    $output | ForEach-Object {
        Write-Log $logStep2 $_
        if ($_ -match "Nmap scan report|OS details|Aggressive OS guesses") {
            Write-Host "  >> $_" -ForegroundColor Green
        }
        elseif ($_ -match "Stats:|Initiating|Completed") {
            Write-Host "  > $_" -ForegroundColor DarkCyan
        }
    }

    # Parser la sortie pour identifier les serveurs Windows Server / Linux Server
    $servers   = [System.Collections.Generic.List[object]]::new()
    $currentIP = $null
    $currentHN = $null

    foreach ($line in $output) {
        if ($line -match 'Nmap scan report for\s+(\S+)\s+\((\d[\d.]+)\)') {
            $currentHN = $Matches[1]; $currentIP = $Matches[2]
        } elseif ($line -match 'Nmap scan report for\s+(\d[\d.]+)') {
            $currentIP = $Matches[1]; $currentHN = $Matches[1]
        }

        if ($currentIP -and $line -match 'OS details:|Aggressive OS guesses:') {
            $osDetail = ($line -replace 'OS details:|Aggressive OS guesses:', '').Trim()
            $isServer = $osDetail -match 'Windows Server|Server \d{4}|Linux.*[Ss]erver|Ubuntu Server|Debian|CentOS|Red Hat|RHEL|Hyper-V'

            if ($isServer) {
                $existing = $servers | Where-Object { $_.IP -eq $currentIP }
                if (-not $existing) {
                    $servers.Add([PSCustomObject]@{
                        Hostname = $currentHN
                        IP       = $currentIP
                        OSDetail = $osDetail.Substring(0, [Math]::Min($osDetail.Length, 80))
                    })
                    Write-Log $logStep2 "  SERVEUR detecte : $currentIP ($currentHN) - $osDetail"
                }
            }
        }
    }

    # Fallback : si nmap n'a pas pu fingerprinter l'OS (pare-feu bloque ICMP TTL)
    # on inclut les hotes avec ports serveur ouverts (445, 3389, 5985, 22)
    Write-Host "  Scan ports serveur (fallback)..." -ForegroundColor DarkCyan
    $fallbackOutput = & $NmapExe -p 22,445,3389,5985,5986 --open -T4 -v --stats-every 15s @ipList 2>&1
    $fallbackOutput | ForEach-Object {
        Write-Log $logStep2 "  [fallback] $_"
        if ($_ -match "open|Stats:|Initiating") {
            Write-Host "  > [fallback] $_" -ForegroundColor DarkGray
        }
    }

    $currentIP = $null; $currentHN = $null
    foreach ($line in $fallbackOutput) {
        if ($line -match 'Nmap scan report for\s+(\S+)\s+\((\d[\d.]+)\)') {
            $currentHN = $Matches[1]; $currentIP = $Matches[2]
        } elseif ($line -match 'Nmap scan report for\s+(\d[\d.]+)') {
            $currentIP = $Matches[1]; $currentHN = $Matches[1]
        }
        if ($currentIP -and $line -match '(22|445|3389|5985|5986)/tcp\s+open') {
            if (-not ($servers | Where-Object { $_.IP -eq $currentIP })) {
                $servers.Add([PSCustomObject]@{
                    Hostname = $currentHN
                    IP       = $currentIP
                    OSDetail = "Inconnu (ports serveur ouverts - fingerprint echoue)"
                })
                Write-Log $logStep2 "  SERVEUR fallback : $currentIP ($currentHN)"
            }
        }
    }

    Write-Log $logStep2 "Serveurs identifies : $($servers.Count)"
    return $servers
}

# ─────────────────────────────────────────────────────────────
# ETAPE 3 : TEST WINRM
# ─────────────────────────────────────────────────────────────

function Invoke-Step3-WinRMCheck {
    param([string]$NmapExe, [object[]]$Servers)

    Write-Step "ETAPE 3 : Test WinRM (ports 5985/5986)"
    Write-Log $logStep3 "Debut Step3 - $($Servers.Count) serveurs a tester"

    $ipList = @($Servers | Select-Object -ExpandProperty IP)
    $ips    = $ipList -join " "
    Write-Host "  Test WinRM (ports 5985/5986)..." -ForegroundColor DarkCyan
    $output = & $NmapExe -p 5985,5986 --open -T4 -v --stats-every 10s @ipList 2>&1
    $output | ForEach-Object {
        Write-Log $logStep3 $_
        if ($_ -match "open|Stats:|Nmap scan report|Initiating") {
            Write-Host "  > $_" -ForegroundColor DarkCyan
        }
    }

    # IPs avec WinRM ouvert
    $winrmIPs  = [System.Collections.Generic.List[string]]::new()
    $currentIP = $null

    foreach ($line in $output) {
        if ($line -match 'Nmap scan report for\s+\S+\s+\((\d[\d.]+)\)') {
            $currentIP = $Matches[1]
        } elseif ($line -match 'Nmap scan report for\s+(\d[\d.]+)') {
            $currentIP = $Matches[1]
        }
        if ($currentIP -and $line -match '598[56]/tcp\s+open') {
            if (-not $winrmIPs.Contains($currentIP)) {
                $winrmIPs.Add($currentIP)
                Write-Log $logStep3 "  WinRM ouvert : $currentIP"
            }
        }
    }

    # Enrichir les objets serveur
    foreach ($srv in $Servers) {
        $srv | Add-Member -NotePropertyName WinRM -NotePropertyValue ($winrmIPs.Contains($srv.IP)) -Force
    }

    Write-Log $logStep3 "Serveurs avec WinRM : $($winrmIPs.Count) / $($Servers.Count)"
    return $Servers
}

# ─────────────────────────────────────────────────────────────
# ETAPE 4 : VALIDATION UTILISATEUR
# ─────────────────────────────────────────────────────────────

function Invoke-Step4-UserValidation {
    param([object[]]$Servers)

    Write-Step "ETAPE 4 : Validation des cibles par l'utilisateur"

    $validated = [System.Collections.Generic.List[object]]::new()

    foreach ($srv in $Servers) {
        $winrmStatus = if ($srv.WinRM) { "OUI" } else { "NON" }
        $color       = if ($srv.WinRM) { "Green" } else { "DarkGray" }

        Write-Host ""
        Write-Host "  +-- Serveur detecte ----------------------------------" -ForegroundColor Cyan
        Write-Host "  |  Hostname : $($srv.Hostname)" -ForegroundColor White
        Write-Host "  |  IP       : $($srv.IP)" -ForegroundColor White
        Write-Host "  |  OS       : $($srv.OSDetail)" -ForegroundColor White
        Write-Host "  |  WinRM    : $winrmStatus" -ForegroundColor $color
        Write-Host "  +------------------------------------------------" -ForegroundColor Cyan

        $skip = $false

        if (-not $srv.WinRM) {
            Write-Host "  WinRM inactif sur cette machine. " -ForegroundColor DarkGray -NoNewline
            $include = Confirm-Action "Inclure quand meme dans le CSV (sans audit distant) ?"
            if (-not $include) {
                Write-Host "  -> Ignore." -ForegroundColor DarkGray
                continue
            }
            $skip = $true
            $skipRaison = "WinRM inactif"
        } else {
            $launch = Confirm-Action "  Inclure et auditer ce serveur ?" "Yellow"
            if (-not $launch) {
                $skip = $true
                Write-Host "  Raison du skip (laisser vide si aucune) : " -ForegroundColor DarkGray -NoNewline
                $skipRaison = Read-Host
                if (-not $skipRaison) { $skipRaison = "Skippe manuellement" }
                Write-Host "  -> Skip : $skipRaison" -ForegroundColor DarkGray
            } else {
                $skipRaison = ""
            }
        }

        $validated.Add([PSCustomObject]@{
            Hostname    = $srv.Hostname
            IP          = $srv.IP
            OSDetail    = $srv.OSDetail
            WinRM       = $srv.WinRM
            Skip        = if ($skip) { "Oui" } else { "Non" }
            SkipRaison  = $skipRaison
        })
    }

    return $validated
}

# ─────────────────────────────────────────────────────────────
# ETAPE 5 : EXPORT CSV + LANCEMENT AUDIT DISTANT
# ─────────────────────────────────────────────────────────────

function Invoke-Step5-AuditRemote {
    param([object[]]$Targets)

    Write-Step "ETAPE 5 : Export CSV et lancement des audits distants"

    # Export CSV resultats reconnaissance
    $Targets | Select-Object Hostname, IP, OSDetail, WinRM, Skip, SkipRaison |
        Export-Csv -Path $csvResults -NoTypeInformation -Encoding UTF8
    Write-Host "  CSV reconnaissance : $csvResults" -ForegroundColor Green

    # Recuperer les IPs locales pour exclure la machine courante
    $localIPs = @(
        [System.Net.Dns]::GetHostAddresses([System.Net.Dns]::GetHostName()) |
        Where-Object { $_.AddressFamily -eq "InterNetwork" } |
        ForEach-Object { $_.ToString() }
    )
    Write-Host "  IPs locales detectees : $($localIPs -join ', ')" -ForegroundColor DarkGray

    $toAudit = $Targets | Where-Object {
        $_.Skip -eq "Non" -and $_.WinRM -eq $true -and $_.IP -notin $localIPs
    }

    # Signaler les machines locales exclues
    $selfTargets = $Targets | Where-Object { $_.IP -in $localIPs }
    if ($selfTargets) {
        foreach ($s in $selfTargets) {
            Write-Host "  [SKIP] $($s.IP) est la machine locale - audit distant ignore." -ForegroundColor DarkGray
        }
    }

    if (-not $toAudit) {
        Write-Host "  Aucune machine distante a auditer (toutes skippees, WinRM inactif, ou machine locale)." -ForegroundColor Yellow
        return
    }

    Write-Host "  $($toAudit.Count) machine(s) distante(s) a auditer via WinRM." -ForegroundColor Cyan

    $mainScript = Join-Path $root "main.ps1"
    if (-not (Test-Path $mainScript)) {
        Write-Warning "main.ps1 introuvable dans $root - audit distant impossible."
        return
    }

    foreach ($target in $toAudit) {
        Write-Host ""
        Write-Host "  >> Audit distant : $($target.Hostname) ($($target.IP))" -ForegroundColor Cyan

        # Determiner si la machine est dans le domaine courant
        $currentDomain = (Get-CimInstance Win32_ComputerSystem).Domain
        $inDomain      = $target.Hostname -match "\.$([regex]::Escape($currentDomain))$" -or
                         $target.Hostname -eq $target.IP -eq $false

        $credParams = @{}

        # Ajouter l IP dans TrustedHosts si connexion par IP
        # (requis quand la cible n est pas joignable via Kerberos/FQDN)
        $connectTarget = $target.IP
        if ($target.Hostname -ne $target.IP -and $target.Hostname -match "\." ) {
            # On a un FQDN - tenter par nom d abord (Kerberos si meme domaine)
            $connectTarget = $target.Hostname
        }

        $currentTrusted = (Get-Item WSMan:\localhost\Client\TrustedHosts -ErrorAction SilentlyContinue).Value
        $alreadyTrusted = $currentTrusted -eq "*" -or $currentTrusted -split "," -contains $target.IP

        if (-not $alreadyTrusted) {
            Write-Host "  Ajout de $($target.IP) dans TrustedHosts..." -ForegroundColor DarkGray
            $newTrusted = if ($currentTrusted -and $currentTrusted -ne "") { "$currentTrusted,$($target.IP)" } else { $target.IP }
            Set-Item WSMan:\localhost\Client\TrustedHosts -Value $newTrusted -Force -ErrorAction SilentlyContinue
        }

        if (-not $inDomain -or $target.Hostname -eq $target.IP) {
            Write-Host "  Machine hors domaine - credentials requis." -ForegroundColor Yellow
            try {
                $cred = Get-Credential -Message "Credentials pour $($target.Hostname) ($($target.IP))"
                $credParams['Credential'] = $cred
            } catch {
                Write-Warning "Credentials annules - $($target.Hostname) ignore."
                continue
            }
        }

        # Dossier de sortie distant pour les CSV
        $remoteOutputDir = "C:\Windows\Temp\Audit_$runId"

        try {
            # Tester la connexion WinRM
            $session = New-PSSession -ComputerName $connectTarget @credParams -ErrorAction Stop

            # Creer le dossier temporaire sur la machine distante
            Invoke-Command -Session $session -ScriptBlock {
                param($dir)
                New-Item -ItemType Directory -Force -Path $dir | Out-Null
                # Sous-dossiers attendus par les scripts
                foreach ($sub in @("Comptes","HV","VM")) {
                    New-Item -ItemType Directory -Force -Path "$dir\$sub" | Out-Null
                }
            } -ArgumentList $remoteOutputDir

            # Copier tous les scripts PS sur la machine distante
            # Scripts exclus de l execution distante via WinRM :
            #   00_* : scripts locaux uniquement (reconnaissance reseau, etc.)
            #   12_Audit_Comptes_365.ps1 : necessite Connect-MgGraph interactif + tenant M365,
            #                              pas applicable sur un serveur AD/HV distant
            $excludeRemote = @(
                "^00_",
                "^12_Audit_Comptes_365"
            )
            $psScripts = Get-ChildItem -Path (Join-Path $root "Scripts\PowerShell") -Filter "*.ps1" -File |
                         Where-Object {
                             $name = $_.Name
                             -not ($excludeRemote | Where-Object { $name -match $_ })
                         }
            Write-Host "    Scripts a executer a distance ($($psScripts.Count)) :" -ForegroundColor DarkGray
            $psScripts | ForEach-Object { Write-Host "      - $($_.Name)" -ForegroundColor DarkGray }
            foreach ($s in $psScripts) {
                Copy-Item -Path $s.FullName -Destination $remoteOutputDir -ToSession $session -Force
            }

            # Executer chaque script a distance
            foreach ($s in $psScripts) {
                $remotePath = "$remoteOutputDir\$($s.Name)"
                Write-Host "    - Execution distante : $($s.Name)" -ForegroundColor DarkCyan

                try {
                    Invoke-Command -Session $session -ScriptBlock {
                        param($script, $csvDir)
                        $env:AUDIT_TEMP_CSV     = $csvDir
                        $env:AUDIT_TEMP_COMPTES = "$csvDir\Comptes"
                        $env:AUDIT_TEMP_HV      = "$csvDir\HV"
                        $env:AUDIT_TEMP_VM      = "$csvDir\VM"
                        & $script
                    } -ArgumentList $remotePath, $remoteOutputDir -ErrorAction Stop
                } catch {
                    Write-Warning "    Erreur sur $($s.Name) : $_"
                }
            }

            # Rapatrier les CSV
            $localDest = Join-Path $tempCsv $target.Hostname
            New-Item -ItemType Directory -Force -Path $localDest | Out-Null

            $remoteFiles = Invoke-Command -Session $session -ScriptBlock {
                param($dir)
                Get-ChildItem -Path $dir -Filter "*.csv" -Recurse | Select-Object -ExpandProperty FullName
            } -ArgumentList $remoteOutputDir

            foreach ($rf in $remoteFiles) {
                $localFile = Join-Path $localDest (Split-Path $rf -Leaf)
                Copy-Item -Path $rf -Destination $localFile -FromSession $session -Force
                Write-Host "    CSV rapatrie : $(Split-Path $rf -Leaf)" -ForegroundColor Green
            }

            # Nettoyage dossier temp distant
            Invoke-Command -Session $session -ScriptBlock {
                param($dir) Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue
            } -ArgumentList $remoteOutputDir

            Remove-PSSession $session
            Write-Host "  << $($target.Hostname) termine." -ForegroundColor Green

        } catch {
            Write-Warning "Connexion WinRM echouee sur $($target.Hostname) ($($target.IP)) : $_"
        }
    }
}

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

Write-Step "RECONNAISSANCE RESEAU - Framework Audit" "Magenta"
Write-Host "  Logs        : $recoDir" -ForegroundColor DarkGray
Write-Host "  CSV resultats : $csvResults" -ForegroundColor DarkGray

# Avertissement
Write-Host ""
Write-Host "  /!\ AVERTISSEMENT /!\" -ForegroundColor Red
Write-Host "  Ce script va effectuer un scan reseau actif (Nmap)." -ForegroundColor Yellow
Write-Host "  Assure-toi d'avoir les autorisations necessaires." -ForegroundColor Yellow
Write-Host "  Le scan peut declencher des alertes sur les equipements sensibles" -ForegroundColor Yellow
Write-Host '  (firewalls, IDS/IPS, equipements reseau, systemes industriels OT/ICS).' -ForegroundColor Yellow
Write-Host ""

if (-not (Confirm-Action "Confirmes-tu avoir les droits pour scanner ce reseau ?" "Red")) {
    Write-Host "Scan annule." -ForegroundColor DarkGray
    exit 0
}

# Saisie des plages reseau
Write-Host ""
Write-Host "Entrez les plages reseau a scanner (une par ligne, ligne vide pour terminer)." -ForegroundColor Cyan
Write-Host '  Formats acceptes : 192.168.1.0/24  ou  10.0.0.1-50  ou  192.168.1.5' -ForegroundColor DarkGray
Write-Host ""

$ranges = [System.Collections.Generic.List[string]]::new()
while ($true) {
    Write-Host "  Plage $($ranges.Count + 1) : " -ForegroundColor Yellow -NoNewline
    $input = Read-Host
    if ([string]::IsNullOrWhiteSpace($input)) {
        if ($ranges.Count -eq 0) {
            Write-Host "  Au moins une plage est requise." -ForegroundColor Red
            continue
        }
        break
    }
    $ranges.Add($input.Trim())
}

Write-Host ""
Write-Host "  Plages a scanner :" -ForegroundColor Cyan
$ranges | ForEach-Object { Write-Host "    - $_" -ForegroundColor White }

# Etape 0 : Nmap
$nmapExe = Get-NmapExe

# Etape 1 : Hosts up
$hostsUp = Invoke-Step1-HostDiscovery -NmapExe $nmapExe -Ranges $ranges

if (-not $hostsUp -or $hostsUp.Count -eq 0) {
    Write-Host "Aucun hote actif trouve sur les plages specifiees. Fin du script." -ForegroundColor Yellow
    exit 0
}

# Etape 2 : Identification serveurs
$servers = Invoke-Step2-ServerIdentification -NmapExe $nmapExe -Hosts $hostsUp

if (-not $servers -or $servers.Count -eq 0) {
    Write-Host "Aucun serveur identifie parmi les hotes actifs. Fin du script." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "  $($servers.Count) serveur(s) identifie(s)." -ForegroundColor Green

# Etape 3 : Test WinRM
$servers = Invoke-Step3-WinRMCheck -NmapExe $nmapExe -Servers $servers

# Etape 4 : Validation utilisateur
$targets = Invoke-Step4-UserValidation -Servers $servers

if (-not $targets -or $targets.Count -eq 0) {
    Write-Host "Aucune cible retenue. Fin du script." -ForegroundColor Yellow
    exit 0
}

# Etape 5 : Export + Audit distant
Invoke-Step5-AuditRemote -Targets $targets

Write-Host ""
Write-Host ("-" * 60) -ForegroundColor Magenta
Write-Host "  RECONNAISSANCE TERMINEE" -ForegroundColor Magenta
Write-Host "  Logs     : $recoDir" -ForegroundColor DarkGray
Write-Host "  CSV      : $csvResults" -ForegroundColor DarkGray
Write-Host ("-" * 60) -ForegroundColor Magenta