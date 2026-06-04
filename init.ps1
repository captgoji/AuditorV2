# init.ps1
# Version : 1.6
# Derniere modif : Npcap installation interactive (version gratuite sans /S)
# Lance automatiquement par main.ps1.

param(
    [switch]$Force  # Regenere tous les templates meme s'ils existent deja
)

$root           = $PSScriptRoot
$pythonDir      = Join-Path $root "Scripts\Python"
$templatesDir   = Join-Path $root "Templates_Excel"
$logsDir        = Join-Path $root "Logs"
$transcriptFile = Join-Path $logsDir "Init_$((Get-Date -Format 'yyyyMMdd_HHmmss')).log"

New-Item -ItemType Directory -Force -Path $logsDir | Out-Null
Start-Transcript -Path $transcriptFile -Append -Force
Write-Host "[$(Get-Date -Format HH:mm:ss)] Init demarre" -ForegroundColor Cyan
Write-Host "[$(Get-Date -Format HH:mm:ss)] Transcript : $transcriptFile" -ForegroundColor Cyan

# Recharger le PATH depuis le registre pour detecter Python installe recemment
# (evite de reinstaller Python a chaque lancement)
try {
    $regEnv    = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'
    $regPath   = (Get-ItemProperty -Path $regEnv -Name PATH -ErrorAction Stop).PATH
    $userPath  = [System.Environment]::GetEnvironmentVariable("PATH", "User")
    $env:PATH  = "$regPath;$userPath"
} catch {}

# ==============================================================================
# PYTHON
# ==============================================================================

function Install-Python {
    Write-Host "Python introuvable. Tentative de telechargement et installation..." -ForegroundColor Yellow

    $pythonVersion = "3.12.4"
    $installer     = "$env:TEMP\python-installer.exe"
    $downloadUrl   = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-amd64.exe"

    try {
        Write-Host "  Telechargement de Python $pythonVersion..." -ForegroundColor Yellow
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installer -UseBasicParsing
    } catch {
        Stop-Transcript
    Write-Error "Telechargement echoue : $_. Installe Python manuellement depuis https://www.python.org/downloads/ en cochant 'Add Python to PATH', puis relance main.ps1."
        exit 1
    }

    Write-Host "  Installation silencieuse de Python..." -ForegroundColor Yellow
    Start-Process -FilePath $installer -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0" -Wait -NoNewWindow
    Remove-Item $installer -Force -ErrorAction SilentlyContinue

    # Rafraichir le PATH de la session courante
    $machinePath = [System.Environment]::GetEnvironmentVariable("PATH", "Machine")
    $userPath    = [System.Environment]::GetEnvironmentVariable("PATH", "User")
    $env:PATH    = "$machinePath;$userPath"

    if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
        Stop-Transcript
    Write-Error "Python installe mais introuvable dans le PATH. Ferme et reouvre PowerShell, puis relance main.ps1."
        exit 1
    }

    Write-Host "Python $pythonVersion installe avec succes." -ForegroundColor Green
}

# Verifier ou installer Python
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Install-Python
}

$pyVersion = python --version 2>&1
Write-Host "Python detecte : $pyVersion" -ForegroundColor Green

# ==============================================================================
# OPENPYXL
# ==============================================================================

$openpyxlCheck = python -c "import openpyxl" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "openpyxl manquant. Installation en cours..." -ForegroundColor Yellow
    python -m pip install openpyxl --quiet
    if ($LASTEXITCODE -ne 0) {
        Stop-Transcript
    Write-Error "Impossible d'installer openpyxl."
        exit 1
    }
    Write-Host "openpyxl installe." -ForegroundColor Green
}



# ==============================================================================
# MICROSOFT GRAPH (requis par 12_Audit_Comptes_365.ps1)
# ==============================================================================

Write-Host "[$(Get-Date -Format HH:mm:ss)] Verification Microsoft.Graph..." -ForegroundColor Cyan

$mgModule = Get-Module -ListAvailable -Name Microsoft.Graph -ErrorAction SilentlyContinue
if (-not $mgModule) {
    Write-Host "  Microsoft.Graph absent. Installation en cours (peut prendre quelques minutes)..." -ForegroundColor Yellow
    try {
        # Forcer NuGet sans interaction
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers -ErrorAction SilentlyContinue | Out-Null
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module Microsoft.Graph -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
        Write-Host "  Microsoft.Graph installe avec succes." -ForegroundColor Green
    } catch {
        Write-Warning "Installation Microsoft.Graph echouee : $_"
        Write-Warning "12_Audit_Comptes_365.ps1 sera ignore si le module est absent."
    }
} else {
    $mgVer = ($mgModule | Sort-Object Version -Descending | Select-Object -First 1).Version
    Write-Host "  Microsoft.Graph detecte : v$mgVer" -ForegroundColor Green
}

# ==============================================================================
# NPCAP (requis par Nmap pour les scans reseau)
# ==============================================================================

$npcapInstaller = Join-Path $root "Tools\Nmap\npcap-installer.exe"

if (Test-Path $npcapInstaller) {
    Write-Host "[$(Get-Date -Format HH:mm:ss)] Verification Npcap..." -ForegroundColor Cyan

    $needInstall = $false
    $npcapService = Get-Service -Name npcap -ErrorAction SilentlyContinue

    if (-not $npcapService) {
        Write-Host "  Npcap absent - installation requise." -ForegroundColor Yellow
        $needInstall = $true
    } else {
        # Tester la compatibilite avec Nmap via un scan loopback
        $nmapExe = Join-Path $root "Tools\Nmap\nmap.exe"
        if (Test-Path $nmapExe) {
            $testOut = & $nmapExe -sn 127.0.0.1 2>&1 | Out-String
            if ($testOut -match "Could not import all necessary Npcap") {
                Write-Host "  Npcap incompatible avec Nmap 7.99 - mise a jour requise." -ForegroundColor Yellow
                $needInstall = $true
            } else {
                Write-Host "  Npcap OK." -ForegroundColor Green
            }
        }
    }

    if ($needInstall) {
        # Desinstaller l ancienne version si presente
        # Chercher dans les deux ruches de registre, prendre le premier resultat
        $npcapUninstall = @(
            Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue
            Get-ChildItem "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction SilentlyContinue
        ) | ForEach-Object {
            Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
        } | Where-Object {
            $_.DisplayName -like "*Npcap*"
        } | Select-Object -First 1

        if ($npcapUninstall) {
            Write-Host "  Desinstallation Npcap v$($npcapUninstall.DisplayVersion)..." -ForegroundColor Yellow
            $uninstallStr = $npcapUninstall.UninstallString

            # Npcap utilise un desinstalleur NSIS (exe direct avec /S)
            # Extraire le chemin de l exe entre guillemets si present
            if ($uninstallStr -match '"([^"]+)"') {
                $uninstallExe = $Matches[1]
            } else {
                $uninstallExe = $uninstallStr.Trim() -split ' ' | Select-Object -First 1
            }

            if (Test-Path $uninstallExe) {
                $proc = Start-Process $uninstallExe -ArgumentList "/S" -Wait -PassThru -NoNewWindow
                Write-Host "  Desinstallation terminee (code $($proc.ExitCode))." -ForegroundColor Green
            } else {
                Write-Warning "  Desinstalleur introuvable : $uninstallExe"
            }
            Start-Sleep -Seconds 5
        } else {
            Write-Host "  Aucune installation Npcap precedente trouvee dans le registre." -ForegroundColor DarkGray
        }

        # La version gratuite de Npcap ne supporte pas l installation silencieuse (/S)
        # L installeur est lance en mode interactif - l utilisateur doit cliquer Next
        Write-Host ""
        Write-Host "  *** ACTION REQUISE ***" -ForegroundColor Yellow
        Write-Host "  L installeur Npcap 1.88 va s ouvrir." -ForegroundColor Yellow
        Write-Host "  Completez l installation manuellement puis revenez ici." -ForegroundColor Yellow
        Write-Host "  Recommande : cocher 'WinPcap API-compatible mode'" -ForegroundColor DarkGray
        Write-Host ""
        $proc = Start-Process -FilePath $npcapInstaller -Wait -PassThru
        if ($proc.ExitCode -eq 0 -or $proc.ExitCode -eq 3010) {
            Start-Sleep -Seconds 3
            Start-Service -Name npcap -ErrorAction SilentlyContinue
            Write-Host "  Npcap 1.88 installe avec succes." -ForegroundColor Green
        } else {
            Write-Warning "Installation Npcap annulee ou echouee (code $($proc.ExitCode))."
            Write-Warning "La reconnaissance reseau ne fonctionnera pas sans Npcap."
        }
    }
} else {
    Write-Host "[$(Get-Date -Format HH:mm:ss)] npcap-installer.exe absent de Tools\Nmap\ - Npcap ignore." -ForegroundColor DarkGray
}

# ==============================================================================
# GENERATION DES TEMPLATES
# ==============================================================================

New-Item -ItemType Directory -Force -Path $templatesDir | Out-Null

$scripts = Get-ChildItem -Path $pythonDir -Filter "Template*.py" -File -ErrorAction SilentlyContinue

if (-not $scripts) {
    Write-Warning "Aucun script Template*.py trouve dans $pythonDir"
    exit 0
}

# Exposer le dossier de sortie aux scripts Python
$env:AUDIT_TEMPLATES_DIR = $templatesDir

Write-Host "`n=== GENERATION DES TEMPLATES EXCEL ===" -ForegroundColor Cyan

$results = @()

foreach ($s in $scripts) {
    $keyword = $s.BaseName -replace "Template", "" -replace "Audit", ""
    $existing = Get-ChildItem -Path $templatesDir -Filter "*.xlsx" -ErrorAction SilentlyContinue |
                Where-Object { $_.BaseName -replace "_","" -match [regex]::Escape($keyword) }

    if ($existing -and -not $Force) {
        Write-Host "  [SKIP] $($s.Name) -> $($existing.Name) existe deja" -ForegroundColor DarkGray
        $results += [PSCustomObject]@{ Script = $s.Name; Status = "SKIP"; Detail = $existing.Name }
        continue
    }

    Write-Host "  [GEN]  $($s.Name)" -ForegroundColor Yellow
    $output = python $s.FullName 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "         OK" -ForegroundColor Green
        $results += [PSCustomObject]@{ Script = $s.Name; Status = "OK"; Detail = ($output | Select-Object -Last 1) }
    } else {
        Write-Host "         ERREUR" -ForegroundColor Red
        Write-Host "         $output" -ForegroundColor Red
        $results += [PSCustomObject]@{ Script = $s.Name; Status = "ERREUR"; Detail = $output }
    }
}

Write-Host "`n=== RESUME ===" -ForegroundColor Cyan
$results | Format-Table -AutoSize

$ok     = ($results | Where-Object Status -eq "OK").Count
$skip   = ($results | Where-Object Status -eq "SKIP").Count
$errors = ($results | Where-Object Status -eq "ERREUR").Count

Write-Host "Generes : $ok  |  Ignores : $skip  |  Erreurs : $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
Write-Host "Templates disponibles dans : $templatesDir`n" -ForegroundColor Yellow

Stop-Transcript
if ($errors -gt 0) { exit 1 } else { exit 0 }