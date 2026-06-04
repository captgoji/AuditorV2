# Version : 1.6
# Derniere modif : ErrorActionPreference Continue pour eviter propagation erreurs dotsource
param(
    [string[]] $Include    = @("*.ps1"),
    [string[]] $Exclude    = @("main.ps1"),
    [switch]   $DryRun,
    [switch]   $StopOnErr
)

$ErrorActionPreference = "Continue"  # Stop propage les erreurs des scripts dotsoources
$root           = $PSScriptRoot
$psDir          = Join-Path $root "Scripts\PowerShell"
$tempCsv        = Join-Path $root "Temp_CSV"
$logs           = Join-Path $root "Logs"
$initScript     = Join-Path $root "init.ps1"
$templatesDir   = Join-Path $root "Templates_Excel"
$pythonDir      = Join-Path $root "Scripts\Python"
$transcriptFile = Join-Path $logs "Audit_$((Get-Date -Format 'yyyyMMdd_HHmmss')).log"

# Creer le dossier Logs avant le transcript
New-Item -ItemType Directory -Force -Path $logs | Out-Null
Start-Transcript -Path $transcriptFile -Append -Force
Write-Host "[$(Get-Date -Format HH:mm:ss)] =============================" -ForegroundColor Cyan
Write-Host "[$(Get-Date -Format HH:mm:ss)] Audit demarre" -ForegroundColor Cyan
Write-Host "[$(Get-Date -Format HH:mm:ss)] Root        : $root" -ForegroundColor Cyan
Write-Host "[$(Get-Date -Format HH:mm:ss)] Transcript  : $transcriptFile" -ForegroundColor Cyan
Write-Host "[$(Get-Date -Format HH:mm:ss)] =============================" -ForegroundColor Cyan

# ---------------------------------------------------------------------
# ETAPE 0 : Detection automatique des templates manquants
# ---------------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $templatesDir | Out-Null
# init.ps1 est toujours lance pour verifier/installer les prerequis
# (Python, openpyxl, Microsoft.Graph, Npcap, templates Excel)
if (-not (Test-Path $initScript)) {
    Write-Error "init.ps1 introuvable dans $root."
    exit 1
}

Write-Host "[$(Get-Date -Format HH:mm:ss)] Lancement init.ps1 (verification des prerequis)..." -ForegroundColor Cyan
$initProc = Start-Process powershell.exe `
    -ArgumentList @("-NoProfile","-ExecutionPolicy","Bypass","-File",$initScript) `
    -Wait -PassThru -NoNewWindow

if ($initProc.ExitCode -ne 0) {
    Write-Error "init.ps1 a echoue. Consulte les erreurs ci-dessus."
    exit 1
}
Write-Host "[$(Get-Date -Format HH:mm:ss)] Prerequis OK.`n" -ForegroundColor Green

# ---------------------------------------------------------------------
# ETAPE 1 : Creation des dossiers
# ---------------------------------------------------------------------
foreach ($sub in @("Comptes","HV","VM")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $tempCsv $sub) | Out-Null
}

# Variables d'environnement exposees aux scripts enfants
$env:AUDIT_TEMP_CSV     = $tempCsv
$env:AUDIT_TEMP_COMPTES = Join-Path $tempCsv "Comptes"
$env:AUDIT_TEMP_HV      = Join-Path $tempCsv "HV"
$env:AUDIT_TEMP_VM      = Join-Path $tempCsv "VM"
$env:AUDIT_RUN_ID       = (Get-Date -Format "yyyyMMdd_HHmmss")

# Chemins NTFS a auditer - separateur ";" entre les chemins
# Peut etre defini avant le lancement : $env:AUDIT_NTFS_PATHS = "C:\Partages\RH;C:\Data"
# Si non defini, demande interactive au lancement
if (-not $env:AUDIT_NTFS_PATHS) {
    Write-Host ""
    Write-Host "=== AUDIT NTFS ===" -ForegroundColor Cyan
    Write-Host "Entrez les chemins a auditer (separateur ';'), ou appuyez sur Entree pour ignorer :" -ForegroundColor Yellow
    Write-Host "  Ex: C:\Partages\RH;C:\Partages\IT;C:\Partages\Commun" -ForegroundColor DarkGray
    $userInput = Read-Host "Chemins NTFS"
    $env:AUDIT_NTFS_PATHS = $userInput.Trim()
    if (-not $env:AUDIT_NTFS_PATHS) {
        Write-Host "Aucun chemin defini - 50_Audit_NTFS sera ignore.`n" -ForegroundColor DarkGray
    }
}

# ---------------------------------------------------------------------
# Fonctions
# ---------------------------------------------------------------------
function Get-AuditScripts {
    param([string]$Folder, [string[]]$Include, [string[]]$Exclude)

    if (-not (Test-Path $Folder)) {
        throw "Dossier introuvable : $Folder"
    }

    $all = Get-ChildItem -Path $Folder -Filter *.ps1 -File -Recurse

    $sel = @()
    foreach ($i in $Include) { $sel += ($all | Where-Object { $_.Name -like $i }) }
    if ($Exclude) {
        foreach ($e in $Exclude) { $sel = $sel | Where-Object { $_.Name -notlike $e } }
    }
    $sel = $sel | Select-Object -Unique

    $final = foreach ($s in $sel) {
        if ($s.BaseName -match '^[_.]') { continue }
        $firstKB = Get-Content -Path $s.FullName -TotalCount 40 -ErrorAction SilentlyContinue -Encoding UTF8
        if ($firstKB -match '#\s*AUDIT:\s*disabled') { continue }
        $s
    }

    $final | Sort-Object @{
        Expression = {
            $m = [regex]::Match($_.Name, '^(?<n>\d{2,})[ _-]')
            if ($m.Success) { [int]$m.Groups['n'].Value } else { [int]::MaxValue }
        }}, Name
}

function Invoke-AuditScript {
    param([Parameter(Mandatory=$true)][string]$ScriptPath)

    $name   = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath)
    $log    = Join-Path $logs ("{0}_{1}.log" -f $name, $env:AUDIT_RUN_ID)
    $result = [PSCustomObject]@{
        Script   = $name
        Path     = $ScriptPath
        Start    = (Get-Date)
        ExitCode = $null
        Log      = $log
        Status   = "PENDING"
        End      = $null
    }

    if ($DryRun) {
        Write-Host "DRY-RUN >> $name ($ScriptPath)" -ForegroundColor Yellow
        $result.Status = "DRY-RUN"
        return $result
    }

    # Scripts interactifs (Read-Host) : lances directement dans la console courante
    # Critere : nom commence par 00_ ou contient le tag # AUDIT: interactive
    $isInteractive = ($name -match '^00_') -or
                     ((Get-Content $ScriptPath -TotalCount 10 -ErrorAction SilentlyContinue) -match '#\s*AUDIT:\s*interactive')

    Write-Host ">> Execution : $name" -ForegroundColor Cyan
    try {
        if ($isInteractive) {
            # Dotsource dans le processus courant : herite de la console et du Read-Host
            # Le transcript main.ps1 capture deja toute la sortie console
            Write-Host "   [Mode interactif]" -ForegroundColor DarkYellow
            try {
                . $ScriptPath
                $exitCode = 0
            } catch {
                $exitCode = 1
                Write-Warning "Erreur dans $name : $_"
            }
            # Log minimal - la sortie complete est dans le transcript Audit_*.log
            "[Execution interactive - voir transcript Audit_$($env:AUDIT_RUN_ID).log]" |
                Out-File -FilePath $log -Encoding UTF8
        } else {
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName  = "powershell.exe"
            $psi.Arguments = @("-NoProfile","-ExecutionPolicy","Bypass","-File",$ScriptPath) -join " "
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError  = $true
            $psi.UseShellExecute        = $false

            $p      = [System.Diagnostics.Process]::Start($psi)
            $stdout = $p.StandardOutput.ReadToEnd()
            $stderr = $p.StandardError.ReadToEnd()
            $p.WaitForExit()
            $exitCode = $p.ExitCode

            $stdout | Out-File -FilePath $log -Encoding UTF8
            if ($stderr) {
                "`n--- STDERR ---`n$stderr" | Out-File -FilePath $log -Append -Encoding UTF8
            }
        }

        $result.ExitCode = $exitCode
        $result.Status   = if ($exitCode -eq 0) { "OK" } else { "ERROR" }

        if ($exitCode -ne 0) {
            Write-Warning ("{0} termine avec code {1}. Voir log : {2}" -f $name, $exitCode, $log)
            if ($StopOnErr) { throw "Arret sur erreur (StopOnErr)" }
        } else {
            Write-Host ("{0} OK. Log : {1}" -f $name, $log) -ForegroundColor Green
        }
    } catch {
        $result.ExitCode = -1
        $result.Status   = "EXCEPTION"
        Write-Error ("{0} a echoue : {1}" -f $name, $_.Exception.Message)
        if ($StopOnErr) { throw }
    } finally {
        $result.End = (Get-Date)
    }
    return $result
}

# ---------------------------------------------------------------------
# ETAPE 1 : Collecte (scripts PowerShell)
# ---------------------------------------------------------------------
Write-Host "=== COLLECTE ===" -ForegroundColor Cyan

$scripts = Get-AuditScripts -Folder $psDir -Include $Include -Exclude $Exclude

if (-not $scripts) {
    Write-Host "Aucun script trouve dans $psDir (motifs Include/Exclude ?)" -ForegroundColor Yellow
    Stop-Transcript
    exit 0
}

Write-Host "Scripts trouves :" -ForegroundColor DarkCyan
$scripts | ForEach-Object { "  - " + $_.Name } | Write-Host

$results = @()
foreach ($s in $scripts) {
    $results += Invoke-AuditScript -ScriptPath $s.FullName
}

# ---------------------------------------------------------------------
# RESUME
# ---------------------------------------------------------------------
Write-Host "`n=== RESUME ===" -ForegroundColor Cyan
$results | Select-Object Script,Status,ExitCode,Start,End,Log | Format-Table -AutoSize

$ok     = ($results | Where-Object Status -eq "OK").Count
$errors = ($results | Where-Object Status -in @("ERROR","EXCEPTION")).Count
Write-Host "OK : $ok  |  Erreurs : $errors" -ForegroundColor $(if ($errors -gt 0) { "Red" } else { "Green" })
Write-Host "CSV disponibles dans : $tempCsv`n" -ForegroundColor Yellow
Stop-Transcript