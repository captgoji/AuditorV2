# NewAuditroV2 — Framework d'Audit Active Directory & Infrastructure Windows

> Framework PowerShell modulaire pour auditer des environnements Windows : comptes AD, VMs, hyperviseurs Hyper-V, GPO, réplication, permissions NTFS et comptes Microsoft 365. Inclut une brique de reconnaissance réseau automatisée via Nmap avec exécution distante via WinRM.

---

## Sommaire

- [Fonctionnalités](#fonctionnalités)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Structure du projet](#structure-du-projet)
- [Utilisation](#utilisation)
- [Modules d'audit](#modules-daudit)
- [Reconnaissance réseau](#reconnaissance-réseau)
- [Exécution distante WinRM](#exécution-distante-winrm)
- [Sorties](#sorties)
- [Configuration](#configuration)
- [Compatibilité](#compatibilité)

---

## Fonctionnalités

- **Audit local ou distant** — s'exécute sur la machine courante ou sur des cibles distantes découvertes automatiquement via WinRM
- **Reconnaissance réseau** — scan Nmap multi-étapes (ping sweep → OS fingerprinting → test WinRM) avec validation utilisateur avant toute action distante
- **Prérequis automatiques** — `init.ps1` installe Python, openpyxl, Microsoft.Graph et Npcap sans intervention manuelle (sauf Npcap)
- **Exports Excel structurés** — chaque CSV collecté peut être injecté dans un template Excel formaté via le script python process_excels
- **Logging complet** — transcript PowerShell + logs Nmap par étape pour chaque exécution
- **Modulaire** — chaque script est indépendant, préfixé par numéro d'ordre, activable/désactivable en ajoutant un "_" devant le nom du fichier

---

## Prérequis

### Machine qui lance le framework

| Composant | Version minimale | Installé automatiquement |
|---|---|---|
| Windows Server / Windows 10+ | — | — |
| PowerShell | 5.1+ | — |
| Python | 3.12+ | ✅ via `init.ps1` |
| openpyxl | 3.x | ✅ via `init.ps1` |
| Microsoft.Graph | 2.x | ✅ via `init.ps1` (si `12_Audit_Comptes_365.ps1` présent) |
| Npcap | 1.88 | ❌ manuel via `init.ps1` (interactif) — voir Installation |
| Nmap | 7.99 (portable) | ✅ via `init.ps1` |

### Cibles distantes auditées via WinRM

- WinRM activé (`Enable-PSRemoting -Force`)
- Compte administrateur local ou de domaine
- Ports 5985 (HTTP) ou 5986 (HTTPS) ouverts

### Modules RSAT (optionnels — skip propre si absents)

| Module | Scripts concernés |
|---|---|
| ActiveDirectory | `10_`, `11_`, `13_` |
| GroupPolicy | `60_Audit_GPO` |
| repadmin | `20_Audit_VMs`, `30_Audit_Replication` |
| Hyper-V | `40_Audit_Hyperviseurs` |

---

## Installation

### 1. Cloner le dépôt

```powershell
git clone https://github.com/<votre-org>/NewAuditroV2.git
cd NewAuditroV2
```

### 2. Préparer Nmap (portable)

Si la version de Nmap n'est pas à jours dans le dépôt télécharger la version **zip** (pas l'installeur) depuis [nmap.org/download.html](https://nmap.org/download.html) et extraire dans `Tools\Nmap\`.

De même pour Npcap si il n'est pas à jours télécharger **Npcap** depuis [npcap.com](https://npcap.com/#download) ou [nmap.org/download.html](https://nmap.org/download.html) (ils proposent aussi l'installateur Npcap), renommer le en `npcap-installer.exe` et placer dans `Tools\Nmap\`.

Ci-dessous la liste des fichiers necessaire au fonctionnement :

```
Tools\
└─ Nmap\
    ├─ nmap.exe
    ├─ npcap-installer.exe
    ├─ ncat.exe
    ├─ nping.exe
    ├─ libcrypto-3.dll
    ├─ libssl-3.dll
    ├─ libssh2.dll
    ├─ zlibwapi.dll
    ├─ nmap-os-db
    ├─ nmap-services
    ├─ nmap-service-probes
    ├─ nmap-mac-prefixes
    ├─ nmap-protocols
    ├─ nmap-rpc
    ├─ nse_main.lua
    ├─ scripts\
    └─ nselib\
```

### 3. Lancer

```powershell
# Depuis la racine du projet, en tant qu'Administrateur
.\main.ps1
```

`init.ps1` s'exécute automatiquement au premier lancement et installe les dépendances manquantes.

---

## Structure du projet

```
NewAuditroV2\
├─ main.ps1                             ← orchestrateur principal
├─ init.ps1                             ← prérequis : Python, Graph, Npcap, templates
├─ modules_config.json                  ← config injection CSV → Excel
├─ Process_excels.py                    ← injection CSV → Excel
├─ Scripts\
│   ├─ PowerShell\
│   │   ├─ 00_Reconnaissance.ps1        ← scan réseau + audit distant WinRM
│   │   ├─ 10_Audit_comptes_T0.ps1      ← comptes Tier 0 / adminCount=1
│   │   ├─ 11_Audit_comptes.ps1         ← tous les comptes AD ou locaux
│   │   ├─ 12_Audit_Comptes_365.ps1     ← comptes Microsoft 365 via Graph
│   │   ├─ 13_Arborescence_OU.ps1       ← arborescence des OUs AD
│   │   ├─ 20_Audit_VMs.ps1             ← inventaire VM (OS, IP, WinRM, HV)
│   │   ├─ 30_Audit_Replication.ps1     ← état réplication AD (repadmin)
│   │   ├─ 40_Audit_Hyperviseurs.ps1    ← inventaire hyperviseurs Hyper-V
│   │   ├─ 50_Audit_NTFS.ps1            ← permissions NTFS sur partages
│   │   └─ 60_Audit_GPO.ps1             ← inventaire et liaison des GPO
│   └─ Python\
│       └─ Template*.py                 ← générateurs de templates Excel
├─ Templates_Excel\                     ← templates .xlsx générés par init.ps1
├─ Tools\
│   └─ Nmap\                            ← Nmap portable + npcap-installer.exe
├─ Temp_CSV\                            ← CSV collectés (Comptes\, HV\, VM\)
├─ Excel\                               ← Excel finaux générés par Process_excels.py
└─ Logs\                                ← transcripts et logs Nmap par étape
```

---

## Utilisation

### Lancement standard

```powershell
# En tant qu'Administrateur
.\main.ps1
```

Le framework exécute automatiquement dans l'ordre :

1. `init.ps1` — vérifie et installe les prérequis
2. Demande interactive des chemins NTFS à auditer
3. `00_Reconnaissance.ps1` — scan réseau interactif (optionnel)
4. `10_` à `60_` — scripts d'audit locaux

### Options de main.ps1

```powershell
# Inclure uniquement certains scripts
.\main.ps1 -Include "10_*","20_*"

# Exclure des scripts
.\main.ps1 -Exclude "main.ps1","00_*","12_*"

# Arrêter au premier script en erreur
.\main.ps1 -StopOnErr

# Mode simulation (liste les scripts sans les exécuter)
.\main.ps1 -DryRun
```

### Désactiver un script

Préfixer le fichier avec `_` — il sera ignoré par `main.ps1` :

```powershell
Rename-Item "Scripts\PowerShell\30_Audit_Replication.ps1" "_30_Audit_Replication.ps1"
```

---

## Modules d'audit

### `10_Audit_comptes_T0.ps1`
Collecte les comptes avec `adminCount=1` (Tier 0) ou les comptes locaux si pas de DC. Exporte : nom, type, date création, statut, groupes, services, tâches planifiées.

### `11_Audit_comptes.ps1`
Tous les comptes AD (hors Tier 0) ou locaux. Mêmes champs que T0.

### `12_Audit_Comptes_365.ps1`
Comptes Microsoft 365 via `Connect-MgGraph`. Exporte : UPN, licences, date dernière connexion, date dernier changement mot de passe, applications activées.

> ⚠️ Requiert une authentification interactive MgGraph. Non exécuté sur les cibles distantes WinRM.

### `13_Arborescence_OU.ps1`
Génère un fichier `.txt` représentant l'arborescence des OUs Active Directory.

### `20_Audit_VMs.ps1`
Inventaire de la VM : OS, IP, MAC, S/N, rôles installés, hyperviseur hôte (via KVP Hyper-V), statut WinRM, réplication AD, activation Windows.

### `30_Audit_Replication.ps1`
Lit le CSV de `20_Audit_VMs` et teste la réplication AD via `repadmin /showrepl` pour chaque VM listée.

> Dépendance : nécessite que `20_Audit_VMs.ps1` ait été exécuté au préalable.

### `40_Audit_Hyperviseurs.ps1`
Inventaire de l'hyperviseur Hyper-V : IPs physiques/vEthernet, MACs, rôles, processeurs, RAM, stockage, VMs hébergées, appartenance cluster.

### `50_Audit_NTFS.ps1`
Audit des permissions NTFS sur les dossiers spécifiés. Activé via la variable `AUDIT_NTFS_PATHS` ou la saisie interactive au lancement de `main.ps1`.

```powershell
# Prédéfinir les chemins avant lancement
$env:AUDIT_NTFS_PATHS = "C:\Partages\RH;C:\Partages\IT"
.\main.ps1
```

### `60_Audit_GPO.ps1`
Inventaire des GPO du domaine : créateur, date de création, OUs liées, statut. Nécessite RSAT GPMC.

---

## Reconnaissance réseau

`00_Reconnaissance.ps1` effectue un scan réseau en 3 étapes avant d'auditer les cibles distantes.

### Étapes

```
Étape 1 — Ping sweep          nmap -sn -T4 --open <plages>
Étape 2 — OS fingerprinting   nmap -O --osscan-guess -T4 <IPs>
           Fallback            nmap -p 22,445,3389,5985,5986 --open -T4 <IPs>
Étape 3 — Test WinRM           nmap -p 5985,5986 --open -T4 <IPs>
```

### Déroulement

1. Avertissement et confirmation obligatoire avant tout scan
2. Saisie interactive des plages réseau (une par ligne)
3. Scan par étapes avec logs séparés dans `Logs\Reconnaissance_<runId>\`
4. Validation utilisateur serveur par serveur (inclusion / skip avec raison)
5. Export CSV `Temp_CSV\Reconnaissance_<runId>.csv`
6. Exécution distante sur les cibles validées

### Logs générés

```
Logs\Reconnaissance_YYYYMMDD_HHMMSS\
├─ step1_scan_updown.log
├─ step1_hosts_up.txt
├─ step2_scan_os.log
└─ step3_scan_winrm.log
```

---

## Exécution distante WinRM

Quand une cible est validée et que WinRM est actif, le framework :

1. Ajoute l'IP dans `TrustedHosts` si nécessaire
2. Tente une connexion Kerberos (FQDN) ou demande des credentials (IP nue / hors domaine)
3. Copie les scripts d'audit dans `C:\Windows\Temp\Audit_<runId>\` sur la cible
4. Exécute chaque script avec les variables d'environnement `AUDIT_TEMP_*`
5. Rapatrie les CSV dans `Temp_CSV\<hostname>\`
6. Nettoie le dossier temporaire distant

### Scripts exclus de l'exécution distante

| Script | Raison |
|---|---|
| `00_Reconnaissance.ps1` | Local uniquement — évite un scan non contrôlé depuis la cible |
| `12_Audit_Comptes_365.ps1` | Requiert authentification MgGraph interactive |

---

## Sorties

### CSV collectés (`Temp_CSV\`)

| Fichier | Script source |
|---|---|
| `<HOST>_Audit_comptes_T0.csv` | `10_Audit_comptes_T0.ps1` |
| `<HOST>_Audit_comptes.csv` | `11_Audit_comptes.ps1` |
| `<HOST>_Audit_Comptes_365.csv` | `12_Audit_Comptes_365.ps1` |
| `<HOST>_AD_OU_Tree.txt` | `13_Arborescence_OU.ps1` |
| `<HOST>_Audit_VMs.csv` | `20_Audit_VMs.ps1` |
| `<HOST>_Audit_Replication.csv` | `30_Audit_Replication.ps1` |
| `<HOST>_Audit_Hyperviseurs.csv` | `40_Audit_Hyperviseurs.ps1` |
| `<HOST>_Audit_NTFS.csv` | `50_Audit_NTFS.ps1` |
| `<HOST>_Audit_GPO.csv` (dans Comptes\) | `60_Audit_GPO.ps1` |
| `Reconnaissance_<runId>.csv` | `00_Reconnaissance.ps1` |

### Excel générés (`Excel\`)

Lancé manuellement après collecte :

```powershell
python Process_excels.py
# Options
python Process_excels.py --force   # recopie les templates même si Excel existe
python Process_excels.py --ts      # ajoute un timestamp aux noms de fichiers
```

### Logs (`Logs\`)

| Fichier | Contenu |
|---|---|
| `Audit_<runId>.log` | Transcript complet de main.ps1 |
| `Init_<runId>.log` | Transcript de init.ps1 |
| `<Script>_<runId>.log` | Sortie de chaque script d'audit |
| `Reconnaissance_<runId>\` | Logs Nmap par étape |

---

## Configuration

### `modules_config.json`

Définit le mapping CSV → Excel pour `Process_excels.py` :

```json
{
  "modules": [
    {
      "name": "comptes",
      "template": "Templates_Excel/Template_Audit_Comptes.xlsx",
      "out_file": "Excel/Audit_Comptes.xlsx",
      "sheet": "Comptes",
      "csv_glob": "Temp_CSV/Comptes/*Audit_comptes.csv",
      "start_row": 1,
      "header_map": {}
    }
  ]
}
```

Le champ `header_map` permet de renommer les colonnes du CSV pour les faire correspondre aux en-têtes du template Excel (utilisé notamment pour `12_Audit_Comptes_365`).

### Chemins NTFS

```powershell
# Dans main.ps1 ou avant le lancement
$env:AUDIT_NTFS_PATHS = "C:\Partages\RH;C:\Partages\IT;C:\Partages\Commun"
```

---

## Compatibilité

| Environnement | Statut |
|---|---|
| Windows Server 2016/2019/2022 | ✅ Testé |
| Windows Server 2012 R2 | ⚠️ Non testé |
| PowerShell 5.1 | ✅ Requis minimum |
| PowerShell 7+ | ✅ Compatible |
| DC Active Directory | ✅ Toutes fonctionnalités |
| Serveur membre de domaine | ✅ Fonctionnalités RSAT selon installation |
| Hyperviseur Hyper-V standalone | ✅ Scripts non-DC skippés proprement |
| Machine hors domaine (WORKGROUP) | ✅ Audit local uniquement |

---

## Licence

Ce projet est à usage interne. Toute redistribution doit mentionner l'auteur original.

---

*Framework développé pour les audits de sécurité Active Directory et infrastructure Windows.*