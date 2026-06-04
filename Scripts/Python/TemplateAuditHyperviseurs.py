import os

# Repertoire de sortie :
#   - AUDIT_TEMPLATES_DIR si defini (appel depuis init.ps1)
#   - Sinon : dossier Templates_Excel a la racine du framework (2 niveaux au-dessus de Scripts/Python)
_script_dir   = os.path.dirname(os.path.abspath(__file__))
_root_dir     = os.path.normpath(os.path.join(_script_dir, '..', '..'))
_default_dir  = os.path.join(_root_dir, 'Templates_Excel')
output_dir    = os.environ.get('AUDIT_TEMPLATES_DIR', _default_dir)
os.makedirs(output_dir, exist_ok=True)

output_file = os.path.join(output_dir, "Template_Audit_Hyperviseurs.xlsx")

# sinon dossier Templates_Excel a cote du script (execution manuelle)

# sinon ecrit dans le meme dossier que ce script (execution manuelle)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
# Dossier de sortie relatif au script

wb = Workbook()

# Feuille 1 : Hyperviseurs
ws = wb.active
ws.title = "Hyperviseurs"

headers = [
    "Nom",                  # A
    "IP (Physique)",        # B
    "IP (vEthernet)",       # C
    "Adresse MAC",          # D
    "Rôles",                # E
    "Type d'hyperviseur",   # F
    "OS",                   # G
    "Intégré au domaine",   # H
    "Nom de Domaine",       # I
    "Cluster/Standalone",   # J
    "Processeurs",          # K
    "RAM",                  # L
    "Stockage",             # M
    "Nb VM hébergées"       # N
]

for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, color="FFFFFF", name="Arial")
    cell.fill = PatternFill("solid", fgColor="4F81BD")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[chr(64+col)].width = 25

table = Table(displayName="Table_Hyperviseurs", ref="A1:N10")
style = TableStyleInfo(
    name="TableStyleMedium9", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style
ws.add_table(table)

# Feuille 2 : Référentiels
ref = wb.create_sheet("Référentiels")

hyperviseur_types = ["VMware ESXi", "Microsoft Hyper-V", "Proxmox VE", "XenServer", "Nutanix AHV", "Autre"]
os_list           = ["Windows Server 2016", "Windows Server 2019", "Windows Server 2022",
                     "Linux Ubuntu", "Linux Debian", "Linux RHEL", "ESXi", "Autre"]
domain_list       = ["Oui", "Non"]
raid_list         = ["RAID 0", "RAID 1", "RAID 5", "RAID 6", "RAID 10", "JBOD", "Aucun"]
cluster_list      = ["Cluster", "Standalone"]

ref.append(["Type d'hyperviseur"] + hyperviseur_types)
ref.append(["OS"]                 + os_list)
ref.append(["Intégré au domaine"] + domain_list)
ref.append(["Stockage"]           + raid_list)
ref.append(["Cluster/Standalone"] + cluster_list)

# Listes déroulantes
dv_hyperviseur = DataValidation(type="list", formula1="Référentiels!$B$1:$G$1", allow_blank=True)
ws.add_data_validation(dv_hyperviseur)
dv_hyperviseur.add("F2:F10")

dv_os = DataValidation(type="list", formula1="Référentiels!$B$2:$I$2", allow_blank=True)
ws.add_data_validation(dv_os)
dv_os.add("G2:G10")

dv_domain = DataValidation(type="list", formula1="Référentiels!$B$3:$C$3", allow_blank=True)
ws.add_data_validation(dv_domain)
dv_domain.add("H2:H10")

dv_storage = DataValidation(type="list", formula1="Référentiels!$B$4:$H$4", allow_blank=True)
ws.add_data_validation(dv_storage)
dv_storage.add("M2:M10")

dv_cluster = DataValidation(type="list", formula1="Référentiels!$B$5:$C$5", allow_blank=True)
ws.add_data_validation(dv_cluster)
dv_cluster.add("J2:J10")

wb.save(output_file)
print(f"Fichier généré : {output_file}")