import os

# Repertoire de sortie :
#   - AUDIT_TEMPLATES_DIR si defini (appel depuis init.ps1)
#   - Sinon : dossier Templates_Excel a la racine du framework (2 niveaux au-dessus de Scripts/Python)
_script_dir   = os.path.dirname(os.path.abspath(__file__))
_root_dir     = os.path.normpath(os.path.join(_script_dir, '..', '..'))
_default_dir  = os.path.join(_root_dir, 'Templates_Excel')
output_dir    = os.environ.get('AUDIT_TEMPLATES_DIR', _default_dir)
os.makedirs(output_dir, exist_ok=True)

output_file = os.path.join(output_dir, "Template_Audit_Comptes.xlsx")

# sinon dossier Templates_Excel a cote du script (execution manuelle)

# sinon ecrit dans le meme dossier que ce script (execution manuelle)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
# Dossier de sortie relatif au script

wb = Workbook()

# Feuille 1 : Comptes
ws = wb.active
ws.title = "Comptes"

headersTab1 = [
    "Ordinateur",       # A
    "Name",             # B
    "Type",             # C
    "Date de creation", # D  ← virgule manquante corrigée
    "Active",           # E
    "Description",      # F
    "PasswordLastSet",  # G
    "Groupes",          # H
    "Services",         # I
    "ScheduledTask"     # J
]

for col, header in enumerate(headersTab1, start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, color="FFFFFF", name="Arial")
    cell.fill = PatternFill("solid", fgColor="4BACC6")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[cell.column_letter].width = 25

# Tableau structuré A1:J200 (10 colonnes réelles)
table = Table(displayName="Table_Comptes", ref="A1:J200")
style = TableStyleInfo(
    name="TableStyleMedium9", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style
ws.add_table(table)

# Feuille 2 : Référentiels
ref = wb.create_sheet("Référentiels")

yes_no         = ["Oui", "Non"]
account_types  = ["Local", "Domain", "Built-in", "gMSA", "sMSA", "Service", "Computer$", "Other"]

ref.append(["Enabled"] + yes_no)           # ligne 1 : B1:C1
ref.append(["Type"]    + account_types)    # ligne 2 : B2:I2

# Listes déroulantes
# Colonne E (Active)
dv_enabled = DataValidation(type="list", formula1="Référentiels!$B$1:$C$1", allow_blank=True)
ws.add_data_validation(dv_enabled)
dv_enabled.add("E2:E200")

# Colonne C (Type)
dv_type = DataValidation(type="list", formula1="Référentiels!$B$2:$I$2", allow_blank=True)
ws.add_data_validation(dv_type)
dv_type.add("C2:C200")

wb.save(output_file)
print(f"Fichier généré : {output_file}")