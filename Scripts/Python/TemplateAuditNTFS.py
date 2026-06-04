import os

# Repertoire de sortie :
#   - AUDIT_TEMPLATES_DIR si defini (appel depuis init.ps1)
#   - Sinon : dossier Templates_Excel a la racine du framework (2 niveaux au-dessus de Scripts/Python)
_script_dir   = os.path.dirname(os.path.abspath(__file__))
_root_dir     = os.path.normpath(os.path.join(_script_dir, '..', '..'))
_default_dir  = os.path.join(_root_dir, 'Templates_Excel')
output_dir    = os.environ.get('AUDIT_TEMPLATES_DIR', _default_dir)
os.makedirs(output_dir, exist_ok=True)

output_file = os.path.join(output_dir, "Template_Audit_NTFS.xlsx")

# sinon dossier Templates_Excel a cote du script (execution manuelle)

# sinon ecrit dans le meme dossier que ce script (execution manuelle)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()

# Feuille 1 : NTFS
ws = wb.active
ws.title = "NTFS"

headers = [
    "Folder",               # A
    "Depth",                # B
    "IdentityReference",    # C
    "FileSystemRights",     # D
    "IsInherited",          # E
]

for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = Font(bold=True, color="FFFFFF", name="Arial")
    cell.fill = PatternFill("solid", fgColor="4F81BD")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[cell.column_letter].width = 30 if col == 1 else 20

# Largeur spécifique colonne B (Depth : courte)
ws.column_dimensions["B"].width = 10

table = Table(displayName="Table_NTFS", ref="A1:E200")
style = TableStyleInfo(
    name="TableStyleMedium9", showFirstColumn=False,
    showLastColumn=False, showRowStripes=True, showColumnStripes=False)
table.tableStyleInfo = style
ws.add_table(table)

# Feuille 2 : Référentiels
ref = wb.create_sheet("Référentiels")

inherited_list = ["Oui", "Non"]
rights_list    = [
    "FullControl", "Modify", "ReadAndExecute", "Read",
    "Write", "ListDirectory", "Delete", "TakeOwnership", "ChangePermissions"
]

ref.append(["IsInherited"]     + inherited_list)
ref.append(["FileSystemRights"] + rights_list)

# Listes déroulantes
dv_inherited = DataValidation(type="list", formula1="Référentiels!$B$1:$C$1", allow_blank=True)
ws.add_data_validation(dv_inherited)
dv_inherited.add("E2:E200")

dv_rights = DataValidation(type="list", formula1="Référentiels!$B$2:$J$2", allow_blank=True)
ws.add_data_validation(dv_rights)
dv_rights.add("D2:D200")

wb.save(output_file)
print(f"Fichier généré : {output_file}")