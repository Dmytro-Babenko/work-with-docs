from pathlib import Path
from openpyxl import load_workbook
from win32com.client import Dispatch

def get_info_from_exel(exel_file: Path, sheet_name='info') -> dict[str:any]:
    wb = load_workbook(exel_file, data_only=True, read_only=True)
    info = {}
    ws = wb[sheet_name]
    for k, v in ws.iter_rows(values_only=True):
        info[k] = v
    return info

def export_image(exel_file: Path, sheet_name='gr') -> dict[str:Path]:
    graphs = {}
    gr_folder = exel_file.parent.joinpath(sheet_name)
    if not gr_folder.exists():
        gr_folder.mkdir()
    app = Dispatch('Excel.Application')
    wb = app.Workbooks.Open(Filename=exel_file)
    app.DisplayAlerts = False

    i = 1
    gr_sheet = wb.Worksheets(sheet_name)
    for chartObject in gr_sheet.ChartObjects():
        gr_name = f'{sheet_name}{i}.png'
        gr_path = gr_folder.joinpath(gr_name)
        graphs[gr_name] = gr_path
        chartObject.Chart.Export(gr_path)
        i += 1
    wb.Close(SaveChanges=False, Filename=str(exel_file))
    return graphs

def first_sheet_name(exel_file: Path) -> str:
    wb = load_workbook(exel_file, read_only=True)
    sheet_name = wb.sheetnames[0]
    return sheet_name

def is_sheet_exist(exel_file: Path, sheet_name: str) -> bool:
    if sheet_name in load_workbook(exel_file, read_only=True).sheetnames:
        return True
    else:
        return False


