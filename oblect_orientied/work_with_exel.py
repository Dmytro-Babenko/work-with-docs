from pathlib import Path
from openpyxl import load_workbook

def first_sheet_name(exel_file: Path) -> str:
    '''Return name of the first sheet in Exel file'''
    wb = load_workbook(exel_file, read_only=True)
    sheet_name = wb.sheetnames[0]
    wb.close()
    return sheet_name

def is_sheet_exist(exel_file: Path, sheet_name: str) -> bool:
    '''Return True if sheet exists, else False'''
    wb = load_workbook(exel_file, read_only=True)
    if sheet_name in wb.sheetnames:
        wb.close()
        return True
    else:
        wb.close()
        return False
        
def existing_sheet(func):
    def inner(element, *args, **kwargs):
        output = func(element, *args, **kwargs)
        while not output:
            print('There no such sheet on exel file')
            output = func(element,  *args, **kwargs)
        return output
    return inner

@existing_sheet
def get_sheet_name(element, base, *args):
    sheet_name = input(f'\nWrite name of the {element}: ')
    if not is_sheet_exist(base, sheet_name):
        sheet_name = None
    return sheet_name

def rename_sheet(element, base, new_name, *args):
    '''Rename Exel sheet from old name to new'''
    book = load_workbook(base)
    old_name = get_sheet_name(element, base)
    book[old_name].title = new_name
    book.save(base)
    book.close()
    return new_name


def sheet_names(exel_file: Path) -> str:
    '''Return names list of the sheets in Exel file'''
    wb = load_workbook(exel_file, read_only=True)
    sheet_names = wb.sheetnames
    wb.close()
    return sheet_names

