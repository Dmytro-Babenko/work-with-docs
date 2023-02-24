from pathlib import Path
from work_with_exel import is_sheet_exist, first_sheet_name
from openpyxl import load_workbook

GRAPH_SYMBOL = 'gr'
TEMPLATE_SYMBOL = 'бланк'
RESULT_FOLDER_NAME = 'виконані'
EXTENTION = ['Sheet with charts']

MAIN_SETTINGS = {
    'base_folder': None,
    'Exel': None,
    'Word': None,
    'Sheet with information': None,
    'Sheet with charts': None,
    'Result file name': None,
    'Result folder': None
}

BASE = {
    'Exel': 'base_folder',
    'Word': 'base_folder',
    'Sheet with information': 'Exel',
    'Sheet with charts': 'Exel',
    'Result folder': 'base_folder'
}

def confirmation(element, value) -> bool:

    if value == None:
        return False
    
    if element in EXTENTION:
        return True
    
    confirmation = input(f'\n{element}: {value}\nIf you confirm print "yes", else - "no" and press Enter: ')
    while True:
        confirmation = confirmation.lower().strip()
        if confirmation == 'yes':
            return True
        elif confirmation == 'no':
            return False
        else:
            confirmation = input('Sorry, write "yes" to confirm, or "no" in others cases: ')

def existing_path(func):
    def inner(element, *args, **kwargs):
        output = func(element, *args, **kwargs)
        while not output.exists():
            print('There no file with such path')
            output = func(element,  *args, **kwargs
            )
        return output
    return inner

def existing_sheet(func):
    def inner(element, *args, **kwargs):
        output = func(element, *args, **kwargs)
        while not output:
            print('There no such sheet on exel file')
            output = func(element,  *args, **kwargs)
        return output
    return inner

@existing_path
def get_path_by_name(element, base, *args):
    inp = input(f'\nWrite name of the {element}: ')
    path = base.joinpath(inp)
    return path

@existing_path
def get_fullpath(element, *args):
    path = Path(input(f'\nWrite path of the {element}: '))
    return path

@existing_sheet
def get_sheet_name(element, base, *args):
    sheet_name = input(f'\nWrite name of the {element}: ')
    if not is_sheet_exist(base, sheet_name):
        sheet_name = None
    return sheet_name

def rename_sheet(element, base, new_name, *args):
    book = load_workbook(base)
    old_name = get_sheet_name(element, base)
    book[old_name].title = new_name
    book.save(base)
    book.close()
    return new_name

def get_name(element, *args):
    name = input(f'\n{element}: ')
    return name
    
USER_SETTINGS_FUNC = {
    'base_folder': get_fullpath,
    'Exel': get_path_by_name,
    'Word': get_path_by_name,
    'Sheet with information': get_sheet_name,
    'Sheet with charts': rename_sheet,
    'Result file name': get_name,
    'Result folder': get_name
}
    
def get_defoult_doc(base: Path):
    pattern = f'*{TEMPLATE_SYMBOL}*.docx'
    output = next(base.glob(pattern))
    return output

def get_defoult_exel(base:Path):
    pattern = f'*{TEMPLATE_SYMBOL}*.xlsx'
    output = next(base.glob(pattern))
    return output

def get_defoult_infosheet(base:Path):
    return first_sheet_name(base)

def get_defoult_chartsheet(base:Path):
    if is_sheet_exist(base, GRAPH_SYMBOL):
        return GRAPH_SYMBOL
    return None

def get_defoult_resultfolder(base: Path):
    output = RESULT_FOLDER_NAME
    return output

def no_defoult_settings(base: Path):
    return None

DEFOULT_SETTINGS_FUNC = {
    'base_folder': no_defoult_settings,
    'Exel': get_defoult_exel,
    'Word': get_defoult_doc,
    'Sheet with information': get_defoult_infosheet,
    'Sheet with charts': get_defoult_chartsheet,
    'Result file name': no_defoult_settings,
    'Result folder': get_defoult_resultfolder
}

def choose_func(element, hendler):
    return hendler.get(element)

def get_settings(settings: dict):
    
    def choose_base(element):
        base = settings.get(BASE.get(element))
        return base
    
    for element in settings:
        base = choose_base(element)
        get_defoult_value = choose_func(element, DEFOULT_SETTINGS_FUNC)
        settings[element] = get_defoult_value(base)
        value = settings[element]
        is_confirm = confirmation(element, value)
        if not is_confirm:
            get_user_value = choose_func(element, USER_SETTINGS_FUNC)
            settings[element] = get_user_value(element, base, GRAPH_SYMBOL)
    return settings

MAIN_SETTINGS = get_settings(MAIN_SETTINGS)

