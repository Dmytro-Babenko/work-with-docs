from pathlib import Path
from work_with_exel import is_sheet_exist, first_sheet_name

GRAPH_SYMBOL = 'gr'
TEMPLATE_SYMBOL = 'бланк'
RESULT_FOLDER_NAME = 'виконані'
EXTENTION = ['Sheet with charts']
    
# def get_defoult_doc(base: Path):
#     pattern = f'*{TEMPLATE_SYMBOL}*.docx'
#     output = next(base.glob(pattern))
#     return output

# def get_defoult_exel(base:Path):
#     pattern = f'*{TEMPLATE_SYMBOL}*.xlsx'
#     output = next(base.glob(pattern))
#     return output

# def get_defoult_gr_folder(base:Path):
#     pattern = GRAPH_SYMBOL
#     output = next(base.glob(pattern))
#     return output

def get_defoult_by_pattern(base: Path, pattern: str):
    try:
        output = next(base.glob(pattern))
    except:
        return None
    return output

def get_defoult_infosheet(base:Path, *args):
    if base.suffix == '.xlsx':
        return first_sheet_name(base)
    return None

def get_defoult_chartsheet(base:Path, *args):
    if base.suffix == '.xlsx' and is_sheet_exist(base, GRAPH_SYMBOL):
        return GRAPH_SYMBOL
    return None

def get_defoult_resultfolder(*args):
    output = RESULT_FOLDER_NAME
    return output

def no_defoult_settings(*args):
    return None

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

@existing_path
def get_path_by_name(element, base, *args):
    inp = input(f'\nWrite name of the {element}: ')
    path = base.joinpath(inp)
    return path

@existing_path
def get_fullpath(element, *args):
    path = Path(input(f'\nWrite path of the {element}: '))
    return path

def get_name(element, *args):
    name = input(f'\n{element}: ')
    return name





