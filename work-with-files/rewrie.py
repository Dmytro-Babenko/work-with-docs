from pathlib import Path
from work_with_exel import is_sheet_exist
from openpyxl import load_workbook
arr = ['1', '2', '3', '4', '5']

GRAPH_SYMBOL = 'gr'
TEMPLATE_SYMBOL = 'бланк'
RESULT_FOLDER_NAME = 'виконані'


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
}

def get_unique(element, base):
    def options(func):
        def inner(arr, *args, **kwargs):
            output = func(arr)
            while output not in arr:
                output = input(f'Please write {element} of the {base}')
            return output
        return inner
    return options

@get_unique('name', 'doc')
def first_input(arr):
    a = input('Write name that: ')
    return a
# print(first_input(arr))

# @get_unique('name', 'doc')
# def sheet_name()

# def make_dir(func):
#     def inner(path: Path, *args):
#         if not path.exists():
#             path.mkdir()
#         func(path)
#     return inner

# path = Path(r'C:\Users\Lenovo\Desktop\заказы\new')

# @make_dir
# def print_folder(path: Path):
#     for item in path.iterdir():
#         print(item)
#     pass

def get_settings(settings):
    def confirmation(func) -> bool:
            def inner(element, *args, **kwargs):
                value = settings[element]
                if value == None:
                    return func(element, *args, **kwargs)
                confirmation = input(f'\n{element}: {value}\nIf you confirm print "yes", else - "no" and press Enter: ')
                while True:
                    confirmation = confirmation.lower().strip()
                    if confirmation == 'yes':
                        return value
                    elif confirmation == 'no':
                        return func(element, *args, **kwargs)
                    else:
                        confirmation = input('Sorry, write "yes" to confirm, or "no" in others cases: ')
            return inner


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

    def choose_base(element, settings):
        base = settings.get(BASE.get(element))
        return base

    @confirmation
    @existing_path
    def get_path_by_name(element, base, *args):
        inp = input(f'Write name of the {element}: ')
        path = base.joinpath(inp)
        return path

    @confirmation
    @existing_path
    def get_fullpath(element, *args):
        path = Path(input(f'Write path of the {element}: '))
        return path

    @confirmation
    @existing_sheet
    def get_sheet_name(element, base, *args):
        sheet_name = input(f'Write name of the {element}: ')
        if sheet_name not in load_workbook(base).sheetnames:
            sheet_name = None
        return sheet_name

    @confirmation
    def rename_sheet(element, base, new_name, *args):
        book = load_workbook(base)
        old_name = get_sheet_name(element, base)
        book[old_name].name = new_name
        book.save(str(base.absolute()))
        book.close()
        return new_name

    @confirmation
    def get_name(element, *args):
        name = input(f'{element}: ')
        return name
        

    hendler = {
        'base_folder': get_fullpath,
        'Exel': get_path_by_name,
        'Word': get_path_by_name,
        'Sheet with information': get_sheet_name,
        'Sheet with charts': rename_sheet,
        'Result file name': get_name
    }

    def choose_func(element, hendler):
        return hendler[element]
        

    def get_defoult_doc(base: Path):
        pattern = f'*{TEMPLATE_SYMBOL}*.docx'
        output = next(base.glob(pattern))
        return output





path = Path(r'D:\test').glob('*бланк*.docx')
print(next(path))
for i in path:
    print(i)

Path(r'D:\test\1').mkdir(exist_ok=True)