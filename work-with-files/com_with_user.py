from pathlib import Path
from work_with_exel import is_sheet_exist


def get_path(word: str, parent_folder=None) -> Path:
    n_p = 'name' if parent_folder else 'path'
    inp = input(f'Write {n_p} of the {word}: ')
    # if parent_folder:
    #     path = parent_folder.joinpath(inp)
    path = parent_folder.joinpath(inp) if parent_folder else Path(inp)
    while True:
        if path.exists():
            break
        else:
            inp = input(f'Write {n_p} of the {word}: ')
            path = parent_folder.joinpath(inp) if parent_folder else Path(inp)
    return path

def get_sheet_name(exel_path: Path, word: str) -> str:
    name = input(f'Write name of the exel sheet with {word}: ')
    while not is_sheet_exist(exel_path, name):
        name = input(f'There no this sheet in the file. Write name of the exel sheet with {word}: ')
    return name

def ask_to_crate_worksheet(exel_path: Path, sheet_name: str) -> str:
    while not is_sheet_exist(exel_path, sheet_name):
        input(f'There no list {sheet_name} in the {exel_path.name}, please create it, save and press Enter')
    pass

def is_setting_confirm(category, value) -> bool:
    confirmation = input(f'\n{category}: {value}\nIf you confirm print "yes", else - "no" and press Enter: ')
    while True:
        confirmation = confirmation.lower().strip()
        if confirmation == 'yes':
            return True
        elif confirmation == 'no':
            return False
        else:
            confirmation = input('Sorry, write "yes" to confirm, or "no" in others cases: ')







