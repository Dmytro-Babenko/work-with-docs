from pathlib import Path
from com_with_user import get_path, get_sheet_name, ask_to_crate_worksheet, is_setting_confirm
from work_with_exel import is_sheet_exist, first_sheet_name, rename_sheet

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

# констнта типов 
# функция исходя из типа определяет функцию которую надо использовать для определения настройки

def reset_settings(settings: dict, *exeptions) -> dict:
    for key in settings.keys():
        if key not in exeptions:
            settings[key] = None
    return settings

def get_main_default_settings(main_settings: dict) -> dict:
    main_settings['base_folder'] = get_path('main folder')
    for file in main_settings['base_folder'].iterdir():
        if file.match(f'{TEMPLATE_SYMBOL}*.*'):
            if file.suffix == '.docx':
                main_settings['Word'] = file
            elif file.suffix == '.xlsx':
                main_settings['Exel'] = file
                main_settings['Sheet with information'] = first_sheet_name(file)
                if is_sheet_exist(file, GRAPH_SYMBOL):
                    main_settings['Sheet with charts'] = GRAPH_SYMBOL
    main_settings['Result folder'] = main_settings['base_folder'].joinpath(RESULT_FOLDER_NAME) 
    if not main_settings['Result folder'].exists():
        main_settings['Result folder'].mkdir()
    return main_settings

def get_user_setting(category, settings: dict): # функция возвращает функцию переписать, для разных параметров переписать через дикораторы
    if category == 'Sheet with charts':
        old_name = get_sheet_name(settings['Exel'], 'charts')
        rename_sheet(settings['Exel'], old_name, GRAPH_SYMBOL)
        return GRAPH_SYMBOL

    if category == 'Sheet with information':
        return get_sheet_name(settings['Exel'], 'text info')

    if category == 'Result file name':
        return input('\nWrite name of the final file: ')

    else:
        return get_path(category, settings['base_folder'])

def get_user_settings(settings: dict, *extention) -> dict:
    print('\nPlease check settings')
    for category, value in settings.items():
        value = value.name if isinstance(value, Path) else value

        if value == None:
            settings[category] = get_user_setting(category, settings)

        elif category not in extention and not is_setting_confirm(category, value):
            settings[category] = get_user_setting(category, settings)

    return settings

def get_main_user_settings():
    main_settings = get_main_default_settings(MAIN_SETTINGS)
    main_settings = get_user_settings(main_settings, 'base_folder', 'Sheet with charts', 'Result folder')
    return main_settings
        

            



        

    