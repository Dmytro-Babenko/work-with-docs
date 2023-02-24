from work_with_exel import get_sheet_name, rename_sheet
import funcs_for_settings

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
   
USER_SETTINGS_FUNC = {
    'base_folder': funcs_for_settings.get_fullpath,
    'Exel': funcs_for_settings.get_path_by_name,
    'Word': funcs_for_settings.get_path_by_name,
    'Sheet with information': get_sheet_name,
    'Sheet with charts': rename_sheet,
    'Result file name': funcs_for_settings.get_name,
    'Result folder': funcs_for_settings.get_name
}

DEFOULT_SETTINGS_FUNC = {
    'base_folder': funcs_for_settings.no_defoult_settings,
    'Exel': funcs_for_settings.get_defoult_exel,
    'Word': funcs_for_settings.get_defoult_doc,
    'Sheet with information': funcs_for_settings.get_defoult_infosheet,
    'Sheet with charts': funcs_for_settings.get_defoult_chartsheet,
    'Result file name': funcs_for_settings.no_defoult_settings,
    'Result folder': funcs_for_settings.get_defoult_resultfolder
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
        is_confirm = funcs_for_settings.confirmation(element, value)
        if not is_confirm:
            get_user_value = choose_func(element, USER_SETTINGS_FUNC)
            settings[element] = get_user_value(element, base, funcs_for_settings.GRAPH_SYMBOL)
    return settings

MAIN_SETTINGS = get_settings(MAIN_SETTINGS)

