import tkinter as tkt
from tkinter import messagebox, filedialog
import funcs_for_settings as ffs
from pathlib import Path
from main_functions import make_word_document, make_template

MAIN_FUNCS = {
    'exel': make_word_document,
    'template Word': make_template
}

DEFOULT_SETTINGS_FUNC = {
    'exel': ffs.get_defoult_by_pattern,
    'word': ffs.get_defoult_by_pattern,
    'sheet with information': ffs.get_defoult_infosheet,
    'sheet with charts': ffs.get_defoult_chartsheet,
    'result file name': ffs.no_defoult_settings,
    'result folder': ffs.get_defoult_resultfolder,
    'template Word': ffs.get_defoult_by_pattern,
    'folder with tamplates': ffs.get_defoult_by_pattern
}

DEFOULT_PATTERNS = {
    'word': f'*{ffs.TEMPLATE_SYMBOL}*.docx',
    'template Word': f'*{ffs.TEMPLATE_SYMBOL}*.docx',
    'exel': f'*{ffs.TEMPLATE_SYMBOL}*.xlsx',
    'folder with tamplates': ffs.GRAPH_SYMBOL
}

FILES_GROUPS = {
    'word': (filedialog.askopenfilename, {'filetypes': [('Word documents', '*.docx')]}),
    'template Word': (filedialog.askopenfilename, {'filetypes': [('Word documents', '*.docx')]}),
    'exel': (filedialog.askopenfilename, {'filetypes': [('Exel documents', '*.xlsx')]}),
    'folder with tamplates': (filedialog.askdirectory, {})
}

# KEY_ELEMENTS = {'exel', 'template Word'}

CONFIRMATION = {
    'exel': ffs.is_path_exist,
    'word': ffs.is_path_exist,
    'sheet with information': ffs.is_sheet,
    'sheet with charts': ffs.is_sheet,
    'result file name': ffs.no_checking,
    'result folder': ffs.no_checking,
    'template Word': ffs.is_path_exist,
    'folder with tamplates': ffs.is_path_exist
}

BASE = {
    'exel': ffs.folder_base,
    'word': ffs.folder_base,
    'sheet with information': ffs.file_base,
    'sheet with charts': ffs.file_base,
    'result file name': ffs.folder_base,
    'result folder': ffs.folder_base,
    'template Word': ffs.folder_base,
    'folder with tamplates': ffs.folder_base
}

def choose_func(element, hendler):
    return hendler.get(element)

def make_fields(settings, root: tkt.Tk):
    field_variables = {}
    # option_menus = {}
    # exel_sheets = []

    def get_key_elements():
        key_element, *_ = set(MAIN_FUNCS) & set(settings)
        file = field_variables[key_element].get()
        file = Path(file)
        base = ffs.folder_base(file)
        return file, base, key_element
    
    def get_settings():
        file, base, key_element = get_key_elements()
        is_ok = True
        for element, var in field_variables.items():
            value= var.get()
            func = choose_func(element, CONFIRMATION)
            base = choose_func(element, BASE)(file)
            comfirmation = func(value, base)
            if comfirmation == True:
                settings[element] = value
            else:
                error = f'{element.upper()}: irrelevant value'
                messagebox.showerror(title='ERROR', message=error)
                is_ok = False
                break
        if is_ok:
            responce = messagebox.askokcancel(message='You realy confirm it?')    
            if responce:
                main_func = choose_func(key_element, MAIN_FUNCS)
                main_func(settings)
        pass
    
    def reset():
        for var in field_variables.values():
            var.set('')
        pass
    
    def entry_rewrite(element, value):
        entry = root.children[f'e_{element}']
        entry.delete(0, tkt.END)
        if value:
            entry.insert(0, value)
        pass

    def var_rewrite(element, value):
        var = field_variables.get(element)
        value = value if value else ''
        var.set(value)
        pass

    def update(*args):
        file, base, *_ = get_key_elements()
        for element in filter(lambda x: x not in MAIN_FUNCS, settings):
            func = choose_func(element, DEFOULT_SETTINGS_FUNC)
            pattern = DEFOULT_PATTERNS.get(element)
            base = choose_func(element, BASE)(file)
            defoult_value = func(base, pattern)
            var_rewrite(element, defoult_value)
        pass

    def get_path(element):
        def inner():
            func = FILES_GROUPS[element][0]
            kwargs = FILES_GROUPS[element][1]
            selected = func(**kwargs)
            selected_path = Path(selected)
            var_rewrite(element, selected_path)
            pass
        return inner  
    
    def make_label(position, element):
        label = tkt.Label(root, name=f'l_{element}', text=element)
        label.grid(row=position, column=0, sticky=tkt.E)
        pass

    def make_button(position, element):
        command = get_path(element)
        button = tkt.Button(root, text='Choose', name=f'b_{element}', command=command)
        button.grid(row=position, column=2)
        pass

    def make_var(element):
        var = tkt.StringVar(root, name=f'v_{element}')
        field_variables[element]=var
        if element in MAIN_FUNCS:
            var.trace('w', update)
        return var

    def make_entry(position, element, var):
        entry = tkt.Entry(root, borderwidth=2, width=60, name=f'e_{element}', textvariable=var)
        entry.grid(row=position, column=1, sticky=tkt.W, padx = 2)
        pass
        
    def make_option_menu(position, element, var, values):
        option_menu = tkt.OptionMenu(root, var, values)
        option_menu.grid(row=position, column=1, sticky=tkt.W, padx = 2)
        option_menus[element] = option_menu
        pass

    def make_row(position, element):
        make_label(position, element)
        var = make_var(element)
        # if BASE[element] == file_base:
        #     make_option_menu(position, element, var, exel_sheets)
        # else:
        make_entry(position, element, var)
        if element in DEFOULT_PATTERNS:
            make_button(position, element)
        pass
    
    for i, k in enumerate(settings):
        make_row(i, k)

    confirm_button = tkt.Button(root, text='Confirm', name='confirm', command= get_settings)
    confirm_button.grid(row=i+1, column=2, padx=2)
    confirm_button = tkt.Button(root, text='Reset', name='reset', command= reset)
    confirm_button.grid(row=i+1, column=1, pady=2, sticky=tkt.E, padx=2)
    pass
   
def settings_root(settings):     
    main_root = tkt.Tk()
    main_root.title('Settings')
    main_root.resizable(width=0, height=0)

    work_frame = tkt.Frame(main_root)
    work_frame.grid(row=0, column=0)
    make_fields(settings, work_frame)
    main_root.mainloop()

    pass

# settings_root(MAIN_SETTINGS)
# settings_root(TEMPLATE_SETTINGS)
