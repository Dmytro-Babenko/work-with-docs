import tkinter as tkt
from tkinter import messagebox, filedialog
import funcs_for_settings
from pathlib import Path
from main import make_it

MAIN_SETTINGS = {
    'exel': None,
    'word': None,
    'result folder': None,
    'sheet with information': None,
    'sheet with charts': None,
    'result file name': None,
}

DEFOULT_SETTINGS_FUNC = {
    'exel': funcs_for_settings.get_defoult_by_pattern,
    'word': funcs_for_settings.get_defoult_by_pattern,
    'sheet with information': funcs_for_settings.get_defoult_infosheet,
    'sheet with charts': funcs_for_settings.get_defoult_chartsheet,
    'result file name': funcs_for_settings.no_defoult_settings,
    'result folder': funcs_for_settings.get_defoult_resultfolder,
    'template Word': funcs_for_settings.get_defoult_by_pattern,
    'folder with tamplates': funcs_for_settings.get_defoult_by_pattern
}

DEFOULT_PATTERNS = {
    'word': f'*{funcs_for_settings.TEMPLATE_SYMBOL}*.docx',
    'template Word': f'*{funcs_for_settings.TEMPLATE_SYMBOL}*.docx',
    'exel': f'*{funcs_for_settings.TEMPLATE_SYMBOL}*.xlsx',
    'folder with tamplates': funcs_for_settings.GRAPH_SYMBOL
}

FILES_GROUPS = {
    'word': (filedialog.askopenfilename, {'filetypes': [('Word documents', '*.docx')]}),
    'template Word': (filedialog.askopenfilename, {'filetypes': [('Word documents', '*.docx')]}),
    'exel': (filedialog.askopenfilename, {'filetypes': [('Exel documents', '*.xlsx')]}),
    'folder with tamplates': (filedialog.askdirectory, {})
}

KEY_ELEMENTS = ('exel', 'template Word')

CONFIRMATION = {
    'exel': funcs_for_settings.is_path_exist,
    'word': funcs_for_settings.is_path_exist,
    'sheet with information': funcs_for_settings.is_sheet,
    'sheet with charts': funcs_for_settings.is_sheet,
    'result file name': funcs_for_settings.no_checking,
    'result folder': funcs_for_settings.no_checking,
    'template Word': funcs_for_settings.is_path_exist,
    'folder with tamplates': funcs_for_settings.is_path_exist
}

def folder_base(file: Path):
    base = file.parent
    return base

def file_base(file: Path):
    return file

BASE = {
    'exel': folder_base,
    'word': folder_base,
    'sheet with information': file_base,
    'sheet with charts': file_base,
    'result file name': folder_base,
    'result folder': folder_base,
    'template Word': folder_base,
    'folder with tamplates': folder_base
}

def choose_func(element, hendler):
    return hendler.get(element)

def make_fields(settings, root):
    def get_settings():
        root_elements = root.children.values()
        is_ok = True
        for k, v in zip(
            filter(lambda y: isinstance(y, tkt.Label), root_elements),
            filter(lambda y: isinstance(y, tkt.Entry), root_elements)):
            element = k.cget('text')
            value = v.get()
            func = choose_func(element, CONFIRMATION)
            base = choose_func(element, BASE)(Path(value_var.get()))
            result = func(value, base)
            if result == True:
                settings[element] = value
            else:
                error = f'{element.upper()}: irrelevant value'
                messagebox.showerror(title='ERROR', message=error)
                is_ok = False
                break
        if is_ok:
            responce = messagebox.askokcancel(message='you realy confirm it?')    
            if responce:
                make_it(settings)
        pass
    
    def entry_rewrite(element, value):
        entry = root.children[f'e_{element}']
        entry.delete(0, tkt.END)
        if value:
            entry.insert(0, value)
        pass

    def update(*args):
        for element in filter(lambda x: x not in KEY_ELEMENTS, settings):
            func = choose_func(element, DEFOULT_SETTINGS_FUNC)
            pattern = DEFOULT_PATTERNS.get(element)
            base = choose_func(element, BASE)(Path(value_var.get()))
            defoult_value = func(base, pattern)
            entry_rewrite(element, defoult_value)
        pass

    def get_path(element):
        def inner():
            func = FILES_GROUPS[element][0]
            kwargs = FILES_GROUPS[element][1]
            selected = func(**kwargs)
            selected_path = Path(selected)
            entry_rewrite(element, selected_path)
            pass
        return inner  
    
    def make_label(position, element):
        label = tkt.Label(root, name=f'l_{element}', text=element)
        label.grid(row=position, column=0, sticky=tkt.E)
        pass

    def make_button(position, element):
        command = get_path(element)
        button = tkt.Button(root, text='...', name=f'b_{element}', command=command)
        button.grid(row=position, column=2, padx = 2)
        pass

    def make_entry(position, element):
        if element in KEY_ELEMENTS:
            entry = tkt.Entry(root, borderwidth=2, width=40, name=f'e_{element}', textvariable=value_var)
        else:
            entry = tkt.Entry(root, borderwidth=2, width=40, name=f'e_{element}')
        entry.grid(row=position, column=1, sticky=tkt.W, padx = 2)
        pass
        
    def make_row(position, element):

        make_label(position, element)
        make_entry(position, element)
        if element in DEFOULT_PATTERNS:
            make_button(position, element)
        pass

    value_var = tkt.StringVar(root)
    value_var.trace('w', update)

    for i, k in enumerate(settings):
        make_row(i, k)
    
    confirm_button = tkt.Button(root, text='Confirm', name='confirm', command= get_settings)
    confirm_button.grid(row=i+1, column=0, columnspan=3)
    
    pass
   

def settings_root(settings):     
    main_root = tkt.Tk()
    work_frame = tkt.Frame(main_root)
    work_frame.grid(row=0, column=0)
    make_fields(settings, work_frame)
    main_root.mainloop()
    pass

settings_root(MAIN_SETTINGS)

        
# root = tkt.Tk()






# confirm = tkt.Button(text='Confirm', name='confirm')
# confirm.grid(row=i+1, column=0, columnspan=3)










# root_elements = root.children.values()
# for a, b in zip(
#     filter(lambda y: isinstance(y, tkt.Label), root_elements),
#     filter(lambda y: isinstance(y, tkt.Entry), root_elements)):
#     # print(a.cget('text'))
#     # print(b.get())






