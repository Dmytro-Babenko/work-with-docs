import tkinter as tkt
from tkinter import messagebox, ttk, filedialog
import funcs_for_settings
from pathlib import Path

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

def button_func(element):
    def get_path():
        func = FILES_GROUPS[element][0]
        kwargs = FILES_GROUPS[element][1]
        selected = func(**kwargs)
        selected_path = Path(selected)
        root.setvar(name=f'v_{element}', value=selected_path)
        MAIN_SETTINGS[element] = selected_path
        pass
    return get_path



def update(*args):
    for element in filter(lambda x: x!='exel', MAIN_SETTINGS):
        func = choose_func(element, DEFOULT_SETTINGS_FUNC)
        pattern = DEFOULT_PATTERNS.get(element)
        base = choose_func(element, BASE)(Path(a.get()))
        defoult_value = func(base, pattern)
        root.setvar(name=f'v_{element}', value='')
        # print((1, root.getvar(name=f'v_{element}')))
        if defoult_value:
            root.setvar(name=f'v_{element}', value=defoult_value)
        
root = tkt.Tk()

base_var = tkt.StringVar(name='s')
dct = {}

for i, (k, v) in enumerate(MAIN_SETTINGS.items()):

    k = k.lower()
    value_var = tkt.StringVar(name=f'v_{k}')
    dct[k]=value_var

    label = tkt.Label(root, name=f'l_{k}', text=k)
    label.grid(row=i, column=0, sticky=tkt.E)


    entry = tkt.Entry(root, borderwidth=2, width=40, name=f'e_{k}', textvariable=value_var)
    entry.grid(row=i, column=1, sticky=tkt.W, padx = 2)

    command = button_func(k)
    button = tkt.Button(text='...', name=f'b_{k}', command=command)
    button.grid(row=i, column=2)

confirm = tkt.Button(text='Confirm', name='confirm')
confirm.grid(row=i+1, column=0, columnspan=3)

root.getvar(name='v_exel')
a = dct['exel']
a.trace('w', update)

root.mainloop()






# root_elements = root.children.values()
# for a, b in zip(
#     filter(lambda y: isinstance(y, tkt.Label), root_elements),
#     filter(lambda y: isinstance(y, tkt.Entry), root_elements)):
#     # print(a.cget('text'))
#     # print(b.get())






