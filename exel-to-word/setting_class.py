from pathlib import Path
from tkinter import filedialog, messagebox
from work_with_exel import is_sheet_exist, first_sheet_name, sheet_names
import tkinter as tk

DOC_EXTENSION = '.docx'
EXEL_EXTENSION = '.xlsx'
TEMPLATE_SYMBOL = 'бланк'
GRAPH_SYMBOL = 'gr'
RESULT_FOLDER_NAME = 'виконані'

BUTTON_TEXT = 'Choose'
FIELD_CONFIGURATION = {
    'borderwidth': 2, 
    'width': 100
}

class SettinsElement:
    def __init__(self, name, base=None, tk_variable=None) -> None:
        self.name = name
        self.base = base
        self.tk_variable = tk_variable
        pass

    def clear_vr(self):
        self.tk_variable.set('')
        pass

    def is_value_exist(func):
        def inner(self):
            if self.tk_variable.get():
                return func(self)
            return None
        return inner
    
    def make_label(self, root):
        label = tk.Label(root, name=f'l_{self.name}', text=self.name.capitalize())
        return label
    
    def make_var(self, root):
        self.tk_variable = tk.StringVar(root, name=f'v_{self.name}')
        return self.tk_variable
    
    def make_entry(self, root):
        entry = tk.Entry(root, name=f'e_{self.name}', textvariable=self.tk_variable, **FIELD_CONFIGURATION)
        return entry
    
    def get_value(self):
        value = self.tk_variable.get()
        return value
    
    def set_options(self, *args):
        pass
    
    def make_button(self, root):
        pass

    def get_base(self, *args):
        pass

    def set_defoult_value(self):
        pass

    @is_value_exist
    def get_confirmation(self):
        return True


class PathElement(SettinsElement):
    def __init__(self, name, word_pattern, base=None, tk_variable=None) -> None:
        super().__init__(name, base, tk_variable)
        self.pattern = f'*{word_pattern}*'
        self.base = None

    def get_value(self):
        value = Path(self.tk_variable.get())
        return value

    @SettinsElement.is_value_exist
    def get_confirmation(self):
        value = self.get_value()
        return value.exists()
    
    def get_base(self, file: Path):
        self.base = file.parent
        pass
    
    def set_defoult_value(self):
        
        try:
            self.tk_variable.set(next(self.base.glob(self.pattern)))
        except:
            self.clear_vr()
        pass
    
    def choose_value(self):
        pass
    
    def make_button(self, root):
        command = self.choose_value
        button = tk.Button(root, text=BUTTON_TEXT, name=f'b_{self.name}', command=command)
        return button
    

class DocElement(PathElement):
    def __init__(self, name, word_pattern, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, base, tk_variable)
        self.pattern = f'*{word_pattern}*{DOC_EXTENSION}'
    
    def choose_value(self):
        value = filedialog.askopenfilename(filetypes=[('Word documents', f'*{DOC_EXTENSION}')])
        self.tk_variable.set(value)
        pass
    
class ExelElement(PathElement):
    def __init__(self, name, word_pattern, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, base, tk_variable)
        self.pattern = f'*{word_pattern}*{EXEL_EXTENSION}'
    
    def choose_value(self):
        value = filedialog.askopenfilename(filetypes=[('Exel documents', f'*{EXEL_EXTENSION}')])
        self.tk_variable.set(value)
        pass

class FolderElement(PathElement):
    def __init__(self, name, word_pattern, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, base, tk_variable)

    def choose_value(self):
        value  = filedialog.askdirectory()
        self.tk_variable.set(value)
        pass

class SheetElement(SettinsElement):
    def __init__(self, name, base=None, tk_variable=None) -> None:
        super().__init__(name, base, tk_variable)
        self.options = None
    
    def is_base_exel(func):
        def inner(self):
            if self.base and self.base.suffix == EXEL_EXTENSION:
                return func(self)
            self.clear_vr()
            self.options.delete(0, tk.END)
            return None
        return inner
    
    def get_base(self, file: Path):
        self.base = file
        pass
    
    @SettinsElement.is_value_exist
    @is_base_exel
    def get_confirmation(self):
        value = self.get_value()
        return is_sheet_exist(self.base, value)
    
    def make_button(self, root):
        button = tk.Menubutton(
            root, text="Choose",borderwidth=2, relief="raised"
            )
        menu = tk.Menu(button, tearoff=False)
        button.configure(menu=menu)
        self.options = menu
        return button
    
    @is_base_exel
    def set_options(self):
        options_lst = sheet_names(self.base)
        self.options.delete(0, tk.END)
        for opt in options_lst:
            self.options.add_radiobutton(
                label=opt, variable=self.tk_variable, value=opt
                )
        pass
    
class FirstSheetElement(SheetElement):
    def __init__(self, name, base=None, tk_variable=None) -> None:
        super().__init__(name, base, tk_variable)

    @SheetElement.is_base_exel
    def set_defoult_value(self):
        value = first_sheet_name(self.base)
        self.tk_variable.set(value)
        pass

class NamedSheetElement(SheetElement):
    def __init__(self, name, pattern, base=None, tk_variable=None) -> None:
        super().__init__(name, base, tk_variable)
        self.pattern = pattern

    @SheetElement.is_base_exel
    def set_defoult_value(self):
        if is_sheet_exist(self.base, self.pattern):
            self.tk_variable.set(self.pattern)
            pass

DATA_MAIN_SETTINGS = (
    ExelElement('exel', TEMPLATE_SYMBOL),
    DocElement('word', TEMPLATE_SYMBOL),
    ExelElement('exel with charts', GRAPH_SYMBOL),
    FolderElement('result folder', RESULT_FOLDER_NAME),
    FirstSheetElement('sheet with information'),
    NamedSheetElement('sheet with charts', GRAPH_SYMBOL),
    SettinsElement('result file name'),
)

DATA_TEMPLATE_SETTINGS = (
    DocElement('template word', TEMPLATE_SYMBOL),
    FolderElement('folder with tamplates', GRAPH_SYMBOL),
)


