from collections import UserDict
from pathlib import Path
from tkinter import filedialog
from work_with_exel import is_sheet_exist, first_sheet_name, GRAPH_SYMBOL
import tkinter as tk

DOC_EXTENSION = '.docx'
EXEL_EXTENSION = '.xlsx'
TEMPLATE_SYMBOL = 'бланк'
RESULT_FOLDER_NAME = 'виконані'
EXTENTION = ['Sheet with charts']
BUTTON_TEXT = 'Chose'

FIELD_CONFIGURATION = {
    'borderwidth': 2, 
    'width': 80
}

class SettinsElement:
    def __init__(self, name, value=None, base=None, tk_variable=None) -> None:
        self.name = name
        self.value = value
        self.base = base
        self.tk_variable = tk_variable
        pass

    def is_value_exist(func):
        def inner(self):
            if self.value:
                return func(self)
            return None
        return inner
    
    def make_label(self, root):
        label = tk.Label(root, name=f'l_{self.name}', text=self.name)
        return label
    
    def make_var(self, root):
        self.tk_variable = tk.StringVar(root, name=f'v_{self.name}')
        return self.tk_variable
    
    def make_entry(self, root):
        entry = tk.Entry(root, name=f'e_{self.name}', textvariable=self.tk_variable, **FIELD_CONFIGURATION)
        return entry

    def make_button(self, root):
        pass

    def get_base(self):
        pass

    def get_defoult_value(self):
        pass

    def get_confirmation(self):
        pass

class PathElement(SettinsElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)
        self.pattern = f'*{word_pattern}*'
        self.base = None

    @SettinsElement.is_value_exist
    def get_confirmation(self):
        return self.value.exists()
    
    def get_base(self, file: Path):
        self.base = file.parent
        pass
    
    def get_defoult_value(self):
        try:
            self.value = next(self.base.glob(self.pattern))
        except:
            return None
        return self.value

class DocElement(PathElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, value, base, tk_variable)
        self.pattern = f'*{word_pattern}*{DOC_EXTENSION}'
    
    def choose_value(self):
        self.value = Path(filedialog.askopenfilename(filetypes=[('Word documents', f'*{DOC_EXTENSION}')]))
        return self.value
    
class ExelElement(PathElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, value, base, tk_variable)
        self.pattern = f'*{word_pattern}*.{EXEL_EXTENSION}'
    
    def choose_value(self):
        self.value = Path(filedialog.askopenfilename(filetypes=[('Exel documents', f'*{EXEL_EXTENSION}')]))
        return self.value

class FolderElement(PathElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, value, base, tk_variable)

    def choose_value(self):
        self.value = Path(filedialog.askdirectory())
        return self.value

class SheetElement(SettinsElement):
    def __init__(self, name, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)
    
    def is_base_exel(func):
        def inner(self):
            if self.base and self.base.suffix == EXEL_EXTENSION:
                func(self)
            return None
        return inner
    
    def get_base(self, file: Path):
        self.base = file
        pass
    
    @SettinsElement.is_value_exist
    @is_base_exel
    def get_confirmation(self):
        return is_sheet_exist(self.base, self.value)

    def get_defoult_value(self):
        return super().get_defoult_value()
    
class FirstSheetElement(SheetElement):
    def __init__(self, name, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)

    @SheetElement.is_base_exel
    def get_defoult_value(self):
        self.value = first_sheet_name(self.base)
        return self.value

class NamedSheetElement(SheetElement):
    def __init__(self, name, pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)
        self.pattern = pattern

    @SheetElement.is_base_exel
    def get_defoult_value(self):
        if is_sheet_exist(self.base, self.pattern):
            self.value = first_sheet_name(self.base)
            return self.value

class Setting():
    def __init__(self, setting_name, data) -> None:
        self.name = setting_name
        self.data = {x.name: x for x in data}
        self.key_element = data[0]
        pass

DATA_MAIN_SETTINGS = (
    ExelElement('exel', TEMPLATE_SYMBOL),
    DocElement('word', TEMPLATE_SYMBOL),
    SettinsElement('result folder'),
    FirstSheetElement('sheet with information'),
    NamedSheetElement('sheet with charts', GRAPH_SYMBOL),
    SettinsElement('result file name'),
)

main_settings = Setting('Main setting', DATA_MAIN_SETTINGS)
print(main_settings.key_element.name)
# a = DocElement('asd', 'бланк')
# a.get_base(Path(r'D:\test\1.xlsx'))
# print(a.get_defoult_value())
# print(a.value)
# print(a.get_confirmation())

