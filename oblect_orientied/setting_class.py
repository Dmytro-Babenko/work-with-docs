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
            if self.tk_variable.get():
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

    def get_value(self):
        value = self.tk_variable.get()
        return value
    
    def make_button(self, root):
        pass

    def get_base(self, *args):
        pass

    def set_defoult_value(self):
        pass

    def get_confirmation(self):
        pass

class PathElement(SettinsElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)
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
            pass
        pass
    
    def choose_value(self):
        pass
    
    def make_button(self, root):
        command = self.choose_value
        button = tk.Button(root, text=BUTTON_TEXT, name=f'b_{self.name}', command=command)
        return button
    

class DocElement(PathElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, value, base, tk_variable)
        self.pattern = f'*{word_pattern}*{DOC_EXTENSION}'
    
    def choose_value(self):
        value = filedialog.askopenfilename(filetypes=[('Word documents', f'*{DOC_EXTENSION}')])
        self.tk_variable.set(value)
        pass
    
class ExelElement(PathElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, value, base, tk_variable)
        self.pattern = f'*{word_pattern}*.{EXEL_EXTENSION}'
    
    def choose_value(self):
        value = filedialog.askopenfilename(filetypes=[('Exel documents', f'*{EXEL_EXTENSION}')])
        self.tk_variable.set(value)
        pass

class FolderElement(PathElement):
    def __init__(self, name, word_pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, word_pattern, value, base, tk_variable)

    def choose_value(self):
        value  = filedialog.askdirectory()
        self.tk_variable.set(value)
        pass

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
        value = self.get_value()
        return is_sheet_exist(self.base, value)

    def set_defoult_value(self):
        return super().set_defoult_value()
    
class FirstSheetElement(SheetElement):
    def __init__(self, name, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)

    @SheetElement.is_base_exel
    def set_defoult_value(self):
        value = first_sheet_name(self.base)
        self.tk_variable.set(value)
        pass

class NamedSheetElement(SheetElement):
    def __init__(self, name, pattern, value=None, base=None, tk_variable=None) -> None:
        super().__init__(name, value, base, tk_variable)
        self.pattern = pattern

    @SheetElement.is_base_exel
    def set_defoult_value(self):
        if is_sheet_exist(self.base, self.pattern):
            self.tk_variable.set(self.pattern)
            pass

class Setting():
    def __init__(self, setting_name, data) -> None:
        self.name = setting_name
        self.data = {x.name: x for x in data}
        self.key_element = data[0]
        pass

    def make_setting_root(self):
        main_root = tk.Tk()
        main_root.title(self.name)
        main_root.resizable(width=0, height=0)
        return main_root
    
    def update(self, *args):
        file = self.key_element.get_value()
        for element in filter(lambda x: x != self.key_element, self.data.values()):
            element.get_base(file)
            element.set_defoult_value()
        pass
    
    def add_trace(self):
        self.key_element.tk_variable.trace('w', self.update)
    
    def make_fields(self, root):
        for i, element in enumerate(self.data.values()):
            label = element.make_label(root)
            element.make_var(root)
            entry = element.make_entry(root)
            button = element.make_button(root)

            label.grid(row=i, column=0, sticky=tk.E)
            entry.grid(row=i, column=1, sticky=tk.W, padx = 2)
            if button:
                button.grid(row=i, column=2)
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
main_root = main_settings.make_setting_root()
main_settings.make_fields(main_root)
main_settings.add_trace()
main_root.mainloop()

