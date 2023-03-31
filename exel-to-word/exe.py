from setting_class import DATA_MAIN_SETTINGS
from docxtpl import DocxTemplate
from clean_folder.sort import find_free_name
import tkinter as tk
from tkinter import messagebox

from documents import ExelComplex

CONFIRM = 'Confirm'
RESET = 'Reset'
CONFIRM_MESSAGE = 'You realy confirm it?'
DONE_MESSAGE = 'Successfully completed'
ERROR_MESSAGE = 'Something go wrong'
FIELD_CONFIGURATION = {
    'borderwidth': 2, 
    'width': 100
}

PASS_VARIABLE = 'pass_var'
PASSWORD = '1111'
PASSWORD_BUTTON = 'Confirm password'
PASSWORD_ERROR = 'Wrong password'
PASS_CONFIGURATION = {
    'borderwidth': 2, 
    'width': 20
}

class PasswordFrame(tk.Frame):
    def __init__(self, parent, controller, data):
        tk.Frame.__init__(self, parent)
        pass_label = tk.Label(self, text='Password:')
        pass_var = tk.StringVar(self, name=PASS_VARIABLE)
        pass_entry = tk.Entry(self, show='*', textvariable=pass_var, **PASS_CONFIGURATION)
        pass_button = tk.Button(self, text=PASSWORD_BUTTON, command=lambda: self.confirm_password(controller))
        pass_label.grid(row=0, column=0)
        pass_entry.grid(row=0, column=1, padx=5)
        pass_button.grid(row=1, column=0, columnspan=2)   

    def confirm_password(self, controller):
        password = self.getvar(name=PASS_VARIABLE)
        if password == PASSWORD:
            controller.show_frame(MainFrame)
        else:
            messagebox.showerror(message=PASSWORD_ERROR)
        pass   

class MainFrame(tk.Frame):
    def __init__(self, parent, controller, data):
        tk.Frame.__init__(self, parent)
        self.data = {x.name: x for x in data}
        self.key_element = data[0]

        for i, element in enumerate(self.data.values()):
            label = element.make_label(self)
            element.make_var(self)
            entry = element.make_entry(self)
            button = element.make_button(self)

            label.grid(row=i, column=0, sticky=tk.E)
            entry.grid(row=i, column=2, sticky=tk.W, padx = 2)
            if button:
                button.grid(row=i, column=1)

        self.key_element.tk_variable.trace_add('write', self.update)

        confirm_button = tk.Button(self, text=CONFIRM, name=CONFIRM.lower(), command=lambda: self.get_settings())
        confirm_button.grid(row=i+1, column=2, padx=2, pady=2, sticky=tk.W)
        confirm_button = tk.Button(self, text=RESET, name=RESET.lower(), command=self.reset)
        confirm_button.grid(row=i+1, column=1, pady=2, sticky=tk.E)
        pass

    def update(self, *args):
        file = self.key_element.get_value()
        for element in filter(lambda x: x != self.key_element, self.data.values()):
            element.get_base(file)
            element.set_defoult_value()
            element.set_options()
        pass

    def reset(self):
        for element in self.data.values():
            element.clear_vr()
        pass

    def get_settings(self):
        def is_all_confirm(self):
            for element in self.data.values():
                confirmation = element.get_confirmation()
                if not confirmation:
                    error = f'{element.name.upper()}: irrelevant value'
                    messagebox.showerror(title='ERROR', message=error)
                    return False
            return True
        
        if is_all_confirm(self):
            responce = messagebox.askokcancel(message=CONFIRM_MESSAGE)    
            if responce:
                try:
                    self.exe_program()
                except:
                    messagebox.showerror(message=ERROR_MESSAGE)
                else:
                    messagebox.showinfo(message=DONE_MESSAGE)
        pass

    def exe_program(self):
        exel_path = self.data['exel'].get_value()
        doc_path = self.data['word'].get_value()
        exel_gr_path = self.data['exel with charts'].get_value()
        text_sheet = self.data['sheet with information'].get_value()
        graph_sheet = self.data['sheet with charts'].get_value()
        result_file_name = self.data['result file name'].get_value()
        result_folder = self.data['result folder'].get_value()

        doc = DocxTemplate(doc_path)
        exel = ExelComplex(exel_path, exel_gr_path, text_sheet, graph_sheet)
        info = exel.get_info()
        doc_result_path = find_free_name(result_file_name, result_folder, doc_path.suffix)[1]
        doc.render(info)
        doc.save(doc_result_path)
        exel.change_graph_file()
        exel.copy_paste_charts(doc_result_path)
        pass

class MainWingow(tk.Tk):
    def __init__(self, data, screenName: str | None = None, baseName: str | None = None, className: str = "Tk", useTk: bool = True, sync: bool = False, use: str | None = None) -> None:
        super().__init__(screenName, baseName, className, useTk, sync, use)
        self.data = data
        self.frames = {}
        self.current_frame = None

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        for F in (PasswordFrame, MainFrame):
            frame = F(parent=container, controller=self, data=data)
            self.frames[F] = frame
            # frame.grid(row=0, column=0)

        self.show_frame(PasswordFrame)
        self.current_frame = self.frames[MainFrame]

    def show_frame(self, page_name):
        if self.current_frame:
            self.current_frame.grid_forget()
        frame: tk.Frame = self.frames[page_name]
        frame.grid(row=0, column=0)
        self.current_frame = frame
    
if __name__ == '__main__':
    root = MainWingow(DATA_MAIN_SETTINGS)
    root.resizable(width=0, height=0)
    root.mainloop()

