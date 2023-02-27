import tkinter as tkt
from tkinter import messagebox, ttk

# root = tkt.Tk()
# root.title('Calculator')

"""
3 виджета
1: папка
2: Ворд, Ексель, финальная папка, финальное имя
3: лист ексель с инфой лист с графиками

Подтверждение папки вызывает возможность заполнять 2 виджет и тд
После всех подтверждений появляется кнопка погнали
"""

MAIN_SETTINGS = {
    'base_folder': None,
    'Exel': None,
    'Word': None,
    'Sheet with information': None,
    'Sheet with charts': None,
    'Result file name': None,
    'Result folder': None
}

TEMPLATE_SETTINGS = {
    'base_folder': None,
    'template Word': None,
    'folder with tamplates': None,
}

BASE = {
    'base_folder': None,
    'Exel': 'base_folder',
    'Word': 'base_folder',
    'Sheet with information': 'Exel',
    'Sheet with charts': 'Exel',
    'Result folder': 'base_folder',
    'Result file name': 'base_folder',
    'template Word':'base_folder',
    'folder with tamplates': 'base_folder'
}

FIRST = ['base_folder']
SECOND = [
    'Exel',
    'Word',
    'Result file name',
    'Result folder'
]

LAST = [
    'Sheet with information',
    'Sheet with charts',
]



def get_settings(settings: dict):

    def choose_base(element):
        base = settings.get(BASE.get(element))
        return base


def make_main_root(settings: dict):

    FIRST = ['base_folder']
    SECOND = [
    'Exel',
    'Word',
    'Result file name',
    'Result folder'
    ]

    LAST = [
    'Sheet with information',
    'Sheet with charts',
    ]

    frames = [FIRST, SECOND, LAST]
    def  generator(lst):
        for i in lst:
            yield i
    l = 0

    gen = generator(frames)

    def make_frame(lst):
        frame = tkt.Frame(root)
        i = 0
        for k in lst:
            label = tkt.Label(frame, name=k.lower(), text=k)
            label.grid(row=i, column=0, sticky=tkt.W)

            e = tkt.Entry(frame, borderwidth=2, width=40)
            if settings[k]:
                e.insert(0, settings[k])
            e.grid(row=i, column=1, sticky=tkt.W, padx = 2)
            i += 1
        button = tkt.Button(frame, text='next', command=lambda: next_frame(frame))
        button.grid(row=i, column=0, sticky=tkt.N, columnspan=2)
        return frame

    def next_frame(frame):
        def save_settings():
            for a in filter(lambda x: isinstance(x, tkt.Label), frame.children):
                print(a)
        save_settings()
        frame.grid_forget()
        fr = next(gen)
        frame = make_frame(fr)
        frame.grid(row=0, column=0)
    



    root = tkt.Tk()
    frame = tkt.Frame(root)
    button = tkt.Button(frame, text='next', command=lambda: next_frame(frame))
    button.grid(row=0, column=0, sticky=tkt.N, columnspan=2)
    frame.grid(row=0, column=0)

    return root



SETTINGS = {'Word': 'blank.docx', 'Exel': 'blank.xlsx'}

root = make_main_root(MAIN_SETTINGS)

root.mainloop()


def button_click(number):
    def inner():
        current = e.get()
        e.delete(0, tkt.END)
        e.insert(0, current + str(number))
    return inner

def b_clear():
    e.delete(0, tkt.END)
    pass

result = 0

def adding():
    current_number = int(e.get())
    global result
    result += current_number
    e.delete(0, tkt.END)

def eq():
    current_number = int(e.get())
    result += current_number
    e.delete(0, tkt.END)
    e.insert(0, str(result))
    result = 0

# for i in range(10):
#     func = button_click(i)
#     button = tkt.Button(root, text=str(i), padx=40, pady=20, command=func)
#     if i:
#         button.grid(row=(i-1)//3+1, column=(i-1)%3)
#     else:
#         button.grid(row=4, column = 0)

# button_add = tkt.Button(root, text='+', padx=40, pady=20, command=adding)
# button_equal = tkt.Button(root, text='=', padx=40, pady=20, command=eq)
# button_clear = tkt.Button(root, text='Clear', padx=120, pady=20, command=b_clear)

# button_add.grid(row=4, column=1)
# button_clear.grid(row=5, column=0, columnspan=3)
# button_equal.grid(row=4, column=2)



    










