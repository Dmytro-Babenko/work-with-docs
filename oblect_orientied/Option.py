import tkinter as tk

class OptionMenu(tk.OptionMenu):
    def __init__(self, *args, **kw):
        self._command = kw.get("command")
        tk.OptionMenu.__init__(self, *args, **kw)
    def addOption(self, label):
        self["menu"].add_command(
            label=label,command=tk._setit(tk.StringVar(), label, self._command)
            )
        
    def add_options(self, *nums):
        for num in nums:
            self["menu"].add_command(
                label=num,command=tk._setit(tk.StringVar(), num, self._command)
            )

    def print_menu(self):
        print(self['menu'].cget('value'))
        # for o in self['menu']:
        #     print(o)

    def my_remove(self):
            # options.set('') # remove default selection only, not the full list
            self['menu'].delete(0,'end') # remove full list 


# root = tk.Tk()

# # var = tk.OptionMenu()

# var = tk.StringVar(root)
# option = OptionMenu(root, var, '0', )
# option.pack()
# option.addOption('2')
# # print(option.value)
# root.mainloop()

if  __name__ == "__main__":
    root = tk.Tk()
    var = tk.StringVar()
    label = tk.Label(root, textvariable=var)
    menubutton = tk.Menubutton(root, text="Choose",
                            borderwidth=2, relief="raised")
    menu = tk.Menu(menubutton, tearoff=False)
    menubutton.configure(menu=menu)
    
    menu.add_radiobutton(label="One", variable=var, value="One")
    menu.add_radiobutton(label="Two", variable=var, value="Two")
    menu.add_radiobutton(label="Three", variable=var, value="Three")
    # menu.delete(0, tk.END) 
    # menu.add_radiobutton(label="Three", variable=var, value="Three")

    label.pack(side="bottom", fill="x")
    menubutton.pack(side="top")


    root.mainloop()