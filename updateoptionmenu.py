import tkinter as tk
import setupwip


class Dropdown():
    def __init__(self, parent, options, entry):
        self.parent = parent
        self.options = options
        self.entry = entry
        self.om_variable = tk.StringVar(self.parent)        
        self.om_variable.trace('w', self.option_select)
        
        self.om = tk.OptionMenu(self.parent, self.om_variable, *self.options)
        
    def update_option_menu(self):
        menu = self.om["menu"]
        menu.delete(0,"end")        
        for string in self.options:
            menu.add_command(label=string, command=lambda value=string: self.om_variable.set(value))
                        
    def add_option(self, to_add):
        self.options.append(to_add)        
        print(self.options)

    def del_option(self, to_del):             
        self.options.remove(to_del)        
        print(self.options)
                           
    def option_select(self, *args):
        print(self.om_variable.get())