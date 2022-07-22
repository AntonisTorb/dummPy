#from pathlib import Path  # core python module
#import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui

DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal9"
THEMES = sg.theme_list()
current_theme = DEFAULT_THEME
# ADD CATEGORIES
def new_layout(i,cat_name):
    return [[sg.T(cat_name), sg.InputText(key=("-CAT-", i))]]

def new_window(theme=current_theme):
        menu_def = [["File", ["New", "Open", "---", "Settings", "---", "Exit"]], 
                    ["Help", ["Documentation", "About"]]]
        data_categories = ("Name", "Age", "Country")
        if theme:
            sg.theme(theme)
        layout = [
        [sg.MenubarCustom(menu_def)],
        [sg.B("Generate", auto_size_button=True), sg.T("Choose Data Type:"), sg.Combo(data_categories, default_value= "Name", key = "-CAT-", size=(15, len(data_categories)), readonly= True), sg.B("Add Category", key= "-ADD-")],
        [sg.HorizontalSeparator()],
        [sg.Column([""], key= "-Column-")],
        [sg.B("Reset"), sg.B("Change theme")]
        ]
        return sg.Window("test", layout, font= DEFAULT_FONT, enable_close_attempted_event= True)

def main_window():
    window = new_window()

    def reset(theme):
        event, values = sg.Window("RESET", 
            [[sg.T("WARNING! If you reset you will lose all progress!"), sg.OK(), sg.Cancel()]], font= DEFAULT_FONT, modal= True
            ).read(close= True)
        if event == "OK":
            nonlocal window
            window.close()
            window = new_window(theme)

    def change_theme():
        event, values = sg.Window("Choose theme: ",
            [[sg.Combo(THEMES, key= "-THEME-", readonly= True, default_value= DEFAULT_THEME), sg.OK(), sg.Cancel()]], font= DEFAULT_FONT
            ).read(close= True)
        if event == "OK":
            theme_temp = values["-THEME-"]
            event, values = sg.Window("WARNING!",[
            [sg.T("This will reset the app and all progress will be lost!")],
            [sg.OK(), sg.Cancel()]
            ], font= DEFAULT_FONT).read(close= True)
            if event == "OK":
                nonlocal window
                global current_theme 
                current_theme= theme_temp
                window.close()
                window = new_window(current_theme)
    i = 1
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Exit":
            event, values = sg.Window("Exit",
            [[sg.T("Confirm Exit")], 
            [sg.OK(), sg.Cancel()]], font= DEFAULT_FONT
            ).read(close= True)
            if event == "OK":
                break
        if event == "-ADD-":
            if i < 15:
                window.extend_layout(window["-Column-"], new_layout(i, values["-CAT-"]))
            i += 1
        if event == "Change theme":
            change_theme()
        if event == "Reset":
            reset(current_theme)
    window.close()
        
if __name__ == "__main__":
    main_window()