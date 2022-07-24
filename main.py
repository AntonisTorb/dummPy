#from pathlib import Path  # core python module
#import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui

DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal9"
THEMES = sg.theme_list()
current_theme = DEFAULT_THEME
window_position= (None, None)
excel_column = [chr(chNum) for chNum in list(range(ord('A'),ord('Z')+1))]

# ---------- ADD CATEGORIES ----------#
def new_layout(i):
    data_categories = ("Name", "Age", "Country")
    return [[sg.Combo(data_categories, default_value= "Age", key = ("-CAT-", i), size= (15, len(data_categories)), readonly= True), sg.B("Configure", key=("-CONF-", i))]]

# ---------- MAIN WINDOW SETUP FOR INTIALIZATION AND RESET ----------#
def new_window(pos, theme= current_theme):
        menu_def = [["File", ["New", "Open", "---", "Settings", "---", "Exit"]], 
                    ["Help", ["Documentation", "About"]]]
        
        if theme:
            sg.theme(theme)
        layout = [
            [sg.MenubarCustom(menu_def)],
            [sg.B("Generate", auto_size_button=True), sg.B("Add Category", key= "-ADDCAT-")],
            [sg.HorizontalSeparator()],
            [sg.Column([""], key= "-Column-")],
            [sg.B("Reset"), sg.B("Change theme")]
        ]
        return sg.Window("test", layout, font= DEFAULT_FONT, enable_close_attempted_event= True, location= pos)

# ---------- ADJUST POPUP WINDOW LOCATION ----------#
def position_correction(dx,dy):
    corrected = list(window_position)
    corrected[0] += dx
    corrected[1] += dy
    return tuple (corrected)

# ---------- MAIN WINDOW AND LOGIC LOOP ----------#
def main_window():
    global window_position
    window = new_window(window_position)
    
    def configure(elem,val):
        if val[("-CAT-", elem)] == "Name":
            print("Name")
            # layout1 = 
            # cat_window = sg.Window("Name",
            # [
            # [sg.Column(),sg.Column()]
            # ])
        if val[("-CAT-", elem)] == "Age":
            layout1 = [
                [sg.T("Min:",size = 4, justification= "r"), sg.I(size= (10,1), key= "-AGEMIN-"),sg.T("1 < Min < 120",font= DEFAULT_FONT+ ("italic",))],
                [sg.T("Max:",size = 4, justification= "r"), sg.I(size= (10,1), key= "-AGEMAX-"),sg.T("1 < Max < 120",font= DEFAULT_FONT+ ("italic",))],
                [sg.HorizontalSeparator()],
                [sg.T("Add to Column:"), sg.Combo(excel_column, key = ("-COLUMN-", elem), readonly=True),sg.B("OK"), sg.B("Cancel")]
            ]
            cat_locatiion = position_correction(-10,50)
            cat_window = sg.Window("Age", layout1, font= DEFAULT_FONT, modal= True, location= cat_locatiion)
            while True:
                event, values = cat_window.read()
                if event == sg.WINDOW_CLOSED or event == "Cancel":
                    break
                if event == "OK":
                    try:
                        if values[("-COLUMN-", elem)] == "":
                            sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                            [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
                        else:
                            if int(values["-AGEMAX-"]) < int(values["-AGEMIN-"]):
                                sg.Window("ERROR!", [[sg.T("Please ensure that Min < Max")],
                                [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
                            elif int(values["-AGEMIN-"]) < 120 and int(values["-AGEMIN-"]) > 0 and int(values["-AGEMAX-"]) < 120 and int(values["-AGEMAX-"]) > 0 and int(values["-AGEMAX-"]) > int(values["-AGEMIN-"]):
                                sg.Popup("done", font= DEFAULT_FONT, modal= True)
                                break
                            else:
                                sg.Window("ERROR!", [[sg.T("Please ensure that values are between 1 and 120")],
                                [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
                    except:
                        sg.Window("ERROR!", [[sg.T("Please ensure the values are integers")],
                        [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
            cat_window.close()




        if val[("-CAT-", elem)] == "Country":
            print("Country")
        
    def reset():
        pos_reset = position_correction(30,50)
        event, values = sg.Window("WARNING!", 
            [[sg.T("This will reset the app and all progress will be lost!")],
            [sg.Push(), sg.OK(), sg.Cancel(),sg.Push()]], font= DEFAULT_FONT, modal= True, location= pos_reset
            ).read(close= True)
        if event == "OK":
            nonlocal window
            window.close()
            window = new_window(window_position,current_theme)
            

    def change_theme():
        pos_theme = position_correction(100,50)
        event, values = sg.Window("Choose theme: ",
            [[sg.Combo(THEMES, key= "-THEME-", readonly= True, default_value= DEFAULT_THEME), sg.OK(), sg.Cancel()]], font= DEFAULT_FONT, location= pos_theme
            ).read(close= True)
        if event == "OK":
            pos_confirm = position_correction(100,50)
            theme_temp = values["-THEME-"]
            event, values = sg.Window("WARNING!",[
                [sg.T("This will reset the app and all progress will be lost!")],
                [sg.Push(), sg.OK(), sg.Cancel(), sg.Push()]
            ], font= DEFAULT_FONT, location= pos_confirm).read(close= True)
            if event == "OK":
                nonlocal window
                global current_theme 
                current_theme= theme_temp
                window.close()
                window = new_window(window_position, current_theme)
                
             
    i = 0
    while True:
        event, values = window.read()
        window_position = window.current_location()
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Exit":
            pos_exit = position_correction(250,50)
            event, values = sg.Window("Exit",[
                [sg.T("Confirm Exit")], 
                [sg.OK(), sg.Cancel()]
            ], font= DEFAULT_FONT, location= pos_exit, modal= True).read(close= True)
            if event == "OK":
                break
        if event == "-ADDCAT-":
            i += 1
            if i < 15:
                window.extend_layout(window["-Column-"], new_layout(i))
        j = i
        while j > 0:
            if event == ("-CONF-", j):
                configure(j,values)
                break
            j -= 1
        if event == "Change theme":
            change_theme()
        if event in ("Reset", "New"):
            reset()
    window.close()
        
if __name__ == "__main__":
    main_window()