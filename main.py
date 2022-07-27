from pathlib import Path
from random import choice, randint, uniform  # core python module
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui

DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal1"

# ---------- ADJUST POPUP WINDOW LOCATION ----------#
def position_correction(winpos,dx,dy):
    corrected = list(winpos)
    corrected[0] += dx
    corrected[1] += dy
    return tuple (corrected)

def read_sample_data(sheet):
    a = pd.read_excel("SAMPLE_DATA.xlsx", sheet_name= sheet, usecols= [0,0], header=None).squeeze("columns")
    return a.values.tolist()
    
#----------- ACTION DEPENDING ON DATA TYPE SELECTED -----------#
def configure(evt, win_pos, rows):
    EXCEL_COLUMN = [chr(chNum) for chNum in list(range(ord('A'),ord('Z')+1))]
    #------------ FOR DATA NAME AND EMAIL ----------# 
    if evt == "-NAME-":
        f_names = read_sample_data("FIRST NAME")
        l_names = read_sample_data("LAST NAME")
        cat_location = position_correction(win_pos,50,50)
        layout_name = [
            [sg.T("First Name Samples"), sg.Push(), sg.T("Last Name Samples")],
            [sg.Listbox(f_names, size= (15,5), key= "-FNAMES-"), sg.Push(), sg.Listbox(l_names, size= (15,5), key= "-LNAMES-")],
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN1-", readonly= True),sg.Push(), sg.B("Sample Reload")],
            [sg.Checkbox("Add email address on column:", key= "-EMAIL-"),sg.Combo(EXCEL_COLUMN, key = "-COLUMN2-", readonly= True)],
            [sg.Push(),sg.B("OK"), sg.B("Back"), sg.Push()]
        ]
        cat_window = sg.Window("Name", layout_name, font= DEFAULT_FONT, modal= True, location= cat_location)
        #----------- DATA NAME LOOP ------------#
        while True:
            event, values = cat_window.read()
            cat_location = cat_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Back":
                break
            if event == "Sample Reload":
                f_names = read_sample_data("FIRST NAME")
                l_names = read_sample_data("LAST NAME")
                cat_window.Element("-FNAMES-").update(f_names)
                cat_window.Element("-LNAMES-").update(l_names)
            if event == "OK":
                if values["-COLUMN1-"] == "" or (values["-COLUMN2-"] == "" and values ["-EMAIL-"]):
                    sg.Window("ERROR!", [[sg.T("Please specify the target Column(s)")],
                    [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                elif values["-COLUMN1-"] == values["-COLUMN2-"] and values ["-EMAIL-"]:
                    sg.Window("ERROR!", [[sg.T("Please specify different Columns for name and e-mail")],
                    [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                else:
                    names = []
                    if values ["-EMAIL-"]:
                        domains = read_sample_data("EMAIL DOMAIN")
                        emails = []
                        for i in range(rows):
                            first = choice(f_names)
                            last = choice(l_names)
                            names.append(f"{first} {last}")
                            domain = choice(domains)
                            delta = randint(0,99)
                            emails.append(f"{first}.{last}.{delta}@{domain}")
                        #!!!!!!!!!!!!!   NEED TO ADD TO DF ON SELECTED COLUMN  !!!!!!!!!!!!!#
                        #print (emails)
                        #print (names)
                    else:
                        for i in range(rows):
                            first = choice(f_names)
                            last = choice(l_names)
                            names.append(f"{first} {last}")
                        #!!!!!!!!!!!!!   NEED TO ADD TO DF ON SELECTED COLUMN  !!!!!!!!!!!!!#
                        #print (names)
                    
                    sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 100, 40))
        cat_window.close()
    #------------ FOR DATA NUMBER ----------#    
    if evt == "-NUMBER-":
        DECIMALS = [i for i in range(8)]
        layout_num = [
            [sg.T("Min:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMIN-")],
            [sg.T("Max:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMAX-")],
            [sg.T("Decimals:"), sg.Push(), sg.Combo(DECIMALS, default_value= 0, key = "-DECIMALS-", readonly=True)],
            [sg.HorizontalSeparator()],
            [sg.Push(),sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True),sg.B("Clear"),sg.Push()],
            [sg.Push(),sg.B("OK"), sg.B("Back"),sg.Push()]
        ]
        cat_location = position_correction(win_pos,50,50)
        cat_window = sg.Window("Number", layout_num, font= DEFAULT_FONT, modal= True, location= cat_location)
        #----------- DATA NUMBER LOOP -----------#
        while True:
            event, values = cat_window.read()
            cat_location = cat_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Back":
                break
            if event == "Clear":
                cat_window.Element("-NUMMIN-").Update("")
                cat_window.Element("-NUMMAX-").Update("")
            if event == "OK":
                try:
                    if values["-COLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                        [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                    else:
                        if float(values["-NUMMAX-"]) < float(values["-NUMMIN-"]):
                            sg.Window("ERROR!", [[sg.T("Please ensure that Min < Max")],
                            [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 40, 40)).read(close= True)
                        else:
                            numbers= []
                            for i in range (rows):
                                numbers.append(round(uniform(float(values["-NUMMIN-"]), float(values["-NUMMAX-"])), values["-DECIMALS-"]))
                            #!!!!!!!!!!!!!   NEED TO ADD TO DF ON SELECTED COLUMN  !!!!!!!!!!!!!#
                            #print(numbers)
                            sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 100, 40))                    
                except:
                    sg.Window("ERROR!", [[sg.T("Please ensure the values are integers")],
                    [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,5,40)).read(close= True)
        cat_window.close()
    #------------ FOR DATA LOCATION ----------# 
    if evt == "-LOCATION-":
        street_name_1 = read_sample_data("STREET NAME 1")
        street_name_2 = read_sample_data("STREET NAME 2")
        city_state = read_sample_data("CITY AND STATE")
        cat_location = position_correction(win_pos,50,50)
        layout_loc = [
            [sg.T("Street Name 1 Samples"), sg.Push(), sg.T("Street Name 2 Samples"), sg.T("City and State Samples")],
            [sg.Listbox(street_name_1, size= (15,5), key= "-SNAMES1-"), sg.Push(), sg.Listbox(street_name_2, size= (15,5), key= "-SNAMES2-"), sg.Push(), sg.Listbox(city_state, size= (15,5), key= "-CITYSTATE-")],
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True),sg.Push(), sg.B("Sample Reload")],
            [sg.Push(),sg.B("OK"), sg.B("Back"), sg.Push()]
        ]
        cat_window = sg.Window("Location", layout_loc, font= DEFAULT_FONT, modal= True, location= cat_location)
        #----------- DATA LOCATION LOOP ------------#
        while True:
            event, values = cat_window.read()
            cat_location = cat_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Back":
                break
            if event == "Sample Reload":
                street_name_1 = read_sample_data("STREET NAME 1")
                street_name_2 = read_sample_data("STREET NAME 2")
                city_state = read_sample_data("CITY AND STATE")
                cat_window.Element("-SNAMES1-").update(street_name_1)
                cat_window.Element("-SNAMES2-").update(street_name_2)
                cat_window.Element("-CITYSTATE-").update(city_state)
            if event == "OK":
                if values["-COLUMN-"] == "":
                    sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                    [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                else:
                    locations = []
                    for i in range(rows):
                        s_number = randint(1,200)
                        s_name_1 = choice(street_name_1)
                        s_name_2 = choice(street_name_2)
                        citystate = choice(city_state)
                        postal_code = randint(10000,99999)
                        locations.append(f"{s_number} {s_name_1} {s_name_2}, {citystate}, {postal_code}")
                    #!!!!!!!!!!!!!   NEED TO ADD TO DF ON SELECTED COLUMN  !!!!!!!!!!!!!#
                    #print (locations)
                    sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 100, 40))
        cat_window.close()

def reset(win, win_pos,theme):
    pos_reset = position_correction(win_pos,30,50)
    while True:  
        reset_window= sg.Window("WARNING!", 
            [[sg.T("This will reset the app and all progress will be lost!")],
            [sg.Push(), sg.OK(), sg.Cancel(),sg.Push()]], font= DEFAULT_FONT, modal= True, location= pos_reset)
        event, values = reset_window.read(close= True)
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "OK":
            win.close()
            window = new_main_window(win_pos, theme)
            return window
    reset_window.close()
    return win

def change_theme(theme,win,win_pos):
    THEMES = sg.theme_list()
    theme_position = position_correction(win_pos,100,50)
    theme_window = sg.Window("Choose theme: ",
    [[sg.Combo(THEMES, key= "-THEME-", readonly= True, default_value= theme), sg.OK(), sg.Cancel()]], font= DEFAULT_FONT, location= theme_position)
    while True:
        event, values = theme_window.read()
        theme_position = theme_window.current_location()
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "OK":
            confirm_position = position_correction(theme_position,100,50)
            theme_temp = values["-THEME-"]
            event, values = sg.Window("WARNING!",[
                [sg.T("This will reset the app and all progress will be lost!")],
                [sg.Push(), sg.OK(), sg.Cancel(), sg.Push()]
            ], font= DEFAULT_FONT, location= confirm_position).read(close= True)
            if event == "OK":
                current_theme = theme_temp
                theme_window.close()
                win.close()
                window = new_main_window(win_pos, current_theme)
                return current_theme, window
    theme_window.close()
    return theme, win

# ---------- MAIN WINDOW SETUP FOR INTIALIZATION AND RESET ----------#
def new_main_window(pos, theme= DEFAULT_THEME):
    menu_def = [["File", ["New", "---", "Load", "Save", "---", "Theme", "---", "Exit"]], 
                ["Help", ["Get Started", "About"]]]
    if theme:
        sg.theme(theme)
    layout = [
        [sg.MenubarCustom(menu_def)],
        [sg.B("Output directory:", key= "-BROWSEOUT-"), sg.I(key= "-OUTPUT-",size = (50,1), default_text= Path.cwd(), disabled= True)],
        [sg.HorizontalSeparator()],
        [sg.T("Excel File Name:"), sg.I(key= "-FILENAME-", size= (30,1), default_text= "dummPy"), sg.B("Set number of rows:", key= "-SETROWS-"), sg.I(size= (5,1), disabled_readonly_background_color="Light Green",default_text= 10, key= "-ROWS-"),],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Data Types:",font= DEFAULT_FONT+ ("bold",)), sg.Push()],
        [sg.T("Name and e-mail"), sg.Push(),sg.B("Configure", key= "-NAME-", disabled= True, disabled_button_color= ("#f2557a",None))],
        [sg.T("Number"), sg.Push(),sg.B("Configure", key= "-NUMBER-", disabled= True, disabled_button_color= ("#f2557a",None))],
        [sg.T("Location"), sg.Push(),sg.B("Configure", key= "-LOCATION-", disabled= True, disabled_button_color= ("#f2557a",None))],
        [sg.HorizontalSeparator()],
        [sg.B("Reset"), sg.Push(), sg.B("Preview Dataframe"), sg.B("Generate", disabled= True, disabled_button_color= ("#f2557a",None))]
    ]
    return sg.Window("test", layout, font= DEFAULT_FONT, enable_close_attempted_event= True, location= pos)

# ---------- MAIN WINDOW AND LOGIC LOOP ----------#
def main_window():
    DATA_CATEGORIES = ("-NAME-", "-NUMBER-", "-LOCATION-")
    current_theme = DEFAULT_THEME
    window_position= (None, None)
    window = new_main_window(window_position)
    
    #--------- MAIN LOOP ---------#              
    while True:
        event, values = window.read()
        window_position = window.current_location()
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Exit":
            pos_exit = position_correction(window_position,250,50)
            event, values = sg.Window("Exit",[
                [sg.T("Confirm Exit")], 
                [sg.OK(), sg.Cancel()]
            ], font= DEFAULT_FONT, location= pos_exit, modal= True).read(close= True)
            if event == "OK":
                break
        if event == "-SETROWS-":
            try:
                rows = int(values["-ROWS-"])
                if rows < 1 or rows > 9999:
                    sg.Window("ERROR!", [[sg.T("Please enter an integer between 1 and 9999")],
                    [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
                else:
                    excel_rows = int(values["-ROWS-"])
                    window.Element("-ROWS-").update(disabled= True)
                    window.Element("-SETROWS-").update(disabled= True)
                    window.Element("-NAME-").update(disabled= False)
                    window.Element("-NUMBER-").update(disabled= False)
                    window.Element("-LOCATION-").update(disabled= False)
                    window.Element("Generate").update(disabled= False)
            except:
                sg.Window("ERROR!", [[sg.T("Number of rows must be an integer")],
                [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
        if event in DATA_CATEGORIES :
            configure(event, window_position, excel_rows)
        if event == "-BROWSEOUT-":
            path_temp = values["-OUTPUT-"]
            folder = sg.popup_get_folder('', no_window=True)
            if folder:
                window["-OUTPUT-"].update(Path(folder))
            else:
                window["-OUTPUT-"].update(path_temp)
        if event == "Theme":
            current_theme, window = change_theme(current_theme, window, window_position)
        if event in ("Reset", "New"):
            try:
                window = reset(window, window_position, current_theme)
            except:
                None
        if event == "Generate":
            filename = values["-FILENAME-"]
            #---------- CHECK FILENAME LENGTH AND FOR ILLEGAL FILENAME CHARACTERS -----------#
            if len(filename) < 1 or len(filename) > 30:
                print("1<filename characters<30")
            else:
                pattern = ("<", ">", ":", "/", "\"", "\\", "|", "?", "*")
                no_special = True
                for i in pattern:
                    if i in filename:
                        print ("no special chars plz")
                        no_special = False
                        break
                if no_special:
                    #!!!!!!!!!!!!!   GENERATE EXCEL FILE WITH DATAFRAME DATA HERE   !!!!!!!!!!!!!!!!!#
                    print(filename)
               
    window.close()
        
if __name__ == "__main__":
    main_window()