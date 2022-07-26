from pathlib import Path  # core python module
#import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
#---------- !NEED TO TURN GLOBALS INTO PARAMETERS OR MOVE IN FUNCTIONS! -----------#
DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal1"

# ---------- ADJUST POPUP WINDOW LOCATION ----------#
def position_correction(winpos,dx,dy):
    corrected = list(winpos)
    corrected[0] += dx
    corrected[1] += dy
    return tuple (corrected)

def configure(evt, win_pos):
    EXCEL_COLUMN = [chr(chNum) for chNum in list(range(ord('A'),ord('Z')+1))]
    if evt in ("-FIRSTNAME-", "-LASTNAME-"):
        print(evt)
        # layout1 = 
        # cat_window = sg.Window("Name",
        # [
        # [sg.Column(),sg.Column()]
        # ])
    if evt == "-AGE-":
        layout1 = [
            [sg.T("Min:",size = 4, justification= "r"), sg.I(size= (10,1), key= "-AGEMIN-"),sg.T("1 < Min < Max",font= DEFAULT_FONT+ ("italic",))],
            [sg.T("Max:",size = 4, justification= "r"), sg.I(size= (10,1), key= "-AGEMAX-"),sg.T("Min < Max < 120",font= DEFAULT_FONT+ ("italic",))],
            [sg.HorizontalSeparator()],
            [sg.Push(),sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly=True),sg.B("Clear"),sg.Push()],
            [sg.Push(),sg.B("OK"), sg.B("Cancel"),sg.Push()]
        ]
        cat_location = position_correction(win_pos,-10,50)
        cat_window = sg.Window("-AGE-", layout1, font= DEFAULT_FONT, modal= True, location= cat_location)
        #----------- DATA -AGE- LOOP -----------#
        while True:
            event, values = cat_window.read()
            cat_location = cat_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Cancel":
                break
            if event == "Clear":
                cat_window.Element("-AGEMIN-").Update("")
                cat_window.Element("-AGEMAX-").Update("")
            if event == "OK":
                try:
                    if values["-COLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                        [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,20,40)).read(close= True)
                    else:
                        if int(values["-AGEMAX-"]) < int(values["-AGEMIN-"]):
                            sg.Window("ERROR!", [[sg.T("Please ensure that Min < Max")],
                            [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,40,40)).read(close= True)
                        elif int(values["-AGEMIN-"]) < 120 and int(values["-AGEMIN-"]) > 0 and int(values["-AGEMAX-"]) < 120 and int(values["-AGEMAX-"]) > 0 and int(values["-AGEMAX-"]) > int(values["-AGEMIN-"]):
                            #---------- !NEED TO ADD COLUMN DATA TO DF! ------------#
                            sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,100,40))
                            #break
                        else:
                            sg.Window("ERROR!", [[sg.T("Please ensure that values are between 1 and 120")],
                            [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,-50,40)).read(close= True)
                except:
                    sg.Window("ERROR!", [[sg.T("Please ensure the values are integers")],
                    [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,5,40)).read(close= True)
        cat_window.close()
    if evt == "-COUNTRY-":
            print(evt)

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
        [sg.T("Excel File Name:"), sg.I(key= "-FILENAME-", size= (30,1), default_text= "dummPy"), sg.B("Set number of rows:", key= "-SETROWS-"), sg.I(size= (5,1), disabled_readonly_background_color="Light Green",default_text= 100, key= "-ROWS-"),],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Data Types:",font= DEFAULT_FONT+ ("bold",)), sg.Push()],
        [sg.T("First Name"), sg.Push(),sg.B("Configure", key= "-FIRSTNAME-")],
        [sg.T("Last Name"), sg.Push(),sg.B("Configure", key= "-LASTNAME-")],
        [sg.T("Age"), sg.Push(),sg.B("Configure", key= "-AGE-")],
        [sg.T("Country"), sg.Push(),sg.B("Configure", key= "-COUNTRY-")],
        [sg.HorizontalSeparator()],
        [sg.B("Reset"), sg.Push(), sg.B("Preview Dataframe"), sg.B("Generate", disabled= True)]
    ]
    return sg.Window("test", layout, font= DEFAULT_FONT, enable_close_attempted_event= True, location= pos,finalize=True)

# ---------- MAIN WINDOW AND LOGIC LOOP ----------#
def main_window():
    DATA_CATEGORIES = ("-FIRSTNAME-","-LASTNAME-", "-AGE-", "-COUNTRY-")
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
                    window.Element("-ROWS-").update(disabled= True)
                    window.Element("-SETROWS-").update(disabled= True)
                    window.Element("Generate").update(disabled= False)
            except:
                sg.Window("ERROR!", [[sg.T("Number of rows must be an integer")],
                [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
        if event in DATA_CATEGORIES :
            configure(event,window_position)
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
                    #----------- !GENERATE EXCEL FILE WITH DATAFRAME DATA HERE! -----------#
                    print(filename)
               
    window.close()
        
if __name__ == "__main__":
    main_window()