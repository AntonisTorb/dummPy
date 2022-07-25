#from pathlib import Path  # core python module
#import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
#---------- !NEED TO TURN GLOBALS INTO PARAMETERS OR MOVE IN FUNCTIONS! -----------#
DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal1"
THEMES = sg.theme_list()
EXCEL_COLUMN = [chr(chNum) for chNum in list(range(ord('A'),ord('Z')+1))]

# ---------- MAIN WINDOW AND LOGIC LOOP ----------#
def main_window():
    DATA_CATEGORIES = ("FIRSTNAME","LASTNAME", "AGE", "COUNTRY")
    current_theme = DEFAULT_THEME
    window_position= (None, None)
    
    rows_amount = 100

# ---------- MAIN WINDOW SETUP FOR INTIALIZATION AND RESET ----------#
    def new_window(pos, theme= DEFAULT_THEME):
        menu_def = [["File", ["New", "Open", "---", "Settings", "---", "Exit"]], 
                    ["Help", ["Documentation", "About"]]]

        if theme:
            sg.theme(theme)
        layout = [
            [sg.MenubarCustom(menu_def)],
            [sg.B("Generate", auto_size_button=True), sg.B("Set number of rows:", key= "-SETROWS-"), sg.I(size= (5,1), disabled_readonly_background_color="Light Green",default_text= rows_amount, key= "-ROWS-")],
            [sg.HorizontalSeparator()],
            [sg.T("First Name"), sg.Push(),sg.B("Configure", key= "FIRSTNAME")],
            [sg.T("Last Name"), sg.Push(),sg.B("Configure", key= "LASTNAME")],
            [sg.T("Age"), sg.Push(),sg.B("Configure", key= "AGE")],
            [sg.T("Country"), sg.Push(),sg.B("Configure", key= "COUNTRY")],
            [sg.B("Reset"), sg.B("Change theme")]
        ]
        return sg.Window("test", layout, font= DEFAULT_FONT, enable_close_attempted_event= True, location= pos,finalize=True)

    window = new_window(window_position)
    
    # ---------- ADJUST POPUP WINDOW LOCATION ----------#
    def position_correction(winpos,dx,dy):
        corrected = list(winpos)
        corrected[0] += dx
        corrected[1] += dy
        return tuple (corrected)

    def configure(evt):
        if evt in ("FIRSTNAME", "LASTNAME"):
            print(evt)
            # layout1 = 
            # cat_window = sg.Window("Name",
            # [
            # [sg.Column(),sg.Column()]
            # ])
        if evt == "AGE":
            layout1 = [
                [sg.T("Min:",size = 4, justification= "r"), sg.I(size= (10,1), key= "-AGEMIN-"),sg.T("1 < Min < Max",font= DEFAULT_FONT+ ("italic",))],
                [sg.T("Max:",size = 4, justification= "r"), sg.I(size= (10,1), key= "-AGEMAX-"),sg.T("Min < Max < 120",font= DEFAULT_FONT+ ("italic",))],
                [sg.HorizontalSeparator()],
                [sg.Push(),sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly=True),sg.B("Clear"),sg.Push()],
                [sg.Push(),sg.B("OK"), sg.B("Cancel"),sg.Push()]
            ]
            cat_location = position_correction(window_position,-10,50)
            cat_window = sg.Window("Age", layout1, font= DEFAULT_FONT, modal= True, location= cat_location)
            #----------- DATA AGE LOOP -----------#
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
        if evt == "COUNTRY":
             print(evt)
        
    def reset():
        pos_reset = position_correction(window_position,30,50)
        event, values = sg.Window("WARNING!", 
            [[sg.T("This will reset the app and all progress will be lost!")],
            [sg.Push(), sg.OK(), sg.Cancel(),sg.Push()]], font= DEFAULT_FONT, modal= True, location= pos_reset).read(close= True)
        if event == "OK":
            nonlocal window
            window.close()
            window = new_window(window_position,current_theme)
            
    def change_theme(theme,win):
        theme_position = position_correction(window_position,100,50)
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
                    window = new_window(window_position, current_theme)
                    return current_theme, window
        theme_window.close() 
        return theme, win

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
            except:
                sg.Window("ERROR!", [[sg.T("Number of rows must be an integer")],
                [sg.Push(), sg.OK(), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
        if event in DATA_CATEGORIES :
            configure(event)
        if event == "Change theme":
            current_theme, window = change_theme(current_theme, window)
        if event in ("Reset", "New"):
            reset()
    window.close()
        
if __name__ == "__main__":
    main_window()