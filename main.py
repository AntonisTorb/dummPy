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
    sample = pd.read_excel("SAMPLE_DATA.xlsx", sheet_name= sheet, usecols= [0,0], header=None).squeeze("columns")
    return sample.values.tolist()

#------------ SORT THE DICTIONARY BY KEY (COLUMN NAME) AND ADD BLANCK COLUMNS WHEN MISSING ----------#
def dict_to_sorted_df(dict, rows):
    dict2 = {}
    max_col =0
    #-------- DETERMINE THE LAST COLUMN OF DICT ---------#
    for i in dict:
        if ord(i) > max_col:
            max_col = ord(i)
    #----------- CREATE A LIST WITH COLUMN NAMES UP TO THE LAST -----------#
    excel = [chr(ch_num) for ch_num in list(range(ord('A'), max_col + 1))]
    #----------- NEW DICT (DICT2) WITH VALUES FROM PREVIOUS, SORTED, AND WITH BLANCK COLUMNS WHERE THERE IS NO DATA ---------#
    for col in excel:
        try:
            dict2[col] = dict[col]
        except: #---------- IF THE COLUMN DOES NOT EXIST IN DICT ----------#
            dict2[col] = []
            for row in range(rows):
                dict2[col].append("")
    return dict2

def preview_dataframe(dict, rows, winpos):
    temp_df = pd.DataFrame.from_dict(dict_to_sorted_df(dict, rows))
    headers = list(temp_df.head())
    val =  temp_df.values.tolist()
    layout = [[sg.Table(val,headings= headers,num_rows= 20, display_row_numbers= True)],
    [sg.OK()]]
    event, values = sg.Window("Dataframe Preview",layout, modal= True, font=DEFAULT_FONT, location=winpos).read(close= True)

#----------- ACTION DEPENDING ON DATA TYPE SELECTED -----------#
def configure(evt, win_pos, rows, dict):
    EXCEL_COLUMN = [chr(chNum) for chNum in list(range(ord('A'), ord('Z')+1))]
    #------------ FOR DATA NAME AND EMAIL ----------# 
    if evt == "-NAME-":
        f_names = read_sample_data("FIRST NAME")
        l_names = read_sample_data("LAST NAME")
        cat_location = position_correction(win_pos,50,50)
        layout_name = [
            [sg.T("First Name Samples"), sg.Push(), sg.T("Last Name Samples")],
            [sg.Listbox(f_names, size= (15,5), key= "-FNAMES-"), sg.Push(), sg.Listbox(l_names, size= (15,5), key= "-LNAMES-")],
            [sg.HorizontalSeparator()],
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMNNAME-", readonly= True),sg.Push(), sg.B("Sample Reload")],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Add email address on column:", key= "-EMAIL-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNMAIL-", readonly= True)],
            [sg.HorizontalSeparator()],
            [sg.Push(), sg.B("Preview Dataframe"), sg.Push()],
            [sg.Push(),sg.B("Add", button_color= ("#292e2a","#5ebd78")), sg.B("Back", button_color= ("#ffffff","#bf365f")), sg.Push()]
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
            if event == "Preview Dataframe":
                preview_dataframe(dict, rows, cat_location)
            if event == "Add":
                if values["-COLUMNNAME-"] == "" or (values["-COLUMNMAIL-"] == "" and values ["-EMAIL-"]):
                    sg.Window("ERROR!", [[sg.T("Please specify the target Column(s)")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                elif values["-COLUMNNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"]:
                    sg.Window("ERROR!", [[sg.T("Please specify different Columns for name and e-mail")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                else:
                    colname1 = values["-COLUMNNAME-"]
                    names = {colname1:[]}
                    if values ["-EMAIL-"]:
                        domains = read_sample_data("EMAIL DOMAIN")
                        colname2 = values["-COLUMNMAIL-"]
                        emails = {colname2:[]}
                        for i in range(rows):
                            first = choice(f_names)
                            last = choice(l_names)
                            names[colname1].append(f"{first} {last}")
                            domain = choice(domains)
                            delta = randint(0,99)
                            emails[colname2].append(f"{first}.{last}.{delta}@{domain}")
                        dict.update(names)
                        dict.update(emails)
                    else:
                        for i in range(rows):
                            first = choice(f_names)
                            last = choice(l_names)
                            names[colname1].append(f"{first} {last}")
                        dict.update(names)
                    
                    sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 100, 40), button_color= ("#292e2a","#5ebd78"))
        cat_window.close()
    #------------ FOR DATA NUMBER ----------#    
    if evt == "-NUMBER-":
        DECIMALS = [i for i in range(9)]
        layout_num = [
            [sg.T("Min:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMIN-")],
            [sg.T("Max:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMAX-")],
            [sg.T("Decimals:"), sg.Push(), sg.Combo(DECIMALS, default_value= 0, key = "-DECIMALS-", readonly=True)],
            [sg.HorizontalSeparator()],
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True),sg.Push(),sg.B("Clear")],
            [sg.HorizontalSeparator()],
            [sg.Push(),sg.B("Preview Dataframe"), sg.Push()],
            [sg.HorizontalSeparator()],
            [sg.Push(),sg.B("Add", button_color= ("#292e2a","#5ebd78")), sg.B("Back", button_color= ("#ffffff","#bf365f")),sg.Push()]
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
            if event == "Preview Dataframe":
                preview_dataframe(dict, rows, cat_location)
            if event == "Add":
                try:
                    if values["-COLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                        [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                    else:
                        if float(values["-NUMMAX-"]) < float(values["-NUMMIN-"]):
                            sg.Window("ERROR!", [[sg.T("Please ensure that Min < Max")],
                            [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 40, 40)).read(close= True)
                        else:
                            colname = values["-COLUMN-"]
                            numbers = {colname:[]}
                            if values["-DECIMALS-"] == 0:
                                for i in range (rows):
                                    numbers[colname].append(randint(int(values["-NUMMIN-"]), int(values["-NUMMAX-"])))
                            else:
                                for i in range (rows):
                                    numbers[colname].append(round(uniform(float(values["-NUMMIN-"]), float(values["-NUMMAX-"])), values["-DECIMALS-"]))
                            dict.update(numbers)
                            sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 100, 40), button_color= ("#292e2a","#5ebd78"))                    
                except:
                    sg.Window("ERROR!", [[sg.T("Please ensure the values are numbers")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,5,40)).read(close= True)
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
            [sg.HorizontalSeparator()],
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True),sg.Push(), sg.B("Preview Dataframe"), sg.Push(), sg.B("Sample Reload")],
            [sg.HorizontalSeparator()],
            [sg.Push(),sg.B("Add", button_color= ("#292e2a","#5ebd78")), sg.B("Back", button_color= ("#ffffff","#bf365f")), sg.Push()]
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
            if event == "Preview Dataframe":
                preview_dataframe(dict, rows, cat_location)
            if event == "Add":
                if values["-COLUMN-"] == "":
                    sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                else:
                    colname = values["-COLUMN-"]
                    locations = {colname:[]}
                    for i in range(rows):
                        s_number = randint(1,200)
                        s_name_1 = choice(street_name_1)
                        s_name_2 = choice(street_name_2)
                        citystate = choice(city_state)
                        postal_code = randint(10000,99999)
                        locations[colname].append(f"{s_number} {s_name_1} {s_name_2}, {citystate}, {postal_code}")
                    dict.update(locations)
                    sg.Popup("done", font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 100, 40), button_color= ("#292e2a","#5ebd78"))
        cat_window.close()

def reset(win, win_pos,theme, dict):
    pos_reset = position_correction(win_pos,30,50)
    while True:  
        reset_window= sg.Window("WARNING!", 
            [[sg.T("This will reset the app and all progress will be lost!")],
            [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")),sg.Push()]], font= DEFAULT_FONT, modal= True, location= pos_reset)
        event, values = reset_window.read(close= True)
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "OK":
            dict.clear()
            win.close()
            window = new_main_window(win_pos, theme)
            return window
    reset_window.close()
    return win

def change_theme(theme, win, win_pos, dict):
    THEMES = sg.theme_list()
    theme_position = position_correction(win_pos,100,50)
    theme_window = sg.Window("Choose theme: ",
    [[sg.Combo(THEMES, key= "-THEME-", readonly= True, default_value= theme), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Cancel()]], font= DEFAULT_FONT, location= theme_position)
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
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")), sg.Push()]
            ], font= DEFAULT_FONT, location= confirm_position).read(close= True)
            if event == "OK":
                dict.clear()
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
        [sg.B("Output directory:", key= "-BROWSEOUT-"), sg.I(key= "-OUTPUT-", size= (50,1), default_text= Path.cwd(), disabled= True)],
        [sg.HorizontalSeparator()],
        [sg.T("Excel File Name:"), sg.I(key= "-FILENAME-", size= (30,1), default_text= "dummPy"), sg.B("Set number of rows:", key= "-SETROWS-"), sg.I(size= (5,1), disabled_readonly_background_color="#5ebd78",default_text= 10, key= "-ROWS-"),],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Data Types:",font= DEFAULT_FONT+ ("bold",)), sg.Push()],
        [sg.T("Name and e-mail"), sg.Push(),sg.B("Configure", key= "-NAME-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Number"), sg.Push(),sg.B("Configure", key= "-NUMBER-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Location"), sg.Push(),sg.B("Configure", key= "-LOCATION-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.HorizontalSeparator()],
        [sg.B("Reset"), sg.Push(), sg.B("Preview Dataframe", disabled= True, disabled_button_color= ("#f2557a", None)), sg.B("Generate", disabled= True, disabled_button_color= ("#f2557a", None), button_color= ("#292e2a","#5ebd78"))]
    ]
    return sg.Window("test", layout, font= DEFAULT_FONT, enable_close_attempted_event= True, location= pos)

# ---------- MAIN WINDOW AND LOGIC LOOP ----------#
def main_window():
    DATA_CATEGORIES = ("-NAME-", "-NUMBER-", "-LOCATION-")
    current_theme = DEFAULT_THEME
    window_position= (None, None)
    window = new_main_window(window_position)
    dictionary= {}
    
    #--------- MAIN LOOP ---------#              
    while True:
        event, values = window.read()
        window_position = window.current_location()
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Exit":
            pos_exit = position_correction(window_position,250,50)
            event, values = sg.Window("Exit",[
                [sg.T("Confirm Exit")], 
                [sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f"))]
            ], font= DEFAULT_FONT, location= pos_exit, modal= True).read(close= True)
            if event == "OK":
                break
        if event == "-SETROWS-":
            try:
                rows = int(values["-ROWS-"])
                if rows < 1 or rows > 9999:
                    sg.Window("ERROR!", [[sg.T("Please enter an integer between 1 and 9999")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
                else:
                    excel_rows = int(values["-ROWS-"])
                    window.Element("-ROWS-").update(disabled= True)
                    window.Element("-SETROWS-").update(disabled= True)
                    window.Element("-NAME-").update(disabled= False)
                    window.Element("-NUMBER-").update(disabled= False)
                    window.Element("-LOCATION-").update(disabled= False)
                    window.Element("Preview Dataframe").update(disabled= False)
                    window.Element("Generate").update(disabled= False)
            except:
                sg.Window("ERROR!", [[sg.T("Number of rows must be an integer")],
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True).read(close= True)
        if event in DATA_CATEGORIES :
            configure(event, window_position, excel_rows, dictionary)
            #print (dictionary)
        if event == "-BROWSEOUT-":
            path_temp = values["-OUTPUT-"]
            folder = sg.popup_get_folder('', no_window=True)
            if folder:
                window["-OUTPUT-"].update(Path(folder))
            else:
                window["-OUTPUT-"].update(path_temp)
        if event == "Theme":
            current_theme, window = change_theme(current_theme, window, window_position, dictionary)
        if event in ("Reset", "New"):
            try:
                window = reset(window, window_position, current_theme, dictionary)
            except:
                None
        if event == "Preview Dataframe":
                preview_dataframe(dictionary, excel_rows, window_position)
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
                    # print(filename)
                    # df2 = pd.DataFrame.from_dict(dictionary)
                    # print (df2)
                    # df2.to_excel(f"{filename}.xlsx", header= False, index= False)


                    #dict_to_sorted_df(dictionary,excel_rows)
                    df2 = dict_to_sorted_df(dictionary,excel_rows)
                    pd.DataFrame.from_dict(df2).to_excel(f"{filename}.xlsx", header= False, index= False)

    window.close()
        
if __name__ == "__main__":
    main_window()