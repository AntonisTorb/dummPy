from pathlib import Path  # core python module
from random import choice, randint, uniform  # core python module
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui

DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal1"

# ---------- ADJUST POPUP WINDOW LOCATION ----------#
def position_correction(winpos, dx, dy):
    corrected = list(winpos)
    corrected[0] += dx
    corrected[1] += dy
    return tuple (corrected)

def read_sample_data(sheet):
    sample = pd.read_excel("SAMPLE_DATA.xlsx", sheet_name= sheet, usecols= [0,0], header=None).squeeze("columns")
    return sample.values.tolist()

#------------ SORT THE DICTIONARY BY KEY (COLUMN NAME) AND ADD BLANCK COLUMNS WHEN MISSING ----------#
def dict_sort_for_df(dict, rows):
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
            for row in range(rows + 1):
                dict2[col].append("")
    return dict2

def preview_dataframe(dict, rows, winpos):
    temp_df = pd.DataFrame.from_dict(dict_sort_for_df(dict, rows))
    headers = list(temp_df.head())
    val =  temp_df.values.tolist()
    layout = [[sg.Table(val, headings= headers, num_rows= 20, display_row_numbers= True)],
        [sg.OK(button_color= ("#292e2a", "#5ebd78"))]]
    sg.Window("Dataframe Preview", layout, modal= True, font= DEFAULT_FONT, location= winpos, grab_anywhere= True).read(close= True)

#----------- ACTION DEPENDING ON DATA TYPE SELECTED -----------#
def configure(evt, win_pos, rows, dict):
    EXCEL_COLUMN = [chr(chNum) for chNum in list(range(ord('A'), ord('Z')+1))]#----------- LIST OF EXCEL COLUMN CHARACTERS, A TO Z ----------#
    #------------ FOR DATA NAME AND EMAIL ----------# 
    if evt == "-NAME-":
        f_names = read_sample_data("FIRST NAME")
        l_names = read_sample_data("LAST NAME")
        cat_location = position_correction(win_pos, 100, 50)

        layout_name = [
            [sg.T("First Name Samples"), sg.Push(), sg.T("Last Name Samples")],
            [sg.Listbox(f_names, size= (20,5), key= "-FNAMES-"), sg.Push(), sg.Listbox(l_names, size= (20,5), key= "-LNAMES-")],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Add First Name on Column:", key = "-FIRSTNAME-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNFIRSTNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-FIRSTNAMETITLE-", default_text= "First Name")],
            [sg.Checkbox("Add Last Name on Column:", key = "-LASTNAME-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNLASTNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-LASTNAMETITLE-", default_text= "Last Name")],
            [sg.T("Else Add Full Name on Column:"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-NAMETITLE-", default_text= "Name")],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Add email address on column:", key= "-EMAIL-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNMAIL-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-EMAILTITLE-", default_text= "E-mail Address")],
            [sg.HorizontalSeparator()],
            [sg.B("Preview"), sg.B("Sample Reload"), sg.Push(), sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))],
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
            if event == "Preview":
                preview_dataframe(dict, rows, cat_location)
            if event == "Add":
                #---------- DECLARING CONDITIONS FOR  VISIBILITY ---------#
                unspecified_column = ((values["-COLUMNNAME-"] == "" and values["-FIRSTNAME-"] == False and values["-LASTNAME-"] == False) or #--------- SEPARATE NAMES ARE NOT SELECTED AND FULLNAME COLUMN NOT SELECTED --------#
                (values["-COLUMNFIRSTNAME-"] == "" and values["-FIRSTNAME-"]) or #--------- IF FIRST NAME SELECTED AND COLUMN NOT SELECTED --------#
                (values["-COLUMNLASTNAME-"] == "" and values["-LASTNAME-"]) or #--------- IF LAST NAME SELECTED AND COLUMN NOT SELECTED ---------#
                (values["-COLUMNMAIL-"] == "" and values ["-EMAIL-"])) #---------- IF EMAIL SELECTED AND COLUMN NOT SELECTED ---------#
                same_column = ((values["-COLUMNNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"] and values["-FIRSTNAME-"] == False and values["-LASTNAME-"] == False) or #---------- SEPARATE NAMES ARE NOT SELECTED, EMAIL SELECTED, AND NAME COLUMN == EMAIL COLUMN --------#
                (values["-COLUMNFIRSTNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"] and values["-FIRSTNAME-"]) or #--------- FIRST NAME AND EMAIL SELECTED AND SAME COLUMN SELECTED ---------#
                (values["-COLUMNLASTNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"] and values["-LASTNAME-"]) or #--------- LAST NAME AND EMAIL SELECTED AND SAME COLUMN SELECTED ---------#
                (values["-COLUMNFIRSTNAME-"] == values["-COLUMNLASTNAME-"] and values ["-FIRSTNAME-"] and values["-LASTNAME-"])) #--------- FIRST NAME AND LAST NAME SELECTED AND SAME COLUMN SELECTED ---------#
                    
                if (unspecified_column):
                    sg.Window("ERROR!", [[sg.T("Please specify the target Column(s)")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)
                elif (same_column):
                    sg.Window("ERROR!", [[sg.T("Please specify different Columns for selected data")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 50, 100)).read(close= True)
                else: #------------ LOGIC FOR ADDING THE DATA -----------#
                    col_names = "" #--------- COLUMN NAMES FOR SEPARATE NAME VALUES, TO DECLARE CONFIRMATION --------------#
                    mail_fnames = [] #--------- IF FIRST NAMES SELECTED, VALUES SAVED TO USE FOR EMAIL GENERATION -----------#
                    mail_lnames = [] #--------- IF LAST NAMES SELECTED, VALUES SAVED TO USE FOR EMAIL GENERATION -----------#
                    if values["-FIRSTNAME-"]: #------------ ADDING FIRST NAME ----------#
                        first_col_name = values["-COLUMNFIRSTNAME-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-FIRSTNAMETITLE-"]) < 21:
                            fnames = {first_col_name:[values["-FIRSTNAMETITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            for i in range(rows):
                                first = choice(f_names)
                                fnames[first_col_name].append(first)
                                if values ["-EMAIL-"]:
                                    mail_fnames.append(first)
                            col_names = first_col_name
                            dict.update(fnames)
                        else:
                            sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                [sg.T("between 1 and 20 characters long")],
                                [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)                   
                    if values["-LASTNAME-"]: #----------- ADDING LAST NAME ---------#
                        last_col_name = values["-COLUMNLASTNAME-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-LASTNAMETITLE-"]) < 21:
                            lnames = {last_col_name:[values["-LASTNAMETITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            for i in range(rows):
                                last = choice(l_names)
                                lnames[last_col_name].append(last)
                                if values ["-EMAIL-"]:
                                    mail_lnames.append(last)
                            if col_names == "": #--------- IF NO COLUMN FOR FIRST NAME -----------#
                                col_names = last_col_name
                            else:
                                col_names = f"{col_names}, {last_col_name}"
                            dict.update(lnames)
                        else:
                            sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                [sg.T("between 1 and 20 characters long")],
                                [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)         
                    if values["-FIRSTNAME-"] == False and values["-LASTNAME-"] == False: # --------- ADDING FULL NAME ----------#
                        full_col_name = values["-COLUMNNAME-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-NAMETITLE-"]) < 21:
                            names = {full_col_name:[values["-NAMETITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            if values ["-EMAIL-"]: # -------- IF ADDING E-MAIL AS WELL ---------#
                                domains = read_sample_data("EMAIL DOMAIN")
                                mail_col_name = values["-COLUMNMAIL-"]
                                if 0 < len(values["-EMAILTITLE-"]) < 21:
                                    emails = {mail_col_name:[values["-EMAILTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                                    for i in range(rows):
                                        first = choice(f_names)
                                        last = choice(l_names)
                                        names[full_col_name].append(f"{first} {last}")
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{first}.{last}.{delta}@{domain}")
                                    dict.update(names)
                                    dict.update(emails)
                                    sg.Window("Done", [[sg.T(f"Data added on columns {full_col_name} and {mail_col_name}")],
                                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 130, 100)).read(close= True)
                                else:
                                    sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                        [sg.T("between 1 and 20 characters long")],
                                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)
                            else: # --------- IF ADDING ONLY NAME ---------#
                                for i in range(rows):
                                    first = choice(f_names)
                                    last = choice(l_names)
                                    names[full_col_name].append(f"{first} {last}")
                                dict.update(names)
                                sg.Window("Done", [[sg.T(f"Data added on column {full_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 170, 100)).read(close= True)
                        else:
                            sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                    [sg.T("between 1 and 20 characters long")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)                    
                    else: #----------- IF SEPERATE DATA ADDED, DETERMINE E-MAIL IF SELECTED AND SHOW CONFIRMATION -----------#
                        if values ["-EMAIL-"]: # -------- IF ADDING E-MAIL AS WELL ---------#
                            domains = read_sample_data("EMAIL DOMAIN")
                            mail_col_name = values["-COLUMNMAIL-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                            if 0 < len(values["-EMAILTITLE-"]) < 21:
                                emails = {mail_col_name:[values["-EMAILTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                                if len(mail_fnames) == 0: #--------- IF FIRST NAME MISSING, ONLY USE LAST NAME FOR EMNAIL ------------#
                                    for i in range(rows):
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{mail_lnames[i]}.{delta}@{domain}")
                                elif len(mail_lnames) == 0: #--------- IF LAST NAME MISSING, ONLY USE FIRST NAME FOR EMNAIL ------------#
                                    for i in range(rows):
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{mail_fnames[i]}.{delta}@{domain}")
                                else:
                                    for i in range(rows):
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{mail_fnames[i]}.{mail_lnames[i]}.{delta}@{domain}")
                                dict.update(emails) #---------- ADD THE EMAILS TO DICTIONARY ---------#
                                sg.Window("Done", [[sg.T(f"Data added on columns {col_names} and {mail_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 130, 100)).read(close= True)
                            else:
                                sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                    [sg.T("between 1 and 20 characters long")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)
                        else:
                            sg.Window("Done", [[sg.T(f"Data added on column(s) {col_names}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 170, 100)).read(close= True)
        cat_window.close()
    #------------ FOR DATA NUMBER ----------#    
    if evt == "-NUMBER-":
        DECIMALS = [i for i in range(9)]
        layout_num = [
            [sg.T("Min:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMIN-")],
            [sg.T("Max:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMAX-")],
            [sg.T("Decimals:"), sg.Push(), sg.Combo(DECIMALS, default_value= 0, key = "-DECIMALS-", readonly=True)],
            [sg.HorizontalSeparator()],
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True)],
            [sg.T("With title:"), sg.I(size= (20, 1), key= "-NUMTITLE-", default_text= "Number")],
            [sg.HorizontalSeparator()],
            [sg.B("Clear"), sg.B("Preview"), sg.Push(),sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
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
            if event == "Preview":
                preview_dataframe(dict, rows, cat_location)
            if event == "Add":
                try:
                    if values["-COLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                    else:
                        if float(values["-NUMMAX-"]) < float(values["-NUMMIN-"]):
                            sg.Window("ERROR!", [[sg.T("Please ensure that Min < Max")],
                            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 40, 40)).read(close= True)
                        else:
                            num_col_name = values["-COLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                            if 0 < len(values["-NUMTITLE-"]) < 21:
                                numbers = {num_col_name:[values["-NUMTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                                if values["-DECIMALS-"] == 0:
                                    for i in range (rows):
                                        numbers[num_col_name].append(randint(int(values["-NUMMIN-"]), int(values["-NUMMAX-"])))
                                else:
                                    for i in range (rows):
                                        numbers[num_col_name].append(round(uniform(float(values["-NUMMIN-"]), float(values["-NUMMAX-"])), values["-DECIMALS-"]))
                                dict.update(numbers)
                                sg.Window("Done", [[sg.T(f"Data added on column {num_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 170, 100)).read(close= True)
                            else:
                                sg.Window("ERROR", [[sg.T("Please ensure the column title is")],
                                    [sg.T("between 1 and 20 characters long")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)
                except:
                    sg.Window("ERROR!", [[sg.T("Please ensure the values are numbers")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location,5,40)).read(close= True)
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
            [sg.T("Add to Column:"), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-LOCTITLE-", default_text= "Location")],
            [sg.HorizontalSeparator()],
            [sg.B("Preview"), sg.B("Sample Reload"), sg.Push(), sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
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
            if event == "Preview":
                preview_dataframe(dict, rows, cat_location)
            if event == "Add":
                if values["-COLUMN-"] == "":
                    sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 20, 40)).read(close= True)
                else:
                    loc_col_name = values["-COLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                    if 0 < len(values["-LOCTITLE-"]) < 21:
                        locations = {loc_col_name:[values["-LOCTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                        for i in range(rows):
                            s_number = randint(1,200)
                            s_name_1 = choice(street_name_1)
                            s_name_2 = choice(street_name_2)
                            citystate = choice(city_state)
                            postal_code = randint(10000, 99999)
                            locations[loc_col_name].append(f"{s_number} {s_name_1} {s_name_2}, {citystate}, {postal_code}")
                        dict.update(locations)
                        sg.Window("Done", [[sg.T(f"Data added on column {loc_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 170, 100)).read(close= True)
                    else:
                        sg.Window("ERROR", [[sg.T("Please ensure the column title is")],
                            [sg.T("between 1 and 20 characters long")],
                            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(cat_location, 120, 100)).read(close= True)
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
        [sg.B("Reset"), sg.Push(), sg.B("Preview", disabled= True, disabled_button_color= ("#f2557a", None)), sg.B("Generate", disabled= True, disabled_button_color= ("#f2557a", None), button_color= ("#292e2a","#5ebd78"))]
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
                    window.Element("Preview").update(disabled= False)
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
        if event == "Preview":
                preview_dataframe(dictionary, excel_rows, window_position)
        if event == "Generate":
            filename = values["-FILENAME-"]
            #---------- CHECK FILENAME LENGTH AND FOR ILLEGAL FILENAME CHARACTERS -----------#
            if len(filename) < 1 or len(filename) > 50:
                sg.Window("ERROR!", [[sg.T("Filename must be between 1 and 50 characters long")],
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 150, 150)).read(close= True)
            else:
                pattern = ("<", ">", ":", "/", "\"", "\\", "|", "?", "*")
                no_special = True
                for i in pattern:
                    if i in filename:
                        sg.Window("ERROR!", [[sg.T("Please do not use any of the following characters in filename:")],
                            [sg.Push(), sg.T("<>:/\|?*"), sg.Push()],
                            [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 125, 140)).read(close= True)
                        no_special = False
                        break
                if no_special:
                    df2 = dict_sort_for_df(dictionary,excel_rows)
                    pd.DataFrame.from_dict(df2).to_excel(f"{filename}.xlsx", header= False, index= False)
                    sg.Popup(f"{filename}.xlsx created", font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 150, 140), button_color= ("#292e2a","#5ebd78"))

    window.close()
        
if __name__ == "__main__":
    main_window()