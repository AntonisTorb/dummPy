import ast  # core python module
import webbrowser  # core python module
from pathlib import Path  # core python module
from random import choice, randint, uniform  # core python module
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui

DEFAULT_FONT = ("Arial", 14)
DEFAULT_THEME = "DarkTeal1"
PATTERN1 = ( "<" , ">" , ":" , "/" , "\"" , "\\" , "|" , "?" , "*") #--------- ILLEGAL FILE NAME CHARACTERS ----------#
PATTERN2= ( ":" , "/" , "\\" , "?" , "*" , "[" , "]" ) #----------- ILLEGAL SHEET NAME CHARACTERS -----------#
EXCEL_COLUMN = [chr(chNum) for chNum in list(range(ord('A'), ord('Z')+1))]#----------- LIST OF EXCEL COLUMN CHARACTERS, A TO Z ----------#

# ---------- ADJUST POPUP WINDOW LOCATION ----------#
def position_correction(winpos, dx, dy):
    corrected = list(winpos)
    corrected[0] += dx
    corrected[1] += dy
    return tuple(corrected)

def read_sample_data(sheet):
    try: 
        sample = pd.read_excel("SAMPLE_DATA.xlsx", sheet_name= sheet, usecols= [0, 0], header= None).squeeze("columns")
        return sample.values.tolist()
    except:
        sg.Window("ERROR!", [[sg.T(f"{sheet} sample file is missing or corrupted!")],
            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, icon= "icon.ico").read(close= True)
        return []
    
#------------ SORT THE DICTIONARY BY KEY (COLUMN NAME) AND ADD BLANCK COLUMNS WHEN MISSING ----------#
def dict_sort_for_df(dict, rows):
    new_dict = {}
    max_col = 0
    #-------- DETERMINE THE LAST COLUMN OF DICT ---------#
    for key in dict:
        if ord(key) > max_col:
            max_col = ord(key)
    #----------- CREATE A LIST WITH COLUMN NAMES UP TO THE LAST -----------#
    excel = [chr(ch_num) for ch_num in list(range(ord('A'), max_col + 1))]
    #----------- NEW DICT WITH VALUES FROM PREVIOUS, SORTED, AND WITH BLANCK COLUMNS WHERE THERE IS NO DATA ---------#
    for col in excel:
        try:
            new_dict[col] = dict[col]
        except: #---------- IF THE COLUMN DOES NOT EXIST IN DICT ----------#
            new_dict[col] = []
            for row in range(rows + 1):
                new_dict[col].append("")
    return new_dict

def preview_dataframe(dict, rows, winpos):
    temp_df = pd.DataFrame.from_dict(dict_sort_for_df(dict, rows))
    headers = list(temp_df.head())
    val =  temp_df.values.tolist()
    preview_layout = [[sg.Table(val, headings= headers, num_rows= 20, display_row_numbers= True)],
        [sg.OK(button_color= ("#292e2a", "#5ebd78"))]]
    sg.Window("Dataframe Preview", preview_layout, modal= True, font= DEFAULT_FONT, location= winpos, grab_anywhere= True, icon= "icon.ico").read(close= True)

def save_dict(dict, winpos):
    save_dir = Path.cwd()
    save_pos = position_correction(winpos, -60, 80)
    save_layout = [
        [sg.T("Save directory:"), sg.Push(), sg.I(key = "-SAVEDIR-", disabled= True, default_text = save_dir), sg.B("Browse")],
        [sg.T("Filename:"), sg.I(key = "-SAVEFILE-", default_text= "Save1"), sg.B("Save", button_color= ("#292e2a", "#5ebd78")),  sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
    ]
    save_window = sg.Window("Save", save_layout, modal= True, font= DEFAULT_FONT, location= save_pos, icon= "icon.ico")
    while True:
        cancelled = False
        event, values = save_window.read()
        save_win_location = save_window.current_location()
        if event == sg.WINDOW_CLOSED or event == "Back":
            cancelled = True
            break
        if event == "Browse":
            save_dir = sg.popup_get_folder('', no_window=True)
            if save_dir:
                save_window["-SAVEDIR-"].update(Path(save_dir))
        if event == "Save":
            no_illegal = True
            filename = values["-SAVEFILE-"]
            for illegal in PATTERN1:
                if illegal in filename:
                    sg.Window("ERROR!", [[sg.T("Please do not use any of the following characters in filename:")],
                        [sg.Push(), sg.T("<>:/\|?*"), sg.Push()],
                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(save_win_location, 125, 70), icon= "icon.ico").read(close= True)
                    no_illegal = False
                    break
            if no_illegal:
                filepath = Path(values["-SAVEDIR-"]) / f"{filename}.txt"
                savefile = open(filepath,"w")
                savefile.write( str(dict))
                savefile.close()
                sg.Window("Saved", [[sg.T(f"Dictionary saved as {filename}.txt")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(save_win_location, 180, 80), icon= "icon.ico").read(close= True)
                break
    save_window.close()
    return cancelled

def load_dict(winpos):
    load_pos = position_correction(winpos, -40, 80)
    load_layout = [
        [sg.T("Load File:"), sg.I(key = "-LOADFILE-", disabled= True), sg.B("Browse")],
        [sg.Push(), sg.B("Load", button_color= ("#292e2a", "#5ebd78"), disabled= True, disabled_button_color= ("#f2557a", None)),  sg.B("Back", button_color= ("#ffffff", "#bf365f")), sg.Push()]
    ]
    load_window = sg.Window("Save", load_layout, modal= True, font= DEFAULT_FONT, location= load_pos, icon= "icon.ico")
    while True:
        event, values = load_window.read()
        load_win_location = load_window.current_location()
        if event == sg.WINDOW_CLOSED or event == "Back":
            break
        if event == "Browse":
            file_to_load = sg.popup_get_file('',file_types=(("text Files", "*.txt*"),), no_window=True)
            if file_to_load:
                load_window["-LOADFILE-"].update(Path(file_to_load))
                load_window.Element("Load").update(disabled= False)
        if event == "Load":
            loadfile = open(file_to_load,"r")
            temp_dict = loadfile.readline()
            loadfile.close()
            try:
                dict = ast.literal_eval(temp_dict) #----------- STRING TO DICTIONARY ----------#
                rows = len(next(iter(dict.values()))) - 1 #-------- MINUS 1 FOR TITLE ----------#
                sg.Window("Loaded", [[sg.T("Dictionary Loaded")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(load_win_location, 250, 30), icon= "icon.ico").read(close= True)
                break
            except:
                sg.Window("ERROR", [[sg.T("Save file is invalid or corrupted")],
                     [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(load_win_location, 250, 30), icon= "icon.ico").read(close= True)
    try:
        load_window.close()
        return rows, ast.literal_eval(temp_dict)
    except: #----------- IF WE CANCEL WITHOUT LOADING -----------#
        load_window.close()

#----------- ACTION DEPENDING ON DATA TYPE SELECTED -----------#
def configure(evt, win_pos, rows, dict):
    #------------ FOR DATA NAME AND EMAIL ----------# 
    if evt == "-NAME-":
        f_names = read_sample_data("FIRST NAME")
        l_names = read_sample_data("LAST NAME")
        domains = read_sample_data("EMAIL DOMAIN")
        pos_name = position_correction(win_pos, -25, 80)
        layout_name = [
            [sg.T("First Name Samples"), sg.Push(), sg.T("Last Name Samples"), sg.Push(), sg.T("E-mail Domain Samples")],
            [sg.Listbox(f_names, size= (18,5), key= "-FNAMES-"), sg.Push(), sg.Listbox(l_names, size= (18,5), key= "-LNAMES-"), sg.Push(), sg.Listbox(domains, size= (18,5), key= "-DOMAINS-")],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Add First Name on Column:", key = "-FIRSTNAME-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNFIRSTNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-FIRSTNAMETITLE-", default_text= "First Name")],
            [sg.Checkbox("Add Last Name on Column:", key = "-LASTNAME-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNLASTNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-LASTNAMETITLE-", default_text= "Last Name")],
            [sg.T("Else Add Full Name on Column:"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-NAMETITLE-", default_text= "Name")],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Add email address on column:", key= "-EMAIL-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMNMAIL-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-EMAILTITLE-", default_text= "E-mail Address")],
            [sg.HorizontalSeparator()],
            [sg.B("Preview"), sg.B("Sample Reload"), sg.Push(), sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))],
        ]
        name_window = sg.Window("Name and E-mail", layout_name, font= DEFAULT_FONT, modal= True, location= pos_name, icon= "icon.ico")
        #----------- DATA NAME LOOP ------------#
        while True:
            event, values = name_window.read()
            pos_name = name_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Back":
                break
            if event == "Sample Reload":
                f_names = read_sample_data("FIRST NAME")
                l_names = read_sample_data("LAST NAME")
                domains = read_sample_data("EMAIL DOMAIN")
                name_window.Element("-FNAMES-").update(f_names)
                name_window.Element("-LNAMES-").update(l_names)
                name_window.Element("-DOMAINS-").update(domains)
            if event == "Preview":
                preview_dataframe(dict, rows, pos_name)
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
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 160, 100), icon= "icon.ico").read(close= True)
                elif (same_column):
                    sg.Window("ERROR!", [[sg.T("Please specify different Columns for selected data")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 100, 100), icon= "icon.ico").read(close= True)
                else: #------------ LOGIC FOR ADDING THE DATA -----------#
                    col_names = "" #--------- COLUMN NAMES FOR SEPARATE NAME VALUES, TO DECLARE CONFIRMATION --------------#
                    mail_fnames = [] #--------- IF FIRST NAMES SELECTED, VALUES SAVED TO USE FOR EMAIL GENERATION -----------#
                    mail_lnames = [] #--------- IF LAST NAMES SELECTED, VALUES SAVED TO USE FOR EMAIL GENERATION -----------#
                    if values["-FIRSTNAME-"]: #------------ ADDING FIRST NAME ----------#
                        first_col_name = values["-COLUMNFIRSTNAME-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-FIRSTNAMETITLE-"]) < 21:
                            fnames = {first_col_name:[values["-FIRSTNAMETITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            for row in range(rows):
                                first = choice(f_names)
                                fnames[first_col_name].append(first)
                                if values ["-EMAIL-"]:
                                    mail_fnames.append(first)
                            col_names = first_col_name
                            dict.update(fnames)
                        else:
                            sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                [sg.Push(), sg.T("between 1 and 20 characters long"), sg.Push()],
                                [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)                   
                    if values["-LASTNAME-"]: #----------- ADDING LAST NAME ---------#
                        last_col_name = values["-COLUMNLASTNAME-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-LASTNAMETITLE-"]) < 21:
                            lnames = {last_col_name:[values["-LASTNAMETITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            for row in range(rows):
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
                                [sg.Push(), sg.T("between 1 and 20 characters long"), sg.Push()],
                                [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)         
                    if values["-FIRSTNAME-"] == False and values["-LASTNAME-"] == False: # --------- ADDING FULL NAME ----------#
                        full_col_name = values["-COLUMNNAME-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-NAMETITLE-"]) < 21:
                            names = {full_col_name:[values["-NAMETITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            if values ["-EMAIL-"]: # -------- IF ADDING E-MAIL AS WELL ---------#
                                mail_col_name = values["-COLUMNMAIL-"]
                                if 0 < len(values["-EMAILTITLE-"]) < 21:
                                    emails = {mail_col_name:[values["-EMAILTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                                    for row in range(rows):
                                        first = choice(f_names)
                                        last = choice(l_names)
                                        names[full_col_name].append(f"{first} {last}")
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{first}.{last}.{delta}@{domain}")
                                    dict.update(names)
                                    dict.update(emails)
                                    sg.Window("Done", [[sg.T(f"Data added on columns {full_col_name} and {mail_col_name}")],
                                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)
                                else:
                                    sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                        [sg.Push(), sg.T("between 1 and 20 characters long"), sg.Push()],
                                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)
                            else: # --------- IF ADDING ONLY NAME ---------#
                                for row in range(rows):
                                    first = choice(f_names)
                                    last = choice(l_names)
                                    names[full_col_name].append(f"{first} {last}")
                                dict.update(names)
                                sg.Window("Done", [[sg.T(f"Data added on column {full_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)
                        else:
                            sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                    [sg.Push(), sg.T("between 1 and 20 characters long"), sg.Push()],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)                    
                    else: #----------- IF SEPERATE DATA ADDED, DETERMINE E-MAIL IF SELECTED AND SHOW CONFIRMATION -----------#
                        if values ["-EMAIL-"]: # -------- IF ADDING E-MAIL AS WELL ---------#
                            mail_col_name = values["-COLUMNMAIL-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                            if 0 < len(values["-EMAILTITLE-"]) < 21:
                                emails = {mail_col_name:[values["-EMAILTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                                if len(mail_fnames) == 0: #--------- IF FIRST NAME MISSING, ONLY USE LAST NAME FOR EMNAIL ------------#
                                    for row in range(rows):
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{mail_lnames[row]}.{delta}@{domain}")
                                elif len(mail_lnames) == 0: #--------- IF LAST NAME MISSING, ONLY USE FIRST NAME FOR EMNAIL ------------#
                                    for row in range(rows):
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{mail_fnames[row]}.{delta}@{domain}")
                                else:
                                    for row in range(rows):
                                        domain = choice(domains)
                                        delta = randint(0,99)
                                        emails[mail_col_name].append(f"{mail_fnames[row]}.{mail_lnames[row]}.{delta}@{domain}")
                                dict.update(emails) #---------- ADD THE EMAILS TO DICTIONARY ---------#
                                sg.Window("Done", [[sg.T(f"Data added on columns {col_names} and {mail_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)
                            else:
                                sg.Window("ERROR", [[sg.T("Please ensure the column titles are")],
                                    [sg.Push(), sg.T("between 1 and 20 characters long"), sg.Push()],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)
                        else:
                            sg.Window("Done", [[sg.T(f"Data added on column(s) {col_names}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_name, 200, 100), icon= "icon.ico").read(close= True)
        name_window.close()
    #------------ FOR DATA NUMBER ----------#    
    if evt == "-NUMBER-":
        DECIMALS = [dec for dec in range(9)]
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
        pos_num = position_correction(win_pos, 180, 80)
        num_window = sg.Window("Number", layout_num, font= DEFAULT_FONT, modal= True, location= pos_num, icon= "icon.ico")
        #----------- DATA NUMBER LOOP -----------#
        while True:
            event, values = num_window.read()
            pos_num = num_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Back":
                break
            if event == "Clear":
                num_window.Element("-NUMMIN-").Update("")
                num_window.Element("-NUMMAX-").Update("")
            if event == "Preview":
                preview_dataframe(dict, rows, pos_num)
            if event == "Add":
                try:
                    if values["-COLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_num, 15, 40), icon= "icon.ico").read(close= True)
                    else:
                        if float(values["-NUMMAX-"]) < float(values["-NUMMIN-"]):
                            sg.Window("ERROR!", [[sg.T("Please ensure that Min < Max")],
                            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_num, 30, 40), icon= "icon.ico").read(close= True)
                        else:
                            num_col_name = values["-COLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                            if 0 < len(values["-NUMTITLE-"]) < 21:
                                numbers = {num_col_name:[values["-NUMTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                                if values["-DECIMALS-"] == 0:
                                    for row in range (rows):
                                        numbers[num_col_name].append(randint(int(values["-NUMMIN-"]), int(values["-NUMMAX-"])))
                                else:
                                    for row in range (rows):
                                        numbers[num_col_name].append(round(uniform(float(values["-NUMMIN-"]), float(values["-NUMMAX-"])), values["-DECIMALS-"]))
                                dict.update(numbers)
                                sg.Window("Done", [[sg.T(f"Data added on column {num_col_name}")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_num, 50, 40), icon= "icon.ico").read(close= True)
                            else:
                                sg.Window("ERROR", [[sg.Push(), sg.T("Please ensure the column title is"), sg.Push()],
                                    [sg.T("between 1 and 20 characters long")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_num, 10, 40), icon= "icon.ico").read(close= True)
                except:
                    sg.Window("ERROR!", [[sg.T("Please ensure the values are numbers")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_num, -10, 40), icon= "icon.ico").read(close= True)
        num_window.close()
    #------------ FOR DATA LOCATION ----------# 
    if evt == "-LOCATION-":
        street_name_1 = read_sample_data("STREET NAME 1")
        street_name_2 = read_sample_data("STREET NAME 2")
        city_and_state = read_sample_data("CITY AND STATE")
        pos_loc = position_correction(win_pos, -10, 80)
        layout_loc = [
            [sg.T("Street Name 1 Samples"), sg.Push(), sg.T("Street Name 2 Samples"), sg.T("City and State Samples")],
            [sg.Listbox(street_name_1, size= (15,5), key= "-SNAMES1-"), sg.Push(), sg.Listbox(street_name_2, size= (15,5), key= "-SNAMES2-"), sg.Push(), sg.Listbox(city_and_state, size= (15,5), key= "-CITYSTATE-")],
            [sg.HorizontalSeparator()],
            [sg.T("Add full address (incl. Street No. and Postal Code) on Column:"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-COLUMN-", readonly= True)],
            [sg.T("With title:"), sg.Push(), sg.I(size= (40, 1), key= "-LOCTITLE-", default_text= "Location")],
            [sg.HorizontalSeparator()],
            [sg.Checkbox("Add Street, City and State in 3 seperate columns:", key= "-SEPARATELOC-"), sg.Push(), sg.Combo(EXCEL_COLUMN, key = "-STREETCOLUMN-", readonly= True), sg.Combo(EXCEL_COLUMN, key = "-CITYCOLUMN-", readonly= True), sg.Combo(EXCEL_COLUMN, key = "-STATECOLUMN-", readonly= True)],
            [sg.HorizontalSeparator()],
            [sg.B("Preview"), sg.B("Sample Reload"), sg.Push(), sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
        ]
        loc_window = sg.Window("Location", layout_loc, font= DEFAULT_FONT, modal= True, location= pos_loc, icon= "icon.ico")
        #----------- DATA LOCATION LOOP ------------#
        while True:
            event, values = loc_window.read()
            pos_loc = loc_window.current_location()
            if event == sg.WINDOW_CLOSED or event == "Back":
                break
            if event == "Sample Reload":
                street_name_1 = read_sample_data("STREET NAME 1")
                street_name_2 = read_sample_data("STREET NAME 2")
                city_and_state = read_sample_data("CITY AND STATE")
                loc_window.Element("-SNAMES1-").update(street_name_1)
                loc_window.Element("-SNAMES2-").update(street_name_2)
                loc_window.Element("-CITYSTATE-").update(city_and_state)
            if event == "Preview":
                preview_dataframe(dict, rows, pos_loc)
            if event == "Add":
                if values["-SEPARATELOC-"]:
                    if values["-STREETCOLUMN-"] == "" or values["-CITYCOLUMN-"] == "" or values["-STATECOLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Columns")],
                        [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_loc, 180, 80), icon= "icon.ico").read(close= True)
                    elif values["-STREETCOLUMN-"] == values["-CITYCOLUMN-"] or values["-STREETCOLUMN-"] == values["-STATECOLUMN-"] or values["-STATECOLUMN-"] == values["-CITYCOLUMN-"]:
                        sg.Window("ERROR!", [[sg.T("Please specify different Columns for selected data")],
                            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_loc, 100, 100), icon= "icon.ico").read(close= True)
                    else:
                        street_col_name = values["-STREETCOLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        city_col_name = values["-CITYCOLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        state_col_name = values["-STATECOLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        streets = {street_col_name:["Street"]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                        cities = {city_col_name:["City"]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                        states = {state_col_name:["State"]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                        for row in range(rows):
                            s_name_1 = choice(street_name_1)
                            s_name_2 = choice(street_name_2)
                            citystate = choice(city_and_state)
                            city = citystate[:-4]
                            state = citystate[-2:]
                            streets[street_col_name].append(f"{s_name_1} {s_name_2}")
                            cities[city_col_name].append(city)
                            states[state_col_name].append(state)
                        dict.update(streets)
                        dict.update(cities)
                        dict.update(states)
                        sg.Window("Done", [[sg.T(f"Data added on columns {street_col_name}, {city_col_name} and {state_col_name}")],
                            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_loc, 150, 80), icon= "icon.ico").read(close= True)
                else:
                    if values["-COLUMN-"] == "":
                        sg.Window("ERROR!", [[sg.T("Please specify the target Column")],
                        [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_loc, 180, 80), icon= "icon.ico").read(close= True)
                    else:
                        loc_col_name = values["-COLUMN-"] #--------- COLUMN NAME TO USE AS DICTIONARY KEY---------#
                        if 0 < len(values["-LOCTITLE-"]) < 21:
                            locations = {loc_col_name:[values["-LOCTITLE-"]]} #-------- COLUMN AS KEY IN DICTIONARY AND TITLE AS FIRST ELEMENT OF ASSOCIATED LIST -------#
                            for row in range(rows):
                                s_number = randint(1,200)
                                s_name_1 = choice(street_name_1)
                                s_name_2 = choice(street_name_2)
                                citystate = choice(city_and_state)
                                postal_code = randint(10000, 99999)
                                locations[loc_col_name].append(f"{s_number} {s_name_1} {s_name_2}, {citystate}, {postal_code}")
                            dict.update(locations)
                            sg.Window("Done", [[sg.T(f"Data added on column {loc_col_name}")],
                                        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_loc, 200, 80), icon= "icon.ico").read(close= True)
                        else:
                            sg.Window("ERROR", [[sg.Push(), sg.T("Please ensure the column title is"), sg.Push()],
                                [sg.T("between 1 and 20 characters long")],
                                [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(pos_loc, 150, 80), icon= "icon.ico").read(close= True)
        loc_window.close()

def reset(win, win_pos, theme, dict):
    if len(dict) == 0:
        win.close()
        window = new_main_window(win_pos, theme)
        return window
    reset_position = position_correction(win_pos, 80, 80)   
    reset_window= sg.Window("WARNING!", [
        [sg.T("This will reset the app and all progress will be lost!")],
        [sg.Push(), sg.T("Would you like to save your progress?"), sg.Push()],
        [sg.Push(), sg.B("Save and Reset"), sg.B("Reset", button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")),sg.Push()]], font= DEFAULT_FONT, modal= True, location= reset_position, icon= "icon.ico")
    while True:
        event, values = reset_window.read()
        reset_position = reset_window.current_location()
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "Save and Reset":
            save_cancelled = save_dict(dict, position_correction(reset_position, -80, 0))
            if save_cancelled == False:
                dict.clear()
                reset_window.close()
                win.close()
                window = new_main_window(win_pos, theme)
                return window
        if event == "Reset":
            dict.clear()
            reset_window.close()
            win.close()
            window = new_main_window(win_pos, theme)
            return window
    reset_window.close()
    return win

def change_theme(theme, win, win_pos, dict):
    THEMES = sg.theme_list()
    theme_position = position_correction(win_pos, 100, 80)
    theme_window = sg.Window("Choose theme: ",
    [[sg.Combo(THEMES, key= "-THEME-", readonly= True, default_value= theme), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Cancel()]], font= DEFAULT_FONT, modal= True, location= theme_position)
    while True:
        event, values = theme_window.read()
        theme_position = theme_window.current_location()
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "OK":
            selected_theme = values["-THEME-"]
            theme_window.close()
            win.close()
            window = new_main_window(win_pos, selected_theme)
            if len(dict) != 0: #------------- IF THE DICTIONARY HAS DATA, ALSO SET THE ROWS AND ENALE THE BUTTONS IN MAIN WINDOW ----------#
                rows = len(next(iter(dict.values()))) - 1 #-------- MINUS 1 FOR TITLE ----------#
                window["-ROWS-"].update(rows)
                rows_are_set(window)
            return selected_theme, window
    theme_window.close()
    return theme, win

def rows_are_set(win):
    win.Element("-ROWS-").update(disabled= True)
    win.Element("-SETROWS-").update(disabled= True)
    win.Element("-NAME-").update(disabled= False)
    win.Element("-CLEARCOLUMN-").update(disabled= False)
    win.Element("-COLUMNTOCLEAR-").update(disabled= False)
    win.Element("-NUMBER-").update(disabled= False)
    win.Element("-LOCATION-").update(disabled= False)
    win.Element("Preview").update(disabled= False)
    win.Element("Generate").update(disabled= False)

# ---------- MAIN WINDOW SETUP FOR INTIALIZATION AND RESET ----------#
def new_main_window(pos, theme= DEFAULT_THEME):
    menu_def = [["File", ["New", "---", "Load", "Save", "---", "Theme", "---", "Exit"]], 
                ["Help", ["Get Started", "About"]]]
    if theme:
        sg.theme(theme)
    layout = [
        [sg.MenubarCustom(menu_def)],
        [sg.B("Output directory:", key= "-BROWSEOUT-"), sg.I(key= "-OUTPUT-", size= (40,1), default_text= Path.cwd(), disabled= True)],
        [sg.HorizontalSeparator()],
        [sg.T("Excel File Name:"), sg.Push(), sg.I(key= "-FILENAME-", size= (30,1), default_text= "dummPy")],
        [sg.T("Excel Sheet Name:"), sg.Push(), sg.I(key= "-SHEETNAME-", size= (30,1), default_text= "Sheet1")],
        [sg.HorizontalSeparator()],
        [sg.B("Set number of rows:", key= "-SETROWS-"), sg.I(size= (5,1), disabled_readonly_background_color="#5ebd78", default_text= 100, key= "-ROWS-"), sg.Push(), sg.B("Clear Data in Column:", key= "-CLEARCOLUMN-", disabled= True, disabled_button_color= ("#f2557a", None)), sg.Combo(EXCEL_COLUMN, key = "-COLUMNTOCLEAR-", readonly= True, disabled= True)],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Data Types:", font= DEFAULT_FONT+ ("bold",)), sg.Push()],
        [sg.T("Name and e-mail"), sg.Push(), sg.B("Configure", key= "-NAME-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Number"), sg.Push(),sg.B("Configure", key= "-NUMBER-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Location"), sg.Push(),sg.B("Configure", key= "-LOCATION-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.HorizontalSeparator()],
        [sg.B("Reset"), sg.Push(), sg.B("Preview", disabled= True, disabled_button_color= ("#f2557a", None)), sg.B("Generate", disabled= True, disabled_button_color= ("#f2557a", None), button_color= ("#292e2a","#5ebd78"))]
    ]
    return sg.Window("DummPy", layout, font= DEFAULT_FONT, enable_close_attempted_event= True, location= pos, icon= "icon.ico").finalize()

# ---------- MAIN WINDOW AND LOGIC LOOP ----------#
def main_window():
    DATA_CATEGORIES = ("-NAME-", "-NUMBER-", "-LOCATION-")
    current_theme = DEFAULT_THEME
    window_position= (None, None)
    window = new_main_window(window_position)
    dictionary= {}               
    while True: #------------ MAIN LOOP -------------#
        event, values = window.read()
        window_position = window.current_location()
        if event == sg.WINDOW_CLOSE_ATTEMPTED_EVENT or event == "Exit":
            if len(dictionary) == 0: #--------- IMMEDIATE EXIT ----------#
                break
            else:
                exit_pos = position_correction(window_position, 170, 80)
                exit_window = sg.Window("Exit",[
                    [sg.T("Confirm Exit or Save and Exit")], 
                    [sg.Push(), sg.B("Save and Exit"), sg.B("Exit", button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")), sg.Push()]], font= DEFAULT_FONT, location= exit_pos, modal= True, icon= "icon.ico")
                event, values = exit_window.read()
                exit_pos = exit_window.current_location()
                if event == sg.WINDOW_CLOSED or event == "Cancel":
                    exit_window.close()
                if event == "Save and Exit":
                    exit_window.close()
                    save_cancelled = save_dict(dictionary, position_correction(exit_pos, -160, 10))
                    if save_cancelled == False: 
                        break
                if event == "Exit":
                    break
        if event == "-SETROWS-":
            try:
                rows = int(values["-ROWS-"])
                if rows < 1 or rows > 99999:
                    sg.Window("ERROR!", [[sg.T("Please enter an integer between 1 and 9999")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 120, 80), icon= "icon.ico").read(close= True)
                else:
                    excel_rows = rows
                    rows_are_set(window)
            except:
                sg.Window("ERROR!", [[sg.T("Number of rows must be an integer")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 170, 80), icon= "icon.ico").read(close= True)
        if event in DATA_CATEGORIES:
            window.hide()
            configure(event, window_position, excel_rows, dictionary)
            window.un_hide()
        if event == "-CLEARCOLUMN-":
            column = values["-COLUMNTOCLEAR-"]
            if column in dictionary.keys():
                dictionary.pop(column, None)
                sg.Window("Done", [[sg.T(f"Data in Column {column} have been removed")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 170, 150), icon= "icon.ico").read(close= True)             
            else:
                sg.Window("ERROR!", [[sg.T(f"Column {column} does not exist in dictionary")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 170, 150), icon= "icon.ico").read(close= True)
        if event == "Save":
            if len(dictionary) == 0:
                sg.Window("ERROR!", [[sg.T("Cannot save an empty dictionary")],
                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 170, 80), icon= "icon.ico").read(close= True)
            else:
                save_cancelled = save_dict(dictionary, window_position) #--------- UNUSED VARIABLE, JUST NEED THE OPERATIONS HERE, NOT THE RETURNED VALUE ---------#
        if event == "Load":
            try:
                excel_rows, dictionary = load_dict(window_position)
                window["-ROWS-"].update(excel_rows)
                rows_are_set(window)
            except: #--------- IF WE CANCEL THE LOAD, RETURNS NONE AND GIVES ERROR, SO EXCEPTION NEEDED -----------#
                None
        if event == "About":
            url = "https://github.com/AntonisTorb/dummPy"
            about_window = sg.Window("About", [
                [sg.Push(), sg.T("~~DummPy~~", font= ("Arial", 30)), sg.Push()],
                [sg.T("Random Data Generator", font= ("Arial", 20))],
                [sg.Push(), sg.T("Version 1.0"), sg.Push()],
                [sg.HorizontalSeparator()],
                [sg.Push(), sg.T("Github Repository", key="-URL-", enable_events= True, tooltip= url, text_color= "Blue", background_color= "Grey",font= DEFAULT_FONT + ("underline",)), sg.Push()],
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 150, 80), icon= "icon.ico")
            while True:
                event, values = about_window.read()
                if event == sg.WINDOW_CLOSED or event == "OK":
                    break
                if event == "-URL-":
                    webbrowser.open(url)
            about_window.close()
        if event == "Get Started":
            webbrowser.open("https://github.com/AntonisTorb/dummPy/blob/main/User%20Guide.pdf")
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
            sheetname = values["-SHEETNAME-"]
            #---------- CHECK FILENAME LENGTH AND FOR ILLEGAL FILENAME CHARACTERS -----------#
            if len(filename) < 1 or len(filename) > 50:
                sg.Window("ERROR!", [[sg.T("File name must be between 1 and 50 characters long")],
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 80, 150), icon= "icon.ico").read(close= True)
            elif len(sheetname) < 1 or len(sheetname) > 31:
                sg.Window("ERROR!", [[sg.T("Sheet name must be between 1 and 31 characters long")],
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 80, 150), icon= "icon.ico").read(close= True)
            elif len(dictionary) == 0:
                sg.Window("ERROR!", [[sg.T("Cannot create empty sheet")],
                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 190, 150), icon= "icon.ico").read(close= True)
            else:
                no_illegal = True
                for illegal in PATTERN1:
                    if illegal in filename:
                        sg.Window("ERROR!", [[sg.T("Please do not use any of the following characters in file name:")],
                            [sg.Push(), sg.T("<>:/\|?*"), sg.Push()],
                            [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 50, 140), icon= "icon.ico").read(close= True)
                        no_illegal = False
                        break
                for illegal in PATTERN2:
                    if illegal in sheetname:
                        sg.Window("ERROR!", [[sg.T("Please do not use any of the following characters in sheet name:")],
                            [sg.Push(), sg.T(":/\?*[]"), sg.Push()],
                            [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 50, 140), icon= "icon.ico").read(close= True)
                        no_illegal = False
                        break
                if no_illegal:
                    df2 = dict_sort_for_df(dictionary, excel_rows)
                    file = values["-OUTPUT-"]
                    filepath = f"{file}\{filename}.xlsx"
                    file_exists = Path(filepath).exists()
                    if file_exists:
                        existing_file = pd.ExcelFile(filepath)
                        existing_sheets = existing_file.sheet_names
                        existing_file.close()
                        if sheetname in existing_sheets:
                            replace_window = sg.Window("Warning", [[sg.T(f"There is already a sheet named {sheetname} in {filename}.xlsx.")],
                                [sg.Push(),sg.T("Would you like to replace it?"), sg.Push()],
                                [sg.Push(), sg.B("Replace", button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 60, 140), icon= "icon.ico")
                            while True:
                                event, values = replace_window.read()
                                replace_pos = replace_window.current_location()
                                if event == sg.WINDOW_CLOSED or event == "Cancel":
                                    break
                                if event == "Replace":
                                    try:
                                        with pd.ExcelWriter(filepath, mode="a", if_sheet_exists= "replace") as writer:
                                            pd.DataFrame.from_dict(df2).to_excel(writer, header= False, index= False, sheet_name= sheetname)
                                            sg.Window("Done", [[sg.T(f"Data created in Sheet named {sheetname} in file named {filename}.xlsx.")],
                                                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(replace_pos, -30, 60), icon= "icon.ico").read(close= True)
                                        break
                                    except: #----------- IF THE FILE IS OPEN ------------#
                                        sg.Window("ERROR!", [[sg.T("Please close the excel file")],
                                            [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(replace_pos, 125, 60), icon= "icon.ico").read(close= True)
                            replace_window.close()
                        else:
                            try:
                                with pd.ExcelWriter(filepath, mode="a") as writer:
                                    pd.DataFrame.from_dict(df2).to_excel(writer, header= False, index= False, sheet_name= sheetname)
                                    sg.Window("Done", [[sg.T(f"Data created in Sheet named {sheetname} in file named {filename}.xlsx.")],
                                        [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 50, 140), icon= "icon.ico").read(close= True)
                            except: #----------- IF THE FILE IS OPEN ------------#
                                sg.Window("ERROR!", [[sg.T("Please close the excel file")],
                                    [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(replace_pos, 125, 60), icon= "icon.ico").read(close= True)
                    else: #---------- IF FILE DOES NOT EXIST, CREATE IT -----------#
                        with pd.ExcelWriter(filepath) as writer:
                            pd.DataFrame.from_dict(df2).to_excel(writer, header= False, index= False, sheet_name= sheetname)
                            sg.Window("Done", [[sg.T(f"Data created in Sheet named {sheetname} in file named {filename}.")],
                                [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]], font= DEFAULT_FONT, modal= True, location= position_correction(window_position, 50, 140), icon= "icon.ico").read(close= True)
    window.close()
        
if __name__ == "__main__":
    main_window()