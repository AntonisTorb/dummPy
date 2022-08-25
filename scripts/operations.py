import webbrowser  # core python module
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
from pathlib import Path  # core python module
import scripts.global_constants as global_constants
import scripts.output_to_external as output_to_external

# ---------- adjust popup window location -----#
def position_correction(win_pos, dx, dy):
    corrected = list(win_pos)
    corrected[0] += dx
    corrected[1] += dy
    return tuple(corrected)

#----- sort the dictionary by key (column name) and add blanck columns when missing -----#
def dict_sort_for_df(dict, rows):
    new_dict = {}
    last_col_ch_num = 0
    #----- determine the character number of the last column of the dict -----#
    for key in dict:
        if ord(key) > last_col_ch_num:
            last_col_ch_num = ord(key)
    excel = [chr(ch_num) for ch_num in list(range(ord('A'), last_col_ch_num + 1))] #----- create a list with column names up to the last -----#
    #----- new dict with values from previous, sorted, and with blanck columns where there is no data -----#
    for col in excel:
        try:
            new_dict[col] = dict[col]
        except KeyError: #----- if the column does not exist in dict, add blanks -----#
            new_dict[col] = []
            for row in range(rows + 1):
                new_dict[col].append("")
    return new_dict

#----- shows a preview of the dataframe that will be used to create the excel file -----#
def preview_dataframe(dict, rows, win_pos):
    temp_df = pd.DataFrame.from_dict(dict_sort_for_df(dict, rows))
    headers = list(temp_df.head())
    val =  temp_df.values.tolist()
    preview_layout = [
        [sg.Table(val, headings= headers, num_rows= 20, display_row_numbers= True)],
        [sg.OK(button_color= ("#292e2a", "#5ebd78"))]
    ]
    sg.Window("Dataframe Preview", preview_layout, modal= True, font= global_constants.DEFAULT_FONT, location= win_pos, grab_anywhere= True, icon= "icon.ico").read(close= True)

#----- resets the main window to the default state -----#
def reset(win, win_pos, theme, dict, saved_dict):
    if len(dict) == 0 or dict == saved_dict: #----- if the dictionay is empty, or has been last saved or loaded, reset is instant -----#
        win.close()
        window = new_main_window(win_pos, theme)
        return window
    reset_window= sg.Window("WARNING!", [
        [sg.T("This will reset the app and all progress will be lost!")],
        [sg.Push(), sg.T("Would you like to save your progress?"), sg.Push()],
        [sg.Push(), sg.B("Save and Reset"), sg.B("Reset", button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")),sg.Push()]
    ], font= global_constants.DEFAULT_FONT, modal= True, location= position_correction(win_pos, 80, 80)  , icon= "icon.ico")
    while True:
        event, values = reset_window.read()
        reset_position = reset_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Cancel":
                break
            case "Save and Reset":
                try:
                    output_to_external.save_dict(dict, position_correction(reset_position, -80, 0))
                    dict.clear()
                    reset_window.close()
                    win.close()
                    window = new_main_window(win_pos, theme)
                    return window
                except RuntimeError: #----- if save cancelled -----#
                    None
            case "Reset":
                dict.clear()
                reset_window.close()
                win.close()
                window = new_main_window(win_pos, theme)
                return window
    reset_window.close()
    return win #----- if canceling reset, return the current main window -----#

#----- changes the theme of all the windows -----#
def change_theme(theme, win, win_pos, dict):
    THEMES = sg.theme_list()
    theme_position = position_correction(win_pos, 100, 80)
    theme_window = sg.Window("Choose theme: ",[
        [sg.Combo(THEMES, key= "-THEME-", readonly= True, default_value= theme), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Cancel()]
    ], font= global_constants.DEFAULT_FONT, modal= True, location= theme_position)
    while True:
        event, values = theme_window.read()
        theme_position = theme_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Cancel":
                break
            case "OK":
                selected_theme = values["-THEME-"]
                theme_window.close()
                win.close()
                window = new_main_window(win_pos, selected_theme)
                if len(dict) != 0: #----- if the dictionary has data, also set the rows and enale the buttons in main window -----#
                    rows = len(next(iter(dict.values()))) - 1 #----- minus 1 for title -----#
                    window["-ROWS-"].update(rows)
                    rows_are_set(window)
                return selected_theme, window
    theme_window.close()
    return theme, win #----- if canceling theme change, return the current theme and main window -----# 

#----- main window setup for intialization and reset -----#
def new_main_window(win_pos= (None, None), theme= global_constants.DEFAULT_THEME):
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
        [sg.B("Set number of rows:", key= "-SETROWS-"), sg.I(size= (5,1), disabled_readonly_background_color="#5ebd78", default_text= 100, key= "-ROWS-"), sg.Push(), sg.B("Clear Data in Column:", key= "-CLEARCOLUMN-", disabled= True, disabled_button_color= ("#f2557a", None)), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMNTOCLEAR-", readonly= True, disabled= True)],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Data Types:", font= global_constants.DEFAULT_FONT+ ("bold",)), sg.Push()],
        [sg.T("Name and e-mail"), sg.Push(), sg.B("Configure", key= "-NAME-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Number"), sg.Push(), sg.B("Configure", key= "-NUMBER-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Location"), sg.Push(), sg.B("Configure", key= "-LOCATION-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.T("Date"), sg.Push(), sg.B("Configure", key= "-DATE-", disabled= True, disabled_button_color= ("#f2557a", None))],
        [sg.HorizontalSeparator()],
        [sg.B("Reset"), sg.Push(), sg.B("Preview", disabled= True, disabled_button_color= ("#f2557a", None)), sg.B("Generate", disabled= True, disabled_button_color= ("#f2557a", None), button_color= ("#292e2a","#5ebd78"))]
    ]
    return sg.Window("DummPy", layout, font= global_constants.DEFAULT_FONT, enable_close_attempted_event= True, location= win_pos, icon= "icon.ico").finalize()

#----- shows the "about" window -----#
def about_window(win_pos):
    url = "https://github.com/AntonisTorb/dummPy"
    about_window = sg.Window("About", [
        [sg.Push(), sg.T("~~DummPy~~", font= ("Arial", 30)), sg.Push()],
        [sg.T("Random Data Generator", font= ("Arial", 20))],
        [sg.Push(), sg.T("Version 1.0.1"), sg.Push()],
        [sg.HorizontalSeparator()],
        [sg.Push(), sg.T("Github Repository", key="-URL-", enable_events= True, tooltip= url, text_color= "Blue", background_color= "Grey",font= global_constants.DEFAULT_FONT + ("underline",)), sg.Push()],
        [sg.Push(), sg.OK(button_color= ("#292e2a","#5ebd78")), sg.Push()]
    ], font= global_constants.DEFAULT_FONT, modal= True, location= win_pos, icon= "icon.ico")
    while True:
        event, values = about_window.read()
        match event:
            case sg.WINDOW_CLOSED | "OK":
                break
            case "-URL-":
                webbrowser.open(url)
    about_window.close()

#----- disables/enables elements when rows are set -----#
def rows_are_set(win):
    win.Element("-ROWS-").update(disabled= True)
    win.Element("-SETROWS-").update(disabled= True)
    win.Element("-NAME-").update(disabled= False)
    win.Element("-CLEARCOLUMN-").update(disabled= False)
    win.Element("-COLUMNTOCLEAR-").update(disabled= False)
    win.Element("-NUMBER-").update(disabled= False)
    win.Element("-LOCATION-").update(disabled= False)
    win.Element("-DATE-").update(disabled= False)
    win.Element("Preview").update(disabled= False)
    win.Element("Generate").update(disabled= False)

