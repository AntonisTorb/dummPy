import ast  # core python module
from pathlib import Path  # core python module
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
import operations
import messages
import global_constants


#----- class to read, store and reload the sample data -----#
class read_sample_data:

    data = {'FIRST NAME': [], 'LAST NAME': [], 'EMAIL DOMAIN': [], 'STREET NAME 1': [], 'STREET NAME 2': [], 'CITY AND STATE': []}
    
    def read_all(self): #----- read all sample data initially -----#
        try:
            sample_file = pd.ExcelFile("SAMPLE_DATA.xlsx")
            sheets = self.data.keys() #----- sheet names are predetermined -----#
            for sheet in sheets:
                try:
                    sample_sheet = pd.read_excel(sample_file, sheet_name= sheet, usecols= [0, 0], header= None).squeeze("columns").values.tolist()
                    self.data[sheet] = sample_sheet  
                except ValueError: #----- sheet is missing in "SAMPLE_DATA.xlsx" file-----#
                    messages.one_line_error_handler(f"{sheet} sample sheet is missing", (None, None))  
        except FileNotFoundError: #----- "SAMPLE_DATA.xlsx" file is missing from script/executable directory -----#
            messages.one_line_error_handler("Sample file is missing!", (None, None))

    def read_datatype(self, datatype, win_pos): #----- reloading sample data -----#
        match datatype: #----- determine the datatype to reload, no need to reload the entire file again -----#
            case "NAME":
                sheets = ["FIRST NAME", "LAST NAME", "EMAIL DOMAIN"]
            case "LOCATION":
                sheets = ["STREET NAME 1", "STREET NAME 2", "CITY AND STATE"]
        try:
            for sheet in sheets:
                try:
                    self.data[sheet] = pd.read_excel("SAMPLE_DATA.xlsx", sheet_name= sheet, usecols= [0, 0], header= None).squeeze("columns")
                except ValueError: #----- sheet is missing in "SAMPLE_DATA.xlsx" file -----#
                    messages.one_line_error_handler(f"{sheet} sample SHEET is missing", operations.position_correction(win_pos, 150, 100))
                    self.data[sheet] = []
        except FileNotFoundError: #----- "SAMPLE_DATA.xlsx" file is missing from script/executable directory -----#
            messages.one_line_error_handler("Sample file is missing!", operations.position_correction(win_pos, 150, 100))

#----- load dictionary from external text file -----#
def load_dict(win_pos):
    load_layout = [
        [sg.T("Load File:"), sg.I(key = "-LOADFILE-", disabled= True), sg.B("Browse")],
        [sg.Push(), sg.B("Load", button_color= ("#292e2a", "#5ebd78"), disabled= True, disabled_button_color= ("#f2557a", None)),  sg.B("Back", button_color= ("#ffffff", "#bf365f")), sg.Push()]
    ]
    load_window = sg.Window("Save", load_layout, modal= True, font= global_constants.DEFAULT_FONT, location= operations.position_correction(win_pos, -40, 80), icon= "icon.ico")
    while True:
        event, values = load_window.read()
        load_win_location = load_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Back":
                break
            case "Browse":
                file_to_load = sg.popup_get_file('',file_types=(("text Files", "*.txt*"),), no_window=True)
                if file_to_load:
                    load_window["-LOADFILE-"].update(Path(file_to_load))
                    load_window.Element("Load").update(disabled= False)
            case "Load":
                with open(file_to_load,"r") as loadfile:
                    temp_dict = loadfile.readline()
                try:
                    dict = ast.literal_eval(temp_dict) #----- string to dictionary -----#
                    rows = len(next(iter(dict.values()))) - 1 #----- minus 1 for title -----#
                    messages.operation_successful("Dictionary Loaded", operations.position_correction(load_win_location, 250, 30))
                    break
                except: #----- unable to specify excveption type as no message is shown, just load window becomes unresponsive if invalid input (not a dictionary) -----#
                    messages.one_line_error_handler("Save file is invalid or corrupted", operations.position_correction(load_win_location, 250, 30))
    load_window.close()
    return rows, dict

