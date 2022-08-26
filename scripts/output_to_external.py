from pathlib import Path  # core python module
import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install pysimplegui
import scripts.operations as operations
import scripts.global_constants as global_constants
import scripts.messages as messages

#----- save current dictionary in external text file -----#
def save_dict(dict, win_pos):
    save_dir = Path.cwd()
    save_layout = [
        [sg.T("Save directory:"), sg.Push(), sg.I(key = "-SAVEDIR-", disabled= True, default_text = save_dir), sg.B("Browse")],
        [sg.T("Filename:"), sg.I(key = "-SAVEFILE-", default_text= "Save1"), sg.B("Save", button_color= ("#292e2a", "#5ebd78")),  sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
    ]
    save_window = sg.Window("Save", save_layout, modal= True, font= global_constants.DEFAULT_FONT, location= operations.position_correction(win_pos, -60, 80), icon= "icon.ico")
    while True:
        cancelled = False
        event, values = save_window.read()
        save_win_location = save_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Back":
                cancelled = True
                break
            case "Browse":
                save_dir = sg.popup_get_folder('', no_window=True)
                if save_dir:
                    save_window["-SAVEDIR-"].update(Path(save_dir))
            case "Save":
                no_illegal = True
                filename = values["-SAVEFILE-"]
                for illegal in global_constants.PATTERN1:
                    if illegal in filename:
                        messages.two_line_error_handler("Please do not use any of the following characters in filename:", "<>:/\|?*", operations.position_correction(save_win_location, 125, 70))
                        no_illegal = False
                        break
                if no_illegal:
                    filepath = Path(values["-SAVEDIR-"]) / f"{filename}.txt"
                    with open(filepath,"w") as savefile:
                        savefile.write(str(dict))
                        messages.operation_successful(f"Dictionary saved as {filename}.txt", operations.position_correction(save_win_location, 180, 80))
                    break
    save_window.close()
    if cancelled: #----- since save is called from functions such as reset, we want to avoid further operations if cancelled, so raising exception and handling it in relevant spots -----#
        raise RuntimeError("Save operations cancelled")

#----- generates the excel file or appends the sheet on existing file using the data provided -----#
def excel_file_generation(file_name, sheet_name, folder, dict, win_pos, rows):
    PATTERN2= ( ":" , "/" , "\\" , "?" , "*" , "[" , "]" ) #----- illegal sheet name characters -----#
    #----- checking file name length -----#
    if len(file_name) < 1 or len(file_name) > 50:
        messages.one_line_error_handler("File name must be between 1 and 50 characters long", operations.position_correction(win_pos, 80, 150))
    elif len(sheet_name) < 1 or len(sheet_name) > 31:
        messages.one_line_error_handler("Sheet name must be between 1 and 31 characters long", operations.position_correction(win_pos, 80, 150))
    elif len(dict) == 0:
        messages.one_line_error_handler("Cannot create empty sheet", operations.position_correction(win_pos, 190, 150))
    else:
        no_illegal = True
        for illegal in global_constants.PATTERN1: #----- checking file name for illegal characters -----#
            if illegal in file_name:
                messages.two_line_error_handler("Please do not use any of the following characters in file name:", "<>:/\|?*", operations.position_correction(win_pos, 50, 140))
                no_illegal = False
                break
        for illegal in PATTERN2: #----- checking sheet name for illegal characters -----#
            if illegal in sheet_name:
                messages.two_line_error_handler("Please do not use any of the following characters in sheet name:", ":/\?*[]", location= operations.position_correction(win_pos, 50, 140))
                no_illegal = False
                break
        if no_illegal:
            df2 = operations.dict_sort_for_df(dict, rows)
            filepath = f"{folder}\{file_name}.xlsx"
            file_exists = Path(filepath).exists()
            if file_exists:
                with pd.ExcelFile(filepath) as existing_file:
                    existing_sheets = existing_file.sheet_names
                if sheet_name in existing_sheets: #----- file name exists, and sheet name exists, so confirmation of overwrite needed -----#
                    replace_window = sg.Window("Warning", [[sg.T(f"There is already a sheet named {sheet_name} in {file_name}.xlsx.")],
                        [sg.Push(),sg.T("Would you like to replace it?"), sg.Push()],
                        [sg.Push(), sg.B("Replace", button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")), sg.Push()]], font= global_constants.DEFAULT_FONT, modal= True, location= operations.position_correction(win_pos, 60, 140), icon= "icon.ico")
                    while True:
                        event, values = replace_window.read()
                        replace_pos = replace_window.current_location()
                        match event:

                            case sg.WINDOW_CLOSED | "Cancel":
                                break
                            case "Replace":
                                try:
                                    with pd.ExcelWriter(filepath, mode="a", if_sheet_exists= "replace") as writer:
                                        pd.DataFrame.from_dict(df2).to_excel(writer, header= False, index= False, sheet_name= sheet_name)
                                    messages.operation_successful(f"Data created in Sheet named {sheet_name} in file named {file_name}.xlsx.", operations.position_correction(replace_pos, -30, 60))
                                    break
                                except PermissionError: #----- if the file is open -----#
                                    messages.one_line_error_handler("Please close the excel file", operations.position_correction(replace_pos, 125, 60))
                    replace_window.close()
                else: #----- file name exists, but sheet name does not, so we can append -----#
                    try:
                        with pd.ExcelWriter(filepath, mode="a") as writer:
                            pd.DataFrame.from_dict(df2).to_excel(writer, header= False, index= False, sheet_name= sheet_name)
                        messages.operation_successful(f"Data created in Sheet named {sheet_name} in file named {file_name}.xlsx.", operations.position_correction(win_pos, 50, 140))
                    except PermissionError: #----- if the file is open -----#
                        messages.one_line_error_handler("Please close the excel file", operations.position_correction(win_pos, 200, 140))
            else: #----- if file does not exist, create it -----#
                with pd.ExcelWriter(filepath) as writer:
                    pd.DataFrame.from_dict(df2).to_excel(writer, header= False, index= False, sheet_name= sheet_name)
                messages.operation_successful(f"Data created in Sheet named {sheet_name} in file named {file_name}.xlsx.", operations.position_correction(win_pos, 50, 140))