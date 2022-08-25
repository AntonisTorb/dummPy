import scripts.configuration as configuration
import scripts.input_from_external as input_from_external
import scripts.messages as messages
import scripts.operations as operations
import scripts.output_to_external as output_to_external
import scripts.global_constants as global_constants
import webbrowser  # core python module
from pathlib import Path  # core python module
import PySimpleGUI as sg  # pip install pysimplegui
from random import choice, randint, uniform  # core python module

#----- main window and logic loop -----#
def main_window():
    current_theme = global_constants.DEFAULT_THEME
    window = operations.new_main_window()
    sample_data = input_from_external.read_sample_data()
    sample_data.read_all()
    dictionary= {}
    last_saved = {}

    while True: #----- main loop -----#
        event, values = window.read()
        window_position = window.current_location()
        match event: #----- event handling -----#
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "Exit":
                if len(dictionary) == 0 or dictionary == last_saved: #----- immediate exit -----#
                    break
                else:
                    exit_pos = operations.position_correction(window_position, 170, 80)
                    exit_window = sg.Window("Exit",[
                        [sg.T("Confirm Exit or Save and Exit")], 
                        [sg.Push(), sg.B("Save and Exit"), sg.B("Exit", button_color= ("#292e2a","#5ebd78")), sg.Cancel(button_color= ("#ffffff","#bf365f")), sg.Push()]
                    ], font= global_constants.DEFAULT_FONT, location= exit_pos, modal= True, icon= "icon.ico")
                    event, values = exit_window.read()
                    exit_pos = exit_window.current_location()
                    if event == sg.WINDOW_CLOSED or event == "Cancel":
                        exit_window.close()
                    if event == "Save and Exit":
                        exit_window.close()
                        try:
                            operations.save_dict(dictionary, operations.position_correction(exit_pos, -160, 10))
                            break
                        except RuntimeError: #----- if save cancelled -----#
                            None
                    if event == "Exit":
                        break
            case "-SETROWS-":          
                try:
                    rows = int(values["-ROWS-"])
                    if rows < 1 or rows > 99999:
                        messages.one_line_error_handler("Please enter an integer between 1 and 9999", operations.position_correction(window_position, 120, 80))
                    else:
                        excel_rows = rows
                        operations.rows_are_set(window)
                except ValueError: #----- if not an integer -----#
                    messages.one_line_error_handler("Number of rows must be an integer", operations.position_correction(window_position, 170, 80))
            case "-NAME-":
                window.hide()
                configuration.configure_name_email(window_position, excel_rows, sample_data, dictionary)
                window.un_hide()
            case "-NUMBER-":
                window.hide()
                configuration.configure_number(window_position, excel_rows, dictionary)
                window.un_hide()
            case "-LOCATION-":
                window.hide()
                configuration.configure_location(window_position, excel_rows, sample_data, dictionary)
                window.un_hide()
            case "-DATE-":
                #window.hide()
                configuration.configure_date(window_position, excel_rows, dictionary)
                #window.un_hide()
            case "-CLEARCOLUMN-":
                column = values["-COLUMNTOCLEAR-"]
                if column in dictionary.keys():
                    dictionary.pop(column, None)
                    messages.operation_successful(f"Data in Column {column} have been removed", operations.position_correction(window_position, 170, 150))
                else:
                    messages.one_line_error_handler(f"Column {column} does not exist in dictionary", operations.position_correction(window_position, 170, 150))
            case "Save":
                if len(dictionary) == 0:
                    messages.one_line_error_handler("Cannot save an empty dictionary", operations.position_correction(window_position, 170, 80))
                else:
                    try:
                        operations.save_dict(dictionary, window_position)
                        last_saved = dictionary.copy()
                    except RuntimeError: #----- if save cancelled -----#
                        None
            case "Load":
                try:
                    excel_rows, dictionary = input_from_external.load_dict(window_position)
                    last_saved = dictionary.copy()
                    window["-ROWS-"].update(excel_rows)
                    operations.rows_are_set(window)
                except UnboundLocalError: #----- if we cancel the load we get an error, so exception needed and returning none -----#
                    None
            case "About":
                operations.about_window(operations.position_correction(window_position, 150, 80))
            case "Get Started":
                webbrowser.open("https://github.com/AntonisTorb/dummPy/blob/main/User%20Guide.pdf")
            case "-BROWSEOUT-":
                path_temp = values["-OUTPUT-"]
                folder = sg.popup_get_folder('', no_window=True)
                if folder:
                    window["-OUTPUT-"].update(Path(folder))
                else:
                    window["-OUTPUT-"].update(path_temp)
            case "Theme":
                current_theme, window = operations.change_theme(current_theme, window, window_position, dictionary)
            case "Reset" | "New":
                window = operations.reset(window, window_position, current_theme, dictionary, last_saved)
            case "Preview":
                operations.preview_dataframe(dictionary, excel_rows, window_position)
            case "Generate":
                filename = values["-FILENAME-"]
                sheetname = values["-SHEETNAME-"]
                folderpath = values["-OUTPUT-"]
                output_to_external.excel_file_generation(filename, sheetname, folderpath, dictionary, window_position, excel_rows)
    window.close()
        
if __name__ == "__main__":
    main_window()