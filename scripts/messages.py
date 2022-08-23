import PySimpleGUI as sg  # pip install pysimplegui
import scripts.global_constants as global_constants

# ---------- shorter error message -----#
def one_line_error_handler(text, position):
    sg.Window("ERROR!", [
        [sg.Push(), sg.T(text), sg.Push()],
        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]
    ], font= global_constants.DEFAULT_FONT, location= position, modal= True, icon= "icon.ico").read(close= True)

# ---------- longer error message -----#
def two_line_error_handler(text1, text2, position):
    sg.Window("ERROR!", [
        [sg.Push(), sg.T(text1), sg.Push()],
        [sg.Push(), sg.T(text2), sg.Push()],
        [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]
    ], font= global_constants.DEFAULT_FONT, modal= True, location= position, icon= "icon.ico").read(close= True)

# ---------- successful operation message -----#
def operation_successful(text, position):
    sg.Window("Done", [[sg.T(text)],
            [sg.Push(), sg.OK(button_color= ("#292e2a", "#5ebd78")), sg.Push()]], font= global_constants.DEFAULT_FONT, modal= True, location= position, icon= "icon.ico").read(close= True)

