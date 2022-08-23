from random import choice, randint, uniform  # core python module
import PySimpleGUI as sg  # pip install pysimplegui
import scripts.operations as operations
import scripts.messages as messages
import scripts.global_constants as global_constants

#----- called if we are adding e-mail address, repeat operations -----#
def email_generation(domains, dict, column, row, *, seperate, column2= None):
    domain = choice(domains)
    delta = randint(0,99)
    name1 = dict[column][row]
    name2 = ""
    if seperate: #----- if adding seperate first name and last name, we need to include both for the emial generation -----#
        name2 = f"{dict[column2][row]}"
    return f"{name1}{name2}{delta}@{domain}".replace(" ", "").lower() #----- replacing blancks if we added full name, and lowercase -----#

#----- configure name and email address generation and storage based on input -----#
def configure_name_email(win_pos, rows, sample, dict):
    first_names = sample.data["FIRST NAME"]
    last_names = sample.data["LAST NAME"]
    domains = sample.data["EMAIL DOMAIN"]
    pos_name = operations.position_correction(win_pos, -25, 80)

    layout_name = [
        [sg.T("First Name Samples"), sg.Push(), sg.T("Last Name Samples"), sg.Push(), sg.T("E-mail Domain Samples")],
        [sg.Listbox(first_names, size= (18,5), key= "-FNAMES-"), sg.Push(), sg.Listbox(last_names, size= (18,5), key= "-LNAMES-"), sg.Push(), sg.Listbox(domains, size= (18,5), key= "-DOMAINS-")],
        [sg.HorizontalSeparator()],
        [sg.Checkbox("Add Full Name on Column:", key= "-FULLNAME-", enable_events= True, default= True), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMNNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-NAMETITLE-", default_text= "Name")],
        [sg.HorizontalSeparator()],
        [sg.Checkbox("Add First Name on Column:", key = "-FIRSTNAME-", enable_events= True), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMNFIRSTNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-FIRSTNAMETITLE-", default_text= "First Name")],
        [sg.Checkbox("Add Last Name on Column:", key = "-LASTNAME-", enable_events= True), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMNLASTNAME-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-LASTNAMETITLE-", default_text= "Last Name")],
        [sg.HorizontalSeparator()],
        [sg.Checkbox("Add email address on column:", key= "-EMAIL-"), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMNMAIL-", readonly= True), sg.T("with title:"), sg.I(size= (20, 1), key= "-EMAILTITLE-", default_text= "E-mail Address")],
        [sg.HorizontalSeparator()],
        [sg.B("Preview"), sg.B("Sample Reload"), sg.Push(), sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))],
    ]
    name_window = sg.Window("Name and E-mail", layout_name, font= global_constants.DEFAULT_FONT, modal= True, location= pos_name, icon= "icon.ico")

    #----- DATA NAME LOOP -----#
    while True:
        event, values = name_window.read()
        pos_name = name_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Back":
                break
            case "Sample Reload":
                sample.read_datatype("NAME", pos_name)
                first_names = sample.data["FIRST NAME"]
                last_names = sample.data["LAST NAME"]
                domains = sample.data["EMAIL DOMAIN"]
                name_window.Element("-FNAMES-").update(first_names)
                name_window.Element("-LNAMES-").update(last_names)
                name_window.Element("-DOMAINS-").update(domains)
            case "Preview":
                operations.preview_dataframe(dict, rows, pos_name)
            case "-FULLNAME-":
                if not values["-FULLNAME-"] and not values["-FIRSTNAME-"] and not values["-LASTNAME-"]: #----- this prevents all 3 checkboxes being false at once. if the only true checkbox gets turned false, turns it back to true -----#
                    name_window.Element("-FULLNAME-").update(True)
                name_window.Element("-FIRSTNAME-").update(False) #----- options are mutually exclusive between full name and seperate first and last name -----#
                name_window.Element("-LASTNAME-").update(False)
            case "-FIRSTNAME-":
                if not values["-FIRSTNAME-"] and not values["-FULLNAME-"] and not values["-LASTNAME-"]: #----- this prevents all 3 checkboxes being false at once. if the only true checkbox gets turned false, turns it back to true -----#
                    name_window.Element("-FIRSTNAME-").update(True)
                name_window.Element("-FULLNAME-").update(False)
            case "-LASTNAME-":
                if not values["-LASTNAME-"] and not values["-FULLNAME-"] and not values["-FIRSTNAME-"]: #----- this prevents all 3 checkboxes being false at once. if the only true checkbox gets turned false, turns it back to true -----#
                    name_window.Element("-LASTNAME-").update(True)
                name_window.Element("-FULLNAME-").update(False)
            case "Add":
                #----- declaring error conditions for readability -----#
                unspecified_column = ((values["-COLUMNNAME-"] == "" and values["-FULLNAME-"]) or #----- if fullname selected and column not selected -----#
                    (values["-COLUMNFIRSTNAME-"] == "" and values["-FIRSTNAME-"]) or #----- if first name selected and column not selected -----#
                    (values["-COLUMNLASTNAME-"] == "" and values["-LASTNAME-"]) or #----- if last name selected and column not selected -----#
                    (values["-COLUMNMAIL-"] == "" and values ["-EMAIL-"])) #----- if email selected and column not selected -----#
                same_column = ((values["-COLUMNNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"] and values["-FULLNAME-"]) or #----- full name email selected, and name column == email column -----#
                    (values["-COLUMNFIRSTNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"] and values["-FIRSTNAME-"]) or #----- first name and email selected and same column selected -----#
                    (values["-COLUMNLASTNAME-"] == values["-COLUMNMAIL-"] and values ["-EMAIL-"] and values["-LASTNAME-"]) or #----- last name and email selected and same column selected -----#
                    (values["-COLUMNFIRSTNAME-"] == values["-COLUMNLASTNAME-"] and values ["-FIRSTNAME-"] and values["-LASTNAME-"])) #----- first name and last name selected and same column selected -----#                       
                if (unspecified_column):
                    messages.one_line_error_handler("Please specify the target Column(s)", operations.position_correction(pos_name, 160, 100))
                elif (same_column):
                    messages.one_line_error_handler("Please specify different Columns for selected data", operations.position_correction(pos_name, 100, 100))
                else: #----- logic for adding the data -----#
                    columns_of_names = "" #----- columns where name data were added, to compose confirmation message-----#
                    if values["-FULLNAME-"]: # --------- adding full name -----#
                        col_full_name = values["-COLUMNNAME-"] #----- column name to use as dictionary key-----#
                        if len(values["-NAMETITLE-"]) < 0 or len(values["-NAMETITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                        else: #----- operations if all checks pass -----#
                            names = {col_full_name:[values["-NAMETITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                            for _ in range(rows):
                                first = choice(first_names)
                                last = choice(last_names)
                                names[col_full_name].append(f"{first} {last}")
                            columns_of_names = col_full_name
                            dict.update(names)
                    if values["-FIRSTNAME-"]: #----- adding first name -----#
                        col_first_name = values["-COLUMNFIRSTNAME-"] #----- column name to use as dictionary key-----#
                        if len(values["-FIRSTNAMETITLE-"]) < 0 or len(values["-FIRSTNAMETITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                        else: #----- operations if all checks pass -----#
                            fnames = {col_first_name:[values["-FIRSTNAMETITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                            for row in range(rows):
                                first = choice(first_names)
                                fnames[col_first_name].append(first)
                            columns_of_names = col_first_name
                            dict.update(fnames)
                    if values["-LASTNAME-"]: #----- adding last name -----#
                        col_last_name = values["-COLUMNLASTNAME-"] #----- column name to use as dictionary key-----#
                        if len(values["-LASTNAMETITLE-"]) < 0 or len(values["-LASTNAMETITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                        else: #----- operations if all checks pass -----#
                            lnames = {col_last_name:[values["-LASTNAMETITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                            for row in range(rows):
                                last = choice(last_names)
                                lnames[col_last_name].append(last)
                            if columns_of_names == "": #----- if no column for first name -----#
                                columns_of_names = col_last_name
                            else:
                                columns_of_names = f"{columns_of_names}, {col_last_name}"
                            dict.update(lnames)
                    if values ["-EMAIL-"]: # -------- if adding e-mail as well -----#
                        mail_col_name = values["-COLUMNMAIL-"] #----- column name to use as dictionary key-----#
                        if len(values["-EMAILTITLE-"]) < 0 or len(values["-EMAILTITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                        else:
                            emails = {mail_col_name:[values["-EMAILTITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                            if values["-FULLNAME-"]: #----- full name for email -----#
                                generated_mail_addresses = [email_generation(domains, dict, values["-COLUMNNAME-"], row, seperate= False) for row in range(1, rows + 1)]
                                emails[mail_col_name] += generated_mail_addresses
                            elif values["-FIRSTNAME-"] and not values["-LASTNAME-"]: #----- if last name not added, only use first name for email -----#
                                generated_mail_addresses = [email_generation(domains, dict, values["-COLUMNFIRSTNAME-"], row, seperate= False) for row in range(1, rows + 1)]
                                emails[mail_col_name] += generated_mail_addresses
                            elif values["-LASTNAME-"] and not values["-FIRSTNAME-"]: #----- if first name not added, only use last name for email -----#
                                generated_mail_addresses = [email_generation(domains, dict, values["-COLUMNLASTNAME-"], row, seperate= False) for row in range(1, rows + 1)]
                                emails[mail_col_name] += generated_mail_addresses
                            elif values["-FIRSTNAME-"] and values["-LASTNAME-"]: #----- both first and last name for email -----#
                                generated_mail_addresses = [email_generation(domains, dict, values["-COLUMNFIRSTNAME-"], row, seperate= True, column2= values["-COLUMNLASTNAME-"]) for row in range(1, rows + 1)]
                                emails[mail_col_name] += generated_mail_addresses
                            dict.update(emails) #----- add the emails to dictionary -----#
                            messages.operation_successful(f"Data added on columns {columns_of_names} and {mail_col_name}", operations.position_correction(pos_name, 200, 100))
                    else: #----- when not adding email, show confirmation for the name data added -----#
                        messages.operation_successful(f"Data added on column(s) {columns_of_names}", operations.position_correction(pos_name, 200, 100))
    name_window.close()

#----- configure number generation and storage based on input -----#
def configure_number(win_pos, rows, dict):
    DECIMALS = [dec for dec in range(9)] #----- up to 8 decimal places -----#
    layout_num = [
        [sg.T("Min:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMIN-")],
        [sg.T("Max:",size = 4), sg.Push(), sg.I(size= (10,1), key= "-NUMMAX-")],
        [sg.T("Decimals:"), sg.Push(), sg.Combo(DECIMALS, default_value= 0, key = "-DECIMALS-", readonly=True)],
        [sg.HorizontalSeparator()],
        [sg.T("Add to Column:"), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMN-", readonly= True)],
        [sg.T("With title:"), sg.I(size= (20, 1), key= "-NUMTITLE-", default_text= "Number")],
        [sg.HorizontalSeparator()],
        [sg.B("Clear"), sg.B("Preview"), sg.Push(),sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
    ]
    pos_num = operations.position_correction(win_pos, 180, 80)
    num_window = sg.Window("Number", layout_num, font= global_constants.DEFAULT_FONT, modal= True, location= pos_num, icon= "icon.ico")
    #----- data number loop -----#
    while True:
        event, values = num_window.read()
        pos_num = num_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Back":
                break
            case "Clear":
                num_window.Element("-NUMMIN-").Update("")
                num_window.Element("-NUMMAX-").Update("")
            case "Preview":
                operations.preview_dataframe(dict, rows, pos_num)
            case "Add":
                try:
                    if values["-COLUMN-"] == "":
                        messages.one_line_error_handler("Please specify the target Column", operations.position_correction(pos_num, 15, 40))
                    else:
                        if float(values["-NUMMAX-"]) < float(values["-NUMMIN-"]):
                            messages.one_line_error_handler("Please ensure that Min < Max", operations.position_correction(pos_num, 30, 40))
                        else:
                            num_col_name = values["-COLUMN-"] #----- column name to use as dictionary key-----#
                            if len(values["-NUMTITLE-"]) < 0 or len(values["-NUMTITLE-"]) > 21:
                                messages.two_line_error_handler("Please ensure the column title is", "between 1 and 20 characters long", operations.position_correction(pos_num, -50, 40))
                            else: #----- operations if all checks pass -----#
                                numbers = {num_col_name:[values["-NUMTITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                                if values["-DECIMALS-"] == 0:
                                    for row in range (rows):
                                        numbers[num_col_name].append(randint(int(values["-NUMMIN-"]), int(values["-NUMMAX-"])))
                                else:
                                    for row in range (rows):
                                        numbers[num_col_name].append(round(uniform(float(values["-NUMMIN-"]), float(values["-NUMMAX-"])), values["-DECIMALS-"]))
                                dict.update(numbers)
                                messages.operation_successful(f"Data added on column {num_col_name}", operations.position_correction(pos_num, 50, 40))
                except ValueError: #----- not a number in input fields, or float if 0 decimals selected -----#
                    messages.two_line_error_handler("Please ensure the input values are numbers", "Please ensure that if no decimals are selected, values are integers", operations.position_correction(pos_num, -130, 40))
    num_window.close()

#----- configure location generation and storage based on input -----#
def configure_location(win_pos, rows, sample, dict):
    street_name_1 = sample.data["STREET NAME 1"]
    street_name_2 = sample.data["STREET NAME 2"]
    city_and_state = sample.data["CITY AND STATE"]
    pos_loc = operations.position_correction(win_pos, -10, 80)
    layout_loc = [
        [sg.T("Street Name 1 Samples"), sg.Push(), sg.T("Street Name 2 Samples"), sg.T("City and State Samples")],
        [sg.Listbox(street_name_1, size= (15,5), key= "-SNAMES1-"), sg.Push(), sg.Listbox(street_name_2, size= (15,5), key= "-SNAMES2-"), sg.Push(), sg.Listbox(city_and_state, size= (15,5), key= "-CITYSTATE-")],
        [sg.HorizontalSeparator()],
        [sg.Checkbox("Add full address (incl. Street No. and Postal Code) on Column:", key= "-FULLLOC-", enable_events= True, default= True), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMN-", readonly= True)],
        [sg.T("With title:"), sg.Push(), sg.I(size= (40, 1), key= "-LOCTITLE-", default_text= "Location")],
        [sg.HorizontalSeparator()],
        [sg.Checkbox("Add Street, City and State in 3 seperate columns:", enable_events= True, key= "-SEPARATELOC-"), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-STREETCOLUMN-", readonly= True), sg.Combo(global_constants.EXCEL_COLUMN, key = "-CITYCOLUMN-", readonly= True), sg.Combo(global_constants.EXCEL_COLUMN, key = "-STATECOLUMN-", readonly= True)],
        [sg.HorizontalSeparator()],
        [sg.B("Preview"), sg.B("Sample Reload"), sg.Push(), sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
    ]
    loc_window = sg.Window("Location", layout_loc, font= global_constants.DEFAULT_FONT, modal= True, location= pos_loc, icon= "icon.ico")
    #----- data location loop -----#
    while True:
        event, values = loc_window.read()
        pos_loc = loc_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSED | "Back":
                break
            case "Sample Reload":
                sample.read_datatype("LOCATION", pos_loc)
                street_name_1 = sample.data["STREET NAME 1"]
                street_name_2 = sample.data["STREET NAME 2"]
                city_and_state = sample.data["CITY AND STATE"]
                loc_window.Element("-SNAMES1-").update(street_name_1)
                loc_window.Element("-SNAMES2-").update(street_name_2)
                loc_window.Element("-CITYSTATE-").update(city_and_state)
            case "Preview":
                operations.preview_dataframe(dict, rows, pos_loc)
            case "-FULLLOC-":
                if not values["-FULLLOC-"] and not values["-SEPARATELOC-"]:
                    loc_window.Element("-FULLLOC-").update(True)
                loc_window.Element("-SEPARATELOC-").update(False)
            case "-SEPARATELOC-":
                if not values["-SEPARATELOC-"] and not values["-FULLLOC-"]:
                    loc_window.Element("-SEPARATELOC-").update(True)
                loc_window.Element("-FULLLOC-").update(False)
            case "Add":
                if values["-FULLLOC-"]:
                    if values["-COLUMN-"] == "":
                        messages.one_line_error_handler("Please specify the target Column", operations.position_correction(pos_loc, 180, 80))
                    else:
                        loc_col_name = values["-COLUMN-"] #----- column name to use as dictionary key-----#
                        if len(values["-LOCTITLE-"]) < 0 or len(values["-LOCTITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column title is", "between 1 and 20 characters long", operations.position_correction(pos_loc, 150, 80))
                        else:
                            locations = {loc_col_name:[values["-LOCTITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                            for row in range(rows):
                                s_number = randint(1,200)
                                s_name_1 = choice(street_name_1)
                                s_name_2 = choice(street_name_2)
                                citystate = choice(city_and_state)
                                postal_code = randint(10000, 99999)
                                locations[loc_col_name].append(f"{s_number} {s_name_1} {s_name_2}, {citystate}, {postal_code}")
                            dict.update(locations)
                            messages.operation_successful(f"Data added on column {loc_col_name}", operations.position_correction(pos_loc, 200, 80))
                if values["-SEPARATELOC-"]:
                    if values["-STREETCOLUMN-"] == "" or values["-CITYCOLUMN-"] == "" or values["-STATECOLUMN-"] == "":
                        messages.one_line_error_handler("Please specify the target Columns", operations.position_correction(pos_loc, 180, 80))
                    elif values["-STREETCOLUMN-"] == values["-CITYCOLUMN-"] or values["-STREETCOLUMN-"] == values["-STATECOLUMN-"] or values["-STATECOLUMN-"] == values["-CITYCOLUMN-"]:
                        messages.one_line_error_handler("Please specify different Columns for selected data", operations.position_correction(pos_loc, 100, 100))
                    else:
                        street_col_name = values["-STREETCOLUMN-"] #----- column name to use as dictionary key-----#
                        city_col_name = values["-CITYCOLUMN-"] #----- column name to use as dictionary key-----#
                        state_col_name = values["-STATECOLUMN-"] #----- column name to use as dictionary key-----#
                        streets = {street_col_name:["Street"]} #----- column as key in dictionary and title as first element of associated list -----#
                        cities = {city_col_name:["City"]} #----- column as key in dictionary and title as first element of associated list -----#
                        states = {state_col_name:["State"]} #----- column as key in dictionary and title as first element of associated list -----#
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
                        messages.operation_successful(f"Data added on columns {street_col_name}, {city_col_name} and {state_col_name}", operations.position_correction(pos_loc, 150, 80))         
    loc_window.close()

