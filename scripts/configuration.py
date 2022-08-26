from random import choice, randint, uniform  # core python module
from datetime import date, timedelta  # core python module
import PySimpleGUI as sg  # pip install pysimplegui
import scripts.operations as operations
import scripts.messages as messages
import scripts.global_constants as global_constants


#----- configure name and email address generation and storage based on input -----#
def configure_name_email(win_pos, rows, sample, dict):
    first_names = sample.data["FIRST NAME"]
    last_names = sample.data["LAST NAME"]
    domains = sample.data["EMAIL DOMAIN"]
    pos_name = operations.position_correction(win_pos, -25, 80)

    def name_generation(column, title, columns_string, *, fnames= [], lnames= []):
        names_dict = {column:[title]} #----- column as key in dictionary and title as first element of associated list -----#
        for row in range(rows):
            if len(fnames) != 0:
                random_f_name = choice(fnames)
            else:
                random_f_name = ""
            if len(lnames) != 0:
                random_l_name = choice(lnames)
            else:
                random_l_name = ""
            random_name = f"{random_f_name} {random_l_name}"
            names_dict[column].append(random_name)
        if len(columns_string) == 0: #----- if no column for first name -----#
            columns_string = column
        else: #----- else append it to the existing one -----#
            columns_string = f"{columns_string}, {column}"
        return names_dict, columns_string

    #----- generation of the e-mail addresses -----#
    def email_generation(domains, column_email, title, column_name, *, seperate, column_name_2= None):
        emails_dict = {column_email:[title]}
        for row in range(1, rows + 1):
            domain = choice(domains)
            delta = randint(0,99)
            name1 = dict[column_name][row]
            name2 = ""
            if seperate: #----- if adding seperate first name and last name, we need to include both for the emial generation -----#
                name2 = f"{dict[column_name_2][row]}"
            random_email = f"{name1}{name2}{delta}@{domain}".replace(" ", "").lower() #----- replacing blanks if we added full name, and lowercase -----#
            emails_dict[column_email].append(random_email)
        return emails_dict

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
    name_window = sg.Window("Name and E-mail", layout_name, font= global_constants.DEFAULT_FONT, enable_close_attempted_event= True, modal= True, location= pos_name, icon= "icon.ico")

    while True: #----- DATA NAME LOOP -----#
        event, values = name_window.read()
        pos_name = name_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "Back":
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
                    columns_of_names = "" #----- string containing columns where name data were added, to compose confirmation message-----#
                    name_added = False #----- using this to avoid adding e-mail if error in adding name (column title length error for example) -----#
                    title_error = False
                    if values["-FULLNAME-"]: # --------- if adding full name -----#
                        if len(values["-NAMETITLE-"]) < 1 or len(values["-NAMETITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                            title_error = True
                        elif not title_error: #----- operations if all checks pass -----#
                            names, columns_of_names = name_generation(values["-COLUMNNAME-"], values["-NAMETITLE-"], columns_of_names, fnames= first_names, lnames= last_names)
                            dict.update(names)
                    if values["-FIRSTNAME-"]: #----- if adding first name -----#
                        if len(values["-FIRSTNAMETITLE-"]) < 1 or len(values["-FIRSTNAMETITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                            title_error = True
                        elif not title_error: #----- operations if all checks pass -----#
                            fnames, columns_of_names = name_generation(values["-COLUMNFIRSTNAME-"], values["-FIRSTNAMETITLE-"], columns_of_names, fnames= first_names)
                            dict.update(fnames)
                    if values["-LASTNAME-"]: #----- if adding last name -----#
                        if len(values["-LASTNAMETITLE-"]) < 1 or len(values["-LASTNAMETITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                            title_error = True
                        elif not title_error: #----- operations if all checks pass -----#
                            lnames, columns_of_names = name_generation(values["-COLUMNLASTNAME-"], values["-LASTNAMETITLE-"], columns_of_names, lnames= last_names)
                            dict.update(lnames)
                    if values ["-EMAIL-"] and not title_error: # -------- if adding e-mail as well -----#
                        mail_col_name = values["-COLUMNMAIL-"] #----- column name to use as dictionary key-----#
                        if len(values["-EMAILTITLE-"]) < 1 or len(values["-EMAILTITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column titles are", "between 1 and 20 characters long", operations.position_correction(pos_name, 200, 100))
                        elif not title_error:
                            if values["-FULLNAME-"]: #----- full name for email -----#
                                generated_email_addresses = email_generation(domains, values["-COLUMNMAIL-"], values["-EMAILTITLE-"], values["-COLUMNNAME-"], seperate= False)
                            elif values["-FIRSTNAME-"] and not values["-LASTNAME-"]: #----- if last name not added, only use first name for email -----#
                                generated_email_addresses = email_generation(domains, values["-COLUMNMAIL-"], values["-EMAILTITLE-"], values["-COLUMNFIRSTNAME-"], seperate= False)
                            elif values["-LASTNAME-"] and not values["-FIRSTNAME-"]: #----- if first name not added, only use last name for email -----#
                                generated_email_addresses = email_generation(domains, values["-COLUMNMAIL-"], values["-EMAILTITLE-"], values["-COLUMNLASTNAME-"], seperate= False)
                            elif values["-FIRSTNAME-"] and values["-LASTNAME-"]: #----- both first and last name for email -----#
                                generated_email_addresses = email_generation(domains, values["-COLUMNMAIL-"], values["-EMAILTITLE-"], values["-COLUMNFIRSTNAME-"], seperate= True, column_name_2= values["-COLUMNLASTNAME-"])
                            dict.update(generated_email_addresses) #----- add the emails to dictionary -----#
                            messages.operation_successful(f"Data added on columns {columns_of_names} and {mail_col_name}", operations.position_correction(pos_name, 200, 100))
                    elif not title_error: #----- when not adding email, show confirmation for the name data added -----#
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
    num_window = sg.Window("Number", layout_num, font= global_constants.DEFAULT_FONT, enable_close_attempted_event= True, modal= True, location= pos_num, icon= "icon.ico")
    #----- data number loop -----#
    while True:
        event, values = num_window.read()
        pos_num = num_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "Back":
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
                            col_num = values["-COLUMN-"] #----- column name to use as dictionary key-----#
                            if len(values["-NUMTITLE-"]) < 1 or len(values["-NUMTITLE-"]) > 21:
                                messages.two_line_error_handler("Please ensure the column title is", "between 1 and 20 characters long", operations.position_correction(pos_num, -50, 40))
                            else: #----- operations if all checks pass -----#
                                numbers = {col_num:[values["-NUMTITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                                if values["-DECIMALS-"] == 0:
                                    for row in range (rows):
                                        numbers[col_num].append(randint(int(values["-NUMMIN-"]), int(values["-NUMMAX-"])))
                                else:
                                    for row in range (rows):
                                        numbers[col_num].append(round(uniform(float(values["-NUMMIN-"]), float(values["-NUMMAX-"])), values["-DECIMALS-"]))
                                dict.update(numbers)
                                messages.operation_successful(f"Data added on column {col_num}", operations.position_correction(pos_num, 50, 40))
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
    loc_window = sg.Window("Location", layout_loc, font= global_constants.DEFAULT_FONT, enable_close_attempted_event= True, modal= True, location= pos_loc, icon= "icon.ico")
    #----- data location loop -----#
    while True:
        event, values = loc_window.read()
        pos_loc = loc_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "Back":
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
                        loc_col = values["-COLUMN-"] #----- column name to use as dictionary key-----#
                        if len(values["-LOCTITLE-"]) < 1 or len(values["-LOCTITLE-"]) > 21:
                            messages.two_line_error_handler("Please ensure the column title is", "between 1 and 20 characters long", operations.position_correction(pos_loc, 150, 80))
                        else:
                            locations = {loc_col:[values["-LOCTITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                            for row in range(rows):
                                s_number = randint(1,200)
                                s_name_1 = choice(street_name_1)
                                s_name_2 = choice(street_name_2)
                                citystate = choice(city_and_state)
                                postal_code = randint(10000, 99999)
                                locations[loc_col].append(f"{s_number} {s_name_1} {s_name_2}, {citystate}, {postal_code}")
                            dict.update(locations)
                            messages.operation_successful(f"Data added on column {loc_col}", operations.position_correction(pos_loc, 200, 80))
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

#----- configure date generation and storage based on input -----#
def configure_date(win_pos, rows, dict):
    YEAR_RANGE = [year for year in range(1900, 2101)]
    MONTH_RANGE = [month for month in range (1, 13)]
    MONTH_FORMATS = ("MM", "Month")
    YEAR_FORMATS = ("YYYY", "YY")
    DATE_FORMATS_MONTH_INT = ("Day/MM/Year", "Year/MM/Day", "Day-MM-Year", "Year-MM-Day")
    DATE_FORMATS_MONTH_STR = ("Day Month Year", "Month Day Year")
    
    #----- when year is configured in date data type, update the day element if there are transitions between leap and non leap years -----#
    def update_day_for_leap_year_transitions(day_element, val, win, days):
        RANGE = [day_element for day_element in range (1, days + 1)]
        try:
            selected_day = int(val[day_element])
        except ValueError: #----- if the day has not been selected yet -----#
            selected_day = ""
        win.Element(day_element).update(values= RANGE)
        try:
            if selected_day <= days: #----- if the existing value is less than the maximum of the new range, maintain it -----#
                win.Element(day_element).update(selected_day)
        except TypeError: #----- cannot perform comparisons between string and int, if selected_day from value error exception above -----#
            win.Element(day_element).update("")     

    #----- configuration of the year in date data type -----#
    def configure_year(year_element, month_element, day_element, val, win, year_set):
        if val[month_element] == "": #----- initial setting -----#
            win.Element(month_element).update(disabled= False)
        if val[month_element] == 2: #----- if the year_element changes while the month_element is set to February, we might have to update the days range due to leap years -----#
            if val[year_element] % 4 == 0 and not year_set % 4 == 0: #----- changing from non leap year_element to leap year_element -----#
                update_day_for_leap_year_transitions(day_element, val, win, days= 29)
            if not val[year_element] % 4 == 0 and year_set % 4 == 0: #----- changing from leap year_element to non leap year_element -----#
                update_day_for_leap_year_transitions(day_element, val, win, days= 28)
        return val[year_element]

    #----- configuration of the month in date data type -----#
    def configure_month(year_element, month_element, day_element, val, win):
        if val[month_element] in (1,3,5,7,8,10,12): #----- set to 31 day month_element -----#
            days = 31
        elif val[month_element] in (4,6,9,11): #----- set to 30 day month_element -----#
            days = 30
        elif val[month_element] == 2: #----- for February set to 29 or 28 days based on year_element -----#
            leap_year = val[year_element] % 4 == 0
            if leap_year:
                days = 29
            else:
                days = 28
        DAY_RANGE = [day for day in range(1, days + 1)]
        try:
            selected_day = int(val[day_element])
        except ValueError:
            selected_day = 32
        win.Element(day_element).update(values= DAY_RANGE, disabled= False)
        if selected_day <= days:
            win.Element(day_element).update(selected_day)

    #----- matching user input format to code format -----#
    def determine_format(user_input):
        match user_input:
            case "MM":
                return "%m"
            case "Month":
                return "%B"
            case "YYYY":
                return "%Y"
            case "YY":
                return "%y"
            case "Day/MM/Year":
                return "{day_format}/{month_format}/{year_format}"
            case "Year/MM/Day":
                return "{year_format}/{month_format}/{day_format}"
            case "Day-MM-Year":
                return "{day_format}-{month_format}-{year_format}"
            case "Year-MM-Day":
                return "{year_format}-{month_format}-{day_format}"
            case "Day Month Year":
                return "{day_format} {month_format} {year_format}"
            case "Month Day Year":
                return "{month_format} {day_format} {year_format}"
    
    layout_date = [
        [sg.T("", size = 10), sg.T("YYYY", size = 5, justification= "center"), sg.T(" /"), sg.T("MM", size = 3, justification= "center"), sg.T(" /"), sg.T("DD", size = 3, justification= "center"), sg.Push()],
        [sg.T("Start:", size = 10), sg.Combo(YEAR_RANGE, size= 5, key= "-MINYEAR-", enable_events= True, readonly= True), sg.Combo(MONTH_RANGE, size= 4, key= "-MINMONTH-", enable_events= True, disabled= True, readonly= True), sg.Combo(values= [""], size= 4, key= "-MINDAY-", disabled= True, readonly= True), sg.Push()],
        [sg.T("Finish:", size = 10), sg.Combo(YEAR_RANGE, size= 5, key= "-MAXYEAR-", enable_events= True, readonly= True), sg.Combo(MONTH_RANGE, size= 4, key= "-MAXMONTH-", enable_events= True, disabled= True, readonly= True), sg.Combo(values= [""], size= 4, key= "-MAXDAY-", disabled= True, readonly= True), sg.Push()],
        [sg.HorizontalSeparator()],
        [sg.T("Add to Column:"), sg.Push(), sg.Combo(global_constants.EXCEL_COLUMN, key = "-COLUMN-", readonly= True)],
        [sg.T("with title:"), sg.Push(), sg.I(size= (20, 1), key= "-DATETITLE-", default_text= "Date")],
        [sg.HorizontalSeparator()],
        [sg.T("Set the formats for:")],
        [sg.T("Month:"), sg.Push(), sg.Combo(MONTH_FORMATS, size = 5, key= "-MONTHFORMAT-", readonly= True, default_value= MONTH_FORMATS[0], enable_events= True)],
        [sg.T("Year:"), sg.Push(),sg.Combo(YEAR_FORMATS, size = 5, key= "-YEARFORMAT-", readonly= True, default_value= YEAR_FORMATS[0])],
        [sg.T("Date:"), sg.Push(),sg.Combo(DATE_FORMATS_MONTH_INT, size = 13, key= "-DATEFORMAT-", readonly= True, default_value= DATE_FORMATS_MONTH_INT[0])],
        [sg.HorizontalSeparator()],
        [sg.B("Clear"), sg.B("Preview"), sg.Push(),sg.B("Add", button_color= ("#292e2a", "#5ebd78")), sg.B("Back", button_color= ("#ffffff", "#bf365f"))]
    ]
    pos_date = operations.position_correction(win_pos, 140, 80)
    date_window = sg.Window("Date", layout_date, font= global_constants.DEFAULT_FONT, enable_close_attempted_event= True, modal= True, location= pos_date, icon= "icon.ico")
    min_year_set = 0 #----- used for leap year transition calculations -----#
    max_year_set = 0 #----- used for leap year transition calculations -----#
    #----- data date loop -----#
    while True:
        event, values = date_window.read()
        pos_date = date_window.current_location()
        match event: #----- actions to perform based on event -----#
            case sg.WINDOW_CLOSE_ATTEMPTED_EVENT | "Back":
                break
            case "Clear": #----- restart with all default values in UI -----#
                date_window.close()
                configure_date(operations.position_correction(pos_date, -140, -80), rows, dict)
                break
            case "Preview":
                operations.preview_dataframe(dict, rows, pos_date)
            case "-MINYEAR-":
                min_year_set = configure_year("-MINYEAR-", "-MINMONTH-", "-MINDAY-", values, date_window, min_year_set)
            case "-MAXYEAR-":
                max_year_set = configure_year("-MAXYEAR-", "-MAXMONTH-", "-MAXDAY-", values, date_window, max_year_set)
            case "-MINMONTH-":
                configure_month("-MINYEAR-", "-MINMONTH-", "-MINDAY-", values, date_window)
            case "-MAXMONTH-":
                configure_month("-MAXYEAR-", "-MAXMONTH-", "-MAXDAY-", values, date_window)
            case "-MONTHFORMAT-":
                if values["-MONTHFORMAT-"] == "Month":
                    date_window.Element("-DATEFORMAT-").update(DATE_FORMATS_MONTH_STR[0], values= DATE_FORMATS_MONTH_STR)
                elif values["-MONTHFORMAT-"] == "MM":
                    date_window.Element("-DATEFORMAT-").update(DATE_FORMATS_MONTH_INT[0], values= DATE_FORMATS_MONTH_INT)
            case "Add":
                if values["-COLUMN-"] == "":
                    messages.one_line_error_handler("Please specify the target Column", operations.position_correction(pos_date, 25, 80))
                else:
                    try:
                        min_year = values["-MINYEAR-"]
                        min_month = values["-MINMONTH-"]
                        min_day = values["-MINDAY-"]
                        max_year = values["-MAXYEAR-"]
                        max_month = values["-MAXMONTH-"]
                        max_day = values["-MAXDAY-"]
                        min_date = date(min_year, min_month, min_day)
                        max_date = date(max_year, max_month, max_day)
                        if min_date > max_date:
                            messages.two_line_error_handler("Please ensure that the start date", "is before the finish date", operations.position_correction(pos_date, 30, 80))
                        else: #----- logic for adding the data -----#
                            col_date = values["-COLUMN-"] #----- column name to use as dictionary key-----#
                            if len(values["-DATETITLE-"]) < 1 or len(values["-DATETITLE-"]) > 21:
                                messages.two_line_error_handler("Please ensure the column title is", "between 1 and 20 characters long", operations.position_correction(pos_date, 25, 80))
                            else: #----- operations if all checks pass -----#
                                dates = {col_date:[values["-DATETITLE-"]]} #----- column as key in dictionary and title as first element of associated list -----#
                                delta_date = max_date - min_date
                                month_format_input = values["-MONTHFORMAT-"]
                                year_format_input = values["-YEARFORMAT-"]
                                date_format_input = values["-DATEFORMAT-"]
                                day_format = "%d"
                                month_format = determine_format(month_format_input)
                                year_format = determine_format(year_format_input)
                                date_format = determine_format(date_format_input).format(day_format= day_format, month_format= month_format, year_format= year_format) #eval(determine_format(date_format_input)) #----- evaluate the f-string -----#
                                for row in range(rows):
                                    random_day = randint(0, delta_date.days)
                                    random_date = date.strftime(min_date + timedelta(days= random_day), date_format)
                                    dates[col_date].append(random_date)
                                dict.update(dates)
                                messages.operation_successful(f"Data added on column {col_date}", operations.position_correction(pos_date, 70, 80))
                    except TypeError:
                        messages.one_line_error_handler("Please ensure that all the date values have been populated.", operations.position_correction(pos_date, -80, 80))
    date_window.close()