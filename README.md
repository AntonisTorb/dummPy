# DummPy
An application that produces fake (dummy) data for Data analysis practice.

Please take a look at the [User Guide](https://github.com/AntonisTorb/dummPy/blob/main/User%20Guide.pdf) before using the application, where you can find instructions on how to run and use the application. If you discover any bugs or have any suggestions, feel free to open an issue. Pull requests are welcome as well, but since I am a begginer (it took me 2 weeks and a lot of google searches to reach version 1.0.0), I might not fully understand your code, so please be patient :) 

## For the future:
- Add Random string as a configurable data type.
- More proofing/testing and optimization.

## Recent changes (v1.1.0):

- New data type: Date
	- Generate random dates between two given start and finish dates.
	- See the User Guide for full information.
- Fixed a bug in "main.py", where the "save_dict" fuction was called from the wrong module.
- Removed the "ast" dependancy from "input_from_external.py", replaced "ast.literal_eval" with "eval".
- Updated all functions that created a PySimpleGUI window to generate a "CLOSE_ATTEMPTED_EVENT" when the window is closed via the "X" button.
	- This is to avoid the unnecessary warning due to the "current_location" PySimpleGUI method.
- Further refactored the functions in "configuration.py".