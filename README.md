# DummPy
An application that produces fake (dummy) data for Data analysis practice.

Please take a look at the [User Guide](https://github.com/AntonisTorb/dummPy/blob/main/User%20Guide.pdf) before using the application, where you can find instructions on how to run and use the application. If you discover any bugs or have any suggestions, feel free to open an issue. Pull requests are welcome as well, but since I am a begginer (it took me 2 weeks and a lot of google searches to reach version 1.0.0), I might not fully understand your code, so please be patient :) 

## For the future:
- Add Date as a configurable data type.
- More proofing/testing and optimization.

## Recent changes (v1.0.1):
- Major refactoring of code.</br>
  - Reduced function length across the board by refactoring the code into seperate functions.
  - Reduced main file length by creating separate files to be imported.
- Improved error handling.
- Improved loading of sample data.</br>
	- Sample data are now loaded on startup.
- Added a way to check if the existing dictionary has been last saved or loaded, in order to improve exit and reset operations.</br>
	- Now if the dictionary has been last saved or loaded, exit/reset operations are instant.
- Removed variables that were only used once in the saving and loading operations.
- Replaced "if" statements with "match" statements in all event handling instances.
- Improved file opening with the proper syntax so closing is no longer necessary after opening.
- Remaning of variables to be more uniform across the code.
- Further commented the code for context.