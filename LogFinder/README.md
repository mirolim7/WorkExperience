# LogFinder
This is the desktopp application for finding error sample numbers through the log files, which search whole file for a phrases "Проба с номером" and  "не найдена" and takes the value between them and returns at the screen. <br /><br />
For running LogFinder.py, it is important to ensure that all needed libraries were installed. <br /><br />
To make a desktop app, .exe file, run ```pyinstaller.exe --onefile --windowed --icon=icon.ico LogFinder.py``` to make as one .exe file. In case you want it to be as whole folder with needed packages and file just remove ```--onefile```. <br /><br />
