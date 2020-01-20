# TD Ameritrade Trading History Classifier

- Latest Stable Version: None
- Latest Beta Version: 2.0.0

## Changelog

* Beta 2.0.1 (2020-01-XX)
    * NEW: introduce pipenv
    * UPDATE: move .vscode/settings.json out of version control

* Beta 2.0.0 (2020-01-19)
    * NEW: implement a new option allow users to decide the coloring method in spreading sheets
    * NEW: implement date format checking mechanism
    * NEW: project dependency document: requirements.txt
    * NEW: add screenshots for main and setting window
    * FIX: some typo

* Beta 1.3.2 (2020-01-15)
    * NEW: adding the instruction for program localization to README.md
    * FIX: using a new approach to fix the incorrect behavior in chk_update() method
    * FIX: some typo
    * FIX: incorrect excel error log export behavior in excelProcessor() method
    * FIX: progress bar incorrect behavior
    * UPDATE: some sentences for better understanding

* Beta 1.3.1 (2020-01-13)
    * FIX: Not functional export excel process log checkbox
    * UPDATE: extract check update function to chk_update()

* Beta 1.3 (2020-01-10)
    * NEW: implement progress bar to show current status
    * NEW: allow user to select the theme of program GUI
    * NEW: implement author and official website information in setting window
    * NEW: implement check update function in setting window
    * NEW: add screenshots for main and setting window

* Beta 1.2 (2020-01-08)
    * NEW: implement more trading types sorting function (thanks threeSecGun from PTT)
    * NEW: moving the note for transactions type to new remark sheet
    * NEW: implement file status check function in excelProcessor()
    * NEW: implement handing the error code function after click process history button
    * NEW: localization for file status sheet
    * NEW: optimize processing the custom symbol input function
    * NEW: move setting cells color code to a better understanding position
    * UPDATE: change the name of some variables for better code understanding

* Beta 1.1 (2020-01-06)
    * NEW: implement more trading types sorting function (thanks threeSecGun from PTT)
    * NEW: skip not important record feature
    * NEW: implement adding new trading history to existing file feature
    * FIX: error message not showing correctly on log sheet
    * FIX: when save file the file version won't handle correctly
    * FIX: setting file loading function
    * UPDATE: GUI for easy understanding
    * UPDATE: localization for trading type indication

* Beta 1.0 (2020-01-04)
    * First release