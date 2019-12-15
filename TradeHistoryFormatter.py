import os
import time
import datetime
import shutil
import PySimpleGUI as sg
from openpyxl import load_workbook
import ctypes

# setup window layout
layout = [[sg.Text('Load trade history file:')],
          [sg.Text('File:', justification='right'), sg.InputText(), sg.FileBrowse(), sg.Button('Load File')],
          [sg.Text('_' * 100, size=(70, 1))],
          [sg.Button('Update History'), sg.Button('Exit')],
          [sg.Text('Result:'), sg.Text("", key='Result')]]

# rendering window
window = sg.Window('Trade History Formatter', auto_size_text=True,
                   default_element_size=(40, 10)).Layout(layout)

# Main Program
while True:
    event, values = window.Read()
    if event == 'Load File':
        fileName = 'abc.xls'
        MessageBox = ctypes.windll.user32.MessageBoxW
        if not os.path.isfile(fileName):
            MessageBox(None, "File not exist!", "File Operation", 0)
    elif event == 'Update History':        
        window.Element('Result').Update('Success')
        pass 
    elif event is None or event == 'Exit':
        break
    print('event: ', event, '\nvalues:', values) # debug message
    event, values = window.Read()

"""     if event is None or event == 'Exit':
        break
    elif event == 'Verify':
        tempList = str(values[0]).split(",")        
        # count = Solution.heightChecker(tempList)
        window.Element('Result').Update(count)
    print(event, values, count) """

window.Close()
