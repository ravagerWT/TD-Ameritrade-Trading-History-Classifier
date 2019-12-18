import os
import time
import datetime
import shutil
import PySimpleGUI as sg
from openpyxl import load_workbook
import ctypes

# openpyxl.utils.cell.coordinate_from_string('B3') // ('B', 3)
# openpyxl.utils.cell.column_index_from_string(a[0])  // 2
# openpyxl.utils.cell.coordinate_to_tuple('B3')  // (3, 2)

#// TODO:實作介面語言字串全域變數

# load excel file to be processed
def loadExcelFile(filePath):
    fileName = os.path.basename(filePath)
    os.chdir(os.path.dirname(filePath))
    MessageBox = ctypes.windll.user32.MessageBoxW
    if not os.path.isfile(fileName):
        # sg.popup("File not exist!")
        MessageBox(None, "File not exist!", "File Operation", 0)
    else:        
        return fileName

# load program setting from setting.json
def loadSetting():
    pass

# edit program setting in setting.json
def editSetting():
    pass

# save setting to setting.json
def saveSetting():
    pass

# load HMI langage from lang.json
def loadLang(langCode):
    langString = 0
    return langString

# setup window layout
def setWindow(langString):
    #// TODO: 實作載入目標語言字串
    # setup window layout
    layout = [[sg.Text('Program Setting' + ':')],
              [sg.FileBrowse('Load Setting File'),
               sg.Button('Open Setting Editor')],
              [sg.Text('_' * 100, size=(70, 1))],
              [sg.Text('Load trade history file' + ':')],
              [sg.Text('File' + ':', justification='right'),
               sg.InputText(), sg.FileBrowse(), sg.Button('Load File')], [sg.Text('Result' + ':'), sg.Text('', key='loadingResult')],
              [sg.Text('_' * 100, size=(70, 1))],
              [sg.Button('Update History'), sg.Button('Exit')],
              [sg.Text('Result' + ':'), sg.Text('', key='Result')]]
    # rendering window
    window = sg.Window('Trade History Formatter', auto_size_text=True,
                       default_element_size=(40, 10)).Layout(layout)
    return window

# Main Program
def main(window):
    fileName = None
    MessageBox = ctypes.windll.user32.MessageBoxW
    while True:
        event, values = window.Read()
        if event == 'Load File':
            if values['Browse'] is None or values['Browse'] == '':                
                MessageBox(None, "Please select file first!", "File Operation", 0)
            else:
                fileName = loadExcelFile(values['Browse'])
                window.Element('loadingResult').Update('Success')  #showing loading result
        elif event == 'Load Setting File':
            #// TODO:實作載入設定檔
            pass
        elif event == 'Open Setting Editor':
            #// TODO:實作設定檔編輯功能
            pass
        elif event == 'Update History':
            if fileName is None or fileName == '':                
                MessageBox(None, "Please select file first!", "File Operation", 0)
            else:
                wb = load_workbook(fileName)
                #// TODO: 實作excel檔處理
                window.Element('Result').Update('Success')  #showing process result
        elif event is None or event == 'Exit':
            break
        print('event: ', event, '\nvalues:', values)  # debug message
        event, values = window.Read()

    window.Close()

if __name__ == '__main__':
    lang = loadLang('default')
    window = setWindow(lang)
    main(window)
