import os
import time
import datetime
import shutil
import PySimpleGUI as sg
from openpyxl import load_workbook
# import openpyxl.worksheet
import ctypes

# openpyxl.utils.cell.coordinate_from_string('B3')  // ('B', 3)
# openpyxl.utils.cell.column_index_from_string(a[0])  // 2
# openpyxl.utils.cell.coordinate_to_tuple('B3')  // (3, 2)

#// TODO:實作介面語言字串全域變數
sg.change_look_and_feel('Dark Blue 3')  # windows colorful

# load program setting from setting.json
def loadSetting():
    pass

# edit program setting in setting.json
def editSetting():
    pass

# save setting to setting.json
def saveSetting():
    pass

# load GUI langage from lang.json
def loadLang(langCode):
    langString = 0
    return langString

# setup window layout
def setWindow(langString):
    #// TODO: 實作載入目標語言字串
    #// TODO: 實作版本號介面
    #// TODO: 簡化或整合檔案載入及處理GUI介面
    # setup window layout
    layout = [[sg.Text('Program Setting' + ':')],
              [sg.FileBrowse('Load Setting File'), sg.Button('Open Setting Editor')],
              [sg.Text('_' * 100, size=(70, 1))],
              [sg.Text('Load trade history file' + ':')],
              [sg.Text('File' + ':', justification='right'), sg.InputText('', key = 'it_filePath'), sg.FileBrowse(file_types=(("Spreadsheet Files", "*.xls"),("Spreadsheet Files", "*.xlsx"),)), sg.Button('Load File')],
              [sg.Text('Result' + ':'), sg.Text('', size=(20, 1), key='loadingResult')],
              [sg.Text('_' * 100, size=(70, 1))],
              [sg.Button('Update History'), sg.Button('Exit')],
              [sg.Text('Result' + ':'), sg.Text('', size=(20,1), key='Result')]]
    # rendering window
    window = sg.Window('Trade History Formatter', auto_size_text=True, default_element_size=(40, 10), resizable=False).Layout(layout)
    return window

# load excel file to be processed
def loadExcelFile(filePath):
    fileName = os.path.basename(filePath)
    os.chdir(os.path.dirname(filePath))
    MessageBox = ctypes.windll.user32.MessageBoxW
    if not os.path.isfile(fileName):
        # sg.popup("File not exist!") # build-in pipup window
        MessageBox(None, "File not exist!", "File Operation", 0)
    else:        
        return fileName

# excel processor
def excelProcessor(fileName, symbol_list = []):
    sheet_list = ['Sorted trade history','ORDINARY DIVIDEND','W-8 WITHHOLDING','WIRING INFO','Ver']    
    # loading workbook
    wb = load_workbook(fileName)
    # create sheets
    for i in range(5):
        if not sheet_list[i] in wb.sheetnames:
            wb.create_sheet(sheet_list[i])
    
    ws_tran = wb["transactions"]
    ws_STH = wb["Sorted trade history"]
    ws_OD = wb["ORDINARY DIVIDEND"]
    ws_W8 = wb["W-8 WITHHOLDING"]
    ws_WI = wb["WIRING INFO"]
    ws_ver = wb["Ver"]

    # setting layout
    ws_STH['A2'] = 'DATE' # Sorted trade history
    ws_OD['A1'] = 'DATE' # ORDINARY DIVIDEND
    ws_W8['A1'] = 'DATE' # W-8 WITHHOLDING
    ws_WI['A1'] = 'DATE' # WIRING INFO
    ws_WI['B1'] = 'Amount'
    ws_ver['A1'] = 'DATE' # Ver
    ws_ver['B1'] = 'Ver'

    counter = 0
    temp_tr_date = ''
    # gather all stock symbol from 'E5'
    for i in range(2, ws_tran.max_row):
        tr_date = ws_tran.cell(i, 1)
        tr_description = ws_tran.cell(i, 3)
        tr_quan = ws_tran.cell(i, 4)
        tr_symbol = ws_tran.cell(i, 5)
        tr_price = ws_tran.cell(i, 6)
        tr_fee = ws_tran.cell(i, 7)
        tr_amount = ws_tran.cell(i, 8)
        # processing sheets format by stock symbols
        if not tr_symbol.value in symbol_list and not tr_symbol.value == None:
            counter += 1
            symbol_list.append(ws_tran.cell(i, 5).value)
            ws_STH.cell(1, counter*4-2).value = ws_tran.cell(i,
                                                             5).value  # stock symbol
            ws_STH.cell(2, counter*4-2).value = 'Quantity'
            ws_STH.cell(2, counter*4-1).value = 'Price'
            ws_STH.cell(2, counter*4).value = 'Fee'
            ws_STH.cell(2, counter*4+1).value = 'Amount'
            # // TODO:待實作儲存格合併功能及顏色區別
            ws_OD.cell(1, counter+1).value = ws_tran.cell(i, 5).value
            ws_W8.cell(1, counter+1).value = ws_tran.cell(i, 5).value
            # print(ws_tran.cell(i, 5).value)  # debug message

        # sorting trade history
        if 'WIRE INCOMING' in tr_description.value:
            ws_WI.insert_rows(2)  # add new row
            ws_WI.cell(2, 1).value = tr_date.value  # date
            ws_WI.cell(2, 2).value = tr_amount.value  # amount
        elif 'Bought' in tr_description.value:
            symbol_index = symbol_list.index(tr_symbol.value)
            if not tr_date.value == temp_tr_date:
                ws_STH.insert_rows(3)  # add new row
                temp_tr_date = tr_date.value
            ws_STH.cell(3, symbol_index*4+2).value = tr_quan.value
            ws_STH.cell(3, symbol_index*4+3).value = tr_price.value
            ws_STH.cell(3, symbol_index*4+4).value = tr_fee.value
            ws_STH.cell(3, symbol_index*4+5).value = tr_amount.value

        #// TODO: 待確認關鍵字
        elif 'Sold' in tr_description.value:
            ws_STH.insert_rows(3)  # add new row
            pass
        elif 'ORDINARY DIVIDEND' in tr_description.value:
            pass
        elif 'WITHHOLDING' in tr_description.value:
            pass
        #// TODO: 待確認以下關鍵字：出金
        else:
            #// TODO: 實作輸出log檔功能
            print('not in the known keyword: ' + ws_tran.cell(i, 3).value)

    wb.save('transactions_forTest_r1.xlsx')
    
# Main Program
def main(window):
    fileName = None
    MessageBox = ctypes.windll.user32.MessageBoxW
    while True:
        event, values = window.Read()
        if event == 'Load File':
            if values['it_filePath'] is None or values['it_filePath'] == '':                
                MessageBox(None, "Please select file first!", "File Operation", 0)
            else:
                fileName = loadExcelFile(values['it_filePath'])
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
                excelProcessor(fileName)                
                window.Element('Result').Update('Success')  #showing process result
        elif event is None or event == 'Exit':
            break
        print('event: ', event, '\nvalues:', values)  # debug message
        event, values = window.Read()

    window.Close()

if __name__ == '__main__':
    loadSetting()
    lang = loadLang('default')
    window = setWindow(lang)
    main(window)
