from openpyxl import load_workbook
from TradeHistoryFormatter import excelProcessor

if __name__ == '__main__':
    excelProcessor('transactions_forTest.xlsx')