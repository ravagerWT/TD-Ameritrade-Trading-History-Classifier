import json


class Lang:

    def __init__(self, data) -> None:

        self.lang = data['lang']

        self.gui_title = data['GUI']['title']
        self.gui_program_setting = data['GUI']['program_setting']
        self.gui_load_setting_file = data['GUI']['load_setting_file']
        self.gui_open_setting_editor = data['GUI']['open_setting_editor']
        self.gui_load_trade_history_file = data['GUI']['load_trade_history_file']
        self.gui_file = data['GUI']['file']
        self.gui_load_file = data['GUI']['load_file']
        self.gui_spreadsheet_files = data['GUI']['spreadsheet_files']
        self.gui_result = data['GUI']['result']
        self.gui_process_history = data['GUI']['process_history']
        self.gui_exit = data['GUI']['exit']
        self.gui_success = data['GUI']['success']
        self.gui_fail = data['GUI']['fail']
        self.gui_ver = data['GUI']['ver']

        self.msg_box_file_op_title = data['msg_box']['file_op_title']
        self.msg_box_file_not_exist = data['msg_box']['file_not_exist']
        self.msg_box_select_file_first = data['msg_box']['select_file_first']

        self.xls_sheet_names = data['excel']['sheet_names']  # list

        self.xls_tt_date = data['excel']['table_title']['date']
        self.xls_tt_quantity = data['excel']['table_title']['quantity']
        self.xls_tt_symbol = data['excel']['table_title']['symbol']
        self.xls_tt_price = data['excel']['table_title']['price']
        self.xls_tt_commission = data['excel']['table_title']['commission']
        self.xls_tt_amount = data['excel']['table_title']['amount']
        self.xls_tt_ver = data['excel']['table_title']['ver']
        self.xls_tt_event = data['excel']['table_title']['event']
        self.xls_tt_msg = data['excel']['table_title']['message']

        self.log_evt_description_keyword_missing = data[
            'xls_error_log']['event']['description_keyword_missing']
        self.log_evt_withholding_symbol_missing = data[
            'xls_error_log']['event']['withholding_symbol_missing']

        self.log_msg_description_keyword_missing = data[
            'xls_error_log']['message']['description_keyword_missing']
        self.log_msg_on = data['xls_error_log']['message']['on']
        self.log_msg_withholding_symbol_missing = data[
            'xls_error_log']['message']['withholding_symbol_missing']
        self.log_msg_th_row = data['xls_error_log']['message']['th_row']
        self.log_msg_found_error = data['xls_error_log']['message']['found_error']


if __name__ == '__main__':
    # open json file and read
    with open('lang_enUS.json', 'r', encoding="utf-8") as reader:
        data = json.loads(reader.read())
    
    a = Lang(data)
    print(a.lang)
    print(a.gui_title)
    print(a.log_msg_description_keyword_missing)
    print(a.xls_sheet_names[3])