import json


class Lang:

    def __init__(self, data) -> None:

        self.lang = data['lang']
        # GUI
        self.gui_title = data['GUI']['title']
        self.gui_program_setting = data['GUI']['program_setting']
        self.gui_setting_loaded = data['GUI']['setting_loaded']
        self.gui_load_setting_file = data['GUI']['load_setting_file']
        self.gui_settings_file = data['GUI']['settings_file']
        self.gui_apply_settings = data['GUI']['apply_settings']
        self.gui_open_setting_editor = data['GUI']['open_setting_editor']
        self.gui_load_trade_history_file = data['GUI']['load_trade_history_file']
        self.gui_file = data['GUI']['file']
        self.gui_load_file = data['GUI']['load_file']  # unused
        self.gui_spreadsheet_files = data['GUI']['spreadsheet_files']
        self.gui_result = data['GUI']['result']
        self.gui_process_history = data['GUI']['process_history']
        self.gui_rem_last_xls_file = data['GUI']['rem_last_xls_file']
        self.gui_exp_error_log = data['GUI']['exp_error_log']
        self.gui_exit = data['GUI']['exit']
        self.gui_success = data['GUI']['success']
        self.gui_fail = data['GUI']['fail']
        self.gui_ver = data['GUI']['ver']

        # GUI-settings
        self.st_author = data['settings']['author']
        self.st_check_update = data['settings']['check_update']
        self.st_website = data['settings']['website']
        self.st_setting_window_title = data['settings']['setting_window_title']
        self.st_localization = data['settings']['localization']
        self.st_gui_theme = data['settings']['gui_theme']
        self.st_xls_fmt_setting = data['settings']['xls_fmt_setting']
        self.st_odd_col_color = data['settings']['odd_col_color']
        self.st_even_col_color = data['settings']['even_col_color']
        self.st_disp_date_fmt = data['settings']['disp_date_fmt']
        self.st_performance_title = data['settings']['performance_title']
        self.st_coloring_method_title = data['settings']['coloring_method_title']
        self.st_coloring_method_option = data['settings']['coloring_method_option']
        self.st_backup_settings = data['settings']['backup_settings']
        self.st_ok = data['settings']['ok']
        self.st_cancel = data['settings']['cancel']

        # msg box
        self.msg_box_file_op_title = data['msg_box']['file_op_title']
        self.msg_box_file_not_exist = data['msg_box']['file_not_exist']
        self.msg_box_file_processed = data['msg_box']['file_processed']
        self.msg_box_trading_sht_not_exist = data['msg_box']['trading_sht_not_exist']
        self.msg_box_select_file_first = data['msg_box']['select_file_first']
        self.msg_box_settings_file_not_change = data['msg_box']['settings_file_not_change']
        self.msg_box_color_fmt_wrong_title = data['msg_box']['color_fmt_wrong_title']
        self.msg_box_odd_col_color_fmt = data['msg_box']['odd_col_color_fmt']
        self.msg_box_even_col_color_fmt = data['msg_box']['even_col_color_fmt']
        self.msg_box_date_fmt_wrong_title = data['msg_box']['date_fmt_wrong_title']
        self.msg_box_date_col_fmt_wrong = data['msg_box']['date_col_fmt_wrong']
        self.msg_box_format_wrong = data['msg_box']['format_wrong']
        self.msg_box_chk_update_title = data['msg_box']['chk_update_title']
        self.msg_box_ver_up_to_date = data['msg_box']['ver_up_to_date']
        self.msg_box_need_update = data['msg_box']['need_update']
        self.msg_box_running_higher_ver_program = data['msg_box']['running_higher_ver_program']
        self.msg_box_running_unofficial_program = data['msg_box']['running_unofficial_program']
        self.msg_box_release_ver_not_exist = data['msg_box']['release_ver_not_exist']

        # Excel relate
        self.xls_sheet_names = data['excel']['sheet_names']  # sheet list
        # optional sheet list
        self.xls_opt_sheet_names = data['excel']['opt_sheet_names']

        self.xls_tt_date = data['excel']['table_title']['date']
        self.xls_tt_quantity = data['excel']['table_title']['quantity']
        self.xls_tt_symbol = data['excel']['table_title']['symbol']
        self.xls_tt_price = data['excel']['table_title']['price']
        self.xls_tt_commission = data['excel']['table_title']['commission']
        self.xls_tt_amount = data['excel']['table_title']['amount']
        self.xls_tt_ver = data['excel']['table_title']['ver']
        self.xls_tt_event = data['excel']['table_title']['event']
        self.xls_tt_msg = data['excel']['table_title']['message']
        self.xls_tt_remark = data['excel']['table_title']['remark']

        # remark message
        self.xls_msg_rebate = data['excel']['msg']['rebate']
        self.xls_msg_client_req_e_funding_dist = data['excel']['msg']['client_req_e_funding_dist']
        self.xls_msg_client_req_e_funding_rec = data['excel']['msg']['client_req_e_funding_rec']
        self.xls_msg_intra_account_transfer = data['excel']['msg']['intra_account_transfer']
        self.xls_msg_bold_indication_ITA = data['excel']['msg']['bold_indication_ITA']
        self.xls_msg_italic_mandi_exchange = data['excel']['msg']['italic_mandi_exchange']
        self.xls_msg_italic_qual_div = data['excel']['msg']['italic_qual_div']
        self.xls_msg_red_div_short = data['excel']['msg']['red_div_short']
        self.xls_msg_blue_nontax_div = data['excel']['msg']['blue_nontax_div']
        self.xls_msg_bold_cashInLieu = data['excel']['msg']['bold_cashInLieu']

        # excel log
        self.log_evt_description_keyword_missing = data[
            'xls_error_log']['event']['description_keyword_missing']
        self.log_evt_withholding_symbol_missing = data[
            'xls_error_log']['event']['withholding_symbol_missing']
        self.log_evt_transaction_symbol_missing = data[
            'xls_error_log']['event']['transaction_symbol_missing']
        self.log_evt_event_skip = data['xls_error_log']['event']['event_skip']

        self.log_msg_description_keyword_missing = data[
            'xls_error_log']['message']['description_keyword_missing']
        self.log_msg_withholding_symbol_missing = data[
            'xls_error_log']['message']['withholding_symbol_missing']
        self.log_msg_transaction_symbol_missing = data[
            'xls_error_log']['message']['transaction_symbol_missing']
        self.log_msg_event_skip = data['xls_error_log']['message']['event_skip']
        self.log_msg_found_error = data['xls_error_log']['message']['found_error']

        # excel status log
        self.xls_stat_sym_list = data['xls_stat_log']['sym_list']
        self.xls_stat_sym_qty = data['xls_stat_log']['sym_qty']


if __name__ == '__main__':
    # open json file and read
    with open('lang_enUS.json', 'r', encoding="utf-8") as reader:
        data = json.loads(reader.read())

    a = Lang(data)
    print(a.lang)
    print(a.gui_title)
    print(a.log_msg_description_keyword_missing)
    print(a.xls_sheet_names[3])
