import json


class Thfset:

    def __init__(self, data) -> None:

        self.gen_set_lang = data['general']['set_lang']
        self.gen_ava_lang = data['general']['ava_lang']
        self.gen_ava_lang_code = data['general']['ava_lang_code']
        self.gen_last_transaction_file_path = data['general']['last_transaction_file_path']
        self.gen_setting_file_path = data['general']['setting_file_path']
        self.gen_exp_error_log = data['general']['exp_error_log']

        self.sht_trns_row_pos_startToWrite = data['sht_trns']['row_pos_startToWrite']
        self.sht_trns_col_pos_tr_date = data['sht_trns']['col_pos_tr_date']
        self.sht_trns_col_pos_tr_description = data['sht_trns']['col_pos_tr_description']
        self.sht_trns_col_pos_qty = data['sht_trns']['col_pos_qty']
        self.sht_trns_col_pos_symbol = data['sht_trns']['col_pos_symbol']
        self.sht_trns_col_pos_price = data['sht_trns']['col_pos_price']
        self.sht_trns_col_pos_fee = data['sht_trns']['col_pos_fee']
        self.sht_trns_col_pos_amount = data['sht_trns']['col_pos_amount']

        self.xls_fmt_color_for_odd_column = data['xls_fmt']['color_for_odd_column']
        self.xls_fmt_color_for_even_column = data['xls_fmt']['color_for_even_column']
        self.xls_fmt_display_date_format = data['xls_fmt']['display_date_format']


if __name__ == '__main__':
    # open json file and read
    with open('settings.json', 'r', encoding="utf-8") as reader:
        lang_reader = json.loads(reader.read())

    a = Thfset(lang_reader)
    print(a.gen_set_lang)
    print(a.gen_ava_lang[1])
    print(a.gen_ava_lang_code[1])
