import language
import json

if __name__ == '__main__':
    # open json file and read
    with open('lang_enUS.json', 'r', encoding="utf-8") as reader:
        data = json.loads(reader.read())
    
    a = language.Lang(data)
    print(a.lang)
    print(a.gui_title)
    print(a.log_msg_description_keyword_missing)
    print(a.xls_sheet_names[3])
