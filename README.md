# TD Ameritrade Trading History Classifier
![Release](https://img.shields.io/github/v/release/ravagerWT/TD-Ameritrade-Trading-History-Classifier)
![Release Date](https://img.shields.io/github/release-date/ravagerWT/TD-Ameritrade-Trading-History-Classifier)
![Commits Since](https://img.shields.io/github/commits-since/ravagerWT/TD-Ameritrade-Trading-History-Classifier/latest/develop)
![License](https://img.shields.io/github/license/ravagerWT/TD-Ameritrade-Trading-History-Classifier)

![](screenshot/Main%20GUI%20Beta2.0.0.jpg?raw=true)

This python program help user easily classify the transactions history by symbol and account activity into several spreading sheets.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine.

### Prerequisites & Dependency

```
Python
PySimpleGUI (package)
openpyxl (package)
```

### Installing

**Step1. Install python:**

Follow the [Python Official Install instruction](https://docs.python.org/3/using/windows.html) to install it.

Or follow the quick step below:
1. [Download](https://www.python.org/downloads/) python from official website.
2. Double click the installer to install python (recommended) or extract the portable version python to your desired folder.
3. If you haven't enable the `setting environment variables option` in the installer or you're using portable version, please follow the instruction to [setting environment variables](https://docs.python.org/3/using/windows.html#configuring-python).


**Step2. Install python packages:**

Windows:

1. Open CMD or Powershell
2. Type following command and hit enter to install packages
   - `pip install PySimpleGUI`
   - `pip install openpyxl` 

**Step3. Get the program:**

Clone this repo to your local machine using
```git
git clone https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier
```
Or

[Download](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier/releases/latest) project source code file and extract it.

## How to use

1. Get transaction file:
   - Download your transaction file from TD Ameritrade website.  The transaction file downloaded from TD Ameritrade website should be .csv file.
   - Use MicroSoft Excel or any compatible software to open downloaded transaction file and save as a new standard .xls or .xlsx file.
2. Navigate to the main project folder.
3. Double click `TradingHistoryClassifier.py` to run the program.
4. Follow the GUI guide.

## FAQ

Q: What environment does this program run?

*A: I develop and test it under Windows 10, python 3.7.  You could try to run it under different environment, but I can't promise it will run smoothly.*

Q: Could I add new trading history into processed spreadsheet file by this program?

*A: Yes, please following instruction to add it.*
- Step1: Delete transactions sheet from processed file (ex.my_transactions_r0.xlsx).
- Step2: Duplicate the new transactions sheet from new downloaded file (ex.transactions.xlsx) to processed file and save it.  Make sure all trading history are correct because this program won't check it for you.
- Step3: Open program and select modified file to classify.

Q: How to change program interface language?

*A: Click `Open Setting Editor` button and choose language for the interface.*

## ChangeLog

Please see the [CHANGELOG.md](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier/blob/develop/CHANGELOG.md) file for details.

The version control of this project follows [Semantic Versioning 2.0.0](https://semver.org/) system after Beta 1.3.1.

## Contributing

1. Fork it.
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

## Localization
1. Duplicate one copy of the `lang_enUS.json` in program file folder.
2. Rename the file to `lang_xxYY.json`.  The `xx` shall be the language code defined in [ISO 639-1 International standards for language codes](https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes) and the `YY` shall be two-letter country codes defined in [ISO 3166-1 Two-letter country codes](https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2).
3. Open new `lang_xxYY.json` by the proper text editor and translate all sentences to desired language.  If you don't know where to modify the file, please open `lang_zhTW.json` in the same program file folder.  By comparing `lang_enUS.json` and `lang_zhTW.json`, you are able to know where you have to modify the file. 
4. Open `settings.json` and find `ava_lang_for_GUI` section.  Add the language name in the format `zzzzz (xxYY)` under `ava_lang_for_GUI` section.  The definition for `xx` and `YY` are same as above point.2 mentioned.  The `zzzzz` shall be the language name you want to call.
5. If you do every steps mentioned above correctly.  You should be able to use the language you want in the program.
6. Finally, don't forget to create a new Pull Request to help other people using the same language.

## Authors

* **RavagerWT** - *Initial work* - [TD-Ameritrade-Trading-History-Classifier](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier)

## License

This project is licensed under the GPLv3 License - see the [LICENSE.md](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier/blob/master/LICENSE.md) file for details.
