# TD Ameritrade Trading History Classifier

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

Or follow the qucik step below:
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

Download project .zip file and extract it.

## How to use

1. Get transaction file:
   - Download your transaction file from TD Ameritrade website.  The transaction file downloaded from TD Ameritrade website should be .csv file.
   - Use MicroSoft Excel or any compatible software to open downloaded transaction file and save as a new standard .xls or .xlsx file.
2. Navigate to the main project folder.
3. Double click `TradingHistoryClassifier.py` to run the program.
4. Follow the GUI guide.

## FAQ

Q: What environment does this program run?

*A: I develop and test it under Windows 10, python 3.7.  You could try to run it under differnt environment, but I can't promise it will run smoothly.*

Q: Could I add new trading history into processed spreadsheet file by this program?

*A: Add new trading history into a existing processed file is not supported by current version.  You need to manually integrate them into existing spreadsheet file.*

Q: How to change program interface language?

*A: Click `Open Setting Editor` button and choose language for the interface.*

## ChangeLog

Please see the [CHANGELOG.md](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier/blob/develop/CHANGELOG.md) file for details.

## Contributing

1. Fork it.
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

## Authors

* **RavagerWT** - *Initial work* - [TD-Ameritrade-Trading-History-Classifier](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier)

## License

This project is licensed under the GPLv3 License - see the [LICENSE.md](https://github.com/ravagerWT/TD-Ameritrade-Trading-History-Classifier/blob/master/LICENSE.md) file for details.
