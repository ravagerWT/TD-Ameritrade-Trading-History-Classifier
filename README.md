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
```
git clone https://github.com/your_username_/Project-Name.git
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

1. Add new trading history into a existing processed file is not supported by current version.  You need to manually integrate them into existing spreadsheet file.

## Contributing

1. Fork it (<https://github.com/yourname/yourproject/fork>)
2. Create your feature branch (`git checkout -b feature/fooBar`)
3. Commit your changes (`git commit -am 'Add some fooBar'`)
4. Push to the branch (`git push origin feature/fooBar`)
5. Create a new Pull Request

## Authors

* **RavagerWT** - *Initial work* - [Trading-History-Classifier](https://github.com/your-project-name)

## License

This project is licensed under the GPLv3 License - see the [LICENSE.md](LICENSE.md) file for details
