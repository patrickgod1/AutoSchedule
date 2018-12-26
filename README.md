# AutoSchedule
AutoSchedule is an application designed to automate the following:

* Downloading the Section Schedule Daily Summary from Destiny and creating a sorted/formatted report for:
	* ABSW
	* Belmont Center
	* Golden Bear Center
	* San Francisco Center

## Getting Started
These instructions will help you get started with using the application.

### Prerequisites
The following must be installed:
* Google Chrome
* Microsoft Excel
* Access to Destiny via CalNet ID

### Usage
1. Check the box next to the function(s) you would like to use
2. Fill in the require fields.
3. Click "Start" when ready. The output files will be saved to your "Save Path" location.
4. When prompted in Chrome, log in using CalNet credentials.
5. Click "Exit" to close the application.

Note: Runtime may vary depending on the number of days/classes that need the signs to be created for.

## Development
These instructions will get you a copy of the project up and running on your local machine for development.

### Built With
* [Python 3.6](https://docs.python.org/3/) - The scripting language used.
* [Pandas](https://pandas.pydata.org/) - Data structure/anaylsis tool used.
* [Selenium](https://selenium-python.readthedocs.io/) - Web crawling automation framework.
* [xlsxwriter](https://xlsxwriter.readthedocs.io/) - Used to create Microsoft Excel documents (Daily Schedule)
* [PyQt5](https://pypi.org/project/PyQt5/) - Framework used to create GUI.
* [QtDesigner](http://doc.qt.io/qt-5/qtdesigner-manual.html) - GUI builder tool.
* [PyInstaller](https://www.pyinstaller.org/) - Used to create executable for release.

### Running the Script
Run the following command to installer all the required Python modules:
```
pip install -r requirements.txt
```
To run the application:
```
.\AutoSchedule.py
```

### Compiling using PyInstaller

The project files includes a batch file (Windows platform only) with commands to run to compile into an executable. 

Other development platforms can run the following command in Terminal:

```
pyinstaller AutoSchedule.spec .\AutoSchedule.py
```
You may need to modify the file paths if not in same current working directory.

## Screenshot
![autoschedule](https://user-images.githubusercontent.com/41496510/50427756-9caf5980-0864-11e9-9337-17838e05c914.png)

## Authors
* **Patrick Yu** - *Initial work* - [patrickgod1](https://github.com/patrickgod1) - UC Berkeley Extension