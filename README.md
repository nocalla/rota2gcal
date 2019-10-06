# rota2gcal
This is a project that uploads a very specific timetable format from an Excel spreadsheet to Google Calendar. Details of the format are in a below section.

## Getting started

### 1. Download & Install Python

+ Download Python 3 directly from the Python Software Foundation's website: https://www.python.org/downloads/.

+ Install it, making sure to select the options to "Add Python to PATH" and the optional feature of installing "pip".

### 2. Download Project

+ Download the latest release from [here](https://github.com/nocalla/rota2gcal/releases). Really all you need is the `rota2gcal.py` file, but the batch file is useful to run the script quickly on Windows.
+ Extract the zip file to whatever location on your harddrive makes sense.

### 3. Install prerequisites

This is where things will get a little difficult if you're not used to the command line.
+ Open the folder where you've saved the `rota2gcal.py` file in Explorer
+ Type cmd in the address bar
+ In the terminal window that opens, type the following and hit Enter:
`pip install -r requirements.txt`
+ Close the terminal window

### 4. Activate Google Calendar API

This is how the script gets access to the calendar. You have to set yourself up as a "developer" to get access to the calendar tools. Brief description here.
+ Go to the Google Developers Console and login.
+ Use your Gmail or Google credentials; create an account if needed
+ Click "Create Project" button
+ Enter a Project Name (mutable, human-friendly string only used in the console)
+ Enter a Project ID (immutable, must be unique and not already taken)
+ Once project has been created, click "Enable an API" button and enable "Google Calendar"
+ You'll need to add Credentials to the project. There should be a button to enable those somewhere in the API site. 
+ Once you've got the Credentials created, download the file and save it to the same folder as `rota2gcal.py` and rename it to `client_secret.json`.

### 5. Initialisation

+ Run the script, either by typing `python rota2gcal.py` in a terminal window opened as described above, or by double clicking the `rota2gcal.bat` file.
+ You'll get an error, that's okay.
+ Now, edit the newly-created `rota2gcal.conf` file - this is where you set the name of the calendar you're writing to, the name of the person whose shifts you're looking for etc. Example configuration below. I'd recommend leaving the Source Folder blank, unless you download the timetable file to somewhere apart from the default Downloads folder on Windows. The Shared Calendar URL can be ignored too.
```
    [DEFAULT]
    Source Folder =
    Calendar Name = Test Calendar
    Shared Calendar URL =
    Person = Murgatroyd
    Event Location = Eiffel Tower
```
+ The first time you run the script in normal use you'll get a Google sign-in pop-up. Sign in and give it the permissions for accessing your calendar.
+ Make sure the `Calendar Name` field matches a calendar that you actually have in your Google Calendar.

### 6. That's it!

You're ready to go.

## Actual Use
+ Put the timetable file in the Source Folder or download folder
+ Double-click the `rota2gcal.bat` file to run the script or run it from the command line
+ Select the correct file from the list by typing the number and hitting Enter
+ Watch as your calendar magically fills up

## Spreadsheet Format Details
+ .xlsx format
+ `B1` shows the date that the current week starts on in the dd/mm/YYY format
+ Employee names are laid out in `column A`
+ The days of the week are laid out from `C2:H2`
+ The script ignores any cell with a formula in it
+ Each worksheet in the Excel workbook is named "Week 1", "Week 2", etc. This just indicates the order of the weeks and doesn't need to correspond to the official "Weeknumber" of the particular week in the calendar
+ Shifts must contain a double hyphen (i.e.: `--` ) between times, e.g. `9--6`
+ "hol" designates holidays - these will not be entered as events on the calendar
+ "off" designates days off - these will be entered as all-day events
