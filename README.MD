# Excel Add-In which adds a graphical user interface to use the rounding functions

Currently the user interface is available in English and German. The language is chosen by the add-in according to the language of the Excel GUI.

To execute the unit tests in the module mdl_UnitTestRound.bas the the COM add-in rubberduck needs be installed (https://github.com/rubberduck-vba/Rubberduck).
(Thank you guys for your work)

## Installation

Download the file download_inoRound.zip to your system and unzip the content.

In Excel use File - Options

![fileoptions](./images/fileoptions.png)

Choose tab Add-ins and select "Manage Excel-Add-Ins" and click go.

![addins](./images/addins.png)

Use "Browse" to select inoRound.xlam and make sure Inoround is ticked.

Now the Add-In is available.

## Usage

The Add-In is available on the Home Tab between Number and Styles

![screenshot](./images/screenshot.PNG)

To round figures in calculations select the cells and use:

1 to round flexible

2 to round up

3 to round down

The number of digits to round to is defined at 4.

If 5 "Round numbers"  is ticked given figures are rounded as well.

6 remove existing round