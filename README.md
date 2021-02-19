# Office phish templates
Tricks the target into enabling content (macros) with fake messages.
Once enabled, uses macros to reduce the risk of suspision from target user via verious methods.

## Notes to increase success
* Saving as 97-2003 document (eg. .xls) gives an old-style icon without the small script icon that you will see for e.g. .xlsm.
* Clear unwanted meta-data via Info -> Check for Issues -> Inspect Document
* Add wanted meta-data to reduce risk of suspicion

##Methods:
###Excel
Hide the sheet containing the fake message, show a legitimate-looking sheet, and opens calc.exe
1. Add macro code to 'ThisWorkbook' function 'WorkbookOpen()' (Developer tab -> Visual Basic).
2. Fill 'Sheet1' with legitimate-looking data
3. Hide 'Sheet1'
4. Protect sheet 'Excel version mismatch'. Uncheck 'Select (un)locked cells'

If you change sheet names, you must do so too in the code that switches sheets!

#### Demonstration
![excel-demo](/excel-demo.gif)

#### Fake message
![excel-document](/excel-document.PNG)

#### Code to switch worksheets and open calc.exe
![excel-code](/excel-code.PNG)

### Word
Fake error message popup, which when closed opens notepad.exe
1. Add macro code to 'ThisDocument' function 'MessageClosed()' (Developer tab -> Visual Basic).
2. Edit popup failure code to a unique identifier. Can be used to e.g. verify that the target enabled macros.
3. Review -> Restrict Editing: Allow only 'Filling in forms'.

#### Demonstration
![word-demo](/word-demo.gif)

#### Fake message
![word-document](/word-document.PNG)

#### Code to switch worksheets and open calc.exe
![word-code](/word-code.PNG)
