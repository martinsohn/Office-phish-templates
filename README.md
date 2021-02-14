# Excel phish template
Tricks the user into enabling content (macros) with a fake "Excel version mismatch" message.
Once enabled, uses macros to hide the message and show a legitimate-looking sheet which reduces risk of suspision from target user.

1. Add macro code to 'ThisWorkbook' (Developer tab -> Visual Basic). Opens calc.exe by default.
2. Fill 'Sheet1' with legitimate-looking data
3. Hide 'Sheet1'
4. Protect sheet 'Excel version mismatch'. Uncheck 'Select (un)locked cells'
5. Save as Excel Macro-Enabled Workbook (.xlsm), or(Excel 97-2003 Workbook (.xls). Saving as .xls uses an old-style Excel icon without the small script icon that you will see for .xlsm.

Notes:
* If you change sheet names, you must do so too in the code that switches sheets!
* Clear unwanted workbook meta-data via Info -> Check for Issues -> Inspect Document
* Add wanted workbook meta-data to reduce risk of suspicion

## Demonstration
![demo](/demo.gif)

## Fake message
![document](/document.PNG)

## Code to switch worksheets and open calc.exe
![code](/code.PNG)
