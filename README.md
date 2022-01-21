# Office phish templates
Tricks the target into enabling content (macros) with fake messages.
Once enabled, uses macros to reduce the risk of suspision from target user via verious methods.

Templates are available with a example macro code or without (macro code for each template can be seen in screenshots below):
* With macros
  * .xlsm
  * .docm
* Without macros
  * .xlsx
  * .docx

## Notes to defend against these
* Disable macros in all Office applications
  * https://www.microsoft.com/security/blog/2016/03/22/new-feature-in-office-2016-can-block-macros-and-help-prevent-infection/
* Detect and block Office extensions used for macro documents in: spam filter, AV, Sysmon, etc.
  * Some extensions
    * .docm=Word.DocumentMacroEnabled.12
    * .dotm=Word.TemplateMacroEnabled.12
    * .xlam=Excel.AddInMacroEnabled
    * .xlm=Excel.Macrosheet
    * .xlsb=Excel.SheetBinaryMacroEnabled.12
    * .xlsm=Excel.SheetMacroEnabled.12
    * .xltm=Excel.TemplateMacroEnabled
    * .potm=PowerPoint.TemplateMacroEnabled.12
    * .ppsm=PowerPoint.SlideShowMacroEnabled.12
    * .pptm=PowerPoint.ShowMacroEnabled.12
    * .sldm=PowerPoint.SlideMacroEnabled.12
  * Be aware that regular Office extensions may also contain macros, these include (but not limited to): .xls, .doc, .rtf, .wbk
  * [Office Macros – file extensions, file format (content), and a few handling stereotypes…](https://www.hexacorn.com/blog/2016/11/05/office-macros-file-extensions-file-format-content-and-a-few-handling-stereotypes/)
* Create user awareness!

## Notes to increase phishing success
* Saving as 97-2003 document (eg. .xls) gives an old-style icon without the small script icon that you will see for e.g. .xlsm.
* Clear unwanted meta-data via Info -> Check for Issues -> Inspect Document
* Add wanted meta-data to reduce risk of suspicion

## Methods

### Excel
Hide the sheet containing the fake message, show a legitimate-looking sheet, and opens calc.exe
1. Add macro code to 'ThisWorkbook' function 'WorkbookOpen()' (Developer tab -> Visual Basic).
2. Fill 'Sheet1' with legitimate-looking data
3. Hide 'Sheet1'
4. Protect sheet 'Excel version mismatch'. Uncheck 'Select (un)locked cells'

If you change sheet names, you must do so too in the code that switches sheets!

#### Demonstration
![](/excel-demo.gif)

#### Document presented to user
![](/excel-document.PNG)

#### Code to perform method
![](/excel-code.PNG)

### Word
Fake error message popup, which when closed opens notepad.exe
1. Add macro code to 'ThisDocument' function 'MessageClosed()' (Developer tab -> Visual Basic).
2. Edit popup failure code to a unique identifier. Can be used to e.g. verify that the target enabled macros.
3. Review -> Restrict Editing: Allow only 'Filling in forms'.

#### Demonstration
![](/word-demo.gif)

#### Document presented to user
![](/word-document.png)

#### Code to perform method
![](/word-code.png)
