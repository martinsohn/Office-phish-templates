# Office phish templates and defense recommendations
Tricks the target into enabling content (macros) with fake messages.
Once enabled, uses macros to reduce the risk of suspision from target user via verious methods.

Templates are available with a example macro code or without (macro code for each template can be seen in screenshots below):
* With macros
  * .xlsm
  * .docm
* Without macros
  * .xlsx
  * .docx

## Notes to defend against macro attacks
* Enforce security settings for macros in all Office applications 
  * Official Microsoft documentation: [Plan security settings for VBA macros in Office 2016](https://docs.microsoft.com/en-us/DeployOffice/security/plan-security-settings-for-vba-macros-in-office)
  * My recommendation is to enforce ***Disable all except digitally signed macros*** and ***Block macros from running in Office files from the Internet***, even partly deploying such policy reduces risk.
* Address risk of certain Office extensions used for macro documents
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
  * Block in spam filter, mailflow rules, EDR/AV
  * Remove default file associations, or associate with Notepad
    * Check current associations with cmd.exe: `assoc | findstr /i "word excel powerpoint"`
  * Add to detection rules
    * Sysmon - Event ID 11: File Creation Events, Event ID 23: FileDelete
  * Be aware that regular Office extensions may also contain macros, these include (but not limited to): .xls, .doc, .rtf, .wbk
* Enforce [attack surface reduction rules](https://docs.microsoft.com/en-us/microsoft-365/security/defender-endpoint/attack-surface-reduction?view=o365-worldwide), here are some which relates directly to Office
  * Block all Office applications from creating child processes
  * Block Office applications from creating executable content
  * Block Office applications from injecting code into other processes
  * Block Office communication application from creating child processes
  * Block Win32 API calls from Office macros
* Create user awareness
  * The Danish sikkerdigital.dk has [provides user-focuced awareness material in both English and Danish at headline "Medarbejderpakken".](https://sikkerdigital.dk/virksomhed/test-og-vaerktoejer)
* More information on macros and misconceptions: [Office Macros – file extensions, file format (content), and a few handling stereotypes…](https://www.hexacorn.com/blog/2016/11/05/office-macros-file-extensions-file-format-content-and-a-few-handling-stereotypes/)

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
