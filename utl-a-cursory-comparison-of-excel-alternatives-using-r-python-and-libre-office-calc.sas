%let pgm=utl-a-cursory-comparison-of-excel-alternatives-using-r-python-and-libre-office-calc;

%stop_submission;

A cursory comparison of excel alternatives using R Python and Libre Office Calc

related
https://www.xelplus.com/python-in-excel-vs-vba/

/*                     _
 ___  ___   __ _ _ __ | |__   _____  __   ___  _ __
/ __|/ _ \ / _` | `_ \| `_ \ / _ \ \/ /  / _ \| `_ \
\__ \ (_) | (_| | |_) | |_) | (_) >  <  | (_) | | | |
|___/\___/ \__,_| .__/|_.__/ \___/_/\_\  \___/|_| |_|
                |_|
*/

Bottom line for mouse surfing use MS Excel, even when recording a macro. For a programatic control and embeded python scipts of use python.

FYI: Back in the day sas had a product call FSCALC.

github
https://tinyurl.com/56ew7kad
https://github.com/rogerjdeangelis/utl-a-cursory-comparison-of-excel-alternatives-using-r-python-and-libre-office-calc

/*___
|  _ \
| |_) |
|  _ <
|_| \_\

*/

Use Excel for manual, interactive spreadsheet and macro work.

Use R for automated, large-scale data manipulation, importing/exporting, and programmatic formula writing (not for interactive macro creation or execution).

Feature                 MS Excel                                 R
Creating spreadsheets   User-friendly GUI                        Programmatic via packages (e.g., openxlsx, readxl)
Reading spreadsheets    Native, direct open and view             Use packages like readxl, openxlsx, xlsx
Adding formulas         Direct cell input, GUI-based formulas    Can write formulas, not live formulas, export fomulas back to Excel
Creating macros         VBA macro editor, record & write macros  Cannot create native Excel macros; can automate Excel via R
Executing Excel macros  Directly in Excel; easy execution        Can trigger macro execution via R (e.g., using RDCOMClient)
ExcelGraphics           Superior Graphics                        Can programtically add grphics
Excek Syles             Wide range of style suppor               Can read and write styles, can eaven read colors

Use Excel for manual, interactive spreadsheet and macro work.

Use R for automated, large-scale data manipulation,
 importing/exporting, and programmatic formula writing
(not for interactive macro creation or execution).

/* _ _                       __  __ _                      _
| (_) |__  _ __ ___    ___  / _|/ _(_) ___ ___    ___ __ _| | ___
| | | `_ \| `__/ _ \  / _ \| |_| |_| |/ __/ _ \  / __/ _` | |/ __|
| | | |_) | | |  __/ | (_) |  _|  _| | (_|  __/ | (_| (_| | | (__
|_|_|_.__/|_|  \___|  \___/|_| |_| |_|\___\___|  \___\__,_|_|\___|

*/

LibreOffice Calc is more than adequate for most standard spreadsheet needs and offers some macro support for VBA,
but is much better suited for users who don't rely heavily on advanced Excel macros or integrations.

Feature              MS excel                                            Libre Office CALC

Creating/Reading     Seamless across formats, strong collaboration       Excellent format support, best for ODF and basic Excel
Formula Support      Full, advanced, leading industry standards          Very strong, some Excel-specific features may not port
Macro Language       VBA                                                 LibreOffice Basic, Python, JavaScript, BeanShell
Excel Macro Support  Full                                                Some support, requires editing/tweaks for VBA
Macro Editor         Robust VBA IDE                                      Decent built-in IDE, but less advanced
Macro Execution      All macros run as intended                          Many simple Excel macros work, advanced ones less so
Cost                 Paid, subscription/licensing                        100% Free and Open Source
Platforms            Windows, macOS, iOS, Android, Web                   Windows, macOS, Linux, some BSDs


/*           _   _
 _ __  _   _| |_| |__   ___  _ __
| `_ \| | | | __| `_ \ / _ \| `_ \
| |_) | |_| | |_| | | | (_) | | | |
| .__/ \__, |\__|_| |_|\___/|_| |_|
|_|    |___/
*/

1. Creating and Reading Excel Files

| Task         | MS Excel (Native)                                    | Python with Excel (Libraries or Cloud)                                |
|--------------|------------------------------------------------------|-----------------------------------------------------------------------|
| Creating     | Point-and-click GUI, instant file creation           | - Programmatic file creation (pandas, openpyxl, xlsxwriter, etc.)[3]  |
|              | Instantly start entering data and formatting         | - Write scripts to specify sheet structure, datatypes, cell contents  |
| Reading      | GUI file open, drag-and-drop                         | - Read whole sheets, selected cells, or ranges via code               |
|              | Supports all Excel formulas and formatting by default| - Can add logic to preprocess/analyze data as you read                |

2. Adding & Using Formulas

| Task            | MS Excel (Native)                                 | Python with Excel (Libraries or Cloud)                                |
|-----------------|---------------------------------------------------|-----------------------------------------------------------------------|
| Adding formulas | Enter formulas directly (`=A1+B1`) in cells       | - Write formulas into cells with code (`cell.formula = '=A1+B1'`      |
|                 | Use built-in function wizard                      | - With pandas/xlwings, perform logic/calculation in Python, to cells  |
| Calculating     | Formulas update automatically as you edit         | - Calculations in Python cells via “Python in Excel” cloud or script

3. Macros & Automation

| Task               | MS Excel (Native, VBA Macros)                  | Python with Excel (Libraries,Cloud,xlwings,pyxll, or Python in Excel)  |
|--------------------|------------------------------------------------|------------------------------------------------------------------------|
| Creating Macros    |Record or hand-write VBA                        | Write Python functions as Excel macros pyxll and xlwings, Python & COM |
| Running Macros     |Run directly Excel menu,assign buttons events   | Use xlwings/pyxll to expose Python scripts as macros callable from Excel
|                    |Macros are embedded in workbook as `.xlsm`      | Can trigger Excel VBA macros from Python scripts or automate workflow  |
| Macro Security     |Macros can be disabled                          | Python macros via xlwings/pyxll require add-in installations;          |
| Assigning to GUI   |Easy assignment to buttons, shapes              | Requires manual setup to call Python macros via button macro name      |

4. Executing Macros

| Task      | MS Excel (Native VBA)                                   | Python with Excel (Local/Cloud/COM integration)
|-----------|---------------------------------------------------------|------------------------------------------------------------------------|
| Execution | - Immediate, runs within Excel                          | Via xlwings, pyxll, or win32com: Python launches Excel,opens,runs macro|
| Automation| - Can chain macros or trigger from events (open/save)   | Full automation with Python scripts: open Excel, run macro, close/save |


Key Distinctions and Notes**

- **Excel Native (VBA):**
  - Deepest integration, best suited for users already familiar with Excel/VBA.
  - Macros stored inside workbook; portable but could face restrictions due to security settings.
  - Limited for complex data processing and new libraries (e.g., ML, web scraping).

- **Python with Excel:**
  - More powerful for complex logic, data analysis and automation.
  - “Python in Excel” (cloud-based) lets you run Python directly inside Excel cells, but this code executes on Microsoft Azure, not locally
  - Macro handling is possible but typically needs add-ons (like pyxll, xlwings).
  - Running native Excel (VBA) macros from Python is supported via libraries like `win32com`.[1][7]
  - Not every workflow (esp. company-internal with tight security) will allow full Python-Excel integration.

***

*When to Use Which?**

- **Use Native Excel when:**
    - You rely heavily on robust GUI, built-in Excel features, embedding/re-using complex macros, or require tight Office integration.

- **Use Python with Excel when:**
    - You need advanced analytics, want to automate across multiple files, need access to modern data science/ML tools,
      or when “code as logic” is preferable to “point/click”.

***

**Summary:**
- **MS Excel** is best for direct, interactive file work and native automation with macros.
- **Python with Excel** (via libraries or cloud) offers programmatic flexibility, access to advanced analytics,
and scalable automation, but often requires extra configuration for interacting with Excel macros and GUI elements.

/*                     _                        __  __
 ___  ___   __ _ _ __ | |__   _____  __   ___  / _|/ _|
/ __|/ _ \ / _` | `_ \| `_ \ / _ \ \/ /  / _ \| |_| |_
\__ \ (_) | (_| | |_) | |_) | (_) >  <  | (_) |  _|  _|
|___/\___/ \__,_| .__/|_.__/ \___/_/\_\  \___/|_| |_|
                |_|
*/
/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
