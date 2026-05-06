# xlUtils — Core Excel Utility Library

**Module:** `xlUtils`  
**Version at archive:** v37.02 (2020-10-08)  
**Used by:** most Excel and shared automation modules in this codebase

---

## Purpose

The main shared utility library for Excel automation. Reached v37 over ~8 years of development, accumulating tools across many areas:

- Sheet, shape, and workbook management
- Named range utilities (create, scope, find/replace, copy)
- String/text processing (capitalisation, sanitisation)
- File and folder operations
- Data connection and link management
- Conditional formatting replication
- Formula manipulation
- Chart image export
- Scheduler / auto-close timer
- Read-only file management
- MySQL ODBC driver detection
- WMI-based utilities (OS detection, Excel process count)

Incorporates `modKeyState` (Chip Pearson) for Shift key detection.

---

## Module-level constants

```vb
Private Const sAutomationUser As String = "pcinsightvm|pcopsvm"
```

Used throughout to detect VM automation accounts and bypass interactive prompts.

File extension constants (`xxls`, `xxlx`, `xxlm`, `xxlb`, `xxtm`, `xcsv`) are used in file conversion and export routines.

---

## Functions and Subroutines

### Version

#### `xlUtilsVer`
```vb
Function xlUtilsVer() As String
```
Returns the module version string (`xlUtilsCodeVer`).

---

### Sheet and Workbook Utilities

#### `xlU_SelectLotsOfSheets`
```vb
Sub xlU_SelectLotsOfSheets(ByVal str As String, Optional ByRef wb As Excel.Workbook)
```
Selects all sheets in `wb` whose name contains `str`. Pass `str = ""` to select all sheets. Useful for group formatting operations.

---

#### `xlU_ShowAllSheets`
```vb
Sub xlU_ShowAllSheets(Optional ByRef wb As Excel.Workbook)
```
Sets all sheets in the workbook to `Visible = True`. Works on `ActiveWorkbook` if no workbook is specified.

---

#### `xlU_ShowAllObjects`
```vb
Sub xlU_ShowAllObjects(Optional ByRef ws As Excel.Worksheet)
```
Sets all shapes on a sheet to `Placement = xlMoveAndSize`. Allows column/row deletion when hidden comments or shapes would otherwise block it.

---

#### `xlU_DeleteAllObjects`
```vb
Sub xlU_DeleteAllObjects(Optional ByRef ws As Excel.Worksheet)
```
Deletes all shapes (including comments, charts, images, buttons) from a sheet.

---

#### `xlU_SafeToQuitExcel`
```vb
Function xlU_SafeToQuitExcel(
    Optional ByVal QuitIfSafeOtherwiseCloseThisWorkbookWithoutSaving As Boolean,
    Optional ByVal DisableCloseEvents As Boolean
) As Boolean
```
Returns `True` if all open workbooks (excluding `PERSONAL.XLSB` and `ThisWorkbook`) are saved. If `QuitIfSafeOtherwiseCloseThisWorkbookWithoutSaving = True`, either quits Excel (if safe) or closes ThisWorkbook without saving. Used at the end of overnight automation runs.

---

#### `xlU_WorkbookIsReadOnly`
```vb
Function xlU_WorkbookIsReadOnly(
    Optional wb As Excel.Workbook,
    Optional xluXLapp As Excel.Application
) As Boolean
```
Checks whether a workbook is read-only by inspecting the window caption for `"Read-Only"`, rather than using `wb.ReadOnly`. More reliable for SharePoint-hosted workbooks where `.ReadOnly` is unreliable.

---

#### `xlU_IsExcelRunning`
```vb
Function xlU_IsExcelRunning(Optional bDisplayMsg As Boolean) As Integer
```
Returns the count of running `Excel.exe` processes via WMI. Useful before opening Excel instances from other Office applications.

---

### Shape and Comment Utilities

#### `xlU_Reset_Comment_Sizes`
```vb
Sub xlU_Reset_Comment_Sizes(
    Optional ByRef ws As Excel.Worksheet,
    Optional ByVal CmtHeight As Single,
    Optional ByVal CmtWidth As Single
)
```
Resets all comments on a sheet to a standard size. Default: 60 × 100 (height × width). Pass custom sizes as needed.

---

#### `xlU_Comments_MoveOnly`
```vb
Sub xlU_Comments_MoveOnly(Optional ByRef ws As Excel.Worksheet)
```
Sets all comment boxes on a sheet to `Placement = xlMove` (moves with cells, doesn't resize). Prevents comments from obscuring content after row/column insertion.

---

### String and Text Utilities

#### `xlU_Clean_Special`
```vb
Function xlU_Clean_Special(
    ByVal str As String,
    Optional ByVal CropLength As Boolean = True,
    Optional ByVal OnlyFilename As Boolean,
    Optional ByVal OnlyVBObjectName As Boolean,
    Optional ByVal OnlySheetName As Boolean
) As String
```
Removes invalid special characters from a string. Modes:

| Mode | Removes |
|------|---------|
| Default | `~ " # % & * : < > ? { | } [ ] ..` |
| `OnlyFilename` | Above + `\ /` |
| `OnlyVBObjectName` | Above + `\ / - (space) , . ( ) '` → replaced with `_`; invalid leading chars stripped; max 35 chars |
| `OnlySheetName` | Above + crops to 31 chars |

`CropLength = True` silently crops instead of showing an error MsgBox.

---

#### `xlU_Remove_Spaces`
```vb
Function xlU_Remove_Spaces(Optional ByVal str As String) As String
```
Removes trailing spaces from a string, or from all cells in `Selection` if no string is provided. Shows a warning MsgBox before modifying cells (not suitable for silent automation).

---

#### `xlU_Text_Capitalise_Phrase`
```vb
Function xlU_Text_Capitalise_Phrase(
    Optional ByRef rPhrases As Excel.Range,
    Optional ByVal sWordOrPhrase As String
) As String
```
Title-cases a string or a range of cells. Delegates to `xlU_Text_Capitalise_Words` for the word-level logic. Handles hyphenated words, slash-separated phrases, and bracketed words.

---

#### `xlU_Text_Capitalise_Words`
```vb
Function xlU_Text_Capitalise_Words(ByVal s As String) As String
```
Core capitalisation engine. After initial title-casing, applies a large set of overrides for common two-letter words (`of`, `in`, `UK`, `St`, `BV`, `NV`), company name abbreviations (`DHL`, `GmbH`, `PLC`, `USAF`, `TNT`), and compound forms (`SRL`, `LLP`, `Sp Z Oo`, `PO Box`, `T/A`, `C/O`).

---

#### `xlU_Text_Capitalise_Words_Exceptions`
```vb
Function xlU_Text_Capitalise_Words_Exceptions(ByVal w As String) As String
```
Exception handler called per-word by `xlU_Text_Capitalise_Words`. Applies word-level rules: 2-character words go fully uppercase; alphanumeric codes go uppercase (unless ordinal like `1st`, `2nd`); specific company/country codes corrected.

---

#### `xlU_Numeric_To_Text`
```vb
Sub xlU_Numeric_To_Text(ByRef xlU_Range As Excel.Range)
```
Prefixes numeric cell values with an apostrophe to force text storage — useful for preserving leading zeros in codes. Note: Excel may have already converted the value before this runs (e.g. `07` → `7`).

---

### File and Folder Utilities

#### `xlU_FileFolderExists`
```vb
Function xlU_FileFolderExists(ByVal strFullPath As String) As Boolean
```
Returns `True` if a file or folder exists at the specified path. Source: Ken Puls (www.excelguru.ca).

---

#### `xlU_EmptyFolder`
```vb
Sub xlU_EmptyFolder(
    ByVal fdr As String,
    Optional ByVal AlsoRmDir As Boolean,
    Optional ByVal DoMsgs As Boolean
)
```
Deletes all files from a folder. Optionally removes the folder itself (`AlsoRmDir = True`). Shows a warning MsgBox before proceeding; not suitable for silent automation.

---

#### `xlU_List_Dir_Contents`
```vb
Sub xlU_List_Dir_Contents(Optional ByVal SpecifyPath As String)
```
Creates a new sheet named `"xlU_List_Dir_Contents"` in the active workbook, listing all filenames from the specified folder (default: the workbook's own directory). Replaces any existing sheet with that name.

---

#### `xlU_GetSpecialFolderNames`
```vb
Function xlU_GetSpecialFolderNames(Optional ByVal DoDebug As Boolean = True)
```
Prints Windows special folder paths (Desktop, Documents, Favorites, etc.) to the Immediate window or MsgBox. Reference/debug utility only.

---

#### `xlU_Deploy_Folders`
```vb
Sub xlU_Deploy_Folders(
    ByVal BasePath As String,
    Optional ByVal LevelLookup As String,
    Optional ByVal sLevel3Header As String = "Team",
    Optional ByVal sLevel2Header As String = "Manager",
    Optional ByVal sLevel1Header As String = "Destination",
    Optional ByVal AddAlphanumericSubfolders As Boolean = False,
    Optional ByVal DebugMode As Boolean = False
)
```
Creates a hierarchical folder structure under `BasePath` driven by a lookup Excel file. The lookup file must have columns matching the specified level header names. Up to 3 nesting levels. Optionally adds A–Z and `0-9` sub-folders within each level-1 folder.

---

### Named Range Utilities

#### `xlU_Ranges_Add_Named_After_Column_Headers`
```vb
Sub xlU_Ranges_Add_Named_After_Column_Headers(
    Optional ByRef xlU_Worksheet As Excel.Worksheet,
    Optional IncludeHeaders As Boolean = True
)
```
Creates workbook-scoped named ranges for each column, named `"[SheetName]_[ColumnHeader]"` and sized to match the data length in column 1.

---

#### `xlU_Ranges_Set_To_Column_1_Data_Rows`
```vb
Sub xlU_Ranges_Set_To_Column_1_Data_Rows()
```
Adjusts the length of all named ranges on query sheets to match the actual data row count in column 1. Skips `_FilterDatabase` names and `v_` prefix (validation) names. Useful after a data refresh changes the number of rows.

---

#### `xlU_Ranges_Change_Scope`
```vb
Function xlU_Ranges_Change_Scope(
    ByVal xluWorksheet As Excel.Worksheet,
    ByVal xluWorkbook As Excel.Workbook,
    Optional ByVal xluScopeChange As Byte = 0,
    Optional ByVal xluDeleteOriginalName As Boolean = False
) As Boolean
```
Moves named ranges between worksheet scope and workbook scope. `xluScopeChange = 1` promotes worksheet names to workbook; `= 2` demotes workbook names to worksheet.

---

#### `xlU_Transfer_Ranges`
```vb
Function xlU_Transfer_Ranges(ByRef srcws As Excel.Worksheet, tgtws As Excel.Worksheet)
```
Copies all named ranges from `srcws` to `tgtws`, replacing the source sheet name with the target sheet name in both the range name and reference.

---

#### `xlU_Range_Names_FindAndReplace`
```vb
Sub xlU_Range_Names_FindAndReplace(
    ByVal oldstring As String,
    ByVal newstring As String,
    Optional WithinWorkbook As Excel.Workbook,
    Optional WithinSheet As Excel.Worksheet
)
```
Find-and-replace in named range names (not their values). Operates on all names in the active workbook if no scope is specified.

---

#### `xlU_Clean_Name_Errors`
```vb
Function xlU_Clean_Name_Errors(
    Optional ByVal bShowMsgBox As Boolean = True,
    Optional ByVal bDeleteErrors As Boolean = True,
    Optional ByVal bDeleteExternal As Boolean = True
) As Integer
```
Removes `#REF!` named ranges and named ranges that reference external local/network paths. Returns the count of deleted names. Useful after moving a workbook, which leaves orphaned names pointing to old paths.

---

### Data Connection and Link Utilities

#### `xlU_RemoveAllConnections`
```vb
Sub xlU_RemoveAllConnections(ByRef wb As Excel.Workbook)
```
Removes all data connections from the workbook except the first (which is necessary for pivot table functions). Note: deletes connection 1 repeatedly in a loop — relies on the list re-indexing on each delete.

---

#### `xlU_RemoveUnusedConnections`
```vb
Sub xlU_RemoveUnusedConnections(ByRef wb As Excel.Workbook)
```
**DISABLED** — body exits immediately with a MsgBox warning. The original implementation caused serious problems and was never fixed.

---

#### `xlU_BreakLinks`
```vb
Function xlU_BreakLinks(ByRef wb As Excel.Workbook) As Boolean
```
Breaks all external Excel workbook links. Returns `False` if there are no links or on error; `True` on success.

---

#### `xlU_UpdateLinks`
```vb
Function xlU_UpdateLinks(Optional wb As Excel.Workbook) As Boolean
```
Updates all external Excel links. Returns `False` on error.

---

### Formula and Cell Utilities

#### `xlU_Find_And_Replace_Text`
```vb
Sub xlU_Find_And_Replace_Text(
    ByVal oldstring As String,
    ByVal newstring As String,
    Optional ByRef RangeToFindAndReplace As Excel.Range
)
```
Find-and-replace in cell values. Works on `Selection` if no range specified. Bypasses sheet protection restrictions that block the normal Find & Replace dialog.

---

#### `xlU_Add_IFERROR`
```vb
Sub xlU_Add_IFERROR(
    Optional ByVal ErrorResultRef,
    Optional ByRef AddToRange As Excel.Range
)
```
Wraps `=IFERROR([existing formula], ErrorResultRef)` around formulas in the selected/specified range. Skips cells that are already wrapped in `IFERROR`. Works on `Selection` if no range specified.

---

#### `xlU_Replace_In_Formulas`
```vb
Sub xlU_Replace_In_Formulas(
    ByVal sOldText As String,
    ByVal sNewText As String,
    Optional ByRef DoInRange As Excel.Range
)
```
Find-and-replace within formula strings only (cells starting with `=`). Leaves non-formula cells untouched. Not case-sensitive (module uses `Option Compare Text`).

---

#### `xlU_TransferValidationList`
```vb
Function xlU_TransferValidationList(
    ByRef vSource As Excel.Range,
    ByRef vTarget As Excel.Range
) As Boolean
```
Copies the full data validation definition from one single cell to another: type, operator, formula, messages, error title, all flags. Returns `True` on success.

---

### File Format Conversion

#### `xlU_Convert_File`
```vb
Function xlU_Convert_File(
    ByRef PathFile As String,
    ByVal FileFormat As XlFileFormat,
    ByVal SourceIsXML As Boolean,
    Optional ByVal DeleteOriginal As Boolean
) As Boolean
```
Opens a file and resaves it in a different format (CSV, XLS, XLSX, XLSM). Creates a temp copy first, so the original is only deleted if `DeleteOriginal = True` and the conversion succeeds. Supported `FileFormat` values: `6` (CSV), `56` (XLS), `51` (XLSX), `52` (XLSM).

---

#### `xlU_Export_Single_Sheets`
```vb
Sub xlU_Export_Single_Sheets(
    ByVal OutputPath As String,
    ByVal OutputXLSX As Boolean,
    ByVal OutputXLS As Boolean,
    ByVal OutputCSV As Boolean,
    Optional ByVal xlU_Password As String
)
```
Copies each sheet to a separate workbook and saves in the chosen format(s). Output filename: `[ThisWorkbook.Name (no extension)] [SheetName].[ext]`. Strips empty rows below data and breaks external links in each output file.

---

### Conditional Formatting

#### `xlU_CF_Replicate_Rows`
```vb
Sub xlU_CF_Replicate_Rows(
    ByRef SourceRange As Excel.Range,
    ByRef TargetRange As Excel.Range
)
```
Copies conditional formatting from a single source row to a target range (which can be many rows). Replicates up to 3 conditions per column: type, operator, formula(s), font bold/italic/underline/strikethrough/colour, interior colour, and border style/weight/colour.

**Constraints:** source and target must have the same column count; source must be 1 row; source row 1 should be the same as (or intersect with) the target row 1, otherwise formula row references will be misaligned.

---

### Chart and Image Utilities

#### `xlU_Save_Chart_As_Image`
```vb
Function xlU_Save_Chart_As_Image(ByVal ChartNumber As String) As String
```
Exports a named chart from the `"EMAIL CHARTS"` sheet as a GIF file. For automation users (`pcinsightvm`, `pcopsvm`), saves to `\\company.local\Group Data\Automation\Testing\Chart Images\`; for all others, saves to `%TEMP%`. Filename includes the Monday date of the current week.

Returns the full path to the saved GIF, for use as `ImagePath1` in `modEmail.SendEmail`.

---

### Read-Only File Management

#### `xlU_Remove_Read_Only`
```vb
Sub xlU_Remove_Read_Only()
```
Removes the read-only file attribute from the current workbook and reopens it for editing via `ChangeFileAccess xlReadWrite`. Waits 2 seconds after `SetAttr` to allow the file system to update. Silent for automation users.

---

#### `xlU_Save_Make_Read_Only_And_Close`
```vb
Sub xlU_Save_Make_Read_Only_And_Close()
```
Saves the current workbook, sets its file attribute to read-only, and closes it (via a temp save to `%TEMP%\test.xlsb`). Only works on `.xlsb` files. Silent for automation users; shows a confirmation dialog for interactive users.

---

### Scheduler / Auto-Close Timer

#### `xlU_Remind_User`
```vb
Sub xlU_Remind_User(
    Optional ByVal BypassMessageBox As Boolean = True,
    Optional AutoCloseAndSave As Boolean = True
)
```
Implements a 5-minute inactivity timer using `Application.OnTime`. Designed to auto-close and save a workbook if a user leaves it open (preventing overnight automation from being blocked).

Call from `Worksheet_SelectionChange` to reset the timer on each user interaction. The timer only runs when the workbook name does not contain `"(master)"` and the user is not an automation account.

`RunWhen` is a module-level public `Double` variable holding the scheduled fire time.

---

### MySQL Driver Detection

#### `get_MySQLDriverName`
```vb
Function get_MySQLDriverName() As String
```
Returns the name of the installed MySQL ODBC driver by enumerating `HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers` via WMI `StdRegProv`. Returns the first registry value name containing `"mysql"` (case-insensitive). Returns an empty string if no MySQL driver is installed.

Used to build dynamic connection strings for MySQL data sources without hardcoding the driver name.

---

### Design Mode Helpers

#### `xlU_Exit_Design_Mode` / `xlU_Enter_Design_Mode`
```vb
Sub xlU_Exit_Design_Mode()
Sub xlU_Enter_Design_Mode()
```
Programmatically toggle Excel design mode via the CommandBar. Useful after deploying ActiveX controls. Note: `xlU_Enter_Design_Mode` may toggle rather than reliably enter.

---

### Interactive / Miscellaneous

#### `xlU_Cut_Multiple_Rows_to_New_Location`
```vb
Sub xlU_Cut_Multiple_Rows_to_New_Location(
    Optional ByVal DeleteEmptyRows As Boolean,
    Optional ByRef ws As Excel.Worksheet
)
```
Prompts for a named range (`temp_range`) and a destination blank row number, then moves all non-blank rows from the range to that destination. Requires the named range to be created manually first.

#### `xlU_Extract_Outlook_Emails`
```vb
Sub xlU_Extract_Outlook_Emails()
```
Reads `"Name <email>; ..."` text from the active cell and extracts addresses into cells below. Interactive utility for extracting email lists from Outlook pastes.

#### `xlU_Pause_for_Timeout`
```vb
Function xlU_Pause_for_Timeout(Optional ByVal TimeOutInSecs As Long) As Boolean
```
Intended countdown that user can cancel by holding Shift. **Marked as non-functional in source code comments** ("it doesn't work!??"). The Shift key detection via `IsShiftKeyDown` does not work reliably inside a tight VBA loop.

---

## Dependencies

- Excel object model (early binding — requires `Microsoft Excel Object Library` reference)
- WMI (`winmgmts`) for `get_MySQLDriverName` and `xlU_IsExcelRunning`
- Windows Registry via WMI `StdRegProv` for `get_MySQLDriverName`
- `modKeyState` (Chip Pearson) — integrated directly into this module for Shift key detection

---

## Known quirks

- **`Option Base 1`**: added at v27. All array declarations in this module are 1-based. Watch for off-by-one errors if adapting code.
- **`Option Compare Text`**: `xlU_Replace_In_Formulas` is explicitly not case-sensitive because of this module-level setting.
- **`xlU_Pause_for_Timeout`**: documented as broken. Do not rely on it.
- **`xlU_RemoveUnusedConnections`**: disabled entirely — exits immediately.
- **`xlU_CF_Replicate_Rows` formula alignment**: conditional format formulas are copied verbatim, not adjusted for the target row. Works correctly only if the source row is also row 1 of the target range.
- **`xlU_Remove_Spaces` and `xlU_EmptyFolder`**: both show confirmation MsgBoxes and are not suitable for silent automation.
- **Late binding throughout**: most Excel object variables are declared `As Object` rather than `As Excel.Worksheet` etc., even though the module uses early-bound references. This was intentional to allow the module to be loaded from non-Excel hosts (Access, Outlook).
