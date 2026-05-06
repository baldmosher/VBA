# accUtils — Access Utility Library

**Module:** `accUtils`  
**Version at archive:** v7.00 (2020-11-09)  
**References required:** Microsoft ActiveX Data Objects 2.0 Library; Microsoft Office Access Database Engine Objects Library

---

## Purpose

Shared utility library used across all Access automation modules. Covers three areas:

- **Table and query relinking** — update linked table paths and query SQL strings after moving a backend database or changing a server name
- **SQL Server connectivity** — open an ADODB connection and execute stored procedures via Windows authentication
- **Import/Export specification management** — update and run saved Access import specs with dynamic paths

Used by `accAuto`, `accCommercialAudits`, `accCommercialCompliance`, and `accRunCode`.

---

## Module-level variables

```vb
Private connDB As New ADODB.Connection
Private rs As New ADODB.Recordset
Private strSQL As String
Private strConnectionstring As String
Private strServer As String
Private strDBase As String
Private strUser As String
Private strPwd As String
```

These are module-level, so `connDB` persists for the lifetime of the Access session. Call `ado_ConnectDatabase` before each `ado_ExecStoredProcedure` call, or verify `connDB.State` before reuse.

---

## Functions and Subroutines

### `ado_ExecStoredProcedure`

```vb
Public Sub ado_ExecStoredProcedure(ByVal SProcName As String)
```

Executes a named SQL Server stored procedure via ADODB. Calls `ado_ConnectDatabase` to establish the connection first.

On error, sends a failure email to the admin address (via `modEmail.SendEmail`) and shows a MsgBox only if not running as the automation user.

**Parameters:**
- `SProcName` — name of the stored procedure to execute (e.g. `"dbo.usp_ProcessAuditData"`)

---

### `ado_ConnectDatabase`

```vb
Public Sub ado_ConnectDatabase(
    Optional ByRef strServer As String = "BI-01",
    Optional ByRef strDBase As String = "PandC_Sandpit"
)
```

Opens an ADODB connection to a SQL Server database using Windows authentication (trusted connection). If `strPwd` is non-blank, falls back to SQL authentication.

Closes an existing open connection before re-opening.

**Parameters:**
- `strServer` — SQL Server hostname or IP (default: `"BI-01"`)
- `strDBase` — database name (default: `"PandC_Sandpit"`)

---

### `acU_PrimKey`

```vb
Public Function acU_PrimKey(tblName As String) As String
```

Returns the primary key field name for a given table by iterating the table's `Indexes` collection. Prints the field name to the Immediate window.

Note: the function signature returns `As String` but the implementation uses `Debug.Print` rather than assigning the return value — the function name is never assigned. This is a known issue in the source (v5.01 comment says "changed to function" but the body wasn't updated).

**Parameters:**
- `tblName` — name of the local or linked table

---

### `acU_RelinkTables`

```vb
Public Sub acU_RelinkTables(
    ByVal OldBasePath As String,
    ByVal NewBasePath As String,
    Optional ByVal AlsoChangeQueries As Boolean = True,
    Optional ByVal AlsoChangeImportExport As Boolean = True
)
```

Relinks all linked tables in the current database by replacing `OldBasePath` with `NewBasePath` in the connection string. Tables where the original path is missing are skipped with a `Debug.Print` entry. Tables where the new path fails are flagged with a critical MsgBox.

Optionally chains to `acU_ChangeQueryPaths` and `acU_Relink_SavedImportExport` for the same replacement.

**Parameters:**
- `OldBasePath` — path fragment to find (e.g. `"\\OldServer\Share\"`)
- `NewBasePath` — replacement path (e.g. `"\\NewServer\Share\"`)
- `AlsoChangeQueries` — also update query SQL (default: `True`)
- `AlsoChangeImportExport` — also update saved import/export specs (default: `True`)

**Usage:**
```vb
acU_RelinkTables "\\OldShare\Data\", "\\NewShare\Data\"
```

---

### `acU_ChangeQueryPaths`

```vb
Public Sub acU_ChangeQueryPaths(
    ByVal OldBasePath As String,
    ByVal NewBasePath As String
)
```

Find-and-replace in the SQL of all `QueryDef` objects in the current database. Warns with a confirmation dialog before proceeding. Handles a known Access bug where bracket doubling (`[[`, `]]`) occurs during replacement.

**Parameters:**
- `OldBasePath` — text to find
- `NewBasePath` — replacement text

**Quirk:** warns if the leading/trailing backslash pattern doesn't match between old and new paths, because asymmetry can break query syntax.

---

### `acU_DB_Path`

```vb
Public Function acU_DB_Path() As String
```

Returns the folder path of the current database (including trailing backslash), equivalent to `ThisWorkbook.Path` in Excel. Parses `CurrentDb.Name` by finding the last backslash.

Only works for local disk or UNC network databases — not for HTTP/web-hosted databases.

---

### `acU_Relink_SavedImportExport`

```vb
Sub acU_Relink_SavedImportExport(
    ByVal OldBasePath As String,
    ByVal NewBasePath As String
)
```

Updates the XML of all saved Import/Export specifications in `CurrentProject.ImportExportSpecifications`, replacing `OldBasePath` with `NewBasePath`. Useful after moving a linked CSV or Excel source file to a new location.

---

### `acU_ImportDirListing`

```vb
Function acU_ImportDirListing(
    strPath As String,
    YourTableName As String,
    YourTableFieldName As String,
    Optional strFilter As String
)
```

Imports the filenames from a folder into a specified table and field using `Dir$` iteration and `INSERT` SQL. Author: CARDA Consultants Inc. (2007-01-19).

**Parameters:**
- `strPath` — folder path, trailing backslash optional (will be appended)
- `YourTableName` — target Access table name
- `YourTableFieldName` — field name to insert filenames into
- `strFilter` — file extension filter (e.g. `"csv"`); use `"*"` for all files (default: `"*"`)

---

### `MyExcelTransfer`

```vb
Public Sub MyExcelTransfer(myTempTable As String, myPath As String)
```

Executes an existing import spec (`myTempTable`) against a new file path (`myPath`). Creates a temporary copy of the spec, replaces the hardcoded path placeholder `"\\MyComputer\ChangeThis"` with the supplied path, executes it, then deletes the temp spec.

Note: the path placeholder `"\\MyComputer\ChangeThis"` must be present in the saved spec's XML — this is a convention to update before use.

---

### `fixImportSpecs`

```vb
Public Sub fixImportSpecs(myTable As String, strFind As String, strRepl As String)
```

Find-and-replace within the XML of a named Import/Export specification. A simpler alternative to `MyExcelTransfer` when you want to permanently update a spec rather than use a temp copy.

---

### `MyExcelChangeName`

```vb
Public Sub MyExcelChangeName(OldName As String, NewName As String)
```

Renames an existing Import/Export specification. Note: there is a bug in the source — `mySpec` is never set to the old spec before deletion, so this will error. Included for completeness.

---

## Dependencies

- `modEmail` — for failure notification email in `ado_ExecStoredProcedure`
- SQL Server accessible via Windows authentication from the Access session's machine account

---

## Known quirks

- `acU_PrimKey` has a return-value bug: the function is declared to return `As String` but never assigns the return value. Only outputs to the Immediate window via `Debug.Print`.
- `MyExcelChangeName` has a bug: `mySpec` is unset before `mySpec.Delete`, causing a runtime error.
- `connDB` is module-level and persists — if the database server becomes unreachable mid-session, callers will see a connection error from `ado_ExecStoredProcedure`.
- `acU_ChangeQueryPaths` fires a MsgBox confirmation on every call, which blocks automation.
