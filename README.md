# VBA Legacy Automation Archive

A reference archive of Excel, Outlook, and Access VBA automation code written circa 2012–2020. Original workbooks and databases are not included — this is source code only, preserved for reference.

---

**baldmosher™ disclaimer:** all code was lifted, stolen, and adapted from the internet. Stack Overflow, forums, etc. Eventually I realised I knew more than the internet, so there's some hefty development in here, but I claim no copyright, nor do I recognise any of this as my IP, and I leave this legacy behind for anyone to utilise, fully open source / free use licence, do whatever you want with it, coding is fun, figure stuff out! But it's 2026... I'm pretty sure VBA is effectively dead now, and coding is almost automated, but I am fond of it, and I still use it sometimes. My top tip: don't use Copilot for creating VBA. Always happy to guide people (and I am available for hire) so do get in touch! tinyurl.com/baldmosherTipJar

---

Sensitive values (email addresses, hostnames, IP addresses, personal names) have been redacted and replaced with placeholders.

## Structure

| Folder | Contents |
|--------|----------|
| `access/` | Access automation modules (`acc*`) |
| `excel/` | Excel automation modules (`xl*`) |
| `outlook/` | Outlook automation modules, class modules, and session files (`ol*`, `*.cls`) |
| `modules/` | Shared utility modules used across applications (`mod*`) |
| `misc/` | Supporting scripts and data files |

---

## Access

| File | Description |
|------|-------------|
| `accAuto.txt` | AutoExec automation controller for the pcinsightvm Access database; schedules and runs overnight routines |
| `accAuto_pcopsvm.txt` | Variant of accAuto configured for the pcopsvm virtual machine |
| `accCommercialAudits.txt` | Loads and processes JIRA/CSV data via the Commercial Audits Access database; executes stored procedures on SQL Server |
| `accCommercialCompliance.txt` | Loads compliance data via the Commercial Compliance Access database |
| `accReportTrail.txt` | Appends and imports TXT log files into a Report Trail Access backend database |
| `accRunCode.txt` | Utility runner; truncates staging tables in SQL Server via stored procedures |
| `accUtils.txt` | Shared Access utility library; database path helpers, relinking, query path updates, stored procedure execution |

## Excel

| File | Description |
|------|-------------|
| `xlAmendNames.txt` | Renames Named Ranges and sheet tabs by adding or replacing text |
| `xlArchive.txt` | Archives report outputs to a network share with versioning |
| `xlAudit.txt` | Manages audit trail TXT log files; maps drive letters to UNC paths |
| `xlAuditLogger.txt` | Updates SQL compliance tables with exception data from Excel user submissions |
| `xlBackgroundRefresh.txt` | Refreshes all data connections as background queries, waits for completion, then updates pivots and recalculates |
| `xlFormatter.txt` | Applies standard named number formats to selected cells |
| `xlFormulaConvert.txt` | Converts IFERROR formulas to IF(blank-check) equivalents for PowerPivot compatibility |
| `xlPCIFormat.txt` | Applies standard report header, border, and column width formatting |
| `xlPDFprint.txt` | Deprecated — superseded by xlScratchpad |
| `xlPivots.txt` | PivotTable and PowerPivot utilities; refresh control and calculation-complete detection |
| `xlRCX.txt` | Extends tracking rows in an exposure tracker workbook on each save |
| `xlScratchpad.txt` | Creates a clean PDF-ready copy of a report sheet; strips hidden rows/columns and applies print settings |
| `xlSelfServeReport.txt` | Self-serve One Way Fees report triggered by email; manages pivot connections and output |
| `xlSelfServeReservations.txt` | Self-serve reservation data report; email-triggered with GDPR password protection for PII output |
| `xlSimple.txt` | Lightweight report refresh controller for simpler non-VM automated reports |
| `xlSprocRefresh.txt` | Triggers SQL stored procedure execution from Excel with timeout handling and failure email |
| `xlSwapColumnRefforTableField.txt` | Find-and-replace in formulas to swap column references for structured table field references |
| `xlTradingSummary.txt` | Filter and button controls for a Trading Summary report (Pay Now/Local, NA Domestic, PNPL toggles) |
| `xlUtils.txt` | Core Excel utility library (reached v37); path handling, connection management, chart export, MySQL driver detection |
| `xlVBAUpdater.txt` | Scans workbooks on a network share and replaces VBA modules from master copies; manages path migrations |

## Outlook

| File | Description |
|------|-------------|
| `olArchive.txt` | Early prototype: moves items between shared mailbox folders |
| `olArchiver.txt` | Moves emails between MAPI folders across mailboxes; basis for archiving routines |
| `olAttachmentProcess.txt` | Saves email attachments from Outlook to a network share and imports to SQL |
| `olAuditCops.txt` | Logs COps team emails (PriceMatch, ProductQuery, StopSale mailboxes) to a SQL audit database |
| `olAuditProcess.txt` | Processes COps audit emails and writes structured data to SQL via pattern matching on email body text |
| `olReminders.txt` | Manages scheduled Outlook reminder appointments used to trigger automated archive routines |
| `olSaveAttachtoDisk.txt` | Saves email attachments to disk and calls SQL import routine (earlier/simpler version of olAttachmentProcess) |
| `olTriggers.txt` | Main email-triggered report engine; receives formatted request emails, opens Excel, runs reports, replies with output (v5.31) |
| `olTriggers - Copy.txt` | Variant of olTriggers for a different virtual machine (forked at v5.21) |
| `olWorkflow.txt` | Stub/prototype for a supplier email auto-filing workflow (incomplete) |
| `ThisOutlookSession.cls` | Application-level Outlook session handler; wires up reminder and receive events to trigger archive and report automation (VM7) |
| `ThisOutlookSession-PCInsightVM.cls` | ThisOutlookSession variant for PCInsightVM (VM2) |
| `ThisOutlookSession-pcopsvm.cls` | ThisOutlookSession variant for pcopsvm; uses scheduled reminders to trigger routines |

## Modules

| File | Description |
|------|-------------|
| `modAppsOffice.txt` | Launches and controls external Excel and Access application instances from VBA |
| `modAuto.txt` | Automation orchestrator; coordinates the full overnight report run sequence across modEmail, modFile, xlAudit, xlBackgroundRefresh, xlPivots, xlUtils |
| `modEmail.txt` | Comprehensive email utility library (reached v13); Outlook object model and CDO fallback, HTML formatting, attachment handling, UNC link conversion |
| `modExportExceltoSQL.txt` | Exports Google Analytics data from an Excel sheet to SQL Server via ADODB |
| `modExtra.txt` | Stub extension point; placeholder subs for custom pre/post automation hooks |
| `modFile.txt` | File system utility library; file age checks, newest-file finder, read-only testing, deletion with age filter |
| `modRecursiveFolders.txt` | Recurses a folder tree and processes Excel files found within |
| `modSleep.txt` | Sleep API declaration and file-size-proportional delay helper |
| `modValidationSummary.txt` | Processes a Report Validation Summary workbook; validates network path and corrects relative hyperlinks |
| `modVBAUpdater.txt` | Programmatically replaces individual VBA modules within open workbooks |
| `modZip.txt` | ZIP file creation and extraction utility; wraps Windows Shell and a third-party zip library |

## Misc

| File | Description |
|------|-------------|
| `Rate Load times.txt` | Tab-separated reference data: rate loading time estimates by destination and supplier |
| `Remap Network Drives.txt` | VBScript snippet using WMI to list current network drive mappings |
