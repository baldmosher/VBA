# modEmail — Email Utility Library

**Module:** `modEmail`  
**Version at archive:** v13.17 (2020-11-12)  
**Used by:** nearly every automation module in this codebase

---

## Purpose

The most widely-used module in the codebase. Provides a single interface for sending rich emails from VBA across multiple host environments (desktop Excel, Outlook, Access on a Windows Server VM). Key features:

- **Dual sending path:** Outlook object model as primary, CDO (SMTP) as fallback
- **HTML email support:** with embedded images (`%HTMLIMAGEn%`) and pasted Excel ranges (`%HTMLTABLEn%`)
- **Up to 3 file attachments** with size checking (10 MB limit)
- **VM-aware routing:** detects automation user accounts and switches sending mode automatically
- **Audit trail emails** for tracking report access
- **Link cleanup:** converts `M:\` drive letter links in HTML to UNC `\\server\share\` clickable hyperlinks

---

## Constants

| Constant | Value | Purpose |
|----------|-------|---------|
| `modEmailCodeVer` | `"modEmail v13.17"` | Version string for inclusion in outbound emails |
| `cBlackHoleEmailTo` | `"blackhole@company.com"` | Test destination — absorbs emails without delivery |
| `cIgnoreAuditEmailsForTheseUsers` | pipe-delimited list | Users excluded from `Email_Audit` logging |
| `eMaxAttSize` | `10,000,000` | Max attachment size in bytes (10 MB) |
| `sPathYDrive` | `"\\company.local\Group Data\P&C\"` | UNC for Y: drive (unused after v13.13) |
| `sPathMDrive` | `"\\company.local\Group Data\"` | UNC for M: drive (used in link replacement) |

---

## Functions and Subroutines

### `modEmailVer`

```vb
Function modEmailVer() As String
```

Returns the module version string (`modEmailCodeVer`). Useful for including in debug output or status emails.

---

### `cAdminEmailTo`

```vb
Function cAdminEmailTo() As String
```

Returns the admin notification email address, selected based on `Environ("Username")`. Returns an empty string if the current user is neither `pcopsvm` nor `pcinsightvm` — callers must handle this.

**Requires updating** for each deployment — user accounts are hardcoded.

---

### `cTriggerEmailTo`

```vb
Function cTriggerEmailTo() As String
```

Returns the automation trigger mailbox address for the current user account. Used as the From/ReplyTo address for outbound automation emails.

**Requires updating** for each deployment.

---

### `SendEmail`

```vb
Function SendEmail(
    ByVal Email_Recipient As String,
    Optional ByVal Email_RecipientCC As String,
    Optional ByVal Email_RecipientBCC As String,
    Optional ByVal Email_Subject As String,
    Optional ByVal Email_BodyText As String,
    Optional ByVal DisplayMsg As Boolean = False,
    Optional ByVal AttachmentPath As String,
    Optional ByVal AttachmentPath2 As String,
    Optional ByVal AttachmentPath3 As String,
    Optional ByVal Email_SendOnBehalfOf As String,
    Optional Email_AsHTML As Boolean,
    Optional ByVal Email_CDO As Boolean,
    Optional ByVal Email_ReplyTo As String,
    Optional ByVal BypassVersionNumber As Boolean,
    Optional ByVal BypassSentItems As Boolean = True,
    Optional ByVal ImagePath1 As String,
    ... ImagePath2 through ImagePath6 ...
) As Byte
```

The primary email sending function. Returns: `0` = success, `1` = general send failure, `2` = header failure, `3` = HTML/image failure (non-fatal, continues).

**Sending logic:**

1. If running on Windows Server → force CDO
2. If running as an automation user account → force CDO
3. Otherwise → try Outlook object model; fall back to CDO on failure

**HTML image embedding (`%HTMLIMAGEn%`):**

Place `%HTMLIMAGE1%` through `%HTMLIMAGE6%` in `Email_BodyText`. Supply corresponding `ImagePath1–6` file paths. The image is attached silently (position 0) and the placeholder is replaced with an `<img src="cid:filename">` tag. Use `xlU_Save_Chart_As_Image` to generate chart images.

**HTML table embedding (`%HTMLTABLEn%`):**

Place `%HTMLTABLE1%` through `%HTMLTABLE6%` in `Email_BodyText`. The function looks for Named Ranges `Email_TableHTML1` through `Email_TableHTML6` in the active workbook and calls `RangetoHTML` to convert them. Only works when called from Excel.

**Version footer:** unless `BypassVersionNumber = True`, the module version, machine name, and username are appended to the body.

**Usage example:**
```vb
modEmail.SendEmail "recipient@company.com", , , "Report ready" _
    , "<p>Your report is attached.</p>" _
    , False, "\\server\share\report.xlsx" _
    , , , , True   ' Email_AsHTML = True
```

---

### `Email_via_CDO`

```vb
Function Email_via_CDO(
    ByVal Email_From As String,
    ByVal Email_Sender As String,
    ByVal Email_Recipient As String,
    Optional ByVal Email_RecipientCC As String,
    Optional ByVal Email_RecipientBCC As String,
    Optional ByVal Email_Subject As String,
    Optional ByVal Email_BodyText As String,
    Optional Email_AsHTML As Boolean,
    Optional ByVal AttachmentPath1 As String,
    Optional ByVal AttachmentPath2 As String,
    Optional ByVal AttachmentPath3 As String
) As Byte
```

Sends via Windows CDO (Collaboration Data Objects) using a direct SMTP connection to the internal mail server IP. Returns `0` = success, `1` = send failure, `2` = config failure.

The SMTP server address is hardcoded as `10.x.x.x` (redacted; was the internal mail relay IP). Port 25, no authentication.

Enforces the 10 MB attachment size limit (`eMaxAttSize`) before attaching each file.

---

### `cDefaultEmail`

```vb
Function cDefaultEmail() As String
```

Returns the current user's email address by reading the default MAPI Inbox folder's parent name. This is the mailbox's display name/address rather than `Environ("UserName")`.

Unreliable when called from outside Outlook, or when multiple Outlook sessions are open (always picks the last-opened session's mailbox).

---

### `getOperatingSystem`

```vb
Function getOperatingSystem()
```

Returns the OS caption and version string (e.g. `"Microsoft Windows Server 2016 10.0.14393"`) via a WMI query. Used internally to detect Windows Server environments where Outlook automation isn't available.

Author: Daniel Pineault, CARDA Consultants Inc. (2012).

---

### `RangetoHTML`

```vb
Function RangetoHTML(rng As Object) As String
```

Converts an Excel range to an HTML string suitable for embedding in an email body. Does this by copying the range to a temporary workbook, publishing it as an `.htm` file, reading the file contents, then cleaning up.

Only works from Excel (shows an error MsgBox otherwise). The temp file is stored in `%TEMP%` with a timestamp filename. Passes the result through `fn_replace_report_links` to fix any M:\ drive links.

Adapted from Ron de Bruin's `RangetoHTML` (2006), with bugfixes through 2018.

---

### `fn_replace_report_links`

```vb
Function fn_replace_report_links(ByVal sOriginalCode As String)
```

Converts `M:\path\to\file` links in HTML (which appear as `&quot;M:\path%20to%20file&quot;`) to proper UNC `<a href="\\company.local\Group Data\...">` HTML tags.

Called automatically by `RangetoHTML`. Errors if a link contains spaces (HTML `%20` is handled, but literal spaces in the path are not).

---

### `Email_Audit`

```vb
Sub Email_Audit(Optional ByVal Email_Message As String)
```

Sends a structured audit trail email to the trigger mailbox, recording which user ran which workbook. Skips silently for users listed in `cIgnoreAuditEmailsForTheseUsers` (automation accounts).

The body includes: workbook path, application name/version, 32/64-bit flag, Excel username, Windows username, and the optional `Email_Message` string.

Intended to be called from automation workbooks whenever a user triggers a report refresh.

---

### `Email_Sender_Address`

```vb
Function Email_Sender_Address(ByRef olMail As Object)
```

Extracts the SMTP email address from an Outlook `MailItem`. If the sender type is `"EX"` (Exchange/MAPI), uses `GetExchangeUser().PrimarySmtpAddress` to get the real SMTP address rather than the internal Exchange X.500 address.

Used when processing inbound emails in trigger/audit routines.

---

### `Outlook_Clean_Emails`

```vb
Sub Outlook_Clean_Emails(Optional ByVal AlsoSeparate As Boolean)
```

Interactive utility sub. Reads the active cell's text (expected to be a paste from Outlook's recipient field, in `"Name <email@address.com>; ..."` format) and extracts just the email addresses.

- Default: writes semicolon-separated email string into the cell below
- `AlsoSeparate = True`: splits each address into a separate cell going downward (useful for sorting)

Deduplicates addresses. Intended for manual use in an Excel workbook, not for automation.

---

## Dependencies

- **Outlook object model** (for `SendEmail` primary path, `cDefaultEmail`, `Outlook_Clean_Emails`)
- **CDO library** (`Microsoft CDO for Windows 2000 Library`) for `Email_via_CDO`
- **Excel object model** for `RangetoHTML` and `Email_Audit` (when called from Excel)
- **WMI** (via `CreateObject("winmgmts:...")`) for `getOperatingSystem`
- **WScript.Network** for username detection throughout

---

## Known quirks and gotchas

- **`SendEmail` and server environments:** On Windows Server 2016+, Outlook COM automation is unavailable. The function detects this via `getOperatingSystem` and routes through CDO automatically.
- **CDO SMTP IP:** `Email_via_CDO` uses a hardcoded internal mail server IP. Update this for any new deployment.
- **`cDefaultEmail` unreliable outside Outlook:** When called from Excel or Access, it creates a new Outlook instance — which always returns the last-opened session's mailbox, not necessarily the current user's.
- **`RangetoHTML` and automation:** The HTML publish step (`PublishObjects.Add`) can cause issues in unattended automation (comment in source: "!! this is causing problems in automation"). Also requires `xlReferenceStyle = xlA1` to be set first (v11.25 bugfix).
- **`Email_ReplyTo` auto-correction:** If the reply-to address matches the trigger mailbox, it is automatically replaced with `"P&CInsight@company.com"` — a hardcoded domain-specific behaviour.
- **`BypassSentItems` default is `True`:** Automation emails are deleted after send by default (`.DeleteAfterSubmit = True`). Automation user accounts (`pcinsightvm`, `pcopsvm`) override this and always save to Sent Items.
- **Multiple recipients:** separate with `";"` — the function splits on `";"` and resolves each address individually. Exchange internal display names (without `@`) will also resolve if Outlook is connected.
