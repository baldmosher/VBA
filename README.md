# VBA Legacy Automation Archive

A reference archive of Excel, Outlook, and Access VBA automation code written circa 2012–2018. Original workbooks and databases are not included — this is source code only, preserved for reference.

## Structure

| Folder | Contents |
|--------|----------|
| `access/` | Access automation modules (`acc*`) |
| `excel/` | Excel automation modules (`xl*`) |
| `outlook/` | Outlook automation modules, class modules, and session files (`ol*`, `*.cls`) |
| `modules/` | Shared utility modules used across applications (`mod*`) |
| `misc/` | Supporting scripts and data files |

## Notes

- Files are `.txt` exports of VBA modules (`.bas`/`.cls` originals)
- Sensitive values (email addresses, hostnames, IP addresses) have been redacted and replaced with `@company.com`, `company.local`, `10.x.x.x` placeholders
- Three `ThisOutlookSession` variants reflect deployment on different virtual machines
