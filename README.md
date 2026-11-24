# office-automation-mcp-servers

Small, single-file [FastMCP](https://github.com/modelcontextprotocol/python-sdk) servers used by a help-desk supervisor with Claude Code.

| Server | What it does |
|---|---|
| `outlook-desktop` | Outlook via COM: search/read mail, drafts (new and in-thread), guarded `send_draft`, rules, calendar, address-book lookup, attachments. Strips internal instructions from quoted history before sending. |
| `excel-office`, `word-office` | Read/write tables, **explain workbook logic** (formulas, defined names, `.xlsb` macros), Word -> PDF, PDF -> images. |
| `filesystem-trabajo` | Sandboxed file search/read, tolerant of Unicode-normalized names in OneDrive; reads Teams `.vtt` transcripts. |
| `onedrive-graph` | Read-only OneDrive through Microsoft Graph (MSAL interactive browser sign-in, DPAPI-encrypted token cache; the password is only ever typed in the browser). |
| `power-apps` | List environments, apps, connectors and connections. |
| `sqlserver-sap` | Read-only SQL Server access (SAP Business One) with a least-privilege grant script. |
| `inventory-api` | Read-only client for a distributor inventory API and its key table. |
| `windows-ops` | ping, tracert, nslookup, port test, services, disk, top processes, event log. |

Shared helpers: `_odoo_env.py`, `_sap_sql_env.py`, `_stod_env.py`, ... store/read credentials in the **Windows Credential Manager** (`keyring`), `_msauth*.py` handle MSAL token caches, `_paths.py` the venv paths.

## Español
Colección de MCP chicos para trabajo diario en Windows y Microsoft 365. Ningún secreto vive en archivos: se guardan en el Administrador de credenciales.

> **Note / Nota:** sanitized portfolio version of production work built for a corporate IT help desk. Company names, people, hosts, IDs and data were replaced with placeholders; secrets never lived in the code (they sit in the Windows Credential Manager). / Versión anonimizada de un trabajo real de mesa de ayuda corporativa: nombres, hosts, IDs y datos fueron reemplazados; los secretos nunca estuvieron en el código.
