"""MCP: consultas SOLO LECTURA a SQL Server (SAP Business One).

Config por variables de entorno (nunca en código):
  SAP_SQL_SERVER, SAP_SQL_DATABASE (en el registro del MCP)
Usuario y contraseña: Administrador de credenciales de Windows (`_sap_sql_env.py set --clipboard`)
  o, como respaldo, SAP_SQL_USER / SAP_SQL_PASSWORD.
  - Usuario `DOMINIO\\usuario` (cuenta de dominio, p. ej. CONSULTORIA\\mesa-lectura-sap): autenticación de Windows
    con esas credenciales (LogonUser NEW_CREDENTIALS + Trusted_Connection).
  - Usuario sin `\\`: login SQL (UID/PWD). Sin usuario: autenticación integrada de la sesión.
"""
import os, re, sys
import pyodbc
from mcp.server.fastmcp import FastMCP

pyodbc.pooling = False   # una conexión reutilizada conservaría las credenciales suplantadas
mcp = FastMCP("sqlserver-sap")
FORBIDDEN = re.compile(
    r"\b(insert|update|delete|drop|alter|create|truncate|merge|exec|execute|grant|revoke|into|xp_\w+|sp_\w+|openrowset|opendatasource|backup|restore|shutdown)\b",
    re.I)
MAX_ROWS = 500


def _credentials():
    """(usuario, contraseña): variables de entorno o Administrador de credenciales (_sap_sql_env.py)."""
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    import _sap_sql_env
    return _sap_sql_env.credentials()


def _conn():
    server, db = os.environ.get("SAP_SQL_SERVER"), os.environ.get("SAP_SQL_DATABASE")
    if not server or not db:
        raise RuntimeError("Faltan SAP_SQL_SERVER / SAP_SQL_DATABASE")
    user, pwd = _credentials()
    base = (f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={db};"
            "ApplicationIntent=ReadOnly;TrustServerCertificate=yes")
    if user and "\\" in user:
        # Cuenta de dominio (DOMINIO\usuario): autenticación de Windows con esas credenciales,
        # como `runas /netonly`: el token solo cambia las credenciales de red del hilo.
        import win32security
        dominio, nombre = user.split("\\", 1)
        token = win32security.LogonUser(nombre, dominio, pwd,
                                        win32security.LOGON32_LOGON_NEW_CREDENTIALS,
                                        win32security.LOGON32_PROVIDER_WINNT50)
        win32security.ImpersonateLoggedOnUser(token)
        try:
            return pyodbc.connect(base + ";Trusted_Connection=yes", timeout=15, readonly=True)
        finally:
            win32security.RevertToSelf()
            token.Close()
    if user:                                   # login SQL clásico
        return pyodbc.connect(base + f";UID={user};PWD={pwd}", timeout=15, readonly=True)
    return pyodbc.connect(base + ";Trusted_Connection=yes", timeout=15, readonly=True)


def _run(sql: str, params=()):
    with _conn() as c:
        cur = c.cursor()
        cur.execute(sql, params)
        cols = [d[0] for d in cur.description]
        rows = cur.fetchmany(MAX_ROWS)
        return cols, [[str(v) if v is not None else None for v in r] for r in rows]


@mcp.tool()
def list_tables(name_filter: str = "") -> list[str]:
    """Tablas de la BD (filtro LIKE opcional, ej. 'OINV')."""
    _, rows = _run("SELECT TABLE_SCHEMA+'.'+TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME LIKE ? ORDER BY 1",
                   (f"%{name_filter}%",))
    return [r[0] for r in rows]


@mcp.tool()
def describe_table(table: str) -> list[dict]:
    """Columnas y tipos de una tabla."""
    cols, rows = _run("SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS "
                      "WHERE TABLE_NAME=? ORDER BY ORDINAL_POSITION", (table.split(".")[-1],))
    return [dict(zip(cols, r)) for r in rows]


@mcp.tool()
def query(sql: str) -> dict:
    """Ejecuta un único SELECT (máx. 500 filas). Cualquier otra sentencia es rechazada."""
    s = sql.strip().rstrip(";")
    if ";" in s or not re.match(r"^\s*(select|with)\b", s, re.I) or FORBIDDEN.search(s):
        raise ValueError("Solo se permite un SELECT de solo lectura")
    cols, rows = _run(s)
    return {"columns": cols, "rows": rows, "truncated": len(rows) >= MAX_ROWS}


if __name__ == "__main__":
    mcp.run()
