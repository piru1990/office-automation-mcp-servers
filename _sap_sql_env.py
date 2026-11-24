"""Credencial de la cuenta de dominio de solo lectura para SQL Server de SAP B1 (MCP sqlserver-sap).

Usuario y contraseña viven en el Administrador de credenciales de Windows (credencial genérica
"sap/sql-lectura"), no en ~/.claude.json ni en el código. La contraseña nunca se imprime.

Uso:
  python _sap_sql_env.py set --clipboard [--user DOMINIO\\usuario]
        guarda la contraseña que está en el portapapeles y lo limpia (la pega el supervisor desde el
        correo/PDF de el gerente; así no pasa por el chat). Usuario por defecto: CONSULTORIA\\mesa-lectura-sap
  python _sap_sql_env.py set     pide usuario y contraseña (sin eco)
  python _sap_sql_env.py status [--db BASE] [--server HOST]
        dice si hay credencial y prueba la conexión (SUSER_SNAME, DB_NAME, SELECT en OUSR/USR5)
  python _sap_sql_env.py borrar  elimina la credencial
Seguro de correr con `!` dentro de Claude Code: nunca imprime la contraseña.
"""
import os
import sys

TARGET = "sap/sql-lectura"
DEFAULT_USER = r"CONSULTORIA\mesa-lectura-sap"
DEFAULT_SERVER = "10.0.0.167"
DEFAULT_DB = "SBO_COMPANY"          # producción; en pruebas (166) la base se llama SBO_COMPANY_TEST


def _cred_read() -> tuple[str, str]:
    try:
        import win32cred
        c = win32cred.CredRead(TARGET, win32cred.CRED_TYPE_GENERIC)
        blob = c["CredentialBlob"]
        pwd = blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)
        return c["UserName"], pwd
    except Exception:
        return "", ""


def cred_write(user: str, password: str) -> None:
    import win32cred
    win32cred.CredWrite({
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": TARGET,
        "UserName": user,
        "CredentialBlob": password,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        "Comment": "Cuenta de dominio de solo lectura para SQL Server de SAP B1 (MCP sqlserver-sap)",
    }, 0)


def credentials() -> tuple[str, str]:
    """(usuario, contraseña): variables SAP_SQL_USER / SAP_SQL_PASSWORD o Administrador de credenciales."""
    user, pwd = os.environ.get("SAP_SQL_USER", ""), os.environ.get("SAP_SQL_PASSWORD", "")
    if not (user and pwd):
        user, pwd = _cred_read()
    return user, pwd


def source() -> str:
    if os.environ.get("SAP_SQL_USER") and os.environ.get("SAP_SQL_PASSWORD"):
        return "variables de entorno SAP_SQL_USER / SAP_SQL_PASSWORD"
    if all(_cred_read()):
        return "Administrador de credenciales de Windows"
    return ""


def _portapapeles_leer() -> str:
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    except Exception:
        return ""
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def _portapapeles_limpiar() -> None:
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
    finally:
        win32clipboard.CloseClipboard()


def _arg(name: str, default: str = "") -> str:
    if name in sys.argv:
        i = sys.argv.index(name)
        if i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return default


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "status":
        src = source()
        user, _ = credentials()
        print(f"Credencial SAP SQL: {src or 'NO ENCONTRADA'}" + (f" (usuario {user})" if user else ""))
        if not src:
            return
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "sqlserver-sap"))
        os.environ.setdefault("SAP_SQL_SERVER", _arg("--server", DEFAULT_SERVER))
        os.environ.setdefault("SAP_SQL_DATABASE", _arg("--db", DEFAULT_DB))
        import server
        try:
            cols, rows = server._run(
                "SELECT SUSER_SNAME() AS login, DB_NAME() AS base, "
                "HAS_PERMS_BY_NAME('dbo.OUSR','OBJECT','SELECT') AS select_ousr, "
                "HAS_PERMS_BY_NAME('dbo.USR5','OBJECT','SELECT') AS select_usr5, "
                "HAS_PERMS_BY_NAME('dbo.OUSR','OBJECT','INSERT') AS insert_ousr")
            print(dict(zip(cols, rows[0])))
        except Exception as e:
            print(f"conexión FALLÓ: {e}")
    elif cmd == "set":
        if "--clipboard" in sys.argv:
            user = _arg("--user", DEFAULT_USER)
            pwd = _portapapeles_leer().strip()
            if not pwd or len(pwd) > 128 or "\n" in pwd:
                raise SystemExit("El portapapeles no tiene una contraseña (copie solo la contraseña); no se guardó nada.")
            cred_write(user, pwd)
            _portapapeles_limpiar()
            print(f"Guardada para {user} en el Administrador de credenciales (sap/sql-lectura); portapapeles limpiado.")
        else:
            import getpass
            user = input(f"Usuario [{DEFAULT_USER}]: ").strip() or DEFAULT_USER
            pwd = getpass.getpass("Contraseña (no se muestra): ").strip()
            if not pwd:
                raise SystemExit("Vacía; no se guardó nada.")
            cred_write(user, pwd)
            print("Guardada en el Administrador de credenciales (sap/sql-lectura).")
    elif cmd == "borrar":
        import win32cred
        try:
            win32cred.CredDelete(TARGET, win32cred.CRED_TYPE_GENERIC)
            print("Credencial borrada.")
        except Exception:
            print("No había credencial guardada.")
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main()
