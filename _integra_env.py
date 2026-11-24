"""Credenciales del portal de soporte de ERP-PARTNER (osTicket) para el MCP integra-soporte.

Usuario y contraseña viven en el Administrador de credenciales de Windows (credencial
genérica "integra/portal"), no en ~/.claude.json ni en el código. Nunca se imprimen.

Orden de búsqueda: variables ERP_PARTNER_USER / ERP_PARTNER_PASS -> Administrador de credenciales.

Uso (en una ventana de PowerShell propia, porque getpass no funciona dentro de Claude Code):
  python _integra_env.py status   dice si hay credencial y prueba el ingreso
  python _integra_env.py set      pide usuario y contraseña (sin eco) y los guarda
"""
import os
import sys
from pathlib import Path

TARGET = "integra/portal"
DEFAULT_URL = "https://support.erp-partner.example.com"


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
        "Comment": "Usuario y contraseña del portal ERP-PARTNER para el MCP integra-soporte",
    }, 0)


def credentials() -> tuple[str, str, str]:
    """(url, usuario, contraseña). Lanza RuntimeError si no hay credencial."""
    url = os.environ.get("ERP_PARTNER_URL") or DEFAULT_URL
    user, pwd = os.environ.get("ERP_PARTNER_USER", ""), os.environ.get("ERP_PARTNER_PASS", "")
    if not (user and pwd):
        user, pwd = _cred_read()
    if not (user and pwd):
        raise RuntimeError(f"No hay credencial de ERP-PARTNER. Guárdela con: python {Path(__file__)} set")
    return url.rstrip("/"), user, pwd


def source() -> str:
    if os.environ.get("ERP_PARTNER_USER") and os.environ.get("ERP_PARTNER_PASS"):
        return "variables de entorno ERP_PARTNER_USER / ERP_PARTNER_PASS"
    if all(_cred_read()):
        return "Administrador de credenciales de Windows"
    return ""


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "status":
        src = source()
        print(f"Credencial ERP-PARTNER: {src or 'NO ENCONTRADA'}")
        if src:
            sys.path.insert(0, str(Path(__file__).parent / "integra-soporte"))
            import server
            try:
                print(server._test())
            except Exception as e:
                print(f"ingreso FALLÓ: {e}")
    elif cmd == "set":
        import getpass
        user = input("Usuario de ERP-PARTNER (correo con el que ingresa al portal): ").strip()
        pwd = getpass.getpass("Contraseña de ERP-PARTNER (no se muestra): ").strip()
        if not user or not pwd:
            raise SystemExit("Faltan datos; no se guardó nada.")
        cred_write(user, pwd)
        print("Guardada en el Administrador de credenciales.")
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main()
