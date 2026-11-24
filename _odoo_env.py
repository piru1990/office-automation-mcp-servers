"""Credenciales de Odoo para el MCP odoo-helpdesk y sus scripts.

La API key vive en el Administrador de credenciales de Windows (credencial genérica
"odoo-helpdesk/api-key"), no en ~/.claude.json. Nunca se imprime.

Orden de búsqueda: variable de entorno ODOO_API_KEY -> Administrador de credenciales
-> (legado) registro del MCP en ~/.claude.json.

Uso:
  python _odoo_env.py status            dice dónde está la key y prueba la conexión
  python _odoo_env.py set               pide la key (sin eco) y la guarda
  python _odoo_env.py set --clipboard   la toma del portapapeles y lo limpia
  python _odoo_env.py migrate           copia la key de ~/.claude.json al Administrador de credenciales
"""
import json
import os
import sys
from pathlib import Path

TARGET = "odoo-helpdesk/api-key"
DEFAULT_URL = "https://example.odoo.com"


def _cred_read() -> str:
    try:
        import win32cred
        blob = win32cred.CredRead(TARGET, win32cred.CRED_TYPE_GENERIC)["CredentialBlob"]
        return blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)
    except Exception:
        return ""


def _cred_write(key: str) -> None:
    import win32cred
    win32cred.CredWrite({
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": TARGET,
        "UserName": "odoo-helpdesk",
        "CredentialBlob": key,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        "Comment": "API key de Odoo para el MCP odoo-helpdesk",
    }, 0)


def _legacy() -> tuple[str, str]:
    """(url, key) del registro del MCP en ~/.claude.json, si todavía está ahí."""
    try:
        cfg = json.loads((Path.home() / ".claude.json").read_text(encoding="utf-8"))
    except Exception:
        return "", ""
    servers = [cfg.get("mcpServers", {})] + [p.get("mcpServers", {}) for p in cfg.get("projects", {}).values()]
    for s in servers:
        env = (s.get("odoo-helpdesk") or {}).get("env") or {}
        if env.get("ODOO_API_KEY"):
            return env.get("ODOO_URL", ""), env["ODOO_API_KEY"]
    return "", ""


def source() -> str:
    if os.environ.get("ODOO_API_KEY"):
        return "variable de entorno ODOO_API_KEY"
    if _cred_read():
        return "Administrador de credenciales de Windows"
    if _legacy()[1]:
        return "~/.claude.json (legado, en claro)"
    return ""


def credentials() -> tuple[str, str]:
    """(url, api_key). Lanza RuntimeError si no hay key."""
    url = os.environ.get("ODOO_URL") or DEFAULT_URL
    key = os.environ.get("ODOO_API_KEY") or _cred_read()
    if not key:
        legacy_url, key = _legacy()
        url = os.environ.get("ODOO_URL") or legacy_url or DEFAULT_URL
    if not key:
        raise RuntimeError("No hay API key de Odoo. Guárdela con: python "
                           f"{Path(__file__)} set")
    return url, key


def load_env() -> None:
    """Deja ODOO_URL y ODOO_API_KEY en os.environ (para los scripts que importan server.py)."""
    try:
        url, key = credentials()
    except RuntimeError as e:
        raise SystemExit(str(e))
    os.environ["ODOO_URL"], os.environ["ODOO_API_KEY"] = url, key


def _test() -> str:
    sys.path.insert(0, str(Path(__file__).parent / "odoo-helpdesk"))
    load_env()
    import server
    ctx = server._call("res.users", "context_get")
    return f"conexión OK (uid {ctx.get('uid')})"


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "status":
        src = source()
        print(f"API key: {src or 'NO ENCONTRADA'}")
        if src:
            try:
                print(_test())
            except Exception as e:
                print(f"conexión FALLÓ: {e}")
    elif cmd == "set":
        if "--clipboard" in sys.argv:
            import win32clipboard
            win32clipboard.OpenClipboard()
            try:
                key = str(win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)).strip()
                win32clipboard.EmptyClipboard()
            finally:
                win32clipboard.CloseClipboard()
        else:
            import getpass
            key = getpass.getpass("API key de Odoo (no se muestra): ").strip()
        if len(key) < 20 or " " in key:
            raise SystemExit("Eso no parece una API key; no se guardó nada.")
        _cred_write(key)
        os.environ.pop("ODOO_API_KEY", None)
        print("Guardada en el Administrador de credenciales.")
        try:
            print(_test())
        except Exception as e:
            print(f"conexión FALLÓ: {e}")
    elif cmd == "migrate":
        if _cred_read():
            print("Ya hay una key en el Administrador de credenciales; no se tocó.")
            return
        key = _legacy()[1]
        if not key:
            raise SystemExit("No hay key en ~/.claude.json que migrar.")
        _cred_write(key)
        print("Copiada al Administrador de credenciales. Falta quitarla de ~/.claude.json.")
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main()
