"""Sesión de Microsoft Graph para el OneDrive de ex-técnico (user@example.com), solo lectura.

- Inicio de sesión interactivo en el navegador del sistema: la contraseña y el segundo factor los
  escribe el supervisor en la ventana de Microsoft. Nunca se piden por chat ni se guardan.
- Caché de tokens PROPIA (cifrada con DPAPI) en mcp-servers/.tokens/msal_cache_extech.bin; no se
  mezcla con la sesión de Power Platform de el supervisor (`_msauth.py`).
- Respeta HTTPS_PROXY / HTTP_PROXY (urllib las lee del entorno).

Uso (en una terminal propia; con `!` dentro de Claude Code):
  python _msauth_graph.py login [graph|office]   inicia sesión (por defecto 'graph')
  python _msauth_graph.py status                  muestra la cuenta de la sesión
  python _msauth_graph.py logout                  borra la sesión y la caché

Variables opcionales: MS_TENANT_ID, MS_LOGIN_HINT_LPRADO.
"""
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

import msal
from msal_extensions import FilePersistenceWithDataProtection, PersistedTokenCache

TENANT_ID = os.environ.get("MS_TENANT_ID", "<ENV_OR_TENANT_ID>")
LOGIN_HINT = os.environ.get("MS_LOGIN_HINT_LPRADO", "user@example.com")
# Clientes públicos de Microsoft. Si el tenant bloquea el primero, `login office`.
CLIENTS = {
    "graph": "14d82eec-204b-4c2f-b7e8-296a70dab67e",   # Microsoft Graph Command Line Tools
    "office": "d3590ed6-52b3-4102-aeff-aad2292ab01c",  # Microsoft Office
}
# Solo lectura. `Files.Read` (archivos propios) alcanza para el OneDrive de la cuenta.
SCOPES = ["https://graph.microsoft.com/Files.Read", "https://graph.microsoft.com/User.Read"]
GRAPH = "https://graph.microsoft.com/v1.0"

BASE = Path(__file__).resolve().parent / ".tokens"
CACHE_FILE = BASE / "msal_cache_extech.bin"
CLIENT_FILE = BASE / "extech_client.txt"
TIMEOUT = 60

_app: msal.PublicClientApplication | None = None


def _client_name() -> str:
    try:
        name = CLIENT_FILE.read_text(encoding="utf-8").strip()
        if name in CLIENTS:
            return name
    except OSError:
        pass
    return "graph"


def _get_app(client: str | None = None) -> msal.PublicClientApplication:
    global _app
    if client:
        BASE.mkdir(exist_ok=True)
        CLIENT_FILE.write_text(client, encoding="utf-8")
        _app = None
    if _app is None:
        BASE.mkdir(exist_ok=True)
        cache = PersistedTokenCache(FilePersistenceWithDataProtection(str(CACHE_FILE)))
        _app = msal.PublicClientApplication(
            CLIENTS[client or _client_name()],
            authority=f"https://login.microsoftonline.com/{TENANT_ID}", token_cache=cache)
    return _app


def _check_account(result: dict) -> None:
    """Aborta si el inicio de sesión no fue con la cuenta de el ex-técnico (evita leer el OneDrive de otra)."""
    claims = result.get("id_token_claims") or {}
    upn = (claims.get("preferred_username") or claims.get("upn") or "").lower()
    if upn and upn != LOGIN_HINT.lower():
        _get_app().remove_account(_get_app().get_accounts()[0]) if _get_app().get_accounts() else None
        raise RuntimeError(f"La sesión es de {upn}, no de {LOGIN_HINT}. Se descartó; repite el login con la cuenta correcta.")


def get_token(interactive: bool = False) -> str:
    app = _get_app()
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result and interactive:
        result = app.acquire_token_interactive(SCOPES, login_hint=LOGIN_HINT, timeout=300,
                                               prompt="select_account")
        if result and "access_token" in result:
            _check_account(result)
    if not result:
        raise RuntimeError("No hay sesión de extech@. Ejecuta: python _msauth_graph.py login")
    if "access_token" not in result:
        raise RuntimeError(f"Inicio de sesión rechazado: {result.get('error')}: "
                           f"{str(result.get('error_description'))[:600]}")
    return result["access_token"]


def api(path_or_url: str, params: dict | None = None) -> Any:
    """GET autenticado a Microsoft Graph (solo lectura); devuelve el JSON."""
    url = path_or_url if path_or_url.startswith("https://") else GRAPH + path_or_url
    if params:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, method="GET", headers={
        "Authorization": f"Bearer {get_token()}", "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} en GET {url.split('?')[0]}: {detail[:1200]}") from None


def download_url(url: str, dest: Path) -> int:
    """Baja una URL prefirmada de Graph (@microsoft.graph.downloadUrl, sin Authorization) a `dest`."""
    if not url.startswith("https://"):
        raise RuntimeError("URL de descarga no válida")
    req = urllib.request.Request(url, method="GET")
    n = 0
    with urllib.request.urlopen(req, timeout=300) as resp, open(dest, "xb") as f:
        while chunk := resp.read(1 << 20):
            f.write(chunk)
            n += len(chunk)
    return n


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "login":
        client = sys.argv[2] if len(sys.argv) > 2 else "graph"
        if client not in CLIENTS:
            sys.exit(f"Cliente desconocido: {client}. Usa: {', '.join(CLIENTS)}")
        _get_app(client)
        get_token(interactive=True)
        me = api("/me", {"$select": "displayName,userPrincipalName"})
        print(f"OK ({client}): {me.get('displayName')} <{me.get('userPrincipalName')}>")
    elif cmd == "logout":
        for acc in _get_app().get_accounts():
            _get_app().remove_account(acc)
        for f in (CACHE_FILE, CLIENT_FILE):
            f.unlink(missing_ok=True)
        print("Sesión cerrada y caché borrada")
    else:
        print([a.get("username") for a in _get_app().get_accounts()] or "Sin sesión")
