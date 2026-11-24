"""Autenticación compartida para los MCP de Power Platform (MSAL, cliente público).

- Inicio de sesión interactivo en el navegador del sistema (nunca se pide ni guarda contraseña).
- Caché de tokens cifrada con DPAPI en mcp-servers/.tokens/ (solo la descifra este usuario de Windows).
- Respeta HTTPS_PROXY / HTTP_PROXY.

Variables de entorno opcionales:
  MS_TENANT_ID   tenant de Entra (por defecto el de Acme)
  MS_CLIENT_ID   id de aplicación (cliente público) a usar para el inicio de sesión
  MS_LOGIN_HINT  UPN sugerido en la pantalla de inicio de sesión
"""
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

import msal
from msal_extensions import FilePersistenceWithDataProtection, PersistedTokenCache

TENANT_ID = os.environ.get("MS_TENANT_ID", "<ENV_OR_TENANT_ID>")
# Cliente público de Microsoft Azure PowerShell (el que usan los módulos oficiales de Power Apps).
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "1950a258-227b-4e31-a9cf-717495945fc2")
LOGIN_HINT = os.environ.get("MS_LOGIN_HINT", "")
CACHE_FILE = Path(__file__).resolve().parent / ".tokens" / "msal_cache.bin"
TIMEOUT = 60

_app: msal.PublicClientApplication | None = None


def _get_app() -> msal.PublicClientApplication:
    global _app
    if _app is None:
        CACHE_FILE.parent.mkdir(exist_ok=True)
        cache = PersistedTokenCache(FilePersistenceWithDataProtection(str(CACHE_FILE)))
        _app = msal.PublicClientApplication(
            CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}", token_cache=cache)
    return _app


def get_token(scope: str, interactive: bool = True) -> str:
    """Token de acceso para `scope` (p. ej. 'https://service.flow.microsoft.com//.default')."""
    app = _get_app()
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent([scope], account=accounts[0])
    if not result and interactive:
        result = app.acquire_token_interactive([scope], login_hint=LOGIN_HINT or None, timeout=300)
    if not result:
        raise RuntimeError("No hay sesión iniciada. Ejecuta: python _msauth.py login")
    if "access_token" not in result:
        raise RuntimeError(f"Inicio de sesión rechazado: {result.get('error')}: "
                           f"{str(result.get('error_description'))[:600]}")
    return result["access_token"]


def api(method: str, url: str, scope: str, params: dict | None = None, body: Any = None) -> Any:
    """Llamada REST autenticada; devuelve el JSON de la respuesta."""
    if params:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": f"Bearer {get_token(scope, interactive=False)}",
        "Content-Type": "application/json; charset=utf-8",
        "Accept": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {"status": resp.status}
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} en {method} {url.split('?')[0]}: {detail[:1500]}") from None


SCOPES = {
    "flow": "https://service.flow.microsoft.com//.default",
    "powerapps": "https://service.powerapps.com//.default",
}

if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "login":
        for name, scope in SCOPES.items():
            get_token(scope, interactive=True)
            print(f"OK {name}")
    elif cmd == "logout":
        for acc in _get_app().get_accounts():
            _get_app().remove_account(acc)
        print("Sesión cerrada")
    else:
        print([a.get("username") for a in _get_app().get_accounts()] or "Sin sesión")
