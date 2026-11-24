"""MCP: Power Apps (api.powerapps.com), solo lectura.

Sesión: `python ..\\_msauth.py login` una vez (navegador).
Variable opcional: PP_ENVIRONMENT (por defecto el entorno Default de Acme).
"""
import os
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import _msauth  # noqa: E402

mcp = FastMCP("power-apps")
RO = ToolAnnotations(readOnlyHint=True)

ENV = os.environ.get("PP_ENVIRONMENT", "Default-<ENV_OR_TENANT_ID>")
BASE = "https://api.powerapps.com/providers/Microsoft.PowerApps"
API = {"api-version": "2016-11-01"}
SCOPE = _msauth.SCOPES["powerapps"]


def _env_filter(environment: str) -> dict:
    return {**API, "$filter": f"environment eq '{environment or ENV}'"}


@mcp.tool(annotations=RO)
def list_environments() -> list[dict]:
    """Entornos de Power Platform a los que el usuario tiene acceso."""
    data = _msauth.api("GET", f"{BASE}/environments", SCOPE, API)
    out = []
    for e in data.get("value", []):
        p = e.get("properties", {})
        out.append({"id": e.get("name"), "name": p.get("displayName"), "default": p.get("isDefault"),
                    "sku": p.get("environmentSku"), "region": e.get("location")})
    return out


@mcp.tool(annotations=RO)
def list_apps(environment: str = "") -> list[dict]:
    """Aplicaciones (canvas y model-driven) visibles para el usuario en el entorno."""
    data = _msauth.api("GET", f"{BASE}/apps", SCOPE, _env_filter(environment))
    out = []
    for a in data.get("value", []):
        p = a.get("properties", {})
        out.append({"id": a.get("name"), "name": p.get("displayName"), "type": a.get("appType") or p.get("appType"),
                    "owner": (p.get("owner") or {}).get("displayName"),
                    "modified": p.get("lastModifiedTime"), "play_url": p.get("appPlayUri") or p.get("appOpenUri")})
    return out


@mcp.tool(annotations=RO)
def get_app(app_id: str) -> dict:
    """Detalle de una aplicación: propietario, conexiones que usa, versión, URLs."""
    return _msauth.api("GET", f"{BASE}/apps/{app_id}", SCOPE, API)


@mcp.tool(annotations=RO)
def list_connectors(name_filter: str = "", environment: str = "") -> list[dict]:
    """Conectores disponibles en el entorno (nombre, nivel Standard/Premium)."""
    data = _msauth.api("GET", f"{BASE}/apis", SCOPE, _env_filter(environment))
    out = []
    for c in data.get("value", []):
        p = c.get("properties", {})
        if name_filter.lower() in ((c.get("name") or "") + (p.get("displayName") or "")).lower():
            out.append({"id": c.get("name"), "name": p.get("displayName"), "tier": p.get("tier"),
                        "custom": p.get("isCustomApi")})
    return out


@mcp.tool(annotations=RO)
def list_connections(environment: str = "") -> list[dict]:
    """Conexiones creadas por el usuario en el entorno."""
    data = _msauth.api("GET", f"{BASE}/connections", SCOPE, _env_filter(environment))
    out = []
    for c in data.get("value", []):
        p = c.get("properties", {})
        out.append({"name": c.get("name"), "api": p.get("apiId", "").rsplit("/", 1)[-1],
                    "display": p.get("displayName"),
                    "status": [s.get("status") for s in p.get("statuses", [])]})
    return out


if __name__ == "__main__":
    mcp.run()
