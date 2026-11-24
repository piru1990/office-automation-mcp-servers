"""MCP: OneDrive de ex-técnico (user@example.com) por Microsoft Graph. SOLO LECTURA.

Sesión: `python ..\\_msauth_graph.py login` una vez (navegador; el supervisor escribe la contraseña de el ex-técnico).
No hay ninguna herramienta que escriba, mueva, comparta o borre.
"""
import re
import sys
import tempfile
import urllib.parse
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import _msauth_graph as ms  # noqa: E402

mcp = FastMCP("onedrive-graph")
RO = ToolAnnotations(readOnlyHint=True)

FIELDS = "id,name,size,lastModifiedDateTime,folder,file,webUrl,parentReference"
TEXT_EXT = {".txt", ".csv", ".md", ".json", ".xml", ".eml", ".log", ".sql", ".html", ".htm", ".tsv"}
MAX_TEXT_BYTES = 5 * 1024 * 1024


def _item(i: dict) -> dict:
    parent = (i.get("parentReference") or {}).get("path", "")
    return {
        "id": i.get("id"), "name": i.get("name"),
        "type": "carpeta" if "folder" in i else "archivo",
        "size": i.get("size"), "modified": i.get("lastModifiedDateTime"),
        "path": (parent.split("root:", 1)[-1] or "/") if parent else "",
        "hijos": (i.get("folder") or {}).get("childCount"),
        "url": i.get("webUrl"),
    }


def _ref(item_id: str = "", path: str = "") -> str:
    """Segmento de Graph para un elemento por id o por ruta relativa a la raíz del OneDrive."""
    if item_id:
        if not re.fullmatch(r"[A-Za-z0-9!_\-]+", item_id):
            raise ValueError("id de elemento no válido")
        return f"/me/drive/items/{item_id}"
    p = path.strip().replace("\\", "/").strip("/")
    if ".." in p.split("/"):
        raise ValueError("Ruta no válida")
    return "/me/drive/root" if not p else "/me/drive/root:/" + urllib.parse.quote(p) + ":"


def _all(url: str, params: dict, top: int) -> list[dict]:
    out: list[dict] = []
    data = ms.api(url, params)
    while True:
        out += data.get("value", [])
        nxt = data.get("@odata.nextLink")
        if not nxt or len(out) >= top:
            return out[:top]
        data = ms.api(nxt)


@mcp.tool(annotations=RO)
def whoami() -> dict:
    """Cuenta con la que está iniciada la sesión y tamaño del OneDrive (comprueba que sea extech@)."""
    me = ms.api("/me", {"$select": "displayName,userPrincipalName,accountEnabled"})
    try:
        q = ms.api("/me/drive", {"$select": "id,driveType,quota"})
        drive = {"tipo": q.get("driveType"), "usado": (q.get("quota") or {}).get("used"),
                 "total": (q.get("quota") or {}).get("total")}
    except RuntimeError as e:
        drive = {"error": str(e)[:300]}
    return {"usuario": me.get("displayName"), "upn": me.get("userPrincipalName"), "drive": drive}


@mcp.tool(annotations=RO)
def list_dir(path: str = "", top: int = 200) -> list[dict]:
    """Contenido de una carpeta del OneDrive (ruta relativa a la raíz; vacío = raíz). Más reciente primero."""
    rows = _all(f"{_ref(path=path)}/children",{"$select": FIELDS, "$top": 200, "$orderby": "lastModifiedDateTime desc"}, top)
    return [_item(r) for r in rows]


@mcp.tool(annotations=RO)
def search(query: str, top: int = 25) -> list[dict]:
    """Busca en el OneDrive por nombre y contenido (p. ej. 'leyenda', 'nota de crédito', 'Ainnova')."""
    q = query.replace("'", "''")
    rows = _all(f"/me/drive/root/search(q='{urllib.parse.quote(q)}')", {"$select": FIELDS, "$top": 50}, top)
    return [_item(r) for r in rows]


@mcp.tool(annotations=RO)
def recent(top: int = 25) -> list[dict]:
    """Archivos usados o modificados recientemente."""
    rows = _all("/me/drive/recent", {"$top": 50}, top)
    return [_item(r) for r in rows]


@mcp.tool(annotations=RO)
def download(item_id: str = "", path: str = "", dest_dir: str = "") -> dict:
    """Descarga un archivo a `dest_dir` (por defecto una carpeta temporal propia) sin sobrescribir.
    Luego léelo con excel-office, word-office o Read. No lo ejecuta ni lo abre."""
    meta = ms.api(_ref(item_id, path))  # sin $select: así sí trae @microsoft.graph.downloadUrl
    if "folder" in meta:
        raise ValueError("Es una carpeta; usa list_dir")
    url = meta.get("@microsoft.graph.downloadUrl")
    if not url:
        raise RuntimeError("Graph no devolvió URL de descarga para ese elemento")
    out_dir = Path(dest_dir) if dest_dir else Path(tempfile.gettempdir()) / "onedrive-graph"
    out_dir.mkdir(parents=True, exist_ok=True)
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", meta.get("name") or "archivo")
    dest = out_dir / name
    k = 1
    while dest.exists():
        dest = out_dir / f"{dest.stem} ({k}){dest.suffix}"
        k += 1
    n = ms.download_url(url, dest)
    return {"ruta": str(dest), "bytes": n, "nombre": meta.get("name")}


@mcp.tool(annotations=RO)
def read_text(item_id: str = "", path: str = "", max_chars: int = 20000, offset: int = 0) -> dict:
    """Texto de un archivo de texto plano (txt, csv, md, json, xml, eml, html, sql…), hasta 5 MB.
    Para Word/Excel/PDF usa `download` y las herramientas de oficina."""
    meta = ms.api(_ref(item_id, path))  # sin $select: así sí trae @microsoft.graph.downloadUrl
    name = meta.get("name") or ""
    if "folder" in meta:
        raise ValueError("Es una carpeta; usa list_dir")
    if Path(name).suffix.lower() not in TEXT_EXT:
        raise ValueError(f"{name}: no es texto plano; usa download y excel-office/word-office/Read")
    if (meta.get("size") or 0) > MAX_TEXT_BYTES:
        raise ValueError("Archivo mayor de 5 MB; usa download")
    with tempfile.TemporaryDirectory() as tmp:
        dest = Path(tmp) / "f"
        ms.download_url(meta["@microsoft.graph.downloadUrl"], dest)
        raw = dest.read_bytes()
    for enc in ("utf-8-sig", "cp1252"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    return {"nombre": name, "total_chars": len(text), "offset": offset,
            "texto": text[offset:offset + max_chars]}


if __name__ == "__main__":
    mcp.run()
