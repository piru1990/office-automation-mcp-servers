"""MCP: acceso acotado (solo lectura) a Documents, Downloads, Escritorio, OneDrive - Acme y mcp-servers.

Tolera nombres con tildes escritas de otra forma (ver ..\\_paths.py).
"""
import re
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import _paths  # noqa: E402

HOME = Path.home()
ROOTS = [(HOME / n).resolve() for n in ("Documents", "Downloads", "Desktop", "OneDrive - Acme, S.A",
                                        "OneDrive - Acme, S.A/Escritorio", "mcp-servers")]
# dentro de mcp-servers no se exponen tokens, el perfil del navegador ni el entorno virtual
EXCLUIDOS = {".tokens", "edge-profile", ".venv", "__pycache__"}
MAX_READ = 200_000
mcp = FastMCP("filesystem-trabajo")
RO = ToolAnnotations(readOnlyHint=True)


def _safe(p: str) -> Path:
    path = _paths.resolve(p).resolve()
    if not any(path == r or r in path.parents for r in ROOTS):
        raise ValueError(f"Ruta fuera de las carpetas permitidas: {[str(r) for r in ROOTS]}")
    if EXCLUIDOS & set(path.parts):
        raise ValueError(f"Carpeta excluida: {sorted(EXCLUIDOS & set(path.parts))}")
    return path


@mcp.tool(annotations=RO)
def allowed_roots() -> list[str]:
    """Carpetas raíz permitidas."""
    return [str(r) for r in ROOTS if r.exists()]


@mcp.tool(annotations=RO)
def list_dir(path: str) -> list[dict]:
    """Lista el contenido de un directorio permitido."""
    d = _safe(path)
    out = []
    for e in sorted(d.iterdir(), key=lambda x: (not x.is_dir(), x.name.lower()))[:500]:
        if e.name in EXCLUIDOS:
            continue
        try:
            st = e.stat()
            out.append({"name": e.name, "path": str(e), "dir": e.is_dir(), "size": st.st_size, "modified": st.st_mtime})
        except OSError:
            pass
    return out


@mcp.tool(annotations=RO)
def read_text_file(path: str, max_chars: int = 50_000, offset: int = 0) -> str:
    """Lee un archivo de texto (UTF-8, tolerante a errores) desde el carácter `offset`, hasta max_chars.
    Para archivos largos, llamar de nuevo con offset += max_chars. Para .vtt/.srt use read_transcript."""
    f = _safe(path)
    return f.read_text(encoding="utf-8", errors="replace")[offset: offset + min(max_chars, MAX_READ)]


RE_VOZ = re.compile(r"<v ([^>]+)>(.*?)</v>", re.S)
RE_ID_CUE = re.compile(r"^(\d+|[0-9a-f]{8}-[0-9a-f-]{20,}(/\S+)?)$", re.I)


def _transcripcion(crudo: str) -> list[tuple[str, str]]:
    """Texto de un .vtt/.srt como lista (hablante, texto), uniendo los turnos seguidos del mismo hablante."""
    lineas = [ln.strip() for ln in crudo.splitlines()]
    lineas = [ln for ln in lineas if ln and "-->" not in ln and not ln.startswith(("WEBVTT", "NOTE"))
              and not RE_ID_CUE.match(ln)]
    cuerpo = " ".join(lineas)
    turnos = [(h.strip(), " ".join(t.split())) for h, t in RE_VOZ.findall(cuerpo)]
    if not turnos:  # sin etiquetas de hablante
        return [("", re.sub(r"<[^>]+>", "", cuerpo))]
    out: list[tuple[str, str]] = []
    for h, t in turnos:
        if not t:
            continue
        if out and out[-1][0] == h:
            out[-1] = (h, out[-1][1] + " " + t)
        else:
            out.append((h, t))
    return out


@mcp.tool(annotations=RO)
def read_transcript(path: str, offset: int = 0, max_chars: int = 30_000) -> dict:
    """Transcripción limpia de un .vtt o .srt (p. ej. grabaciones de Teams): sin marcas de tiempo ni
    identificadores, un renglón por turno "[Hablante] texto". Devuelve el tramo desde `offset`, el total de
    caracteres y `siguiente_offset` (None si ya terminó) para seguir leyendo."""
    turnos = _transcripcion(_safe(path).read_text(encoding="utf-8-sig", errors="replace"))
    hablantes = sorted({h for h, _ in turnos if h})
    cortos = {h: f"H{i + 1}" for i, h in enumerate(hablantes)}
    texto = "\n".join(f"[{cortos[h]}] {t}" if h else t for h, t in turnos)
    fin = offset + max(1, min(max_chars, MAX_READ))
    return {"hablantes": {v: k for k, v in cortos.items()}, "total_chars": len(texto), "offset": offset,
            "siguiente_offset": fin if fin < len(texto) else None, "texto": texto[offset:fin]}


@mcp.tool(annotations=RO)
def search_files(root: str, pattern: str, max_results: int = 100) -> list[str]:
    """Busca archivos por glob (ej. '*.xlsx') recursivamente bajo root. Devuelve las rutas tal como
    están en disco (úselas así en las demás herramientas)."""
    r = _safe(root)
    res = []
    for p in r.rglob(pattern):
        if EXCLUIDOS & set(p.parts):
            continue
        res.append(str(p))
        if len(res) >= max_results:
            break
    return res


if __name__ == "__main__":
    mcp.run()
