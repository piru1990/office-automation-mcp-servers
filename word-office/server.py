"""MCP: lectura de documentos Word y conversión a PDF, con Word por COM (pywin32).

Abre siempre en solo lectura y en una instancia propia de Word (DispatchEx), así no toca
los documentos que el usuario tenga abiertos. No modifica ningún .docx.
"""
import os
import sys
from contextlib import contextmanager
from pathlib import Path

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP, Image
from mcp.types import ToolAnnotations

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import _paths  # noqa: E402

mcp = FastMCP("word-office")
RO = ToolAnnotations(readOnlyHint=True)
WD_WITHIN_TABLE = 12
WD_EXPORT_PDF = 17
EXT = (".docx", ".doc", ".docm", ".dotx", ".rtf")


@contextmanager
def _abrir(path: str):
    ruta = str(_paths.resolve(path))
    if not os.path.isfile(ruta):
        raise FileNotFoundError(ruta)
    if not ruta.lower().endswith(EXT):
        raise ValueError(f"Extensión no soportada; use {EXT}")
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(ruta, ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            yield doc
        finally:
            doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


def _celda(c) -> str:
    return c.Range.Text.rstrip("\r\x07").replace("\r", " / ").strip()


def _filas(tb, max_rows: int) -> list[list[str]]:
    filas: list[list[str]] = []
    for c in tb.Range.Cells:
        if c.RowIndex > max_rows:
            break
        while len(filas) < c.RowIndex:
            filas.append([])
        filas[c.RowIndex - 1].append(_celda(c))
    return filas


@mcp.tool(annotations=RO)
def describe_docx(path: str, max_chars: int = 200, max_items: int = 400) -> dict:
    """Esquema de un Word en orden de lectura: párrafos con texto (recortados a max_chars, con su estilo
    si es título o lista) y, donde va cada tabla, su número, tamaño y primera fila.
    Use read_docx_table para leer una tabla completa."""
    with _abrir(path) as doc:
        tablas = [(t.Range.Start, t.Range.End, t) for t in doc.Tables]
        vistas: set[int] = set()
        items: list[str] = []
        truncado = False
        for p in doc.Paragraphs:
            if len(items) >= max_items:
                truncado = True
                break
            r = p.Range
            if r.Information(WD_WITHIN_TABLE):
                k = next((i for i, (a, b, _) in enumerate(tablas) if a <= r.Start < b), None)
                if k is not None and k not in vistas:
                    vistas.add(k)
                    tb = tablas[k][2]
                    cab = _filas(tb, 1)
                    items.append(f"[tabla {k + 1}: {tb.Rows.Count} filas x {tb.Columns.Count} columnas] "
                                 f"primera fila: {cab[0] if cab else []}")
                continue
            texto = r.Text.strip().replace("\x0c", "[salto de página]")
            if not texto:
                continue
            if not texto.strip("/\x01 "):  # párrafo que solo contiene imágenes o gráficas
                items.append("[imagen]")
                continue
            estilo = p.Style.NameLocal
            marca = f"<{estilo}> " if estilo.lower() not in ("normal",) else ""
            items.append(marca + (texto[:max_chars] + "…" if len(texto) > max_chars else texto))
        return {
            "archivo": os.path.basename(path),
            "paginas": doc.ComputeStatistics(2),
            "parrafos": doc.Paragraphs.Count,
            "tablas": len(tablas),
            "imagenes": doc.InlineShapes.Count + doc.Shapes.Count,
            "contenido": items,
            "truncado": truncado,
        }


@mcp.tool(annotations=RO)
def read_docx_table(path: str, index: int, max_rows: int = 200) -> dict:
    """Lee una tabla de un Word (index desde 1, como en describe_docx) como lista de filas."""
    with _abrir(path) as doc:
        if not 1 <= index <= doc.Tables.Count:
            raise ValueError(f"El documento tiene {doc.Tables.Count} tablas")
        tb = doc.Tables(index)
        return {"tabla": index, "filas_totales": tb.Rows.Count, "filas": _filas(tb, max_rows)}


@mcp.tool()
def docx_to_pdf(path: str, dest: str = "") -> str:
    """Exporta un Word a PDF. dest por defecto: mismo nombre y carpeta con .pdf. No sobrescribe."""
    destino = os.path.abspath(dest) if dest else os.path.splitext(os.path.abspath(path))[0] + ".pdf"
    if not destino.lower().endswith(".pdf"):
        raise ValueError("dest debe terminar en .pdf")
    if os.path.exists(destino):
        raise FileExistsError(f"{destino} ya existe; elija otro nombre")
    with _abrir(path) as doc:
        doc.ExportAsFixedFormat(destino, WD_EXPORT_PDF)
        paginas = doc.ComputeStatistics(2)
    return f"Creado {destino} ({paginas} páginas)"


def _paginas(spec: str, total: int) -> list[int]:
    """'1-3,5' -> [0, 1, 2, 4] (índices desde 0), dentro de las páginas existentes."""
    out: list[int] = []
    for parte in (spec or "1").replace(" ", "").split(","):
        a, _, b = parte.partition("-")
        ini, fin = int(a), int(b or a)
        out += [n - 1 for n in range(ini, fin + 1) if 1 <= n <= total]
    return list(dict.fromkeys(out))


@mcp.tool(annotations=RO)
def pdf_info(path: str, max_chars: int = 3000) -> dict:
    """Páginas de un PDF, su tamaño y el texto de las primeras (cortado a max_chars)."""
    import pymupdf
    with _paths.local_copy(path) as p, pymupdf.open(p) as doc:
        texto = ""
        for pg in doc:
            texto += pg.get_text()
            if len(texto) >= max_chars:
                break
        return {"archivo": os.path.basename(str(path)), "paginas": doc.page_count,
                "tamano_pt": [round(doc[0].rect.width), round(doc[0].rect.height)] if doc.page_count else None,
                "texto": texto[:max_chars], "texto_truncado": len(texto) > max_chars}


@mcp.tool(annotations=RO, structured_output=False)
def pdf_to_images(path: str, pages: str = "1", dpi: int = 80) -> list:
    """Devuelve páginas de un PDF como imágenes para REVISARLAS VISUALMENTE (formato, tablas, saltos).
    pages: '1', '1-3' o '1,4' (máximo 6 por llamada). dpi 80 basta para ver el diseño; suba a 120-150
    para leer letra pequeña."""
    import pymupdf
    with _paths.local_copy(path) as p, pymupdf.open(p) as doc:
        idx = _paginas(pages, doc.page_count)[:6]
        if not idx:
            raise ValueError(f"El PDF tiene {doc.page_count} páginas")
        return [Image(data=doc[i].get_pixmap(dpi=max(40, min(dpi, 200))).tobytes("png"), format="png") for i in idx]


@mcp.tool()
def pdf_save_pages(path: str, dest_dir: str, pages: str = "1", dpi: int = 110) -> list[str]:
    """Guarda páginas de un PDF como PNG en dest_dir (no sobrescribe) y devuelve las rutas."""
    import pymupdf
    os.makedirs(dest_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(str(_paths.resolve(path))))[0]
    out = []
    with _paths.local_copy(path) as p, pymupdf.open(p) as doc:
        for i in _paginas(pages, doc.page_count)[:30]:
            destino = os.path.join(dest_dir, f"{base}_p{i + 1}.png")
            if os.path.exists(destino):
                raise FileExistsError(f"{destino} ya existe")
            doc[i].get_pixmap(dpi=max(40, min(dpi, 300))).save(destino)
            out.append(destino)
    return out


if __name__ == "__main__":
    mcp.run()
