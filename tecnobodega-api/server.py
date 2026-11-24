"""MCP: acceso a la API de distribuidores de Sub-D (Mesa de Ayuda).

- SQL Server Tecno_distribuidores (10.0.0.180): palabras clave y bitácora de la API.
- API HTTP DistribuidoresAPI (login JWT + inventario).
Las claves nunca se devuelven al modelo: se enmascaran en toda salida.
Credencial SQL: _tecno_env.py (Administrador de credenciales de Windows).
"""
import html as html_mod
import os
import re
import sys
from pathlib import Path

import httpx
import pyodbc
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import _tecno_env  # noqa: E402

mcp = FastMCP("inventory-api")
RO = ToolAnnotations(readOnlyHint=True)

API_URL = os.environ.get("TECNO_API_URL", "http://api.sub-d.example.com/DistribuidoresAPI").rstrip("/")
DOCS_DIR = r"C:\Users\<user>\OneDrive - Acme, S.A\Solicitud de APIs"
SECRET_COL = re.compile(r"clave|palabra|pass|pwd|secret|token|hash", re.I)
ID_RE = re.compile(r"^[A-Za-z0-9_]{1,30}$")
MAX_ROWS = 200


def _conn():
    server, db, user, pwd = _tecno_env.credentials()
    cs = (f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={db};UID={user};PWD={pwd};"
          "TrustServerCertificate=yes")
    return pyodbc.connect(cs, timeout=15)


def _test() -> str:
    with _conn() as c:
        row = c.cursor().execute("SELECT DB_NAME(), SUSER_SNAME()").fetchone()
    return f"conexión OK: base {row[0]}, usuario {row[1]}"


def _mask(col: str, v):
    if v is None:
        return None
    s = str(v)
    if SECRET_COL.search(col):
        return "****" + s[-2:] if len(s) > 4 else "****"
    return s


def _rows(sql: str, params=()):
    with _conn() as c:
        cur = c.cursor()
        cur.execute(sql, params)
        cols = [d[0] for d in cur.description]
        data = cur.fetchmany(MAX_ROWS)
    return [{k: _mask(k, v) for k, v in zip(cols, r)} for r in data]


def _tabla(nombre: str) -> str:
    """Valida que sea una tabla existente (evita inyección en el nombre)."""
    if not ID_RE.match(nombre):
        raise ValueError("Nombre de tabla no válido")
    return nombre


@mcp.tool(annotations=RO)
def esquema(tabla: str = "") -> dict:
    """Tablas de Tecno_distribuidores o, con `tabla`, sus columnas y tipos (solo lectura)."""
    if not tabla:
        return {"tablas": [r["t"] for r in _rows(
            "SELECT TABLE_SCHEMA+'.'+TABLE_NAME AS t FROM INFORMATION_SCHEMA.TABLES ORDER BY 1")]}
    return {"columnas": _rows(
        "SELECT COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_NAME=? ORDER BY ORDINAL_POSITION", (_tabla(tabla),))}


@mcp.tool(annotations=RO)
def muestra(tabla: str, filas: int = 5) -> list[dict]:
    """Primeras filas de una tabla; las columnas que parecen clave/token salen enmascaradas."""
    n = max(1, min(int(filas), 50))
    return _rows(f"SELECT TOP {n} * FROM [{_tabla(tabla)}]")


@mcp.tool(annotations=RO)
def buscar(tabla: str, columna: str, valor: str, limit: int = 20) -> list[dict]:
    """Filas de una tabla donde `columna` = valor (p. ej. PalabrasDistribuidor, el código del distribuidor).
    Las columnas secretas salen enmascaradas."""
    if not ID_RE.match(columna):
        raise ValueError("Nombre de columna no válido")
    n = max(1, min(int(limit), MAX_ROWS))
    return _rows(f"SELECT TOP {n} * FROM [{_tabla(tabla)}] WHERE [{columna}] = ?", (valor,))


# ------------------------------------------------------------------ página web DistribuidoresWEB (WebForms)
WEB_URL = os.environ.get("TECNO_WEB_URL", "https://api.sub-d.example.com/DistribuidoresWEB").rstrip("/")
_web: dict = {"client": None}
_HIDDEN_RE = re.compile(r'<input[^>]+type="hidden"[^>]*>', re.I)


def _campos_ocultos(html: str) -> dict:
    out = {}
    for tag in _HIDDEN_RE.findall(html):
        n = re.search(r'name="([^"]+)"', tag)
        v = re.search(r'value="([^"]*)"', tag)
        if n:
            out[n.group(1)] = v.group(1) if v else ""
    return out


def _web_login() -> httpx.Client:
    """Cliente httpx con sesión iniciada en DistribuidoresWEB (credencial del Administrador de credenciales)."""
    if _web["client"] is not None:
        return _web["client"]
    _, _, user, pwd = _tecno_env.credentials()
    c = httpx.Client(follow_redirects=True, timeout=30)
    r = c.get(f"{WEB_URL}/Login.aspx")
    data = _campos_ocultos(r.text)
    data.update({"txtUsuario": user, "txtPassword": pwd, "__EVENTTARGET": "btnLogin", "__EVENTARGUMENT": ""})
    r = c.post(f"{WEB_URL}/Login.aspx", data=data)
    if "Login.aspx" in str(r.url) and 'id="txtPassword"' in r.text:
        c.close()
        raise RuntimeError("La página rechazó la credencial guardada (revise con _tecno_env.py set).")
    _web["client"] = c
    return c


def _web_get(path: str) -> str:
    c = _web_login()
    r = c.get(f"{WEB_URL}/{path.lstrip('/')}")
    if 'id="txtPassword"' in r.text:  # sesión caducada
        _web["client"] = None
        r = _web_login().get(f"{WEB_URL}/{path.lstrip('/')}")
    return r.text


@mcp.tool(annotations=RO)
def claves_web(codigo: str = "", dump: bool = False) -> dict:
    """Lista las palabras clave de la página DistribuidoresWEB (PlbClave.aspx) con la clave ENMASCARADA.
    codigo: filtra por código de distribuidor (ej. C00191). dump=True guarda el HTML de la página en la
    carpeta temporal de la sesión para revisar su estructura (el HTML contiene claves; no se imprime)."""
    html = _web_get("Views/PlbClave.aspx")
    if dump:
        p = Path(os.environ.get("TEMP", ".")) / "tecno_plbclave.html"
        p.write_text(html, encoding="utf-8")
    filas = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        celdas = [re.sub(r"<[^>]+>", "", t).strip() for t in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        celdas = [x for x in celdas if x]
        if len(celdas) >= 3:
            filas.append({"id": celdas[0], "codigo": celdas[1], "clave": _mask("clave", celdas[2])})
    if codigo:
        filas = [f for f in filas if f["codigo"].upper() == codigo.upper()]
    return {"total": len(filas), "filas": filas[:MAX_ROWS], "url": str(WEB_URL) + "/Views/PlbClave.aspx"}


def _filas_web(html: str) -> list[tuple[str, str, str]]:
    out = []
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I):
        # html.unescape: la página muestra la «ñ» como &#241;; la clave real lleva el carácter.
        c = [html_mod.unescape(re.sub(r"<[^>]+>", "", t)).strip()
             for t in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S | re.I)]
        c = [x for x in c if x]
        if len(c) >= 3:
            out.append((c[0], c[1], c[2]))
    return out


def _filas_codigo(codigo: str) -> list[tuple[str, str, str]]:
    """Filas (id, código, clave) de un distribuidor, de la más reciente (id mayor) a la más antigua."""
    f = [x for x in _filas_web(_web_get("Views/PlbClave.aspx")) if x[1].upper() == codigo.upper()]
    return sorted(f, key=lambda x: int(x[0]) if x[0].isdigit() else -1, reverse=True)


def _clave_web_fila(codigo: str) -> tuple[str, str, int] | None:
    """(id, clave en claro, cantidad de filas) de la clave MÁS RECIENTE del distribuidor. Solo uso interno del
    servidor; la clave nunca se devuelve al modelo. Un código puede tener varias filas (duplicados)."""
    f = _filas_codigo(codigo)
    return (f[0][0], f[0][2], len(f)) if f else None


def _clave_web(codigo: str) -> str | None:
    fila = _clave_web_fila(codigo)
    return fila[1] if fila else None


def _api_post(ruta: str, body: dict, token: str | None = None) -> dict:
    h = {"Content-Type": "application/json"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    r = httpx.post(f"{API_URL}/{ruta}", json=body, headers=h, timeout=45)
    try:
        j = r.json()
    except Exception:
        j = {"texto": r.text[:300]}
    return {"http": r.status_code, "json": j}


def _token(codigo: str) -> tuple[str | None, str]:
    clave = _clave_web(codigo)
    if not clave:
        return None, f"{codigo} no tiene palabra clave en DistribuidoresWEB."
    r = _api_post("api/login/authenticate", {"UserName": codigo, "Password": clave})
    t = r["json"] if isinstance(r["json"], str) else None
    if r["http"] != 200 or not t:
        return None, f"Login rechazado (HTTP {r['http']})."
    return t, "ok"


def _consulta(codigo: str, ruta: str, extra: dict | None = None) -> dict:
    t, msg = _token(codigo)
    if not t:
        return {"ok": False, "motivo": msg}
    body = {"CodigoDistribuidor": codigo, **(extra or {})}
    r = _api_post(f"apI/Inventario/{ruta}", body, t)
    j = r["json"] if isinstance(r["json"], dict) else {}
    return {"ok": r["http"] == 200 and j.get("CodigoRespuesta") == "00000", "http": r["http"],
            "codigo_respuesta": j.get("CodigoRespuesta"), "descripcion": j.get("DescripcionRespuesta"),
            "objeto": j.get("objeto")}


@mcp.tool(annotations=RO)
def distribuidor_info(codigo: str) -> dict:
    """¿El distribuidor tiene palabra clave en DistribuidoresWEB? Devuelve id, código y clave enmascarada."""
    f = _filas_codigo(codigo)
    if not f:
        return {"codigo": codigo, "tiene_clave": False}
    r = {"codigo": codigo, "tiene_clave": True, "id": f[0][0], "clave": _mask("clave", f[0][2]),
         "nota": "id = la fila más reciente (la que usan copiar_clave y probar_acceso)"}
    if len(f) > 1:
        r["duplicados"] = len(f)
        r["filas"] = [{"id": x[0], "clave": _mask("clave", x[2])} for x in f]
        r["aviso"] = ("El código tiene más de una palabra clave: confirmar con cuál se entregó a su "
                      "cliente antes de enviar credenciales.")
    return r


def _portapapeles_poner(txt: str) -> None:
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        if txt:
            win32clipboard.SetClipboardText(txt, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def _limpiar_portapapeles_en(seg: int, esperado: str) -> None:
    """Vacía el portapapeles tras `seg` segundos, solo si aún contiene la clave copiada."""
    import threading
    import time

    def _t():
        time.sleep(seg)
        try:
            import win32clipboard
            win32clipboard.OpenClipboard()
            try:
                actual = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            except Exception:
                actual = None
            finally:
                try:
                    win32clipboard.CloseClipboard()
                except Exception:
                    pass
            if actual == esperado:
                _portapapeles_poner("")
        except Exception:
            pass
    threading.Thread(target=_t, daemon=True).start()


@mcp.tool()
def copiar_clave(codigo: str, limpiar_en: int = 90) -> dict:
    """Copia la palabra clave del distribuidor al PORTAPAPELES de el supervisor (para pegarla en el borrador de
    credenciales). La clave nunca se devuelve ni se imprime: solo se informa la máscara. Se limpia sola."""
    fila = _clave_web_fila(codigo)
    if not fila:
        return {"codigo": codigo, "copiada": False, "motivo": "sin palabra clave en DistribuidoresWEB"}
    id_, clave, n = fila
    _portapapeles_poner(clave)
    seg = max(15, min(int(limpiar_en), 600))
    _limpiar_portapapeles_en(seg, clave)
    r = {"codigo": codigo, "copiada": True, "id": id_, "clave": _mask("clave", clave), "se_limpia_en_seg": seg}
    if n > 1:
        r["duplicados"] = n
        r["aviso"] = (f"{codigo} tiene {n} palabras clave; se copió la más reciente (id {id_}). Ver "
                      "distribuidor_info y confirmar cuál se entregó al cliente.")
    return r


@mcp.tool(annotations=RO)
def probar_acceso(codigo: str) -> dict:
    """Prueba de punta a punta lo que haría el distribuidor: login con su clave guardada + consulta Inventario.
    Devuelve OK/error, cantidad de artículos y 3 de muestra. No devuelve token ni clave. Queda en la bitácora
    del servicio como un acceso de ese distribuidor.
    OJO: el OK NO prueba que la clave sea la correcta: el login acepta credenciales inválidas (tema del token
    con Persona X, 19-nov-2026). Solo confirma que el servicio responde."""
    r = _consulta(codigo, "Inventario")
    obj = r.pop("objeto", None) or []
    r["articulos"] = len(obj)
    r["muestra"] = [{k: a.get(k) for k in ("CodigoArticulo", "NombreArticulo", "Existencia", "PrecioQuetzales")}
                    for a in obj[:3]]
    return r


@mcp.tool(annotations=RO)
def inventario(codigo: str, articulo: str = "", limit: int = 20) -> dict:
    """Inventario que ve el distribuidor. Con `articulo`, el detalle de ese artículo (InventarioDetalle)."""
    r = _consulta(codigo, "InventarioDetalle", {"CodigoArticulo": articulo}) if articulo else _consulta(codigo, "Inventario")
    obj = r.pop("objeto", None) or []
    r["total"] = len(obj)
    r["articulos"] = obj[:max(1, min(int(limit), 100))]
    return r


@mcp.tool(annotations=RO)
def inventario_bodega(codigo: str, articulo: str) -> dict:
    """Existencia de un artículo por bodega, tal como la ve el distribuidor (InventarioBodega)."""
    return _consulta(codigo, "InventarioBodega", {"CodigoArticulo": articulo})


@mcp.tool(annotations=RO)
def api_ping() -> dict:
    """Comprueba que el servicio HTTP de la API responde (sin credenciales)."""
    try:
        r = httpx.post(f"{API_URL}/api/login/authenticate", json={"UserName": "_ping", "Password": "_ping"},
                       timeout=15)
        return {"url": API_URL, "http": r.status_code, "respuesta": r.text[:200]}
    except Exception as e:
        return {"url": API_URL, "error": str(e)}


if __name__ == "__main__":
    mcp.run()
