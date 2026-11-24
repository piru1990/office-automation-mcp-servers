"""MCP: lectura/escritura de Excel y CSV (openpyxl/pandas; Excel por COM para .xlsb/.xls y para ver fórmulas).

Las lecturas trabajan sobre una copia temporal (ver ..\\_paths.py): así no fallan ni se quedan colgadas
cuando el archivo está abierto en Excel o sincronizándose con OneDrive, y toleran nombres con tildes
escritas de otra forma.
"""
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import _paths  # noqa: E402

mcp = FastMCP("excel-office")
RO = ToolAnnotations(readOnlyHint=True)
TEXTO = (".csv", ".txt", ".tsv")


POR_COM = (".xlsb", ".xls")  # formatos que pandas/openpyxl no leen aquí: se abren con Excel
LIMITE_COM = 240             # segundos antes de cerrar a la fuerza la instancia de Excel


def _con_excel(path: str, fn, limite: int = LIMITE_COM):
    """Ejecuta fn(workbook) con el libro abierto en una instancia propia de Excel: solo lectura, macros
    deshabilitadas, sin avisos ni actualización de vínculos. Si pasa de `limite` segundos, mata esa instancia."""
    import subprocess
    import threading

    res: dict = {}

    def trabajo(ruta: str):
        import pythoncom
        import win32com.client
        import win32process
        pythoncom.CoInitialize()
        xl = None
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
            res["pid"] = win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.AutomationSecurity = 3  # msoAutomationSecurityForceDisable: no corre macros
            xl.EnableEvents = False
            xl.AskToUpdateLinks = False
            wb = xl.Workbooks.Open(ruta, UpdateLinks=0, ReadOnly=True, AddToMru=False)
            try:
                res["ok"] = fn(wb)
            finally:
                wb.Close(False)
        except Exception as e:  # noqa: BLE001
            # sin el traceback: sus marcos retienen objetos COM y dejarían la instancia de Excel viva
            res["err"] = RuntimeError(f"{type(e).__name__}: {e}")
        finally:
            if xl is not None:
                try:
                    xl.Quit()
                except Exception:  # noqa: BLE001
                    pass
            xl = None
            pythoncom.CoUninitialize()

    with _paths.local_copy(path) as p:
        h = threading.Thread(target=trabajo, args=(str(p),), daemon=True)
        h.start()
        h.join(limite)
        if h.is_alive():
            if res.get("pid"):
                subprocess.run(["taskkill", "/PID", str(res["pid"]), "/F"], capture_output=True)
            h.join(15)
            raise TimeoutError(f"Excel no terminó de leer {Path(path).name} en {limite} s; se cerró esa instancia")
    if "err" in res:
        raise res["err"]
    return res["ok"]


def _matriz(v) -> list[list]:
    """Range.Value / Range.Formula a lista de filas (una sola celda llega como escalar)."""
    if not isinstance(v, tuple):
        return [[v]]
    return [list(f) for f in v]


def _hoja(wb, sheet: str):
    return wb.Worksheets(sheet) if sheet else wb.Worksheets(1)


def _df_com(path: str, sheet: str, max_rows: int, skip_rows: int):
    import pandas as pd

    def leer(wb):
        ur = _hoja(wb, sheet).UsedRange
        filas = min(ur.Rows.Count - skip_rows, max_rows + 1)
        if filas < 1:
            return []
        ws = ur.Worksheet
        r0, c0 = ur.Row + skip_rows, ur.Column
        rng = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + filas - 1, c0 + ur.Columns.Count - 1))
        return [[None if x is None else (x if isinstance(x, (int, float, str, bool)) else str(x)) for x in f]
                for f in _matriz(rng.Value)]

    m = _con_excel(path, leer)
    if not m:
        return pd.DataFrame()
    cab = [str(c) if c is not None else f"col{i + 1}" for i, c in enumerate(m[0])]
    return pd.DataFrame(m[1:], columns=cab)


@mcp.tool(annotations=RO)
def list_sheets(path: str) -> list[str]:
    """Hojas de un libro (.xlsx, .xlsm, .xlsb, .xls), incluidas las ocultas."""
    if Path(path).suffix.lower() in POR_COM:
        return _con_excel(path, lambda wb: [ws.Name for ws in wb.Worksheets])
    import openpyxl
    with _paths.local_copy(path) as p:
        wb = openpyxl.load_workbook(p, read_only=True)
        try:
            return wb.sheetnames
        finally:
            wb.close()


@mcp.tool(annotations=RO)
def read_table(path: str, sheet: str = "", max_rows: int = 100, skip_rows: int = 0) -> str:
    """Lee .xlsx/.xlsm/.xlsb/.csv y devuelve las primeras filas en formato CSV. skip_rows salta filas
    iniciales (títulos antes del encabezado). Funciona aunque el archivo esté abierto en Excel."""
    import pandas as pd
    if Path(path).suffix.lower() in POR_COM:
        df = _df_com(path, sheet, max_rows, skip_rows)
        return f"columnas: {list(df.columns)}\nfilas: {len(df)}\n" + df.to_csv(index=False)
    with _paths.local_copy(path) as p:
        if p.suffix.lower() in TEXTO:
            df = pd.read_csv(p, nrows=max_rows, skiprows=skip_rows, encoding_errors="replace", sep=None, engine="python")
        else:
            df = pd.read_excel(p, sheet_name=sheet or 0, nrows=max_rows, skiprows=skip_rows)
    return f"columnas: {list(df.columns)}\nfilas: {len(df)}\n" + df.to_csv(index=False)


@mcp.tool(annotations=RO)
def describe_table(path: str, sheet: str = "", max_rows: int = 200_000) -> str:
    """Estadísticas descriptivas de una hoja/CSV (hasta max_rows filas)."""
    import pandas as pd
    if Path(path).suffix.lower() in POR_COM:
        df = _df_com(path, sheet, max_rows, 0)
        return f"shape={df.shape}\n{df.dtypes.to_string()}\n\n{df.describe(include='all').to_string()}"
    with _paths.local_copy(path) as p:
        df = pd.read_csv(p, sep=None, engine="python", nrows=max_rows, encoding_errors="replace") \
            if p.suffix.lower() in TEXTO else pd.read_excel(p, sheet_name=sheet or 0, nrows=max_rows)
    return f"shape={df.shape}\n{df.dtypes.to_string()}\n\n{df.describe(include='all').to_string()}"


@mcp.tool(annotations=RO)
def workbook_logic(path: str, sheets: list[str] | None = None, max_rows: int = 60, max_cols: int = 60,
                   max_chars: int = 60_000) -> str:
    """Lógica de un libro de Excel (cualquier formato, también con macros): por cada hoja su visibilidad,
    rango usado, tablas y tablas dinámicas, y el contenido celda por celda mostrando las FÓRMULAS (no su
    resultado); al final los nombres definidos. Las hojas con más de max_rows filas se resumen al encabezado
    y 3 filas de muestra. sheets limita a esas hojas. Las macros no se ejecutan ni se extraen."""
    def leer(wb):
        out = []
        for ws in wb.Worksheets:
            if sheets and ws.Name not in sheets:
                continue
            ur = ws.UsedRange
            filas, cols = ur.Rows.Count, ur.Columns.Count
            tablas = [f"{t.Name}:{t.Range.GetAddress(False, False)}" for t in ws.ListObjects]
            pivotes = [pt.Name for pt in ws.PivotTables()]
            out.append(f"=== [{ws.Name}] visible={'sí' if ws.Visible == -1 else 'no'} rango={ur.GetAddress(False, False)}"
                       f" tablas={tablas} dinamicas={pivotes}")
            n = filas if filas <= max_rows else min(filas, 4)
            if filas > max_rows:
                out.append(f"  (hoja de {filas} filas: se muestran las primeras {n})")
            c = min(cols, max_cols)
            rng = ws.Range(ws.Cells(ur.Row, ur.Column), ws.Cells(ur.Row + n - 1, ur.Column + c - 1))
            for i, fila in enumerate(_matriz(rng.Formula)):
                for j, v in enumerate(fila):
                    if v is None or v == "":
                        continue
                    s = str(v)
                    out.append(f"  {ws.Cells(ur.Row + i, ur.Column + j).GetAddress(False, False)} = "
                               f"{s if len(s) <= 300 else s[:300] + '…'}")
        out.append("=== NOMBRES DEFINIDOS")
        out += [f"  {nm.Name} -> {nm.RefersTo}" for nm in wb.Names]
        return "\n".join(out)

    txt = _con_excel(path, leer)
    if len(txt) > max_chars:
        txt = txt[:max_chars] + f"\n… (cortado: {len(txt)} caracteres en total; pida menos hojas con `sheets`)"
    return txt


@mcp.tool()
def write_table(path: str, columns: list[str], rows: list[list], sheet: str = "Hoja1") -> str:
    """Crea un .xlsx nuevo con columnas y filas (no sobrescribe archivos existentes)."""
    import pandas as pd
    p = Path(path)
    if p.exists():
        raise FileExistsError(f"{p} ya existe; elija otro nombre")
    if p.suffix.lower() != ".xlsx":
        raise ValueError("solo .xlsx")
    pd.DataFrame(rows, columns=columns).to_excel(p, sheet_name=sheet, index=False)
    return f"Creado {p} ({len(rows)} filas)"


@mcp.tool()
def add_sheet(path: str, sheet: str, columns: list[str], rows: list[list]) -> str:
    """Agrega una hoja NUEVA a un .xlsx existente (no modifica las hojas que ya tiene). Falla si la
    hoja ya existe o si el archivo está abierto en Excel."""
    import openpyxl
    p = _paths.resolve(path)
    if p.suffix.lower() != ".xlsx":
        raise ValueError("solo .xlsx")
    wb = openpyxl.load_workbook(p)
    try:
        if sheet in wb.sheetnames:
            raise ValueError(f"La hoja '{sheet}' ya existe en {p.name}")
        ws = wb.create_sheet(sheet)
        ws.append(columns)
        for r in rows:
            ws.append(r)
        wb.save(p)
    finally:
        wb.close()
    return f"Hoja '{sheet}' agregada a {p} ({len(rows)} filas)"


if __name__ == "__main__":
    mcp.run()
