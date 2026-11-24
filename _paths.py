"""Rutas de archivos para los MCP locales.

Dos problemas de OneDrive que costaron tiempo en sesiones anteriores:
  * nombres con tilde "combinada" (o + U+0301 en vez de ó): la ruta escrita no coincide con la del disco;
  * archivos abiertos en Excel/Word o sincronizándose: abrirlos directamente falla o se queda colgado.

resolve() encuentra el archivo aunque la tilde esté escrita de otra forma, y local_copy() entrega una
copia temporal de solo lectura para trabajar sobre ella.
"""
import os
import shutil
import tempfile
import unicodedata
from contextlib import contextmanager
from pathlib import Path


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s).casefold()


def resolve(path: str | os.PathLike) -> Path:
    """Ruta existente equivalente a `path`, comparando cada tramo sin importar cómo está compuesta la tilde.
    Si no existe ni así, lanza FileNotFoundError."""
    p = Path(os.path.expandvars(str(path))).expanduser()
    if not p.is_absolute():
        p = Path.cwd() / p
    if p.exists():
        return p
    actual = Path(p.anchor)
    for parte in p.parts[1:]:
        cand = actual / parte
        if cand.exists():
            actual = cand
            continue
        try:
            match = next((e for e in actual.iterdir() if _nfc(e.name) == _nfc(parte)), None)
        except OSError:
            match = None
        if match is None:
            raise FileNotFoundError(f"No existe '{parte}' dentro de {actual}")
        actual = match
    return actual


def _copy_shared(src: Path, dst: Path) -> None:
    """Copia leyendo con todos los permisos de uso compartido, para archivos abiertos en otra aplicación."""
    import win32con
    import win32file
    h = win32file.CreateFile(str(src), win32con.GENERIC_READ,
                             win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                             None, win32con.OPEN_EXISTING, win32con.FILE_ATTRIBUTE_NORMAL, None)
    try:
        with open(dst, "wb") as out:
            while True:
                _, data = win32file.ReadFile(h, 1024 * 1024)
                if not data:
                    break
                out.write(data)
    finally:
        h.Close()


@contextmanager
def local_copy(path: str | os.PathLike):
    """Entrega una copia temporal del archivo (misma extensión) y la borra al terminar."""
    src = resolve(path)
    if not src.is_file():
        raise FileNotFoundError(f"{src} no es un archivo")
    tmpdir = Path(tempfile.mkdtemp(prefix="mcp_copy_"))
    dst = tmpdir / ("copia" + src.suffix)
    try:
        try:
            _copy_shared(src, dst)
        except Exception:
            shutil.copyfile(src, dst)
        yield dst
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
