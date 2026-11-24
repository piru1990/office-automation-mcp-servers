"""Credenciales de STOD (portal interno de Acme) para el MCP stod.

Usuario y contraseña viven en el Administrador de credenciales de Windows (credencial
genérica "stod/portal"), no en ~/.claude.json ni en el código. Nunca se imprimen.

Orden de búsqueda: variables STOD_USER / STOD_PASS -> Administrador de credenciales.

Uso:
  python _stod_env.py status   dice si hay credencial y prueba el ingreso
  python _stod_env.py set      pide usuario y contraseña (sin eco) y los guarda
  python _stod_env.py nuevaclave   GENERA una temporal aleatoria con el formato de STOD, la guarda y la copia al
                               portapapeles sin mostrarla (se limpia en 90 s). Recomendado: una por alta.
  python _stod_env.py copiarclave  vuelve a copiar la guardada al portapapeles (para entregarla al usuario)
  python _stod_env.py setclave pide una temporal elegida por usted (sin eco) y la guarda
  python _stod_env.py borrarclave  la borra (hacerlo al terminar las altas)
Seguro de correr con `!` dentro de Claude Code: nunca imprime la contraseña.
"""
import os
import sys
from pathlib import Path

TARGET = "stod/portal"
DEFAULT_URL = "http://stod.net/stodnet"


def _cred_read() -> tuple[str, str]:
    try:
        import win32cred
        c = win32cred.CredRead(TARGET, win32cred.CRED_TYPE_GENERIC)
        blob = c["CredentialBlob"]
        pwd = blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)
        return c["UserName"], pwd
    except Exception:
        return "", ""


def cred_write(user: str, password: str) -> None:
    import win32cred
    win32cred.CredWrite({
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": TARGET,
        "UserName": user,
        "CredentialBlob": password,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        "Comment": "Usuario y contraseña de STOD para el MCP stod",
    }, 0)


TARGET_CLAVE = "stod/clave_inicial"


def clave_inicial() -> str:
    """Contraseña temporal para usuarios nuevos de STOD. La guarda el supervisor con `setclave`;
    el MCP la usa al crear un usuario y jamás la devuelve ni la imprime."""
    try:
        import win32cred
        c = win32cred.CredRead(TARGET_CLAVE, win32cred.CRED_TYPE_GENERIC)
        blob = c["CredentialBlob"]
        return blob.decode("utf-16-le") if isinstance(blob, bytes) else str(blob)
    except Exception:
        return ""


def clave_write(password: str) -> None:
    import win32cred
    win32cred.CredWrite({
        "Type": win32cred.CRED_TYPE_GENERIC,
        "TargetName": TARGET_CLAVE,
        "UserName": "clave_inicial",
        "CredentialBlob": password,
        "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        "Comment": "Contraseña temporal para altas de usuario en STOD (MCP stod)",
    }, 0)


ESPECIALES = "@#$%^&+="          # los que acepta el validador de STOD
LIMPIAR_EN = 90                  # segundos que la clave se queda en el portapapeles


def clave_valida(c: str) -> bool:
    """Formato que exige STOD: 8+ caracteres con mayúscula, dígito y uno de @#$%^&+=."""
    return (len(c) >= 8 and any(x.isupper() for x in c) and any(x.isdigit() for x in c)
            and any(x in ESPECIALES for x in c))


def generar_clave(largo: int = 12) -> str:
    """Contraseña aleatoria (secrets) que cumple el formato de STOD, sin caracteres ambiguos (0 O o 1 l I)."""
    import secrets
    may, min_, dig = "ABCDEFGHJKLMNPQRSTUVWXYZ", "abcdefghijkmnpqrstuvwxyz", "23456789"
    todo = may + min_ + dig + ESPECIALES
    while True:
        c = [secrets.choice(may), secrets.choice(min_), secrets.choice(dig), secrets.choice(ESPECIALES)]
        c += [secrets.choice(todo) for _ in range(largo - len(c))]
        secrets.SystemRandom().shuffle(c)
        c = "".join(c)
        if clave_valida(c):
            return c


def _portapapeles_poner(txt: str) -> None:
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        if txt:
            win32clipboard.SetClipboardText(txt, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def _portapapeles_leer() -> str:
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
    except Exception:
        return ""
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:       # algunas versiones de pywin32 ya lo cerraron tras leer (error 1418)
            pass


def copiar_clave() -> None:
    """Copia la temporal guardada al portapapeles SIN imprimirla y programa su borrado (proceso aparte)."""
    import subprocess
    c = clave_inicial()
    if not c:
        raise SystemExit("No hay contraseña temporal guardada (use nuevaclave o setclave).")
    _portapapeles_poner(c)
    subprocess.Popen([sys.executable, str(Path(__file__)), "_limpiar"], close_fds=True,
                     creationflags=0x00000008 | 0x00000200)      # DETACHED_PROCESS | NEW_PROCESS_GROUP


def _limpiar() -> None:
    import time
    time.sleep(LIMPIAR_EN)
    c = clave_inicial()
    if c and _portapapeles_leer() == c:
        _portapapeles_poner("")


def credentials() -> tuple[str, str, str]:
    """(url, usuario, contraseña). Lanza RuntimeError si no hay credencial."""
    url = os.environ.get("STOD_URL") or DEFAULT_URL
    user, pwd = os.environ.get("STOD_USER", ""), os.environ.get("STOD_PASS", "")
    if not (user and pwd):
        user, pwd = _cred_read()
    if not (user and pwd):
        raise RuntimeError(f"No hay credencial de STOD. Guárdela con: python {Path(__file__)} set")
    return url.rstrip("/"), user, pwd


def source() -> str:
    if os.environ.get("STOD_USER") and os.environ.get("STOD_PASS"):
        return "variables de entorno STOD_USER / STOD_PASS"
    if all(_cred_read()):
        return "Administrador de credenciales de Windows"
    return ""


def main() -> None:
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "status":
        src = source()
        print(f"Credencial STOD: {src or 'NO ENCONTRADA'}")
        if src:
            sys.path.insert(0, str(Path(__file__).parent / "stod"))
            import server
            try:
                print(server._test())
            except Exception as e:
                print(f"ingreso FALLÓ: {e}")
    elif cmd == "set":
        import getpass
        user = input("Usuario de STOD: ").strip()
        pwd = getpass.getpass("Contraseña de STOD (no se muestra): ").strip()
        if not user or not pwd:
            raise SystemExit("Faltan datos; no se guardó nada.")
        cred_write(user, pwd)
        print("Guardada en el Administrador de credenciales.")
    elif cmd == "setclave":
        import getpass
        pwd = getpass.getpass("Contraseña TEMPORAL para usuarios nuevos de STOD (no se muestra): ").strip()
        if not pwd:
            raise SystemExit("Vacía; no se guardó nada.")
        clave_write(pwd)
        print("Guardada en el Administrador de credenciales (stod/clave_inicial).")
    elif cmd == "nuevaclave":
        clave_write(generar_clave())
        copiar_clave()
        print(f"Contraseña temporal NUEVA generada (formato STOD), guardada en stod/clave_inicial y copiada al "
              f"portapapeles (se limpia sola en {LIMPIAR_EN} s). No se muestra. Péguela donde la vaya a entregar "
              f"o vuelva a copiarla con: _stod_env.py copiarclave")
    elif cmd == "copiarclave":
        copiar_clave()
        print(f"Contraseña temporal copiada al portapapeles (se limpia sola en {LIMPIAR_EN} s). No se muestra.")
    elif cmd == "_limpiar":
        _limpiar()
    elif cmd == "borrarclave":
        import win32cred
        try:
            win32cred.CredDelete(TARGET_CLAVE, win32cred.CRED_TYPE_GENERIC)
            print("Contraseña temporal borrada.")
        except Exception:
            print("No había contraseña temporal guardada.")
    else:
        raise SystemExit(__doc__)


if __name__ == "__main__":
    main()
