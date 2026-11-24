"""MCP: Outlook clásico por COM (pywin32). Lee correo y reglas, crea borradores.

No borra correo. Solo envía con send_draft(confirmar=True), un borrador ya existente y con OK de el supervisor.
Requiere Outlook clásico (Office16) con el perfil del usuario configurado; si Outlook no está abierto,
COM lo inicia en segundo plano.
"""
import ctypes
import mimetypes
import os
import re
import struct
import subprocess
import tempfile
import time
import unicodedata
import uuid
from datetime import datetime
from html import unescape
from email.message import EmailMessage
from email.utils import format_datetime
from typing import Any

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

mcp = FastMCP("outlook-desktop")
RO = ToolAnnotations(readOnlyHint=True)

# OlDefaultFolders
DEFAULT_FOLDERS = {"inbox": 6, "sent": 5, "drafts": 16, "deleted": 3, "junk": 23, "outbox": 4}
# OlRuleActionType que sacan correo del buzón
ACTION_TYPES = {1: "move_to_folder", 2: "assign_category", 3: "delete", 4: "delete_permanently",
                5: "copy_to_folder", 6: "forward", 7: "forward_as_attachment", 8: "redirect",
                9: "server_reply", 10: "template", 11: "flag", 13: "clear_categories",
                14: "importance", 15: "sensitivity", 17: "play_sound", 19: "mark_read",
                23: "desktop_alert", 41: "stop"}
OL_EMBEDDED_ITEM = 5
PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
BODY_MAX = 20000
EMAIL_RE = re.compile(r"^[^@\s<>]+@[^@\s<>]+\.[^@\s<>.]+$")
# Firma de el supervisor: una imagen (la misma de Outlook web), que se incrusta al final de los borradores.
FIRMA_PATH = os.environ.get("OUTLOOK_FIRMA") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "firma.png")
FIRMA_ANCHO_MAX = 648  # ancho con que Outlook web muestra la firma


class OutlookNoDisponible(RuntimeError):
    """No se pudo conectar con Outlook por COM. `elevado` = el proceso corre como administrador."""

    def __init__(self, msg: str, elevado: bool = False):
        super().__init__(msg)
        self.elevado = elevado


def _elevado() -> bool:
    """True si este proceso corre elevado (administrador, integridad Alta)."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _outlook_corriendo() -> bool:
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"], capture_output=True,
                             text=True, timeout=15, creationflags=0x08000000).stdout
        return "outlook.exe" in out.lower()
    except Exception:
        return False


def _app():
    """Application de Outlook: GetActiveObject, luego Dispatch, con un reintento. Si falla, explica por qué."""
    pythoncom.CoInitialize()
    ultimo: Exception | None = None
    for intento in (1, 2):
        for fn in (lambda: win32com.client.GetActiveObject("Outlook.Application"),
                   lambda: win32com.client.Dispatch("Outlook.Application")):
            try:
                return fn()
            except Exception as e:
                ultimo = e
        if intento == 1:
            time.sleep(2)
    elevado, corriendo = _elevado(), _outlook_corriendo()
    if elevado and corriendo:
        raise OutlookNoDisponible(
            "Claude Code está elevado (administrador) y Outlook no: COM no conecta entre niveles distintos. "
            f"Reinicie Claude Code sin 'Ejecutar como administrador'. Detalle: {ultimo}", elevado=True)
    if not corriendo:
        raise OutlookNoDisponible(
            f"Outlook no está abierto y no se pudo iniciar en segundo plano. Ábralo e intente de nuevo. Detalle: {ultimo}",
            elevado=elevado)
    raise OutlookNoDisponible(f"Outlook está abierto pero COM no responde: {ultimo}", elevado=elevado)


def _ns():
    app = _app()
    return app, app.GetNamespace("MAPI")


def _store(ns, mailbox: str):
    """Almacén (buzón) abierto en Outlook cuyo nombre contiene `mailbox`; vacío = el buzón por defecto."""
    if not mailbox:
        return ns.DefaultStore
    hits = [s for s in ns.Stores if mailbox.lower() in (s.DisplayName or "").lower()]
    if len(hits) != 1:
        raise ValueError(f"'{mailbox}' coincide con {len(hits)} buzones abiertos en Outlook: "
                         f"{[s.DisplayName for s in ns.Stores]}")
    return hits[0]


def _folder(ns, name: str, mailbox: str = ""):
    """'inbox', 'sent', 'drafts', 'deleted', 'junk' o una ruta desde la raíz del buzón ('Bandeja de entrada/Sub').
    mailbox: otro buzón abierto en Outlook (compartido o adicional), por parte de su nombre."""
    key = (name or "inbox").strip().lower()
    store = _store(ns, mailbox)
    if key in DEFAULT_FOLDERS:
        return store.GetDefaultFolder(DEFAULT_FOLDERS[key]) if mailbox else ns.GetDefaultFolder(DEFAULT_FOLDERS[key])
    folder = store.GetRootFolder()
    for part in [p for p in name.replace("\\", "/").split("/") if p]:
        match = next((f for f in folder.Folders if f.Name.lower() == part.lower()), None)
        if match is None:
            raise ValueError(f"No existe la carpeta '{part}' dentro de '{folder.Name}'")
        folder = match
    return folder


def _get(obj, attr: str, default: Any = None) -> Any:
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _dt(value) -> str | None:
    if value is None:
        return None
    try:
        return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second).isoformat()
    except Exception:
        return str(value)


def _smtp(recipient) -> str:
    try:
        return recipient.PropertyAccessor.GetProperty(PR_SMTP)
    except Exception:
        return _get(recipient, "Address", "") or ""


def _sender(item) -> str:
    addr = _get(item, "SenderEmailAddress", "") or ""
    if addr.upper().startswith("/O="):
        try:
            addr = item.Sender.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            pass
    return addr


def _summary(item) -> dict:
    return {
        "entry_id": _get(item, "EntryID"),
        "class": _get(item, "MessageClass"),
        "subject": _get(item, "Subject"),
        "sender_name": _get(item, "SenderName"),
        "sender": _sender(item),
        "received": _dt(_get(item, "ReceivedTime") or _get(item, "CreationTime")),
        "unread": _get(item, "UnRead"),
        "attachments": _get(_get(item, "Attachments"), "Count", 0),
    }


def _detail(item, body_max: int = BODY_MAX) -> dict:
    d = _summary(item)
    recips = _get(item, "Recipients")
    to, cc = [], []
    if recips is not None:
        for r in recips:
            (cc if _get(r, "Type") == 2 else to).append({"name": _get(r, "Name"), "email": _smtp(r)})
    body = _get(item, "Body", "") or ""
    d.update({
        "to": to, "cc": cc,
        "conversation_topic": _get(item, "ConversationTopic"),
        "conversation_id": _get(item, "ConversationID"),
        "sent": _dt(_get(item, "SentOn")),
        "body": body[:body_max],
        "body_truncated": len(body) > body_max,
        "attachment_names": [_get(a, "FileName") or _get(a, "DisplayName") for a in (_get(item, "Attachments") or [])],
    })
    return d


@mcp.tool(annotations=RO)
def list_folders(parent: str = "", depth: int = 2, mailbox: str = "") -> list[dict]:
    """Carpetas de un buzón (ruta, total y no leídos). parent vacío = raíz. mailbox vacío = buzón por
    defecto; si no, parte del nombre de otro buzón abierto en Outlook."""
    _, ns = _ns()
    root = _folder(ns, parent, mailbox) if parent else _store(ns, mailbox).GetRootFolder()
    out: list[dict] = []

    def walk(folder, path: str, level: int):
        for f in folder.Folders:
            p = f"{path}/{f.Name}" if path else f.Name
            out.append({"path": p, "items": _get(_get(f, "Items"), "Count", 0), "unread": _get(f, "UnReadItemCount", 0)})
            if level < depth:
                walk(f, p, level + 1)

    walk(root, parent.strip("/"), 1)
    return out


@mcp.tool(annotations=RO)
def search_messages(folder: str = "inbox", subject_contains: str = "", sender_contains: str = "",
                    since: str = "", unread_only: bool = False, limit: int = 25,
                    body_contains: str = "", until: str = "", offset: int = 0, mailbox: str = "") -> list[dict]:
    """Busca en una carpeta, del más reciente al más antiguo. since/until: 'YYYY-MM-DD' (hora local,
    ambos inclusive). body_contains busca en el cuerpo. offset salta los primeros N resultados (paginar).
    mailbox: parte del nombre de otro buzón abierto en Outlook (vacío = el propio).
    Los informes de no entrega (NDR) aparecen con class 'REPORT.IPM.Note.NDR'."""
    _, ns = _ns()
    items = _folder(ns, folder, mailbox).Items
    items.Sort("[ReceivedTime]", True)
    clauses = []
    if body_contains:
        clauses.append("\"urn:schemas:httpmail:textdescription\" LIKE '%" + body_contains.replace("'", "''") + "%'")
    if until:
        datetime.strptime(until, "%Y-%m-%d")
        clauses.append("\"urn:schemas:httpmail:datereceived\" <= '" + until + " 23:59'")
    if subject_contains:
        clauses.append("\"urn:schemas:httpmail:subject\" LIKE '%" + subject_contains.replace("'", "''") + "%'")
    if sender_contains:
        s = sender_contains.replace("'", "''")
        clauses.append("(\"urn:schemas:httpmail:fromemail\" LIKE '%" + s + "%' OR \"urn:schemas:httpmail:fromname\" LIKE '%" + s + "%')")
    if since:
        datetime.strptime(since, "%Y-%m-%d")
        clauses.append("\"urn:schemas:httpmail:datereceived\" >= '" + since + " 00:00'")
    if unread_only:
        clauses.append("\"urn:schemas:httpmail:read\" = 0")
    if clauses:
        items = items.Restrict("@SQL=" + " AND ".join(clauses))
    out = []
    for i, item in enumerate(items):
        if i < offset:
            continue
        out.append(_summary(item))
        if len(out) >= max(1, min(limit, 200)):
            break
    return out


@mcp.tool(annotations=RO)
def get_message(entry_id: str, include_embedded: bool = True) -> dict:
    """Mensaje completo por entry_id. Con include_embedded, también lee los correos adjuntos
    (p. ej. el mensaje original dentro de un NDR)."""
    _, ns = _ns()
    item = ns.GetItemFromID(entry_id)
    d = _detail(item)
    if include_embedded:
        embedded = []
        for att in (_get(item, "Attachments") or []):
            if _get(att, "Type") != OL_EMBEDDED_ITEM:
                continue
            path = os.path.join(tempfile.mkdtemp(prefix="olmcp_"), "embedded.msg")
            try:
                att.SaveAsFile(path)
                inner = ns.OpenSharedItem(path)
                embedded.append(_detail(inner))
                inner.Close(1)  # olDiscard
            except Exception as e:
                embedded.append({"error": str(e), "name": _get(att, "DisplayName")})
            finally:
                try:
                    os.remove(path)
                    os.rmdir(os.path.dirname(path))
                except OSError:
                    pass
        d["embedded_messages"] = embedded
    return d


# --- Directorio (libreta global de direcciones, GAL) -------------------------------------------------
_GAL_CACHE: dict = {"t": 0.0, "rows": []}
_GAL_TTL = 1800  # segundos; la GAL cambia poco y recorrerla por COM tarda ~10 s
_GAL_TIPOS = {0: "usuario", 6: "usuario_remoto", 1: "lista_distribucion", 5: "lista_distribucion",
              9: "lista_distribucion", 7: "sala_equipo", 8: "sala_equipo"}


def _sin_tildes(s: str) -> str:
    return unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()


def _gal_rows(ns) -> list[dict]:
    """Todas las entradas de la GAL (con caché). Solo lectura."""
    if _GAL_CACHE["rows"] and time.time() - _GAL_CACHE["t"] < _GAL_TTL:
        return _GAL_CACHE["rows"]
    entries = ns.GetGlobalAddressList().AddressEntries
    rows = []
    for k in range(1, entries.Count + 1):
        try:
            ae = entries.Item(k)
            eu = None
            try:
                eu = ae.GetExchangeUser()
            except Exception:
                pass
            smtp = (_get(eu, "PrimarySmtpAddress", "") or "") if eu is not None else ""
            if not smtp:
                try:
                    smtp = ae.PropertyAccessor.GetProperty(PR_SMTP) or ""
                except Exception:
                    smtp = ""
            rows.append({
                "nombre": _get(ae, "Name", "") or "",
                "correo": smtp.lower(),
                "puesto": (_get(eu, "JobTitle", "") or "") if eu is not None else "",
                "departamento": (_get(eu, "Department", "") or "") if eu is not None else "",
                "empresa": (_get(eu, "CompanyName", "") or "") if eu is not None else "",
                "oficina": (_get(eu, "OfficeLocation", "") or "") if eu is not None else "",
                "tipo": _GAL_TIPOS.get(_get(ae, "DisplayType", -1), "otro"),
            })
        except Exception:
            continue
    _GAL_CACHE.update(t=time.time(), rows=rows)
    return rows


@mcp.tool(annotations=RO)
def buscar_directorio(consulta: str, limite: int = 10, incluir_listas: bool = False) -> list[dict]:
    """Busca personas en la libreta global de direcciones (GAL) de Exchange de Acme y devuelve su correo
    corporativo real con puesto, departamento y empresa. Útil para saber el correo de alguien que en Odoo
    figura como <código>@employee.example. Solo lectura, sin contraseñas: usa el perfil de Outlook.
    consulta: nombre en cualquier orden y sin tildes ('barbara mendez', 'Aguilar, Mario') o parte de un
    correo ('user@'). Deben aparecer todas las palabras. Las listas de distribución se omiten salvo
    incluir_listas=True. Ojo: un correo de puesto (p. ej. contauxiliar9@) puede ser el de la persona."""
    palabras = [w for w in re.split(r"[\s,;]+", _sin_tildes(consulta)) if w]
    if not palabras:
        raise ValueError("consulta vacía")
    _, ns = _ns()
    hits = []
    for r in _gal_rows(ns):
        if r["tipo"] == "lista_distribucion" and not incluir_listas:
            continue
        nombre = _sin_tildes(r["nombre"])
        resto = " ".join(_sin_tildes(r[k]) for k in ("correo", "puesto", "departamento"))
        if all(w in nombre or w in resto for w in palabras):
            en_nombre = sum(1 for w in palabras if w in nombre)
            hits.append((-en_nombre, nombre, r))
    hits.sort(key=lambda h: (h[0], h[1]))
    return [h[2] for h in hits[:max(1, min(limite, 50))]]


def _rule_dict(rule) -> dict:
    actions = []
    try:
        for i in range(1, rule.Actions.Count + 1):
            a = rule.Actions.Item(i)
            if not _get(a, "Enabled"):
                continue
            entry = {"type": ACTION_TYPES.get(a.ActionType, str(a.ActionType))}
            recips = _get(a, "Recipients")
            if recips is not None:
                entry["recipients"] = [_smtp(r) or _get(r, "Name") for r in recips]
            fld = _get(a, "Folder")
            if fld is not None:
                entry["folder"] = _get(fld, "FolderPath")
            actions.append(entry)
    except Exception as e:
        actions.append({"error": str(e)})
    return {"name": rule.Name, "enabled": bool(rule.Enabled), "order": rule.ExecutionOrder,
            "type": "receive" if rule.RuleType == 0 else "send", "actions": actions}


@mcp.tool(annotations=RO)
def list_rules() -> list[dict]:
    """Reglas del buzón por defecto, con sus acciones (reenvíos, redirecciones, mover a carpeta…)."""
    _, ns = _ns()
    return [_rule_dict(r) for r in ns.DefaultStore.GetRules()]


@mcp.tool()
def set_rule_enabled(name: str, enabled: bool) -> dict:
    """Activa o desactiva una regla por su nombre exacto. No la borra."""
    _, ns = _ns()
    rules = ns.DefaultStore.GetRules()
    matches = [r for r in rules if r.Name == name]
    if len(matches) != 1:
        raise ValueError(f"Se esperaba 1 regla llamada '{name}' y hay {len(matches)}")
    matches[0].Enabled = bool(enabled)
    rules.Save()
    return _rule_dict(matches[0])


def _add_recipients(item, addrs: str, tipo: int) -> list[str]:
    """Agrega destinatarios (separados por ';' o ',') como Para=1, CC=2 o CCO=3 y los resuelve contra la
    libreta. Omite los que ya están en el mensaje. Devuelve los que no se pudieron resolver."""
    ya = set()
    for r in item.Recipients:
        ya.add((_get(r, "Name", "") or "").lower())
        ya.add((_smtp(r) or "").lower())
    sin_resolver = []
    for addr in [a.strip() for a in re.split(r"[;,]", addrs or "")]:
        if not addr or addr.lower() in ya:
            continue
        ya.add(addr.lower())
        if "@" in addr and not EMAIL_RE.match(addr):
            sin_resolver.append(addr)
            continue
        r = item.Recipients.Add(addr)
        r.Type = tipo
        if not r.Resolve():
            sin_resolver.append(addr)
    return sin_resolver


def _recipients_out(item) -> list[dict]:
    tipos = {1: "to", 2: "cc", 3: "bcc"}
    return [{"type": tipos.get(_get(r, "Type"), "to"), "name": _get(r, "Name"), "email": _smtp(r),
             "resolved": bool(_get(r, "Resolved"))} for r in item.Recipients]


def _descartar(item, sin_resolver: list[str]):
    """Cierra sin guardar y falla: no se deja un borrador con destinatarios que Outlook marcaría en rojo."""
    try:
        item.Close(1)  # olDiscard
    except Exception:
        pass
    raise ValueError(f"No se pudieron resolver estos destinatarios: {sin_resolver}. "
                     "Usar la dirección de correo completa. No se creó el borrador.")


def _firma_html(item) -> str:
    """Incrusta la imagen de firma (adjunto oculto con cid) y devuelve su HTML; '' si no hay archivo."""
    if not os.path.isfile(FIRMA_PATH):
        return ""
    with open(FIRMA_PATH, "rb") as f:
        cab = f.read(24)
    ancho, alto = FIRMA_ANCHO_MAX, 0
    if cab[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", cab[16:24])
        ancho = min(FIRMA_ANCHO_MAX, w)
        alto = round(h * ancho / w)
    cid = f"firma-{uuid.uuid4().hex}"
    att = item.Attachments.Add(FIRMA_PATH)
    # Outlook decide solo si el adjunto es "oculto" (lo marca al enviar, por estar referenciado en el HTML).
    att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
    dims = f' width="{ancho}"' + (f' height="{alto}"' if alto else "")
    estilo = f"width:{ancho}px;" + (f"height:{alto}px;" if alto else "") + "max-width:100%"
    return f'<p><img src="cid:{cid}"{dims} style="{estilo}" alt="Firma"></p>'


def _rutas_adjuntos(attachments: list[str] | None) -> list[str]:
    rutas = [os.path.abspath(a) for a in (attachments or [])]
    faltan = [r for r in rutas if not os.path.isfile(r)]
    if faltan:
        raise FileNotFoundError(f"No existen estos adjuntos: {faltan}")
    return rutas


def _draft_out(item, rutas: list[str], signature: bool, firmado: bool) -> dict:
    out = {"entry_id": item.EntryID, "subject": item.Subject, "to": item.To, "cc": item.CC,
           "recipients": _recipients_out(item), "attachments": [os.path.basename(r) for r in rutas],
           "signature": firmado, "sent": False}
    if signature and not firmado:
        out["warning"] = f"Borrador sin firma: no existe {FIRMA_PATH}"
    return out


EML_DIR = os.path.join(tempfile.gettempdir(), "outlook-mcp-eml")


def _lista(addrs: str) -> list[str]:
    out = [a.strip() for a in re.split(r"[;,]", addrs or "") if a.strip()]
    malos = [a for a in out if not EMAIL_RE.match(a)]
    if malos:
        raise ValueError(f"Modo .eml: use direcciones completas; no válidas: {malos}")
    return out


def _crear_eml(to: str, cc: str, bcc: str, subject: str, html_body: str, rutas: list[str],
               signature: bool, motivo: str) -> dict:
    """Plan B sin COM: escribe un .eml con X-Unsent: 1 y lo abre con explorer.exe, que lo lanza en el
    contexto del usuario no elevado; Outlook lo muestra como borrador para revisar y guardar. No envía."""
    para, copia, oculta = _lista(to), _lista(cc), _lista(bcc)
    if not para:
        raise ValueError("Modo .eml: falta el destinatario (to).")
    msg = EmailMessage()
    msg["X-Unsent"] = "1"
    msg["Subject"] = subject
    msg["To"] = ", ".join(para)
    if copia:
        msg["Cc"] = ", ".join(copia)
    if oculta:
        msg["Bcc"] = ", ".join(oculta)
    msg["Date"] = format_datetime(datetime.now().astimezone())
    firma_cid, firma_bytes, ancho, alto = "", b"", FIRMA_ANCHO_MAX, 0
    if signature and os.path.isfile(FIRMA_PATH):
        with open(FIRMA_PATH, "rb") as f:
            firma_bytes = f.read()
        if firma_bytes[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", firma_bytes[16:24])
            ancho = min(FIRMA_ANCHO_MAX, w)
            alto = round(h * ancho / w)
        firma_cid = f"firma-{uuid.uuid4().hex}"
    firma = (f'<p><img src="cid:{firma_cid}" width="{ancho}"' + (f' height="{alto}"' if alto else "") +
             f' style="max-width:100%" alt="Firma"></p>') if firma_cid else ""
    pos = html_body.lower().rfind("</body")
    html = html_body[:pos] + firma + html_body[pos:] if pos >= 0 else html_body + firma
    msg.set_content(re.sub(r"<[^>]+>", "", re.sub(r"</(p|li|h\d)>|<br\s*/?>", "\n", html, flags=re.I)))
    msg.add_alternative(html, subtype="html")
    if firma_cid:
        msg.get_payload()[1].add_related(firma_bytes, "image", "png", cid=f"<{firma_cid}>", filename="firma.png",
                                         disposition="inline")
    for r in rutas:
        tipo = mimetypes.guess_type(r)[0] or "application/octet-stream"
        a, b = tipo.split("/", 1)
        with open(r, "rb") as f:
            msg.add_attachment(f.read(), maintype=a, subtype=b, filename=os.path.basename(r))
    os.makedirs(EML_DIR, exist_ok=True)
    seguro = "".join(c if c.isalnum() or c in " ._-" else "_" for c in subject)[:60].strip() or "borrador"
    ruta = os.path.join(EML_DIR, f"{datetime.now():%Y%m%d_%H%M%S}_{seguro}.eml")
    with open(ruta, "wb") as f:
        f.write(bytes(msg))
    subprocess.Popen(["explorer.exe", ruta])
    return {"mode": "eml", "path": ruta, "subject": subject, "to": para, "cc": copia, "bcc": oculta,
            "attachments": [os.path.basename(r) for r in rutas], "signature": bool(firma_cid), "sent": False,
            "motivo": motivo,
            "nota": "Se abrió en Outlook como borrador sin guardar: revíselo y guárdelo o envíelo desde esa ventana. "
                    "Los destinatarios no se resolvieron contra la libreta."}


@mcp.tool()
def create_draft(to: str, subject: str, html_body: str, cc: str = "",
                 attachments: list[str] | None = None, bcc: str = "", signature: bool = True,
                 modo: str = "auto") -> dict:
    """Crea un borrador en la carpeta Borradores. NO lo envía. to/cc/bcc: direcciones completas separadas
    por ';'; se resuelven contra la libreta y, si alguna no resuelve, falla sin crear el borrador.
    attachments: rutas absolutas de archivos locales a adjuntar.
    signature=True agrega al final la firma de el supervisor (imagen): NO escribir firma, nombre ni cargo en
    html_body; terminar en la despedida ("Saludos,").
    modo: 'auto' (COM; si Claude Code está elevado cae a un .eml que se abre en Outlook para revisar y
    guardar) o 'eml' (forzar ese camino)."""
    rutas = _rutas_adjuntos(attachments)
    if modo == "eml" or os.environ.get("OUTLOOK_FORZAR_EML") == "1":
        return _crear_eml(to, cc, bcc, subject, html_body, rutas, signature, "solicitado")
    try:
        app, _ = _ns()
    except OutlookNoDisponible as e:
        if not e.elevado:
            raise
        return _crear_eml(to, cc, bcc, subject, html_body, rutas, signature, str(e))
    mail = app.CreateItem(0)
    malos = _add_recipients(mail, to, 1) + _add_recipients(mail, cc, 2) + _add_recipients(mail, bcc, 3)
    if malos:
        _descartar(mail, malos)
    mail.Subject = subject
    firma = _firma_html(mail) if signature else ""
    pos = html_body.lower().rfind("</body")
    mail.HTMLBody = html_body[:pos] + firma + html_body[pos:] if pos >= 0 else html_body + firma
    for r in rutas:
        mail.Attachments.Add(r)
    mail.Save()
    return _draft_out(mail, rutas, signature, bool(firma))


# --- Historial citado sin instrucciones internas (el gerente, 24-nov-2026) -------------------------------------------
# Al responder al usuario, el historial citado no debe llevar correos con instrucciones internas (p. ej. Persona X
# → Técnico B). Quien cuenta como «interno» está en internos.txt (una entrada por línea, correo o parte del nombre).
INTERNOS_PATH = os.environ.get("OUTLOOK_INTERNOS") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "internos.txt")
_MARCAS_BLOQUE = [
    re.compile(r"<div[^>]*style=['\"][^'\">]*border-top:\s*solid\s*#E1E1E1", re.I),   # encabezado de Outlook de escritorio
    re.compile(r"<div[^>]*id=['\"]?divRplyFwdMsg", re.I),                              # encabezado de Outlook web / móvil
    re.compile(r"-{3,}\s*(?:Original Message|Mensaje original|Mensaje reenviado)\s*-{3,}", re.I),
]
_DE_RE = re.compile(r"\b(?:De|From)\s*:\s*(.{1,200}?)\s*(?:Enviado(?: el)?|Sent|Para|To|Fecha|Date)\s*:\s*(.{0,60})", re.I | re.S)


def _internos() -> list[str]:
    try:
        with open(INTERNOS_PATH, encoding="utf-8") as f:
            return [ln.strip().lower() for ln in f if ln.strip() and not ln.lstrip().startswith("#")]
    except OSError:
        return []


def _texto_plano(fragmento: str) -> str:
    return " ".join(unescape(re.sub(r"<[^>]+>", " ", fragmento)).split())


def _bloques_citados(html: str) -> list[dict]:
    """Parte el HTML de una respuesta en bloques por los encabezados de mensajes citados. El bloque 0 es el mensaje
    al que se responde. Cada bloque: {inicio, fin, de, fecha}. Lista vacía si no reconoce ningún encabezado."""
    marcas = sorted((m.start(), n) for n, rx in enumerate(_MARCAS_BLOQUE) for m in rx.finditer(html))
    limpios: list[int] = []
    ultimo_tipo = -1
    for i, tipo in marcas:
        if limpios and tipo != ultimo_tipo and i - limpios[-1] < 400:
            continue  # el mismo encabezado reconocido por dos marcas distintas (p. ej. «Original Message» + div)
        limpios.append(i)
        ultimo_tipo = tipo
    # el encabezado de Outlook va dentro de un <div> contenedor: se incluye para no dejar etiquetas sueltas
    ajustados = []
    for i in limpios:
        previo = html[:i].rstrip()
        ajustados.append(len(previo) - 5 if previo.lower().endswith("<div>") else i)
    cuerpo_fin = html.lower().rfind("</body")
    cuerpo_fin = cuerpo_fin if cuerpo_fin > 0 else len(html)
    bloques = []
    for n, ini in enumerate(ajustados):
        fin = ajustados[n + 1] if n + 1 < len(ajustados) else cuerpo_fin
        m = _DE_RE.search(_texto_plano(html[ini:min(fin, ini + 3000)]))
        bloques.append({"inicio": ini, "fin": fin, "de": (m.group(1).strip() if m else ""), "fecha": (m.group(2).strip() if m else "")})
    return bloques


def _es_interno(de: str, internos: list[str]) -> bool:
    d = (de or "").lower()
    return bool(d) and any(t in d for t in internos)


def _internos_en_historial(html: str) -> list[dict]:
    """Bloques citados escritos por internos, sin contar el primero (el mensaje al que se responde)."""
    internos = _internos()
    return [b for n, b in enumerate(_bloques_citados(html)) if n > 0 and _es_interno(b["de"], internos)]


def _quitar_internos(html: str) -> tuple[str, list[dict]]:
    quitar = _internos_en_historial(html)
    for b in sorted(quitar, key=lambda x: x["inicio"], reverse=True):
        html = html[:b["inicio"]] + html[b["fin"]:]
    return html, [{"de": b["de"], "fecha": b["fecha"]} for b in quitar]


@mcp.tool()
def create_reply_draft(entry_id: str, html_body: str, reply_all: bool = True, cc_add: str = "",
                       attachments: list[str] | None = None, signature: bool = True,
                       to_fallback: str = "", subject_fallback: str = "", modo: str = "auto",
                       limpiar_internos: bool = True) -> dict:
    """Crea un borrador de RESPUESTA dentro del hilo de un mensaje (conserva asunto, destinatarios y el
    texto citado). NO lo envía. reply_all=False responde solo al remitente. cc_add: copias extra
    (direcciones completas separadas por ';'). attachments: rutas absolutas de archivos locales.
    signature=True agrega la firma de el supervisor (imagen) entre el texto nuevo y el mensaje citado: NO
    escribir firma, nombre ni cargo en html_body; terminar en la despedida ("Saludos,").
    Sin COM (Claude Code elevado, o modo='eml') genera un .eml nuevo: no conserva el hilo ni cita el
    original y necesita to_fallback y subject_fallback (con 'RE: ' ya puesto).
    limpiar_internos=True (por defecto) quita del historial citado los correos escritos por las personas de
    internos.txt (instrucciones internas que el usuario no debe ver; regla de el gerente, 24-nov-2026) y lo informa
    en internos_removidos. No toca el mensaje al que se responde. send_draft se niega a enviar si queda alguno."""
    rutas = _rutas_adjuntos(attachments)
    if modo == "eml" or os.environ.get("OUTLOOK_FORZAR_EML") == "1":
        return _crear_eml(to_fallback, cc_add, "", subject_fallback, html_body, rutas, signature, "solicitado")
    try:
        _, ns = _ns()
    except OutlookNoDisponible as e:
        if not e.elevado:
            raise
        if not (to_fallback and subject_fallback):
            raise OutlookNoDisponible(f"{e} Para usar el modo .eml pase to_fallback y subject_fallback.", True)
        return _crear_eml(to_fallback, cc_add, "", subject_fallback, html_body, rutas, signature, str(e))
    original = ns.GetItemFromID(entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    malos = _add_recipients(reply, cc_add, 2)
    reply.Recipients.ResolveAll()
    if malos:
        _descartar(reply, malos)
    firma = _firma_html(reply) if signature else ""
    nuevo = html_body + firma
    citado = reply.HTMLBody or ""
    pos = citado.lower().find("<body")
    if pos >= 0:
        fin = citado.find(">", pos) + 1
        reply.HTMLBody = citado[:fin] + nuevo + citado[fin:]
    else:
        reply.HTMLBody = nuevo + citado
    out_extra: dict = {}
    if limpiar_internos:
        try:
            limpio, quitados = _quitar_internos(reply.HTMLBody or "")
            if quitados:
                reply.HTMLBody = limpio
            out_extra["internos_removidos"] = quitados
        except Exception as e:  # no se pudo interpretar el HTML: se deja igual y send_draft lo revisará
            out_extra["warning_historial"] = f"No se pudo limpiar el historial: {e}"
    for r in rutas:
        reply.Attachments.Add(r)
    reply.Save()
    return {**_draft_out(reply, rutas, signature, bool(firma)), **out_extra}


ENVIADOS_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "enviados.log")
OL_FOLDER_DRAFTS = DEFAULT_FOLDERS["drafts"]


@mcp.tool()
def send_draft(entry_id: str, confirmar: bool = False, permitir_internos: bool = False) -> dict:
    """ENVÍA un borrador que ya existe en la carpeta Borradores. Solo usar cuando el supervisor haya dado su OK
    explícito para ese borrador (CLAUDE.md: por defecto los borradores NO se envían). Con confirmar=False
    (por defecto) no envía: devuelve asunto, Para/CC y si pasaría las guardas (simulación). Guardas: el
    ítem debe estar en Borradores, no haber sido enviado, tener al menos un destinatario y todos resueltos,
    y el historial citado NO debe contener correos escritos por internos (internos.txt; regla de el gerente,
    24-nov-2026) salvo permitir_internos=True, que solo se usa con OK explícito de el supervisor.
    ANTES de enviar acuses de recibo verificar en Odoo que el ticket siga sin respuesta del personal
    (odoo-helpdesk\\enviar_acuses.py lo hace). Cada envío queda anotado en outlook-desktop\\enviados.log."""
    _, ns = _ns()
    item = ns.GetItemFromID(entry_id)
    drafts = ns.GetDefaultFolder(OL_FOLDER_DRAFTS)
    problemas = []
    try:
        en_borradores = item.Parent.EntryID == drafts.EntryID
    except Exception:
        en_borradores = False
    if not en_borradores:
        problemas.append("el mensaje no está en la carpeta Borradores")
    if _get(item, "Sent", False):
        problemas.append("el mensaje ya fue enviado")
    recs = _recipients_out(item)
    if not recs:
        problemas.append("no tiene destinatarios")
    sin_resolver = [r["email"] or r["name"] for r in recs if not r["resolved"]]
    if sin_resolver:
        problemas.append(f"destinatarios sin resolver: {sin_resolver}")
    if not permitir_internos:
        try:
            for b in _internos_en_historial(item.HTMLBody or ""):
                problemas.append(f"el historial citado incluye un correo de «{b['de']}» ({b['fecha']}): quitarlo del "
                                 "historial antes de enviar (o permitir_internos=True con OK de el supervisor)")
        except Exception as e:
            problemas.append(f"no se pudo revisar el historial citado: {e}")
    out = {"entry_id": entry_id, "subject": item.Subject, "to": item.To, "cc": item.CC,
           "recipients": recs, "problemas": problemas, "sent": False}
    if problemas:
        out["error"] = "No se envió: " + "; ".join(problemas)
        return out
    if not confirmar:
        out["simulacion"] = "Pasa las guardas. Para enviarlo llamar de nuevo con confirmar=True."
        return out
    destino = "; ".join(f"{r['type']}:{r['email'] or r['name']}" for r in recs)
    asunto = item.Subject
    item.Send()
    with open(ENVIADOS_LOG, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}\t{asunto}\t{destino}\n")
    out["sent"] = True
    return out


@mcp.tool()
def save_attachments(entry_id: str, dest_dir: str, name_contains: str = "", include_inline: bool = False) -> list[dict]:
    """Guarda en dest_dir los adjuntos de un mensaje y devuelve sus rutas. No sobrescribe (agrega un
    número). Por defecto omite las imágenes incrustadas en el cuerpo (firmas, logos)."""
    _, ns = _ns()
    item = ns.GetItemFromID(entry_id)
    os.makedirs(dest_dir, exist_ok=True)
    out = []
    for att in (_get(item, "Attachments") or []):
        nombre = _get(att, "FileName") or _get(att, "DisplayName") or "adjunto"
        if name_contains and name_contains.lower() not in nombre.lower():
            continue
        if not include_inline:
            try:
                cid = att.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
            except Exception:
                cid = ""
            if cid and cid in (_get(item, "HTMLBody", "") or ""):
                continue
        if _get(att, "Type") == OL_EMBEDDED_ITEM and not nombre.lower().endswith(".msg"):
            nombre += ".msg"
        seguro = "".join(c if c.isalnum() or c in "._- ()" else "_" for c in nombre)
        base, ext = os.path.splitext(seguro)
        ruta, n = os.path.join(dest_dir, seguro), 1
        while os.path.exists(ruta):
            ruta, n = os.path.join(dest_dir, f"{base} ({n}){ext}"), n + 1
        att.SaveAsFile(ruta)
        out.append({"name": nombre, "path": ruta, "bytes": os.path.getsize(ruta)})
    return out


@mcp.tool(annotations=RO)
def list_events(since: str = "", until: str = "", limit: int = 50) -> list[dict]:
    """Eventos del calendario propio entre since y until ('YYYY-MM-DD', ambos inclusive; por defecto hoy),
    con las repeticiones expandidas, en orden cronológico."""
    import locale
    _, ns = _ns()
    d1 = datetime.strptime(since, "%Y-%m-%d") if since else datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    d2 = (datetime.strptime(until, "%Y-%m-%d") if until else d1).replace(hour=23, minute=59)
    items = ns.GetDefaultFolder(9).Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True
    try:
        locale.setlocale(locale.LC_TIME, "")
    except locale.Error:
        pass
    fmt = lambda d: d.strftime("%x %H:%M")  # noqa: E731  (fecha corta según la configuración regional de Windows)
    items = items.Restrict(f"[Start] <= '{fmt(d2)}' AND [End] >= '{fmt(d1)}'")
    out = []
    for ev in items:
        out.append({"subject": _get(ev, "Subject"), "start": _dt(_get(ev, "Start")), "end": _dt(_get(ev, "End")),
                    "all_day": _get(ev, "AllDayEvent"), "location": _get(ev, "Location"),
                    "organizer": _get(ev, "Organizer"), "required": _get(ev, "RequiredAttendees"),
                    "is_recurring": _get(ev, "IsRecurring"), "entry_id": _get(ev, "EntryID")})
        if len(out) >= max(1, min(limit, 200)):
            break
    return out


@mcp.tool()
def mark_read(entry_ids: list[str], read: bool = True) -> dict:
    """Marca mensajes como leídos (o no leídos)."""
    _, ns = _ns()
    for eid in entry_ids:
        item = ns.GetItemFromID(eid)
        item.UnRead = not read
        item.Save()
    return {"updated": len(entry_ids), "read": read}


@mcp.tool()
def move_message(entry_ids: list[str], folder: str) -> dict:
    """Mueve mensajes a otra carpeta del buzón (ruta como en list_folders). No permite 'deleted'."""
    if folder.strip().lower() in ("deleted", "elementos eliminados", "deleted items"):
        raise ValueError("Este MCP no borra correo")
    _, ns = _ns()
    target = _folder(ns, folder)
    for eid in entry_ids:
        ns.GetItemFromID(eid).Move(target)
    return {"moved": len(entry_ids), "folder": target.FolderPath}


if __name__ == "__main__":
    mcp.run()
