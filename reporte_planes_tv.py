"""
Reporte mensual de suscripciones de TV (Active + Past Due Date).

Cada fin de mes (último día, 23:50):
  1. Descarga de ERPNext (API monthly_subscription_snapshot) las suscripciones en estado Active y
     Past Due Date al corte, de los planes de TV del reporte (TV_PLANES).
  2. Genera el Excel "planes tv {mes} {año}.xlsx" con el formato del ejemplo de septiembre:
     una hoja "planes tv {mes}" con la tabla Tabla1 (Suscripcion, Estado, Centro poblado,
     Ciudad, Codigo Dane municipio, Estrato, custom_television, Plan), una fila por suscripción.
  3. Lo envía por Outlook (Microsoft Graph) a los destinatarios, con copia, mensaje
     personalizado y firma de quien envía.

Uso:
  python reporte_planes_tv.py                      # corte del mes, genera y envía
  python reporte_planes_tv.py --sin-enviar         # solo genera el Excel y una vista previa del correo
  python reporte_planes_tv.py --verificar          # revisa la API de ERPNext, permisos de Azure y archivos
  python reporte_planes_tv.py --corte 2026-10-31 --forzar
  python reporte_planes_tv.py --corte 2026-10-31 --foto /ruta/2026-10.json.gz   # desde una foto guardada
  python reporte_planes_tv.py --para "a@x.co;b@x.co" --cc "c@x.co" --mensaje "Texto del correo"
  python reporte_planes_tv.py --iniciar-sesion     # MAIL_AUTH=delegado: código para microsoft.com/devicelogin
  python reporte_planes_tv.py --completar-sesion   # espera a que se ingrese el código y guarda la sesión

Fecha de corte: si se ejecuta el último día del mes, el corte es ese día; cualquier otro día
se toma como el corte del mes anterior (por si el servidor estaba apagado el último día).
Un corte que ya se envió no se reenvía, salvo con --forzar.

Reglas del reporte:
  - Solo los planes de TV_PLANES. Firestick cuenta como TV aunque en ERPNext tenga TV = 0,
    por eso custom_television es 1 en todas las filas.
  - Una fila por suscripción: si tiene dos planes del reporte, queda el de mayor prioridad
    (el primero en TV_PLANES; ej. WinSports + 2 SetBox antes que 1 SetBox).
  - El estado es el del corte (status_al_corte). El código DANE del municipio viene de ERPNext (o,
    si falta, de la ciudad); el centro poblado es la cabecera municipal (código + "000").

Programación: en el servidor Linux con deploy/reporte-planes-tv.timer (último día del mes,
23:50 hora de Bogotá, igual que el corte del portal de TV).
Configuración: variables TV_*, ERPNEXT_* (API de ERPNext) y de Azure en .env.
Requisitos: pip install -r requirements.txt
"""

import argparse
import base64
import gzip
import json
import logging
import mimetypes
import os
import re
import sys
import time
import unicodedata
from calendar import monthrange
from collections import Counter
from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
from urllib.parse import unquote
from zoneinfo import ZoneInfo

import msal
import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def _ruta(valor: str) -> Path:
    """Rutas relativas se toman desde la carpeta del script (la tarea programada no usa cwd)."""
    p = Path(valor)
    return p if p.is_absolute() else BASE_DIR / p


def _lista(valor: str | None, separadores: str = ";,") -> list[str]:
    return [x.strip() for x in re.split(f"[{separadores}]", valor or "") if x.strip()]


# ---------- Configuración ----------
OUTPUT_DIR = _ruta(os.getenv("OUTPUT_DIR", "reportes"))
PLANTILLA_HTML = _ruta(os.getenv("TV_PLANTILLA_HTML", "plantilla_tv.html"))
MENSAJE_PATH = _ruta(os.getenv("TV_MENSAJE", "mensaje_tv.txt"))
FIRMA_PATH = _ruta(os.getenv("TV_FIRMA", "firma.jpg"))
FIRMA_ANCHO = int(os.getenv("TV_FIRMA_ANCHO", "421"))
ASUNTO = os.getenv("TV_MAIL_SUBJECT", "Reporte de Número de Abonados TV {mes} {año}")
ESTADO_PATH = OUTPUT_DIR / "estado_planes_tv.json"
# Las fechas (corte, envío) van en hora de Colombia aunque el servidor esté en UTC.
ZONA = ZoneInfo(os.getenv("TV_ZONA_HORARIA", "America/Bogota"))

# Planes del reporte, en orden de prioridad (separados por ;). Son los del ejemplo de septiembre.
PLANES = _lista(os.getenv("TV_PLANES") or ";".join([
    "Giga Fibra TV + WinSports + 2 SetBox - Nov-2025",
    "Giga Fibra TV + WinSports + 1 SetBox - Jan-2025",
    "Giga Fibra TV + 1 SetBox - Set-2024",
    "Giga Fibra TV + 1 SetBox - May-2024",
    "Firestick",
]), ";")
ESTADOS = ("Active", "Past Due Date")

# Código DANE (DIVIPOLA) por municipio, según custom_ciudad_l normalizado (sin tildes ni "D.C.").
# Si aparece una ciudad nueva, el log lo avisa y queda sin código hasta agregarla aquí.
DANE_MUNICIPIOS = {
    "bogota": 11001, "soacha": 25754, "tenjo": 25799, "chia": 25175,
    "cota": 25214, "zipaquira": 25899, "tocancipa": 25817,
}

GRAPH = "https://graph.microsoft.com/v1.0"
# Cómo se autentica el envío:
#   aplicacion: la app usa su secreto (permiso Mail.Send de tipo Aplicación, requiere administrador).
#   delegado:   envía con la cuenta de MAIL_SENDER, que inicia sesión una vez con un código
#               (Mail.Send Delegado + "Permitir flujos de clientes públicos" en la app de Azure).
MAIL_AUTH = os.getenv("MAIL_AUTH", "aplicacion").strip().lower()
SESION_PATH = _ruta(os.getenv("MAIL_SESION", ".sesion_correo.json"))  # llave de sesión (no la contraseña)
SESION_PENDIENTE = SESION_PATH.with_name(".sesion_correo_pendiente.json")
SCOPES_DELEGADO = ["https://graph.microsoft.com/Mail.Send"]
LIMITE_GRAPH = 3_500_000  # sendMail admite ~4 MB por petición (adjuntos en base64 incluidos)

MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
               "septiembre", "octubre", "noviembre", "diciembre"]

# Excel: mismo formato que "planes tv septiembre 2026.xlsx".
ENCABEZADOS = ["Suscripcion", "Estado", "Centro poblado", "Ciudad", "Codigo Dane municipio",
               "Estrato", "custom_television", "Plan"]
ANCHOS = {"A": 16.9, "C": 16.7, "E": 23.7, "G": 19.0, "H": 43.0}

# API de ERPNext: método monthly_subscription_snapshot (el mismo del portal CRC). Devuelve cada
# suscripción con su estado reconstruido a la fecha de corte (status_al_corte), sus planes
# (planes_todos, separados por " | "), ciudad, estrato y código DANE.
ERPNEXT_URL = os.getenv("ERPNEXT_URL", "https://next.gigafibra.co").rstrip("/")
ERPNEXT_EMPRESA = os.getenv("ERPNEXT_COMPANY", "GIGA FIBRA SAS")
METODO_SNAPSHOT = "/api/method/monthly_subscription_snapshot"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
_handlers = [logging.FileHandler(OUTPUT_DIR / "reporte_planes_tv.log", encoding="utf-8")]
if sys.stderr:  # sin consola (p. ej. pythonw) no hay stderr
    _handlers.append(logging.StreamHandler())
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", handlers=_handlers)
log = logging.getLogger("reporte_planes_tv")


# ---------- Fecha de corte y estado ----------
def ahora() -> datetime:
    return datetime.now(ZONA)


def fecha_de_corte(valor: str | None) -> date:
    if valor:
        return date.fromisoformat(valor)
    hoy = ahora().date()
    if hoy.day == monthrange(hoy.year, hoy.month)[1]:
        return hoy
    return hoy.replace(day=1) - timedelta(days=1)


def cargar_estado() -> dict:
    if ESTADO_PATH.exists():
        return json.loads(ESTADO_PATH.read_text(encoding="utf-8"))
    return {}


def guardar_estado(estado: dict) -> None:
    ESTADO_PATH.write_text(json.dumps(estado, indent=2, ensure_ascii=False), encoding="utf-8")


# ---------- 1. Datos de ERPNext (API) ----------
def erpnext_get(ruta: str, params: dict | None = None, intentos: int = 3):
    """GET a la API de ERPNext con token de API; reintenta fallas transitorias."""
    clave, secreto = os.getenv("ERPNEXT_API_KEY"), os.getenv("ERPNEXT_API_SECRET")
    if not (clave and secreto):
        raise RuntimeError("Faltan ERPNEXT_API_KEY / ERPNEXT_API_SECRET en .env")
    cabeceras = {"Authorization": f"token {clave}:{secreto}", "Accept": "application/json"}
    for intento in range(1, intentos + 1):
        try:
            resp = requests.get(f"{ERPNEXT_URL}{ruta}", params=params, headers=cabeceras, timeout=300)
        except requests.RequestException as e:
            if intento == intentos:
                raise RuntimeError(f"No fue posible conectarse con ERPNext: {e}") from e
        else:
            if resp.status_code < 400:
                return resp.json()
            if resp.status_code not in (429, 502, 503, 504) or intento == intentos:  # solo fallas pasajeras
                raise RuntimeError(f"ERPNext respondió {resp.status_code} en {ruta}: {resp.text[:300]}")
        log.warning("ERPNext no respondió bien (intento %d/%d); reintento en %ds", intento, intentos, 20 * intento)
        time.sleep(20 * intento)


def descargar_snapshot(corte: date) -> list[dict]:
    """Suscripciones con su estado al corte (API monthly_subscription_snapshot del mes del corte).

    El estado se reconstruye a la fecha de corte aunque se ejecute después; los planes, la ciudad y
    el estrato son los del momento de la consulta (por eso se ejecuta el último día a las 23:50).
    """
    params = {"desde": corte.replace(day=1).isoformat(), "corte": corte.isoformat(), "company": ERPNEXT_EMPRESA}
    datos = (erpnext_get(METODO_SNAPSHOT, params).get("message") or {}).get("data")
    if not isinstance(datos, list):
        raise RuntimeError(f"La API de ERPNext devolvió un formato inesperado: {str(datos)[:300]}")
    log.info("ERPNext devolvió %d suscripciones con estado al %s", len(datos), corte)
    return datos


def leer_foto(ruta: Path, corte: date) -> list[dict]:
    """Foto guardada de la misma API (ej. la que toma el portal CRC el último día a las 23:50).

    Sirve para generar un corte pasado con los planes y la ciudad de ese momento.
    """
    abrir = gzip.open if ruta.suffix == ".gz" else open
    with abrir(ruta, "rt", encoding="utf-8") as fh:
        foto = json.load(fh)
    if foto.get("corte") and foto["corte"] != corte.isoformat():
        raise RuntimeError(f"La foto es del corte {foto['corte']}, no del {corte}")
    log.info("Datos de la foto %s (tomada %s%s)", ruta.name, foto.get("tomada", "?"),
             ", tardía" if foto.get("tardia") else "")
    return foto["snapshot"]["data"]


def _normalizar_ciudad(ciudad: str | None) -> str:
    texto = unicodedata.normalize("NFKD", ciudad or "").encode("ascii", "ignore").decode().lower()
    return re.sub(r",?\s*d\.?\s*c\.?$", "", texto).strip()


def por_plan(datos: list[dict]) -> list[dict]:
    """Registros del snapshot (una suscripción con sus planes) -> una entrada por plan."""
    return [
        {"name": d["subscription"], "status": d.get("status_al_corte"), "ciudad": d.get("ciudad"),
         "estrato": d.get("estrato"), "plan": plan.strip(),
         "dane": d.get("codigo_dane_municipio"), "centro": d.get("codigo_dane_centro_poblado")}
        for d in datos for plan in (d.get("planes_todos") or "").split(" | ") if plan.strip()
    ]


def armar_filas(datos: list[dict]) -> list[list]:
    """Filas del Excel: una por suscripción, con el plan de mayor prioridad y los códigos DANE.

    `datos` trae una entrada por plan: name, status, ciudad, estrato, plan y, si los hay, dane/centro.
    """
    prioridad = {plan: i for i, plan in enumerate(PLANES)}
    elegidas: dict[str, dict] = {}  # conserva el orden en que llegan las suscripciones
    for d in datos:
        if d.get("status") not in ESTADOS or d.get("plan") not in prioridad:
            continue
        actual = elegidas.get(d["name"])
        if actual is None or prioridad[d["plan"]] < prioridad[actual["plan"]]:
            elegidas[d["name"]] = d

    filas, sin_dane = [], Counter()
    for d in elegidas.values():
        dane = d.get("dane") or DANE_MUNICIPIOS.get(_normalizar_ciudad(d.get("ciudad")))
        if not dane:
            sin_dane[d.get("ciudad")] += 1
        centro = d.get("centro") or (dane * 1000 if dane else None)  # cabecera municipal
        filas.append([
            d["name"], d["status"], centro, d.get("ciudad"), dane, d.get("estrato"),
            1,  # todos los planes del reporte cuentan como TV (Firestick también)
            d["plan"],
        ])
    for ciudad, n in sin_dane.items():
        log.warning("Ciudad sin código DANE: '%s' (%d suscripciones). Agrégala en DANE_MUNICIPIOS.", ciudad, n)
    return filas


# ---------- 2. Excel ----------
def generar_excel(filas: list[list], corte: date, destino: Path) -> None:
    if not filas:
        raise RuntimeError("No hay suscripciones de los planes del reporte; no se genera ni envía el Excel")
    wb = Workbook()
    ws = wb.active
    ws.title = f"planes tv {MESES_LARGO[corte.month - 1]}"
    ws.append(ENCABEZADOS)
    for fila in filas:
        ws.append(fila)
    tabla = Table(displayName="Tabla1", ref=f"A1:H{len(filas) + 1}")
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabla)
    for letra, ancho in ANCHOS.items():
        ws.column_dimensions[letra].width = ancho
    wb.save(destino)

    por_plan = Counter(f[7] for f in filas)
    por_estado = Counter(f[1] for f in filas)
    for plan in PLANES:
        log.info("  %-50s %5d", plan, por_plan[plan])
    log.info("  Total %d suscripciones (%s)", len(filas),
             ", ".join(f"{e}: {por_estado[e]}" for e in ESTADOS))
    log.info("Excel guardado en %s (%.0f KB)", destino, destino.stat().st_size / 1024)


# ---------- 3. Correo ----------
def _adjunto_graph(nombre: str, contenido: bytes, cid: str | None = None) -> dict:
    item = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": nombre,
        "contentType": mimetypes.guess_type(nombre)[0] or "application/octet-stream",
        "contentBytes": base64.b64encode(contenido).decode(),
    }
    if cid:
        item["contentId"] = cid
        item["isInline"] = True
    return item


def _rellenar(texto: str, valores: dict) -> str:
    for clave, valor in valores.items():
        texto = texto.replace("{" + clave + "}", str(valor))
    return texto


def _texto_a_html(texto: str) -> str:
    parrafos = re.split(r"\n\s*\n", texto.strip())
    return "\n".join(f'<p style="margin:0 0 12px 0;">{escape(p).replace(chr(10), "<br>")}</p>'
                     for p in parrafos)


def construir_firma() -> tuple[str, list[tuple[str, bytes, str]]]:
    """Devuelve (html, imágenes incrustadas como (nombre, bytes, cid)).

    TV_FIRMA puede ser una imagen (firma.jpg) o el .htm de una firma de Outlook (en Windows,
    %APPDATA%\\Microsoft\\Signatures; copiarlo junto al script en el servidor). Las imágenes
    se incrustan con cid para que se vean en cualquier cliente de correo.
    """
    if not FIRMA_PATH.exists():
        log.warning("No se encontró la firma %s; el correo sale sin firma", FIRMA_PATH)
        return "", []

    if FIRMA_PATH.suffix.lower() not in (".htm", ".html"):
        cid = "firma"
        html = (f'<img src="cid:{cid}" width="{FIRMA_ANCHO}" alt="Firma" '
                f'style="display:block;border:0;">')
        return html, [(f"firma{FIRMA_PATH.suffix.lower()}", FIRMA_PATH.read_bytes(), cid)]

    crudo = FIRMA_PATH.read_bytes()
    try:
        html = crudo.decode("utf-8")
    except UnicodeDecodeError:
        html = crudo.decode("cp1252")
    cuerpo = re.search(r"<body[^>]*>(.*)</body>", html, re.S | re.I)
    html = cuerpo.group(1) if cuerpo else html
    imagenes = []

    def data_uri(m):
        cid = f"firma{len(imagenes) + 1}"
        ext = mimetypes.guess_extension(m.group(1)) or ".png"
        imagenes.append((cid + ext, base64.b64decode(m.group(2)), cid))
        return f'src="cid:{cid}"'

    def archivo_local(m):
        ruta = FIRMA_PATH.parent / unquote(m.group(1))
        if not ruta.exists():
            return m.group(0)
        cid = f"firma{len(imagenes) + 1}"
        imagenes.append((cid + ruta.suffix.lower(), ruta.read_bytes(), cid))
        return f'src="cid:{cid}"'

    html = re.sub(r'src="data:(image/[\w.+-]+);base64,([^"]+)"', data_uri, html)
    html = re.sub(r'src="(?!cid:|https?:|data:)([^"]+)"', archivo_local, html)
    return html, imagenes


def construir_correo(corte: date, total: int, mensaje: str | None = None):
    valores = {
        "fecha_corte": f"{corte:%d/%m/%Y}",
        "mes": MESES_LARGO[corte.month - 1],
        "anio": corte.year,
        "año": corte.year,
        "total": f"{total:,}".replace(",", "."),
    }
    if mensaje is not None:
        cuerpo = _texto_a_html(_rellenar(mensaje.replace("\\n", "\n"), valores))
    else:
        texto = _rellenar(MENSAJE_PATH.read_text(encoding="utf-8"), valores)
        es_html = MENSAJE_PATH.suffix.lower() in (".htm", ".html")
        cuerpo = texto if es_html else _texto_a_html(texto)
    firma, imagenes = construir_firma()
    plantilla = re.sub(r"<!--.*?-->", "", PLANTILLA_HTML.read_text(encoding="utf-8"), flags=re.S)
    html = plantilla.replace("{{mensaje}}", cuerpo).replace("{{firma}}", firma).strip()
    return _rellenar(ASUNTO, valores), html, imagenes


def guardar_vista_previa(asunto, html, imagenes, para, cc, excel: Path) -> Path:
    for nombre, contenido, cid in imagenes:
        tipo = mimetypes.guess_type(nombre)[0] or "image/png"
        html = html.replace(f"cid:{cid}", f"data:{tipo};base64,{base64.b64encode(contenido).decode()}")
    cabecera = (
        '<div style="font:13px Segoe UI,Arial;background:#f3f3f3;padding:10px 14px;'
        'margin-bottom:16px;border-left:4px solid #999;">'
        f"<b>Vista previa (no enviado)</b><br>Para: {escape('; '.join(para)) or '(sin configurar)'}"
        f"<br>CC: {escape('; '.join(cc)) or '-'}<br>Asunto: {escape(asunto)}"
        f"<br>Adjunto: {escape(excel.name)}</div>"
    )
    ruta = OUTPUT_DIR / "vista_previa_correo.html"
    ruta.write_text(f'<meta charset="utf-8">{cabecera}{html}', encoding="utf-8")
    return ruta


def _autoridad() -> str:
    return f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}"


def _escribir_privado(ruta: Path, texto: str) -> None:
    ruta.write_text(texto, encoding="utf-8")
    try:
        os.chmod(ruta, 0o600)
    except OSError:
        pass


def _app_delegada():
    cache = msal.SerializableTokenCache()
    if SESION_PATH.exists():
        cache.deserialize(SESION_PATH.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(os.environ["AZURE_CLIENT_ID"], authority=_autoridad(),
                                       token_cache=cache)
    return app, cache


def iniciar_sesion() -> None:
    """Paso 1 del modo delegado: pide a Microsoft un código para iniciar sesión."""
    app, _ = _app_delegada()
    flujo = app.initiate_device_flow(scopes=SCOPES_DELEGADO)
    if "user_code" not in flujo:
        raise RuntimeError(f"Microsoft no entregó el código: {flujo.get('error_description', flujo)}")
    _escribir_privado(SESION_PENDIENTE, json.dumps(flujo))
    print(f"Abre {flujo['verification_uri']} e ingresa el código {flujo['user_code']}. "
          f"Entra con {os.environ['MAIL_SENDER']}. Vence en {flujo['expires_in'] // 60} minutos.")
    print("Luego ejecuta: reporte_planes_tv.py --completar-sesion (espera a que termines).")


def completar_sesion() -> None:
    """Paso 2 del modo delegado: espera a que se termine el inicio de sesión y guarda la sesión."""
    if not SESION_PENDIENTE.exists():
        raise RuntimeError("No hay un inicio de sesión pendiente; ejecuta primero --iniciar-sesion")
    flujo = json.loads(SESION_PENDIENTE.read_text(encoding="utf-8"))
    app, cache = _app_delegada()
    avisado = False
    while time.time() < flujo["expires_at"]:
        # exit_condition=True: un solo intento por vuelta, para manejar aquí los errores.
        r = app.acquire_token_by_device_flow(flujo, exit_condition=lambda _: True)
        if "access_token" in r:
            break
        espera = flujo.get("interval", 5)
        if r.get("error") not in ("authorization_pending", "slow_down"):
            if "AADSTS7000218" not in r.get("error_description", ""):
                raise RuntimeError(f"No se completó el inicio de sesión: {r.get('error_description', r)}")
            # Se puede activar mientras el código siga vigente. Se espera más entre intentos porque
            # Microsoft corta (AADSTS50196) si recibe muchos intentos fallidos seguidos.
            espera = 30
            if not avisado:
                log.warning("Falta activar 'Permitir flujos de clientes públicos' en la app de Azure; sigo esperando")
                avisado = True
        time.sleep(espera)
    else:
        SESION_PENDIENTE.unlink(missing_ok=True)
        raise RuntimeError("El código venció sin completar el inicio de sesión; repite --iniciar-sesion")
    SESION_PENDIENTE.unlink(missing_ok=True)
    usuario = (r.get("id_token_claims") or {}).get("preferred_username", "")
    if usuario.lower() != os.environ["MAIL_SENDER"].lower():
        raise RuntimeError(f"Se inició sesión como {usuario}, pero MAIL_SENDER es "
                           f"{os.environ['MAIL_SENDER']}; no se guardó la sesión")
    _escribir_privado(SESION_PATH, cache.serialize())
    log.info("Sesión guardada para %s en %s", usuario, SESION_PATH.name)


def _token_delegado() -> str:
    app, cache = _app_delegada()
    cuentas = app.get_accounts(username=os.environ["MAIL_SENDER"])
    r = app.acquire_token_silent(SCOPES_DELEGADO, account=cuentas[0]) if cuentas else None
    if not r or "access_token" not in r:
        raise RuntimeError(f"No hay una sesión válida de {os.environ['MAIL_SENDER']}: inicia sesión con "
                           "--iniciar-sesion y --completar-sesion")
    if cache.has_state_changed:  # Microsoft renovó la llave: guardarla
        _escribir_privado(SESION_PATH, cache.serialize())
    return r["access_token"]


def obtener_token() -> str:
    if MAIL_AUTH == "delegado":
        return _token_delegado()
    app = msal.ConfidentialClientApplication(
        os.environ["AZURE_CLIENT_ID"],
        authority=_autoridad(),
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"No se pudo obtener token: {result.get('error_description')}")
    token = result["access_token"]
    carga = token.split(".")[1]
    roles = json.loads(base64.urlsafe_b64decode(carga + "=" * (-len(carga) % 4))).get("roles", [])
    if "Mail.Send" not in roles:
        raise RuntimeError("La app de Azure no tiene el permiso de aplicación Mail.Send con "
                           "consentimiento de administrador (permisos actuales: %s)" % (roles or "ninguno"))
    return token


def enviar_correo(token, asunto, html, para, cc, excel: Path, imagenes) -> None:
    adjuntos = [_adjunto_graph(excel.name, excel.read_bytes())]
    adjuntos += [_adjunto_graph(nombre, contenido, cid) for nombre, contenido, cid in imagenes]
    peso = sum(len(a["contentBytes"]) for a in adjuntos)
    if peso > LIMITE_GRAPH:
        raise RuntimeError(f"Los adjuntos pesan {peso / 1e6:.1f} MB en base64 y Graph admite ~4 MB "
                           "por correo. Reduce el tamaño de la firma.")
    payload = {
        "message": {
            "subject": asunto,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": d}} for d in para],
            "ccRecipients": [{"emailAddress": {"address": d}} for d in cc],
            "attachments": adjuntos,
        },
        "saveToSentItems": True,
    }
    ruta = "/me/sendMail" if MAIL_AUTH == "delegado" else f"/users/{os.environ['MAIL_SENDER']}/sendMail"
    resp = requests.post(
        f"{GRAPH}{ruta}",
        headers={"Authorization": f"Bearer {token}"},
        json=payload,
        timeout=120,
    )
    if resp.status_code != 202:
        raise RuntimeError(f"Error al enviar ({resp.status_code}): {resp.text}")
    log.info("Correo enviado desde %s a %s%s", os.environ["MAIL_SENDER"], ", ".join(para),
             f" (CC: {', '.join(cc)})" if cc else "")


# ---------- Verificación de la configuración ----------
def verificar(para, cc) -> bool:
    ok = True

    def paso(nombre, fn):
        nonlocal ok
        try:
            detalle = fn()
            log.info("[OK]    %s%s", nombre, f" - {detalle}" if detalle else "")
        except Exception as e:  # noqa: BLE001 - se reporta cualquier fallo
            ok = False
            log.error("[FALLA] %s - %s", nombre, e)

    def api():
        usuario = erpnext_get("/api/method/frappe.auth.get_logged_user", intentos=1).get("message")
        corte = fecha_de_corte(None)
        params = {"desde": corte.replace(day=1).isoformat(), "corte": corte.isoformat(), "company": ERPNEXT_EMPRESA}
        datos = (erpnext_get(METODO_SNAPSHOT, params, intentos=1).get("message") or {}).get("data")
        if not datos:
            raise RuntimeError("monthly_subscription_snapshot no devolvió datos")
        return f"{ERPNEXT_URL} como {usuario}; snapshot al {corte}: {len(datos)} suscripciones"

    def archivos():
        faltan = [str(p) for p in (PLANTILLA_HTML, MENSAJE_PATH, FIRMA_PATH) if not p.exists()]
        if faltan:
            raise FileNotFoundError(", ".join(faltan))

    def destinatarios():
        if not para:
            raise ValueError("TV_MAIL_TO está vacío")
        return f"Para: {'; '.join(para)} | CC: {'; '.join(cc) or '-'}"

    paso("API de ERPNext (ERPNEXT_API_KEY / SECRET)", api)
    modo = "sesión de la cuenta (delegado)" if MAIL_AUTH == "delegado" else "permiso de aplicación"
    paso(f"Azure: Mail.Send con {modo}", lambda: obtener_token() and f"remitente {os.environ['MAIL_SENDER']}")
    paso("Archivos (plantilla, mensaje, firma)", archivos)
    paso("Destinatarios", destinatarios)
    return ok


# ---------- Main ----------
def main() -> None:
    ap = argparse.ArgumentParser(description="Reporte mensual de suscripciones de TV por correo")
    ap.add_argument("--corte", help="fecha de corte AAAA-MM-DD (por defecto, fin de mes)")
    ap.add_argument("--para", help="destinatarios separados por ; (reemplaza TV_MAIL_TO)")
    ap.add_argument("--cc", help="con copia separados por ; (reemplaza TV_MAIL_CC; \"\" para ninguno)")
    ap.add_argument("--mensaje", help="texto del correo (reemplaza el archivo TV_MENSAJE)")
    ap.add_argument("--sin-enviar", action="store_true", help="genera el Excel y la vista previa, sin enviar")
    ap.add_argument("--forzar", action="store_true", help="reenvía aunque el corte ya se haya enviado")
    ap.add_argument("--foto", help="arma el reporte desde una foto guardada (.json.gz) en vez de consultar ERPNext")
    ap.add_argument("--verificar", action="store_true", help="solo revisa la configuración")
    ap.add_argument("--iniciar-sesion", action="store_true", help="modo delegado: muestra el código para iniciar sesión")
    ap.add_argument("--completar-sesion", action="store_true", help="modo delegado: espera el inicio de sesión y lo guarda")
    args = ap.parse_args()

    if args.iniciar_sesion:
        return iniciar_sesion()
    if args.completar_sesion:
        return completar_sesion()

    para = _lista(args.para) if args.para is not None else _lista(os.getenv("TV_MAIL_TO"))
    cc = _lista(args.cc) if args.cc is not None else _lista(os.getenv("TV_MAIL_CC"))

    if args.verificar:
        raise SystemExit(0 if verificar(para, cc) else 1)

    corte = fecha_de_corte(args.corte)
    clave = corte.isoformat()
    estado = cargar_estado()
    log.info("=== Reporte suscripciones TV - corte %s ===", f"{corte:%d/%m/%Y}")

    if not args.sin_enviar:
        if estado.get(clave, {}).get("enviado") and not args.forzar:
            log.info("El corte %s ya se envió el %s. No se hace nada (usa --forzar para reenviar).",
                     clave, estado[clave]["enviado"])
            return
        if not para:
            raise RuntimeError("No hay destinatarios: configura TV_MAIL_TO en .env o usa --para")

    datos = leer_foto(_ruta(args.foto), corte) if args.foto else descargar_snapshot(corte)
    filas = armar_filas(por_plan(datos))
    destino = OUTPUT_DIR / f"planes tv {MESES_LARGO[corte.month - 1]} {corte.year}.xlsx"
    generar_excel(filas, corte, destino)

    asunto, html, imagenes = construir_correo(corte, len(filas), args.mensaje)

    if args.sin_enviar:
        vista = guardar_vista_previa(asunto, html, imagenes, para, cc, destino)
        log.info("Modo --sin-enviar: vista previa del correo en %s", vista)
        return

    enviar_correo(obtener_token(), asunto, html, para, cc, destino, imagenes)
    estado[clave] = {"archivo": destino.name, "enviado": f"{ahora():%Y-%m-%d %H:%M:%S}",
                     "para": para, "cc": cc, "suscripciones": len(filas)}
    guardar_estado(estado)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        log.exception("Falló la generación o el envío del reporte de suscripciones TV")
        raise SystemExit(1)
