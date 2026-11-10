"""
Genera un reporte automáticamente y lo envía por Outlook (Microsoft 365)
usando Microsoft Graph API con autenticación de aplicación (sin usuario interactivo).

Requisitos: pip install msal requests pandas openpyxl sqlalchemy python-dotenv
Configuración: ver .env.example
"""

import base64
import logging
import os
from datetime import date, datetime
from pathlib import Path

import msal
import pandas as pd
import requests
from dotenv import load_dotenv
from sqlalchemy import create_engine

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("report_mailer")

# ---------- Configuración (desde variables de entorno) ----------
TENANT_ID = os.environ["AZURE_TENANT_ID"]
CLIENT_ID = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
SENDER = os.environ["MAIL_SENDER"]                     # buzón que envía, ej: reportes@empresa.com
RECIPIENTS = [r.strip() for r in os.environ["MAIL_TO"].split(";") if r.strip()]
DB_URL = os.environ["DB_URL"]                          # ej: mysql+pymysql://user:pass@host/db
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "./reportes"))
TEMPLATE_PATH = Path(os.getenv("TEMPLATE_PATH", "./plantilla.html"))
# Imágenes de la firma: nombre_cid=ruta;otro_cid=ruta  (en el HTML: <img src="cid:nombre_cid">)
INLINE_IMAGES = {
    k.strip(): Path(v.strip())
    for k, v in (p.split("=", 1) for p in os.getenv("INLINE_IMAGES", "").split(";") if "=" in p)
}

GRAPH = "https://graph.microsoft.com/v1.0"

MIME = {
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pdf": "application/pdf",
}


# ---------- 1. Generar el reporte ----------
def generar_reporte(fecha: date) -> Path:
    """Ejecuta la consulta y guarda un Excel. Ajusta el SQL a tu caso."""
    sql = """
        SELECT fecha, cliente, producto, cantidad, total
        FROM ventas
        WHERE fecha = CURDATE() - INTERVAL 1 DAY
        ORDER BY total DESC
    """
    engine = create_engine(DB_URL)
    df = pd.read_sql(sql, engine)
    log.info("Consulta devolvió %d filas", len(df))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ruta = OUTPUT_DIR / f"reporte_ventas_{fecha:%Y-%m-%d}.xlsx"

    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ventas", index=False)
        ws = writer.sheets["Ventas"]
        for col in ws.columns:  # ancho de columnas automático
            ancho = max(len(str(c.value)) for c in col if c.value is not None) + 2
            ws.column_dimensions[col[0].column_letter].width = min(ancho, 50)

    log.info("Reporte guardado en %s", ruta)
    return ruta


# ---------- 2. Obtener token de Graph ----------
def obtener_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"No se pudo obtener token: {result.get('error_description')}")
    return result["access_token"]


# ---------- 3. Enviar el correo ----------
def _adjunto(ruta: Path, cid: str | None = None) -> dict:
    """Construye un adjunto de Graph. Con cid, va incrustado en el HTML (firma, logo)."""
    item = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": ruta.name,
        "contentType": MIME.get(ruta.suffix.lower(), "application/octet-stream"),
        "contentBytes": base64.b64encode(ruta.read_bytes()).decode(),
    }
    if cid:
        item["contentId"] = cid
        item["isInline"] = True
    return item


def construir_cuerpo(**datos) -> str:
    """Rellena plantilla.html (diseño + firma) con los datos del reporte."""
    plantilla = TEMPLATE_PATH.read_text(encoding="utf-8")
    for clave, valor in datos.items():
        plantilla = plantilla.replace("{{" + clave + "}}", str(valor))
    return plantilla


def enviar_correo(token: str, asunto: str, cuerpo_html: str, adjunto: Path) -> None:
    adjuntos = [_adjunto(adjunto)]
    adjuntos += [_adjunto(ruta, cid) for cid, ruta in INLINE_IMAGES.items() if ruta.exists()]
    payload = {
        "message": {
            "subject": asunto,
            "body": {"contentType": "HTML", "content": cuerpo_html},
            "toRecipients": [{"emailAddress": {"address": r}} for r in RECIPIENTS],
            "attachments": adjuntos,
        },
        "saveToSentItems": True,
    }
    resp = requests.post(
        f"{GRAPH}/users/{SENDER}/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    if resp.status_code != 202:
        raise RuntimeError(f"Error al enviar ({resp.status_code}): {resp.text}")
    log.info("Correo enviado a %s", ", ".join(RECIPIENTS))


# ---------- Main ----------
def main() -> None:
    hoy = date.today()
    reporte = generar_reporte(hoy)

    asunto = f"Reporte de ventas – {hoy:%d/%m/%Y}"
    cuerpo = construir_cuerpo(
        fecha=f"{hoy:%d/%m/%Y}",
        generado=f"{datetime.now():%d/%m/%Y %H:%M}",
    )

    token = obtener_token()
    enviar_correo(token, asunto, cuerpo, reporte)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log.exception("Falló la generación o el envío del reporte")
        raise SystemExit(1)
