# Correo automático - reporte de abonados TV mensual

Cada fin de mes genera el Excel de abonados de TV de Giga Fibra (suscripciones **Active** y **Past Due Date** al corte) y lo envía por Outlook (Microsoft Graph) con un mensaje personalizado y la firma de quien envía.

- **Cuándo:** último día de cada mes a las 23:50, hora de Bogotá.
- **Qué se envía:** `planes tv {mes} {año}.xlsx`, asunto `Reporte de Número de Abonados TV {mes} {año}`.
- **Desde / para:** se configura en `.env` (`MAIL_SENDER`, `TV_MAIL_TO`, `TV_MAIL_CC`).

## El reporte

Una hoja `planes tv {mes}` con la tabla `Tabla1`, una fila por suscripción:

| Suscripcion | Estado | Centro poblado | Ciudad | Codigo Dane municipio | Estrato | custom_television | Plan |
|---|---|---|---|---|---|---|---|

Reglas:

- **Planes** (`TV_PLANES`, en orden de prioridad): WinSports + 2 SetBox, WinSports + 1 SetBox, 1 SetBox Set-2024, 1 SetBox May-2024 y Firestick. No incluye TV Free, Disney ni Discovery.
- **Una fila por suscripción.** Si tiene dos de esos planes, queda el de mayor prioridad.
- **custom_television = 1** en todas las filas: Firestick tiene TV = 0 en ERPNext, pero en este reporte cuenta como TV.
- **Estado al corte:** se toma `status_al_corte` del método de ERPNext `monthly_subscription_snapshot` (el mismo del portal CRC), que reconstruye el estado a la fecha de corte.
- **DANE:** el código del municipio viene de ERPNext (o de la ciudad si falta); el centro poblado es la cabecera municipal (código + `000`). Si aparece una ciudad sin código, el log lo avisa; se agrega en `DANE_MUNICIPIOS`.

El formato se validó contra el reporte de septiembre de 2026: con los mismos datos, el script produce exactamente el mismo archivo (2.100 filas).

## Archivos

| Archivo | Para qué |
|---|---|
| `reporte_planes_tv.py` | Script principal: descarga, genera el Excel y envía |
| `mensaje_tv.txt` | Texto del correo (variables `{mes}`, `{año}`, `{fecha_corte}`, `{total}`) |
| `plantilla_tv.html` | Diseño del correo (`{{mensaje}}`, `{{firma}}`) |
| `firma.jpg` | Firma de quien envía |
| `.env.example` | Configuración de ejemplo (copiar a `.env`) |
| `deploy/` | Servicio y temporizador de systemd del servidor |
| `report_mailer.py`, `plantilla.html` | Plantilla genérica original; el reporte de TV no la usa |

## Uso

```bash
pip install -r requirements.txt
cp .env.example .env   # y completar

python reporte_planes_tv.py --verificar          # revisa API de ERPNext, correo, archivos y destinatarios
python reporte_planes_tv.py --sin-enviar         # genera el Excel y reportes/vista_previa_correo.html
python reporte_planes_tv.py                      # corte del mes: genera y envía
python reporte_planes_tv.py --corte 2026-10-31 --forzar
python reporte_planes_tv.py --corte 2026-10-31 --foto /ruta/2026-10.json.gz --forzar
```

- **Fecha de corte:** si se ejecuta el último día del mes, es ese día; cualquier otro día se toma como el corte del mes anterior.
- **Sin reenvíos:** un corte ya enviado no se vuelve a enviar, salvo con `--forzar` (queda registrado en `reportes/estado_planes_tv.json`).
- **`--foto`:** arma el reporte desde una foto guardada de la misma API. El portal CRC guarda una cada cierre en `/var/www/html/portal_reportes_crc/backend/data/fotos/AAAA-MM.json.gz`; así se generó octubre de 2026 con los datos del 31/10 a las 23:50.

## Envío de correo

`MAIL_AUTH=delegado` envía con la cuenta de `MAIL_SENDER` (permiso `Mail.Send` delegado, sin administrador de Azure):

1. En la app de Azure: *Authentication* → **Permitir flujos de clientes públicos** = Sí.
2. Iniciar sesión una vez:
   ```bash
   python reporte_planes_tv.py --iniciar-sesion     # muestra el código para login.microsoft.com/device
   python reporte_planes_tv.py --completar-sesion   # espera el inicio de sesión y lo guarda
   ```
3. La sesión queda en `.sesion_correo.json` (no se sube a Git) y se renueva con cada envío. Hay que repetir el paso 2 si pasan 90 días sin uso o si cambia la contraseña.

`MAIL_AUTH=aplicacion` usa el secreto de la app (permiso `Mail.Send` de aplicación con consentimiento de administrador).

## Servidor

- **Ruta:** `/var/www/html/monthly_tv_reports` en `10.20.67.35` (Rocky Linux).
- **Protección:** la carpeta está dentro del webroot de nginx, pero tiene permisos `700` de root (y `.env` `600`), así que nginx no puede servir sus archivos (responde 404/403) sin cambiar su configuración.
- **Python:** se ejecuta con `/usr/bin/python3.12` y `PYTHONPATH=.venv/lib/python3.12/site-packages`, porque ejecutar binarios de `.venv/bin` dentro de `/var/www/html` choca con SELinux.
- **Programación:** `deploy/reporte-planes-tv.service` y `.timer` instalados en `/etc/systemd/system`. Si el servidor está apagado a la hora del envío, corre al encender (`Persistent=true`).

```bash
systemctl list-timers reporte-planes-tv.timer              # próxima ejecución
journalctl -u reporte-planes-tv.service                    # salida de cada ejecución
tail /var/www/html/monthly_tv_reports/reportes/reporte_planes_tv.log
systemctl disable --now reporte-planes-tv.timer            # desactivar
```

**Actualizar el código:** copiar los archivos cambiados a la carpeta del servidor (sin tocar `.env`, `.sesion_correo.json`, `.venv` ni `reportes/`).
