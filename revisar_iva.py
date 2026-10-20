"""
Revisión diaria de pedidos sin IVA - Tienda online Real Zaragoza (Magento 2)

Descarga los pedidos del día anterior (hora de Madrid), filtra los envíos a
domicilio y avisa por email de los que no llevan IVA, con el Excel adjunto.

Variables de entorno:
  Magento
    MAGENTO_BASE_URL      (opcional) por defecto https://tienda.realzaragoza.com
    MAGENTO_TOKEN         (recomendado) access token de una Integración
    MAGENTO_USER / MAGENTO_PASSWORD   alternativa: token de administrador
  Email
    SMTP_HOST             (opcional) por defecto smtp.gmail.com
    SMTP_PORT             (opcional) por defecto 587
    SMTP_USER             cuenta que envía (p. ej. avisos@realzaragoza.com)
    SMTP_PASSWORD         contraseña de aplicación de esa cuenta
    MAIL_FROM             (opcional) por defecto SMTP_USER
    MAIL_TO               destinatarios separados por comas
  Comportamiento
    FECHA                 (opcional) día a revisar YYYY-MM-DD; vacío = ayer
    ENVIAR_SI_VACIO       (opcional) "true" para mandar email aunque no haya incidencias
    STATUS_EXCLUIDOS      (opcional) estados a ignorar, separados por comas
    SHIPPING_DESCRIPTION  (opcional) método de envío a revisar
    CRON_DISPARADOR       (lo rellena GitHub Actions) para ejecutar solo a las 9:00 de Madrid
"""

import io
import os
import smtplib
import ssl
import sys
from datetime import date, datetime, time, timedelta
from email.message import EmailMessage
from zoneinfo import ZoneInfo

import pandas as pd
import requests

# ==========================================
# CONFIGURACIÓN
# ==========================================
BASE_URL = os.getenv("MAGENTO_BASE_URL", "https://tienda.realzaragoza.com").rstrip("/")
TZ_MADRID = ZoneInfo("Europe/Madrid")
TZ_UTC = ZoneInfo("UTC")
PAGE_SIZE = 100
TIMEOUT = 60  # segundos por petición
IVA = 0.21

# Los pedidos de ayer a las 9:00 casi nunca están "closed" todavía, así que en vez
# de una lista de estados válidos se descartan los que no son venta real.
STATUS_EXCLUIDOS = {
    s.strip()
    for s in os.getenv("STATUS_EXCLUIDOS", "canceled,pending_payment,paypal_reversed").split(",")
    if s.strip()
}

SHIPPING_DESCRIPTION = os.getenv("SHIPPING_DESCRIPTION", "a domicilio -  Envío")


def normalizar(texto):
    """Minúsculas y espacios colapsados, para no fallar por un doble espacio."""
    return " ".join(str(texto or "").split()).lower()


# ==========================================
# 0. ¿TOCA EJECUTAR? (solo para GitHub Actions)
# ==========================================
def debe_ejecutarse():
    """
    GitHub Actions programa en UTC. Hay dos crons (07:00 y 08:00 UTC) para cubrir
    horario de verano e invierno; aquí se deja pasar solo el que equivale a las
    9:00 en Madrid hoy. Funciona aunque GitHub retrase la ejecución.
    """
    cron = os.getenv("CRON_DISPARADOR", "").strip()
    if not cron:
        return True  # ejecución manual o en otro entorno
    hora_cron_utc = int(cron.split()[1])
    offset_madrid = int(datetime.now(TZ_MADRID).utcoffset().total_seconds() // 3600)
    return hora_cron_utc == 9 - offset_madrid


# ==========================================
# 1. AUTENTICACIÓN
# ==========================================
def obtener_headers():
    token = os.getenv("MAGENTO_TOKEN")
    if not token:
        resp = requests.post(
            f"{BASE_URL}/rest/V1/integration/admin/token",
            json={"username": os.getenv("MAGENTO_USER"), "password": os.getenv("MAGENTO_PASSWORD")},
            timeout=TIMEOUT,
        )
        if resp.status_code != 200:
            raise RuntimeError(f"Error de autenticación en Magento ({resp.status_code}): {resp.text}")
        token = resp.json()
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


# ==========================================
# 2. RANGO DE FECHAS (día anterior, hora de Madrid -> UTC)
# ==========================================
def rango_dia(dia):
    """Magento guarda created_at en UTC; convertimos el día de Madrid a UTC."""
    inicio = datetime.combine(dia, time.min, tzinfo=TZ_MADRID).astimezone(TZ_UTC)
    fin = (datetime.combine(dia, time.min, tzinfo=TZ_MADRID) + timedelta(days=1)).astimezone(TZ_UTC)
    fmt = "%Y-%m-%d %H:%M:%S"
    return inicio.strftime(fmt), fin.strftime(fmt)


# ==========================================
# 3. DESCARGA PAGINADA
# ==========================================
def descargar_pedidos(headers, date_from, date_to):
    url = f"{BASE_URL}/rest/V1/orders"
    pedidos, page = [], 1
    while True:
        params = {
            "searchCriteria[filter_groups][0][filters][0][field]": "created_at",
            "searchCriteria[filter_groups][0][filters][0][value]": date_from,
            "searchCriteria[filter_groups][0][filters][0][condition_type]": "gteq",
            "searchCriteria[filter_groups][1][filters][0][field]": "created_at",
            "searchCriteria[filter_groups][1][filters][0][value]": date_to,
            "searchCriteria[filter_groups][1][filters][0][condition_type]": "lt",
            "searchCriteria[currentPage]": page,
            "searchCriteria[pageSize]": PAGE_SIZE,
        }
        resp = requests.get(url, headers=headers, params=params, timeout=TIMEOUT)
        if resp.status_code != 200:
            raise RuntimeError(f"Error {resp.status_code} en la página {page}: {resp.text}")

        data = resp.json()
        items = data.get("items", []) or []
        pedidos.extend(items)
        print(f"Página {page}: {len(items)} pedidos")

        # Magento devuelve la última página otra vez si te pasas: usamos total_count
        if not items or page * PAGE_SIZE >= data.get("total_count", 0):
            break
        page += 1
    return pedidos


# ==========================================
# 4. ANÁLISIS
# ==========================================
def a_hora_madrid(created_at):
    try:
        dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ_UTC)
        return dt.astimezone(TZ_MADRID).strftime("%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return created_at


def extraer_fila(o):
    ext = o.get("extension_attributes") or {}

    sa = ext.get("shipping_assignments") or []
    address = ((sa[0].get("shipping") or {}).get("address") or {}) if sa else {}

    taxes = ext.get("taxes") or ext.get("applied_taxes") or []
    tax = taxes[0] if isinstance(taxes, list) and taxes else {}

    nombre = f"{o.get('customer_firstname') or ''} {o.get('customer_lastname') or ''}".strip()

    return {
        "id": o.get("increment_id"),
        "status": o.get("status"),
        "fecha (Madrid)": a_hora_madrid(o.get("created_at")),
        "cliente": nombre or pd.NA,
        "email cliente": o.get("customer_email", pd.NA),
        "pagado": float(o.get("base_grand_total") or 0),
        "tax_amount pedido": float(o.get("base_tax_amount") or 0),
        "city": address.get("city", pd.NA),
        "region": address.get("region", pd.NA),
        "country": address.get("country_id", pd.NA),
        "title": tax.get("title", pd.NA),
        "percent": tax.get("percent", pd.NA),
        "amount": tax.get("amount", pd.NA),
    }


def pedidos_sin_iva(pedidos):
    envio_objetivo = normalizar(SHIPPING_DESCRIPTION)
    enviados = [
        o for o in pedidos
        if o.get("status") not in STATUS_EXCLUIDOS
        and normalizar(o.get("shipping_description")) == envio_objetivo
    ]
    print(f"Envíos a domicilio válidos: {len(enviados)}")

    df = pd.DataFrame([extraer_fila(o) for o in enviados])
    if df.empty:
        return df

    # Sin IVA = no hay impuesto aplicado o el impuesto del pedido es 0
    mask = df["title"].isna() | (df["tax_amount pedido"].round(2) == 0)
    df = df[mask].copy()
    df.insert(loc=6, column="supuesto IVA", value=(df["pagado"] - df["pagado"] / (1 + IVA)).round(2))
    return df.reset_index(drop=True)


# ==========================================
# 5. EMAIL
# ==========================================
def enviar_email(asunto, html, adjunto=None, nombre_adjunto=None):
    destinatarios = [d.strip() for d in os.getenv("MAIL_TO", "").split(",") if d.strip()]
    if not destinatarios:
        raise RuntimeError("MAIL_TO está vacío: no hay a quién enviar el aviso")

    smtp_user = os.environ["SMTP_USER"]
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = os.getenv("MAIL_FROM", smtp_user)
    msg["To"] = ", ".join(destinatarios)
    msg.set_content("Este mensaje necesita un cliente de correo con HTML.")
    msg.add_alternative(html, subtype="html")
    if adjunto is not None:
        msg.add_attachment(
            adjunto,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=nombre_adjunto,
        )

    with smtplib.SMTP(os.getenv("SMTP_HOST", "smtp.gmail.com"), int(os.getenv("SMTP_PORT", "587")), timeout=TIMEOUT) as s:
        s.starttls(context=ssl.create_default_context())
        s.login(smtp_user, os.environ["SMTP_PASSWORD"])
        s.send_message(msg)
    print(f"Email enviado a: {', '.join(destinatarios)}")


def html_informe(df, dia, total_pedidos):
    fecha_txt = dia.strftime("%d/%m/%Y")
    if df.empty:
        return f"<p>Revisados {total_pedidos} pedidos del {fecha_txt}: <b>ningún envío sin IVA</b>.</p>"

    columnas = ["id", "status", "fecha (Madrid)", "cliente", "pagado", "supuesto IVA", "city", "region", "country"]
    tabla = df[columnas].to_html(index=False, border=0, na_rep="—", float_format=lambda x: f"{x:,.2f} €")
    return f"""
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#222">
      <p>Se han detectado <b>{len(df)} envío(s) sin IVA</b> en los pedidos del <b>{fecha_txt}</b>
         (de {total_pedidos} pedidos revisados).</p>
      <p>IVA no repercutido estimado: <b>{df['supuesto IVA'].sum():,.2f} €</b></p>
      <style>
        table {{border-collapse:collapse}}
        th {{background:#0b3d91;color:#fff;padding:6px 10px;text-align:left}}
        td {{border-bottom:1px solid #ddd;padding:6px 10px}}
      </style>
      {tabla}
      <p style="color:#666;font-size:12px">Detalle completo en el Excel adjunto. Aviso automático.</p>
    </div>"""


# ==========================================
# MAIN
# ==========================================
def main():
    if not debe_ejecutarse():
        print("No son las 9:00 en Madrid para este cron; se omite la ejecución.")
        return

    fecha_env = os.getenv("FECHA", "").strip()
    dia = date.fromisoformat(fecha_env) if fecha_env else datetime.now(TZ_MADRID).date() - timedelta(days=1)
    date_from, date_to = rango_dia(dia)
    print(f"Revisando pedidos del {dia} (UTC {date_from} -> {date_to})")

    headers = obtener_headers()
    pedidos = descargar_pedidos(headers, date_from, date_to)
    print(f"Total de pedidos descargados: {len(pedidos)}")

    df = pedidos_sin_iva(pedidos)
    print(f"Envíos sin IVA: {len(df)}")

    enviar_si_vacio = os.getenv("ENVIAR_SI_VACIO", "false").lower() == "true"
    if df.empty and not enviar_si_vacio:
        print("Sin incidencias; no se envía email.")
        return

    adjunto = nombre = None
    if not df.empty:
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False)
        adjunto, nombre = buffer.getvalue(), f"Pedidos sin IVA {dia}.xlsx"
        asunto = f"⚠️ {len(df)} envío(s) sin IVA - pedidos del {dia.strftime('%d/%m/%Y')}"
    else:
        asunto = f"✅ Sin envíos sin IVA - pedidos del {dia.strftime('%d/%m/%Y')}"

    enviar_email(asunto, html_informe(df, dia, len(pedidos)), adjunto, nombre)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"❌ {e}")
        try:
            enviar_email(
                "❌ Error en la revisión diaria de pedidos sin IVA",
                f"<p>La revisión automática ha fallado:</p><pre>{e}</pre>",
            )
        except Exception as e_mail:
            print(f"Tampoco se pudo enviar el aviso de error: {e_mail}")
        sys.exit(1)
