# -*- coding: utf-8 -*-
import os
import json
import base64
import http.client
from datetime import datetime
from urllib.parse import quote
import requests
import pandas as pd
from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from playwright.sync_api import sync_playwright

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}


def obtener_datos_dolar():
    conn = http.client.HTTPSConnection("proxy.icap.com.co")
    headers = {
        "Accept": "application/json",
        "Referer": "https://dolar.icap.com.co/dolar/",
        "User-Agent": "Mozilla/5.0",
    }
    conn.request("GET", "/api/v1/ws/public/delay/snapshot?market=71", headers=headers)
    res = conn.getresponse()
    if res.status != 200:
        raise RuntimeError(f"Error HTTP {res.status}: {res.reason}")

    res_json = json.loads(res.read().decode("utf-8"))
    msg_grafico = next(m for m in res_json["messages"] if m.get("id") == 1000)
    a = msg_grafico["result"][0]["datos_grafico_moneda_mercado"]

    sep1 = a.split("data: [")
    data_dolar = sep1[1].split("],type:")[0]
    hora_raw = sep1[2].split("labels: [")[1].split("]}")[0]

    precios = [float(p.strip()) for p in data_dolar.split(",") if p.strip()]
    horas = [h.strip().strip("'\"") for h in hora_raw.split(",") if h.strip()]

    n = min(len(precios), len(horas))
    ahora = datetime.now()
    hoy_str = ahora.strftime("%Y-%m-%d")

    df = pd.DataFrame({"Hora": horas[:n], "Dolar": precios[:n]})
    df["FechaHora"] = pd.to_datetime(hoy_str + " " + df["Hora"])
    df = df.set_index("FechaHora").sort_index()
    return df, ahora


def resumir_intervalo(df, freq):
    resumen = df["Dolar"].resample(freq).agg(
        Apertura="first",
        Maximo="max",
        Minimo="min",
        Cierre="last",
    ).dropna(subset=["Apertura"])
    resumen.index = resumen.index.strftime("%H:%M")
    resumen = resumen.reset_index().rename(columns={"FechaHora": "Hora"})
    return resumen


def capturar_dashboard_set_icap(ruta_png):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1600, "height": 1050})
        page.goto("https://dolar.icap.com.co/dolar/", wait_until="networkidle", timeout=60000)
        page.wait_for_selector("text=PRECIOS", timeout=30000)
        page.wait_for_selector("text=VOLUMEN TRANSADO", timeout=30000)
        page.wait_for_timeout(4000)
        page.screenshot(
            path=ruta_png,
            clip={"x": 0, "y": 115, "width": 1600, "height": 850},
        )
        browser.close()


def escribir_tabla_estilizada(ws, df_tabla, fila_inicio, col_inicio, titulo, color_titulo="1F4E78"):
    thin = Side(border_style="thin", color="D0D7DE")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)

    fill_titulo = PatternFill("solid", fgColor=color_titulo)
    fill_header = PatternFill("solid", fgColor="263238")
    fill_par = PatternFill("solid", fgColor="F4F6F8")
    fill_impar = PatternFill("solid", fgColor="FFFFFF")

    font_titulo = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_dato = Font(name="Calibri", size=10, color="111827")

    num_cols = len(df_tabla.columns)
    col_fin = col_inicio + num_cols - 1

    ws.merge_cells(start_row=fila_inicio, start_column=col_inicio, end_row=fila_inicio, end_column=col_fin)
    c_tit = ws.cell(row=fila_inicio, column=col_inicio, value=titulo)
    c_tit.fill = fill_titulo
    c_tit.font = font_titulo
    c_tit.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[fila_inicio].height = 24

    fila_hdr = fila_inicio + 1
    ws.row_dimensions[fila_hdr].height = 20
    for j, col_name in enumerate(df_tabla.columns):
        cell = ws.cell(row=fila_hdr, column=col_inicio + j, value=col_name)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = borde

    for i, row in enumerate(df_tabla.itertuples(index=False)):
        fila_actual = fila_hdr + 1 + i
        fill_fila = fill_par if i % 2 == 0 else fill_impar
        for j, val in enumerate(row):
            cell = ws.cell(row=fila_actual, column=col_inicio + j, value=val)
            cell.fill = fill_fila
            cell.font = font_dato
            cell.border = borde
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for j, col_name in enumerate(df_tabla.columns):
        col_letter = get_column_letter(col_inicio + j)
        ancho = max(11, len(str(col_name)) + 3)
        ws.column_dimensions[col_letter].width = max(ws.column_dimensions[col_letter].width or 0, ancho)


def poblar_hoja_reporte(ws, df, fecha_str, ruta_png):
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:AC1")
    c_main = ws["A1"]
    c_main.value = f"REPORTE DIARIO DÓLAR SET-FX (USD/COP) - {fecha_str}"
    c_main.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    c_main.fill = PatternFill("solid", fgColor="0F172A")
    c_main.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    resumen_diario = pd.DataFrame([{
        "Fecha": fecha_str,
        "Apertura": df["Dolar"].iloc[0],
        "Maximo": df["Dolar"].max(),
        "Minimo": df["Dolar"].min(),
        "Cierre": df["Dolar"].iloc[-1],
    }])
    escribir_tabla_estilizada(ws, resumen_diario, fila_inicio=3, col_inicio=1,
                              titulo="RESUMEN DIARIO", color_titulo="0D9488")

    img = XLImage(ruta_png)
    img.width = 960
    img.height = 510
    ws.add_image(img, "A7")

    tablas = [
        ("CADA 10 MINUTOS", resumir_intervalo(df, "10min"), "1F4E78"),
        ("CADA 30 MINUTOS", resumir_intervalo(df, "30min"), "2E75B6"),
        ("CADA 60 MINUTOS", resumir_intervalo(df, "60min"), "385723"),
        ("CADA 5 MINUTOS", resumir_intervalo(df, "5min"), "7F6000"),
        ("CADA 15 MINUTOS", resumir_intervalo(df, "15min"), "833C0C"),
    ]

    col_actual = 1
    for titulo, tabla_df, color in tablas:
        escribir_tabla_estilizada(ws, tabla_df, fila_inicio=34, col_inicio=col_actual,
                                  titulo=titulo, color_titulo=color)
        col_separador = get_column_letter(col_actual + len(tabla_df.columns))
        ws.column_dimensions[col_separador].width = 3
        col_actual += len(tabla_df.columns) + 1


def generar_excel_diario(df, fecha_str, ruta_png, ruta_xlsx):
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte_Dolar"
    poblar_hoja_reporte(ws, df, fecha_str, ruta_png)
    wb.save(ruta_xlsx)


def actualizar_excel_mensual(df, fecha_str, ruta_png, ruta_xlsx_mensual):
    if os.path.exists(ruta_xlsx_mensual):
        wb = load_workbook(ruta_xlsx_mensual)
        if fecha_str in wb.sheetnames:
            del wb[fecha_str]
        ws = wb.create_sheet(title=fecha_str)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = fecha_str

    poblar_hoja_reporte(ws, df, fecha_str, ruta_png)
    wb.save(ruta_xlsx_mensual)


# --- INTEGRACIÓN MICROSOFT GRAPH (ONEDRIVE + CORREO) ---
def obtener_token_graph():
    tenant_id = os.getenv("AZURE_TENANT_ID")
    client_id = os.getenv("AZURE_CLIENT_ID")
    client_secret = os.getenv("AZURE_CLIENT_SECRET")

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }
    r = requests.post(url, data=payload, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def construir_ruta_onedrive(anio_str, nombre_archivo):
    base = (os.getenv("ONEDRIVE_BASE_FOLDER") or "").strip("/")
    partes = [p for p in [base, anio_str, nombre_archivo] if p]
    return "/".join(partes)


def descargar_mensual_si_existe(token, user_email, ruta_remota, ruta_local):
    url = f"https://graph.microsoft.com/v1.0/users/{quote(user_email)}/drive/root:/{quote(ruta_remota)}:/content"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code == 200:
        with open(ruta_local, "wb") as f:
            f.write(r.content)


def subir_a_onedrive(token, user_email, ruta_remota, ruta_local):
    url = f"https://graph.microsoft.com/v1.0/users/{quote(user_email)}/drive/root:/{quote(ruta_remota)}:/content"
    with open(ruta_local, "rb") as f:
        r = requests.put(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            },
            data=f,
            timeout=120,
        )
    r.raise_for_status()
    return r.json().get("webUrl", "")


def obtener_url_carpeta_onedrive(token, user_email):
    base = (os.getenv("ONEDRIVE_BASE_FOLDER") or "").strip("/")
    if base:
        url = f"https://graph.microsoft.com/v1.0/users/{quote(user_email)}/drive/root:/{quote(base)}"
    else:
        url = f"https://graph.microsoft.com/v1.0/users/{quote(user_email)}/drive/root"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if r.status_code == 200:
        return r.json().get("webUrl", "")
    return ""


def enviar_correo_graph(token, mail_from, mail_to_str, df, fecha_str, mes_nombre, anio_str, ruta_xlsx_dia, web_url_carpeta):
    apertura = df["Dolar"].iloc[0]
    maximo = df["Dolar"].max()
    minimo = df["Dolar"].min()
    cierre = df["Dolar"].iloc[-1]

    with open(ruta_xlsx_dia, "rb") as f:
        adjunto_b64 = base64.b64encode(f.read()).decode("utf-8")

    bloque_onedrive = ""
    if web_url_carpeta:
        bloque_onedrive = f"""
        <p style="margin-top:16px;">
          El consolidado mensual y el histórico de reportes se encuentran disponibles en la carpeta de OneDrive:
        </p>
        <p>
          <a href="{web_url_carpeta}" style="background-color:#0078D4;color:#FFFFFF;padding:10px 18px;text-decoration:none;border-radius:6px;display:inline-block;font-weight:bold;">
            Abrir Carpeta de Reportes en OneDrive
          </a>
        </p>
        """

    cuerpo_html = f"""
    <div style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#1F2937;">
      <p>Buen día Don Hernan,</p>
      <p>Envío adjunto el <b>Reporte Diario del Dólar SET-FX (USD/COP)</b> correspondiente al <b>{fecha_str}</b>.</p>

      <table style="border-collapse:collapse;margin:14px 0;min-width:420px;">
        <thead>
          <tr style="background-color:#0078D4;color:#FFFFFF;">
            <th style="padding:8px 14px;border:1px solid #D1D5DB;">Apertura</th>
            <th style="padding:8px 14px;border:1px solid #D1D5DB;">Máximo</th>
            <th style="padding:8px 14px;border:1px solid #D1D5DB;">Mínimo</th>
            <th style="padding:8px 14px;border:1px solid #D1D5DB;">Cierre</th>
          </tr>
        </thead>
        <tbody>
          <tr style="background-color:#F9FAFB;text-align:center;font-weight:bold;">
            <td style="padding:8px 14px;border:1px solid #D1D5DB;">${apertura:,.2f}</td>
            <td style="padding:8px 14px;border:1px solid #D1D5DB;">${maximo:,.2f}</td>
            <td style="padding:8px 14px;border:1px solid #D1D5DB;">${minimo:,.2f}</td>
            <td style="padding:8px 14px;border:1px solid #D1D5DB;">${cierre:,.2f}</td>
          </tr>
        </tbody>
      </table>

      {bloque_onedrive}

      <p style="margin-top:20px;color:#6B7280;font-size:12px;">
        Este correo fue generado automáticamente.
      </p>
    </div>
    """

    destinatarios = [
        {"emailAddress": {"address": correo.strip()}}
        for correo in mail_to_str.split(",")
        if correo.strip()
    ]

    mensaje = {
        "message": {
            "subject": f"Reporte Diario Dólar SET-FX (USD/COP) - {fecha_str}",
            "body": {"contentType": "HTML", "content": cuerpo_html},
            "toRecipients": destinatarios,
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": os.path.basename(ruta_xlsx_dia),
                    "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "contentBytes": adjunto_b64,
                }
            ],
        },
        "saveToSentItems": True,
    }

    url = f"https://graph.microsoft.com/v1.0/users/{quote(mail_from)}/sendMail"
    r = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=mensaje,
        timeout=60,
    )
    r.raise_for_status()


def main():
    carpeta_base = os.path.dirname(os.path.abspath(__file__))
    load_dotenv(os.path.join(carpeta_base, ".env"))

    df, ahora = obtener_datos_dolar()

    anio_str = str(ahora.year)
    mes_nombre = MESES_ES[ahora.month]
    fecha_str = ahora.strftime("%Y-%m-%d")
    nombre_excel_mes = f"{mes_nombre} {anio_str}.xlsx"

    carpeta_anio = os.path.join(carpeta_base, "Reportes", anio_str)
    carpeta_dia = os.path.join(carpeta_anio, mes_nombre, fecha_str)
    os.makedirs(carpeta_dia, exist_ok=True)

    # 1. CSV de datos crudos del día
    ruta_csv = os.path.join(carpeta_dia, f"Dolar-{fecha_str}.csv")
    df[["Hora", "Dolar"]].to_csv(ruta_csv, index=False, encoding="utf-8-sig")

    # 2. Captura idéntica del dashboard del día
    ruta_png = os.path.join(carpeta_dia, f"Grafica-Dolar-{fecha_str}.png")
    capturar_dashboard_set_icap(ruta_png)

    # 3. Excel diario (sin Datos_Crudos)
    ruta_xlsx_dia = os.path.join(carpeta_dia, f"Reporte-Dolar-{fecha_str}.xlsx")
    generar_excel_diario(df, fecha_str, ruta_png, ruta_xlsx_dia)

    # 4. Excel mensual (sincronizado con OneDrive)
    ruta_xlsx_mes = os.path.join(carpeta_anio, nombre_excel_mes)
    user_email = os.getenv("ONEDRIVE_USER_EMAIL")
    mail_from = os.getenv("MAIL_FROM")
    mail_to = os.getenv("MAIL_TO")

    token = None
    web_url = ""
    ruta_remota = construir_ruta_onedrive(anio_str, nombre_excel_mes)

    if user_email or (mail_from and mail_to):
        token = obtener_token_graph()

    if token and user_email:
        if not os.path.exists(ruta_xlsx_mes):
            descargar_mensual_si_existe(token, user_email, ruta_remota, ruta_xlsx_mes)

    actualizar_excel_mensual(df, fecha_str, ruta_png, ruta_xlsx_mes)

    if token and user_email:
        subir_a_onedrive(token, user_email, ruta_remota, ruta_xlsx_mes)
        web_url = obtener_url_carpeta_onedrive(token, user_email)
        print(f"Subido/actualizado en OneDrive: {ruta_remota}")

    # 5. Envío de correo diario con el Excel adjunto + enlace a carpeta de OneDrive
    if token and mail_from and mail_to:
        enviar_correo_graph(
            token, mail_from, mail_to, df, fecha_str, mes_nombre, anio_str, ruta_xlsx_dia, web_url
        )
        print(f"Correo enviado desde {mail_from} a {mail_to}")

    print(f"Reporte diario guardado en:  {carpeta_dia}")
    print(f"Reporte mensual actualizado: {ruta_xlsx_mes}")


if __name__ == "__main__":
    main()