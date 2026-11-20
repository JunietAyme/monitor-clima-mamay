import os
import re
import csv
import glob
import json
import shutil
import subprocess
from collections import defaultdict
from datetime import datetime, timedelta, timezone
TZ_LIMA = timezone(timedelta(hours=-5))
def ahora_lima():
    return datetime.now(TZ_LIMA).replace(tzinfo=None)
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart import BarChart, Reference

CARPETA_BASE = os.path.dirname(os.path.abspath(__file__))
ARCHIVO_XLSX = os.path.join(CARPETA_BASE, "Reporte_Clima_Mamay.xlsx")
ARCHIVO_JS_AUTO = os.path.join(CARPETA_BASE, "datos_clima.js")

GALPONES_ARIAS = range(1, 10)      # G01 - G09
GALPONES_CASTILLO = range(10, 18)  # G10 - G17
DIAS_DETALLE_FINO = 7              # Últimos 7 días a 5 min; anteriores a 30 min

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}


def obtener_encargado(num):
    return "Arias" if num in GALPONES_ARIAS else "Castillo"


def _git_exe():
    for c in [
        shutil.which("git"),
        r"C:\Program Files\Git\bin\git.exe",
        r"C:\Program Files\Git\cmd\git.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Git\cmd\git.exe"),
    ]:
        if c and os.path.isfile(c):
            return c
    return None


def sincronizar_github_pages():
    if not os.path.isdir(os.path.join(CARPETA_BASE, ".git")) or not os.path.isfile(ARCHIVO_JS_AUTO):
        return
    git_cmd = _git_exe()
    if not git_cmd:
        return
    try:
        subprocess.run(
            [git_cmd, "add", ".gitignore", "index.html", "datos_clima.js"],
            cwd=CARPETA_BASE, check=True, capture_output=True, timeout=15
        )
        res = subprocess.run([git_cmd, "diff", "--cached", "--quiet"], cwd=CARPETA_BASE, timeout=10)
        if res.returncode != 0:
            msg = "Actualizacion automatica " + ahora_lima().strftime("%Y-%m-%d %H:%M")
            subprocess.run([git_cmd, "commit", "-m", msg], cwd=CARPETA_BASE, check=True, capture_output=True, timeout=15)
            subprocess.run([git_cmd, "push", "origin", "main"], cwd=CARPETA_BASE, check=True, capture_output=True, timeout=25)
            print(f"[{ahora_lima().strftime('%H:%M:%S')}] Dashboard web sincronizado en GitHub Pages.")
    except Exception as e:
        print(f"[{ahora_lima().strftime('%H:%M:%S')}] Aviso sincronizacion web (datos locales a salvo): {e}")


def parse_float(val):
    if val is None:
        return None
    limpio = str(val).replace(",", ".").strip()
    if not limpio or limpio in ("None", "N/D", "nan", "—"):
        return None
    try:
        return float(limpio)
    except ValueError:
        return None


def extraer_num_galpon(texto):
    m = re.search(r"(\d+)", str(texto or ""))
    return int(m.group(1)) if m else 999


def normalizar_fecha_hora(raw_fecha, raw_hora=""):
    f = str(raw_fecha or "").strip()
    h = str(raw_hora or "").strip()

    m_survey = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}:\d{2}))?", f)
    if m_survey:
        dd = m_survey.group(1).zfill(2)
        mm = m_survey.group(2).zfill(2)
        yyyy = m_survey.group(3)
        hhmm = m_survey.group(4).zfill(5) if m_survey.group(4) else "00:00"
        return f"{yyyy}-{mm}-{dd}", hhmm

    if len(f) >= 10 and f[:4].isdigit() and f[4] == "-":
        return f[:10], (h[:5] if len(h) >= 5 else "00:00")

    return None, None


def nombre_hoja_mes(clave_ym):
    try:
        anio, mes = map(int, clave_ym.split("-"))
        return f"{MESES_ES[mes]} {anio}"
    except Exception:
        return clave_ym


def procesar_archivo_csv(ruta_csv, acumulado_diario, serie_dashboard):
    with open(ruta_csv, mode="r", encoding="utf-8-sig", newline="") as f:
        muestra = f.read(2048)
        f.seek(0)
        sep = ";" if ";" in muestra else ","
        reader = csv.reader(f, delimiter=sep)
        headers = next(reader, None)
        if not headers:
            return

        h_low = [h.strip().lower() for h in headers]

        idx_fecha = next((i for i, h in enumerate(h_low) if h in ("fecha", "date")), None)
        idx_hora = next((i for i, h in enumerate(h_low) if h == "hora"), None)
        idx_enc = next((i for i, h in enumerate(h_low) if h == "encargado"), None)
        idx_gal = next((i for i, h in enumerate(h_low) if h in ("galpon", "galpón", "ubicación", "ubicacion")), None)
        idx_sem = next((i for i, h in enumerate(h_low) if h == "semana"), None)

        idx_temp = next((i for i, h in enumerate(h_low) if
                         (h.startswith("temperatura") and "exterior" not in h and "experimentada" not in h and "mín" not in h and "máx" not in h)
                         or h == "temp_actual_c"), None)
        idx_ext = next((i for i, h in enumerate(h_low) if "exterior" in h and "mín" not in h and "máx" not in h), None)
        idx_hum = next((i for i, h in enumerate(h_low) if
                        (h.startswith("humedad") and "exterior" not in h and "mín" not in h and "máx" not in h)
                        or h == "hum_actual_pct"), None)
        idx_exp = next((i for i, h in enumerate(h_low) if "experimentada" in h), None)

        if idx_fecha is None or idx_gal is None or (idx_temp is None and idx_hum is None):
            return

        for cols in reader:
            if len(cols) <= max(idx_fecha, idx_gal):
                continue
            raw_h = cols[idx_hora] if idx_hora is not None and idx_hora < len(cols) else ""
            fecha, hora = normalizar_fecha_hora(cols[idx_fecha], raw_h)
            if not fecha:
                continue

            num = extraer_num_galpon(cols[idx_gal])
            if num == 999:
                continue

            encargado = (
                cols[idx_enc].strip()
                if idx_enc is not None and idx_enc < len(cols) and cols[idx_enc].strip()
                else obtener_encargado(num)
            )
            galpon_lbl = f"Galpón {num:02d}"

            t = parse_float(cols[idx_temp]) if idx_temp is not None and idx_temp < len(cols) else None
            ext = parse_float(cols[idx_ext]) if idx_ext is not None and idx_ext < len(cols) else None
            h = parse_float(cols[idx_hum]) if idx_hum is not None and idx_hum < len(cols) else None
            exp = parse_float(cols[idx_exp]) if idx_exp is not None and idx_exp < len(cols) else None
            sem_val = parse_float(cols[idx_sem]) if idx_sem is not None and idx_sem < len(cols) else None
            sem = int(sem_val) if sem_val is not None else None

            ym = fecha[:7]
            reg = acumulado_diario[(ym, fecha, num)]
            reg["encargado"] = encargado
            reg["galpon"] = galpon_lbl
            if t is not None:
                reg["temps"].append(t)
            if h is not None:
                reg["hums"].append(h)

            clave_punto = (fecha, hora, num)
            prev = serie_dashboard.get(clave_punto)
            if not prev or (exp is not None and prev.get("exp") is None):
                serie_dashboard[clave_punto] = {
                    "fecha": fecha,
                    "hora": hora,
                    "encargado": encargado,
                    "num": num,
                    "galpon": galpon_lbl,
                    "semana": sem if sem is not None else (prev.get("semana") if prev else None),
                    "temp": t,
                    "ext": ext,
                    "hum": h,
                    "exp": exp,
                }


def _compactar_serie_web(serie_dashboard):
    """Mantiene 5 min en los últimos 7 días y agrupa días anteriores a bloques de 30 min."""
    limite_fino = (ahora_lima() - timedelta(days=DIAS_DETALLE_FINO)).strftime("%Y-%m-%d")
    resultado = []
    buckets_antiguos = defaultdict(list)

    for k in sorted(serie_dashboard.keys(), key=lambda x: (x[0], x[1], x[2])):
        item = serie_dashboard[k]
        if item["fecha"] >= limite_fino:
            resultado.append(item)
        else:
            try:
                hh, mm = map(int, item["hora"].split(":"))
                bloque_mm = "00" if mm < 30 else "30"
                hora_bucket = f"{hh:02d}:{bloque_mm}"
            except Exception:
                hora_bucket = item["hora"]
            buckets_antiguos[(item["fecha"], hora_bucket, item["num"])].append(item)

    for (fecha, hora_bucket, num), grupo in sorted(buckets_antiguos.items()):
        base = dict(grupo[-1])
        base["hora"] = hora_bucket
        temps = [x["temp"] for x in grupo if x["temp"] is not None]
        exts = [x["ext"] for x in grupo if x["ext"] is not None]
        hums = [x["hum"] for x in grupo if x["hum"] is not None]
        exps = [x["exp"] for x in grupo if x["exp"] is not None]
        base["temp"] = round(max(temps), 1) if temps else None
        base["ext"] = round(sum(exts) / len(exts), 1) if exts else None
        base["hum"] = round(max(hums), 1) if hums else None
        base["exp"] = round(max(exps), 1) if exps else None
        resultado.append(base)

    resultado.sort(key=lambda x: (x["fecha"], x["hora"], x["num"]))
    return resultado


def exportar_datos_web_auto(serie_dashboard):
    lista_ordenada = _compactar_serie_web(serie_dashboard)
    if os.path.exists(ARCHIVO_JS_AUTO):
        try:
            raw_ant = open(ARCHIVO_JS_AUTO, 'r', encoding='utf-8').read().strip()
            if '=' in raw_ant:
                prev_total = len(json.loads(raw_ant.split('=', 1)[1].rstrip(';').strip()).get('registros', []))
                if prev_total > 0 and len(lista_ordenada) < int(prev_total * 0.85):
                    raise RuntimeError(f'Bloqueo de seguridad: los registros bajarian de {prev_total} a {len(lista_ordenada)}. Abortando escritura.')
        except RuntimeError:
            raise
        except Exception:
            pass
    payload = {
        "actualizado": ahora_lima().strftime("%d/%m/%Y %H:%M:%S"),
        "total": len(lista_ordenada),
        "registros": lista_ordenada,
    }
    contenido_js = "window.DATOS_MAMAY_AUTO = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n"
    tmp_js = ARCHIVO_JS_AUTO + ".tmp"
    with open(tmp_js, mode="w", encoding="utf-8") as f:
        f.write(contenido_js)
    os.replace(tmp_js, ARCHIVO_JS_AUTO)


def actualizar_excel_mensual(ruta_xlsx=ARCHIVO_XLSX):
    acumulado_diario = defaultdict(lambda: {
        "encargado": "",
        "galpon": "",
        "temps": [],
        "hums": []
    })
    serie_dashboard = {}
    if os.path.exists(ARCHIVO_JS_AUTO):
        try:
            raw_js = open(ARCHIVO_JS_AUTO, "r", encoding="utf-8").read().strip()
            if "=" in raw_js:
                json_str = raw_js.split("=", 1)[1].rstrip(";").strip()
                prev_data = json.loads(json_str)
                for reg_prev in prev_data.get("registros", []):
                    clave_p = (reg_prev.get("fecha"), reg_prev.get("hora"), reg_prev.get("num"))
                    if all(clave_p):
                        serie_dashboard[clave_p] = reg_prev
        except Exception as e:
            print(f"Aviso precargando datos_clima.js previo: {e}")

    archivos_csv = [
        p for p in glob.glob(os.path.join(CARPETA_BASE, "*.csv"))
        if not os.path.basename(p).endswith("_backup.csv")
    ]

    for ruta_csv in sorted(archivos_csv):
        procesar_archivo_csv(ruta_csv, acumulado_diario, serie_dashboard)

    if not acumulado_diario:
        print("No se encontraron registros válidos en los CSV.")
        return False

    exportar_datos_web_auto(serie_dashboard)
    sincronizar_github_pages()

    por_mes = defaultdict(list)
    for (ym, fecha, num), reg in sorted(acumulado_diario.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        t_min = min(reg["temps"]) if reg["temps"] else None
        t_max = max(reg["temps"]) if reg["temps"] else None
        h_min = min(reg["hums"]) if reg["hums"] else None
        h_max = max(reg["hums"]) if reg["hums"] else None
        muestras = max(len(reg["temps"]), len(reg["hums"]))

        por_mes[ym].append([
            fecha,
            reg["encargado"],
            reg["galpon"],
            t_min,
            t_max,
            h_min,
            h_max,
            muestras
        ])

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    fill_titulo = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    fill_header = PatternFill(start_color="EA580C", end_color="EA580C", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_alerta = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    font_titulo = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=9, italic=True, color="475569")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="0F172A")
    font_alerta = Font(name="Calibri", size=10, bold=True, color="B91C1C")

    borde_fino = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0"),
    )

    encabezados = [
        "Fecha",
        "Encargado",
        "Galpón",
        "Temperatura Mínima (°C)",
        "Temperatura Máxima (°C)",
        "Humedad Mínima (%)",
        "Humedad Máxima (%)",
        "Lecturas del Día"
    ]

    for ym in sorted(por_mes.keys()):
        titulo_hoja = nombre_hoja_mes(ym)
        ws = wb.create_sheet(title=titulo_hoja)

        ws.merge_cells("A1:H1")
        ws["A1"] = f"INVERSIONES MAMAY E.I.R.L. — REGISTRO DIARIO DE CLIMA ({titulo_hoja.upper()})"
        ws["A1"].font = font_titulo
        ws["A1"].fill = fill_titulo
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 26

        ws.merge_cells("A2:H2")
        ws["A2"] = f"Consolidado BigFarmNet Survey + Monitor en Vivo | Actualizado: {ahora_lima().strftime('%d/%m/%Y %H:%M')}"
        ws["A2"].font = font_sub
        ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[2].height = 18

        ws.row_dimensions[4].height = 24
        for col_idx, texto in enumerate(encabezados, start=1):
            celda = ws.cell(row=4, column=col_idx, value=texto)
            celda.font = font_header
            celda.fill = fill_header
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = borde_fino

        for idx_fila, fila in enumerate(por_mes[ym], start=5):
            ws.row_dimensions[idx_fila].height = 20
            es_par = (idx_fila % 2 == 0)

            for col_idx, valor in enumerate(fila, start=1):
                celda = ws.cell(row=idx_fila, column=col_idx, value=valor)
                celda.font = font_data
                celda.border = borde_fino
                if es_par:
                    celda.fill = fill_zebra

                if col_idx in (1, 2, 3):
                    celda.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in (4, 5):
                    celda.number_format = "0.0"
                    celda.alignment = Alignment(horizontal="right", vertical="center")
                    if col_idx == 5 and valor is not None and valor >= 26.0:
                        celda.fill = fill_alerta
                        celda.font = font_alerta
                elif col_idx in (6, 7):
                    celda.number_format = "0"
                    celda.alignment = Alignment(horizontal="right", vertical="center")
                    if col_idx == 7 and valor is not None and valor > 80.0:
                        celda.fill = fill_alerta
                        celda.font = font_alerta
                else:
                    celda.number_format = "#,##0"
                    celda.alignment = Alignment(horizontal="center", vertical="center")

        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:H{4 + len(por_mes[ym])}"

        # --- TABLAS RESUMEN LATERALES (J:M) PARA GRÁFICOS NATIVOS ---
        resumen_galp = defaultdict(lambda: {"t_prom": [], "t_max": [], "h_max": []})
        resumen_dia = defaultdict(lambda: {"t_min": [], "t_max": [], "h_max": []})

        for fila in por_mes[ym]:
            f_dia, _, g_nom, t_mn, t_mx, _, h_mx, _ = fila
            if t_mn is not None and t_mx is not None:
                resumen_galp[g_nom]["t_prom"].append((t_mn + t_mx) / 2.0)
            if t_mx is not None:
                resumen_galp[g_nom]["t_max"].append(t_mx)
                resumen_dia[f_dia]["t_max"].append(t_mx)
            if t_mn is not None:
                resumen_dia[f_dia]["t_min"].append(t_mn)
            if h_mx is not None:
                resumen_galp[g_nom]["h_max"].append(h_mx)
                resumen_dia[f_dia]["h_max"].append(h_mx)

        # Tabla 1 (J4:M...) - Resumen por Galpón
        headers_g = ["Galpón", "Temp. Prom. (°C)", "Temp. Máx. (°C)", "Hum. Máx. (%)"]
        for c_off, h_txt in enumerate(headers_g, start=10):
            c = ws.cell(row=4, column=c_off, value=h_txt)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde_fino

        galpones_ord = sorted(resumen_galp.keys(), key=extraer_num_galpon)
        for idx_g, g_nom in enumerate(galpones_ord, start=5):
            d_g = resumen_galp[g_nom]
            tp = round(sum(d_g["t_prom"]) / len(d_g["t_prom"]), 1) if d_g["t_prom"] else None
            tx = round(max(d_g["t_max"]), 1) if d_g["t_max"] else None
            hx = round(max(d_g["h_max"]), 0) if d_g["h_max"] else None
            for c_off, val in enumerate([g_nom.replace("Galpón ", "G"), tp, tx, hx], start=10):
                c = ws.cell(row=idx_g, column=c_off, value=val)
                c.font = font_data
                c.border = borde_fino
                c.alignment = Alignment(horizontal="center" if c_off == 10 else "right", vertical="center")
                if c_off in (11, 12):
                    c.number_format = "0.0"
                elif c_off == 13:
                    c.number_format = "0"

        fin_g = 4 + len(galpones_ord)
        if fin_g >= 5:
            chart_bar = BarChart()
            chart_bar.type = "col"
            chart_bar.style = 10
            chart_bar.grouping = "clustered"
            chart_bar.title = f"Comparativa Térmica por Galpón — {titulo_hoja}"
            chart_bar.y_axis.title = "Temperatura (°C)"
            chart_bar.x_axis.title = "Galpón"
            chart_bar.y_axis.scaling.min = 16
            chart_bar.width = 18
            chart_bar.height = 10
            data_bar = Reference(ws, min_col=11, max_col=12, min_row=4, max_row=fin_g)
            cats_bar = Reference(ws, min_col=10, min_row=5, max_row=fin_g)
            chart_bar.add_data(data_bar, titles_from_data=True)
            chart_bar.set_categories(cats_bar)
            ws.add_chart(chart_bar, "O4")

        # Tabla 2 (J24:M...) - Evolución Diaria del Mes
        fila_ini_d = max(24, fin_g + 3)
        ws.row_dimensions[fila_ini_d].height = 24
        headers_d = ["Fecha", "Temp. Mín. (°C)", "Temp. Máx. (°C)", "Hum. Máx. (%)"]
        for c_off, h_txt in enumerate(headers_d, start=10):
            c = ws.cell(row=fila_ini_d, column=c_off, value=h_txt)
            c.font = font_header
            c.fill = fill_header
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde_fino

        dias_ord = sorted(resumen_dia.keys())
        for idx_d, f_dia in enumerate(dias_ord, start=fila_ini_d + 1):
            d_d = resumen_dia[f_dia]
            tmn = round(min(d_d["t_min"]), 1) if d_d["t_min"] else None
            tmx = round(max(d_d["t_max"]), 1) if d_d["t_max"] else None
            hmx = round(max(d_d["h_max"]), 0) if d_d["h_max"] else None
            for c_off, val in enumerate([f_dia[5:], tmn, tmx, hmx], start=10):
                c = ws.cell(row=idx_d, column=c_off, value=val)
                c.font = font_data
                c.border = borde_fino
                c.alignment = Alignment(horizontal="center" if c_off == 10 else "right", vertical="center")
                if c_off in (11, 12):
                    c.number_format = "0.0"
                elif c_off == 13:
                    c.number_format = "0"

        fin_d = fila_ini_d + len(dias_ord)
        if fin_d >= fila_ini_d + 1:
            chart_line = LineChart()
            chart_line.style = 10
            chart_line.title = f"Evolución Térmica Diaria — {titulo_hoja}"
            chart_line.y_axis.title = "Temperatura (°C)"
            chart_line.x_axis.title = "Día (MM-DD)"
            chart_line.y_axis.scaling.min = 16
            chart_line.width = 18
            chart_line.height = 10
            data_line = Reference(ws, min_col=11, max_col=12, min_row=fila_ini_d, max_row=fin_d)
            cats_line = Reference(ws, min_col=10, min_row=fila_ini_d + 1, max_row=fin_d)
            chart_line.add_data(data_line, titles_from_data=True)
            chart_line.set_categories(cats_line)
            for serie in chart_line.series:
                serie.graphicalProperties.line.width = 25000
            ws.add_chart(chart_line, "O24")

        anchos = [14, 14, 14, 24, 24, 20, 20, 17, 3, 12, 18, 18, 16]
        for i, ancho in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(i)].width = ancho

        # --- Resumen por Galpón (Cols J:M) + Gráficos de Barras Nativos ---
        resumen_galp = {}
        for fila in por_mes[ym]:
            _, enc_f, galp_f, _, tmax_f, _, hmax_f, _ = fila
            num_g = extraer_num_galpon(galp_f)
            item_g = resumen_galp.setdefault(num_g, {
                "tag": f"G{num_g:02d}",
                "zona": enc_f,
                "tmax": None,
                "hmax": None,
            })
            if tmax_f is not None and (item_g["tmax"] is None or tmax_f > item_g["tmax"]):
                item_g["tmax"] = tmax_f
            if hmax_f is not None and (item_g["hmax"] is None or hmax_f > item_g["hmax"]):
                item_g["hmax"] = hmax_f

        if resumen_galp:
            headers_res = ["Galpón", "Zona", "Temp. Máx. (°C)", "Hum. Máx. (%)"]
            for offset, h_txt in enumerate(headers_res):
                c_hdr = ws.cell(row=4, column=10 + offset, value=h_txt)
                c_hdr.font = font_header
                c_hdr.fill = fill_titulo
                c_hdr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c_hdr.border = borde_fino

            galps_ord = [resumen_galp[k] for k in sorted(resumen_galp.keys())]
            for idx_g, info_g in enumerate(galps_ord, start=5):
                vals_g = [info_g["tag"], info_g["zona"], info_g["tmax"], info_g["hmax"]]
                for offset, val_g in enumerate(vals_g):
                    cel = ws.cell(row=idx_g, column=10 + offset, value=val_g)
                    cel.font = font_data
                    cel.border = borde_fino
                    if idx_g % 2 == 0:
                        cel.fill = fill_zebra
                    if offset in (0, 1):
                        cel.alignment = Alignment(horizontal="center", vertical="center")
                    elif offset == 2:
                        cel.number_format = "0.0"
                        cel.alignment = Alignment(horizontal="right", vertical="center")
                        if val_g is not None and val_g >= 26.0:
                            cel.fill = fill_alerta
                            cel.font = font_alerta
                    elif offset == 3:
                        cel.number_format = "0"
                        cel.alignment = Alignment(horizontal="right", vertical="center")
                        if val_g is not None and val_g > 80.0:
                            cel.fill = fill_alerta
                            cel.font = font_alerta

            anchos_res = [11, 13, 18, 16]
            for offset, w_res in enumerate(anchos_res):
                ws.column_dimensions[get_column_letter(10 + offset)].width = w_res

            fila_fin_res = 4 + len(galps_ord)
            cats_ref = Reference(ws, min_col=10, min_row=5, max_row=fila_fin_res)

            # Gráfico 1: Temperatura Máxima por Galpón (°C)
            chart_t = BarChart()
            chart_t.type = "col"
            chart_t.style = 10
            chart_t.title = f"Temperatura Máxima por Galpón — {titulo_hoja}"
            chart_t.y_axis.title = "°C"
            chart_t.x_axis.title = "Galpón"
            chart_t.y_axis.scaling.min = 18
            chart_t.legend = None
            chart_t.width = 16
            chart_t.height = 8.5
            data_t = Reference(ws, min_col=12, min_row=4, max_row=fila_fin_res)
            chart_t.add_data(data_t, titles_from_data=True)
            chart_t.set_categories(cats_ref)
            if chart_t.series:
                chart_t.series[0].graphicalProperties.solidFill = "EA580C"
            ws.add_chart(chart_t, f"J{fila_fin_res + 2}")

            # Gráfico 2: Humedad Máxima por Galpón (%)
            chart_h = BarChart()
            chart_h.type = "col"
            chart_h.style = 10
            chart_h.title = f"Humedad Máxima por Galpón — {titulo_hoja}"
            chart_h.y_axis.title = "% HR"
            chart_h.x_axis.title = "Galpón"
            chart_h.y_axis.scaling.min = 50
            chart_h.legend = None
            chart_h.width = 16
            chart_h.height = 8.5
            data_h = Reference(ws, min_col=13, min_row=4, max_row=fila_fin_res)
            chart_h.add_data(data_h, titles_from_data=True)
            chart_h.set_categories(cats_ref)
            if chart_h.series:
                chart_h.series[0].graphicalProperties.solidFill = "0A84FF"
            ws.add_chart(chart_h, f"J{fila_fin_res + 19}")

    try:
        wb.save(ruta_xlsx)
        print(f"[{ahora_lima().strftime('%H:%M:%S')}] Excel y datos web actualizados ({len(serie_dashboard)} lecturas totales).")
        return True
    except PermissionError:
        print(f"[{ahora_lima().strftime('%H:%M:%S')}] Datos web actualizados; cierre {ruta_xlsx} para refrescar el Excel.")
        return False


if __name__ == "__main__":
    actualizar_excel_mensual()




