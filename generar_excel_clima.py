import os
import re
import csv
import glob
import json
import shutil
import subprocess
from collections import defaultdict
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CARPETA_BASE = r"C:\Scripts\Monitor_Mamay"
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
            msg = "Actualizacion automatica " + datetime.now().strftime("%Y-%m-%d %H:%M")
            subprocess.run([git_cmd, "commit", "-m", msg], cwd=CARPETA_BASE, check=True, capture_output=True, timeout=15)
            subprocess.run([git_cmd, "push", "origin", "main"], cwd=CARPETA_BASE, check=True, capture_output=True, timeout=25)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Dashboard web sincronizado en GitHub Pages.")
    except Exception as e:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Aviso sincronizacion web (datos locales a salvo): {e}")


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
    limite_fino = (datetime.now() - timedelta(days=DIAS_DETALLE_FINO)).strftime("%Y-%m-%d")
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
    payload = {
        "actualizado": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
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
        ws["A2"] = f"Consolidado BigFarmNet Survey + Monitor en Vivo | Actualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
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
                    if col_idx == 7 and valor is not None and valor >= 80.0:
                        celda.fill = fill_alerta
                        celda.font = font_alerta
                else:
                    celda.number_format = "#,##0"
                    celda.alignment = Alignment(horizontal="center", vertical="center")

        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:H{4 + len(por_mes[ym])}"

        anchos = [14, 14, 14, 24, 24, 20, 20, 17]
        for i, ancho in enumerate(anchos, start=1):
            ws.column_dimensions[get_column_letter(i)].width = ancho

    try:
        wb.save(ruta_xlsx)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Excel y datos web actualizados ({len(serie_dashboard)} lecturas totales).")
        return True
    except PermissionError:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Datos web actualizados; cierre {ruta_xlsx} para refrescar el Excel.")
        return False


if __name__ == "__main__":
    actualizar_excel_mensual()

