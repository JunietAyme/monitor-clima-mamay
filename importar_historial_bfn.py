import os
import re
import csv
import glob
from datetime import datetime
import openpyxl

CARPETA_BASE = r"C:\Scripts\Monitor_Mamay"
CARPETA_IMPORT = os.path.join(CARPETA_BASE, "importar_excel")
CSV_VIVO = os.path.join(CARPETA_BASE, "historial_clima_mamay.csv")
CSV_HISTORICO = os.path.join(CARPETA_BASE, "historial_historico_excel.csv")

TEMP_PREV = 23.0
TEMP_CRIT = 26.0
HUM_PREV = 72.0
HUM_CRIT = 80.0


def parse_num(val):
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    txt = str(val).replace(",", ".").strip()
    if not txt or txt in ("—", "-", "N/D", "VENTA", "nan", "None"):
        return None
    m = re.search(r"(-?\d+(?:\.\d+)?)", txt)
    return float(m.group(1)) if m else None


def clasificar_estado(t, h):
    if (t is not None and t >= TEMP_CRIT) or (h is not None and h >= HUM_CRIT):
        return "CRITICO"
    if (t is not None and t > TEMP_PREV) or (h is not None and h > HUM_PREV):
        return "PREVENTIVO"
    return "NORMAL"


def fechas_ya_en_vivo():
    """Obtiene fechas (YYYY-MM-DD) ya capturadas en vivo para no duplicarlas."""
    fechas = set()
    if not os.path.isfile(CSV_VIVO):
        return fechas
    with open(CSV_VIVO, mode="r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader, None)
        for row in reader:
            if row and len(row[0]) >= 10:
                fechas.add(row[0][:10])
    return fechas


def extraer_bloques_hoja(ws):
    """Detecta todos los bloques horizontales/verticales con 'fecha' y números de galpón."""
    registros = {}  # (fecha_str, num_galpon) -> (val_min, val_max)
    max_r, max_c = ws.max_row, ws.max_column

    for r in range(1, max_r + 1):
        cols_fecha = []
        for c in range(1, max_c + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip().lower() == "fecha":
                cols_fecha.append(c)

        if not cols_fecha:
            continue

        # Por cada columna 'fecha' encontrada en la fila r, mapear galpones hacia la derecha
        for idx_b, col_f in enumerate(cols_fecha):
            col_limite = cols_fecha[idx_b + 1] - 1 if idx_b + 1 < len(cols_fecha) else max_c
            galpon_cols = []  # lista de (num_galpon, col_izq, col_der)

            for c in range(col_f + 1, col_limite + 1):
                val_g = ws.cell(r, c).value
                if isinstance(val_g, (int, float)) and 1 <= int(val_g) <= 17:
                    num_g = int(val_g)
                    # Verificar que debajo tenga etiquetas MIN / MAX
                    sub1 = str(ws.cell(r + 1, c).value or "").upper()
                    sub2 = str(ws.cell(r + 1, c + 1).value or "").upper()
                    if "MIN" in sub1 or "MAX" in sub1 or "MIN" in sub2 or "MAX" in sub2:
                        galpon_cols.append((num_g, c, c + 1))

            if not galpon_cols:
                continue

            # Leer filas de datos debajo del encabezado MIN/MAX hasta el siguiente bloque o fin
            curr_r = r + 2
            filas_vacias = 0
            while curr_r <= max_r:
                # Si empieza otro bloque de mes o encabezado, detener este bloque
                celda_check = ws.cell(curr_r, col_f).value
                if isinstance(celda_check, str) and celda_check.strip().lower() in ("fecha", "⁰c", "hr%"):
                    break

                if isinstance(celda_check, datetime):
                    filas_vacias = 0
                    fecha_str = celda_check.strftime("%Y-%m-%d")
                    for num_g, c1, c2 in galpon_cols:
                        v1 = parse_num(ws.cell(curr_r, c1).value)
                        v2 = parse_num(ws.cell(curr_r, c2).value)
                        vals = [v for v in (v1, v2) if v is not None]
                        if vals:
                            registros[(fecha_str, num_g)] = (min(vals), max(vals))
                else:
                    filas_vacias += 1
                    if filas_vacias >= 12:
                        break
                curr_r += 1

    return registros


def importar():
    archivos = glob.glob(os.path.join(CARPETA_IMPORT, "*.xlsx"))
    if not archivos:
        print(f"Coloca tu archivo Excel en: {CARPETA_IMPORT}")
        return

    fechas_vivo = fechas_ya_en_vivo()
    temp_map = {}
    hum_map = {}

    for ruta in sorted(archivos):
        print(f"Procesando: {os.path.basename(ruta)}...")
        wb = openpyxl.load_workbook(ruta, data_only=True)
        if "TEMPERATURA" in wb.sheetnames:
            temp_map.update(extraer_bloques_hoja(wb["TEMPERATURA"]))
        if "HUMEDAD" in wb.sheetnames:
            hum_map.update(extraer_bloques_hoja(wb["HUMEDAD"]))

    todas_claves = sorted(set(temp_map.keys()) | set(hum_map.keys()))
    filas_salida = []

    for fecha_str, num_g in todas_claves:
        if fecha_str in fechas_vivo:
            continue  # No duplicar días ya capturados por el monitor en vivo

        t_min, t_max = temp_map.get((fecha_str, num_g), (None, None))
        h_min, h_max = hum_map.get((fecha_str, num_g), (None, None))
        encargado = "Arias" if num_g <= 9 else "Castillo"
        galpon_lbl = f"Galpón {num_g:02d}"

        # Registro de mínima diaria (05:00) y máxima diaria (14:00)
        if t_min is not None or h_min is not None:
            filas_salida.append([
                fecha_str, "05:00:00", encargado, galpon_lbl,
                clasificar_estado(t_min, h_min),
                t_min, None, h_min, None
            ])
        if t_max is not None or h_max is not None:
            filas_salida.append([
                fecha_str, "14:00:00", encargado, galpon_lbl,
                clasificar_estado(t_max, h_max),
                t_max, None, h_max, None
            ])

    with open(CSV_HISTORICO, mode="w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "Fecha", "Hora", "Encargado", "Galpon", "Estado",
            "Temperatura_C", "Temperatura_Exterior_C",
            "Humedad_Pct", "Temperatura_Experimentada_C"
        ])
        writer.writerows(filas_salida)

    print(f"Importación completada: {len(filas_salida)} registros históricos guardados en historial_historico_excel.csv.")


if __name__ == "__main__":
    importar()
