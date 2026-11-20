import os
import re
import csv
import sys
import json
import time
import html
import shutil
import socket
import atexit
import logging
import threading
import subprocess
from logging.handlers import RotatingFileHandler
from datetime import datetime, date

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from generar_excel_clima import actualizar_excel_mensual

# ==============================================================================
# 1. CONFIGURACIÓN GENERAL Y UMBRALES
# ==============================================================================
CARPETA_BASE = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(CARPETA_BASE, "monitor_mamay.log")
LOCK_FILE = os.path.join(CARPETA_BASE, "monitor.lock")
ESTADO_FILE = os.path.join(CARPETA_BASE, "estado_alerta.json")
CACHE_LECTURA_FILE = os.path.join(CARPETA_BASE, "ultima_lectura.json")
ARCHIVO_CSV = os.path.join(CARPETA_BASE, "historial_clima_mamay.csv")
PERFIL_DIR = os.path.join(CARPETA_BASE, "perfil_bfn")

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] [%(levelname)s] %(message)s",
    handlers=[
        RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

CONFIG_SECRETA_FILE = os.path.join(CARPETA_BASE, "config_secreta.json")


def cargar_credenciales():
    if not os.path.isfile(CONFIG_SECRETA_FILE):
        if os.getenv("MAMAY_BFN_USER") and os.getenv("MAMAY_TG_TOKEN"):
            return {}
        raise FileNotFoundError(
            f"Falta {CONFIG_SECRETA_FILE}. Copia config_secreta.example.json a config_secreta.json y completa tus claves."
        )
    with open(CONFIG_SECRETA_FILE, "r", encoding="utf-8-sig") as f:
        return json.load(f)


_CREDS = cargar_credenciales()
TELEGRAM_TOKEN = os.getenv("MAMAY_TG_TOKEN", _CREDS.get("telegram_token", ""))
TELEGRAM_CHAT_ID = os.getenv("MAMAY_TG_CHAT_ID", _CREDS.get("telegram_chat_id", ""))

URL_SURVEY = "https://survey.bfn-fusion.bigdutchman.com/"
URL_DASHBOARD = "https://junietayme.github.io/monitor-clima-mamay/"
NOMBRE_PESTANA = "TEMPERATURA -HUMEDAD"

CORREO_WEB = os.getenv("MAMAY_BFN_USER", _CREDS.get("correo_bfn", ""))
CLAVE_WEB = os.getenv("MAMAY_BFN_PASS", _CREDS.get("clave_bfn", ""))

GALPONES_ARIAS = range(1, 10)      # G01 - G09
GALPONES_CASTILLO = range(10, 18)  # G10 - G17 (16-17 pendientes de red)

# Umbrales Climáticos (Normal -> Preventivo -> Crítico)
TEMP_PREV_LIMITE = 23.0
TEMP_CRIT_LIMITE = 26.0
HUM_PREV_LIMITE = 72.0
HUM_CRIT_LIMITE = 80.0
SENS_PREV_LIMITE = 25.5
SENS_CRIT_LIMITE = 28.0

# Escalamiento y recordatorios climáticos
DELTA_TEMP_ESCALAMIENTO = 0.5
DELTA_HUM_ESCALAMIENTO = 3.0
DELTA_SENS_ESCALAMIENTO = 0.7
RECORDATORIO_CRITICO_MIN = 15

# Umbrales Técnicos de Resiliencia (ciclos de 5 min)
CICLOS_DESCONEXION_ALERTA = 2      # 2 ciclos = 10 min sin señal
CICLOS_SENSOR_CONGELADO = 3       # 36 ciclos = 180 min sin variación mínima

# Resúmenes automáticos
HORAS_RESUMEN = ["08:00", "12:00", "15:00"]
VENTANA_RESUMEN_MIN = 45
ENVIAR_RESUMEN_AL_INICIAR = False
INTERVALO_SEGUNDOS = 3600

INICIO_SERVICIO_DT = datetime.now()
alerta_sesion_enviada = False
lock_estado = threading.Lock()


# ==============================================================================
# 2. CANDADO DE INSTANCIA ÚNICA Y PERSISTENCIA
# ==============================================================================
def pid_activo(pid):
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def adquirir_candado():
    if os.path.isfile(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r", encoding="utf-8") as f:
                pid_previo = int(f.read().strip())
            if pid_previo != os.getpid() and pid_activo(pid_previo):
                logging.warning(f"Ya existe una instancia activa (PID {pid_previo}). Cerrando duplicado.")
                sys.exit(0)
        except Exception:
            pass

    with open(LOCK_FILE, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()))
    atexit.register(liberar_candado)


def liberar_candado():
    try:
        if os.path.isfile(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


def cargar_estado(hoy_str):
    estado_vacio = {"fecha": hoy_str, "galpones": {}, "tecnico": {}, "resumenes": []}
    if not os.path.isfile(ESTADO_FILE):
        return estado_vacio
    try:
        with open(ESTADO_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if data.get("fecha") != hoy_str:
            return estado_vacio
        data.setdefault("galpones", {})
        data.setdefault("tecnico", {})
        data.setdefault("resumenes", [])
        return data
    except Exception:
        return estado_vacio


def guardar_estado(estado):
    try:
        tmp = ESTADO_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(estado, f, indent=2, ensure_ascii=False)
        os.replace(tmp, ESTADO_FILE)
    except Exception as e:
        logging.warning(f"No se pudo guardar estado_alerta.json: {e}")


def guardar_cache_lectura(ahora_dt, galpones):
    try:
        payload = {
            "timestamp": ahora_dt.isoformat(),
            "fecha_hora": ahora_dt.strftime("%d/%m/%Y %H:%M"),
            "galpones": {str(k): v for k, v in galpones.items()},
        }
        tmp = CACHE_LECTURA_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        os.replace(tmp, CACHE_LECTURA_FILE)
    except Exception as e:
        logging.warning(f"No se pudo actualizar ultima_lectura.json: {e}")


def leer_cache_lectura():
    if not os.path.isfile(CACHE_LECTURA_FILE):
        return None
    try:
        with open(CACHE_LECTURA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


# ==============================================================================
# 3. UTILIDADES CLIMÁTICAS Y CLASIFICACIÓN
# ==============================================================================
def hay_internet(host="8.8.8.8", port=53, timeout=3):
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def obtener_encargado(num):
    return "Arias" if num in GALPONES_ARIAS else "Castillo"


def parsear_primer_numero(texto):
    if not texto:
        return None
    limpio = texto.replace("\xa0", " ").strip()
    match = re.search(r"(-?\d+[.,]\d+|-?\d+)", limpio)
    return float(match.group(1).replace(",", ".")) if match else None


def extraer_numero_galpon(texto):
    match = re.search(r"Galp[oó]n\s*0*(\d+)", texto, re.IGNORECASE)
    if match:
        return int(match.group(1))
    match_num = re.search(r"(\d+)", texto)
    return int(match_num.group(1)) if match_num else 999


def clasificar_estado(temp_ref, hum_ref, exp_ref=None):
    es_critico = (
        (temp_ref is not None and temp_ref >= TEMP_CRIT_LIMITE)
        or (hum_ref is not None and hum_ref > HUM_CRIT_LIMITE)
        or (exp_ref is not None and exp_ref >= SENS_CRIT_LIMITE)
    )
    if es_critico:
        return "CRITICO", "🔴 CRÍTICO"

    es_preventivo = (
        (temp_ref is not None and temp_ref > TEMP_PREV_LIMITE)
        or (hum_ref is not None and hum_ref > HUM_PREV_LIMITE)
        or (exp_ref is not None and exp_ref > SENS_PREV_LIMITE)
    )
    if es_preventivo:
        return "PREVENTIVO", "🟡 PREVENTIVO"

    return "NORMAL", "🟢 ÓPTIMO"


def diagnosticar_motivos(g, previo):
    motivos = []
    t, h, exp, ext = g.get("temp"), g.get("hum"), g.get("exp"), g.get("ext")
    t_ant = previo.get("ultima_temp")
    h_ant = previo.get("ultima_hum")

    if t is not None and t >= TEMP_CRIT_LIMITE:
        delta = f" ({t - t_ant:+.1f} °C)" if t_ant is not None else ""
        motivos.append(f"Temp. interior crítica: <b>{t:.1f} °C</b>{delta}")
    elif t is not None and t > TEMP_PREV_LIMITE:
        delta = f" ({t - t_ant:+.1f} °C)" if t_ant is not None else ""
        motivos.append(f"Temp. interior elevada: <b>{t:.1f} °C</b>{delta}")

    if h is not None and h > HUM_CRIT_LIMITE:
        delta = f" ({h - h_ant:+.0f} %)" if h_ant is not None else ""
        motivos.append(f"Humedad crítica: <b>{h:.0f} %</b>{delta}")
    elif h is not None and h > HUM_PREV_LIMITE:
        delta = f" ({h - h_ant:+.0f} %)" if h_ant is not None else ""
        motivos.append(f"Humedad elevada: <b>{h:.0f} %</b>{delta}")

    if exp is not None and exp >= SENS_CRIT_LIMITE:
        motivos.append(f"Sensación térmica crítica: <b>{exp:.1f} °C</b>")
    elif exp is not None and exp > SENS_PREV_LIMITE:
        motivos.append(f"Sensación térmica alta: <b>{exp:.1f} °C</b>")

    if t is not None and ext is not None and (t - ext) >= 3.0:
        motivos.append(f"Diferencial vs. exterior: <b>+{t - ext:.1f} °C</b> (revisar ventilación)")

    return motivos


def toca_resumen_programado(ahora_dt, resumenes_enviados):
    min_actual = ahora_dt.hour * 60 + ahora_dt.minute
    for slot in HORAS_RESUMEN:
        if slot in resumenes_enviados:
            continue
        hh, mm = map(int, slot.split(":"))
        min_obj = hh * 60 + mm
        if 0 <= (min_actual - min_obj) <= VENTANA_RESUMEN_MIN:
            return slot
    return None


# ==============================================================================
# 4. TELEGRAM: ENVÍO + BOT INTERACTIVO BAJO DEMANDA + /SALUD
# ==============================================================================
TECLADO_PERMANENTE = {
    "keyboard": [
        [{"text": "📊 Estado General"}, {"text": "⚠️ Solo Alertas"}],
        [{"text": "🐔 Zona Arias (01-09)"}, {"text": "🐔 Zona Castillo (10-17)"}],
    ],
    "resize_keyboard": True,
    "is_persistent": True,
}


def enviar_telegram(mensaje_html, url_boton=None, texto_boton="Abrir Dashboard en Vivo", chat_destino=None):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    destino = str(chat_destino or TELEGRAM_CHAT_ID)
    link = url_boton or URL_DASHBOARD

    payload = {
        "chat_id": destino,
        "text": mensaje_html,
        "parse_mode": "HTML",
        "disable_web_page_preview": True,
        "reply_markup": TECLADO_PERMANENTE,
    }

    if link.startswith("http"):
        payload["reply_markup"] = {
            "inline_keyboard": [[{"text": texto_boton, "url": link}]]
        }

    try:
        r = requests.post(url, json=payload, timeout=15)
        logging.info(f"Telegram enviado (HTTP {r.status_code})")
        if r.status_code != 200:
            logging.warning(f"Detalle Telegram: {r.text}")
    except requests.exceptions.RequestException as e:
        logging.warning(f"Sin red al enviar Telegram: {e}")
    except Exception as e:
        logging.error(f"Error inesperado enviando Telegram: {e}")


def activar_teclado_consulta(chat_id):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": str(chat_id),
        "text": (
            "🎛️ <b>Panel de consulta bajo demanda activado</b>\n"
            "Usa los botones inferiores para consultar los galpones cuando lo necesites:"
        ),
        "parse_mode": "HTML",
        "reply_markup": TECLADO_PERMANENTE,
    }
    try:
        requests.post(url, json=payload, timeout=10)
    except Exception:
        pass


def responder_salud_sistema(chat_id):
    ahora = datetime.now()
    uptime_seg = int((ahora - INICIO_SERVICIO_DT).total_seconds())
    horas, rem = divmod(uptime_seg, 3600)
    mins, _ = divmod(rem, 60)

    cache = leer_cache_lectura()
    ult_str = cache.get("fecha_hora", "Sin lecturas aún") if cache else "Sin lecturas aún"
    cant_galp = len(cache.get("galpones", {})) if cache else 0

    filas_vivo = 0
    if os.path.isfile(ARCHIVO_CSV):
        try:
            with open(ARCHIVO_CSV, encoding="utf-8-sig") as f:
                filas_vivo = max(0, sum(1 for _ in f) - 1)
        except Exception:
            pass

    msg = (
        f"🩺 <b>DIAGNÓSTICO DE SALUD — MONITOR MAMAY</b>\n"
        f"• <b>Estado servicio:</b> 🟢 Activo (PID {os.getpid()})\n"
        f"• <b>Tiempo activo:</b> {horas}h {mins}m\n"
        f"• <b>Última lectura BFN:</b> {ult_str} ({cant_galp} galpones)\n"
        f"• <b>Registros en vivo (CSV):</b> {filas_vivo:,} filas\n"
        f"• <b>Umbrales vigentes:</b> Temp &gt;{TEMP_PREV_LIMITE}°C / ≥{TEMP_CRIT_LIMITE}°C | "
        f"Hum &gt;{HUM_PREV_LIMITE:.0f}% / &gt;{HUM_CRIT_LIMITE:.0f}%"
    )
    enviar_telegram(msg, chat_destino=chat_id)


def responder_log_sistema(chat_id):
    if str(chat_id) != str(TELEGRAM_CHAT_ID):
        return
    if not os.path.isfile(LOG_FILE):
        enviar_telegram("?? No se encontr? el archivo monitor_mamay.log.", chat_destino=chat_id)
        return
    try:
        with open(LOG_FILE, "r", encoding="utf-8", errors="replace") as f:
            ultimas = f.readlines()[-20:]
        texto_log = "".join(ultimas).strip()
        if len(texto_log) > 3400:
            texto_log = "...\n" + texto_log[-3400:]
        enviar_telegram(
            f"?? <b>?LTIMAS L?NEAS DEL LOG ? MAMAY</b>\n<pre>{html.escape(texto_log)}</pre>",
            chat_destino=chat_id,
        )
    except Exception as e:
        enviar_telegram(f"❌ Error leyendo log: {html.escape(str(e))}", chat_destino=chat_id)


def ejecutar_actualizacion_remota(chat_id):
    if str(chat_id) != str(TELEGRAM_CHAT_ID):
        return
    enviar_telegram("⏳ Descargando última versión desde GitHub...", chat_destino=chat_id)
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        subprocess.run(["git", "fetch", "origin", "main"], cwd=CARPETA_BASE, check=True, capture_output=True, text=True, creationflags=flags)
        subprocess.run(["git", "reset", "--hard", "origin/main"], cwd=CARPETA_BASE, check=True, capture_output=True, text=True, creationflags=flags)
        info_commit = subprocess.check_output(
            ["git", "log", "-n", "1", "--pretty=format:%h — %s (%ad)", "--date=format:%d/%m %H:%M"],
            cwd=CARPETA_BASE, text=True, creationflags=flags
        ).strip()
    except Exception as e:
        enviar_telegram(f"❌ Error al actualizar con Git:\n<pre>{html.escape(str(e))}</pre>", chat_destino=chat_id)
        return

    enviar_telegram(
        f"✅ <b>Código actualizado en producción</b>\n"
        f"? <b>Commit:</b> {html.escape(info_commit)}\n"
        f"?? Reiniciando servicio en segundo plano...",
        chat_destino=chat_id,
    )

    try:
        if os.path.isfile(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass

    pythonw_exe = sys.executable.replace("python.exe", "pythonw.exe")
    if not os.path.isfile(pythonw_exe):
        pythonw_exe = "pythonw.exe"

    detached = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | flags
    subprocess.Popen([pythonw_exe, os.path.join(CARPETA_BASE, "monitor_mamay.py")], cwd=CARPETA_BASE, creationflags=detached, close_fds=True)
    time.sleep(1.5)
    os._exit(0)


def responder_consulta_bajo_demanda(comando_texto, chat_id):
    cache = leer_cache_lectura()
    if not cache or not cache.get("galpones"):
        enviar_telegram("⏳ Aún no hay lecturas registradas en caché. Espera el próximo ciclo de 5 min.", chat_destino=chat_id)
        return

    galpones_dict = {int(k): v for k, v in cache["galpones"].items()}
    fecha_hora = cache.get("fecha_hora", "N/D")

    try:
        dt_lectura = datetime.fromisoformat(cache["timestamp"])
        mins_antig = max(0, int((datetime.now() - dt_lectura).total_seconds() // 60))
        etiqueta_antig = "ahora mismo" if mins_antig == 0 else f"hace {mins_antig} min"
    except Exception:
        etiqueta_antig = ""

    cmd = comando_texto.lower()
    todos = [galpones_dict[n] for n in sorted(galpones_dict)]

    if "arias" in cmd:
        lista = [g for g in todos if g["num"] in GALPONES_ARIAS]
        cuerpo = formatear_resumen_compacto("ZONA ARIAS (Galpones 01–09)", lista)
        enviar_telegram(
            f"<b>CONSULTA EN VIVO — ZONA ARIAS</b>\n"
            f"🕒 Lectura: <b>{fecha_hora}</b> ({etiqueta_antig})\n\n{cuerpo}",
            url_boton=f"{URL_DASHBOARD}?zona=Arias",
            texto_boton="Ver Zona Arias en Gráficas",
            chat_destino=chat_id,
        )
    elif "castillo" in cmd:
        lista = [g for g in todos if g["num"] in GALPONES_CASTILLO]
        cuerpo = formatear_resumen_compacto("ZONA CASTILLO (Galpones 10–17)", lista)
        enviar_telegram(
            f"<b>CONSULTA EN VIVO — ZONA CASTILLO</b>\n"
            f"🕒 Lectura: <b>{fecha_hora}</b> ({etiqueta_antig})\n"
            f"<i>Nota: G16 y G17 pendientes de enlace BFN.</i>\n\n{cuerpo}",
            url_boton=f"{URL_DASHBOARD}?zona=Castillo",
            texto_boton="Ver Zona Castillo en Gráficas",
            chat_destino=chat_id,
        )
    elif "alerta" in cmd:
        en_alerta = [g for g in todos if clasificar_estado(g["temp"], g["hum"], g.get("exp"))[0] != "NORMAL"]
        if not en_alerta:
            enviar_telegram(
                f"✅ <b>TODA LA GRANJA EN RANGO ÓPTIMO</b>\n"
                f"🕒 Lectura: <b>{fecha_hora}</b> ({etiqueta_antig})\n"
                f"Los {len(todos)} galpones conectados están ≤ 23.0 °C y ≤ 72 % HR.",
                chat_destino=chat_id,
            )
        else:
            cuerpo = formatear_resumen_compacto(f"GALPONES FUERA DE RANGO ({len(en_alerta)})", en_alerta)
            enviar_telegram(
                f"⚠️ <b>ALERTAS ACTIVAS EN ESTE MOMENTO</b>\n"
                f"🕒 Lectura: <b>{fecha_hora}</b> ({etiqueta_antig})\n\n{cuerpo}",
                chat_destino=chat_id,
            )
    else:
        sec_arias = formatear_resumen_compacto("ZONA ARIAS (01–09)", [g for g in todos if g["num"] in GALPONES_ARIAS])
        sec_castillo = formatear_resumen_compacto("ZONA CASTILLO (10–17)", [g for g in todos if g["num"] in GALPONES_CASTILLO])
        enviar_telegram(
            f"<b>ESTADO GENERAL — GRANJA MAMAY</b>\n"
            f"🕒 Lectura: <b>{fecha_hora}</b> ({etiqueta_antig})\n\n"
            f"{sec_arias}\n\n{sec_castillo}",
            chat_destino=chat_id,
        )


def escuchar_comandos_telegram():
    offset = 0
    url_updates = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates"
    logging.info("Hilo interactivo de botones Telegram iniciado.")

    while True:
        try:
            r = requests.get(url_updates, params={"offset": offset, "timeout": 25}, timeout=32)
            if r.status_code != 200:
                time.sleep(5)
                continue

            data = r.json()
            for upd in data.get("result", []):
                offset = upd["update_id"] + 1
                msg = upd.get("message") or upd.get("channel_post")
                if not msg:
                    continue
                fecha_msg = msg.get("date", 0)
                if fecha_msg and (time.time() - fecha_msg) > 120:
                    continue
                texto = (msg.get("text") or "").strip()
                chat_id = msg.get("chat", {}).get("id")
                if not texto or not chat_id:
                    continue

                if texto.startswith("/start") or texto.startswith("/botones") or texto.startswith("/menu"):
                    activar_teclado_consulta(chat_id)
                elif texto.startswith("/salud"):
                    responder_salud_sistema(chat_id)
                elif texto.startswith("/log"):
                    responder_log_sistema(chat_id)
                elif texto.startswith("/actualizar"):
                    ejecutar_actualizacion_remota(chat_id)
                elif any(k in texto.lower() for k in ["estado general", "zona arias", "zona castillo", "solo alertas", "/estado"]):
                    responder_consulta_bajo_demanda(texto, chat_id)
        except Exception:
            time.sleep(5)


# ==============================================================================
# 5. EXTRACCIÓN DESDE BIGFARMNET SURVEY
# ==============================================================================
def extraer_filas_visibles(page, galpones):
    datos = page.evaluate(
        """() => {
            const headers = Array.from(document.querySelectorAll('table thead th')).map(th => th.innerText.trim().toLowerCase());
            const filas = Array.from(document.querySelectorAll('table tbody tr')).map(tr =>
                Array.from(tr.querySelectorAll('td')).map(td => td.innerText.trim())
            );
            return { headers, filas };
        }"""
    )
    headers = datos.get("headers", [])
    filas_raw = datos.get("filas", [])

    idx_temp = next((i for i, h in enumerate(headers) if "temperatura" in h and "exterior" not in h and "experimentada" not in h), None)
    idx_ext = next((i for i, h in enumerate(headers) if "exterior" in h), None)
    idx_hum = next((i for i, h in enumerate(headers) if "humedad" in h), None)
    idx_exp = next((i for i, h in enumerate(headers) if "experimentada" in h), None)

    for celdas in filas_raw:
        if len(celdas) < 4 or "Galpon" not in celdas[0]:
            continue
        num = extraer_numero_galpon(celdas[0])
        if num == 999:
            continue

        def val_idx(idx):
            return parsear_primer_numero(celdas[idx]) if idx is not None and idx < len(celdas) else None

        galpones[num] = {
            "num": num,
            "activo": "Activo" in celdas[0],
            "temp": val_idx(idx_temp),
            "ext": val_idx(idx_ext),
            "hum": val_idx(idx_hum),
            "exp": val_idx(idx_exp),
        }


def ampliar_vista_todos_galpones(page, galpones):
    try:
        selector_10 = page.locator("mat-select, [role='combobox'], select, button").filter(
            has_text=re.compile(r"^\s*10\s*$")
        ).last
        if selector_10.count() > 0:
            if selector_10.evaluate("el => el.tagName") == "SELECT":
                selector_10.select_option(label="40")
            else:
                selector_10.click(timeout=2500)
                page.wait_for_timeout(800)
                opt_40 = page.locator("mat-option, [role='option'], li, button").filter(
                    has_text=re.compile(r"^\s*40\s*$")
                ).first
                if opt_40.count() > 0:
                    opt_40.click(timeout=2500)
            page.wait_for_timeout(2500)
    except Exception:
        pass

    extraer_filas_visibles(page, galpones)

    if len(galpones) < 15:
        try:
            btn_sig = page.locator(
                "button[aria-label*='siguiente' i], button[aria-label*='next' i], "
                "button.mat-mdc-paginator-navigation-next, button.mat-paginator-navigation-next"
            ).first
            if btn_sig.count() > 0 and btn_sig.is_enabled():
                btn_sig.click(timeout=3000)
                page.wait_for_timeout(3000)
                extraer_filas_visibles(page, galpones)
        except Exception:
            pass


def leer_tabla_survey():
    global alerta_sesion_enviada
    galpones = {}

    try:
        with sync_playwright() as p:
            import shutil
            canal_navegador = "chrome" if shutil.which("google-chrome") else None
            context = p.chromium.launch_persistent_context(
                user_data_dir=PERFIL_DIR,
                channel=canal_navegador,
                headless=True,
                viewport={"width": 1920, "height": 1400},
                args=["--disable-blink-features=AutomationControlled"],
            )
            try:
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(URL_SURVEY, wait_until="domcontentloaded", timeout=60000)

                for _ in range(30):
                    page.wait_for_timeout(1000)
                    if "login.bfn-fusion" in page.url and page.locator('input[name="password"], #password').count() > 0:
                        if CLAVE_WEB:
                            page.locator('input[name="username"], #username').first.fill(CORREO_WEB)
                            page.locator('input[name="password"], #password').first.fill(CLAVE_WEB)
                            chk = page.locator('#rememberMe, input[name="rememberMe"]').first
                            if chk.count() > 0 and not chk.is_checked():
                                chk.check()
                            page.locator('#kc-login, button[type="submit"], input[type="submit"]').first.click()
                            page.wait_for_url("**/survey.bfn-fusion.bigdutchman.com/**", timeout=45000)
                        else:
                            if not alerta_sesion_enviada:
                                enviar_telegram(
                                    "<b>AVISO DE SISTEMA: SESIÓN BFN CADUCADA</b>\n"
                                    "Ejecute <code>guardar_sesion.py</code> en el equipo para restablecer el acceso."
                                )
                                alerta_sesion_enviada = True
                            raise RuntimeError("Sesión de BigFarmNet caducada (redirigido a login).")

                    if "survey.bfn-fusion.bigdutchman.com" in page.url and page.locator("table").count() > 0:
                        break

                alerta_sesion_enviada = False

                tab = page.get_by_text(NOMBRE_PESTANA, exact=False).first
                if tab.count() > 0:
                    tab.click()

                page.wait_for_function(
                    """() => {
                        const filas = Array.from(document.querySelectorAll('table tbody tr'));
                        return filas.some(tr => tr.innerText.includes('Galpon') && tr.innerText.includes('°C'));
                    }""",
                    timeout=60000,
                )
                try:
                    page.wait_for_function(
                        """() => {
                            const filas = Array.from(document.querySelectorAll('table tbody tr'));
                            return filas.some(tr => {
                                const txt = tr.innerText;
                                return txt.includes('Galpon') && (txt.match(/°C/g) || []).length >= 3;
                            });
                        }""",
                        timeout=10000,
                    )
                except Exception:
                    logging.warning("Las columnas Exterior/Experimentada tardaron más de 10s o llegaron vacías desde BFN.")
                page.wait_for_timeout(1000)

                ampliar_vista_todos_galpones(page, galpones)
                if galpones and all(g["ext"] is None and g["exp"] is None for g in galpones.values()):
                    logging.warning("Datos de Exterior/Experimentada vacíos en primer intento; reintentando en 5s...")
                    page.wait_for_timeout(5000)
                    ampliar_vista_todos_galpones(page, galpones)
                if galpones and all(g["ext"] is None and g["exp"] is None for g in galpones.values()):
                    logging.warning("Ciclo descartado: BFN no cargó Exterior ni Experimentada tras reintento.")
                    galpones = {}
            finally:
                try:
                    context.close()
                except Exception:
                    pass

    except PlaywrightTimeoutError:
        logging.warning("Timeout de red al consultar BigFarmNet Fusion. Se reintentará en el próximo ciclo.")
        return {}
    except Exception as e:
        logging.error(f"Error de conexión o lectura en Survey: {e}")
        return {}

    logging.info(f"Galpones extraídos ({len(galpones)}): {sorted(galpones.keys())}")
    return galpones


# ==============================================================================
# 6. PERSISTENCIA Y FORMATEO DE MENSAJES
# ==============================================================================
def guardar_historial_csv(ahora_dt, galpones):
    existe = os.path.isfile(ARCHIVO_CSV)
    fecha_str = ahora_dt.strftime("%Y-%m-%d")
    hora_min = ahora_dt.strftime("%H:%M")
    claves_recientes = set()
    if existe:
        try:
            with open(ARCHIVO_CSV, mode="r", encoding="utf-8-sig") as fr:
                for linea in fr.readlines()[-40:]:
                    partes = linea.strip().split(";")
                    if len(partes) >= 4:
                        claves_recientes.add((partes[0], partes[1][:5], partes[3]))
        except Exception:
            pass

    with open(ARCHIVO_CSV, mode="a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        if not existe:
            writer.writerow([
                "Fecha", "Hora", "Encargado", "Galpon", "Estado",
                "Temperatura_C", "Temperatura_Exterior_C",
                "Humedad_Pct", "Temperatura_Experimentada_C",
            ])

        for num in sorted(galpones):
            galp_lbl = f"Galpón {num:02d}"
            if (fecha_str, hora_min, galp_lbl) in claves_recientes:
                continue
            g = galpones[num]
            estado, _ = clasificar_estado(g["temp"], g["hum"], g.get("exp"))
            writer.writerow([
                fecha_str,
                ahora_dt.strftime("%H:%M:%S"),
                obtener_encargado(num),
                galp_lbl,
                estado,
                g["temp"],
                g["ext"],
                g["hum"],
                g["exp"],
            ])


def hacer_backup_diario():
    carpeta_backup = os.path.join(CARPETA_BASE, "backups")
    os.makedirs(carpeta_backup, exist_ok=True)
    hoy_str = datetime.now().strftime("%Y-%m-%d")

    pares = [
        (ARCHIVO_CSV, os.path.join(carpeta_backup, f"historial_clima_{hoy_str}.csv"), "CSV"),
        (os.path.join(CARPETA_BASE, "Reporte_Clima_Mamay.xlsx"), os.path.join(carpeta_backup, f"Reporte_Clima_{hoy_str}.xlsx"), "Excel"),
    ]
    for orig, dest, nombre in pares:
        try:
            if os.path.isfile(orig) and not os.path.isfile(dest):
                shutil.copy2(orig, dest)
                logging.info(f"Backup diario de {nombre} creado correctamente.")
        except Exception as e:
            logging.error(f"Error al realizar backup diario de {nombre}: {e}")


def fmt_val(val, sufijo="", dec=1):
    return f"{val:.{dec}f}{sufijo}" if val is not None else "N/D"


def formatear_resumen_compacto(titulo, lista_galpones):
    if not lista_galpones:
        return f"<b>{html.escape(titulo)}</b>\nSin galpones activos."

    lineas = [f"<b>{html.escape(titulo)}</b>"]
    for g in lista_galpones:
        nivel, _ = clasificar_estado(g["temp"], g["hum"], g.get("exp"))
        icono = "🔴" if nivel == "CRITICO" else ("🟡" if nivel == "PREVENTIVO" else "🟢")
        lineas.append(
            f"{icono} <b>G{g['num']:02d}:</b> Int <b>{fmt_val(g['temp'], '°')}</b> · "
            f"Sens <b>{fmt_val(g['exp'], '°')}</b> · Hum <b>{fmt_val(g['hum'], '%', 0)}</b> · "
            f"Ext {fmt_val(g['ext'], '°')}"
        )
    return "\n".join(lineas)


def construir_reporte_completo(galpones, ahora_dt):
    todos = [galpones[n] for n in sorted(galpones)]
    cabecera = (
        f"<b>REPORTE PROGRAMADO — GRANJA MAMAY</b>\n"
        f"📅 {ahora_dt.strftime('%d/%m/%Y')} | 🕒 {ahora_dt.strftime('%H:%M')} hrs"
    )
    sec_arias = formatear_resumen_compacto("ZONA ARIAS (Galpones 01 al 09)", [g for g in todos if g["num"] in GALPONES_ARIAS])
    sec_castillo = formatear_resumen_compacto("ZONA CASTILLO (Galpones 10 al 17)", [g for g in todos if g["num"] in GALPONES_CASTILLO])
    return f"{cabecera}\n\n{sec_arias}\n\n{sec_castillo}"


# ==============================================================================
# 7. AUDITORÍA TÉCNICA DE SENSORES Y ENLACE BFN
# ==============================================================================
def evaluar_salud_sensores(estado, galpones):
    """Detecta galpones desconectados o sensores congelados sin alterar las alertas climáticas."""
    tecnico = estado.setdefault("tecnico", {})
    conocidos = set(int(k) for k in tecnico.keys()) | set(galpones.keys())
    avisos_tecnicos = []

    for num in sorted(conocidos):
        clave = str(num)
        reg_t = tecnico.get(clave, {
            "ausencias": 0,
            "alerta_desconexion": False,
            "firma_previa": None,
            "ciclos_iguales": 0,
            "alerta_congelado": False,
        })
        zona = obtener_encargado(num)

        if num not in galpones:
            reg_t["ausencias"] = reg_t.get("ausencias", 0) + 1
            if reg_t["ausencias"] >= CICLOS_DESCONEXION_ALERTA and not reg_t.get("alerta_desconexion"):
                avisos_tecnicos.append(
                    f"🔌 <b>Pérdida de enlace BFN:</b> Galpón {num:02d} ({zona}) dejó de reportar datos hace {reg_t['ausencias'] * 5} min."
                )
                reg_t["alerta_desconexion"] = True
        else:
            if reg_t.get("alerta_desconexion"):
                avisos_tecnicos.append(
                    f"✅ <b>Enlace BFN restablecido:</b> Galpón {num:02d} ({zona}) volvió a enviar telemetría."
                )
            reg_t["ausencias"] = 0
            reg_t["alerta_desconexion"] = False

            g = galpones[num]
            firma_act = f"{g.get('temp')}|{g.get('hum')}|{g.get('ext')}|{g.get('exp')}"
            if g.get("temp") is not None and firma_act == reg_t.get("firma_previa"):
                reg_t["ciclos_iguales"] = reg_t.get("ciclos_iguales", 0) + 1
                if reg_t["ciclos_iguales"] >= CICLOS_SENSOR_CONGELADO and not reg_t.get("alerta_congelado"):
                    avisos_tecnicos.append(
                        f"🧊 <b>Posible sensor congelado:</b> Galpón {num:02d} ({zona}) repite exactamente "
                        f"{fmt_val(g.get('temp'), ' °C')} / {fmt_val(g.get('hum'), '%', 0)} desde hace {reg_t['ciclos_iguales'] * 5} min."
                    )
                    reg_t["alerta_congelado"] = True
            else:
                reg_t["firma_previa"] = firma_act
                reg_t["ciclos_iguales"] = 0
                reg_t["alerta_congelado"] = False

        tecnico[clave] = reg_t

    return avisos_tecnicos


# ==============================================================================
# 8. ORQUESTACIÓN PRINCIPAL DE CICLO
# ==============================================================================
primer_ciclo = True


def procesar_ciclo(forzar_reporte=False):
    global primer_ciclo
    ahora_dt = datetime.now()
    hoy_str = date.today().isoformat()
    hora_actual = ahora_dt.strftime("%H:%M")

    if not hay_internet():
        logging.warning("Sin conexión a internet. Omitiendo extracción y esperando próximo ciclo.")
        return

    galpones = leer_tabla_survey()
    if not galpones:
        logging.warning("No se obtuvieron datos válidos de Survey en este ciclo.")
        return

    guardar_cache_lectura(ahora_dt, galpones)

    try:
        guardar_historial_csv(ahora_dt, galpones)
    except Exception as e:
        logging.error(f"Error guardando historial CSV: {e}")

    try:
        actualizar_excel_mensual()
    except Exception as e:
        logging.error(f"Error actualizando Excel / datos_clima.js: {e}")

    hacer_backup_diario()

    with lock_estado:
        estado = cargar_estado(hoy_str)
        avisos_tecnicos = evaluar_salud_sensores(estado, galpones)
        alertas_criticas = []
        alertas_preventivas = []
        normalizados = []

        for num in sorted(galpones):
            g = galpones[num]
            t_act, h_act, exp_act = g["temp"], g["hum"], g.get("exp")
            clave_g = str(num)
            nivel_act, _ = clasificar_estado(t_act, h_act, exp_act)

            base_previo = {
                "nivel": "NORMAL",
                "max_temp": 0.0,
                "max_hum": 0.0,
                "max_exp": 0.0,
                "ultima_temp": None,
                "ultima_hum": None,
                "inicio_alerta": None,
                "ultimo_aviso_ts": None,
            }
            guardado = estado["galpones"].get(clave_g)
            previo = {**base_previo, **guardado} if isinstance(guardado, dict) else base_previo

            en_histeresis = (
                previo.get("nivel") in ("PREVENTIVO", "CRITICO")
                and (
                    (t_act is not None and t_act > (TEMP_PREV_LIMITE - 0.4))
                    or (h_act is not None and h_act > (HUM_PREV_LIMITE - 2.0))
                    or (exp_act is not None and exp_act > (SENS_PREV_LIMITE - 0.4))
                )
            )

            if nivel_act == "NORMAL" and not en_histeresis:
                if previo.get("nivel") in ("PREVENTIVO", "CRITICO"):
                    duracion_txt = ""
                    if previo.get("inicio_alerta"):
                        try:
                            mins = max(1, int((ahora_dt - datetime.fromisoformat(previo["inicio_alerta"])).total_seconds() // 60))
                            duracion_txt = f" (duró {mins} min)"
                        except Exception:
                            pass
                    normalizados.append((g, duracion_txt))

                estado["galpones"][clave_g] = {
                    "nivel": "NORMAL",
                    "max_temp": 0.0,
                    "max_hum": 0.0,
                    "max_exp": 0.0,
                    "ultima_temp": t_act,
                    "ultima_hum": h_act,
                    "inicio_alerta": None,
                    "ultimo_aviso_ts": None,
                }
                continue

            nivel_prev = previo.get("nivel", "NORMAL")
            subio_nivel = (nivel_prev == "NORMAL") or (nivel_prev == "PREVENTIVO" and nivel_act == "CRITICO")
            escalo_temp = t_act is not None and previo["max_temp"] > 0 and t_act >= previo["max_temp"] + DELTA_TEMP_ESCALAMIENTO
            escalo_hum = h_act is not None and previo["max_hum"] > 0 and h_act >= previo["max_hum"] + DELTA_HUM_ESCALAMIENTO
            escalo_exp = exp_act is not None and previo["max_exp"] > 0 and exp_act >= previo["max_exp"] + DELTA_SENS_ESCALAMIENTO

            toca_recordatorio_critico = False
            if nivel_act == "CRITICO" and previo.get("ultimo_aviso_ts"):
                try:
                    mins_desde_aviso = (ahora_dt - datetime.fromisoformat(previo["ultimo_aviso_ts"])).total_seconds() / 60.0
                    if mins_desde_aviso >= RECORDATORIO_CRITICO_MIN:
                        toca_recordatorio_critico = True
                except Exception:
                    pass

            debe_notificar = subio_nivel or escalo_temp or escalo_hum or escalo_exp or toca_recordatorio_critico

            if debe_notificar:
                motivos = diagnosticar_motivos(g, previo)
                if toca_recordatorio_critico and not (subio_nivel or escalo_temp or escalo_hum or escalo_exp):
                    motivos.insert(0, f"Persiste en estado crítico (+{RECORDATORIO_CRITICO_MIN} min sin bajar)")

                item_alerta = (g, motivos)
                if nivel_act == "CRITICO":
                    alertas_criticas.append(item_alerta)
                else:
                    alertas_preventivas.append(item_alerta)

                if t_act is not None:
                    previo["max_temp"] = max(previo["max_temp"], t_act)
                if h_act is not None:
                    previo["max_hum"] = max(previo["max_hum"], h_act)
                if exp_act is not None:
                    previo["max_exp"] = max(previo["max_exp"], exp_act)
                if not previo.get("inicio_alerta"):
                    previo["inicio_alerta"] = ahora_dt.isoformat()
                previo["ultimo_aviso_ts"] = ahora_dt.isoformat()

            previo["nivel"] = nivel_act
            previo["ultima_temp"] = t_act
            previo["ultima_hum"] = h_act
            estado["galpones"][clave_g] = previo

        slot_resumen = toca_resumen_programado(ahora_dt, estado["resumenes"])
        enviar_ahora = forzar_reporte

        if primer_ciclo:
            primer_ciclo = False
            if ENVIAR_RESUMEN_AL_INICIAR:
                enviar_ahora = True

        if slot_resumen:
            enviar_ahora = True
            estado["resumenes"].append(slot_resumen)

        guardar_estado(estado)

    # Notificaciones técnicas de hardware/enlace (solo en turno operativo 06:00 - 22:00)
    if avisos_tecnicos and 6 <= ahora_dt.hour < 22:
        enviar_telegram(
            f"🛠️ <b>AVISO TÉCNICO DE TELEMETRÍA — GRANJA MAMAY</b>\n"
            f"📅 {ahora_dt.strftime('%d/%m/%Y')} | 🕒 {hora_actual} hrs\n\n"
            + "\n".join(avisos_tecnicos)
        )

    # Notificaciones climáticas
    if enviar_ahora:
        enviar_telegram(construir_reporte_completo(galpones, ahora_dt))
    else:
        if alertas_criticas or alertas_preventivas:
            todas_alertas = [g for g, _ in alertas_criticas] + [g for g, _ in alertas_preventivas]
            zonas_afectadas = {obtener_encargado(g["num"]) for g in todas_alertas}

            if len(todas_alertas) == 1:
                g_unico = todas_alertas[0]
                link_alerta = f"{URL_DASHBOARD}?galpon=G{g_unico['num']:02d}"
                txt_btn = f"Ver Gráfica Galpón {g_unico['num']:02d}"
            elif len(zonas_afectadas) == 1:
                z_unica = next(iter(zonas_afectadas))
                link_alerta = f"{URL_DASHBOARD}?zona={z_unica}"
                txt_btn = f"Ver Zona {z_unica}"
            else:
                link_alerta = URL_DASHBOARD
                txt_btn = "Abrir Dashboard Climático"

            secciones = []
            if alertas_criticas:
                lineas_c = [
                    f"🔴 <b>ALERTA CRÍTICA DE ESTRÉS TÉRMICO — GRANJA MAMAY</b>\n"
                    f"📅 {ahora_dt.strftime('%d/%m/%Y')} | 🕒 {hora_actual} hrs — <i>Requiere acción inmediata</i>"
                ]
                for g, motivos in alertas_criticas:
                    zona = obtener_encargado(g["num"])
                    detalle_mot = "\n   • ".join(motivos) if motivos else "Umbral crítico superado"
                    lineas_c.append(
                        f"\n<b>Galpón {g['num']:02d} ({zona})</b>\n"
                        f"   • {detalle_mot}\n"
                        f"   • Lectura: Int <b>{fmt_val(g['temp'], ' °C')}</b> | Hum <b>{fmt_val(g['hum'], ' %', 0)}</b> | "
                        f"Sens <b>{fmt_val(g['exp'], ' °C')}</b> | Ext {fmt_val(g['ext'], ' °C')}"
                    )
                secciones.append("\n".join(lineas_c))

            if alertas_preventivas:
                lineas_p = [
                    f"🟡 <b>AVISO PREVENTIVO CLIMÁTICO — GRANJA MAMAY</b>\n"
                    f"📅 {ahora_dt.strftime('%d/%m/%Y')} | 🕒 {hora_actual} hrs — <i>Monitorear ventilación</i>"
                ]
                for g, motivos in alertas_preventivas:
                    zona = obtener_encargado(g["num"])
                    detalle_mot = "\n   • ".join(motivos) if motivos else "Umbral preventivo superado"
                    lineas_p.append(
                        f"\n<b>Galpón {g['num']:02d} ({zona})</b>\n"
                        f"   • {detalle_mot}\n"
                        f"   • Lectura: Int <b>{fmt_val(g['temp'], ' °C')}</b> | Hum <b>{fmt_val(g['hum'], ' %', 0)}</b> | "
                        f"Sens <b>{fmt_val(g['exp'], ' °C')}</b> | Ext {fmt_val(g['ext'], ' °C')}"
                    )
                secciones.append("\n".join(lineas_p))

            enviar_telegram("\n\n".join(secciones), url_boton=link_alerta, texto_boton=txt_btn)

        if normalizados:
            lineas_norm = [
                f"🟢 <b>NORMALIZACIÓN CLIMÁTICA — GRANJA MAMAY</b>\n"
                f"📅 {ahora_dt.strftime('%d/%m/%Y')} | 🕒 {hora_actual} hrs\n"
            ]
            for g, dur in normalizados:
                lineas_norm.append(
                    f"• <b>Galpón {g['num']:02d} ({obtener_encargado(g['num'])})</b> volvió a óptimo{dur}: "
                    f"Int <b>{fmt_val(g['temp'], ' °C')}</b> | Hum <b>{fmt_val(g['hum'], ' %', 0)}</b> | "
                    f"Sens <b>{fmt_val(g['exp'], ' °C')}</b>"
                )
            enviar_telegram("\n".join(lineas_norm))


if __name__ == "__main__":
    modo_nube = "--una-vez" in sys.argv
    if not modo_nube:
        adquirir_candado()
    logging.info(f"Servicio Monitor Mamay iniciado (modo={'nube-1h' if modo_nube else 'local-5m'}).")

    if not modo_nube:
        hilo_bot = threading.Thread(target=escuchar_comandos_telegram, daemon=True)
        hilo_bot.start()

    for _ in range(15):
        if hay_internet():
            break
        time.sleep(3)

    if modo_nube:
        ultimo_error = None
        for intento in range(1, 4):
            try:
                logging.info(f"Iniciando intento {intento}/3 de extracci?n horaria...")
                procesar_ciclo()
                logging.info("Ciclo horario en nube finalizado correctamente.")
                sys.exit(0)
            except Exception as e:
                ultimo_error = e
                logging.error(f"Fallo en intento {intento}/3: {e}")
                if intento < 3:
                    espera = intento * 20
                    logging.info(f"Reintentando en {espera} segundos...")
                    time.sleep(espera)
        logging.critical(f"Abortado tras 3 intentos fallidos: {ultimo_error}")
        sys.exit(1)

    activar_teclado_consulta(TELEGRAM_CHAT_ID)

    while True:
        try:
            procesar_ciclo()
        except KeyboardInterrupt:
            logging.info("Ejecución detenida manualmente por el usuario.")
            break
        except Exception as e:
            logging.error(f"Error no esperado en ciclo principal (continuando): {e}")
        time.sleep(INTERVALO_SEGUNDOS)








