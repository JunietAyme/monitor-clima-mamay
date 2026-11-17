# Monitor Climático Automatizado — Granja Avícola Mamay

Sistema integral de telemetría climática, alerta temprana y analítica histórica para 17 galpones avícolas divididos en dos zonas operativas (**Zona Arias: G01–G09** y **Zona Castillo: G10–G17**).

## Arquitectura del Sistema

1. **Extracción Automatizada (`monitor_mamay.py`):**
   - Servicio residente en Windows (`pythonw`) con bloqueo de instancia única (`monitor.lock`) y espera inteligente de conectividad tras el arranque.
   - Automatización con **Playwright** sobre sesión persistente en *BigFarmNet Fusion Survey* cada 5 minutos (`Temperatura Interior`, `Temperatura Exterior`, `Humedad Relativa` y `Sensación Térmica`).
2. **Motor de Alertas de 3 Niveles y Resiliencia Técnica:**
   - **Preventivo:** `Temp > 23.0 °C` | `Humedad > 72 %` | `Sensación Térmica > 25.5 °C`.
   - **Crítico:** `Temp ≥ 26.0 °C` | `Humedad ≥ 80 %` | `Sensación Térmica ≥ 28.0 °C` (con recordatorio persistente cada 15 minutos y notificación automática de normalización).
   - **Auditoría de Hardware:** Detección de pérdida de enlace con BigFarmNet (`≥ 10 min`) y sensores congelados sin variación (`≥ 90 min`).
3. **Bot Interactivo en Telegram:**
   - Panel de consulta bajo demanda (`Estado General`, `Solo Alertas`, `Zona Arias`, `Zona Castillo`) respondido desde caché local en `< 1 s` y comando técnico `/salud`.
4. **Reportes Ejecutivos y Dashboard Web (`generar_excel_clima.py` + `index.html`):**
   - Consolidación histórica (Enero – Noviembre 2026) deduplicada con lecturas en vivo.
   - Dashboard web en **GitHub Pages** con filtros por zona, mes y jornada, comparativas visuales y exportación directa desde el navegador a Excel multi-hoja estilizado mediante **ExcelJS**.

## Configuración Local

1. Copiar `config_secreta.example.json` como `config_secreta.json` y completar `telegram_token`, `telegram_chat_id` y `correo_bfn`.
2. Ejecutar `guardar_sesion.py` una vez para autenticar la sesión persistente en `perfil_bfn/`.
3. Iniciar el servicio en segundo plano con `pythonw.exe monitor_mamay.py`.
