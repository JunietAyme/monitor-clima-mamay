# Monitor Climático Automatizado — Granja Avícola Mamay

Sistema integral de telemetría climática, alerta temprana y analítica histórica para 18 galpones avícolas divididos en dos zonas operativas (**Zona Arias: G01–G09** y **Zona Castillo: G10–G18**).

## Arquitectura del Sistema

1. **Extracción Automatizada (`monitor_mamay.py` + GitHub Actions):**
   - Ejecución horaria automatizada en la nube (`0 * * * *` con `--una-vez`) y soporte para ejecución continua local con bloqueo de instancia única (`monitor.lock`).
   - Automatización con **Playwright** sobre sesión en *BigFarmNet Fusion Survey* (`Temperatura Interior`, `Temperatura Exterior`, `Humedad Relativa` y `Sensación Térmica`).
2. **Motor de Alertas de 3 Niveles y Resiliencia Técnica:**
   - **Preventivo:** `Temp > 23.0 °C` | `Humedad > 72 %` | `Sensación Térmica > 25.5 °C`.
   - **Crítico:** `Temp >= 26.0 °C` | `Humedad > 80 %` | `Sensación Térmica >= 28.0 °C` (con recordatorio persistente y notificación de normalización).
   - **Auditoría de Hardware:** Detección de pérdida de enlace con BigFarmNet y sensores congelados sin variación sostenida.
3. **Bot Interactivo en Telegram:**
   - Panel de consulta bajo demanda (`Estado General`, `Solo Alertas`, `Zona Arias`, `Zona Castillo`) respondido desde caché local y comandos técnicos.
4. **Reportes Ejecutivos y Dashboard Web (`generar_excel_clima.py` + `index.html`):**
   - Consolidación histórica deduplicada con lecturas en vivo.
   - Dashboard web en **GitHub Pages** con filtros por zona, mes, día y semana operativa, cruce clima × producción y exportación a Excel multi-hoja mediante **ExcelJS**.

## Configuración Local

1. Copiar `config_secreta.example.json` como `config_secreta.json` y completar `telegram_token`, `telegram_chat_id` y `correo_bfn`.
2. Ejecutar `guardar_sesion.py` o `renovar_sesion.ps1` para autenticar o exportar la sesión de BigFarmNet.
3. Ejecutar `python monitor_mamay.py --una-vez` para una prueba puntual o `pythonw.exe monitor_mamay.py` para ejecución continua.
