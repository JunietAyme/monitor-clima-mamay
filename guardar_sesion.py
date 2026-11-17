from playwright.sync_api import sync_playwright

PERFIL_DIR = r"C:\Scripts\Monitor_Mamay\perfil_bfn"
URL_SURVEY = "https://survey.bfn-fusion.bigdutchman.com/"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=PERFIL_DIR,
        headless=False,
        viewport={"width": 1440, "height": 900},
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.goto(URL_SURVEY)
    print("\n>>> 1. Inicia sesión en la ventana que se abrió y entra a 'TEMPERATURA -HUMEDAD'.")
    print(">>> 2. Cuando veas los 15 galpones en pantalla, CIERRA esa ventana del navegador.")
    page.wait_for_event("close", timeout=0)
    context.close()
    print("Sesión guardada correctamente en:", PERFIL_DIR)
