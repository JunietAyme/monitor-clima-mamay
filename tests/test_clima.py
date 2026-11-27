import unittest
from pathlib import Path


class TestMonitorClima(unittest.TestCase):
    def test_workflow_tiene_cron_horario(self):
        yml = Path(".github/workflows/monitor_horario.yml").read_text(encoding="utf-8")
        self.assertIn("0 * * * *", yml)
        self.assertIn("--una-vez", yml)

    def test_historial_usa_humedad_critica_estricta(self):
        src = Path("importar_historial_bfn.py").read_text(encoding="utf-8")
        self.assertIn("h > HUM_CRIT", src)

    def test_guardar_sesion_usa_ruta_relativa(self):
        src = Path("guardar_sesion.py").read_text(encoding="utf-8")
        self.assertIn("Path(__file__).resolve().parent", src)

    def test_dashboard_tiene_galpones_en_linea_y_escape_html(self):
        html = Path("index.html").read_text(encoding="utf-8")
        self.assertIn("galponesEnLinea", html)
        self.assertIn("function escapeHtml", html)
        self.assertIn("idxGalponTiempo", html)


if __name__ == "__main__":
    unittest.main()
