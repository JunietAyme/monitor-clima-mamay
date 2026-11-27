import csv
import importlib.util
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def _cargar_monitor_mamay():
    spec = importlib.util.spec_from_file_location("monitor_mamay", REPO_ROOT / "monitor_mamay.py")
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


mm = _cargar_monitor_mamay()


class TestMonitorClima(unittest.TestCase):
    def test_workflow_tiene_cron_horario(self):
        yml = (REPO_ROOT / ".github/workflows/monitor_horario.yml").read_text(encoding="utf-8")
        self.assertIn("0 * * * *", yml)
        self.assertIn("--una-vez", yml)

    def test_historial_usa_humedad_critica_estricta(self):
        src = (REPO_ROOT / "importar_historial_bfn.py").read_text(encoding="utf-8-sig")
        self.assertIn("h > HUM_CRIT", src)

    def test_guardar_sesion_usa_ruta_relativa(self):
        src = (REPO_ROOT / "guardar_sesion.py").read_text(encoding="utf-8-sig")
        self.assertIn("Path(__file__).resolve().parent", src)

    def test_dashboard_tiene_galpones_en_linea_y_escape_html(self):
        html = (REPO_ROOT / "index.html").read_text(encoding="utf-8")
        self.assertIn("galponesEnLinea", html)
        self.assertIn("function escapeHtml", html)
        self.assertIn("idxGalponTiempo", html)

    def test_clasificar_estado_umbrales(self):
        estado_norm, _ = mm.clasificar_estado(23.0, 72.0, 23.0)
        self.assertEqual(estado_norm, "NORMAL")

        estado_prev, _ = mm.clasificar_estado(23.5, 72.0, 23.5)
        self.assertEqual(estado_prev, "PREVENTIVO")

        estado_crit_temp, _ = mm.clasificar_estado(26.0, 70.0, 26.0)
        self.assertEqual(estado_crit_temp, "CRITICO")

        estado_crit_hum, _ = mm.clasificar_estado(22.0, 81.0, 22.0)
        self.assertEqual(estado_crit_hum, "CRITICO")

    def test_guardar_historial_csv_deduplica_por_minuto_y_galpon(self):
        galpones = {
            1: {"temp": 22.5, "ext": 20.0, "hum": 65.0, "exp": 22.5},
        }
        fijo = datetime(2026, 11, 27, 10, 30, 15)
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_tmp = str(Path(tmpdir) / "historial_test.csv")
            with patch.object(mm, "ARCHIVO_CSV", csv_tmp), patch.object(mm, "ahora_lima", return_value=fijo):
                mm.guardar_historial_csv(fijo, galpones)
                mm.guardar_historial_csv(fijo, galpones)

            with open(csv_tmp, mode="r", encoding="utf-8-sig", newline="") as f:
                filas = list(csv.reader(f, delimiter=";"))
            self.assertEqual(len(filas), 2)


if __name__ == "__main__":
    unittest.main()
