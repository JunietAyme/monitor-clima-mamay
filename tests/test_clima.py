import unittest
from datetime import datetime
import monitor_mamay as mm
import generar_excel_clima as ge


class TestMonitorClimaMamay(unittest.TestCase):
    def test_clasificacion_limites_exactos(self):
        cod, _ = mm.clasificar_estado(22.0, 80.0)
        self.assertEqual(cod, "PREVENTIVO")

        cod, _ = mm.clasificar_estado(22.0, 80.1)
        self.assertEqual(cod, "CRITICO")

        cod, _ = mm.clasificar_estado(26.0, 65.0)
        self.assertEqual(cod, "CRITICO")

        cod, _ = mm.clasificar_estado(22.5, 72.0)
        self.assertEqual(cod, "NORMAL")

    def test_diagnosticar_motivos(self):
        galpon = {"temp": 26.5, "hum": 82.0, "exp": 28.5, "ext": 22.0}
        previo = {"ultima_temp": 25.0, "ultima_hum": 78.0}
        motivos = mm.diagnosticar_motivos(galpon, previo)
        texto = " ".join(motivos)
        self.assertIn("crítica", texto)
        self.assertIn("+1.5 °C", texto)
        self.assertIn("+4 %", texto)
        self.assertIn("Diferencial", texto)

    def test_evaluar_salud_sensores_firma_completa(self):
        estado = {}
        galpon = {1: {"temp": 22.0, "hum": 70.0, "ext": 20.0, "exp": 22.0}}
        mm.evaluar_salud_sensores(estado, galpon)
        self.assertEqual(estado["tecnico"]["1"]["firma_previa"], "22.0|70.0|20.0|22.0")

    def test_ventana_resumen_programado(self):
        dt = datetime(2026, 11, 17, 12, 3)
        self.assertEqual(mm.toca_resumen_programado(dt, []), "12:00")
        self.assertIsNone(mm.toca_resumen_programado(dt, ["12:00"]))

    def test_normalizacion_fechas_excel(self):
        f, h = ge.normalizar_fecha_hora("17/11/2026 08:35")
        self.assertEqual((f, h), ("2026-11-17", "08:35"))
        f2, h2 = ge.normalizar_fecha_hora("2026-11-17", "14:20:05")
        self.assertEqual((f2, h2), ("2026-11-17", "14:20"))


if __name__ == "__main__":
    unittest.main()
