import sys
import tempfile
import unittest
from pathlib import Path


PASTA_SCRIPT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PASTA_SCRIPT))

from lote_usuarios_pwdm_v2 import carregar_emails_xlsx, deduplicar_emails


class LoteUsuariosPwdmTests(unittest.TestCase):
    def test_deduplicar_emails_preserva_ordem(self):
        unicos, duplicados = deduplicar_emails(
            [
                "Usuario1@Empresa.com",
                "usuario2@empresa.com",
                "usuario1@empresa.com",
            ]
        )

        self.assertEqual(unicos, ["usuario1@empresa.com", "usuario2@empresa.com"])
        self.assertEqual(duplicados, ["usuario1@empresa.com"])

    def test_carregar_emails_xlsx_lendo_coluna_email(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "usuarios.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["nome", "email"])
            sheet.append(["Usuario 1", "Usuario1@Empresa.com"])
            sheet.append(["Usuario 2", "usuario2@empresa.com"])
            sheet.append(["Duplicado", "usuario1@empresa.com"])
            sheet.append(["Invalido", "sem-email"])
            sheet.append(["Vazio", ""])
            workbook.save(caminho)

            resultado = carregar_emails_xlsx(caminho)

        self.assertEqual(
            resultado["emailsValidosUnicos"],
            ["usuario1@empresa.com", "usuario2@empresa.com"],
        )
        self.assertEqual(resultado["duplicados"], ["usuario1@empresa.com"])
        self.assertEqual(resultado["invalidos"], [{"linha": 5, "email": "sem-email"}])
        self.assertEqual(resultado["linhasSemEmail"], [6])

    def test_carregar_emails_xlsx_exige_coluna_email(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "usuarios.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["usuario"])
            sheet.append(["usuario1@empresa.com"])
            workbook.save(caminho)

            with self.assertRaisesRegex(ValueError, "coluna chamada 'email'"):
                carregar_emails_xlsx(caminho)


if __name__ == "__main__":
    unittest.main()
