import json
import sys
import tempfile
import unittest
from pathlib import Path


PASTA_SCRIPT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PASTA_SCRIPT))

from projetos_projectwise_pwdm_v2 import carregar_projetos_selecionados, carregar_projetos_selecionados_com_diagnostico


class ProjetosProjectWisePwdmTests(unittest.TestCase):
    def test_carregar_projetos_selecionados_normaliza_connected_project_id(self):
        dados = {
            "concessao": {"nome": "Ecovias Araguaia"},
            "projetos": [
                {
                    "nome": "Acesso 3 Ano",
                    "projectWiseId": "58230",
                    "projectWiseGuid": "7cf8e412-d262-4ac7-a7a2-c65cf3180033",
                    "connectedProjectId": "0f10d4a0-17c4-4ab7-8c46-102982815fbd",
                    "urlParticipantesPwdm": (
                        "https://pwdm.bentley.com/0f10d4a0-17c4-4ab7-8c46-102982815fbd/"
                        "ProjectSettings/0f10d4a0-17c4-4ab7-8c46-102982815fbd/View#PARTICIPANTS"
                    ),
                }
            ],
        }

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "projetos.json"
            caminho.write_text(json.dumps(dados), encoding="utf-8")

            projetos = carregar_projetos_selecionados(caminho)

        self.assertEqual(len(projetos), 1)
        projeto = projetos[0]
        self.assertEqual(projeto["connectSpaceId"], "0f10d4a0-17c4-4ab7-8c46-102982815fbd")
        self.assertEqual(projeto["projectId"], "0f10d4a0-17c4-4ab7-8c46-102982815fbd")
        self.assertEqual(projeto["criterioCruzamento"], "ProjectWise ConnectedProjectId")
        self.assertEqual(projeto["origemProjectWise"]["id"], "58230")

    def test_carregar_projetos_selecionados_ignora_connected_project_id_zerado(self):
        dados = {
            "concessao": {"nome": "Ecovias Araguaia"},
            "projetos": [
                {
                    "nome": "Projeto sem PWDM",
                    "projectWiseId": "1",
                    "connectedProjectId": "00000000-0000-0000-0000-000000000000",
                },
                {
                    "nome": "Projeto com PWDM",
                    "projectWiseId": "2",
                    "connectedProjectId": "0f10d4a0-17c4-4ab7-8c46-102982815fbd",
                },
            ],
        }

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "projetos.json"
            caminho.write_text(json.dumps(dados), encoding="utf-8")

            projetos = carregar_projetos_selecionados(caminho)

        self.assertEqual(len(projetos), 1)
        self.assertEqual(projetos[0]["nome"], "Projeto com PWDM")

    def test_carregar_projetos_selecionados_com_diagnostico_retorna_nome_ignorado(self):
        dados = {
            "concessao": {"nome": "Ecovias Araguaia"},
            "projetos": [
                {
                    "nome": "Projeto sem PWDM",
                    "projectWiseId": "123",
                    "connectedProjectId": "00000000-0000-0000-0000-000000000000",
                },
                {
                    "nome": "Projeto com PWDM",
                    "projectWiseId": "456",
                    "connectedProjectId": "0f10d4a0-17c4-4ab7-8c46-102982815fbd",
                },
            ],
        }

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "projetos.json"
            caminho.write_text(json.dumps(dados), encoding="utf-8")

            projetos, ignorados = carregar_projetos_selecionados_com_diagnostico(caminho)

        self.assertEqual(len(projetos), 1)
        self.assertEqual(len(ignorados), 1)
        self.assertIn("Projeto sem PWDM", ignorados[0])
        self.assertIn("PW ID: 123", ignorados[0])


if __name__ == "__main__":
    unittest.main()
