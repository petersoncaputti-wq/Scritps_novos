import sys
import unittest
from pathlib import Path


PASTA_SCRIPT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PASTA_SCRIPT))

import regras_v2 as pwdm


class RegrasPwdmTests(unittest.TestCase):
    def test_normalizar_permissoes_ativa_exibir_quando_receber_ativo(self):
        permissoes = pwdm.normalizar_permissoes(
            {
                "canView": False,
                "canReceive": True,
                "canIssue": False,
                "canReviewAnswer": False,
                "isAdmin": False,
            }
        )

        self.assertTrue(permissoes["canView"])
        self.assertTrue(permissoes["canReceive"])

    def test_normalizar_permissoes_emitir_exige_receber(self):
        permissoes = pwdm.normalizar_permissoes(
            {
                "canView": True,
                "canReceive": False,
                "canIssue": True,
                "canReviewAnswer": False,
                "isAdmin": False,
            }
        )

        self.assertTrue(permissoes["canReceive"])

    def test_determinar_acao_incluir_participante_atualiza(self):
        acao, descricao = pwdm.determinar_acao_efetiva("incluir", "participante")

        self.assertEqual(acao, "atualizar_participante")
        self.assertIn("ATUALIZAR", descricao)

    def test_determinar_acao_alterar_usuario_ausente_ignora(self):
        acao, descricao = pwdm.determinar_acao_efetiva("alterar", "ausente")

        self.assertEqual(acao, "ignorar_usuario_ausente")
        self.assertIn("IGNORAR", descricao)

    def test_descricao_permissoes_sem_permissoes(self):
        descricao = pwdm.descricao_permissoes({chave: False for chave in pwdm.CHAVES_PERMISSOES})

        self.assertEqual(descricao, "sem permissoes")

    def test_descricao_permissoes_lista_rotulos_esperados(self):
        descricao = pwdm.descricao_permissoes(
            {
                "canView": True,
                "canReceive": True,
                "canIssue": True,
                "canReviewAnswer": True,
                "isAdmin": True,
            }
        )

        self.assertEqual(descricao, "Exibir, Receber, Emitir, Aprovar RFI, Admin")

    def test_mascarar_email_preserva_dominio(self):
        self.assertEqual(pwdm.mascarar_email("usuario.teste@empresa.com"), "us***@empresa.com")

    def test_extrair_ids_de_url_participantes(self):
        connect_space_id_esperado = "11111111-1111-1111-1111-111111111111"
        project_id_esperado = "22222222-2222-2222-2222-222222222222"
        url = (
            "https://pwdm.bentley.com/"
            f"{connect_space_id_esperado}/ProjectSettings/{project_id_esperado}/View#PARTICIPANTS"
        )

        connect_space_id, project_id = pwdm.extrair_ids_de_url(url)

        self.assertEqual(connect_space_id, connect_space_id_esperado)
        self.assertEqual(project_id, project_id_esperado)

    def test_titulo_igual_ignora_acentos_e_caixa(self):
        membro = {"roleTitle": "Coordenacao Tecnica"}

        self.assertTrue(pwdm.titulo_igual(membro, "coordenação técnica"))

    def test_permissoes_iguais_compara_chaves_principais(self):
        membro = {
            "canView": True,
            "canReceive": False,
            "canIssue": True,
            "canReviewAnswer": False,
            "isAdmin": False,
        }
        permissoes = {
            "canView": True,
            "canReceive": False,
            "canIssue": True,
            "canReviewAnswer": False,
            "isAdmin": False,
        }

        self.assertTrue(pwdm.permissoes_iguais(membro, permissoes))


if __name__ == "__main__":
    unittest.main()
