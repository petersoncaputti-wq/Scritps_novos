import sys
import unittest
from pathlib import Path
from unittest.mock import patch


PASTA_SCRIPT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PASTA_SCRIPT))

import gerenciar_participante_pwdm_connected_v2 as connected
from gerenciar_participante_pwdm_v2 import sanitizar_para_log


class ExclusaoParticipantePwdmTests(unittest.TestCase):
    def montar_operacao(self):
        return {
            "status": "participante",
            "acaoSolicitada": "excluir",
            "acaoEfetiva": "excluir_participante",
            "connectSpaceId": "11111111-1111-1111-1111-111111111111",
            "projectId": "22222222-2222-2222-2222-222222222222",
            "membro": {"id": "33333333-3333-3333-3333-333333333333"},
            "permissoes": {},
        }

    @patch.object(connected, "post_json_em_aba")
    def test_exclusao_usa_endpoint_e_payload_capturados(self, post_json):
        post_json.return_value = {"body": {"isCompleted": True}, "status": 200}

        resultado = connected.aplicar_operacao_em_aba(
            object(),
            self.montar_operacao(),
            "usuario@empresa.com",
        )

        self.assertEqual(resultado["status"], "excluido")
        _, _, endpoint, payload = post_json.call_args.args
        self.assertTrue(endpoint.endswith("/GenericDeleteItems"))
        self.assertEqual(payload["className"], "Participants")
        self.assertEqual(
            payload["ids"],
            ["dm:33333333-3333-3333-3333-333333333333:usuario@empresa.com"],
        )

    @patch.object(connected, "post_json_em_aba")
    def test_exclusao_exige_confirmacao_is_completed(self, post_json):
        post_json.return_value = {"body": {"isCompleted": False}, "status": 200}

        with self.assertRaisesRegex(RuntimeError, "nao confirmou"):
            connected.aplicar_operacao_em_aba(
                object(),
                self.montar_operacao(),
                "usuario@empresa.com",
            )

    @patch("builtins.input", return_value="CONFIRMAR EXCLUSAO")
    def test_confirmacao_reforcada_autoriza_exclusao(self, _input):
        self.assertTrue(connected.confirmar_aplicacao_sn([self.montar_operacao()]))

    @patch("builtins.input", return_value="S")
    def test_confirmacao_simples_nao_autoriza_exclusao(self, _input):
        self.assertFalse(connected.confirmar_aplicacao_sn([self.montar_operacao()]))

    def test_log_mascara_email_embutido_no_id_de_exclusao(self):
        valor = "dm:33333333-3333-3333-3333-333333333333:usuario@empresa.com"

        sanitizado = sanitizar_para_log(valor)

        self.assertEqual(
            sanitizado,
            "dm:33333333-3333-3333-3333-333333333333:us***@empresa.com",
        )


if __name__ == "__main__":
    unittest.main()
