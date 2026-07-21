import sys
import tempfile
import unittest
from pathlib import Path


PASTA_SCRIPT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PASTA_SCRIPT))

from incluir_usuarios_projetos_roles_pwdm_v2 import (  # noqa: E402
    LinhaInclusao,
    carregar_linhas_inclusao_xlsx,
    extrair_id_usuario,
    extrair_role_ids_usuario,
    extrair_roles_de_membros,
    normalizar_role,
    resolver_role,
    resolver_projetos_linha,
    valor_bool,
)


class IncluirUsuariosProjetosRolesTests(unittest.TestCase):
    def test_carregar_linhas_inclusao_xlsx_com_role(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "inclusoes.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["email", "projeto", "role"])
            sheet.append(["Usuario1@Empresa.com", "Projeto A", "PWDM Issuer"])
            sheet.append(["sem-email", "Projeto A", "Team Member"])
            workbook.save(caminho)

            resultado = carregar_linhas_inclusao_xlsx(caminho)

        linhas = resultado["linhas"]
        self.assertEqual(len(linhas), 1)
        self.assertEqual(linhas[0].email, "usuario1@empresa.com")
        self.assertEqual(linhas[0].projeto, "Projeto A")
        self.assertEqual(linhas[0].role, "PWDM Issuer")
        self.assertEqual(resultado["invalidas"][0]["motivo"], "email_invalido")

    def test_carregar_linhas_inclusao_xlsx_sem_role(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "inclusoes.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["email", "projeto"])
            sheet.append(["Usuario1@Empresa.com", "Projeto A"])
            workbook.save(caminho)

            resultado = carregar_linhas_inclusao_xlsx(caminho)

        linhas = resultado["linhas"]
        self.assertEqual(len(linhas), 1)
        self.assertEqual(linhas[0].email, "usuario1@empresa.com")
        self.assertEqual(linhas[0].projeto, "Projeto A")
        self.assertEqual(linhas[0].role, "")

    def test_carregar_linhas_inclusao_xlsx_apenas_com_coluna_email(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as pasta:
            caminho = Path(pasta) / "usuarios.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["E-mail"])
            sheet.append(["Usuario1@Empresa.com"])
            sheet.append(["usuario2@empresa.com"])
            workbook.save(caminho)

            resultado = carregar_linhas_inclusao_xlsx(caminho)

        linhas = resultado["linhas"]
        self.assertEqual([linha.email for linha in linhas], ["usuario1@empresa.com", "usuario2@empresa.com"])
        self.assertEqual([linha.projeto for linha in linhas], ["todos", "todos"])

    def test_resolver_projetos_linha_por_projectwise_id(self):
        projetos = [
            {
                "nome": "Projeto A",
                "projectId": "11111111-1111-1111-1111-111111111111",
                "connectSpaceId": "11111111-1111-1111-1111-111111111111",
                "origemProjectWise": {"id": "123", "nome": "Projeto A"},
            }
        ]
        linha = LinhaInclusao(
            linha=2,
            email="usuario@empresa.com",
            projeto="123",
            role="Team Member",
        )

        encontrados, criterio = resolver_projetos_linha(linha, projetos)

        self.assertEqual(criterio, "exato")
        self.assertEqual(encontrados, projetos)

    def test_resolver_projetos_linha_todos(self):
        projetos = [
            {"nome": "Projeto A", "projectId": "1", "connectSpaceId": "1", "origemProjectWise": {}},
            {"nome": "Projeto B", "projectId": "2", "connectSpaceId": "2", "origemProjectWise": {}},
        ]
        linha = LinhaInclusao(
            linha=2,
            email="usuario@empresa.com",
            projeto="todos",
            role="Team Member",
        )

        encontrados, criterio = resolver_projetos_linha(linha, projetos)

        self.assertEqual(criterio, "todos")
        self.assertEqual(encontrados, projetos)

    def test_valor_bool_aceita_s_n_e_x(self):
        self.assertEqual(valor_bool("S"), True)
        self.assertEqual(valor_bool("x"), True)
        self.assertEqual(valor_bool("nao"), False)
        self.assertIsNone(valor_bool("talvez"))

    def test_resolver_role_por_nome_exato(self):
        roles = [
            {"id": "1", "displayName": "Team Member"},
            {"id": "2", "displayName": "PWDM Issuer"},
        ]

        role = resolver_role("pwdm issuer", roles)

        self.assertEqual(role["id"], "2")

    def test_normalizar_role_aceita_role_id(self):
        role = normalizar_role({"roleId": "abc", "name": "Team Member"}, "iTwin", "itwin-1")

        self.assertEqual(role["id"], "abc")
        self.assertEqual(role["displayName"], "Team Member")

    def test_extrair_ids_usuario_em_formatos_diferentes(self):
        usuario = {
            "userId": "user-1",
            "roleIds": ["role-1"],
            "roles": [{"roleId": "role-2"}, {"id": "role-3"}],
        }

        self.assertEqual(extrair_id_usuario(usuario), "user-1")
        self.assertEqual(extrair_role_ids_usuario(usuario), ["role-1", "role-2", "role-3"])

    def test_extrair_roles_de_membros(self):
        usuarios = [
            {
                "email": "usuario@empresa.com",
                "roles": [
                    {"id": "team-member-id", "displayName": "Team Member"},
                    {"roleId": "imodel-id", "name": "iModel", "type": "enterprise"},
                ],
            }
        ]

        roles = extrair_roles_de_membros(usuarios, "itwin-1")

        self.assertEqual([role["displayName"] for role in roles], ["iModel", "Team Member"])
        self.assertEqual(roles[0]["type"], "enterprise")


if __name__ == "__main__":
    unittest.main()
