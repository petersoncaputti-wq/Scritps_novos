from openpyxl import Workbook
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import criar_usuarios as kartado


class CadastroTests(unittest.TestCase):
    def arquivo(self, linhas):
        pasta = tempfile.TemporaryDirectory()
        self.addCleanup(pasta.cleanup)
        caminho = Path(pasta.name) / 'entrada.xlsx'
        workbook = Workbook()
        workbook.active.append(['nome', 'sobrenome', 'email', 'login'])
        for linha in linhas:
            workbook.active.append(linha)
        workbook.save(caminho)
        workbook.close()
        return caminho

    def test_limpa_espacos_e_preserva_login(self):
        caminho = self.arquivo([[' Ana ', ' Silva ', 'ana@example.com', 'Ana.Silva']])
        self.assertEqual(kartado.carregar_excel(caminho)[0]['nome'], 'Ana')

    def test_bloqueia_duplicidade_sem_diferenciar_caixa(self):
        caminho = self.arquivo([['Ana', 'Silva', 'a@example.com', 'Ana'], ['Ana', 'Souza', 'b@example.com', 'ana']])
        with self.assertRaisesRegex(ValueError, 'repetido'):
            kartado.carregar_excel(caminho)

    def test_bloqueia_formulas(self):
        caminho = self.arquivo([['=1+1', 'Silva', 'a@example.com', 'ana']])
        with self.assertRaisesRegex(ValueError, 'formulas'):
            kartado.carregar_excel(caminho)

    def test_ignora_linhas_vazias_preservando_numero(self):
        caminho = self.arquivo([[None] * 4, ['Ana', 'Silva', 'a@example.com', 'ana']])
        self.assertEqual(kartado.carregar_excel(caminho)[0]['linha'], 3)

    def test_aba_inexistente(self):
        caminho = self.arquivo([['Ana', 'Silva', 'a@example.com', 'ana']])
        with self.assertRaisesRegex(ValueError, 'Aba nao encontrada'):
            kartado.carregar_excel(caminho, 'Inexistente')

    def test_recusa_xls(self):
        with self.assertRaisesRegex(ValueError, 'salve como .xlsx'):
            kartado.carregar_excel(Path('entrada.xls'))

    def test_payload_igual_ao_contrato_capturado(self):
        usuario = {'nome': 'Ana', 'sobrenome': 'Silva', 'email': 'a@example.com', 'login': 'ana'}
        data = kartado.payload_usuario(usuario, 'unidade', 'grupo', 'perfil')['data']
        self.assertIs(data['attributes']['createAccessRequest'], True)
        self.assertEqual(data['relationships']['permissions']['data'], {'type': 'UserPermission', 'id': 'perfil'})
        self.assertEqual(data['relationships']['company']['data']['id'], 'unidade')
        self.assertNotIn('password', data['attributes'])

    def test_bloqueia_url_externa_antes_de_enviar_credenciais(self):
        with self.assertRaises(ValueError):
            kartado.Cliente('teste', 'segredo').requisitar('https://outro.example/Company/')

    def test_paginacao(self):
        cliente = kartado.Cliente('teste', 'segredo')
        with patch.object(cliente, 'requisitar', side_effect=[
            (200, {'data': [{'id': '1'}], 'links': {'next': kartado.BASE + 'Company/?page=2'}}),
            (200, {'data': [{'id': '2'}], 'links': {'next': None}}),
        ]) as chamada:
            self.assertEqual(len(cliente.listar('Company/')), 2)
            self.assertEqual(chamada.call_count, 2)

    def test_validacao_offline_nao_cria_cliente(self):
        caminho = self.arquivo([['Ana', 'Silva', 'a@example.com', 'ana']])
        with patch('sys.argv', ['script', '--arquivo', str(caminho), '--validar']), patch.object(kartado, 'Cliente') as cliente:
            kartado.main()
            cliente.assert_not_called()

    def test_previa_consulta_usuarios_no_contexto_da_unidade(self):
        caminho = self.arquivo([['Ana', 'Silva', 'a@example.com', 'ana']])
        unidade = {'id': 'unidade-teste', 'attributes': {'name': 'Unidade teste'},
                   'relationships': {'companyGroup': {'data': {'id': '6d356a3c-8adf-49f7-b545-528da6138f6f'}}}}
        perfil = {'id': 'perfil-teste', 'attributes': {'name': 'Perfil teste'}}
        with patch('sys.argv', ['script', '--arquivo', str(caminho), '--somente-previa']), patch('builtins.input', return_value='teste'), patch.object(kartado.getpass, 'getpass', return_value='senha-ficticia'), patch.object(kartado, 'selecionar', side_effect=[unidade, perfil]), patch.object(kartado, 'Cliente') as classe:
            cliente = classe.return_value
            cliente.listar.side_effect = [[unidade], [perfil], []]
            kartado.main()
            self.assertEqual(cliente.listar.call_args_list[-1].args,
                             ('User/', {'company': 'unidade-teste', 'sort': 'first_name', 'only_company': 'unidade-teste'}))
            cliente.requisitar.assert_not_called()

    def test_fluxo_padrao_exige_aprovacao_antes_do_post(self):
        caminho = self.arquivo([['Ana', 'Silva', 'a@example.com', 'ana']])
        unidade = {'id': 'unidade', 'attributes': {'name': 'Teste'}, 'relationships': {'companyGroup': {'data': {'id': '6d356a3c-8adf-49f7-b545-528da6138f6f'}}}}
        perfil = {'id': 'perfil', 'attributes': {'name': 'Campo'}}
        for aprovacao, esperado in [('NAO', 0), ('', 0), ('CRIAR', 1)]:
            with self.subTest(aprovacao=aprovacao), tempfile.TemporaryDirectory() as saida:
                with patch('sys.argv', ['script', '--arquivo', str(caminho), '--saida', saida]), patch('builtins.input', side_effect=['teste', aprovacao]), patch.object(kartado.getpass, 'getpass', return_value='ficticia'), patch.object(kartado, 'selecionar', side_effect=[unidade, perfil]), patch.object(kartado, 'Cliente') as classe:
                    cliente = classe.return_value
                    cliente.listar.side_effect = [[unidade], [perfil], []]
                    cliente.requisitar.return_value = (201, {'data': {'id': 'novo', 'attributes': {'active': False}}})
                    kartado.main()
                    self.assertEqual(cliente.requisitar.call_count, esperado)

    def test_cancelar_seletor_nao_conecta(self):
        with patch('sys.argv', ['script']), patch.object(kartado, 'selecionar_planilha', return_value=None), patch.object(kartado, 'Cliente') as cliente:
            kartado.main()
            cliente.assert_not_called()

    def test_selecao_visual_valida_arquivo_sem_conectar(self):
        caminho = self.arquivo([['Ana', 'Silva', 'a@example.com', 'ana']])
        with patch('sys.argv', ['script', '--validar']), patch.object(kartado, 'selecionar_planilha', return_value=caminho) as seletor, patch.object(kartado, 'Cliente') as cliente:
            kartado.main()
            seletor.assert_called_once()
            cliente.assert_not_called()


if __name__ == '__main__':
    unittest.main()
