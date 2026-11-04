"""Cadastro em lote Kartado por Excel; nenhuma senha e persistida."""
import argparse
import base64
import getpass
import json
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urljoin, urlsplit
from urllib.request import Request
from uuid import UUID


BASE = 'https://eco.api.kartado.com.br/'


class ApiError(Exception):
    def __init__(self, status, metodo=None, endpoint=None):
        self.status = status
        self.metodo = metodo
        self.endpoint = endpoint
        contexto = f' em {metodo} {endpoint}' if endpoint else ''
        super().__init__(f'API retornou HTTP {status}{contexto}.')


class Cliente:
    def __init__(self, usuario, senha):
        self.authorization = 'Basic ' + base64.b64encode(f'{usuario}:{senha}'.encode()).decode()

    def requisitar(self, caminho, payload=None):
        url = urljoin(BASE, caminho)
        if urlsplit(url).netloc != urlsplit(BASE).netloc or urlsplit(url).scheme != 'https':
            raise ValueError('Endereco fora do ambiente Kartado configurado.')
        dados = None if payload is None else json.dumps(payload).encode('utf-8')
        req = Request(url, data=dados, headers={
            'Authorization': self.authorization,
            'Accept': 'application/vnd.api+json',
            'Content-Type': 'application/vnd.api+json',
        }, method='GET' if payload is None else 'POST')
        # Nao seguir redirecionamentos com credenciais para outro endereco.
        from urllib.request import build_opener, HTTPRedirectHandler
        class SemRedirecionamento(HTTPRedirectHandler):
            def redirect_request(self, req, fp, code, msg, headers, newurl):
                return None
        try:
            with build_opener(SemRedirecionamento).open(req, timeout=45) as resposta:
                return resposta.status, json.load(resposta)
        except HTTPError as erro:
            raise ApiError(erro.code, req.get_method(), urlsplit(url).path) from None

    def listar(self, recurso, filtros=None):
        caminho = recurso + '?' + urlencode({'page': 1, 'page_size': 25, **(filtros or {})})
        itens, paginas = [], set()
        while caminho:
            if caminho in paginas:
                raise ValueError('Paginacao repetida na API.')
            paginas.add(caminho)
            _, resposta = self.requisitar(caminho)
            if not isinstance(resposta.get('data'), list):
                raise ValueError('Resposta de listagem inesperada.')
            itens.extend(resposta['data'])
            caminho = (resposta.get('links') or {}).get('next')
            if isinstance(caminho, dict):
                caminho = caminho.get('href')
        return itens


def carregar_excel(caminho, aba=None):
    caminho = Path(caminho)
    if caminho.suffix.lower() != '.xlsx':
        raise ValueError('Selecione um arquivo Excel .xlsx. Para .xls, salve como .xlsx no Excel.')
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise ValueError('Instale a dependencia: python -m pip install -r Scripts/Kartado/requirements.txt') from None
    from zipfile import BadZipFile
    try:
        workbook = load_workbook(caminho, read_only=True, data_only=False)
    except (BadZipFile, ValueError):
        raise ValueError('Arquivo .xlsx invalido, corrompido ou protegido por senha.') from None
    try:
        if aba and aba not in workbook.sheetnames:
            raise ValueError(f'Aba nao encontrada: {aba}. Disponiveis: {", ".join(workbook.sheetnames)}')
        planilha = workbook[aba] if aba else workbook.active
        linhas = planilha.iter_rows()
        cabecalhos = [str(c.value or '').strip().casefold().replace('e-mail', 'email') for c in next(linhas, [])]
        campos = ['nome', 'sobrenome', 'email', 'login']
        if not set(campos).issubset(cabecalhos):
            raise ValueError('A primeira linha deve ter as colunas: nome, sobrenome, email, login.')
        if any(cabecalhos.count(campo) != 1 for campo in campos):
            raise ValueError('Cabecalhos obrigatorios repetidos na planilha.')
        indices = {campo: cabecalhos.index(campo) for campo in campos}
        usuarios, vistos = [], set()
        for numero, linha in enumerate(linhas, 2):
            celulas = {campo: linha[indice] for campo, indice in indices.items()}
            if all(c.value is None or str(c.value).strip() == '' for c in celulas.values()):
                continue
            if any(c.data_type in ('f', 'e') for c in celulas.values()):
                raise ValueError(f'Linha {numero}: use texto, sem formulas ou erros do Excel nos campos de cadastro.')
            valores = {campo: str(c.value if c.value is not None else '').strip() for campo, c in celulas.items()}
            if not all(valores.values()):
                raise ValueError(f'Linha {numero}: preencha todos os campos.')
            if len(valores['nome']) > 30 or len(valores['sobrenome']) > 150:
                raise ValueError(f'Linha {numero}: nome ou sobrenome excede o limite.')
            if len(valores['login']) > 150 or not re.fullmatch(r'[\w.@+-]+', valores['login']):
                raise ValueError(f'Linha {numero}: login invalido.')
            if len(valores['email']) > 254 or not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+', valores['email']):
                raise ValueError(f'Linha {numero}: email invalido.')
            chave = valores['login'].casefold()
            if chave in vistos:
                raise ValueError(f'Linha {numero}: login repetido no arquivo.')
            vistos.add(chave)
            usuarios.append({'linha': numero, **valores})
    finally:
        workbook.close()
    if not usuarios:
        raise ValueError('Planilha sem usuarios.')
    return usuarios


def selecionar(itens, titulo):
    if not itens:
        raise ValueError(f'Nenhuma opcao disponivel: {titulo}.')
    print('\n' + titulo)
    itens = sorted(itens, key=lambda item: item.get('attributes', {}).get('name', '').casefold())
    for indice, item in enumerate(itens, 1):
        print(f"{indice}. {item.get('attributes', {}).get('name', '(sem nome)')} [{item['id']}]")
    while True:
        escolha = input('Numero: ').strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(itens):
            return itens[int(escolha) - 1]
        print('Selecione um numero da lista.')


def payload_usuario(usuario, unidade, grupo, perfil):
    return {'data': {'type': 'User', 'attributes': {
        'firstName': usuario['nome'], 'lastName': usuario['sobrenome'],
        'email': usuario['email'], 'username': usuario['login'], 'createAccessRequest': True,
    }, 'relationships': {
        'company': {'data': {'type': 'Company', 'id': unidade}},
        'companyGroup': {'data': {'type': 'CompanyGroup', 'id': grupo}},
        'permissions': {'data': {'type': 'UserPermission', 'id': perfil}},
    }}}


def selecionar_planilha():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        raise ValueError('Seletor visual indisponivel neste Python. Instale o suporte Tcl/Tk ou use --arquivo caminho.xlsx.') from None
    raiz = None
    try:
        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes('-topmost', True)
        caminho = filedialog.askopenfilename(
            parent=raiz,
            title='Kartado - Selecionar planilha de usuarios',
            filetypes=[('Planilha Excel (*.xlsx)', '*.xlsx')],
        )
        return Path(caminho) if caminho else None
    except tk.TclError:
        raise ValueError('Nao foi possivel abrir o seletor visual. Use --arquivo caminho.xlsx.') from None
    finally:
        if raiz is not None:
            raiz.destroy()


def main():
    parser = argparse.ArgumentParser(description='Kartado: previa e criacao em lote por Excel .xlsx.')
    parser.add_argument('--arquivo', type=Path, help='Excel .xlsx; se omitido, abre uma janela para selecionar.')
    parser.add_argument('--aba', help='Nome da aba. Padrao: aba ativa do arquivo.')
    parser.add_argument('--validar', action='store_true', help='Valida somente o Excel, sem conexao.')
    modos = parser.add_mutually_exclusive_group()
    modos.add_argument('--somente-previa', action='store_true', help='Encerra apos a previa, sem oferecer criacao.')
    modos.add_argument('--executar', action='store_true', help=argparse.SUPPRESS)
    parser.add_argument('--grupo', help='UUID do CompanyGroup, se nao vier no retorno da unidade.')
    parser.add_argument('--saida', type=Path, default=Path(__file__).parent / 'Logs')
    args = parser.parse_args()
    if args.arquivo is None:
        args.arquivo = selecionar_planilha()
        if args.arquivo is None:
            print('Selecao cancelada. Nenhuma operacao foi realizada.')
            return
    print(f'Planilha selecionada: {args.arquivo}')
    usuarios = carregar_excel(args.arquivo, args.aba)
    print(f'{len(usuarios)} usuario(s) validado(s).')
    if args.validar:
        return
    usuario_api = input('Login da API: ').strip()
    senha = getpass.getpass('Senha da API (nao sera salva): ')
    cliente = Cliente(usuario_api, senha)
    unidade = selecionar(cliente.listar('Company/'), 'Unidade do lote')
    perfil = selecionar(cliente.listar('Permission/', {'company': unidade['id'], 'is_admin': 'false'}), 'Perfil do lote')
    grupo_unidade = (((unidade.get('relationships') or {}).get('companyGroup') or {}).get('data') or {}).get('id')
    if args.grupo and grupo_unidade and args.grupo != grupo_unidade:
        raise ValueError('Grupo informado difere do grupo retornado pela unidade.')
    grupo = grupo_unidade or args.grupo or input('UUID do CompanyGroup confirmado para esta unidade: ').strip()
    UUID(grupo)
    # Mesmo contexto utilizado pela listagem do portal; nao e uma busca global.
    existentes = cliente.listar('User/', {
        'company': unidade['id'],
        'sort': 'first_name',
        'only_company': unidade['id'],
    })
    logins = {str(item.get('attributes', {}).get('username', '')).casefold() for item in existentes}
    print('\nPrevia:')
    for item in usuarios:
        print(f"Linha {item['linha']}: {item['login']} | " + ('IGNORAR: login existente' if item['login'].casefold() in logins else 'CRIAR + SOLICITAR ACESSO'))
    print(f"Unidade: {unidade.get('attributes', {}).get('name')} | Perfil: {perfil.get('attributes', {}).get('name')} | Grupo: {grupo}")
    print('Verificacao de logins existentes limitada a unidade selecionada.')
    print('Usuarios encontrados nao serao alterados. Cadastros de outras unidades podem nao aparecer nesta consulta.')
    if args.somente_previa:
        print('Somente previa. Nenhum cadastro foi enviado.')
        return
    pendentes = sum(item['login'].casefold() not in logins for item in usuarios)
    if not pendentes:
        print('Nenhum usuario novo para criar nesta unidade.')
        return
    print(f'\nAprovacao: criar {pendentes} usuario(s) e solicitar acesso na unidade e perfil acima.')
    if input('Digite CRIAR para aplicar este lote: ').strip() != 'CRIAR':
        print('Execucao cancelada.')
        return
    args.saida.mkdir(parents=True, exist_ok=True)
    caminho_log = args.saida / ('resultado_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.jsonl')
    with caminho_log.open('x', encoding='utf-8') as log:
        for item in usuarios:
            resultado = {'linha': item['linha'], 'login': item['login'], 'unidade': unidade['id'], 'perfil': perfil['id']}
            parar = False
            if item['login'].casefold() in logins:
                resultado['resultado'] = 'ignorado_login_existente'
            else:
                try:
                    status, resposta = cliente.requisitar('User/', payload_usuario(item, unidade['id'], grupo, perfil['id']))
                    data = resposta.get('data') or {}
                    resultado.update(resultado='criado_verificar_aprovacao' if status == 201 else 'resposta_inesperada', http=status, id=data.get('id'), ativo=data.get('attributes', {}).get('active'))
                    logins.add(item['login'].casefold())
                    parar = status != 201
                except ApiError as erro:
                    resultado.update(resultado='erro_http', http=erro.status)
                    parar = erro.status in (401, 403, 429) or erro.status >= 500
                except (URLError, TimeoutError, OSError, ValueError):
                    resultado['resultado'] = 'resultado_incerto_conferir_no_portal_antes_de_repetir'
                    parar = True
            log.write(json.dumps(resultado, ensure_ascii=False) + '\n')
            log.flush()
            os.fsync(log.fileno())
            print(f"{item['login']}: {resultado['resultado']}")
            if parar:
                print('Lote interrompido. Confira o ultimo resultado antes de retomar; linhas seguintes nao foram enviadas.')
                break
    print(f'Relatorio: {caminho_log.resolve()}')


def executar():
    try:
        main()
    except KeyboardInterrupt:
        print('\nInterrompido. Se ocorreu durante um envio, confira o usuario no portal antes de repetir.')
        return 130
    except Exception as erro:
        # Sem variaveis locais, payloads, credenciais ou respostas completas.
        esperado = isinstance(erro, (ValueError, ApiError, URLError, OSError))
        mensagem = str(erro) if esperado else 'Falha inesperada. Consulte o diagnostico.'
        print(f'Falha ({type(erro).__name__}): {mensagem}', file=sys.stderr)
        pasta = Path(__file__).parent / 'Logs'
        diagnostico = {
            'data': datetime.now().isoformat(),
            'python': sys.executable,
            'versao_python': sys.version,
            'tipo': type(erro).__name__,
            'mensagem': mensagem,
            'pilha': [{'arquivo': Path(frame.filename).name, 'linha': frame.lineno, 'funcao': frame.name}
                      for frame in traceback.extract_tb(erro.__traceback__)],
        }
        try:
            pasta.mkdir(parents=True, exist_ok=True)
            destino = pasta / ('diagnostico_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.json')
            destino.write_text(json.dumps(diagnostico, ensure_ascii=False, indent=2), encoding='utf-8')
            print(f'Diagnostico salvo em: {destino.resolve()}', file=sys.stderr)
        except OSError:
            print('Nao foi possivel salvar o diagnostico na pasta Logs.', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(executar())
