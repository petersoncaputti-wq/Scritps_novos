import csv
import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from playwright.sync_api import Browser, BrowserContext, Page, Playwright, sync_playwright

from regras_v2 import (
    ACOES_USUARIO,
    CHAVES_PERMISSOES,
    EMAIL_REGEX,
    ROTULOS_ACAO_USUARIO,
    descricao_permissoes,
    determinar_acao_efetiva,
    extrair_ids_de_url,
    mascarar_email,
    normalizar_permissoes,
    normalizar_texto_chave,
    permissoes_iguais,
    titulo_igual,
)


URL_PORTAL = "https://pwdm.bentley.com"
URL_CONNECT = "https://connect.bentley.com"
ORIGEM_PWDM = "https://pwdm.bentley.com"
ARQUIVO_SESSAO = Path("session.json")
PASTA_LOGS = Path("Logs")
ARQUIVO_CACHE_EXPORT_CONNECT = PASTA_LOGS / "connect_projects_export_latest.csv"
PADRAO_ARVORE_PROJECTWISE = "projectwise_arvore_*.json"

GUID_REGEX = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)
SCORE_MINIMO_CORRESPONDENCIA_FUZZY = 80

ALIASES_PROJECTWISE_PWDM = {
    ("ecovias capixaba", "contorno de vitoria"): ["contorno", "vitoria", "dispositivo", "1"],
    ("ecovias capixaba", "contorno dis"): ["contorno", "vitoria", "dispositivo", "2"],
    ("ecovias capixaba", "contorno marg"): ["contorno", "vitoria", "marginais"],
    ("ecovias capixaba", "vias marginais"): ["marginais", "vitoria"],
}


def preparar_pasta_logs() -> None:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)


def nome_execucao() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def deve_usar_sessao_salva() -> bool:
    return os.getenv("USAR_SESSION_JSON", "").strip().lower() in {"1", "true", "s", "sim", "yes"}


def iniciar_navegador(playwright: Playwright) -> tuple[Browser, BrowserContext, Page]:
    print("[1/9] Abrindo navegador visual...")
    browser = playwright.chromium.launch(headless=False, slow_mo=80)

    opcoes_contexto: dict[str, Any] = {
        "viewport": {"width": 1366, "height": 768},
    }

    if deve_usar_sessao_salva():
        if ARQUIVO_SESSAO.exists():
            print(f"[INFO] Usando sessao salva para teste: {ARQUIVO_SESSAO.resolve()}")
            opcoes_contexto["storage_state"] = str(ARQUIVO_SESSAO)
        else:
            print(f"[AVISO] USAR_SESSION_JSON ativo, mas {ARQUIVO_SESSAO} nao existe.")

    context = browser.new_context(**opcoes_contexto)
    page = context.new_page()
    return browser, context, page


def abrir_portal(page: Page) -> None:
    print(f"[2/9] Acessando portal: {URL_PORTAL}")
    page.goto(URL_PORTAL, wait_until="domcontentloaded")
    print("\nSe necessario, faca login manualmente no navegador.")
    input("Quando estiver autenticado, pressione ENTER no terminal...")


def navegar_com_redirect(page: Page, url: str, timeout: int = 60000) -> None:
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=timeout)
    except Exception as erro:
        mensagem = str(erro)
        if "interrupted by another navigation" not in mensagem:
            raise
        print("[INFO] Navegacao interrompida por redirect automatico; seguindo apos carregar a pagina.")

    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
    except Exception:
        pass


def solicitar_email() -> str:
    while True:
        email = input("\nE-mail do usuario: ").strip().lower()
        if EMAIL_REGEX.match(email):
            return email
        print("[AVISO] Informe um e-mail valido.")


def solicitar_acao_usuario() -> str:
    print("\nO que deseja fazer com esse usuario?")
    print("01. Incluir nos projetos selecionados")
    print("02. Alterar permissoes onde ele ja for participante")
    print("03. Excluir dos projetos selecionados (indisponivel nesta versao)")

    while True:
        escolha = input("Acao desejada [1/2/3]: ").strip().lower()
        escolha = {
            "incluir": "1",
            "inclusao": "1",
            "adicionar": "1",
            "alterar": "2",
            "editar": "2",
            "permissionar": "2",
            "excluir": "3",
            "remover": "3",
        }.get(escolha, escolha)

        acao = ACOES_USUARIO.get(escolha)
        if not acao:
            print("[AVISO] Escolha 1, 2 ou 3.")
            continue
        if acao == "excluir":
            print("[AVISO] Exclusao ainda nao esta habilitada. Escolha incluir ou alterar.")
            continue

        print(f"[OK] Acao selecionada: {ROTULOS_ACAO_USUARIO[acao]}")
        return acao


def solicitar_titulo() -> str:
    titulo = input("\nTitulo/cargo do participante [Participante]: ").strip()
    return titulo or "Participante"


def solicitar_bool(rotulo: str, padrao: bool = False) -> bool:
    sufixo = "S/n" if padrao else "s/N"
    resposta = input(f"{rotulo} [{sufixo}]: ").strip().lower()

    if not resposta:
        return padrao

    return resposta in {"s", "sim", "y", "yes", "1", "true"}


def solicitar_permissoes() -> dict[str, bool]:
    print("\nPermissoes desejadas:")
    can_view = solicitar_bool("Exibir pacote", True)

    if not can_view:
        print("[AVISO] Sem 'Exibir pacote', as outras permissoes nao fazem sentido no PWDM.")

    permissoes = {
        "canView": can_view,
        "canReceive": solicitar_bool("Receber pacote"),
        "canIssue": solicitar_bool("Emitir pacote"),
        "canReviewAnswer": solicitar_bool("Aprovar resposta da RFI"),
        "isAdmin": solicitar_bool("Admin"),
    }
    return normalizar_permissoes(permissoes)


def valor_primeiro(dados: dict[str, Any], chaves: list[str]) -> str:
    for chave in chaves:
        valor = dados.get(chave)
        if valor:
            return str(valor)
    return ""


def montar_chaves_busca(*textos: Any) -> list[str]:
    chaves: list[str] = []
    for texto in textos:
        if isinstance(texto, list):
            chaves.extend(montar_chaves_busca(*texto))
            continue
        chave = normalizar_texto_chave(str(texto or ""))
        if chave:
            chaves.append(chave)
    return list(dict.fromkeys(chaves))


def guids_do_item(item: dict[str, Any]) -> list[str]:
    encontrados: list[str] = []
    for valor in item.values():
        if isinstance(valor, str):
            encontrados.extend(guid.lower() for guid in GUID_REGEX.findall(valor))
    return list(dict.fromkeys(encontrados))


def normalizar_projeto(item: dict[str, Any]) -> Optional[dict[str, Any]]:
    chaves_de_projeto = {
        "teamId",
        "TeamId",
        "ultimateRefId",
        "UltimateRefId",
        "status",
        "Status",
        "type",
        "Type",
        "solutionType",
        "SolutionType",
    }
    if not any(chave in item for chave in chaves_de_projeto):
        return None

    project_id = valor_primeiro(
        item,
        [
            "projectId",
            "ProjectId",
            "projectID",
            "id",
            "Id",
            "instanceId",
            "InstanceId",
        ],
    )
    ids = GUID_REGEX.findall(project_id)
    if not ids:
        return None

    nome = valor_primeiro(
        item,
        [
            "name",
            "Name",
            "displayName",
            "DisplayName",
            "projectName",
            "ProjectName",
            "label",
            "Label",
        ],
    )

    numero = valor_primeiro(item, ["number", "Number", "projectNumber", "ProjectNumber"])
    asset_name = valor_primeiro(item, ["assetName", "AssetName"])
    friendly_type = valor_primeiro(item, ["friendlyTypeName", "FriendlyTypeName"])
    location = valor_primeiro(item, ["location", "Location"])

    ativos = []
    for ativo in item.get("Assets") or item.get("assets") or []:
        if isinstance(ativo, dict):
            ativo_normalizado = normalizar_ativo(ativo)
            if ativo_normalizado:
                ativos.append(ativo_normalizado)

    ativo_do_projeto = normalizar_ativo_do_projeto(item)
    if ativo_do_projeto and not any(ativo["assetId"] == ativo_do_projeto["assetId"] for ativo in ativos):
        ativos.append(ativo_do_projeto)

    return {
        "projectId": ids[0],
        "connectSpaceId": ids[0],
        "nome": nome or ids[0],
        "numero": numero,
        "assetName": asset_name,
        "friendlyTypeName": friendly_type,
        "location": location,
        "ativos": ativos,
        "guids": guids_do_item(item),
        "nomeChave": normalizar_texto_chave(nome or ids[0]),
        "chavesBusca": montar_chaves_busca(nome, numero, asset_name, friendly_type, location, ids[0]),
        "textoBusca": " | ".join(valor for valor in [numero, nome, asset_name, friendly_type, location] if valor),
    }


def normalizar_ativo(item: dict[str, Any]) -> Optional[dict[str, str]]:
    ativo_id = valor_primeiro(
        item,
        [
            "instanceId",
            "InstanceId",
            "assetId",
            "AssetId",
            "id",
            "Id",
        ],
    )

    nome = valor_primeiro(
        item,
        [
            "name",
            "Name",
            "displayName",
            "DisplayName",
            "assetName",
            "AssetName",
            "label",
            "Label",
        ],
    )
    numero = valor_primeiro(item, ["number", "Number", "assetNumber", "AssetNumber"])
    tipo = valor_primeiro(item, ["friendlyTypeName", "FriendlyTypeName", "assetType", "AssetType", "type", "Type"])
    local = valor_primeiro(item, ["location", "Location"])

    ids = GUID_REGEX.findall(ativo_id)
    if ids:
        chave = ids[0]
    elif nome:
        chave = f"{nome}|{tipo}|{local}".lower()
    else:
        return None

    return {
        "assetId": chave,
        "nome": nome or ids[0],
        "numero": numero,
        "tipo": tipo,
        "local": local,
    }


def normalizar_ativo_do_projeto(item: dict[str, Any]) -> Optional[dict[str, str]]:
    nome = valor_primeiro(item, ["assetName", "AssetName"])
    if not nome:
        return None

    return normalizar_ativo(
        {
            "Name": nome,
            "Number": valor_primeiro(item, ["assetNumber", "AssetNumber"]),
            "FriendlyTypeName": valor_primeiro(item, ["friendlyTypeName", "FriendlyTypeName"]),
            "Location": valor_primeiro(item, ["location", "Location"]),
        }
    )


def coletar_projetos_recursivo(valor: Any, projetos: list[dict[str, Any]]) -> None:
    if isinstance(valor, dict):
        projeto = normalizar_projeto(valor)
        if projeto:
            projetos.append(projeto)

        for item in valor.values():
            coletar_projetos_recursivo(item, projetos)

    elif isinstance(valor, list):
        for item in valor:
            coletar_projetos_recursivo(item, projetos)


def deduplicar_projetos(projetos: list[dict[str, Any]]) -> list[dict[str, Any]]:
    deduplicados: dict[str, dict[str, Any]] = {}
    for projeto in projetos:
        project_id = projeto["projectId"]
        atual = deduplicados.get(project_id)
        if not atual:
            deduplicados[project_id] = projeto
            continue

        ativos_por_id = {ativo["assetId"]: ativo for ativo in atual.get("ativos") or []}
        for ativo in projeto.get("ativos") or []:
            ativos_por_id[ativo["assetId"]] = ativo
        atual["ativos"] = list(ativos_por_id.values())

        if len(projeto.get("nome", "")) > len(atual.get("nome", "")):
            atual["nome"] = projeto["nome"]

        chaves = (atual.get("chavesBusca") or []) + (projeto.get("chavesBusca") or [])
        atual["chavesBusca"] = list(dict.fromkeys(chaves))
        textos = [atual.get("textoBusca", ""), projeto.get("textoBusca", "")]
        atual["textoBusca"] = " | ".join(item for item in textos if item)
    return list(deduplicados.values())


def normalizar_projeto_de_link_connect(href: str, texto: str, campos: list[str]) -> Optional[dict[str, Any]]:
    try:
        connect_space_id, project_id = extrair_ids_de_url(href)
    except Exception:
        return None

    campos_limpos = [campo.strip() for campo in campos if campo and campo.strip()]
    if not campos_limpos and texto:
        campos_limpos = [linha.strip() for linha in texto.splitlines() if linha.strip()]

    nome = campos_limpos[0] if campos_limpos else project_id
    numero = campos_limpos[1] if len(campos_limpos) > 1 else ""
    asset_name = campos_limpos[2] if len(campos_limpos) > 2 else ""
    friendly_type = campos_limpos[3] if len(campos_limpos) > 3 else ""
    location = campos_limpos[4] if len(campos_limpos) > 4 else ""
    texto_busca = " | ".join(campos_limpos)

    return {
        "projectId": project_id,
        "connectSpaceId": connect_space_id,
        "nome": nome,
        "numero": numero,
        "assetName": asset_name,
        "friendlyTypeName": friendly_type,
        "location": location,
        "guids": [project_id.lower()],
        "nomeChave": normalizar_texto_chave(nome),
        "chavesBusca": montar_chaves_busca(campos_limpos, project_id),
        "textoBusca": texto_busca or nome,
    }


def normalizar_projeto_de_csv_connect(linha: dict[str, str]) -> Optional[dict[str, Any]]:
    link = linha.get("Link") or linha.get("link") or ""
    try:
        connect_space_id, project_id = extrair_ids_de_url(link)
    except Exception:
        return None

    numero_csv = linha.get("Number") or ""
    nome_csv = linha.get("Name") or ""
    # No export do Connect, "Number" carrega o rotulo mais proximo da pasta ProjectWise.
    nome = numero_csv or nome_csv or project_id
    numero = nome_csv
    asset_name = linha.get("Asset Name") or ""
    asset_id = linha.get("Asset Id") or ""
    friendly_type = linha.get("Type") or ""
    location = linha.get("Location") or ""
    status = linha.get("Status") or ""
    campos = [numero_csv, nome_csv, asset_name, friendly_type, location, status]

    ativos = []
    if asset_name or asset_id:
        ativo = normalizar_ativo(
            {
                "Id": asset_id,
                "Name": asset_name,
                "FriendlyTypeName": friendly_type,
                "Location": location,
            }
        )
        if ativo:
            ativos.append(ativo)

    return {
        "projectId": project_id,
        "connectSpaceId": connect_space_id,
        "nome": nome,
        "numero": numero,
        "assetName": asset_name,
        "friendlyTypeName": friendly_type,
        "location": location,
        "status": status,
        "ativos": ativos,
        "guids": [project_id.lower()],
        "nomeChave": normalizar_texto_chave(nome),
        "chavesBusca": montar_chaves_busca(campos, project_id),
        "textoBusca": " | ".join(campo for campo in campos if campo),
    }


def buscar_projetos_em_endpoint(page: Page, endpoint: str, rotulo: str) -> list[dict[str, Any]]:
    resultado = page.evaluate(
        """
        async ({ endpoint }) => {
            const response = await fetch(endpoint, {
                method: "GET",
                credentials: "include",
                headers: {
                    "Accept": "application/json",
                    "X-Requested-With": "XMLHttpRequest"
                }
            });

            const text = await response.text();
            return {
                ok: response.ok,
                status: response.status,
                statusText: response.statusText,
                contentType: response.headers.get("content-type") || "",
                text
            };
        }
        """,
        {"endpoint": endpoint},
    )

    if not resultado["ok"] or "application/json" not in resultado["contentType"].lower():
        detalhe = str(resultado.get("text") or "").strip().replace("\n", " ")[:250]
        detalhe_msg = f" Detalhe: {detalhe}" if detalhe else ""
        print(f"[AVISO] {rotulo}: HTTP {resultado['status']} {resultado['statusText']}.{detalhe_msg}")
        return []

    dados = json.loads(resultado["text"])
    projetos: list[dict[str, Any]] = []
    coletar_projetos_recursivo(dados, projetos)
    projetos = deduplicar_projetos(projetos)
    print(f"[INFO] {rotulo}: {len(projetos)} projeto(s).")
    return projetos


def buscar_projetos_por_api(page: Page) -> list[dict[str, Any]]:
    print("\n[3/9] Buscando catalogo de projetos disponiveis no Connect/PWDM...")
    navegar_com_redirect(page, f"{URL_CONNECT}/SelectProject/Index")

    projetos_export = buscar_projetos_por_export_connect(page)
    if projetos_export:
        return projetos_export

    endpoints = [
        ("/SelectProject/GetMyProjects", "Meus projetos"),
        ("/SelectProject/GetFavorites", "Favoritos"),
        ("/SelectProject/GetMruList", "Recentes"),
    ]

    projetos: list[dict[str, Any]] = []
    for endpoint, rotulo in endpoints:
        projetos.extend(buscar_projetos_em_endpoint(page, endpoint, rotulo))

    return deduplicar_projetos(projetos)


def decodificar_csv_connect(conteudo: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            texto = conteudo.decode(encoding)
        except UnicodeDecodeError:
            continue
        if "Number,Name,Asset Name" in texto:
            return texto
    return conteudo.decode("utf-8", errors="replace")


def buscar_projetos_por_export_connect(page: Page) -> list[dict[str, Any]]:
    resposta = page.context.request.get(f"{URL_CONNECT}/Admin/GetAllProjectInfo", timeout=60000)
    content_type = resposta.headers.get("content-type", "")
    if not resposta.ok or "csv" not in content_type.lower():
        detalhe = resposta.text()[:160].strip().replace("\n", " ") if resposta.body() else ""
        retry_after = resposta.headers.get("retry-after")
        retry_msg = f" Tente novamente em {retry_after}s." if retry_after else ""
        print(f"[AVISO] Export geral de projetos: HTTP {resposta.status}.{retry_msg} {detalhe}")
        if ARQUIVO_CACHE_EXPORT_CONNECT.exists():
            print(f"[INFO] Usando cache do export geral: {ARQUIVO_CACHE_EXPORT_CONNECT.resolve()}")
            texto_csv = ARQUIVO_CACHE_EXPORT_CONNECT.read_text(encoding="utf-8-sig")
            return projetos_do_csv_connect(texto_csv)
        return []

    conteudo = resposta.body()
    texto_csv = decodificar_csv_connect(conteudo)
    ARQUIVO_CACHE_EXPORT_CONNECT.write_text(texto_csv, encoding="utf-8-sig")
    return projetos_do_csv_connect(texto_csv)


def projetos_do_csv_connect(texto_csv: str) -> list[dict[str, Any]]:
    projetos: list[dict[str, Any]] = []
    for linha in csv.DictReader(io.StringIO(texto_csv)):
        projeto = normalizar_projeto_de_csv_connect(linha)
        if projeto:
            projetos.append(projeto)

    projetos = deduplicar_projetos(projetos)
    print(f"[INFO] Export geral de projetos: {len(projetos)} projeto(s).")
    return projetos


def abrir_aba_busca_connect(page: Page) -> None:
    if not page.url.lower().startswith(f"{URL_CONNECT.lower()}/selectproject/index"):
        navegar_com_redirect(page, f"{URL_CONNECT}/SelectProject/Index")

    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    if page.locator("#projectSearchField").count():
        return

    busca = page.locator("a[href*='SEARCH']").first
    if not busca.count():
        navegar_com_redirect(page, f"{URL_CONNECT}/SelectProject/Index")
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        busca = page.locator("a[href*='SEARCH']").first

    busca.click(timeout=10000)
    page.wait_for_selector("#projectSearchField", state="visible", timeout=20000)


def buscar_projetos_por_termo(page: Page, termo: str) -> list[dict[str, Any]]:
    termo = termo.strip()
    if not termo:
        return []

    try:
        abrir_aba_busca_connect(page)
        if page.locator("#includeInactiveProject").count():
            checkbox = page.locator("#includeInactiveProject").first
            if not checkbox.is_checked():
                checkbox.check()
        page.locator("#projectSearchField").fill(termo)
        page.keyboard.press("Enter")
        page.wait_for_timeout(4500)
    except Exception as erro:
        print(f"[AVISO] Nao consegui usar a busca visual do Connect para '{termo}': {erro}")
        return []

    dados = page.evaluate(
        """
        () => Array.from(document.querySelectorAll("a[href*='projectId=']")).map((link) => ({
            href: link.href,
            text: (link.innerText || link.textContent || "").trim(),
            campos: Array.from(link.querySelectorAll(".btly-list-cell")).map((cell) => (
                cell.getAttribute("title") || cell.innerText || cell.textContent || ""
            ).trim()).filter(Boolean)
        }))
        """
    )

    projetos: list[dict[str, Any]] = []
    for item in dados:
        projeto = normalizar_projeto_de_link_connect(
            str(item.get("href") or ""),
            str(item.get("text") or ""),
            item.get("campos") or [],
        )
        if projeto:
            projetos.append(projeto)

    projetos = deduplicar_projetos(projetos)
    print(f"[INFO] Busca '{termo}': {len(projetos)} projeto(s).")
    return projetos


def buscar_projetos_por_links_visiveis(page: Page) -> list[dict[str, Any]]:
    print("[INFO] Tentando identificar projetos pelos links visiveis da pagina...")
    itens = page.evaluate(
        """
        () => Array.from(document.querySelectorAll("a[href*='projectId=']")).map((link) => ({
            href: link.href,
            text: (link.innerText || link.textContent || "").trim(),
            campos: Array.from(link.querySelectorAll(".btly-list-cell")).map((cell) => (
                cell.getAttribute("title") || cell.innerText || cell.textContent || ""
            ).trim()).filter(Boolean)
        }))
        """
    )

    projetos: list[dict[str, Any]] = []
    for item in itens:
        projeto = normalizar_projeto_de_link_connect(
            str(item.get("href") or ""),
            str(item.get("text") or ""),
            item.get("campos") or [],
        )
        if projeto:
            projetos.append(projeto)

    return deduplicar_projetos(projetos)


def buscar_concessoes_disponiveis(page: Page) -> list[dict[str, Any]]:
    projetos = buscar_projetos_por_api(page)
    if not projetos:
        projetos = buscar_projetos_por_links_visiveis(page)

    if not projetos:
        raise RuntimeError("Nao consegui listar concessoes/projetos disponiveis.")

    projetos.sort(key=lambda item: item.get("nome", "").lower())
    print(f"[OK] {len(projetos)} projeto(s) encontrado(s) no Connect/PWDM.")
    return projetos


def carregar_arvore_projectwise() -> dict[str, Any]:
    arquivos = sorted(PASTA_LOGS.glob(PADRAO_ARVORE_PROJECTWISE), key=lambda item: item.stat().st_mtime, reverse=True)
    if not arquivos:
        raise RuntimeError(
            "Nao encontrei a arvore ProjectWise em Logs. "
            "Execute primeiro: python inspecionar_arvore_projectwise.py"
        )

    arquivo = arquivos[0]
    with arquivo.open("r", encoding="utf-8-sig") as entrada:
        dados = json.load(entrada)

    concessoes = dados.get("concessoes") or []
    if not concessoes:
        raise RuntimeError(f"O arquivo {arquivo.name} nao contem concessoes ProjectWise.")

    print(f"\n[INFO] Arvore ProjectWise carregada: {arquivo.resolve()}")
    print(f"[INFO] Base: {dados.get('caminhoBase', 'Engenharia')} | Concessoes: {len(concessoes)}")
    return dados


def selecionar_concessao_projectwise(arvore: dict[str, Any]) -> dict[str, Any]:
    concessoes = [
        item for item in (arvore.get("concessoes") or [])
        if (item.get("concessao") or {}).get("nome") and item.get("pastaProjetosEncontrada")
    ]

    if not concessoes:
        raise RuntimeError("Nenhuma concessao ProjectWise com pasta de projetos foi encontrada.")

    print("\n[5/9] Concessoes ProjectWise disponiveis:")
    for indice, item in enumerate(concessoes, start=1):
        concessao = item.get("concessao") or {}
        projetos = item.get("projetos") or []
        print(f"{indice:02d}. {concessao.get('nome')} ({len(projetos)} projeto(s))")

    while True:
        selecao = input("\nNumero da concessao desejada: ").strip()
        if selecao.isdigit():
            indice = int(selecao)
            if 1 <= indice <= len(concessoes):
                selecionada = concessoes[indice - 1]
                print(f"[OK] Concessao selecionada: {selecionada['concessao']['nome']}")
                return selecionada

        print("[AVISO] Selecione um numero valido da lista.")


def selecionar_projetos_projectwise(concessao: dict[str, Any]) -> list[dict[str, Any]]:
    nome_concessao = (concessao.get("concessao") or {}).get("nome") or ""
    projetos = [
        {**projeto, "concessaoProjectWise": nome_concessao}
        for projeto in (concessao.get("projetos") or [])
    ]
    if not projetos:
        raise RuntimeError("A concessao selecionada nao contem projetos.")

    print("\n[6/9] Projetos ProjectWise da concessao selecionada:")
    for indice, projeto in enumerate(projetos, start=1):
        print(f"{indice:03d}. {projeto.get('nome')} | PW ID: {projeto.get('id')} | GUID: {projeto.get('guid')}")

    print("\nDigite os numeros separados por virgula, ou 'todos'.")
    selecao = input("Selecao: ").strip().lower()

    if selecao in {"todos", "tudo", "all"}:
        return projetos

    escolhidos: list[dict[str, Any]] = []
    for parte in selecao.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if not parte.isdigit():
            raise ValueError(f"Selecao invalida: {parte}")

        indice = int(parte)
        if indice < 1 or indice > len(projetos):
            raise ValueError(f"Indice fora da lista: {indice}")
        escolhidos.append(projetos[indice - 1])

    if not escolhidos:
        raise ValueError("Nenhum projeto ProjectWise foi selecionado.")

    return escolhidos


def indexar_projetos_pwdm(projetos_pwdm: list[dict[str, Any]]) -> tuple[dict[str, dict[str, Any]], dict[str, list[dict[str, Any]]]]:
    por_guid: dict[str, dict[str, Any]] = {}
    por_nome: dict[str, list[dict[str, Any]]] = {}

    for projeto in projetos_pwdm:
        for guid in projeto.get("guids") or []:
            por_guid[guid.lower()] = projeto
        por_guid[projeto["projectId"].lower()] = projeto

        nome_chave = projeto.get("nomeChave") or normalizar_texto_chave(projeto.get("nome", ""))
        chaves = projeto.get("chavesBusca") or [nome_chave]
        for chave in chaves:
            if chave:
                por_nome.setdefault(chave, []).append(projeto)

    return por_guid, por_nome


def tokens_de_busca(texto: str) -> set[str]:
    tokens = set(normalizar_texto_chave(texto).split())
    expansoes = {
        "duplica": "duplicacao",
        "dupli": "duplicacao",
        "duplic": "duplicacao",
        "duplicacao": "duplicacao",
        "contor": "contorno",
        "dis": "dispositivo",
        "disp": "dispositivo",
        "marg": "marginais",
        "marginal": "marginais",
    }
    for token in list(tokens):
        if token in expansoes:
            tokens.add(expansoes[token])

    composicoes = {
        ("duplica", "o"): "duplicacao",
        ("vit", "ria"): "vitoria",
        ("fund", "o"): "fundao",
        ("ibira", "u"): "ibiracu",
        ("can", "rio"): "canario",
        ("timbu",): "timbui",
    }
    for partes, token_composto in composicoes.items():
        if all(parte in tokens for parte in partes):
            tokens.add(token_composto)

    return {token for token in tokens if token not in {"de", "da", "do", "e"}}


def pontuar_correspondencia(projeto_pw: dict[str, Any], projeto_pwdm: dict[str, Any]) -> int:
    nome_pw = str(projeto_pw.get("nome") or "")
    texto_pwdm = " ".join(
        str(projeto_pwdm.get(campo) or "")
        for campo in ["numero", "nome", "assetName", "friendlyTypeName", "location", "textoBusca"]
    )
    tokens_pw = tokens_de_busca(nome_pw)
    tokens_pwdm = tokens_de_busca(texto_pwdm)
    tokens_nome_pwdm = tokens_de_busca(str(projeto_pwdm.get("nome") or ""))
    if not tokens_pw or not tokens_pwdm:
        return 0

    score = 0
    comuns = tokens_pw & tokens_pwdm
    score += len(comuns) * 12

    if tokens_pw.issubset(tokens_pwdm):
        score += 45
        score -= max(0, len(tokens_nome_pwdm - tokens_pw)) * 8

    if normalizar_texto_chave(nome_pw) == normalizar_texto_chave(str(projeto_pwdm.get("nome") or "")):
        score += 80

    concessao_pw = normalizar_texto_chave(str(projeto_pw.get("concessaoProjectWise") or ""))
    asset_pwdm = normalizar_texto_chave(str(projeto_pwdm.get("assetName") or ""))
    if concessao_pw and asset_pwdm == concessao_pw:
        score += 35

    if "antt" in tokens_pw and "antt" in tokens_pwdm:
        score += 25
    if "antt" not in tokens_pw and "antt" in tokens_pwdm:
        score -= 20

    if "obras" in tokens_pwdm and "antt" not in tokens_pw:
        score += 10

    return score


def buscar_correspondencia_fuzzy_no_catalogo(
    projeto_pw: dict[str, Any],
    projetos_pwdm: list[dict[str, Any]],
) -> Optional[dict[str, Any]]:
    concessao_pw = normalizar_texto_chave(str(projeto_pw.get("concessaoProjectWise") or ""))
    candidatos = projetos_pwdm
    if concessao_pw:
        filtrados = [
            projeto for projeto in projetos_pwdm
            if normalizar_texto_chave(str(projeto.get("assetName") or "")) == concessao_pw
        ]
        if filtrados:
            candidatos = filtrados

    pontuados = [
        (pontuar_correspondencia(projeto_pw, projeto), projeto)
        for projeto in candidatos
    ]
    pontuados = sorted(
        [(score, projeto) for score, projeto in pontuados if score >= SCORE_MINIMO_CORRESPONDENCIA_FUZZY],
        key=lambda item: item[0],
        reverse=True,
    )
    if not pontuados:
        return None

    melhor_score, melhor = pontuados[0]
    segundo_score = pontuados[1][0] if len(pontuados) > 1 else 0
    if segundo_score and melhor_score - segundo_score < 15:
        nomes = ", ".join(projeto.get("nome", "") for _, projeto in pontuados[:4])
        print(f"[AVISO] Match ambiguo para {projeto_pw.get('nome')}: {nomes}")
        return None

    projeto = dict(melhor)
    projeto["origemProjectWise"] = projeto_pw
    projeto["criterioCruzamento"] = f"catalogo Connect por similaridade ({melhor_score})"
    return projeto


def buscar_correspondencia_por_alias(
    projeto_pw: dict[str, Any],
    projetos_pwdm: list[dict[str, Any]],
) -> Optional[dict[str, Any]]:
    chave_alias = (
        normalizar_texto_chave(str(projeto_pw.get("concessaoProjectWise") or "")),
        normalizar_texto_chave(str(projeto_pw.get("nome") or "")),
    )
    tokens_alvo = ALIASES_PROJECTWISE_PWDM.get(chave_alias)
    if not tokens_alvo:
        return None

    concessao_pw = chave_alias[0]
    candidatos: list[dict[str, Any]] = []
    for projeto in projetos_pwdm:
        asset_pwdm = normalizar_texto_chave(str(projeto.get("assetName") or ""))
        if concessao_pw and asset_pwdm != concessao_pw:
            continue

        texto = " ".join(
            str(projeto.get(campo) or "")
            for campo in ["numero", "nome", "assetName", "friendlyTypeName", "location", "textoBusca"]
        )
        tokens_pwdm = tokens_de_busca(texto)
        if set(tokens_alvo).issubset(tokens_pwdm):
            candidatos.append(projeto)

    if chave_alias == ("ecovias capixaba", "vias marginais"):
        sem_contorno = [
            projeto for projeto in candidatos
            if "contorno" not in tokens_de_busca(
                " ".join(str(projeto.get(campo) or "") for campo in ["numero", "nome", "textoBusca"])
            )
        ]
        if sem_contorno:
            candidatos = sem_contorno

    if len(candidatos) != 1:
        if candidatos:
            nomes = ", ".join(projeto.get("nome", "") for projeto in candidatos)
            print(f"[AVISO] Alias ambiguo para {projeto_pw.get('nome')}: {nomes}")
        return None

    projeto = dict(candidatos[0])
    projeto["origemProjectWise"] = projeto_pw
    projeto["criterioCruzamento"] = "alias controlado ProjectWise/PWDM"
    return projeto


def cruzar_projeto_projectwise_com_pwdm(
    projeto_pw: dict[str, Any],
    projetos_pwdm: list[dict[str, Any]],
    por_guid: dict[str, dict[str, Any]],
    por_nome: dict[str, list[dict[str, Any]]],
) -> Optional[dict[str, Any]]:
    guid_pw = str(projeto_pw.get("guid") or "").lower()
    if guid_pw and guid_pw in por_guid:
        projeto = dict(por_guid[guid_pw])
        projeto["origemProjectWise"] = projeto_pw
        projeto["criterioCruzamento"] = "guid"
        return projeto

    nome_chave = normalizar_texto_chave(str(projeto_pw.get("nome") or ""))
    candidatos = por_nome.get(nome_chave) or []
    if len(candidatos) == 1:
        projeto = dict(candidatos[0])
        projeto["origemProjectWise"] = projeto_pw
        projeto["criterioCruzamento"] = "nome"
        return projeto

    por_alias = buscar_correspondencia_por_alias(projeto_pw, projetos_pwdm)
    if por_alias:
        return por_alias

    return buscar_correspondencia_fuzzy_no_catalogo(projeto_pw, projetos_pwdm)


def escolher_melhor_resultado_busca(
    projeto_pw: dict[str, Any],
    candidatos: list[dict[str, Any]],
) -> Optional[dict[str, Any]]:
    if not candidatos:
        return None

    nome_chave_pw = normalizar_texto_chave(str(projeto_pw.get("nome") or ""))
    exatos = [
        projeto for projeto in candidatos
        if nome_chave_pw in (projeto.get("chavesBusca") or [])
        or (projeto.get("nomeChave") or normalizar_texto_chave(projeto.get("nome", ""))) == nome_chave_pw
    ]
    if len(exatos) == 1:
        projeto = dict(exatos[0])
        projeto["origemProjectWise"] = projeto_pw
        projeto["criterioCruzamento"] = "busca Connect por nome exato"
        return projeto

    if len(candidatos) == 1:
        projeto = dict(candidatos[0])
        projeto["origemProjectWise"] = projeto_pw
        projeto["criterioCruzamento"] = "busca Connect resultado unico"
        return projeto

    return None


def buscar_correspondencia_no_connect(
    page: Page,
    projeto_pw: dict[str, Any],
) -> Optional[dict[str, Any]]:
    termos = [
        str(projeto_pw.get("nome") or "").strip(),
        str(projeto_pw.get("guid") or "").strip(),
    ]

    vistos = set()
    for termo in termos:
        if not termo or termo.lower() in vistos:
            continue
        vistos.add(termo.lower())

        candidatos = buscar_projetos_por_termo(page, termo)
        escolhido = escolher_melhor_resultado_busca(projeto_pw, candidatos)
        if escolhido:
            return escolhido

        if len(candidatos) > 1:
            nomes = ", ".join(projeto.get("nome", "") for projeto in candidatos[:5])
            print(f"[AVISO] Busca '{termo}' retornou {len(candidatos)} candidatos; sem match unico. Primeiros: {nomes}")

    return None


def complementar_correspondencias_manuais(nao_encontrados: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not nao_encontrados:
        return []

    resposta = input(
        "\nDeseja informar manualmente a URL PWDM para algum projeto sem correspondencia? [s/N]: "
    ).strip().lower()
    if resposta not in {"s", "sim", "y", "yes"}:
        return []

    manuais: list[dict[str, Any]] = []
    for projeto_pw in nao_encontrados:
        print(f"\nProjeto ProjectWise sem correspondencia: {projeto_pw.get('nome')}")
        print(f"GUID PW: {projeto_pw.get('guid')} | ID PW: {projeto_pw.get('id')}")
        url = input("Cole a URL PWDM de Participantes, ou ENTER para pular: ").strip()
        if not url:
            continue

        try:
            connect_space_id, project_id = extrair_ids_de_url(url)
        except Exception as erro:
            print(f"[AVISO] URL ignorada: {erro}")
            continue

        manuais.append(
            {
                "projectId": project_id,
                "connectSpaceId": connect_space_id,
                "nome": projeto_pw.get("nome") or project_id,
                "numero": "",
                "guids": [project_id.lower()],
                "nomeChave": normalizar_texto_chave(projeto_pw.get("nome") or project_id),
                "origemProjectWise": projeto_pw,
                "criterioCruzamento": "url manual",
            }
        )

    return manuais


def cruzar_projetos_projectwise_com_pwdm(
    page: Page,
    projetos_pw: list[dict[str, Any]],
    projetos_pwdm: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    por_guid, por_nome = indexar_projetos_pwdm(projetos_pwdm)
    cruzados: list[dict[str, Any]] = []
    nao_encontrados: list[dict[str, Any]] = []

    print("\n[INFO] Cruzando projetos ProjectWise com projetos PWDM...")
    for projeto_pw in projetos_pw:
        projeto_pwdm = cruzar_projeto_projectwise_com_pwdm(projeto_pw, projetos_pwdm, por_guid, por_nome)
        if projeto_pwdm:
            cruzados.append(projeto_pwdm)
            print(
                f"[OK] {projeto_pw.get('nome')} -> {projeto_pwdm.get('nome')} "
                f"({projeto_pwdm.get('criterioCruzamento')})"
            )
        else:
            nao_encontrados.append(projeto_pw)
            print(f"[AVISO] Sem correspondencia PWDM: {projeto_pw.get('nome')} | GUID: {projeto_pw.get('guid')}")

    if nao_encontrados:
        print("\n[INFO] Pesquisando no Connect/PWDM os projetos sem correspondencia inicial...")
        pendentes: list[dict[str, Any]] = []
        for projeto_pw in nao_encontrados:
            projeto_encontrado = buscar_correspondencia_no_connect(page, projeto_pw)
            if projeto_encontrado:
                cruzados.append(projeto_encontrado)
                print(
                    f"[OK] {projeto_pw.get('nome')} -> {projeto_encontrado.get('nome')} "
                    f"({projeto_encontrado.get('criterioCruzamento')})"
                )
            else:
                pendentes.append(projeto_pw)
        nao_encontrados = pendentes

    if nao_encontrados:
        arquivo = PASTA_LOGS / f"pwdm_projetos_sem_correspondencia_{nome_execucao()}.json"
        with arquivo.open("w", encoding="utf-8") as saida:
            json.dump(nao_encontrados, saida, indent=2, ensure_ascii=False)
        print(f"[AVISO] {len(nao_encontrados)} projeto(s) sem correspondencia. Log: {arquivo.resolve()}")
        cruzados.extend(complementar_correspondencias_manuais(nao_encontrados))

    if not cruzados:
        raise RuntimeError(
            "Nenhum projeto ProjectWise selecionado foi encontrado no catalogo PWDM/Connect. "
            "Nenhuma alteracao sera preparada."
        )

    print(f"[OK] {len(cruzados)} projeto(s) cruzado(s) com PWDM.")
    return cruzados


def json_contem_email(valor: Any, email: str) -> Optional[dict[str, Any]]:
    if isinstance(valor, dict):
        emails = [
            valor.get("email"),
            valor.get("Email"),
            valor.get("userEmail"),
            valor.get("UserEmail"),
            valor.get("mail"),
            valor.get("Mail"),
        ]
        if any(str(item).lower() == email for item in emails if item):
            return valor

        for item in valor.values():
            encontrado = json_contem_email(item, email)
            if encontrado:
                return encontrado

    elif isinstance(valor, list):
        for item in valor:
            encontrado = json_contem_email(item, email)
            if encontrado:
                return encontrado

    return None


def buscar_usuario_connect(page: Page, project_id: str, email: str) -> Optional[dict[str, Any]]:
    resposta = page.context.request.get(
        f"{URL_CONNECT}/Project/GetAllProjectUsers?projectid={project_id}",
        headers={
            "Accept": "application/json",
            "X-Requested-With": "XMLHttpRequest",
        },
        timeout=30000,
    )

    if not resposta.ok or "application/json" not in resposta.headers.get("content-type", "").lower():
        return None

    return json_contem_email(json.loads(resposta.text()), email)


def validar_cadastro_usuario_pwdm(
    page: Page,
    email: str,
    projetos: list[dict[str, Any]],
) -> dict[str, Any]:
    print(f"\n[4/9] Validando cadastro do usuario no Connect/PWDM: {email}")

    for indice, projeto in enumerate(projetos, start=1):
        print(f"- Verificando {indice}/{len(projetos)}: {projeto['nome']}")
        encontrado = buscar_usuario_connect(page, projeto["projectId"], email)
        if encontrado:
            nome = (
                encontrado.get("name")
                or encontrado.get("Name")
                or encontrado.get("displayName")
                or encontrado.get("DisplayName")
                or email
            )
            print(f"[OK] Usuario identificado: {nome} <{email}>")
            return {
                "email": email,
                "nome": nome,
                "origemProjectId": projeto["projectId"],
                "origemProjetoNome": projeto.get("nome"),
                "dados": encontrado,
                "podeUsarNoPwdm": True,
            }

    raise RuntimeError(
        "Nao encontrei esse e-mail nos usuarios dos projetos acessiveis no Connect/PWDM. "
        "Confirme se o usuario ja existe/cadastrado no ProjectWise/Connect e se esta disponivel "
        "para uso no PWDM antes de preparar qualquer alteracao."
    )


def selecionar_concessoes(projetos: list[dict[str, Any]]) -> list[dict[str, Any]]:
    print("\n[6/9] Selecione os projetos para incluir/permissionar o usuario:")
    for indice, projeto in enumerate(projetos, start=1):
        numero = f" - {projeto['numero']}" if projeto.get("numero") else ""
        print(f"{indice:02d}. {projeto['nome']}{numero}")

    print("\nDigite os numeros separados por virgula, ou 'todos'.")
    selecao = input("Selecao: ").strip().lower()

    if selecao in {"todos", "tudo", "all"}:
        return projetos

    indices: list[int] = []
    for parte in selecao.split(","):
        parte = parte.strip()
        if not parte:
            continue
        if not parte.isdigit():
            raise ValueError(f"Selecao invalida: {parte}")
        indices.append(int(parte))

    escolhidos: list[dict[str, str]] = []
    for indice in indices:
        if indice < 1 or indice > len(projetos):
            raise ValueError(f"Indice fora da lista: {indice}")
        escolhidos.append(projetos[indice - 1])

    if not escolhidos:
        raise ValueError("Nenhuma concessao/projeto foi selecionado.")

    return escolhidos


def endpoint_participantes(connect_space_id: str, project_id: str) -> str:
    return f"/{connect_space_id}/ProjectSettings/{project_id}/UserGroupsAndUsers"


def url_tela_participantes(connect_space_id: str, project_id: str) -> str:
    return f"{ORIGEM_PWDM}/{connect_space_id}/ProjectSettings/{project_id}/View#PARTICIPANTS"


def abrir_tela_e_buscar_dados(page: Page, connect_space_id: str, project_id: str) -> dict[str, Any]:
    url_tela = url_tela_participantes(connect_space_id, project_id)
    endpoint = endpoint_participantes(connect_space_id, project_id)

    pagina_api = page.context.new_page()
    pagina_api.goto(url_tela, wait_until="domcontentloaded")
    try:
        pagina_api.wait_for_load_state("networkidle", timeout=30000)
    except Exception:
        print("[AVISO] Tela ainda tinha requisicoes ativas; seguindo com a consulta.")

    resultado = pagina_api.evaluate(
        """
        async ({ endpoint }) => {
            const response = await fetch(endpoint, {
                method: "GET",
                credentials: "include",
                headers: {
                    "Accept": "application/json",
                    "X-Requested-With": "XMLHttpRequest"
                }
            });

            const text = await response.text();
            return {
                ok: response.ok,
                status: response.status,
                statusText: response.statusText,
                contentType: response.headers.get("content-type") || "",
                text
            };
        }
        """,
        {"endpoint": endpoint},
    )

    if not resultado["ok"] or "application/json" not in resultado["contentType"].lower():
        raise RuntimeError(
            f"Falha ao buscar participantes em {project_id}: "
            f"HTTP {resultado['status']} {resultado['statusText']}. "
            f"Inicio da resposta: {resultado['text'][:300]}"
        )

    dados = json.loads(resultado["text"])
    if not isinstance(dados, dict):
        raise RuntimeError(f"Resposta inesperada ao buscar participantes em {project_id}.")

    return dados


def obter_membro(all_members: list[dict[str, Any]], indice_ou_id: Any) -> dict[str, Any]:
    if isinstance(indice_ou_id, int) and 0 <= indice_ou_id < len(all_members):
        return all_members[indice_ou_id]

    for membro in all_members:
        if str(membro.get("id", "")) == str(indice_ou_id):
            return membro

    return {}


def buscar_usuario_no_projeto(dados: dict[str, Any], email: str) -> dict[str, Any]:
    all_members = dados.get("allMembers") or []
    direct_members_idx = dados.get("directMembersIdx") or []

    for indice in direct_members_idx:
        membro = obter_membro(all_members, indice)
        if str(membro.get("email", "")).lower() == email:
            return {"status": "participante", "membro": membro}

    for membro in all_members:
        if str(membro.get("email", "")).lower() == email:
            return {"status": "candidato", "membro": membro}

    return {"status": "nao_encontrado", "membro": {}}


def permissoes_usuario_atual(dados: dict[str, Any]) -> dict[str, bool]:
    permissoes = (
        dados.get("currentUserPermissions")
        or dados.get("currentUser")
        or dados.get("currentParticipant")
        or dados.get("loggedInUser")
        or dados.get("userPermissions")
    )
    if not isinstance(permissoes, dict):
        return {}

    return {
        "isMember": bool(permissoes.get("isMember")),
        "canView": bool(permissoes.get("canView")),
        "canReceive": bool(permissoes.get("canReceive")),
        "canIssue": bool(permissoes.get("canIssue")),
        "canReviewAnswer": bool(permissoes.get("canReviewAnswer")),
        "isAdmin": bool(permissoes.get("isAdmin")),
        "canInvite": bool(permissoes.get("canInvite") or permissoes.get("canAddParticipant")),
    }


def montar_operacoes(
    page: Page,
    projetos: list[dict[str, Any]],
    email: str,
    acao_usuario: str,
    titulo: str,
    permissoes: dict[str, bool],
) -> list[dict[str, Any]]:
    print("\n[7/9] Lendo participantes dos projetos selecionados...")
    operacoes: list[dict[str, Any]] = []

    for indice, projeto in enumerate(projetos, start=1):
        connect_space_id = projeto["connectSpaceId"]
        project_id = projeto["projectId"]
        print(f"- Projeto {indice}: {projeto['nome']} ({project_id})")
        dados = abrir_tela_e_buscar_dados(page, connect_space_id, project_id)
        usuario = buscar_usuario_no_projeto(dados, email)
        permissoes_atuais = permissoes_usuario_atual(dados)
        acao_efetiva, descricao_acao = determinar_acao_efetiva(acao_usuario, usuario["status"])

        operacoes.append(
            {
                "nomeProjeto": projeto["nome"],
                "origemProjectWise": projeto.get("origemProjectWise"),
                "criterioCruzamento": projeto.get("criterioCruzamento"),
                "connectSpaceId": connect_space_id,
                "projectId": project_id,
                "acaoSolicitada": acao_usuario,
                "acaoEfetiva": acao_efetiva,
                "descricaoAcao": descricao_acao,
                "status": usuario["status"],
                "membro": usuario["membro"],
                "permissoesUsuarioAtual": permissoes_atuais,
                "titulo": titulo,
                "permissoes": permissoes,
            }
        )

    return operacoes


def exibir_previa(operacoes: list[dict[str, Any]], email: str) -> None:
    print("\n[8/9] Previa das acoes:")
    exibir_resumo_operacoes(operacoes)

    for indice, op in enumerate(operacoes, start=1):
        status = op["status"]
        membro = op["membro"]
        nome = membro.get("name") or membro.get("displayName") or membro.get("email") or email
        acao = op.get("descricaoAcao") or determinar_acao_efetiva(
            op.get("acaoSolicitada", "incluir"),
            status,
        )[1]

        print(f"{indice:02d}. Projeto {op.get('nomeProjeto') or op['projectId']}")
        origem_pw = op.get("origemProjectWise") or {}
        if origem_pw:
            print(
                f"    ProjectWise: {origem_pw.get('nome')} | "
                f"PW ID: {origem_pw.get('id')} | GUID: {origem_pw.get('guid')}"
            )
            print(f"    Cruzamento: {op.get('criterioCruzamento')}")
        print(f"    Usuario: {nome} <{email}>")
        print(f"    Situacao: {status}")
        print(f"    Acao   : {acao}")
        atuais = op.get("permissoesUsuarioAtual") or {}
        if atuais:
            print(
                "    Seu acesso: "
                f"membro={atuais.get('isMember')}, "
                f"receber={atuais.get('canReceive')}, "
                f"emitir={atuais.get('canIssue')}, "
                f"admin={atuais.get('isAdmin')}, "
                f"convidar={atuais.get('canInvite')}"
            )
        if op.get("acaoEfetiva") not in {"ignorar_usuario_ausente", "excluir_participante"}:
            print(f"    Titulo : {op['titulo']}")
            print(f"    Perms  : {descricao_permissoes(op['permissoes'])}")
        elif op.get("acaoEfetiva") == "excluir_participante":
            print("    Obs    : exclusao ainda nao sera aplicada automaticamente.")


def contar_operacoes_por_acao(operacoes: list[dict[str, Any]]) -> dict[str, int]:
    contagem: dict[str, int] = {}
    for op in operacoes:
        chave = str(op.get("acaoEfetiva") or op.get("status") or "desconhecido")
        contagem[chave] = contagem.get(chave, 0) + 1
    return contagem


def exibir_resumo_operacoes(operacoes: list[dict[str, Any]]) -> None:
    contagem = contar_operacoes_por_acao(operacoes)
    print(f"Total de projetos: {len(operacoes)}")
    for chave, quantidade in sorted(contagem.items()):
        print(f"- {chave}: {quantidade}")


def confirmar_aplicacao(operacoes: list[dict[str, Any]]) -> bool:
    aplicaveis = [
        op
        for op in operacoes
        if op.get("acaoEfetiva") not in {"ignorar_usuario_ausente", "excluir_participante"}
    ]

    if not aplicaveis:
        print("\n[AVISO] Nao ha alteracoes aplicaveis nesta selecao.")
        return False

    print("\nConfirmacao final:")
    print(f"- Projetos selecionados: {len(operacoes)}")
    print(f"- Operacoes aplicaveis: {len(aplicaveis)}")
    print("- Para aplicar, digite exatamente: CONFIRMAR")
    resposta = input("Confirmacao: ").strip()
    return resposta == "CONFIRMAR"


def url_absoluta_pwdm(endpoint: str) -> str:
    if endpoint.startswith("http://") or endpoint.startswith("https://"):
        return endpoint
    return f"{ORIGEM_PWDM}{endpoint if endpoint.startswith('/') else '/' + endpoint}"


def obter_cookie(context: BrowserContext, nome: str, url: str) -> str:
    for cookie in context.cookies(url):
        if cookie.get("name") == nome:
            return str(cookie.get("value") or "")
    return ""


def extrair_token_html(html: str) -> str:
    padroes = [
        r"name=[\"']__RequestVerificationToken[\"'][^>]*value=[\"']([^\"']+)",
        r"value=[\"']([^\"']+)[\"'][^>]*name=[\"']__RequestVerificationToken[\"']",
        r"name=[\"']__requestverificationtoken[\"'][^>]*value=[\"']([^\"']+)",
        r"value=[\"']([^\"']+)[\"'][^>]*name=[\"']__requestverificationtoken[\"']",
    ]
    for padrao in padroes:
        match = re.search(padrao, html, flags=re.IGNORECASE)
        if match:
            return match.group(1)
    return ""


def obter_token_verificacao(pagina_api: Page, url_tela: str) -> tuple[str, str]:
    token_dom = pagina_api.evaluate(
        """
        () => {
            const input =
                document.querySelector("input[name='__RequestVerificationToken']") ||
                document.querySelector("input[name='__requestverificationtoken']") ||
                Array.from(document.querySelectorAll("input")).find((item) =>
                    (item.name || "").toLowerCase().includes("requestverificationtoken")
                );
            const meta =
                document.querySelector("meta[name='csrf-token']") ||
                document.querySelector("meta[name='RequestVerificationToken']") ||
                document.querySelector("meta[name='__requestverificationtoken']");
            return input?.value || meta?.content || "";
        }
        """
    )
    if token_dom:
        return str(token_dom), "dom"

    url_sem_hash = url_tela.split("#", 1)[0]
    try:
        resposta = pagina_api.context.request.get(url_sem_hash, timeout=30000)
        if resposta.ok:
            token_html = extrair_token_html(resposta.text())
            if token_html:
                return token_html, "html-view"
    except Exception:
        pass

    token_cookie = obter_cookie(pagina_api.context, "__RequestVerificationToken", ORIGEM_PWDM)
    if token_cookie:
        return token_cookie, "context-cookie"

    return "", ""


def post_json(page: Page, url_tela: str, endpoint: str, payload: dict[str, Any]) -> dict[str, Any]:
    pagina_api = page.context.new_page()
    pagina_api.goto(url_tela, wait_until="domcontentloaded")
    try:
        pagina_api.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    url_endpoint = url_absoluta_pwdm(endpoint)
    token_verificacao, origem_token = obter_token_verificacao(pagina_api, url_tela)
    resultado = pagina_api.evaluate(
        """
        async ({ endpoint, payload, referer, tokenVerificacao, origemToken }) => {
            const cookieValue = (name) => {
                const cookie = document.cookie || "";
                const item = cookie.split("; ").find((parte) => parte.startsWith(name + "="));
                return item ? decodeURIComponent(item.slice(name.length + 1)) : "";
            };

            const token =
                document.querySelector("input[name='__RequestVerificationToken']")?.value ||
                document.querySelector("input[name='__requestverificationtoken']")?.value ||
                document.querySelector("meta[name='csrf-token']")?.content ||
                document.querySelector("meta[name='RequestVerificationToken']")?.content ||
                document.querySelector("meta[name='__requestverificationtoken']")?.content ||
                tokenVerificacao ||
                cookieValue("__RequestVerificationToken") ||
                cookieValue("XSRF-TOKEN") ||
                "";

            const headers = {
                "Accept": "application/json, text/plain, */*",
                "Content-Type": "application/json",
                "X-Requested-With": "XMLHttpRequest"
            };

            if (token) {
                headers["__requestverificationtoken"] = token;
                headers["RequestVerificationToken"] = token;
                headers["X-CSRF-TOKEN"] = token;
                headers["X-XSRF-TOKEN"] = token;
            }

            const response = await fetch(endpoint, {
                method: "POST",
                credentials: "include",
                headers,
                referrer: referer,
                body: JSON.stringify(payload)
            });

            const text = await response.text();
            let body = null;
            try {
                body = text ? JSON.parse(text) : null;
            } catch {
                body = text;
            }

            return {
                ok: response.ok,
                status: response.status,
                statusText: response.statusText,
                contentType: response.headers.get("content-type") || "",
                responseUrl: response.url,
                pageUrl: location.href,
                tokenFound: Boolean(token),
                tokenSource: token ? (origemToken || "dom") : "",
                body,
                text
            };
        }
        """,
        {
            "endpoint": url_endpoint,
            "payload": payload,
            "referer": url_tela,
            "tokenVerificacao": token_verificacao,
            "origemToken": origem_token,
        },
    )

    if not resultado["ok"]:
        detalhe = resultado.get("body")
        if detalhe is None:
            detalhe = str(resultado.get("text") or "").strip()[:500] or "sem corpo de resposta"
        payload_log = {chave: ("***" if chave.lower() == "email" else valor) for chave, valor in payload.items()}
        raise RuntimeError(
            f"POST falhou em {url_endpoint}: HTTP {resultado['status']} "
            f"{resultado['statusText']} - {detalhe} "
            f"(pagina={resultado.get('pageUrl')}, resposta={resultado.get('responseUrl')}, "
            f"token={'sim' if resultado.get('tokenFound') else 'nao'}, "
            f"origem_token={resultado.get('tokenSource') or 'nenhuma'}, "
            f"payload={payload_log})"
        )

    return resultado


def sanitizar_para_log(valor: Any, chave: str = "") -> Any:
    chave_lower = chave.lower()

    if chave_lower == "text":
        return "[omitido]"

    if any(segredo in chave_lower for segredo in ["cookie", "authorization"]):
        return "[removido]"

    if "token" in chave_lower and chave_lower not in {"tokensource", "tokenfound"}:
        return "[removido]"

    if isinstance(valor, dict):
        return {item_chave: sanitizar_para_log(item_valor, item_chave) for item_chave, item_valor in valor.items()}

    if isinstance(valor, list):
        return [sanitizar_para_log(item, chave) for item in valor]

    if isinstance(valor, str):
        if EMAIL_REGEX.match(valor):
            return mascarar_email(valor)
        return valor

    return valor


def status_resultado(resultado: dict[str, Any]) -> str:
    return str(resultado.get("status") or "desconhecido")


def http_status_resultado(resultado: dict[str, Any]) -> str:
    dados = resultado.get("resultado")
    if isinstance(dados, dict):
        return str(dados.get("status") or "")
    dados_permissao = resultado.get("resultado_permissao")
    if isinstance(dados_permissao, dict):
        return str(dados_permissao.get("status") or "")
    return ""


def aplicar_operacao(page: Page, op: dict[str, Any], email: str) -> dict[str, Any]:
    status = op["status"]
    acao_efetiva = op.get("acaoEfetiva")
    connect_space_id = op["connectSpaceId"]
    project_id = op["projectId"]
    membro = op["membro"]
    permissoes = normalizar_permissoes(op["permissoes"])
    tela = url_tela_participantes(connect_space_id, project_id)

    if not acao_efetiva:
        acao_efetiva = determinar_acao_efetiva(op.get("acaoSolicitada", "incluir"), status)[0]

    if acao_efetiva == "ignorar_usuario_ausente":
        return {
            "status": "ignorado",
            "mensagem": "Usuario nao esta como participante direto neste projeto.",
        }

    if acao_efetiva == "excluir_participante":
        return {
            "status": "nao_implementado",
            "mensagem": (
                "Exclusao ainda nao foi automatizada. Precisamos capturar e validar o endpoint "
                "de remocao antes de aplicar uma acao destrutiva."
            ),
        }

    if acao_efetiva == "atualizar_participante":
        if permissoes_iguais(membro, permissoes) and titulo_igual(membro, op["titulo"]):
            return {
                "status": "sem_alteracao",
                "mensagem": "Participante ja estava com as permissoes e o titulo desejados.",
            }

        endpoint_permissao = endpoint_participantes(connect_space_id, project_id).replace(
            "UserGroupsAndUsers",
            "ProjectParticipant",
        )
        payload = {
            "email": email,
            "canView": permissoes["canView"],
            "canReceive": permissoes["canReceive"],
            "canIssue": permissoes["canIssue"],
            "canReviewAnswer": permissoes["canReviewAnswer"],
            "isAdmin": permissoes["isAdmin"],
        }

        if membro.get("isPWSynced"):
            endpoint_permissao = endpoint_participantes(connect_space_id, project_id).replace(
                "UserGroupsAndUsers",
                "PWProjectParticipantOrGroup",
            )
            payload["groupId"] = membro.get("id")
        else:
            payload["participantId"] = membro.get("id")

        resultado_permissao = post_json(page, tela, endpoint_permissao, payload)

        payload_titulo = {
            "participantIds": [membro.get("id")],
            "roleTitle": op["titulo"],
        }
        endpoint_titulo = endpoint_participantes(connect_space_id, project_id).replace(
            "UserGroupsAndUsers",
            "ProjectParticipants",
        )
        resultado_titulo = post_json(page, tela, endpoint_titulo, payload_titulo)

        return {
            "status": "atualizado",
            "resultado_permissao": resultado_permissao,
            "resultado_titulo": resultado_titulo,
        }

    if acao_efetiva == "adicionar_existente":
        payload = {
            "participantId": membro.get("id"),
            "roleTitle": op["titulo"],
            "isAdmin": permissoes["isAdmin"],
            "canReviewAnswer": permissoes["canReviewAnswer"],
            "canReceive": permissoes["canReceive"],
            "canIssue": permissoes["canIssue"],
        }
        endpoint = endpoint_participantes(connect_space_id, project_id).replace("UserGroupsAndUsers", "AddExistingPWParticipant")
        resultado = post_json(page, tela, endpoint, payload)
        return {"status": "adicionado", "resultado": resultado}

    if acao_efetiva != "adicionar_por_email":
        raise RuntimeError(f"Acao efetiva desconhecida: {acao_efetiva}")

    payload = {
        "email": email,
        "roleTitle": op["titulo"],
        "isAdmin": permissoes["isAdmin"],
        "canReviewAnswer": permissoes["canReviewAnswer"],
        "canReceive": permissoes["canReceive"],
        "canIssue": permissoes["canIssue"],
    }
    endpoint = endpoint_participantes(connect_space_id, project_id).replace("UserGroupsAndUsers", "ProjectParticipant")
    resultado = post_json(page, tela, endpoint, payload)
    return {"status": "adicionado_por_email", "resultado": resultado}


def aplicar_operacoes(page: Page, operacoes: list[dict[str, Any]], email: str) -> list[dict[str, Any]]:
    print("\n[9/9] Aplicando alteracoes confirmadas...")
    resultados = []

    for indice, op in enumerate(operacoes, start=1):
        print(f"- Projeto {indice}: {op.get('nomeProjeto') or op['projectId']} [{op.get('acaoEfetiva')}]")
        try:
            resultado = aplicar_operacao(page, op, email)
            print(f"  [OK] {resultado['status']}")
        except Exception as erro:
            resultado = {"status": "erro", "erro": str(erro)}
            print(f"  [ERRO] {erro}")

        resultados.append(
            {
                "projectId": op["projectId"],
                "nomeProjeto": op.get("nomeProjeto"),
                "origemProjectWise": op.get("origemProjectWise"),
                "criterioCruzamento": op.get("criterioCruzamento"),
                "email": email,
                "acao_solicitada": op.get("acaoSolicitada"),
                "acao_planejada": op.get("acaoEfetiva") or op["status"],
                "situacao_usuario": op["status"],
                "resultado": resultado,
            }
        )

    return resultados


def salvar_log(operacoes: list[dict[str, Any]], resultados: list[dict[str, Any]]) -> None:
    print("\nSalvando log...")
    execucao = nome_execucao()
    arquivo = PASTA_LOGS / f"pwdm_gerenciamento_{execucao}.json"
    dados = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "operacoes": sanitizar_para_log(operacoes),
        "resultados": sanitizar_para_log(resultados),
    }

    with arquivo.open("w", encoding="utf-8") as saida:
        json.dump(dados, saida, indent=2, ensure_ascii=False)

    print(f"[OK] Log: {arquivo.resolve()}")
    salvar_resumo_csv(resultados, execucao)


def salvar_resumo_csv(resultados: list[dict[str, Any]], execucao: str) -> None:
    arquivo = PASTA_LOGS / f"pwdm_resumo_{execucao}.csv"
    campos = [
        "timestamp",
        "nomeProjeto",
        "projectId",
        "email",
        "acao_solicitada",
        "acao_planejada",
        "situacao_usuario",
        "status",
        "http_status",
    ]

    with arquivo.open("w", encoding="utf-8-sig", newline="") as saida:
        escritor = csv.DictWriter(saida, fieldnames=campos, extrasaction="ignore")
        escritor.writeheader()
        for item in resultados:
            resultado = item.get("resultado") if isinstance(item, dict) else {}
            if not isinstance(resultado, dict):
                resultado = {}
            escritor.writerow(
                {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "nomeProjeto": item.get("nomeProjeto"),
                    "projectId": item.get("projectId"),
                    "email": mascarar_email(str(item.get("email") or "")) if item.get("email") else "",
                    "acao_solicitada": item.get("acao_solicitada"),
                    "acao_planejada": item.get("acao_planejada"),
                    "situacao_usuario": item.get("situacao_usuario"),
                    "status": status_resultado(resultado),
                    "http_status": http_status_resultado(resultado),
                }
            )

    print(f"[OK] Resumo CSV: {arquivo.resolve()}")


def main() -> None:
    preparar_pasta_logs()

    with sync_playwright() as playwright:
        browser: Optional[Browser] = None

        try:
            browser, _context, page = iniciar_navegador(playwright)
            abrir_portal(page)

            print("\nDados da operacao")
            email = solicitar_email()
            projetos_disponiveis = buscar_concessoes_disponiveis(page)
            usuario_validado = validar_cadastro_usuario_pwdm(page, email, projetos_disponiveis)
            print(
                "[OK] Cadastro validado para uso no PWDM: "
                f"{usuario_validado.get('nome')} <{usuario_validado.get('email')}>"
            )
            acao_usuario = solicitar_acao_usuario()

            arvore_projectwise = carregar_arvore_projectwise()
            concessao_selecionada = selecionar_concessao_projectwise(arvore_projectwise)
            projetos_projectwise = selecionar_projetos_projectwise(concessao_selecionada)
            projetos_selecionados = cruzar_projetos_projectwise_com_pwdm(
                page,
                projetos_projectwise,
                projetos_disponiveis,
            )

            if acao_usuario == "excluir":
                print(
                    "\n[AVISO] A exclusao sera listada na previa, mas ainda nao sera aplicada. "
                    "Precisamos capturar o endpoint correto antes de remover participantes."
                )
                titulo = ""
                permissoes = normalizar_permissoes({chave: False for chave in CHAVES_PERMISSOES})
            else:
                titulo = solicitar_titulo()
                permissoes = solicitar_permissoes()

            operacoes = montar_operacoes(
                page,
                projetos_selecionados,
                email,
                acao_usuario,
                titulo,
                permissoes,
            )
            exibir_previa(operacoes, email)

            if confirmar_aplicacao(operacoes):
                resultados = aplicar_operacoes(page, operacoes, email)
            else:
                print("\n[OK] Nenhuma alteracao aplicada.")
                resultados = [{"status": "cancelado_pelo_usuario"}]

            salvar_log(operacoes, resultados)
            print("\nFinalizado.")
            input("Pressione ENTER para fechar o navegador...")

        except KeyboardInterrupt:
            print("\nExecucao interrompida pelo usuario.")
        except Exception as erro:
            print(f"\n[ERRO] {erro}")
            input("Pressione ENTER para fechar o navegador...")
        finally:
            if browser:
                browser.close()


if __name__ == "__main__":
    main()
