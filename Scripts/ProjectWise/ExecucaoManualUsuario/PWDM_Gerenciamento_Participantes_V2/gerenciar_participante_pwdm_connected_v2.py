import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any, Optional
from tkinter import Tk, filedialog

from playwright.sync_api import Browser, sync_playwright

from gerenciar_participante_pwdm_v2 import (
    CHAVES_PERMISSOES,
    abrir_portal,
    determinar_acao_efetiva,
    endpoint_participantes,
    exibir_previa,
    iniciar_navegador,
    buscar_usuario_no_projeto,
    normalizar_permissoes,
    obter_token_verificacao,
    permissoes_iguais,
    permissoes_usuario_atual,
    preparar_pasta_logs,
    salvar_log,
    solicitar_acao_usuario,
    solicitar_email,
    solicitar_permissoes,
    solicitar_titulo,
    titulo_igual,
    url_tela_participantes,
    url_absoluta_pwdm,
)
from lote_usuarios_pwdm_v2 import carregar_emails_xlsx
from projetos_projectwise_pwdm_v2 import carregar_projetos_selecionados_com_diagnostico


PASTA_LOGS = Path("Logs")
SCRIPT_SELETOR_PROJECTWISE = Path("pw_selecionar_projetos_para_pwdm_v2.ps1")


def nome_execucao_connected() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def executar_seletor_projectwise() -> Path:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    arquivo = PASTA_LOGS / f"pwdm_projetos_execucao_connected_{nome_execucao_connected()}.json"
    script = Path(__file__).resolve().parent / SCRIPT_SELETOR_PROJECTWISE

    if not script.exists():
        raise FileNotFoundError(f"Seletor ProjectWise nao encontrado: {script}")

    print("\n[0/9] Selecionando concessao/projetos no ProjectWise...")
    print("Uma janela/login do ProjectWise pode ser aberta se a sessao ainda nao estiver ativa.")

    comando = [
        "powershell.exe",
        "-NoProfile",
        "-MTA",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-Saida",
        str(arquivo),
    ]
    resultado = subprocess.run(comando, cwd=Path(__file__).resolve().parent)
    if resultado.returncode != 0:
        raise RuntimeError(f"Selecao ProjectWise falhou com codigo {resultado.returncode}.")

    if not arquivo.exists():
        raise FileNotFoundError(f"Seletor ProjectWise nao gerou o JSON esperado: {arquivo}")

    return arquivo


def apagar_json_temporario(arquivo: Optional[Path]) -> None:
    if not arquivo:
        return
    try:
        if arquivo.exists():
            arquivo.unlink()
            print(f"[INFO] JSON temporario removido: {arquivo.resolve()}")
    except Exception as erro:
            print(f"[AVISO] Nao foi possivel remover JSON temporario {arquivo}: {erro}")


def carregar_projetos_para_fluxo(arquivo: Path) -> list[dict]:
    print(f"\n[INFO] Usando projetos selecionados pelo ProjectWise nesta execucao: {arquivo.resolve()}")
    projetos, ignorados = carregar_projetos_selecionados_com_diagnostico(arquivo)
    if ignorados:
        print("\n[AVISO] Projeto(s) ignorado(s) por nao terem ConnectedProjectId valido para PWDM:")
        for item in ignorados:
            print(f"- {item}")
        print("Esses projetos nao serao consultados no PWDM nesta execucao.")

    print(f"[OK] {len(projetos)} projeto(s) carregado(s) para PWDM pelo ConnectedProjectId.")
    for indice, projeto in enumerate(projetos, start=1):
        origem = projeto.get("origemProjectWise") or {}
        print(
            f"{indice:03d}. {projeto.get('nome')} | "
            f"PW ID: {origem.get('id')} | ConnectedProjectId: {projeto.get('projectId')}"
        )
    return projetos


def abrir_tela_e_buscar_dados_em_aba(pagina_api, connect_space_id: str, project_id: str) -> dict[str, Any]:
    url_tela = url_tela_participantes(connect_space_id, project_id)
    endpoint = endpoint_participantes(connect_space_id, project_id)

    pagina_api.goto(url_tela, wait_until="domcontentloaded")
    try:
        pagina_api.wait_for_load_state("networkidle", timeout=15000)
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

    if not resultado.get("ok"):
        raise RuntimeError(
            f"Falha ao buscar participantes em {project_id}: "
            f"HTTP {resultado.get('status')} {resultado.get('statusText')}. "
            f"Inicio da resposta: {str(resultado.get('text') or '')[:300]}"
        )

    texto = str(resultado.get("text") or "")
    try:
        import json

        return json.loads(texto)
    except Exception as erro:
        raise RuntimeError(f"Resposta inesperada ao buscar participantes em {project_id}.") from erro


def montar_operacoes_com_aba_unica(
    page,
    projetos: list[dict[str, Any]],
    email: str,
    acao_usuario: str,
    titulo: str,
    permissoes: dict[str, bool],
) -> list[dict[str, Any]]:
    print("\n[7/9] Lendo participantes dos projetos selecionados...")
    print("[INFO] Usando uma unica aba auxiliar para consultar todos os projetos.")
    operacoes: list[dict[str, Any]] = []
    pagina_api = page.context.new_page()

    try:
        for indice, projeto in enumerate(projetos, start=1):
            connect_space_id = projeto["connectSpaceId"]
            project_id = projeto["projectId"]
            print(f"- Projeto {indice}: {projeto['nome']} ({project_id})")
            dados = abrir_tela_e_buscar_dados_em_aba(pagina_api, connect_space_id, project_id)
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
    finally:
        try:
            pagina_api.close()
        except Exception:
            pass

    return operacoes


def montar_operacoes_lote_com_aba_unica(
    page,
    projetos: list[dict[str, Any]],
    emails: list[str],
    acao_usuario: str,
    titulo: str,
    permissoes: dict[str, bool],
) -> list[dict[str, Any]]:
    print("\n[7/9] Lendo participantes dos projetos selecionados para o lote...")
    print("[INFO] Usando uma unica aba auxiliar para consultar todos os projetos.")
    operacoes: list[dict[str, Any]] = []
    pagina_api = page.context.new_page()

    try:
        for indice, projeto in enumerate(projetos, start=1):
            connect_space_id = projeto["connectSpaceId"]
            project_id = projeto["projectId"]
            print(f"- Projeto {indice}: {projeto['nome']} ({project_id})")
            dados = abrir_tela_e_buscar_dados_em_aba(pagina_api, connect_space_id, project_id)
            permissoes_atuais = permissoes_usuario_atual(dados)

            for email in emails:
                usuario = buscar_usuario_no_projeto(dados, email)
                acao_efetiva, descricao_acao = determinar_acao_efetiva(acao_usuario, usuario["status"])

                operacoes.append(
                    {
                        "email": email,
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
    finally:
        try:
            pagina_api.close()
        except Exception:
            pass

    return operacoes


def confirmar_aplicacao_sn(operacoes: list[dict[str, Any]]) -> bool:
    aplicaveis = [
        op
        for op in operacoes
        if op.get("acaoEfetiva") not in {"ignorar_usuario_ausente", "excluir_participante"}
    ]

    if not aplicaveis:
        print("\n[AVISO] Nao ha alteracoes aplicaveis nesta selecao.")
        return False

    print("\nConfirmacao final:")
    emails = sorted({str(op.get("email") or "") for op in operacoes if op.get("email")})
    projetos = sorted({str(op.get("projectId") or "") for op in operacoes if op.get("projectId")})
    if emails:
        print(f"- Usuarios selecionados: {len(emails)}")
    print(f"- Projetos selecionados: {len(projetos) or len(operacoes)}")
    print(f"- Operacoes planejadas: {len(operacoes)}")
    print(f"- Operacoes aplicaveis: {len(aplicaveis)}")
    print("- Responda S para aplicar ou N para cancelar.")

    while True:
        resposta = input("Aplicar alteracoes? [S/N]: ").strip().lower()
        if resposta in {"s", "sim", "y", "yes"}:
            return True
        if resposta in {"n", "nao", "não", "no", ""}:
            return False
        print("[AVISO] Responda apenas S ou N.")


def post_json_em_aba(pagina_api, url_tela: str, endpoint: str, payload: dict[str, Any]) -> dict[str, Any]:
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


def aplicar_operacao_em_aba(pagina_api, op: dict[str, Any], email: str) -> dict[str, Any]:
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

        resultado_permissao = post_json_em_aba(pagina_api, tela, endpoint_permissao, payload)

        payload_titulo = {
            "participantIds": [membro.get("id")],
            "roleTitle": op["titulo"],
        }
        endpoint_titulo = endpoint_participantes(connect_space_id, project_id).replace(
            "UserGroupsAndUsers",
            "ProjectParticipants",
        )
        resultado_titulo = post_json_em_aba(pagina_api, tela, endpoint_titulo, payload_titulo)

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
        resultado = post_json_em_aba(pagina_api, tela, endpoint, payload)
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
    resultado = post_json_em_aba(pagina_api, tela, endpoint, payload)
    return {"status": "adicionado_por_email", "resultado": resultado}


def aplicar_operacoes_com_aba_unica(page, operacoes: list[dict[str, Any]], email: str) -> list[dict[str, Any]]:
    print("\n[9/9] Aplicando alteracoes confirmadas...")
    print("[INFO] Usando uma unica aba auxiliar para aplicar todos os POSTs.")
    resultados = []
    pagina_api = page.context.new_page()

    try:
        for indice, op in enumerate(operacoes, start=1):
            print(f"- Projeto {indice}: {op.get('nomeProjeto') or op['projectId']} [{op.get('acaoEfetiva')}]")
            try:
                resultado = aplicar_operacao_em_aba(pagina_api, op, email)
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
    finally:
        try:
            pagina_api.close()
        except Exception:
            pass

    return resultados


def aplicar_operacoes_lote_com_aba_unica(page, operacoes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    print("\n[9/9] Aplicando alteracoes confirmadas...")
    print("[INFO] Usando uma unica aba auxiliar para aplicar todos os POSTs do lote.")
    resultados = []
    pagina_api = page.context.new_page()

    try:
        for indice, op in enumerate(operacoes, start=1):
            email = str(op.get("email") or "")
            print(
                f"- Operacao {indice}/{len(operacoes)}: "
                f"{email} | {op.get('nomeProjeto') or op['projectId']} [{op.get('acaoEfetiva')}]"
            )
            try:
                resultado = aplicar_operacao_em_aba(pagina_api, op, email)
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
    finally:
        try:
            pagina_api.close()
        except Exception:
            pass

    return resultados


def confirmar_email_informado(email: str) -> None:
    print(f"\nE-mail informado: {email}")
    while True:
        resposta = input("Continuar com este e-mail? [S/N]: ").strip().lower()
        if resposta in {"s", "sim", "y", "yes"}:
            return
        if resposta in {"n", "nao", "não", "no", ""}:
            raise RuntimeError("E-mail nao confirmado. Nenhuma alteracao sera preparada.")
        print("[AVISO] Responda apenas S ou N.")


def solicitar_modo_usuarios() -> str:
    print("\nModo de usuarios")
    print("01. Usuario unico")
    print("02. Lote por planilha .xlsx")
    while True:
        resposta = input("Escolha [1/2]: ").strip()
        if resposta in {"1", "01"}:
            return "unico"
        if resposta in {"2", "02"}:
            return "lote_xlsx"
        print("[AVISO] Escolha 1 ou 2.")


def selecionar_planilha_usuarios_xlsx() -> Path:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha .xlsx de usuarios",
            filetypes=[
                ("Planilhas Excel", "*.xlsx"),
                ("Todos os arquivos", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not caminho:
        raise RuntimeError("Nenhuma planilha foi selecionada.")

    return Path(caminho)


def carregar_emails_lote_interativo() -> list[str]:
    caminho = selecionar_planilha_usuarios_xlsx()
    resumo = carregar_emails_xlsx(caminho)
    emails = list(resumo["emailsValidosUnicos"])

    print(f"\n[OK] Planilha importada: {Path(resumo['arquivo']).resolve()}")
    print(f"- Aba: {resumo['aba']}")
    print(f"- Linhas de dados: {resumo['linhasDados']}")
    print(f"- E-mails encontrados: {resumo['emailsEncontrados']}")
    print(f"- E-mails validos unicos: {len(emails)}")
    print(f"- Duplicados ignorados: {len(resumo['duplicados'])}")
    print(f"- Invalidos ignorados: {len(resumo['invalidos'])}")
    print(f"- Linhas sem e-mail ignoradas: {len(resumo['linhasSemEmail'])}")

    if resumo["invalidos"]:
        print("[AVISO] E-mails invalidos ignorados:")
        for item in resumo["invalidos"][:20]:
            print(f"  - Linha {item['linha']}: {item['email']}")
        if len(resumo["invalidos"]) > 20:
            print(f"  ... mais {len(resumo['invalidos']) - 20} invalido(s).")

    return emails


def confirmar_emails_lote(emails: list[str]) -> None:
    print("\nUsuarios do lote:")
    for indice, email in enumerate(emails, start=1):
        print(f"{indice:03d}. {email}")
    while True:
        resposta = input("Continuar com estes usuarios? [S/N]: ").strip().lower()
        if resposta in {"s", "sim", "y", "yes"}:
            return
        if resposta in {"n", "nao", "não", "no", ""}:
            raise RuntimeError("Lote nao confirmado. Nenhuma alteracao sera preparada.")
        print("[AVISO] Responda apenas S ou N.")


def exibir_previa_lote(operacoes: list[dict[str, Any]], emails: list[str]) -> None:
    print("\n[8/9] Previa consolidada do lote:")
    projetos = sorted({str(op.get("projectId") or "") for op in operacoes if op.get("projectId")})
    aplicaveis = [
        op
        for op in operacoes
        if op.get("acaoEfetiva") not in {"ignorar_usuario_ausente", "excluir_participante"}
    ]
    print(f"- Usuarios: {len(emails)}")
    print(f"- Projetos: {len(projetos)}")
    print(f"- Operacoes planejadas: {len(operacoes)}")
    print(f"- Operacoes aplicaveis: {len(aplicaveis)}")

    for email in emails:
        operacoes_email = [op for op in operacoes if op.get("email") == email]
        exibir_previa(operacoes_email, email)


def main() -> None:
    preparar_pasta_logs()
    arquivo_projetos_execucao: Optional[Path] = None

    with sync_playwright() as playwright:
        browser: Optional[Browser] = None

        try:
            arquivo_projetos_execucao = executar_seletor_projectwise()
            projetos_selecionados = carregar_projetos_para_fluxo(arquivo_projetos_execucao)

            browser, _context, page = iniciar_navegador(playwright)
            abrir_portal(page)

            print("\nDados da operacao")
            modo_usuarios = solicitar_modo_usuarios()
            if modo_usuarios == "lote_xlsx":
                emails = carregar_emails_lote_interativo()
                confirmar_emails_lote(emails)
            else:
                email = solicitar_email()
                confirmar_email_informado(email)
                emails = [email]

            acao_usuario = solicitar_acao_usuario()

            if acao_usuario == "excluir":
                print(
                    "\n[AVISO] A exclusao sera listada na previa, mas ainda nao sera aplicada. "
                    "Precisamos capturar e validar o endpoint correto antes de remover participantes."
                )
                titulo = ""
                permissoes = normalizar_permissoes({chave: False for chave in CHAVES_PERMISSOES})
            else:
                titulo = solicitar_titulo()
                permissoes = solicitar_permissoes()

            if modo_usuarios == "lote_xlsx":
                operacoes = montar_operacoes_lote_com_aba_unica(
                    page,
                    projetos_selecionados,
                    emails,
                    acao_usuario,
                    titulo,
                    permissoes,
                )
                exibir_previa_lote(operacoes, emails)
            else:
                operacoes = montar_operacoes_com_aba_unica(
                    page,
                    projetos_selecionados,
                    emails[0],
                    acao_usuario,
                    titulo,
                    permissoes,
                )
                for op in operacoes:
                    op["email"] = emails[0]
                exibir_previa(operacoes, emails[0])

            if confirmar_aplicacao_sn(operacoes):
                if modo_usuarios == "lote_xlsx":
                    resultados = aplicar_operacoes_lote_com_aba_unica(page, operacoes)
                else:
                    resultados = aplicar_operacoes_com_aba_unica(page, operacoes, emails[0])
            else:
                print("\n[OK] Execucao cancelada pelo usuario. Nenhuma alteracao foi aplicada.")
                resultados = [
                    {
                        "projectId": op.get("projectId"),
                        "nomeProjeto": op.get("nomeProjeto"),
                        "origemProjectWise": op.get("origemProjectWise"),
                        "criterioCruzamento": op.get("criterioCruzamento"),
                        "email": op.get("email") or emails[0],
                        "acao_solicitada": op.get("acaoSolicitada"),
                        "acao_planejada": op.get("acaoEfetiva") or op.get("status"),
                        "situacao_usuario": op.get("status"),
                        "resultado": {
                            "status": "cancelado_pelo_usuario",
                            "mensagem": "Usuario nao confirmou a aplicacao final.",
                        },
                    }
                    for op in operacoes
                ]

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
            apagar_json_temporario(arquivo_projetos_execucao)


if __name__ == "__main__":
    main()
