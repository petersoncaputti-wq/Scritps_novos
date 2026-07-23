import json
import msvcrt
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import Response, sync_playwright

from diagnostico_endpoint_participantes_v2 import (
    headers_seguros,
    mascarar_texto,
    request_post_data_seguro,
    texto_limitado,
)
from gerenciar_participante_pwdm_v2 import iniciar_navegador


PASTA_BASE = Path(__file__).resolve().parent
PASTA_LOGS = PASTA_BASE / "Logs"
METODOS_ALTERACAO = {"DELETE", "POST", "PUT", "PATCH"}
TERMOS_EXCLUSAO = (
    "participant",
    "usergroupsandusers",
    "projectparticipant",
    "member",
    "remove",
    "delete",
)
TEMPO_MAXIMO_CAPTURA_SEGUNDOS = 180


def nome_execucao() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def solicitar_url_participantes() -> str:
    print("\nAbra previamente o projeto desejado no PWDM e copie a URL da tela de participantes.")
    while True:
        url = input("URL da tela PARTICIPANTES: ").strip()
        url_lower = url.lower()
        if url_lower.startswith("https://pwdm.bentley.com/") and "projectsettings" in url_lower:
            return url
        print("[AVISO] Informe uma URL do PWDM contendo ProjectSettings.")


def registrar_resposta(response: Response, chamadas: list[dict[str, Any]]) -> None:
    request = response.request
    metodo = request.method.upper()
    url_lower = response.url.lower()
    if "bentley.com" not in url_lower:
        return

    content_type = response.headers.get("content-type", "")
    chamada: dict[str, Any] = {
        "timestamp": datetime.now().isoformat(timespec="milliseconds"),
        "url": mascarar_texto(response.url),
        "method": metodo,
        "resourceType": request.resource_type,
        "status": response.status,
        "statusText": response.status_text,
        "contentType": content_type,
        "requestHeaders": headers_seguros(request.headers),
        "responseHeaders": headers_seguros(response.headers),
    }

    post_data = request_post_data_seguro(request)
    if post_data:
        chamada["postData"] = post_data

    try:
        corpo = response.text()
        if corpo:
            chamada["responseBody"] = texto_limitado(mascarar_texto(corpo), 6000)
    except Exception as erro:
        chamada["responseBodyError"] = str(erro)

    texto_busca = " ".join(
        [
            url_lower,
            str(chamada.get("postData") or "").lower(),
            str(chamada.get("responseBody") or "").lower(),
        ]
    )
    chamada["alteraDados"] = metodo in METODOS_ALTERACAO
    chamada["candidatoExclusao"] = metodo == "DELETE" or (
        chamada["alteraDados"] and any(termo in texto_busca for termo in TERMOS_EXCLUSAO)
    )
    chamadas.append(chamada)

    if chamada["candidatoExclusao"]:
        print(f"[CANDIDATO] HTTP {response.status} {metodo} {response.url}")
    elif chamada["alteraDados"]:
        print(f"[ALTERACAO] HTTP {response.status} {metodo} {response.url}")


def aguardar_exclusao_sem_bloquear_playwright(page) -> None:
    print(
        f"Voce tem ate {TEMPO_MAXIMO_CAPTURA_SEGUNDOS} segundos. "
        "Pressione ENTER depois que o PWDM confirmar a exclusao."
    )
    inicio = time.monotonic()
    while time.monotonic() - inicio < TEMPO_MAXIMO_CAPTURA_SEGUNDOS:
        # Esta chamada bombeia os eventos de rede do Playwright enquanto aguardamos.
        page.wait_for_timeout(200)
        if not msvcrt.kbhit():
            continue

        tecla = msvcrt.getwch()
        if tecla in {"\r", "\n"}:
            # Dá tempo para a resposta final chegar antes de encerrar a coleta.
            page.wait_for_timeout(1500)
            return

        # Descarta a segunda parte de teclas especiais (setas, F1 etc.).
        if tecla in {"\x00", "\xe0"} and msvcrt.kbhit():
            msvcrt.getwch()

    print("\n[AVISO] Tempo maximo de captura atingido; salvando o que foi coletado.")


def main() -> None:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    url_participantes = solicitar_url_participantes()
    chamadas: list[dict[str, Any]] = []

    with sync_playwright() as playwright:
        browser, _context, page = iniciar_navegador(playwright)
        try:
            print("\nDiagnostico de exclusao de participante")
            print("- O diagnostico apenas observa a rede; ele nao exclui usuarios.")
            print("- Faca login e abra a tela informada.")
            print("- Nao execute outra alteracao durante a captura.")

            page.goto(url_participantes, wait_until="domcontentloaded", timeout=90000)
            input("\nQuando a lista de participantes estiver carregada, pressione ENTER para iniciar...")

            # O listener no contexto inclui a pagina principal, pop-ups, novas abas e frames.
            _context.on("response", lambda response: registrar_resposta(response, chamadas))
            print("\n[CAPTURA ATIVA]")
            print("Exclua agora UM participante de teste pela interface do PWDM.")
            aguardar_exclusao_sem_bloquear_playwright(page)

            alteracoes = [item for item in chamadas if item["alteraDados"]]
            candidatos = [item for item in chamadas if item["candidatoExclusao"]]
            arquivo = PASTA_LOGS / f"pwdm_exclusao_participante_diagnostico_{nome_execucao()}.json"
            conteudo = {
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "urlTela": mascarar_texto(url_participantes),
                "observacao": (
                    "Cookies, autorizacao, tokens antiforgery e enderecos de e-mail "
                    "foram mascarados automaticamente."
                ),
                "totalRespostasBentleyCapturadas": len(chamadas),
                "totalAlteracoesCapturadas": len(alteracoes),
                "totalCandidatosExclusao": len(candidatos),
                "candidatosExclusao": candidatos,
                "todasAlteracoesCapturadas": alteracoes,
                "todasRespostasBentleyCapturadas": chamadas,
            }
            arquivo.write_text(
                json.dumps(conteudo, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

            print(f"\n[OK] Diagnostico salvo em: {arquivo.resolve()}")
            print(f"[OK] Respostas Bentley capturadas: {len(chamadas)}")
            print(f"[OK] Alteracoes capturadas: {len(alteracoes)}")
            print(f"[OK] Candidatos de exclusao: {len(candidatos)}")
            for candidato in candidatos:
                print(
                    f"- {candidato['method']} HTTP {candidato['status']} "
                    f"{candidato['url']}"
                )
            input("\nPressione ENTER para fechar o navegador...")
        finally:
            browser.close()


if __name__ == "__main__":
    main()
