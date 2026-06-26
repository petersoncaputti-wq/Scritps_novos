import argparse
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import Request, Response, sync_playwright

from gerenciar_participante_pwdm_v2 import iniciar_navegador


PASTA_BASE = Path(__file__).resolve().parent
PASTA_LOGS = PASTA_BASE / "Logs"
PROJECT_ID_PADRAO = "3887ad62-2f26-423f-bf8e-801139045ca8"
URL_PADRAO = (
    "https://pwdm.bentley.com/"
    f"{PROJECT_ID_PADRAO}/ProjectSettings/{PROJECT_ID_PADRAO}/View#PARTICIPANTS"
)

DOMINIOS_RELEVANTES = (
    "pwdm.bentley.com",
    "infrastructurecloud.bentley.com",
    "connect.bentley.com",
    "bentley.com",
)

TERMOS_RELEVANTES = (
    "participant",
    "participants",
    "user",
    "users",
    "group",
    "groups",
    "member",
    "members",
    "permission",
    "permissions",
    "role",
    "roles",
    "projectsettings",
    "project",
)

HEADERS_SENSIVEIS = {
    "authorization",
    "cookie",
    "set-cookie",
    "x-csrf-token",
    "x-xsrf-token",
    "x-requestverificationtoken",
    "requestverificationtoken",
    "__requestverificationtoken",
}


def nome_execucao() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def texto_limitado(texto: str, limite: int = 4000) -> str:
    if len(texto) <= limite:
        return texto
    return texto[:limite] + "...[truncado]"


def mascarar_texto(texto: str) -> str:
    texto = re.sub(r"[\w.\-+%]+@[\w.\-]+\.[A-Za-z]{2,}", "[email-removido]", texto)
    texto = re.sub(
        r"(__RequestVerificationToken\"?\s*(?:type=\"hidden\"\s*)?value=\")([^\"]+)",
        r"\1[token-removido]",
        texto,
        flags=re.IGNORECASE,
    )
    return texto


def headers_seguros(headers: dict[str, str]) -> dict[str, str]:
    seguros: dict[str, str] = {}
    for chave, valor in headers.items():
        if chave.lower() in HEADERS_SENSIVEIS:
            seguros[chave] = "[removido]"
        else:
            seguros[chave] = texto_limitado(str(valor), 300)
    return seguros


def parece_relevante(url: str) -> bool:
    url_lower = url.lower()
    if not any(dominio in url_lower for dominio in DOMINIOS_RELEVANTES):
        return False
    return any(termo in url_lower for termo in TERMOS_RELEVANTES)


def reduzir_json(valor: Any, profundidade: int = 0) -> Any:
    if profundidade >= 5:
        return "[profundidade-limitada]"

    if isinstance(valor, dict):
        reduzido: dict[str, Any] = {}
        for indice, (chave, item) in enumerate(valor.items()):
            if indice >= 35:
                reduzido["..."] = f"{len(valor) - indice} chave(s) omitida(s)"
                break
            reduzido[str(chave)] = reduzir_json(item, profundidade + 1)
        return reduzido

    if isinstance(valor, list):
        return [reduzir_json(item, profundidade + 1) for item in valor[:8]]

    if isinstance(valor, str):
        return texto_limitado(mascarar_texto(valor), 300)

    return valor


def extrair_json_preview(texto: str) -> Any:
    try:
        return reduzir_json(json.loads(texto))
    except Exception:
        return None


def request_post_data_seguro(request: Request) -> str:
    try:
        texto = request.post_data or ""
    except Exception:
        return ""
    return texto_limitado(mascarar_texto(texto), 1500)


def registrar_response(response: Response, chamadas: list[dict[str, Any]]) -> None:
    request = response.request
    url = response.url
    if not parece_relevante(url):
        return

    content_type = response.headers.get("content-type", "")
    item: dict[str, Any] = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "url": url,
        "method": request.method,
        "resourceType": request.resource_type,
        "status": response.status,
        "contentType": content_type,
        "requestHeaders": headers_seguros(request.headers),
        "responseHeaders": headers_seguros(response.headers),
    }

    post_data = request_post_data_seguro(request)
    if post_data:
        item["postDataPreview"] = post_data

    content_type_lower = content_type.lower()
    if "application/json" in content_type_lower or "text/" in content_type_lower:
        try:
            texto = response.text()
            item["bodyPreview"] = texto_limitado(mascarar_texto(texto))
            json_preview = extrair_json_preview(texto)
            if json_preview is not None:
                item["jsonPreview"] = json_preview
        except Exception as erro:
            item["bodyErro"] = str(erro)

    chamadas.append(item)
    print(f"[CAPTURADO] HTTP {response.status} {request.method} {url}")


def pontuar_candidato(chamada: dict[str, Any]) -> int:
    url = str(chamada.get("url") or "").lower()
    content_type = str(chamada.get("contentType") or "").lower()
    texto = json.dumps(chamada.get("jsonPreview") or chamada.get("bodyPreview") or "", ensure_ascii=False).lower()

    score = 0
    if "application/json" in content_type:
        score += 5
    for termo in ("participant", "participants", "usergroupsandusers", "projectparticipant", "member", "permission"):
        if termo in url:
            score += 4
        if termo in texto:
            score += 2
    if chamada.get("method") in {"POST", "PUT", "PATCH", "DELETE"}:
        score += 3
    if int(chamada.get("status") or 0) in range(200, 300):
        score += 2
    return score


def selecionar_candidatos(chamadas: list[dict[str, Any]]) -> list[dict[str, Any]]:
    candidatos = []
    for chamada in chamadas:
        score = pontuar_candidato(chamada)
        if score >= 7:
            item = {
                "score": score,
                "method": chamada.get("method"),
                "status": chamada.get("status"),
                "contentType": chamada.get("contentType"),
                "url": chamada.get("url"),
                "postDataPreview": chamada.get("postDataPreview", ""),
                "jsonPreview": chamada.get("jsonPreview"),
                "bodyPreview": chamada.get("bodyPreview", ""),
            }
            candidatos.append(item)

    return sorted(candidatos, key=lambda item: item["score"], reverse=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Captura chamadas de rede para localizar endpoints novos de participantes do PWDM/Infrastructure Cloud."
    )
    parser.add_argument("--url", default=URL_PADRAO, help="URL da tela de participantes do projeto.")
    parser.add_argument("--tempo", type=int, default=45, help="Tempo de captura automatica em segundos.")
    args = parser.parse_args()

    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    chamadas: list[dict[str, Any]] = []

    with sync_playwright() as playwright:
        browser, _context, page = iniciar_navegador(playwright)

        print("Diagnostico de endpoint de participantes")
        print("1. O navegador sera aberto no login/portal Bentley.")
        print("2. Faca login normalmente e espere concluir.")
        print("3. Depois pressione ENTER aqui no terminal.")
        print("4. A captura sera limpa e a tela do projeto sera aberta.")
        print("5. Se a aba PARTICIPANTS/PARTICIPANTES nao carregar sozinha, clique nela manualmente.")
        print("6. Se puder, tente abrir o painel de adicionar/editar participante, mas nao confirme nenhuma alteracao.")
        print("7. Volte ao terminal e pressione ENTER para salvar.")
        print(f"\nURL: {args.url}")

        page.goto("https://pwdm.bentley.com", wait_until="domcontentloaded", timeout=90000)
        input("\nDepois de concluir o login no navegador, pressione ENTER para iniciar a captura...")

        chamadas.clear()
        page.on("response", lambda response: registrar_response(response, chamadas))

        page.goto(args.url, wait_until="domcontentloaded", timeout=90000)
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass

        print(f"\nCapturando por ate {args.tempo} segundo(s). Interaja com a tela de participantes.")
        input("Pressione ENTER para salvar o diagnostico...")

        candidatos = selecionar_candidatos(chamadas)
        arquivo = PASTA_LOGS / f"pwdm_endpoint_participantes_diagnostico_{nome_execucao()}.json"
        with arquivo.open("w", encoding="utf-8") as saida:
            json.dump(
                {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "urlInicial": args.url,
                    "totalChamadas": len(chamadas),
                    "totalCandidatos": len(candidatos),
                    "candidatos": candidatos,
                    "chamadas": chamadas,
                },
                saida,
                indent=2,
                ensure_ascii=False,
            )

        print(f"\n[OK] Diagnostico salvo em: {arquivo.resolve()}")
        print(f"[OK] Candidatos encontrados: {len(candidatos)}")
        for candidato in candidatos[:10]:
            print(f"- score={candidato['score']} {candidato['method']} {candidato['status']} {candidato['url']}")

        input("Pressione ENTER para fechar o navegador...")
        browser.close()


if __name__ == "__main__":
    main()
