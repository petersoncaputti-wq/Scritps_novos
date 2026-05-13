import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, Request, Response, sync_playwright


PASTA_LOGS = Path("Logs")
URL_INICIAL = "https://connect.bentley.com"
DOMINIOS_RELEVANTES = (
    "bentley.com",
    "projectwise",
    "wac",
    "pw",
    "connect",
)
TERMOS_RELEVANTES = (
    "folder",
    "folders",
    "project",
    "projects",
    "workarea",
    "workareas",
    "datasource",
    "datasources",
    "document",
    "documents",
    "tree",
    "navigation",
    "children",
    "child",
)


def nome_execucao() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def preparar_logs() -> None:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)


def texto_limitado(texto: str, limite: int = 3000) -> str:
    if len(texto) <= limite:
        return texto
    return texto[:limite] + "...[truncado]"


def parece_relevante(url: str) -> bool:
    url_lower = url.lower()
    if not any(dominio in url_lower for dominio in DOMINIOS_RELEVANTES):
        return False
    return any(termo in url_lower for termo in TERMOS_RELEVANTES)


def headers_seguros(headers: dict[str, str]) -> dict[str, str]:
    seguros: dict[str, str] = {}
    for chave, valor in headers.items():
        chave_lower = chave.lower()
        if chave_lower in {"authorization", "cookie", "set-cookie", "x-csrf-token", "requestverificationtoken"}:
            seguros[chave] = "[removido]"
        else:
            seguros[chave] = valor
    return seguros


def extrair_amostra_json(texto: str) -> Any:
    try:
        dados = json.loads(texto)
    except Exception:
        return None

    return reduzir_json(dados)


def reduzir_json(valor: Any, profundidade: int = 0) -> Any:
    if profundidade >= 4:
        return "[profundidade-limitada]"

    if isinstance(valor, dict):
        reduzido: dict[str, Any] = {}
        for indice, (chave, item) in enumerate(valor.items()):
            if indice >= 25:
                reduzido["..."] = f"{len(valor) - indice} chave(s) omitida(s)"
                break
            reduzido[chave] = reduzir_json(item, profundidade + 1)
        return reduzido

    if isinstance(valor, list):
        return [reduzir_json(item, profundidade + 1) for item in valor[:5]]

    if isinstance(valor, str):
        if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", valor):
            return "[email-removido]"
        return texto_limitado(valor, 200)

    return valor


def registrar_response(response: Response, chamadas: list[dict[str, Any]]) -> None:
    request = response.request
    url = response.url
    if not parece_relevante(url):
        return

    item: dict[str, Any] = {
        "url": url,
        "method": request.method,
        "resourceType": request.resource_type,
        "status": response.status,
        "requestHeaders": headers_seguros(request.headers),
        "responseHeaders": headers_seguros(response.headers),
    }

    content_type = response.headers.get("content-type", "").lower()
    if "application/json" in content_type or "text/" in content_type:
        try:
            texto = response.text()
            item["bodyPreview"] = texto_limitado(texto)
            amostra_json = extrair_amostra_json(texto)
            if amostra_json is not None:
                item["jsonPreview"] = amostra_json
        except Exception as erro:
            item["bodyErro"] = str(erro)

    chamadas.append(item)
    print(f"[CAPTURADO] HTTP {response.status} {request.method} {url}")


def configurar_captura(page: Page, chamadas: list[dict[str, Any]]) -> None:
    def ao_request(request: Request) -> None:
        if parece_relevante(request.url):
            print(f"[REQ] {request.method} {request.url}")

    page.on("request", ao_request)
    page.on("response", lambda response: registrar_response(response, chamadas))


def main() -> None:
    preparar_logs()
    chamadas: list[dict[str, Any]] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False, slow_mo=80)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        page = context.new_page()
        configurar_captura(page, chamadas)

        print("Diagnostico PW Web - concessoes e projetos")
        print("1. O navegador sera aberto no Connect.")
        print("2. Faca login normalmente.")
        print("3. Navegue ate o PW Web e abra a area onde aparecem concessoes/projetos.")
        print("4. Expanda uma concessao e alguns projetos.")
        print("5. Volte ao terminal e pressione ENTER para salvar o diagnostico.")

        page.goto(URL_INICIAL, wait_until="domcontentloaded")
        input("\nDepois de navegar no PW Web e expandir a arvore desejada, pressione ENTER...")

        arquivo = PASTA_LOGS / f"pw_web_concessoes_diagnostico_{nome_execucao()}.json"
        with arquivo.open("w", encoding="utf-8") as saida:
            json.dump(
                {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "totalChamadas": len(chamadas),
                    "chamadas": chamadas,
                },
                saida,
                indent=2,
                ensure_ascii=False,
            )

        print(f"\n[OK] Diagnostico salvo em: {arquivo.resolve()}")
        input("Pressione ENTER para fechar o navegador...")
        browser.close()


if __name__ == "__main__":
    main()
