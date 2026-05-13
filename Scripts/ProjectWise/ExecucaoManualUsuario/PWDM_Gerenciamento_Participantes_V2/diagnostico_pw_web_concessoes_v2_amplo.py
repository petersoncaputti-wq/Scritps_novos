import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import BrowserContext, Page, Request, Response, sync_playwright


PASTA_LOGS = Path("Logs")
URL_INICIAL = "https://connect.bentley.com"
TIPOS_RECURSO = {"document", "xhr", "fetch"}
DOMINIOS_PERMITIDOS = (
    "bentley.com",
    "bentley",
    "projectwise",
)


def nome_execucao() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def preparar_logs() -> None:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)


def texto_limitado(texto: str, limite: int = 5000) -> str:
    if len(texto) <= limite:
        return texto
    return texto[:limite] + "...[truncado]"


def dominio_permitido(url: str) -> bool:
    url_lower = url.lower()
    return any(dominio in url_lower for dominio in DOMINIOS_PERMITIDOS)


def deve_capturar_request(request: Request) -> bool:
    return request.resource_type in TIPOS_RECURSO and dominio_permitido(request.url)


def headers_seguros(headers: dict[str, str]) -> dict[str, str]:
    seguros: dict[str, str] = {}
    for chave, valor in headers.items():
        chave_lower = chave.lower()
        if chave_lower in {
            "authorization",
            "cookie",
            "set-cookie",
            "x-csrf-token",
            "requestverificationtoken",
            "__requestverificationtoken",
        }:
            seguros[chave] = "[removido]"
        else:
            seguros[chave] = valor
    return seguros


def reduzir_json(valor: Any, profundidade: int = 0) -> Any:
    if profundidade >= 5:
        return "[profundidade-limitada]"

    if isinstance(valor, dict):
        reduzido: dict[str, Any] = {}
        for indice, (chave, item) in enumerate(valor.items()):
            if indice >= 40:
                reduzido["..."] = f"{len(valor) - indice} chave(s) omitida(s)"
                break
            reduzido[chave] = reduzir_json(item, profundidade + 1)
        return reduzido

    if isinstance(valor, list):
        return [reduzir_json(item, profundidade + 1) for item in valor[:10]]

    if isinstance(valor, str):
        if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", valor):
            return "[email-removido]"
        return texto_limitado(valor, 250)

    return valor


def extrair_json_preview(texto: str) -> Any:
    try:
        return reduzir_json(json.loads(texto))
    except Exception:
        return None


def registrar_request(request: Request, chamadas: list[dict[str, Any]]) -> None:
    if not deve_capturar_request(request):
        return

    item = {
        "tipoEvento": "request",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "url": request.url,
        "method": request.method,
        "resourceType": request.resource_type,
        "headers": headers_seguros(request.headers),
        "postDataPreview": texto_limitado(request.post_data or "", 1500),
    }
    chamadas.append(item)
    print(f"[REQ] {request.resource_type} {request.method} {request.url}")


def registrar_response(response: Response, chamadas: list[dict[str, Any]]) -> None:
    request = response.request
    if not deve_capturar_request(request):
        return

    item: dict[str, Any] = {
        "tipoEvento": "response",
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "url": response.url,
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
            json_preview = extrair_json_preview(texto)
            if json_preview is not None:
                item["jsonPreview"] = json_preview
        except Exception as erro:
            item["bodyErro"] = str(erro)

    chamadas.append(item)
    print(f"[RES] HTTP {response.status} {request.resource_type} {request.method} {response.url}")


def configurar_pagina(page: Page, chamadas: list[dict[str, Any]]) -> None:
    page.on("request", lambda request: registrar_request(request, chamadas))
    page.on("response", lambda response: registrar_response(response, chamadas))


def configurar_contexto(context: BrowserContext, chamadas: list[dict[str, Any]]) -> None:
    context.on("page", lambda page: configurar_pagina(page, chamadas))
    context.on("request", lambda request: registrar_request(request, chamadas))
    context.on("response", lambda response: registrar_response(response, chamadas))


def main() -> None:
    preparar_logs()
    chamadas: list[dict[str, Any]] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False, slow_mo=80)
        context = browser.new_context(viewport={"width": 1440, "height": 900})
        configurar_contexto(context, chamadas)
        page = context.new_page()
        configurar_pagina(page, chamadas)

        print("Diagnostico amplo PW Web - concessoes e projetos")
        print("Capturei todas as abas criadas pelo navegador.")
        print("Navegue no PW Web ate a arvore de concessoes/projetos e expanda alguns itens.")
        print("Depois volte aqui e pressione ENTER para salvar.")

        page.goto(URL_INICIAL, wait_until="domcontentloaded")
        input("\nPressione ENTER depois de navegar/expandir concessoes e projetos...")

        arquivo = PASTA_LOGS / f"pw_web_concessoes_diagnostico_amplo_{nome_execucao()}.json"
        with arquivo.open("w", encoding="utf-8") as saida:
            json.dump(
                {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "totalEventos": len(chamadas),
                    "eventos": chamadas,
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
