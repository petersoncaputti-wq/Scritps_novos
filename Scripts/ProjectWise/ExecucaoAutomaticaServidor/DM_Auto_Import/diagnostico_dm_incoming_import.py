import json
import re
from datetime import datetime
from typing import Any
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

from playwright.sync_api import Error, sync_playwright

from config import (
    BROWSER_PROFILE_DIR,
    CONNECT_URL,
    DOMINIOS_RELEVANTES,
    LOGS_DIR,
    OBJETIVO_DIAGNOSTICO,
    TERMOS_RELEVANTES_DM,
    TIPOS_RECURSO_CAPTURADOS,
    USAR_PERFIL_PERSISTENTE,
)


HEADERS_SENSIVEIS = {
    "authorization",
    "cookie",
    "set-cookie",
    "x-csrf-token",
    "x-xsrf-token",
    "x-requestverificationtoken",
    "requestverificationtoken",
}

CHAVES_SENSIVEIS = {
    "access_token",
    "auth",
    "authorization",
    "cookie",
    "email",
    "id_token",
    "jwt",
    "mail",
    "password",
    "refresh_token",
    "requestverificationtoken",
    "set-cookie",
    "token",
    "x-csrf-token",
    "x-requestverificationtoken",
    "x-xsrf-token",
}

PADRAO_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", re.IGNORECASE)
PADRAO_BEARER = re.compile(r"\bBearer\s+[A-Za-z0-9._~+/=-]+", re.IGNORECASE)
PADRAO_TOKEN_QUERY = re.compile(
    r"(?i)(access_token|id_token|refresh_token|token|jwt|code|requestverificationtoken)=([^&#\s]+)"
)


def mascarar_texto(valor: Any) -> Any:
    if not isinstance(valor, str):
        return valor

    valor = PADRAO_EMAIL.sub("[EMAIL_MASCARADO]", valor)
    valor = PADRAO_BEARER.sub("Bearer [TOKEN_MASCARADO]", valor)
    valor = PADRAO_TOKEN_QUERY.sub(r"\1=[TOKEN_MASCARADO]", valor)
    return valor


def chave_sensivel(chave: str) -> bool:
    chave_normalizada = chave.lower()
    return any(termo in chave_normalizada for termo in CHAVES_SENSIVEIS)


def mascarar_estrutura(valor: Any) -> Any:
    if isinstance(valor, dict):
        return {
            chave: "[MASCARADO]" if chave_sensivel(str(chave)) else mascarar_estrutura(item)
            for chave, item in valor.items()
        }

    if isinstance(valor, list):
        return [mascarar_estrutura(item) for item in valor]

    return mascarar_texto(valor)


def mascarar_headers(headers: dict[str, str]) -> dict[str, str]:
    return {
        chave: "[MASCARADO]" if chave.lower() in HEADERS_SENSIVEIS else mascarar_texto(valor)
        for chave, valor in headers.items()
    }


def mascarar_url(url: str) -> str:
    partes = urlsplit(url)
    query_segura = []

    for chave, valor in parse_qsl(partes.query, keep_blank_values=True):
        if chave_sensivel(chave):
            query_segura.append((chave, "[MASCARADO]"))
        else:
            query_segura.append((chave, mascarar_texto(valor)))

    return urlunsplit(
        (
            partes.scheme,
            partes.netloc,
            mascarar_texto(partes.path),
            urlencode(query_segura, doseq=True),
            mascarar_texto(partes.fragment),
        )
    )


def mascarar_post_data(post_data: str | None) -> Any:
    if not post_data:
        return None

    try:
        return mascarar_estrutura(json.loads(post_data))
    except json.JSONDecodeError:
        return mascarar_texto(post_data)


def dominio_relevante(url: str) -> bool:
    url_lower = url.lower()
    return any(dominio in url_lower for dominio in DOMINIOS_RELEVANTES)


def contem_termo_dm(*valores: str | None) -> bool:
    texto = " ".join(valor or "" for valor in valores).lower()
    return any(termo in texto for termo in TERMOS_RELEVANTES_DM)


def evento_base(request) -> dict[str, Any]:
    return {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "tipoEvento": "request",
        "metodo": request.method,
        "resourceType": request.resource_type,
        "url": mascarar_url(request.url),
        "headers": mascarar_headers(request.headers),
        "postData": mascarar_post_data(request.post_data),
        "provavelmenteDM": dominio_relevante(request.url)
        and contem_termo_dm(request.url, request.post_data),
    }


def registrar_response(response, eventos: list[dict[str, Any]]) -> None:
    request = response.request

    if request.resource_type not in TIPOS_RECURSO_CAPTURADOS:
        return

    if not dominio_relevante(request.url):
        return

    try:
        headers = response.headers
    except Error:
        headers = {}

    eventos.append(
        {
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "tipoEvento": "response",
            "metodo": request.method,
            "resourceType": request.resource_type,
            "url": mascarar_url(response.url),
            "status": response.status,
            "statusText": response.status_text,
            "headers": mascarar_headers(headers),
            "provavelmenteDM": contem_termo_dm(response.url),
        }
    )


def salvar_diagnostico(eventos: list[dict[str, Any]]) -> str:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    timestamp_arquivo = datetime.now().strftime("%Y%m%d_%H%M%S")
    eventos_dm = [evento for evento in eventos if evento.get("provavelmenteDM")]

    diagnostico = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "objetivo": OBJETIVO_DIAGNOSTICO,
        "observacoes": [
            "Fase 1 apenas diagnostica. Nenhuma importacao automatica e executada pelo script.",
            "GUIDs foram mantidos para permitir identificar projectId/connectSpaceId/submittalId.",
            "Cookies, authorization headers, tokens, request verification tokens e e-mails foram mascarados quando detectados.",
            "Revise o log antes de compartilhar, pois pode haver dados de negocio em URLs, payloads ou nomes de documentos.",
        ],
        "totalEventos": len(eventos),
        "totalEventosProvavelmenteDM": len(eventos_dm),
        "eventosProvavelmenteDM": eventos_dm,
        "eventosCompletos": eventos,
    }

    caminho_log = LOGS_DIR / f"diagnostico_dm_incoming_import_{timestamp_arquivo}.json"
    caminho_log.write_text(
        json.dumps(diagnostico, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return str(caminho_log)


def main() -> None:
    eventos: list[dict[str, Any]] = []

    print("Iniciando diagnostico do ProjectWise Deliverables Management.")
    print("O navegador sera aberto em modo visual.")
    print("Faca login, navegue ate Incoming/Submittals e execute manualmente o fluxo de teste.")
    print("Quando terminar, volte a este terminal e pressione ENTER para salvar o JSON.")

    with sync_playwright() as p:
        if USAR_PERFIL_PERSISTENTE:
            BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(BROWSER_PROFILE_DIR),
                headless=False,
            )
            browser = None
        else:
            browser = p.chromium.launch(headless=False)
            context = browser.new_context()

        page = context.new_page()

        def on_request(request) -> None:
            if request.resource_type not in TIPOS_RECURSO_CAPTURADOS:
                return

            if not dominio_relevante(request.url):
                return

            eventos.append(evento_base(request))

        context.on("request", on_request)
        context.on("response", lambda response: registrar_response(response, eventos))

        page.goto(CONNECT_URL, wait_until="domcontentloaded")
        input("\nPressione ENTER para finalizar a captura e salvar o diagnostico...")

        caminho_log = salvar_diagnostico(eventos)
        print(f"\nDiagnostico salvo em: {caminho_log}")
        print(f"Total de eventos capturados: {len(eventos)}")
        print(
            "Total de eventos provavelmente relacionados ao DM: "
            f"{sum(1 for evento in eventos if evento.get('provavelmenteDM'))}"
        )

        context.close()
        if browser:
            browser.close()


if __name__ == "__main__":
    main()
