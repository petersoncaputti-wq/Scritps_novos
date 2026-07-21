import json
from datetime import datetime
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, sync_playwright

from gerenciar_participante_pwdm_v2 import iniciar_navegador, preparar_pasta_logs


PASTA_BASE = Path(__file__).resolve().parent
PASTA_LOGS = PASTA_BASE / "Logs"
URL_RBAC_BASE = "https://connect-rbacportal.bentley.com"
TERMOS_RELEVANTES = (
    "role",
    "roles",
    "user",
    "users",
    "member",
    "members",
    "permission",
    "permissions",
    "rbac",
    "manage",
)


def nome_execucao() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def sanitizar_headers(headers: dict[str, str]) -> dict[str, str]:
    removidos = {"authorization", "cookie", "set-cookie"}
    return {
        chave: ("[removido]" if chave.lower() in removidos else valor)
        for chave, valor in headers.items()
    }


def chamada_relevante(url: str) -> bool:
    url_lower = url.lower()
    return any(termo in url_lower for termo in TERMOS_RELEVANTES)


def corpo_resposta_seguro(response) -> Any:
    url_lower = response.url.lower()
    if "getaccesstoken" in url_lower or "getprivatetoken" in url_lower:
        return "[token removido]"

    content_type = response.headers.get("content-type", "")
    if "application/json" not in content_type.lower():
        return ""
    try:
        texto = response.text()
    except Exception:
        return ""
    if len(texto) > 20000:
        return texto[:20000] + "...[truncado]"
    try:
        return json.loads(texto)
    except Exception:
        return texto


def capturar_chamadas(page: Page, chamadas: list[dict[str, Any]]) -> None:
    def on_response(response):
        try:
            request = response.request
            url = response.url
            if not chamada_relevante(url):
                return
            chamadas.append(
                {
                    "url": url,
                    "method": request.method,
                    "status": response.status,
                    "requestHeaders": sanitizar_headers(request.headers),
                    "responseHeaders": sanitizar_headers(response.headers),
                    "postData": request.post_data or "",
                    "body": corpo_resposta_seguro(response),
                }
            )
        except Exception as erro:
            chamadas.append({"erroCaptura": str(erro)})

    page.on("response", on_response)


def extrair_roles_visiveis(page: Page) -> list[dict[str, str]]:
    return page.evaluate(
        """
        () => {
            const textosIgnorados = new Set([
                "Role",
                "Role(s)",
                "Assign roles *",
                "Assign Owner role (Owner gets full access in this iTwin)",
                "ENTERPRISE"
            ]);

            const elementos = Array.from(document.querySelectorAll("body *"));
            const roles = [];
            for (const elemento of elementos) {
                const texto = (elemento.innerText || elemento.textContent || "").trim();
                if (!texto || texto.length > 80 || textosIgnorados.has(texto)) {
                    continue;
                }
                if (/^(Team Member|Imodel|Project Administrator|PWDM Administrator|PWDM Approver|PWDM Issuer|PWDM Receiver|Owner)$/i.test(texto)) {
                    roles.push({
                        nome: texto,
                        tag: elemento.tagName,
                        ariaLabel: elemento.getAttribute("aria-label") || "",
                        title: elemento.getAttribute("title") || "",
                        dataTestId: elemento.getAttribute("data-testid") || ""
                    });
                }
            }

            const vistos = new Set();
            return roles.filter((role) => {
                const chave = role.nome.toLowerCase();
                if (vistos.has(chave)) {
                    return false;
                }
                vistos.add(chave);
                return true;
            });
        }
        """
    )


def salvar_diagnostico(url: str, chamadas: list[dict[str, Any]], roles: list[dict[str, str]]) -> Path:
    PASTA_LOGS.mkdir(parents=True, exist_ok=True)
    arquivo = PASTA_LOGS / f"rbac_roles_membros_diagnostico_{nome_execucao()}.json"
    dados = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "url": url,
        "rolesVisiveis": roles,
        "chamadas": chamadas,
    }
    with arquivo.open("w", encoding="utf-8") as saida:
        json.dump(dados, saida, indent=2, ensure_ascii=False)
    return arquivo


def main() -> None:
    preparar_pasta_logs()
    url = input(
        "\nURL da tela de membros RBAC "
        "[cole a URL ou ENTER para usar a pagina aberta manualmente]: "
    ).strip()

    with sync_playwright() as playwright:
        browser = None
        try:
            browser, _context, page = iniciar_navegador(playwright)
            chamadas: list[dict[str, Any]] = []
            capturar_chamadas(page, chamadas)

            if url:
                page.goto(url, wait_until="domcontentloaded")
            else:
                page.goto(URL_RBAC_BASE, wait_until="domcontentloaded")

            print("\nFaca login se necessario.")
            print("Depois abra a tela Users do projeto, clique em Add e aguarde a lista de roles aparecer.")
            input("Quando o painel de Add user(s) estiver aberto, pressione ENTER...")

            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass

            roles = extrair_roles_visiveis(page)
            arquivo = salvar_diagnostico(page.url, chamadas, roles)

            print("\nRoles visiveis capturados:")
            for role in roles:
                print(f"- {role['nome']}")
            print(f"\n[OK] Diagnostico salvo em: {arquivo.resolve()}")
            input("Pressione ENTER para fechar o navegador...")
        finally:
            if browser:
                browser.close()


if __name__ == "__main__":
    main()
