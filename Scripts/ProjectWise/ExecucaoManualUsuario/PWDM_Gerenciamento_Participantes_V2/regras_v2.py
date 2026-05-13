import re
import unicodedata
from typing import Any
from urllib.parse import parse_qs, urlsplit


GUID_REGEX = re.compile(
    r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"
)
EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
CHAVES_PERMISSOES = ["canView", "canReceive", "canIssue", "canReviewAnswer", "isAdmin"]
ACOES_USUARIO = {
    "1": "incluir",
    "2": "alterar",
    "3": "excluir",
}
ROTULOS_ACAO_USUARIO = {
    "incluir": "Incluir usuario nos projetos",
    "alterar": "Alterar permissoes de usuario ja participante",
    "excluir": "Excluir usuario dos projetos",
}


def normalizar_permissoes(permissoes: dict[str, bool]) -> dict[str, bool]:
    permissoes = dict(permissoes)

    if any(permissoes.get(chave) for chave in ["canReceive", "canIssue", "canReviewAnswer", "isAdmin"]):
        if not permissoes.get("canView"):
            print("[INFO] Ajuste automatico: qualquer permissao ativa exige 'Exibir pacote'.")
        permissoes["canView"] = True

    if permissoes.get("canIssue") and not permissoes.get("canReceive"):
        print("[INFO] Ajuste automatico: 'Emitir pacote' exige tambem 'Receber pacote' no PWDM.")
        permissoes["canReceive"] = True

    if permissoes.get("canReviewAnswer") and not permissoes.get("canReceive"):
        print("[INFO] Ajuste automatico: 'Aprovar resposta da RFI' exige tambem 'Receber pacote' no PWDM.")
        permissoes["canReceive"] = True

    return permissoes


def normalizar_texto_chave(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "")
    texto = "".join(caractere for caractere in texto if not unicodedata.combining(caractere))
    texto = texto.lower()
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = normalizar_mojibake_connect(texto)
    return re.sub(r"\s+", " ", texto).strip()


def normalizar_mojibake_connect(texto: str) -> str:
    substituicoes = {
        r"\bamplia\s+es\b": "ampliacoes",
        r"\badequa\s+o\b": "adequacao",
        r"\bal\s+as\b": "alcas",
        r"\balian\s+a\b": "alianca",
        r"\babadi\s+nia\b": "abadiania",
        r"\buru a u\b": "uruacu",
        r"\bpra\s+a\b": "praca",
        r"\bped\s+gio\b": "pedagio",
        r"\b(?:pontos?\s+de\s+)?nibus\b": "onibus",
    }
    for padrao, substituto in substituicoes.items():
        texto = re.sub(padrao, substituto, texto)
    return texto


def extrair_ids_de_url(url: str) -> tuple[str, str]:
    url_lower = url.lower()
    if (
        "pwdm.bentley.com" not in url_lower
        and "connect.bentley.com/project/index" not in url_lower
        and "connect.bentley.com/project/" not in url_lower
    ):
        raise ValueError(f"URL nao parece ser do PWDM/Connect: {url}")

    partes = urlsplit(url)
    parametros = parse_qs(partes.query)
    project_id_param = parametros.get("projectId") or parametros.get("projectid")
    if project_id_param:
        ids_parametro = GUID_REGEX.findall(project_id_param[0])
        if ids_parametro:
            return ids_parametro[0], ids_parametro[0]

    match_settings = re.search(
        rf"/({GUID_REGEX.pattern})/ProjectSettings/({GUID_REGEX.pattern})/",
        url,
        flags=re.IGNORECASE,
    )
    if match_settings:
        return match_settings.group(1), match_settings.group(2)

    ids = list(dict.fromkeys(GUID_REGEX.findall(url)))
    if len(ids) >= 2:
        return ids[0], ids[1]
    if len(ids) == 1:
        return ids[0], ids[0]

    raise ValueError(f"Nao encontrei GUID de projeto na URL: {url}")


def determinar_acao_efetiva(acao_usuario: str, status_usuario: str) -> tuple[str, str]:
    if acao_usuario == "incluir":
        if status_usuario == "participante":
            return "atualizar_participante", "ATUALIZAR participante ja existente"
        if status_usuario == "candidato":
            return "adicionar_existente", "ADICIONAR participante existente no ProjectWise"
        return "adicionar_por_email", "ADICIONAR por e-mail validado"

    if acao_usuario == "alterar":
        if status_usuario == "participante":
            return "atualizar_participante", "ALTERAR permissoes do participante"
        return "ignorar_usuario_ausente", "IGNORAR: usuario nao esta nos participantes do projeto"

    if acao_usuario == "excluir":
        if status_usuario == "participante":
            return "excluir_participante", "EXCLUIR participante"
        return "ignorar_usuario_ausente", "IGNORAR: usuario nao esta nos participantes do projeto"

    raise ValueError(f"Acao desconhecida: {acao_usuario}")


def descricao_permissoes(permissoes: dict[str, bool]) -> str:
    nomes = []
    if permissoes.get("canView"):
        nomes.append("Exibir")
    if permissoes.get("canReceive"):
        nomes.append("Receber")
    if permissoes.get("canIssue"):
        nomes.append("Emitir")
    if permissoes.get("canReviewAnswer"):
        nomes.append("Aprovar RFI")
    if permissoes.get("isAdmin"):
        nomes.append("Admin")
    return ", ".join(nomes) or "sem permissoes"


def mascarar_email(valor: str) -> str:
    partes = valor.split("@", 1)
    if len(partes) != 2:
        return "***"

    usuario, dominio = partes
    prefixo = usuario[:2] if len(usuario) > 2 else usuario[:1]
    return f"{prefixo}***@{dominio}"


def permissoes_iguais(membro: dict[str, Any], permissoes: dict[str, bool]) -> bool:
    return all(bool(membro.get(chave)) == bool(permissoes.get(chave)) for chave in CHAVES_PERMISSOES)


def titulo_igual(membro: dict[str, Any], titulo: str) -> bool:
    titulo_atual = membro.get("roleTitle") or membro.get("roleLabel") or ""
    return normalizar_texto_chave(str(titulo_atual)) == normalizar_texto_chave(titulo)
