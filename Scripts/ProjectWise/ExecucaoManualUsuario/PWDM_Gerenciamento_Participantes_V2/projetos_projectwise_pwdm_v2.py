import json
from pathlib import Path
from typing import Any, Optional


PADRAO_PROJETOS_SELECIONADOS = "pwdm_projetos_selecionados_pw_*.json"
GUID_ZERO = "00000000-0000-0000-0000-000000000000"


def connected_project_id_valido(valor: Any) -> bool:
    texto = str(valor or "").strip().lower()
    return bool(texto) and texto != GUID_ZERO


def localizar_json_mais_recente(pasta_logs: Path) -> Path:
    arquivos = sorted(
        pasta_logs.glob(PADRAO_PROJETOS_SELECIONADOS),
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum JSON {PADRAO_PROJETOS_SELECIONADOS} encontrado em {pasta_logs.resolve()}."
        )
    return arquivos[0]


def carregar_json(caminho: Path) -> dict[str, Any]:
    with caminho.open("r", encoding="utf-8-sig") as entrada:
        dados = json.load(entrada)
    if not isinstance(dados, dict):
        raise ValueError("JSON de projetos deve ter objeto na raiz.")
    return dados


def normalizar_projeto_selecionado(projeto: dict[str, Any], concessao: dict[str, Any]) -> dict[str, Any]:
    connected_project_id = str(projeto.get("connectedProjectId") or "").strip()
    if not connected_project_id_valido(connected_project_id):
        raise ValueError(f"Projeto sem connectedProjectId: {projeto.get('nome') or projeto}")

    nome = str(projeto.get("nome") or projeto.get("nomeTecnico") or connected_project_id).strip()
    project_id = str(projeto.get("projectId") or connected_project_id).strip()
    connect_space_id = str(projeto.get("connectSpaceId") or connected_project_id).strip()

    origem_projectwise = {
        "nome": nome,
        "id": str(projeto.get("projectWiseId") or ""),
        "guid": str(projeto.get("projectWiseGuid") or ""),
        "connectedProjectId": connected_project_id,
        "concessao": str((concessao or {}).get("nome") or projeto.get("concessaoProjectWise") or ""),
        "projectWiseWebName": str(projeto.get("projectWiseWebName") or ""),
        "projectWiseWebNumber": str(projeto.get("projectWiseWebNumber") or ""),
        "urlParticipantesPwdm": str(projeto.get("urlParticipantesPwdm") or ""),
    }

    return {
        "nome": nome,
        "numero": str(projeto.get("projectWiseWebNumber") or ""),
        "assetName": str((concessao or {}).get("nome") or projeto.get("concessaoProjectWise") or ""),
        "connectSpaceId": connect_space_id,
        "projectId": project_id,
        "guids": [connected_project_id.lower()],
        "nomeChave": "",
        "origemProjectWise": origem_projectwise,
        "criterioCruzamento": "ProjectWise ConnectedProjectId",
        "urlParticipantesPwdm": str(projeto.get("urlParticipantesPwdm") or ""),
    }


def _descricao_projeto_ignorado(projeto: dict[str, Any]) -> str:
    nome = projeto.get("nome") or projeto.get("nomeTecnico") or "<sem nome>"
    project_wise_id = projeto.get("projectWiseId") or projeto.get("id") or ""
    connected_project_id = projeto.get("connectedProjectId") or ""
    return f"{nome} | PW ID: {project_wise_id} | ConnectedProjectId: {connected_project_id or '<vazio>'}"


def carregar_projetos_selecionados_com_diagnostico(
    caminho: Optional[Path] = None,
    pasta_logs: Optional[Path] = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    if caminho is None:
        if pasta_logs is None:
            pasta_logs = Path("Logs")
        caminho = localizar_json_mais_recente(pasta_logs)

    dados = carregar_json(caminho)
    concessao = dados.get("concessao") or {}
    projetos = dados.get("projetos")
    if not isinstance(projetos, list):
        raise ValueError("Campo 'projetos' nao encontrado ou nao e lista.")

    normalizados: list[dict[str, Any]] = []
    ignorados: list[str] = []
    for projeto in projetos:
        if not isinstance(projeto, dict):
            ignorados.append(str(projeto))
            continue
        if not connected_project_id_valido(projeto.get("connectedProjectId")):
            ignorados.append(_descricao_projeto_ignorado(projeto))
            continue
        normalizados.append(normalizar_projeto_selecionado(projeto, concessao))

    if not normalizados:
        detalhes = "; ".join(ignorados[:5])
        raise ValueError(f"Nenhum projeto selecionado possui connectedProjectId. {detalhes}")

    return normalizados, ignorados


def carregar_projetos_selecionados(caminho: Optional[Path] = None, pasta_logs: Optional[Path] = None) -> list[dict[str, Any]]:
    projetos, _ignorados = carregar_projetos_selecionados_com_diagnostico(caminho, pasta_logs)
    return projetos
