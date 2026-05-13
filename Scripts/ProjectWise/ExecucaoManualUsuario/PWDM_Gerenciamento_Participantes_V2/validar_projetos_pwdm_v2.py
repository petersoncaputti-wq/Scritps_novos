import json
import re
import sys
from pathlib import Path
from typing import Any


GUID_RE = re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")
GUID_ZERO = "00000000-0000-0000-0000-000000000000"


def carregar_json(caminho: Path) -> dict[str, Any]:
    with caminho.open("r", encoding="utf-8-sig") as entrada:
        dados = json.load(entrada)
    if not isinstance(dados, dict):
        raise ValueError("JSON raiz deve ser um objeto.")
    return dados


def validar_projeto(projeto: dict[str, Any], indice: int) -> list[str]:
    avisos: list[str] = []
    nome = projeto.get("nome") or projeto.get("nomeTecnico") or f"projeto {indice}"
    connected_project_id = str(projeto.get("connectedProjectId") or "")
    project_id = str(projeto.get("projectId") or "")
    connect_space_id = str(projeto.get("connectSpaceId") or "")
    url = str(projeto.get("urlParticipantesPwdm") or "")

    if not connected_project_id:
        avisos.append(f"{nome}: sem connectedProjectId.")
        return avisos

    if connected_project_id == GUID_ZERO:
        avisos.append(f"{nome}: connectedProjectId zerado; projeto nao esta associado a um ProjectWise Project/PWDM valido.")
        return avisos

    if not GUID_RE.match(connected_project_id):
        avisos.append(f"{nome}: connectedProjectId nao parece GUID valido: {connected_project_id}")

    if project_id != connected_project_id:
        avisos.append(f"{nome}: projectId diferente de connectedProjectId.")

    if connect_space_id != connected_project_id:
        avisos.append(f"{nome}: connectSpaceId diferente de connectedProjectId.")

    esperado = f"https://pwdm.bentley.com/{connected_project_id}/ProjectSettings/{connected_project_id}/View#PARTICIPANTS"
    if url != esperado:
        avisos.append(f"{nome}: urlParticipantesPwdm diferente do padrao esperado.")

    return avisos


def projeto_tem_connected_project_id_valido(projeto: dict[str, Any]) -> bool:
    connected_project_id = str(projeto.get("connectedProjectId") or "")
    return bool(GUID_RE.match(connected_project_id)) and connected_project_id != GUID_ZERO


def main() -> None:
    if len(sys.argv) > 1:
        caminho = Path(sys.argv[1])
    else:
        arquivos = sorted(
            Path("Logs").glob("pwdm_projetos_selecionados_pw_*.json"),
            key=lambda item: item.stat().st_mtime,
            reverse=True,
        )
        if not arquivos:
            raise SystemExit("Nenhum JSON pwdm_projetos_selecionados_pw_*.json encontrado em Logs.")
        caminho = arquivos[0]

    dados = carregar_json(caminho)
    projetos = dados.get("projetos")
    if not isinstance(projetos, list):
        raise ValueError("Campo 'projetos' nao encontrado ou nao e lista.")

    avisos: list[str] = []
    for indice, projeto in enumerate(projetos, start=1):
        if not isinstance(projeto, dict):
            avisos.append(f"Projeto {indice}: item nao e objeto.")
            continue
        avisos.extend(validar_projeto(projeto, indice))

    aptos = [
        projeto
        for projeto in projetos
        if isinstance(projeto, dict) and projeto_tem_connected_project_id_valido(projeto)
    ]

    print(f"Arquivo: {caminho.resolve()}")
    print(f"Concessao: {(dados.get('concessao') or {}).get('nome')}")
    print(f"Projetos selecionados: {len(projetos)}")
    print(f"Projetos aptos PWDM: {len(aptos)}")
    print(f"Projetos nao aptos PWDM: {len(projetos) - len(aptos)}")

    if avisos:
        print("\nAvisos:")
        for aviso in avisos[:80]:
            print(f"- {aviso}")
        if len(avisos) > 80:
            print(f"- ... {len(avisos) - 80} aviso(s) adicional(is)")
    else:
        print("\n[OK] JSON pronto para integrar ao PWDM V2.")


if __name__ == "__main__":
    main()
