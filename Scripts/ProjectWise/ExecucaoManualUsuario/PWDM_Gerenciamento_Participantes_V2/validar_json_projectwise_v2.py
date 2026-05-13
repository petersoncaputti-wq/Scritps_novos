import json
import sys
from pathlib import Path
from typing import Any


def carregar_arquivo(caminho: Path) -> dict[str, Any]:
    with caminho.open("r", encoding="utf-8-sig") as entrada:
        dados = json.load(entrada)
    if not isinstance(dados, dict):
        raise ValueError("JSON raiz deve ser um objeto.")
    return dados


def validar(dados: dict[str, Any]) -> list[str]:
    avisos: list[str] = []
    concessoes = dados.get("concessoes")
    if not isinstance(concessoes, list):
        raise ValueError("Campo 'concessoes' nao encontrado ou nao e lista.")

    for indice, concessao in enumerate(concessoes, start=1):
        if not isinstance(concessao, dict):
            avisos.append(f"Concessao {indice}: item nao e objeto.")
            continue

        nome = concessao.get("nome") or concessao.get("nomeTecnico")
        if not nome:
            avisos.append(f"Concessao {indice}: sem nome.")

        projetos = concessao.get("projetos")
        if not isinstance(projetos, list):
            avisos.append(f"Concessao {indice} ({nome}): campo 'projetos' nao e lista.")
            continue

        if not projetos:
            avisos.append(f"Concessao {indice} ({nome}): nenhum projeto encontrado.")

        for indice_projeto, projeto in enumerate(projetos, start=1):
            if not isinstance(projeto, dict):
                avisos.append(f"Concessao {nome}, projeto {indice_projeto}: item nao e objeto.")
                continue
            if not (projeto.get("nome") or projeto.get("nomeTecnico")):
                avisos.append(f"Concessao {nome}, projeto {indice_projeto}: sem nome.")
            if not (projeto.get("id") or projeto.get("guid")):
                avisos.append(f"Concessao {nome}, projeto {indice_projeto}: sem id/guid.")

    return avisos


def main() -> None:
    if len(sys.argv) > 1:
        caminho = Path(sys.argv[1])
    else:
        arquivos = sorted(Path("Logs").glob("pw_concessoes_projetos_v2_*.json"), key=lambda item: item.stat().st_mtime, reverse=True)
        if not arquivos:
            raise SystemExit("Nenhum JSON pw_concessoes_projetos_v2_*.json encontrado em Logs.")
        caminho = arquivos[0]

    dados = carregar_arquivo(caminho)
    avisos = validar(dados)

    concessoes = dados.get("concessoes") or []
    total_projetos = sum(len(item.get("projetos") or []) for item in concessoes if isinstance(item, dict))

    print(f"Arquivo: {caminho.resolve()}")
    print(f"Concessoes: {len(concessoes)}")
    print(f"Projetos: {total_projetos}")

    if avisos:
        print("\nAvisos:")
        for aviso in avisos[:50]:
            print(f"- {aviso}")
        if len(avisos) > 50:
            print(f"- ... {len(avisos) - 50} aviso(s) adicional(is)")
    else:
        print("\n[OK] Estrutura basica valida.")


if __name__ == "__main__":
    main()
