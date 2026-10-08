import logging
import re
from pathlib import Path
from urllib.parse import quote, urlsplit
import requests
from graph_client import GraphClient

def download_file(graph: GraphClient, drive_id: str, item: dict, destination: Path) -> Path:
    metadata = graph.get(f"drives/{quote(drive_id, safe='')}/items/{quote(item['id'], safe='')}")
    url = metadata.get("@microsoft.graph.downloadUrl", "")
    if urlsplit(url).scheme != "https":
        raise ValueError("Graph nao forneceu URL HTTPS de download.")
    name = item["name"]
    if name in (".", "..") or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((".", " ")) or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", name, re.I):
        raise ValueError("Nome incompat?vel com Windows. Renomeacao automatica ficara para a segunda fase.")
    destination.mkdir(parents=True, exist_ok=True)
    target = destination / name
    if target.exists():
        raise ValueError("Arquivo de destino ja existe; escolha outra pasta para o teste.")
    partial = target.with_name(target.name + ".part")
    logging.info("Download iniciado: %s", name)
    written = 0
    try:
        # Separate request: never forward the Graph bearer token to the download host.
        with partial.open("xb") as output:
            with requests.get(url, stream=True, timeout=(15, 120)) as response:
                if response.status_code != 200:
                    raise RuntimeError(f"Download recusado: HTTP {response.status_code}. Execute o teste novamente.")
                for chunk in response.iter_content(1024 * 1024):
                    output.write(chunk)
                    written += len(chunk)
        if written != metadata["size"]:
            raise RuntimeError("Tamanho baixado diverge do tamanho informado pelo Graph.")
        partial.rename(target)
    except requests.RequestException:
        raise RuntimeError("Falha de rede no download. Arquivo parcial preservado para diagnostico; este POC nao retoma downloads.") from None
    logging.info("Download concluido: %s bytes=%s", name, written)
    return target
