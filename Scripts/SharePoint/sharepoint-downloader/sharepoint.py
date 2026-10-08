import logging
from dataclasses import dataclass
from urllib.parse import urlsplit, parse_qs, unquote, quote
from graph_client import GraphClient, GraphError

@dataclass(frozen=True)
class SharePointURL:
    hostname: str
    path: str

def parse_sharepoint_url(url: str) -> SharePointURL:
    parts = urlsplit(url.strip())
    if parts.scheme != "https" or not parts.hostname or not parts.hostname.endswith(".sharepoint.com") or parts.username or parts.port:
        raise ValueError("Informe uma URL HTTPS de SharePoint Online (*.sharepoint.com).")
    query = parse_qs(parts.query)
    path = query.get("id", query.get("RootFolder", [unquote(parts.path)]))[0]
    if not path.startswith("/") or path.startswith("//"):
        raise ValueError("Caminho remoto invalido.")
    if path.startswith("/:") or ("/Forms/" in path and not query.get("id") and not query.get("RootFolder")):
        raise ValueError("Use a URL direta da pasta ou a URL de visualizacao contendo id/RootFolder; links de compartilhamento nao sao suportados neste POC.")
    return SharePointURL(parts.hostname.lower(), path.rstrip("/"))

def get_site(graph: GraphClient, parsed: SharePointURL) -> dict:
    # Probe from deepest path to root: handles subsites without assuming /sites or /teams.
    segments = parsed.path.strip("/").split("/") if parsed.path else []
    for length in range(len(segments), -1, -1):
        path = "/" + "/".join(segments[:length])
        try:
            site = graph.get(f"sites/{parsed.hostname}:{quote(path, safe='/')}")
            logging.info("Site encontrado: %s Site ID=%s", site.get("displayName"), site["id"])
            return site
        except GraphError as error:
            if error.status != 404:
                raise
    raise ValueError("Site nao encontrado para a URL informada.")

def get_drives(graph: GraphClient, site_id: str) -> list[dict]:
    return list(graph.pages(f"sites/{quote(site_id, safe='')}/drives"))

def resolve_sharepoint_folder(graph: GraphClient, parsed: SharePointURL, drives: list[dict]) -> tuple[dict, dict]:
    candidates = []
    for drive in drives:
        url = urlsplit(drive.get("webUrl", ""))
        root = unquote(url.path).rstrip("/")
        if url.hostname == parsed.hostname and (parsed.path.casefold() == root.casefold() or parsed.path.casefold().startswith(root.casefold() + "/")):
            candidates.append((len(root), root, drive))
    if not candidates:
        raise ValueError("Nenhuma biblioteca corresponde a URL. Verifique a URL direta da pasta.")
    _, root, drive = max(candidates, key=lambda item: item[0])
    relative = parsed.path[len(root):].strip("/")
    endpoint = f"drives/{quote(drive['id'], safe='')}/root"
    if relative:
        endpoint += ":/" + quote(relative, safe="/")
    folder = graph.get(endpoint)
    if "folder" not in folder:
        raise ValueError("A URL aponta para um arquivo; informe uma pasta.")
    logging.info("Biblioteca=%s Drive ID=%s caminho=%s", drive.get("name"), drive["id"], relative or "/")
    return drive, folder

def list_folder(graph: GraphClient, drive_id: str, folder_id: str) -> list[dict]:
    items = list(graph.pages(f"drives/{quote(drive_id, safe='')}/items/{quote(folder_id, safe='')}/children?$top=200"))
    logging.info("Itens=%s arquivos=%s", len(items), sum("file" in item for item in items))
    return items
