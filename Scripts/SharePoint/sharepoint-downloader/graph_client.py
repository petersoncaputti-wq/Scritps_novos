import logging
import time
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from urllib.parse import urlsplit
import requests

class GraphError(RuntimeError):
    def __init__(self, status: int, code: str):
        self.status = status
        messages = {401: "Sessao expirada. Execute novamente para autenticar.",
            403: "Microsoft Graph recusou o acesso. Verifique permissoes, consentimento e acesso do usuario.",
            404: "Site, biblioteca ou pasta nao encontrado."}
        super().__init__(messages.get(status, "Falha na chamada Microsoft Graph.") + f" HTTP {status}; codigo {code}.")

class GraphClient:
    def __init__(self, token: str):
        self.session = requests.Session()
        self.session.headers["Authorization"] = f"Bearer {token}"

    def get(self, path: str) -> dict:
        url = path if path.startswith("https://") else "https://graph.microsoft.com/v1.0/" + path.lstrip("/")
        parsed = urlsplit(url)
        if parsed.scheme != "https" or parsed.netloc != "graph.microsoft.com":
            raise ValueError("URL Graph fora do dominio permitido.")
        for attempt in range(6):
            try:
                response = self.session.get(url, timeout=(15, 90), allow_redirects=False)
            except requests.RequestException:
                logging.warning("Falha de rede Graph tentativa=%s caminho=%s", attempt + 1, parsed.path)
                if attempt == 5:
                    raise RuntimeError("Falha de rede ao acessar Microsoft Graph.") from None
                time.sleep(2 ** attempt)
                continue
            if response.ok:
                return response.json()
            try:
                code = response.json().get("error", {}).get("code", "unknown")
            except ValueError:
                code = "unknown"
            logging.warning("Graph falhou caminho=%s HTTP=%s codigo=%s request-id=%s", parsed.path, response.status_code, code, response.headers.get("request-id"))
            if (response.status_code != 429 and not 500 <= response.status_code <= 599) or attempt == 5:
                raise GraphError(response.status_code, code)
            delay = 2 ** attempt
            retry = response.headers.get("Retry-After")
            if retry:
                try:
                    delay = max(0, float(retry))
                except ValueError:
                    try:
                        delay = max(0, (parsedate_to_datetime(retry) - datetime.now(timezone.utc)).total_seconds())
                    except (ValueError, TypeError):
                        pass
            print(f"Microsoft Graph esta limitando ou adiando requisicoes. Aguardando {delay:.0f}s.")
            time.sleep(delay)
        raise RuntimeError("Limite de tentativas atingido.")

    def pages(self, path: str):
        while path:
            data = self.get(path)
            yield from data.get("value", [])
            path = data.get("@odata.nextLink", "")
