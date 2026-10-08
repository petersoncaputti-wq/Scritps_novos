import logging
import os
from pathlib import Path
import msal
from msal_extensions import FilePersistenceWithDataProtection, PersistedTokenCache
from config import Config

SCOPES = ["https://graph.microsoft.com/Sites.Read.All"]

def authenticate(config: Config, device_code: bool = False) -> str:
    if os.name != "nt":
        raise ValueError("Este POC utiliza cache criptografado pelo Windows DPAPI e requer Windows.")
    cache_dir = Path(os.environ["LOCALAPPDATA"]) / "SharePointDownloader" / config.client_id / config.tenant_id
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache = PersistedTokenCache(FilePersistenceWithDataProtection(str(cache_dir / "token_cache.bin")))
    app = msal.PublicClientApplication(config.client_id,
        authority=f"https://login.microsoftonline.com/{config.tenant_id}", token_cache=cache)
    print("Autenticando com Microsoft...")
    accounts = app.get_accounts()
    account = accounts[0] if len(accounts) == 1 else None
    result = app.acquire_token_silent(SCOPES, account=account) if account else None
    if not result:
        if device_code:
            flow = app.initiate_device_flow(scopes=SCOPES)
            if "user_code" not in flow:
                raise RuntimeError("Nao foi possivel iniciar Device Code. Verifique a configuracao e as politicas do tenant.")
            print(flow["message"])
            result = app.acquire_token_by_device_flow(flow)
        else:
            result = app.acquire_token_interactive(scopes=SCOPES, timeout=300)
    if "access_token" not in result:
        codes = result.get("error_codes", [])
        logging.error("Falha de autenticacao: codigo=%s correlacao=%s", result.get("error"), result.get("correlation_id"))
        if 65001 in codes or 90094 in codes:
            raise RuntimeError("O tenant Microsoft 365 exige consentimento para esta aplicacao. Solicite avaliacao ao administrador.")
        raise RuntimeError("Autenticacao recusada ou cancelada. Verifique App Registration, MFA e politicas do tenant.")
    user = result.get("id_token_claims", {}).get("preferred_username", "indisponivel")
    logging.info("Usuario autenticado: %s", user)
    print("Autenticacao concluida.")
    return result["access_token"]
