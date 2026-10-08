import json
from dataclasses import dataclass
from pathlib import Path

BASE = Path(__file__).resolve().parent

@dataclass(frozen=True)
class Config:
    client_id: str
    tenant_id: str
    sharepoint_url: str = ""
    destination: str = ""

def load_config() -> Config:
    path = BASE / "config.json"
    if not path.exists():
        raise ValueError("Copie config.example.json para config.json e configure Client ID e Tenant ID.")
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    config = Config(**data)
    if not config.client_id or config.client_id == "CLIENT_ID_DO_APP" or not config.tenant_id:
        raise ValueError("Configure Client ID e Tenant ID em config.json.")
    return config
