from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "Logs"
DATA_DIR = BASE_DIR / "data"
BROWSER_PROFILE_DIR = DATA_DIR / "browser_profile"

CONNECT_URL = "https://connect.bentley.com"

USAR_PERFIL_PERSISTENTE = True

DOMINIOS_RELEVANTES = (
    "bentley.com",
    "projectwise",
    "pwise",
    "pwdm",
    "connect",
)

TERMOS_RELEVANTES_DM = (
    "deliverable",
    "deliverables",
    "submittal",
    "submittals",
    "incoming",
    "acknowledge",
    "acknowledged",
    "import",
    "download",
    "package",
    "packages",
    "transmittal",
    "transmittals",
    "document",
    "documents",
    "projectsettings",
    "workarea",
)

TIPOS_RECURSO_CAPTURADOS = {"document", "xhr", "fetch"}

OBJETIVO_DIAGNOSTICO = (
    "Diagnosticar requests/endpoints/payloads do fluxo Incoming/Acknowledge/Import "
    "do ProjectWise Deliverables Management sem executar importacao automatica."
)
