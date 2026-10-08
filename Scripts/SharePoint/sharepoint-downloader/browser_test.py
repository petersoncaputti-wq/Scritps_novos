"""Teste de leitura usando login manual e sessao SharePoint no navegador."""
import logging
import re
import time
from pathlib import Path
from urllib.parse import quote, urlsplit

from config import BASE
from sharepoint import parse_sharepoint_url


def folder_endpoint(web: str, path: str) -> str:
    literal = quote(path.replace("'", "''"), safe="/")
    return f"{web}/_api/web/GetFolderByServerRelativePath(decodedUrl='{literal}')"


def select_destination() -> str:
    import tkinter as tk
    from tkinter import filedialog

    window = tk.Tk()
    window.withdraw()
    try:
        window.attributes('-topmost', True)
        selected = filedialog.askdirectory(
            parent=window,
            title='Selecione a pasta onde salvar o download do SharePoint',
            mustexist=True,
        )
    finally:
        window.destroy()
    if not selected:
        raise KeyboardInterrupt
    return selected


def read_json(page, endpoint: str) -> dict:
    result = page.evaluate("""async endpoint => {
        const r = await fetch(endpoint, {credentials: 'same-origin',
            headers: {Accept: 'application/json;odata=nometadata'}});
        const type = r.headers.get('content-type') || '';
        if (!r.ok || !type.includes('json')) return {status: r.status};
        return {status: r.status, data: await r.json()};
    }""", endpoint)
    if "data" not in result:
        raise RuntimeError(f"SharePoint recusou a leitura: HTTP {result['status']}. Confira login e acesso a pasta.")
    return result["data"]


def wait_for_sharepoint(context, hostname: str):
    """Pump browser events after terminal input; login may open another tab."""
    while True:
        answer = input("Quando a pasta estiver aberta, pressione ENTER aqui (Q para cancelar)... ").strip()
        if answer.upper() == "Q":
            raise KeyboardInterrupt
        deadline = time.monotonic() + 30
        while time.monotonic() < deadline:
            pages = [tab for tab in context.pages if not tab.is_closed()]
            if not pages:
                raise RuntimeError("O navegador foi fechado. Execute o teste novamente.")
            # input() blocks Playwright event dispatch. A browser API call lets
            # navigation and new-tab events catch up before inspecting page.url.
            pages[0].wait_for_timeout(250)
            for tab in reversed(context.pages):
                if not tab.is_closed() and urlsplit(tab.url).hostname == hostname:
                    return tab
        hosts = sorted({urlsplit(tab.url).hostname or "pagina sem dominio"
                        for tab in context.pages if not tab.is_closed()})
        print(f"Ainda aguardando o SharePoint: {hostname}")
        print("Dominios abertos: " + ", ".join(hosts))
        print("Conclua o login/MFA na janela Edge aberta pelo teste e tente novamente.")


def run_browser_test(url: str = "", destination: str = "", download_all: bool = False,
                     workers: int = 3, refresh_inventory: bool = False, diagnose_file: str = '') -> int:
    from playwright.sync_api import sync_playwright

    url = url or input("URL da pasta SharePoint: ").strip()
    parsed = parse_sharepoint_url(url)
    logging.info("ORIGEM hostname=%s pasta=%s", parsed.hostname, parsed.path)
    if download_all and not destination:
        print("Selecione a pasta de destino na janela que sera aberta.")
        destination = select_destination()
    target_dir = Path(destination).expanduser().resolve() if destination else BASE / "downloads_teste"
    print(f"Destino local: {target_dir}")
    logging.info("DESTINO_SELECIONADO caminho=%s", target_dir)
    origin = f"https://{parsed.hostname}"
    with sync_playwright() as pw:
        browser = pw.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=120000)
            print("Faca login e MFA no navegador. Abra a pasta desejada.")
            logging.info("LOGIN_AGUARDANDO hostname=%s", parsed.hostname)
            page = wait_for_sharepoint(context, parsed.hostname)
            logging.info("LOGIN_CONCLUIDO hostname=%s", parsed.hostname)
            # Probe web roots, including subsites, without assuming a library name.
            segments = parsed.path.strip("/").split("/")
            web = None
            for count in range(len(segments), -1, -1):
                candidate = origin + ("/" + "/".join(quote(s, safe="") for s in segments[:count]) if count else "")
                try:
                    info = read_json(page, candidate + "/_api/web?$select=Url,Title")
                    web = info["Url"].rstrip("/")
                    break
                except RuntimeError as error:
                    if "HTTP 404" not in str(error):
                        raise
            if not web or urlsplit(web).hostname != parsed.hostname:
                raise RuntimeError("Nao foi possivel localizar o site SharePoint.")
            if diagnose_file:
                from download_diagnostic import diagnose
                return diagnose(page, web, diagnose_file)
            if download_all:
                from browser_bulk import download_tree
                return download_tree(page, web, parsed.path, target_dir, url,
                                     workers=workers, refresh_inventory=refresh_inventory)
            current_path = parsed.path
            while True:
                endpoint = folder_endpoint(web, current_path)
                folder = read_json(page, endpoint + "?$select=Name,ServerRelativeUrl")
                data = read_json(page, endpoint + "/Files?$select=Name,Length,ServerRelativeUrl&$top=20")
                files = data.get("value", [])
                children = read_json(page, endpoint + "/Folders?$select=Name,ServerRelativeUrl&$top=200")
                folders = children.get("value", [])
                print(f"\nPasta: {folder['ServerRelativeUrl']}")
                print(f"Arquivos exibidos: {len(files)} (ate 20); subpastas exibidas: {len(folders)} (ate 200).")
                for number, child in enumerate(folders, 1):
                    print(f"P{number}. {child['Name']}")
                if children.get("odata.nextLink") or children.get("@odata.nextLink"):
                    print("Ha mais subpastas alem das exibidas nesta pagina.")
                if not folders:
                    break
                choice = input("Digite P e o numero para abrir uma subpasta; ENTER para usar esta pasta: ").strip().upper()
                if not choice:
                    break
                if not re.fullmatch(r"P[0-9]+", choice) or not 1 <= int(choice[1:]) <= len(folders):
                    print("Subpasta invalida. Tente novamente.")
                    continue
                child_path = folders[int(choice[1:]) - 1]["ServerRelativeUrl"]
                if not child_path.startswith(folder["ServerRelativeUrl"].rstrip("/") + "/"):
                    raise RuntimeError("Subpasta retornada fora da pasta solicitada.")
                current_path = child_path
            print(f"Site: {info['Title']}\nPasta: {folder['ServerRelativeUrl']}")
            print("Ate 20 arquivos diretos (teste sem recursao):")
            for number, item in enumerate(files, 1):
                print(f"{number}. {item['Name']} ({item['Length']} bytes)")
            logging.info("Teste navegador: leitura concluida arquivos_exibidos=%s", len(files))
            if not files:
                print("Nenhum arquivo direto encontrado nesta pasta. Se houver subpastas, execute novamente e selecione uma delas.")
                return 0
            selection = input("Numero de um arquivo pequeno para baixar (ENTER para encerrar): ").strip()
            if not selection:
                return 0
            index = int(selection)
            if not 1 <= index <= len(files):
                raise ValueError("Numero de arquivo invalido.")
            item = files[index - 1]
            name = item["Name"]
            if name in (".", "..") or re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((".", " ")) or re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)", name, re.I):
                raise ValueError("Nome de arquivo incompatível com Windows.")
            if int(item["Length"]) > 100 * 1024 * 1024:
                raise ValueError("Escolha um arquivo de ate 100 MB para este teste.")
            target_dir.mkdir(parents=True, exist_ok=True)
            target = target_dir / name
            partial = target.with_name(name + ".part")
            if target.exists() or partial.exists():
                raise ValueError("Arquivo de destino ou parcial ja existe. Escolha outro destino.")
            remote = item["ServerRelativeUrl"]
            if not remote.startswith(folder["ServerRelativeUrl"].rstrip("/") + "/"):
                raise RuntimeError("Arquivo retornado fora da pasta solicitada.")
            download_url = web + "/_api/web/GetFileByServerRelativePath(decodedUrl='" + quote(remote.replace("'", "''"), safe="/") + "')/$value"
            page.evaluate("""url => {
                const a = document.createElement('a'); a.id = 'sharepoint-poc-download';
                a.href = url; a.download = ''; a.textContent = 'Download teste';
                document.body.appendChild(a);
            }""", download_url)
            with page.expect_download(timeout=120000) as pending:
                page.locator("#sharepoint-poc-download").click()
            download = pending.value
            if download.failure():
                raise RuntimeError("O navegador informou falha no download.")
            download.save_as(str(partial))
            if partial.stat().st_size != int(item["Length"]):
                raise RuntimeError("Tamanho baixado diverge do informado pelo SharePoint. Parcial preservado.")
            partial.rename(target)
            print(f"Download validado por tamanho: {target}")
            logging.info("Teste navegador: download concluido bytes=%s", target.stat().st_size)
            return 0
        finally:
            context.close()
            browser.close()
