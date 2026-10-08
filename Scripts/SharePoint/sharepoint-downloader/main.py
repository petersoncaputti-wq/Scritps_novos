import argparse
import logging
from pathlib import Path
from config import BASE, load_config

def main() -> int:
    parser = argparse.ArgumentParser(description="SharePoint Downloader POC")
    parser.add_argument("--device-code", action="store_true", help="Usar Device Code explicitamente, se permitido pelo tenant")
    parser.add_argument("--browser-test", action="store_true", help="Teste com login manual no Edge, sem App Registration")
    parser.add_argument("--browser-all", action="store_true", help="Baixar toda a pasta e subpastas usando login no Edge")
    parser.add_argument("--diagnose-download", action="store_true", help="Comparar metadados e rotas de download de um arquivo com falha")
    parser.add_argument("--diagnose-file", default="", help="Caminho remoto do arquivo; por padrao usa a ultima falha registrada")
    parser.add_argument("--url", default="", help="URL da pasta para o teste pelo navegador")
    parser.add_argument("--destination", default="", help="Pasta local para o teste pelo navegador")
    parser.add_argument("--workers", type=int, choices=range(1, 9), default=3, help="Downloads simultaneos (padrao: 3)")
    parser.add_argument("--refresh-inventory", action="store_true", help="Reler o SharePoint em vez de reutilizar o inventario concluido")
    args = parser.parse_args()
    from log_setup import configure_logging
    run_log, handlers = configure_logging(BASE / 'logs')
    print(f"Log completo desta execucao: {run_log}")
    logging.info("APLICACAO_INICIO modo=%s workers=%s atualizar_inventario=%s", 'diagnostico' if args.diagnose_download else 'browser-all' if args.browser_all else 'browser-test' if args.browser_test else 'graph', args.workers, args.refresh_inventory)
    try:
        if args.diagnose_download:
            from download_diagnostic import choose_target
            args.url, args.diagnose_file = choose_target(args.url, args.diagnose_file)
        if args.browser_test or args.browser_all or args.diagnose_download:
            if args.device_code:
                raise ValueError("Escolha --browser-test ou --device-code.")
            from browser_test import run_browser_test
            return run_browser_test(args.url, args.destination, download_all=args.browser_all,
                                    workers=args.workers, refresh_inventory=args.refresh_inventory,
                                    diagnose_file=args.diagnose_file if args.diagnose_download else '')
        from auth import authenticate
        from graph_client import GraphClient
        from sharepoint import parse_sharepoint_url, get_site, get_drives, resolve_sharepoint_folder, list_folder
        from downloader import download_file
        config = load_config()
        print("SharePoint Downloader POC")
        parsed = parse_sharepoint_url(config.sharepoint_url or input("URL do SharePoint: ").strip())
        destination_text = config.destination or input("Destino: ").strip()
        if not destination_text:
            raise ValueError("Informe um diretorio de destino.")
        destination = Path(destination_text).expanduser().resolve()
        graph = GraphClient(authenticate(config, args.device_code))
        # Simple Graph smoke check using the same read-only scope, before folder resolution.
        site = get_site(graph, parsed)
        logging.info("Chamada inicial Microsoft Graph concluida")
        drive, folder = resolve_sharepoint_folder(graph, parsed, get_drives(graph, site["id"]))
        items = list_folder(graph, drive["id"], folder["id"])
        files = [item for item in items if "file" in item]
        print(f"Site: {site.get('displayName')}\nBiblioteca: {drive['name']}\nPasta: {folder['name']}\nArquivos encontrados: {len(files)} (diretos, sem recursao)")
        for index, item in enumerate(files[:20], 1):
            print(f"{index}. {item['name']} ({item['size']} bytes)")
        if files and input("Deseja baixar um arquivo para testar? [S/N] ").strip().upper() == "S":
            index = int(input(f"Numero do arquivo [1-{min(20, len(files))}]: "))
            if not 1 <= index <= min(20, len(files)):
                raise ValueError("Numero de arquivo invalido.")
            print(f"Arquivo salvo: {download_file(graph, drive['id'], files[index-1], destination)}")
        return 0
    except (KeyboardInterrupt, EOFError):
        logging.warning("APLICACAO_CANCELADA pelo_usuario=true")
        print("Execucao cancelada.")
        return 130
    except Exception as error:
        # Never log raw library exceptions: they may contain signed URLs or tokens.
        logging.error("Execucao interrompida: tipo=%s", type(error).__name__)
        if isinstance(error, (ValueError, RuntimeError)):
            print(f"Erro: {error}")
        else:
            print(f"Falha ({type(error).__name__}). Verifique configuracao, dependencias e conectividade. Consulte logs/sharepoint_downloader.log.")
        return 1
    finally:
        logging.info("APLICACAO_FIM")
        for handler in handlers:
            logging.getLogger().removeHandler(handler)
            handler.close()

if __name__ == "__main__":
    raise SystemExit(main())
