# SharePoint Downloader POC para Windows

## Baixar projetos de uma planilha por prioridade

Execute `Baixar Projetos da Planilha.bat` e informe o caminho do arquivo XLSX. O preparador instala ou reutiliza o ambiente Python. Selecione o destino local e conclua o login no Edge uma vez. O programa conserva a mesma sessao entre projetos; expiracao ou politicas corporativas podem exigir novo login/MFA.

A primeira aba e usada por padrao: A = ano, B = lote, C = SH, E = prioridade numerica positiva, F = link. O leitor extrai o endereco real do hiperlink, aceita links de pasta `/:f:/r/` e nao executa formulas nem altera a planilha. Resumos e notas sem dados de projeto sao ignorados. Links invalidos ou prioridades ausentes aparecem como pendencias.

Os projetos seguem prioridade crescente (1 primeiro), com ordem de linha nos empates. Cada projeto e baixado com suas subpastas antes do proximo. Destinos usam ano, lote, SH e identificador da origem. Falhas de um projeto sao registradas e a fila continua. Pastas duplicadas na lista sao sinalizadas e nao sao repetidas.

```powershell
# Conferir sem abrir navegador nem baixar arquivos:
.\.venv\Scripts\python.exe main.py --project-list "C:\Listas\Migracao.xlsx" --plan-only
# Iniciar a fila:
.\.venv\Scripts\python.exe main.py --project-list "C:\Listas\Migracao.xlsx" --destination "D:\Migracao"
# Distribuir linhas diferentes entre maquinas:
.\.venv\Scripts\python.exe main.py --project-list "C:\Listas\Migracao.xlsx" --project-rows 2 3 4 5 --destination "D:\Migracao"
# Selecionar prioridades ou outra aba:
.\.venv\Scripts\python.exe main.py --project-list "C:\Listas\Migracao.xlsx" --priority 1 --priority 2 --sheet "Planilha1"
```

Cada maquina precisa de seu proprio login. Para preservar prioridade global entre maquinas, distribua primeiro os projetos de maior prioridade. Os filtros mantem a prioridade dentro da selecao de cada maquina, sem coordenacao automatica entre equipamentos.

Para retomar, repita o comando com a mesma lista e destino. Os projetos sao revisitados e o controle de cada pasta reaproveita arquivos existentes conforme as regras de tamanho/versao do modo completo. Preserve os bancos SQLite e parciais. Use `--refresh-inventory` para reler alteracoes na origem.

Cada execucao grava `fila_projetos_resumo_<id>.json` no destino e atualiza esse arquivo apos cada projeto, registrando linha, prioridade, destino e estado. `concluido` indica lote sem falhas reportadas, nao verificacao criptografica. Pendencias da planilha e dos downloads resultam em codigo de saida 1. CTRL+C interrompe a fila e preserva os parciais. A fila nao realiza importacao no ProjectWise.

Na lista analisada em 19/10/2026, ha 74 projetos com links aceitos e duas pendencias nas linhas 56 e 57 (`L17` na coluna F). Corrija esses links para incluir os dois projetos.

## Baixar a pasta completa pelo navegador

Para outros notebooks, copie esta pasta sem `.venv`, logs e downloads. Dê duplo clique em `Baixar Pasta Completa SharePoint.bat` em qualquer local.
O BAT usa sua propria pasta e chama `Preparar-Downloader.ps1`, que cria o ambiente e instala dependencias no primeiro uso. Requer Windows, Microsoft Edge e internet na instalacao.
Se nao houver Python 3.11+ compativel com Tcl/Tk, tenta instalar Python 3.13 para o usuario atual usando `winget install --id Python.Python.3.13 --exact --source winget --scope user --silent --accept-package-agreements --accept-source-agreements`.
Requer winget (Instalador de Aplicativos do Windows); politicas corporativas podem impedir a instalacao. O script valida Python antes de continuar e procura tambem seu caminho padrao para nao depender de reiniciar o terminal apos a instalacao.
Execucoes seguintes reutilizam o ambiente; alteracoes no requirements.txt acionam nova verificacao. Ambientes movidos identificados pelo marcador ou invalidos sao preservados como `.venv_backup_*` e recriados.
Opcional: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Preparar-Downloader.ps1 -PrepareOnly` prepara sem iniciar login/download.

Execute `Baixar Pasta Completa SharePoint.bat` ou `.\.venv\Scripts\python.exe main.py --browser-all`.
Informe a URL da pasta e selecione o destino na janela visual. Cancelar encerra a execucao.
Faca login/MFA no Edge; a transferencia e feita pelo Python no terminal, por streaming.

- Padrao de 3 downloads simultaneos; ajuste com `--workers 1` a `--workers 8`.
- Percorre todas as subpastas e paginas, incluindo pastas vazias. Nao ha limite de 100 MB.
- O inventario inicial conta os documentos e o volume total. A barra mostra bytes transferidos e arquivos concluidos, somando as transferencias simultaneas.
- Durante os downloads, a barra ocupa uma unica linha atualizada, adaptada a largura do terminal, com contador de erros e estado atual. Mensagens por arquivo ficam no log completo; login adicional e resumo final aparecem separadamente no terminal.
- Um inventario concluido e reutilizado na mesma URL/destino. Para detectar arquivos novos ou alteracoes, use `--refresh-inventory`. Inventarios interrompidos sao refeitos.
- O arquivo `.sharepoint_<identificador>.sqlite` registra inventario e conclusoes. Preserve-o.
- Arquivos existentes no caminho correspondente com o mesmo tamanho sao pulados mesmo sem registro anterior. Isso evita duplicacao, mas tamanho igual nao comprova integridade. Se ha registro de outra versao, ou tamanho diferente, o arquivo e preservado e aparece como conflito.
- A pasta da URL e criada dentro do destino. Se o destino selecionado ja tem o mesmo nome da pasta raiz, seu conteudo e usado diretamente, evitando uma segunda pasta raiz.
- Nomes incompat?veis com Windows ou muito longos recebem adaptacao com sufixo hash.
- Parciais ficam em `.spdownload.part` com identificacao em `.json`. Uma nova tentativa usa HTTP Range e If-Range quando ha ETag forte; valida Content-Range e ETag antes de anexar bytes. Se o servidor nao suporta retomada ou muda a versao, o arquivo recomeca do inicio. Parciais antigos `.part.<identificador>` nao sao importados automaticamente.
- CTRL+C interrompe as transferencias e preserva os parciais identificados. A parada pode aguardar uma leitura HTTP ate seu timeout (60 segundos). Execute novamente com a mesma URL/destino para continuar.
- Cada arquivo recebe ate tres tentativas. HTTP 429/503 reduz o lote para um download e aplica uma pausa global, respeitando Retry-After. Sessao expirada solicita novo login no Edge.
- Se a rota REST por caminho retornar HTTP 400/404, tenta a URL direta e depois download.aspx. Cada alternativa conserva sessao, streaming, retomada e validacao de tamanho; a rota usada e registrada no log. Falhas de autenticacao ou integridade nao disparam troca de rota.
- Um bloqueio impede dois lotes simultaneos no mesmo destino. Cookies ficam apenas em memoria; trabalhadores HTTP nao acessam Playwright ou SQLite.
- Erros aparecem no terminal e no log sem tokens ou URLs assinadas. O resumo distingue baixados, existentes e falhas/conflitos.

Exemplos:

```powershell
.\.venv\Scripts\python.exe main.py --browser-all
.\.venv\Scripts\python.exe main.py --browser-all --refresh-inventory
.\.venv\Scripts\python.exe main.py --browser-all --workers 1 --destination "D:\SharePoint"
```

Para evitar sincronizacao concorrente pelo OneDrive, selecione um destino local fora dele, se disponivel.
Os totais refletem o inventario salvo; nao bloqueiam alteracoes no SharePoint. A copia inclui arquivos atuais, sem historico, permissoes ou metadados de biblioteca.
Verificado com testes locais de concorrencia, inventario, deduplicacao, conflitos, retomada Range e reinicio quando o servidor ignora Range. O desempenho e a retomada real dependem do SharePoint corporativo.

## Teste alternativo com navegador

### Logs completos

Para investigar uma falha individual, execute `.\.venv\Scripts\python.exe main.py --diagnose-download`.
Seleciona automaticamente um arquivo da execucao mais recente que tenha falhas, abre o Edge para login e compara metadados via terminal/navegador e amostras de conteudo por caminho REST, ID, URL direta e download.aspx.
Cada resposta e lida ate 32 KiB, sem salvar documentos; erros seguem a captura limitada. Registros `DIAGNOSTICO_RESULTADO` mostram os resultados. Sucesso de uma amostra nao comprova download completo.
Para escolher outro arquivo, informe `--url` da pasta e `--diagnose-file` com o caminho remoto completo.

Cada execucao cria `logs/execucao_AAAAMMDD_HHMMSS_<id>.log`, sem rotacao, e mostra seu caminho no terminal.
Tambem atualiza `logs/sharepoint_downloader.log`, historico rotativo de 5 MB com tres copias anteriores.
O log individual registra identificador da execucao, horario UTC-03:00, thread, origem/destino, login, inventario de pastas e arquivos, tentativas, respostas HTTP, downloads concluidos, arquivos existentes, conflitos, retomadas, reinicios, limitacao, erros, cancelamento e resumo com duracao.
Durante o lote, registra progresso agregado a cada 30 segundos. Cookies, tokens e URLs assinadas nao sao registrados.
Logs individuais sao preservados sem exclusao automatica; podem ocupar espaco em lotes grandes. Reinicie o programa para ativar mudancas de log em uma execucao ja iniciada.
Erros HTTP de download geram `SHAREPOINT_HTTP_ERRO` com codigo/mensagem OData, tipo de resposta e IDs de correlacao validos. Leitura limitada a 16 KiB; mensagem limitada a 1200 caracteres. URLs, emails e credenciais reconhecidas sao removidos. Corpos brutos, innererror e headers de autenticacao nao sao gravados; para HTML, apenas o titulo da pagina e extraido.

Execute `Testar SharePoint pelo Navegador.bat` ou `.\.venv\Scripts\python.exe main.py --browser-test`.
Este modo usa Playwright e o Microsoft Edge instalado; nao requer config.json, Client ID ou Tenant ID.
Informe a URL direta da pasta (ou Forms/AllItems.aspx com id/RootFolder), faca login e MFA manualmente e pressione ENTER no terminal quando a pasta estiver aberta.
O teste procura o site/subsite e le a pasta pela API REST SharePoint dentro da sessao do navegador.
Lista ate 20 arquivos diretos e permite baixar um arquivo de ate 100 MB, validando o tamanho.
Exibe ate 200 subpastas por nivel; digite `P1`, `P2` etc. para entrar em uma delas antes de escolher o arquivo. Nao percorre subpastas automaticamente.
O destino padrao e `downloads_teste` nesta pasta; use `--destination` para escolher outro.
Nao sobrescreve arquivos. A sessao e temporaria, sem exportacao de cookies ou tokens.
Politicas corporativas podem impedir este fluxo; login, leitura e download ainda precisam ser comprovados no tenant real.
Este modo nao implementa recursao, retomada ou renovacao para transferencias em massa.
Instale a dependencia com `.\.venv\Scripts\python.exe -m pip install -r requirements.txt`.

POC de leitura: autentica, descobre site e bibliotecas, resolve a pasta, lista todos os filhos diretos com paginacao e mostra ate 20 arquivos. Permite baixar somente um arquivo por execucao. Nao foi validado contra seu tenant: a comprovacao de acesso depende do registro e do login real.

## Requisitos e instalacao

Windows, Python 3.11 ou superior, acesso a internet, conta corporativa com leitura na pasta e autorizacao para registrar/usar a aplicacao no tenant. Instale Python pelo https://www.python.org/downloads/windows/ com o Python Launcher. Se `py` nao estiver disponivel, utilize `python` nos comandos correspondentes. Execute PowerShell na pasta deste README:

```powershell
py --version
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item config.example.json config.json
notepad config.json
```

Nao e necessario ativar o ambiente virtual ou alterar ExecutionPolicy.

## Microsoft Entra ID: configuracao exata

1. No https://entra.microsoft.com, abra Entra ID > App registrations > New registration. Se nao tiver autorizacao, solicite o registro ao administrador.
2. Nome: SharePoint Downloader POC. Tipo: Public Client/Desktop; Supported account types: Accounts in this organizational directory only (single tenant), recomendado para este POC corporativo.
3. Em Authentication > Add a platform > Mobile and desktop applications, registre exatamente `http://localhost` como Redirect URI. Nao selecione plataforma Web nem SPA.
4. Em Authentication > Advanced settings, habilite Allow public client flows para permitir o fallback Device Code. Esse fluxo continua sujeito a politicas do tenant. O login principal e interativo no navegador, aberto pelo MSAL, sem automacao.
5. Em Overview, copie Application (client) ID para `client_id` e Directory (tenant) ID para `tenant_id`. Use o GUID do tenant corporativo; `common` so faz sentido com registro e contas compativeis com multitenant.
6. Em API permissions > Add a permission > Microsoft Graph > Delegated permissions, adicione **Sites.Read.All**. Esta e a unica permissao Graph solicitada pelo POC. Nao adicione Application permissions, permissoes de escrita, Client Secret ou certificados. `User.Read` nao e necessaria: a identidade vem dos claims do login, sem chamada `/me`.
7. Sites.Read.All delegated tem AdminConsentRequired = No na referencia Microsoft. Mesmo assim, o tenant pode restringir consentimento de usuarios ou exigir aprovacao administrativa. Nesse caso, um administrador deve avaliar e conceder consentimento. A permissao nao amplia os direitos SharePoint do usuario autenticado.
8. MSAL solicita automaticamente os escopos de identidade `openid`, `profile` e `offline_access`, necessarios ao login/cache; eles nao sao permissoes Application. Politicas de consentimento tambem podem controla-los.

Sites.Read.All e a permissao minima documentada para a descoberta de site usada aqui; permite leitura dos sites aos quais o usuario tem acesso. Para restringir a aplicacao a sites selecionados, uma variante com Sites.Selected e concessoes explicitas por site pode ser avaliada com TI, mudando provisionamento e descoberta. Nao foi implementada neste POC.

Referencias oficiais:

- https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens
- https://learn.microsoft.com/en-us/graph/api/site-get?view=graph-rest-1.0
- https://learn.microsoft.com/en-us/graph/permissions-reference#sitesreadall
- https://github.com/AzureAD/microsoft-authentication-extensions-for-python

## Configuracao e execucao

Preencha `client_id` e `tenant_id` em config.json. Deixe `sharepoint_url` e `destination` vazios para solicitar no terminal ou preencha-os. Em JSON, use `D:\\BackupSharePoint` para representar `D:\BackupSharePoint`.

```powershell
.\.venv\Scripts\python.exe main.py
# Fallback explicito, somente se permitido pelo tenant:
.\.venv\Scripts\python.exe main.py --device-code
```

Alternativa: duplo clique em `Executar SharePoint Downloader.bat`, depois da instalacao. Segue o estilo dos lancadores 02.1 (BAT > PowerShell > Python) e 05 (BAT > Python). Os scripts ProjectWise existentes nao foram alterados e sua autenticacao Playwright nao foi reutilizada.

Use URL direta, por exemplo `https://empresa.sharepoint.com/sites/Projeto/Biblioteca/Pasta`, ou uma URL `Forms/AllItems.aspx?id=...` / `RootFolder=...`. Bibliotecas sao identificadas pela webUrl, sem assumir nome Documents. Subsites sao procurados do caminho mais profundo ate a raiz. Links curtos de compartilhamento `/:f:/...` nao sao suportados; abra a pasta e copie a URL de navegacao.

No teste, escolha um arquivo pequeno entre os primeiros 20. Confirme que site, biblioteca e pasta exibidos correspondem ao esperado. Confira o arquivo local apos o download. A quantidade mostrada conta arquivos diretos; pastas e seus descendentes nao entram nessa contagem.

## Diagnostico e seguranca

Log rotativo: `logs/sharepoint_downloader.log`. Registra inicio, identidade disponivel, IDs, caminho, contagem, falhas HTTP/codigos Graph, request-id e downloads. Nao registra tokens, headers Authorization, senhas ou URLs assinadas de download. Erros brutos de bibliotecas nao sao registrados para evitar vazamento acidental.

Cache MSAL: `%LOCALAPPDATA%\SharePointDownloader\<client_id>\<tenant_id>\token_cache.bin`, criptografado por DPAPI para o usuario Windows; MSAL Extensions coordena acesso com lock. O cache contem tokens, nunca senha. Nao copie nem compartilhe esse diretorio. Para trocar de conta, remova o cache dessa aplicacao com o programa fechado. Havendo varias contas em cache, o programa abre login interativo.

## Troubleshooting

- Python/py nao encontrado: instale Python e reabra o terminal.
- ModuleNotFoundError: instale requirements no ambiente virtual e use seu executavel.
- Configuracao ausente: copie config.example.json; nao coloque credenciais no Git.
- Login recusado/redirect mismatch: confira Client ID, Tenant ID e redirect `http://localhost` em Mobile and desktop applications. Para investigar autenticacao, TI pode consultar os sign-in logs; o log local guarda codigo e correlation ID da resposta MSAL.
- Consentimento/MFA/Conditional Access: solicite avaliacao a TI; nao tente contornar. Device Code pode estar bloqueado e nao corrige consentimento.
- 401: execute novamente; o MSAL tenta renovar usando cache. Renovacao em meio a execucao e futura.
- 403: confira leitura do usuario no SharePoint, Sites.Read.All delegated e consentimento. A descoberta interrompe em 403 em vez de ocultar bloqueios.
- 404: confira hostname, site/subsite e URL real da biblioteca/pasta.
- 429: respeita Retry-After numerico ou data HTTP. 5xx e falhas de rede Graph usam backoff, ate seis tentativas.
- Falha no download: este POC usa streaming e verifica tamanho, mas nao retoma nem repete automaticamente transferencias. HTTP do download aparece no terminal. Se houver `.part`, examine/remova ou escolha outro destino antes de repetir.
- Arquivo existente: nao sobrescreve. Nomes invalidos/reservados do Windows sao rejeitados nesta fase.
- Caminho longo/disco cheio/acesso negado: use um destino curto com espaco e permissao de escrita; tratamento completo de nomes/caminhos ficara para a proxima fase.

## Arquitetura e proxima fase

config.py valida configuracao; auth.py autentica; graph_client.py centraliza HTTP/paginacao/retry; sharepoint.py resolve URLs/sites/drives/pastas; downloader.py transfere um arquivo; main.py organiza a CLI.

A listagem deste POC guarda filhos diretos em memoria, sem recursao. Uma transferencia de 3,3 TB exigira inventario persistido (por exemplo SQLite), filas limitadas, renovacao de tokens, downloads paralelos moderados, retomada HTTP Range com validacao de identidade/ETag, regras de nomes Windows, verificacao de existentes, relatorio de erros e progresso/velocidade/ETA. URLs assinadas expiram e devem ser obtidas novamente. Arquivos que mudam durante download precisam de validacao adicional; tamanho igual sozinho nao prova integridade. Um backup de biblioteca tambem nao inclui automaticamente versoes, permissoes e metadados SharePoint.

A segunda fase so deve iniciar apos comprovar login, descoberta, listagem e download pequeno no tenant real.

## Verificacao local

```powershell
.\.venv\Scripts\python.exe -m unittest -v test_poc.py
```

Testes simulam Graph/retry/paginacao, resolucao de subsite/biblioteca, download sem envio do bearer ao host de transferencia e cache DPAPI local. Nao substituem a validacao no tenant.
