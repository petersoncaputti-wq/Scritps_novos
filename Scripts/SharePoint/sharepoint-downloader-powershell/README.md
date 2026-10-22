# Downloader SharePoint em PowerShell

Versao independente do downloader Python em `../sharepoint-downloader`.
O original nao foi alterado. Login manual/MFA pelo Edge, sem App Registration.
Requer Windows PowerShell 5.1 (presente no Windows), Edge e `msedgedriver.exe`.
Nao requer Python, Excel instalado, Selenium ou modulos extras de PowerShell.
`Transfer.cs` e compilado automaticamente por `Add-Type` usando .NET do Windows;
copie todos os arquivos desta pasta para usar em outra maquina. A execucao em
PowerShell 7 nao foi validada: use os atalhos, que chamam Windows PowerShell 5.1.

## Iniciar

- **Iniciar-Downloader.cmd**: baixar uma pasta completa, incluindo subpastas.
- **Baixar-Projetos.cmd**: abrir seletor visual XLSX e baixar projetos por prioridade.
- **Testar-Navegador.cmd**: navegar pelas subpastas e testar um arquivo de ate 100 MB.
- **Diagnosticar-Download.cmd**: investigar o ultimo arquivo com falha registrado.

Os atalhos preservam a janela para leitura de erros. Informe a URL da pasta,
escolha o destino na janela visual e faca login/MFA no Edge aberto pelo script.
Ao concluir login/MFA, o programa confirma acesso pela API e continua automaticamente.
Nao precisa pressionar ENTER no terminal. Se o acesso nao for confirmado em 10
minutos, encerra a espera com mensagem para conferir MFA/permissoes no Edge.
Na fila por planilha, selecione primeiro o XLSX e depois o destino. O Edge abre
a primeira pasta antes da preparacao do resumo da fila. O terminal mostra a pasta
que esta abrindo; o log registra `NAVEGACAO_INICIO` e `NAVEGACAO_FIM`.
Cancelar a selecao da planilha encerra sem abrir Edge. O atalho usa modo STA para
os dialogos do Windows. A navegacao aguarda o documento inicial (eager), sem
esperar todos os recursos de pagina para prosseguir com a autenticacao.
Se o seletor visual nao funcionar, o destino e solicitado por texto.
Cancelar o seletor permite informar o caminho no terminal; ENTER vazio cancela.

O Edge usa um perfil temporario. Nao importa cookies do perfil pessoal, nao
salva cookies/tokens nos logs e nao guarda login permanente. Politicas do tenant
podem impedir o fluxo ou exigir novo login durante a execucao.

## Pasta completa

```powershell
.\Download-SharePoint.ps1
.\Download-SharePoint.ps1 -Destination 'D:\Migracao' -Workers 3
.\Download-SharePoint.ps1 -Destination 'D:\Migracao' -RefreshInventory
.\Download-SharePoint.ps1 -Destination 'D:\Migracao' -OnlyCurrentFolder
.\Download-SharePoint.ps1 -Destination 'D:\Conferencia' -ListOnly
```

`-Url` evita solicitar a origem. Aceita URL direta, `Forms/AllItems.aspx` com `id`
ou `RootFolder`, e links `/:f:/r/` que contem o caminho. Links opacos nao sao aceitos.
`-Recursive` continua aceito, mas a recursao ja e padrao.

- Inclui todas as paginas, subpastas e pastas vazias, sem limite de 100 MB no lote.
- Cria a pasta raiz da URL dentro do destino. Se o destino ja tem esse nome,
  usa-o diretamente. A versao inicial PowerShell salvava diretamente no destino;
  confira essa diferenca antes de reaproveitar downloads daquela versao.
- Faz inventario antes da transferencia, contando arquivos e bytes.
- Padrao de tres downloads simultaneos; `-Workers` aceita 1 a 8.
- Progresso agregado mostra arquivos, bytes, velocidade estimada, ETA e falhas.
  A estimativa considera bytes logicos, inclusive existentes/parciais reutilizados.
- Nomes Windows invalidos, reservados ou longos recebem sufixo hash. Colisoes
  entre nomes locais interrompem o inventario para evitar perda de arquivos.
- Arquivos existentes de mesmo tamanho sao aproveitados se nao houver registro
  de outra versao. Tamanho igual nao comprova integridade criptografica.
- Tamanho ou versao conhecida diferente vira conflito: preserva arquivo e continua.
- Cada arquivo tem ate tres tentativas. HTTP 429/503 reduz novas transferencias
  para uma e aplica pausa global, respeitando Retry-After numerico ou data HTTP.
  Transferencias ja iniciadas podem terminar durante a pausa.
- Sessao recusada solicita novo login/MFA no Edge. Falta de permissao pode
  continuar produzindo 403 mesmo com login correto.
- Em HTTP 400/404 tenta REST por caminho, URL direta e download.aspx. Falhas de
  autenticacao, disco e integridade nao acionam troca de rota.
- Um bloqueio exclusivo impede lotes concorrentes no mesmo destino.
- CTRL+C cancela e preserva parciais; pode aguardar timeout de leitura de 60 s.
- Arquivos publicados somente apos validacao do tamanho; nunca sobrescreve destino.
- Gravacao usa caminhos estendidos do Windows, incluindo documentos, parciais e
  metadados. Nao requer alterar o Registro. Preserva os nomes e a estrutura.

`-ListOnly` nao transfere documentos, mas cria pastas e arquivos de controle do
inventario no destino selecionado. Para ver alteracoes recentes, combine com
`-RefreshInventory`. O inventario nao bloqueia mudancas na origem.

## Retomar

Repita a mesma origem e destino. Preserve a pasta `.sharepoint_<id>`, os arquivos
`.spdownload.part` e seus `.json`. O controle usa:

- `manifest.json`: origem, escopo, totais e marcador de inventario completo;
- `inventory.jsonl`: um registro por arquivo, processado por streaming;
- `files/<hash>.json`: tamanho, versao conhecida e resultado de cada arquivo.

Inventarios interrompidos sao refeitos. Inventarios completos sao reutilizados,
a menos que seja informado `-RefreshInventory` ou o escopo mude.
Parciais identificados tentam Range/If-Range apenas com ETag forte e identidade
compativel. Content-Range e ETag sao verificados antes de anexar bytes. Resposta
200 a Range reinicia a transferencia; divergencia de tamanho invalida retomada.
Parciais sem identificacao sao preservados e reportados como falha.
Em divergencia de tamanho, registra esperado/recebido, Content-Length, ETag e
Content-Encoding. Consulta novamente os metadados no SharePoint antes da nova
tentativa; se tamanho/versao mudou, atualiza a referencia, invalida o inventario
salvo e baixa de novo. Nao aceita a resposta divergente apenas por seu tamanho.

O banco SQLite do Python nao e importado. Arquivos completos sao reconhecidos
por tamanho e registros novos; parciais no formato atual do Python compartilham
os campos `identity`, `size`, `etag`, mas seu aproveitamento depende de origem,
destino, identidade e ETag compativeis. Parciais antigos `.part.<id>` ou `.part`
nao sao importados. Nao execute ambos os programas sobre o mesmo lote ao mesmo tempo.

## Fila por planilha

```powershell
.\Download-SharePoint.ps1 -ProjectList 'C:\Listas\Migracao.xlsx' -PlanOnly
.\Download-SharePoint.ps1 -ProjectList 'C:\Listas\Migracao.xlsx' -Destination 'D:\Migracao'
.\Download-SharePoint.ps1 -ProjectList 'C:\Listas\Migracao.xlsx' -Priority 1,2 -ProjectRows 2,3,4 -Sheet 'Planilha1'
```

Primeira aba por padrao. A = ano, B = lote, C = SH, E = prioridade inteira positiva,
F = URL/hyperlink real. Nao executa formulas, nao abre Excel e nao altera o XLSX.
Ordem de prioridade crescente; empate usa numero da linha. Links duplicados,
prioridades invalidas e links ausentes aparecem como pendencias.
Pastas locais usam ano/lote/SH/hash da origem. Reutiliza a sessao do Edge entre
projetos, permite renovacao e continua a fila apos falha de um projeto.
`fila_projetos_resumo_<id>.json` e atualizado depois de cada projeto, inclusive
se houver interrupcao. Falhas ou pendencias resultam em codigo de saida 1.
`-PlanOnly` confere sem abrir Edge, solicitar destino ou exigir o driver.
Os filtros permitem distribuir linhas por maquina, sem coordenacao global.

## Teste e diagnostico

```powershell
.\Download-SharePoint.ps1 -BrowserTest
.\Download-SharePoint.ps1 -DiagnoseDownload
.\Download-SharePoint.ps1 -Url 'https://empresa.sharepoint.com/sites/Site/Docs/Pasta' -DiagnoseFile '/sites/Site/Docs/Pasta/arquivo.pdf'
```

Teste lista ate 20 arquivos e 200 subpastas por nivel. Selecione `P1`, `P2`, etc.
para navegar; depois escolha um arquivo de ate 100 MB.
Diagnostico usa a ultima falha final ou o caminho informado. Compara metadados
no navegador e resposta no terminal; testa conteudo por caminho, ID quando
disponivel, URL direta e download.aspx. Amostras de ate 32 KiB, sem salvar documentos.
Sucesso da amostra nao comprova download completo.

## Logs

`logs/execucao_<data>_<id>.log` preserva cada execucao em JSON por linha.
`logs/sharepoint_downloader.log` guarda historico rotativo de 5 MB e tres copias.
Registra login, inventario, tentativas, rotas, retomadas, conflitos, erros,
progresso a cada 30 s, resumo e duracao. Erros HTTP incluem status, codigo/mensagem
OData quando disponiveis e SPRequestGuid valido, com leitura limitada a 16 KiB.
URLs, emails e credenciais reconhecidas sao removidos; corpos e cookies nao
sao gravados. Os logs individuais nao sao apagados automaticamente.

## Driver

Confira Edge em `edge://settings/help`; obtenha Edge WebDriver em
https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/ .
Os tres primeiros numeros da versao devem corresponder ao Edge. Extraia
`msedgedriver.exe` nesta pasta ou informe `-DriverPath 'C:\Ferramentas\msedgedriver.exe'`.
Depois de atualizacoes do Edge, pode ser necessario atualizar o driver.
Os atalhos nao alteram politicas de execucao: siga o procedimento de scripts
permitido pela TI caso a maquina bloqueie a execucao.

## Validacao e diferencas

```powershell
powershell.exe -NoProfile -File .\tests\Test-Downloader.ps1
powershell.exe -NoProfile -File .\tests\Test-EdgeDriver.ps1
```

Testes locais usam servidor HTTP em loopback e fixtures XLSX, sem SharePoint,
Python ou Excel. Cobrem streaming, Range/If-Range, resposta 200 ignorando Range,
Content-Range invalido, truncamento e nova tentativa, rotas alternativas, erros
de login, HTML, 429, cancelamento, preservacao de parcial, filtro de credenciais,
fila XLSX, inventario, reaproveitamento, atualizacao, conflitos e bloqueio.
Fixtures ficam no TEMP para inspecao.
O segundo teste valida o Edge/driver instalado em modo headless, com pagina e
JSON locais, sem acessar o SharePoint ou usar sua conta.

O escopo portado e o downloader por sessao do navegador. O modo alternativo
Graph/MSAL/Device Code do Python nao foi incluido, pois exige cadastro de
aplicativo e contraria o caminho escolhido nesta versao. SQLite foi substituido
por JSON/JSONL nativos. Nao ha copia de historico de versoes, permissoes ou
metadados de biblioteca. Caminhos estendidos continuam sujeitos aos limites do
sistema de arquivos e de cada componente. Testes locais nao validam politicas corporativas, estabilidade em
varios TB nem o comportamento real do SharePoint: validar primeiro com lote pequeno.
