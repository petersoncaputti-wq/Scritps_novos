# Central de Scripts ProjectWise

Primeira versão do executor visual das rotinas manuais do ProjectWise.

## Como iniciar

Na Área de Trabalho, execute `Central de Scripts ProjectWise.bat` com duplo clique.

Esse inicializador permanece na Área de Trabalho e localiza o executor em
`Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise`. O arquivo `Iniciar-Executor.bat`
também pode ser usado diretamente dentro da pasta do executor.

Enquanto uma rotina estiver marcada como `legado`, ela será iniciada em um processo separado usando o código original. As novas versões integradas serão criadas somente na pasta `Rotinas`, preservando os originais.

## Catálogo

As opções exibidas estão em `scripts.json`. Para cadastrar uma nova rotina, inclua um objeto em `scripts` com os campos:

- `id`: identificador único;
- `ordem`: posição na listagem;
- `nome`, `categoria` e `descricao`;
- `arquivo`: caminho relativo à pasta do executor;
- `tipo`: `powershell`, `batch` ou `python`;
- `risco`: `baixo`, `medio`, `alto` ou `critico`;
- `dependencias`: itens no formato `module:Nome` ou `command:Nome`.

Rotinas de risco alto ou crítico exigem confirmação antes da execução. Dependências declaradas como ausentes bloqueiam o início e são mostradas no painel de detalhes.

## Logs

A central registra abertura, fechamento, bloqueios e processos iniciados na pasta `Logs`. Os logs internos gerados por cada rotina permanecem em suas pastas originais.
