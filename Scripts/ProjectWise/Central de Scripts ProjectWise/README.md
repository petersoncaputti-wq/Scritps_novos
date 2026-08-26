# Central de Scripts ProjectWise

Primeira versão do executor visual das rotinas manuais do ProjectWise.

## Como iniciar

Na Área de Trabalho, execute `Central de Scripts ProjectWise.bat` com duplo clique.

Esse inicializador permanece na Área de Trabalho e localiza o executor em
`Scritps_novos\Scripts\ProjectWise\Central de Scripts ProjectWise`. O arquivo `Iniciar-Executor.bat`
também pode ser usado diretamente dentro da pasta do executor.

## Pré-requisitos

- Windows PowerShell 5.1;
- módulos `PWPS_DAB` e `ImportExcel`;
- acesso aos datasources ProjectWise cadastrados nas rotinas;
- Python no ambiente virtual `.venv` da raiz do projeto;
- pacotes de `PWDM_Gerenciamento_Participantes_V2\requirements.txt`;
- Chromium instalado pelo Playwright.

Para preparar ou reparar o ambiente Python a partir da raiz do projeto:

```powershell
python -m venv --clear .venv
.\.venv\Scripts\python.exe -m pip install -r "Scripts\ProjectWise\ExecucaoManualUsuario\PWDM_Gerenciamento_Participantes_V2\requirements.txt"
.\.venv\Scripts\python.exe -m playwright install chromium
```

A Central prefere automaticamente esse ambiente virtual e bloqueia as rotinas
PWDM quando um pacote ou o navegador estiver ausente.

As rotinas manuais de servidor usam a janela de login do ProjectWise e não
armazenam senha no código. O parâmetro `-DoNotCreateWorkingDirectory` evita a
dependência de uma pasta de trabalho pertencente a outro perfil do Windows.

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
