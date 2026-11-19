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

## Visual

O visual segue o Ecovias Brand Book (abril/2025) e fica centralizado em `Tema.ps1`:

- cores: verde-escuro `#056B28` (barra de título, botão principal), verde-médio `#1DA013` (seleção) e verde-claro `#72D216`;
- fonte Lexend, carregada de `Fontes\` apenas para o processo da Central (sem instalar no Windows; na falta dela, usa Verdana);
- ícones das rotinas com a fonte Segoe Fluent Icons, dentro da forma "janela para o mundo" do brand book.

Atalhos: `Ctrl+F` busca, setas mudam a seleção, `Enter` abre a rotina, `Esc` limpa a busca.

### Tela das rotinas de console (03, 04, 06, 07, 08, 09 e 10)

Essas rotinas usam o painel compartilhado `Rotinas\03-IncluirUsuariosWeb\Painel.IncluirUsuariosWeb.ps1`.
Os scripts Python/PowerShell não mudam: o painel lê a saída e transforma cada pergunta no controle adequado,
enviando ao script o mesmo texto que seria digitado no console.

| Saída do script | Controle |
|---|---|
| `NN. Nome` ou `[N] Nome` + pergunta de concessão/role | lista com filtro (envia o número) |
| lista de projetos/buscas | lista com caixas de marcação e filtro (envia `1,4,7`; tudo marcado envia `todos`/`T`) |
| `Escolha [1/2]`, `Acao desejada [1/2/3]` | botões com o texto de cada opção |
| `[S/N]` | quadro com o resumo e botões Sim/Não |
| `Confirmacao:` (exclusão) | campo que só libera com `CONFIRMAR EXCLUSAO` |
| `Quando estiver autenticado...` | botão para continuar após o login no navegador |
| qualquer outra pergunta | caixa de texto (comportamento anterior) |

`[n/N]` na saída alimenta a barra de etapas; `[OK]`, `[ERRO]` e `[AVISO]` são coloridos e contados no resumo final.
O console original continua em "Detalhes técnicos" e no arquivo de log. Para mostrar os passos de uma rotina
em "Como funciona", passe `-Etapas` na chamada de `New-IncluirUsuariosWebPanel`.

## Catálogo

As opções exibidas estão em `scripts.json`. Para cadastrar uma nova rotina, inclua um objeto em `scripts` com os campos:

- `id`: identificador único;
- `ordem`: posição na listagem;
- `nome`, `categoria` e `descricao`;
- `arquivo`: caminho relativo à pasta do executor;
- `tipo`: `powershell`, `batch` ou `python`;
- `risco`: `baixo`, `medio`, `alto` ou `critico` (só `critico` aparece em vermelho; `alto` mostra "Pede confirmação");
- `icone`: código do glifo da Segoe Fluent Icons (ex.: `E72E` para cadeado);
- `dependencias`: itens no formato `module:Nome` ou `command:Nome`.

Rotinas de risco alto ou crítico exigem confirmação antes da execução. Dependências declaradas como ausentes bloqueiam o início e são mostradas no painel de detalhes.

## Logs

A central registra abertura, fechamento, bloqueios e processos iniciados na pasta `Logs`. Os logs internos gerados por cada rotina permanecem em suas pastas originais.
