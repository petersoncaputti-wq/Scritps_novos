# Melhorias e automações dos scripts ProjectWise

**Comparativo gerencial e referência técnica**  
Data da análise: 28/09/2026 • Versão: 1.0 • Base versionada: `c499c35`, acrescida das alterações locais identificadas.

## 1. Objetivo e alcance

Este documento consolida as capacidades de automação identificadas no repositório e compara as atividades que elas atendem com o trabalho necessário sem essas automações. Abrange administração de usuários e acessos, criação de projetos, processamento de documentos, emissão de GRDs, inventários, importação e manutenção.

**Como interpretar o comparativo:** a coluna “Antes” descreve um cenário de referência sem a automação, inferido a partir das atividades atendidas pelos scripts. Esse cenário precisa ser validado com os responsáveis pelo processo. Quando há evidência de evolução técnica no histórico ou entre arquivos, ela é indicada separadamente. A existência de código não comprova sua implantação em produção, periodicidade de execução ou resultado operacional.

A análise foi feita por leitura dos scripts, documentação, catálogo da Central e histórico Git. Nenhuma rotina operacional foi executada e nenhum dado do ProjectWise foi alterado. Os benefícios apresentados são qualitativos; não foram fornecidas medições de tempo, custo ou volume que permitam quantificá-los.

## 2. Visão gerencial

O conjunto de scripts permite executar em lote atividades repetitivas, aplicar regras padronizadas e produzir registros para acompanhamento. A Central de Scripts reúne parte dessas rotinas em uma interface, com verificação de dependências e confirmação para operações de maior risco.

O trabalho do operador passa a se concentrar na preparação das entradas, seleção do escopo, revisão das ações e tratamento de exceções. Essa mudança pode reduzir esforço repetitivo e divergências de execução, mas o ganho efetivo depende do volume processado, das regras adotadas e da utilização das ferramentas.

| Dimensão | Antes — cenário de referência | Depois — capacidade identificada | Benefício esperado |
|---|---|---|---|
| Produtividade | Repetição de cadastros e alterações por usuário, projeto ou documento. | Processamento em lote, seleção de múltiplos projetos e uso de planilhas. | Menor esforço operacional por lote. |
| Padronização | Aplicação das regras a cada operação pelo responsável. | Regras de nomenclatura, estrutura, validação e encaminhamento incorporadas aos scripts. | Maior consistência entre projetos e concessões. |
| Rastreabilidade | Necessidade de consolidar manualmente o resultado das operações. | Logs e relatórios em rotinas específicas, com resumos e identificação de falhas. | Mais evidências para acompanhamento e diagnóstico. |
| Usabilidade | Localização e abertura de scripts individuais. | Catálogo visual com dez rotinas e painéis especializados. | Acesso mais organizado às ferramentas. |
| Controle | Conferências dependentes do procedimento adotado pelo operador. | Validações de entrada, prévias, modos de consulta/simulação e confirmações, conforme a rotina. | Mais oportunidades de identificar problemas antes de aplicar alterações. |
| Visibilidade do acervo | Consultas e consolidações separadas. | Inventário, contagem, identificação de nomes repetidos e exportação de dados. | Melhor apoio à análise do acervo documental. |

## 3. Comparativo por processo

As referências T01 a T24 remetem ao catálogo técnico da seção 5. “Depois” significa capacidade presente nos arquivos examinados, sujeita à confirmação de uso em produção.

### 3.1 Administração de usuários, acessos e projetos

| Ref. | Processo | Antes — cenário de referência | Depois — automação identificada |
|---|---|---|---|
| T01 | Acesso às ferramentas | Abrir cada arquivo e conhecer seus requisitos individualmente. | Central com categorias, descrição, verificação de dependências, indicação de risco e logs de execução. |
| T02 | Criação de usuários | Cadastrar dados e conferir cada usuário individualmente. | Leitura de CSV/Excel, validação, criação de usuários Bentley IMS, opção de atualizar existentes e geração de relatório. |
| T03 | Gestão de acessos | Localizar projetos e repetir a consulta ou alteração de acessos. | Interface com árvore de projetos, seleção de escopo, recursos de lote e registro das operações. |
| T04 | Inclusão de usuários no PW Web | Incluir usuários e atribuir funções projeto a projeto. | Entrada por planilha e processamento de usuários/funções nos projetos selecionados. |
| T05 | Participantes PWDM | Consultar e ajustar participantes em cada ambiente/projeto. | Fluxo conectado, cruzamento de projetos ProjectWise/PWDM, regras de operação e registros de resultados. |
| T06 | Usuários inativos | Consultar a situação das contas e aplicar ações individualmente. | Interface para relatório, desativação e exclusão, com confirmações e controles próprios. |
| T07 | Criação de projetos | Preencher propriedades e montar estruturas repetidamente. | Criação por formulário/parâmetros ou planilha, reaproveitando funções comuns e validações de entrada. |
| T08 | Buscas salvas | Configurar filtros e estados em cada projeto. | Criação de buscas padronizadas para estados do fluxo, com seleção das buscas e dos projetos. |
| T09 | Cadastro e dados dos projetos | Transcrever propriedades para planilhas. | Exportação no modelo de criação e extração de inventário técnico dos projetos. |

### 3.2 Recebimento, organização e distribuição documental

| Ref. | Processo | Antes — cenário de referência | Depois — automação identificada |
|---|---|---|---|
| T10 | Validação de documentos | Conferir número, revisão, versão e presença de arquivos correspondentes. | Aplicação de regras de validação, incluindo relação nativo/PDF e exceções previstas para extensões e modelos. |
| T11 | Pastas de disciplina e unidade | Criar e conferir a estrutura de destino a cada necessidade. | Cálculo dos destinos com base nos atributos e criação/verificação das pastas correspondentes. |
| T12 | Encaminhamento de documentos válidos | Localizar destino, verificar documentos anteriores e processar cada item. | Tratamento por regras de destino, consultas com cache e recursos de processamento/cópia nas versões disponíveis. A Central aponta para a v3. |
| T13 | Documentos importados e cópias para unidade | Organizar documentos importados e distribuir arquivos conforme o fluxo. | Rotinas específicas de movimentação de importados e cópia para unidade, incluindo tratamento de documentos anteriores na rotina de cópia. |
| T14 | GRDs para unidade | Selecionar documentos, organizar a emissão e preparar os registros. | Rotina de geração de GRDs com parâmetros de execução, simulação e teste de login. A Central aciona o modo real com confirmação. |
| T15 | Devolução ao projetista | Separar documentos, organizar GRD de saída e atualizar estados/atributos. | Busca de documentos elegíveis, definição da próxima GRD, criação de pasta, cópia e atualização de atributos/estados. |
| T16 | XFDF e documentos superados | Localizar arquivos de anotação associados a PDFs superados e ajustar sua organização. | Rotina de correlação de PDFs em “Superados” com arquivos XFDF e tratamento correspondente. |
| T17 | Notificações travadas | Identificar notificações em execução e intervir nas situações de travamento. | Rotina com consulta, identificação de notificações travadas e função de destravamento. |

### 3.3 Inventário, saneamento e migração

| Ref. | Processo | Antes — cenário de referência | Depois — automação identificada |
|---|---|---|---|
| T18 | Nomenclatura e revisão | Inspecionar e ajustar nomes de arquivos um a um. | Análise de nomes, sugestões de ajuste, relatórios e rotina específica para padronização de revisão R00. |
| T19 | Inventário documental | Consolidar listas e procurar repetições manualmente. | Excel com documentos dos projetos selecionados, sinalização de nomes repetidos, resumo por projeto e opções de versões, arquivo existente e reconciliação. |
| T20 | Exportação de áreas de transferência | Localizar as áreas e exportar seu conteúdo em operações separadas. | Busca recursiva por projeto, seleção de destino, interface gráfica e opção de somente listar. |
| T21 | Contagem de documentos | Consolidar contagens por pasta e subpasta. | Consulta recursiva sem baixar arquivos, com opção de incluir versões e respeito à visibilidade do usuário conectado. |
| T22 | Limpeza de áreas de transferência | Localizar e conferir manualmente os alvos de limpeza. | Interface e modo console, inventário de tamanho, validação de sessão/pastas, confirmações e medição do resultado. Há melhorias locais adicionais ainda não versionadas. |
| T23 | Importação e acompanhamento | Mapear arquivos, cadastrar destinos e consolidar o resultado da migração. | Scripts de mapeamento, importação a partir de dados de projetos e consolidação de relatórios de arquivos analisados/importados. |
| T24 | Estrutura para vídeo registro | Recriar manualmente uma árvore de pastas. | Seleção de pasta-modelo e replicação das subpastas, sem cópia de arquivos. |

## 4. Melhorias técnicas e evidências de evolução

| Melhoria | Referência anterior disponível | Situação atual observada | Efeito prático e limite |
|---|---|---|---|
| Centralização das rotinas | Scripts individuais continuam disponíveis em `ExecucaoManualUsuario`. O histórico registra a Central em `5a2510d`, de 31/07/2026. | Executor visual, catálogo e painéis por rotina. | Organiza a operação; não significa que todos os scripts do repositório estejam na Central. |
| Autenticação das rotinas de servidor | O README registra que anteriormente havia senha no código; `256b1af`, de 26/08/2026, abrange alterações de autenticação. | README define variável de ambiente para execução automática; chamadas pela Central utilizam login gráfico nas rotinas previstas. | Retira a dependência da senha escrita nesses scripts. Não comprova configuração da conta de serviço nem remoção do segredo do histórico. |
| Consultas na movimentação de documentos | Existem múltiplas versões de `MovimentaDocumentosValidos`. | Versões otimizadas incluem caches e medição de tempo; a v3 contém funções para tratar cópia de documento sem arquivo. | Busca reduzir chamadas repetidas e tratar casos específicos. Não há benchmark nesta análise. |
| Evolução do inventário | Histórico registra alterações em `9d68a69`, de 01/09/2026, e `46d269f`, de 16/09/2026. | Código mantém nomes repetidos, sinaliza duplicidades e oferece reconciliação opcional. | Amplia a capacidade de conferência. Nome repetido não comprova conteúdo duplicado. |
| Exportação, contagem e limpeza | `46d269f`, de 16/09/2026, registra arquivos dessas rotinas; `c499c35`, de 18/09/2026, altera a limpeza e sua interface. | Ferramentas específicas com seleção de escopo e interfaces próprias. | Amplia o suporte à manutenção do acervo. Datas de commit não são datas de implantação. |
| Continuidade após falha específica na limpeza | A versão `c499c35` não contém o novo tratamento local mostrado pelo diff. | Alteração local permite continuar após erro 58014 somente nas etapas previstas e com resultado confirmado por inventário; exclui condições de perda de conexão/sessão. | Evita interromper todo o lote em um caso restrito. Mudança local, sem homologação comprovada. |
| Persistência para retomada da limpeza | Nova função ausente na base versionada consultada. | Alteração local introduz gravação por arquivo temporário e substituição com backup; a interface também possui alterações locais. | Mecanismo para preservar estado. A presença da função não comprova validação completa do fluxo de retomada. |

### Situação da Central

O `scripts.json` contém **dez rotinas**: nove marcadas como `integrado` e criação de usuários marcada como `em_migracao`. Existe implementação e painel para criação de usuários, mas a classificação de migração deve permanecer explícita.

Alguns READMEs das subpastas ainda dizem “pasta reservada”, embora haja painéis e classificação de integração no catálogo. Essa documentação precisa ser atualizada para refletir o estado efetivo. A classificação do catálogo é evidência de configuração, não de homologação operacional.

## 5. Catálogo técnico e rastreabilidade

Caminhos relativos à raiz do repositório. Abreviações utilizadas abaixo:

- **C:** `Scripts/ProjectWise/Central de Scripts ProjectWise`
- **M:** `Scripts/ProjectWise/ExecucaoManualUsuario`
- **A:** `Scripts/ProjectWise/ExecucaoAutomaticaServidor`
- **I:** `Scripts/ProjectWise/ImportacaoOneDrive`
- **P:** `M/PWDM_Gerenciamento_Participantes_V2`

| Ref. | Arquivos ou componentes principais | Entradas, saídas e observações técnicas |
|---|---|---|
| T01 | `C/Executor.ps1`, `C/scripts.json`, `C/Rotinas/` | Catálogo JSON; painéis WinForms; logs da Central. Os logs internos das rotinas têm destinos próprios. |
| T02 | `M/03 - Criar Usuarios ProjectWise em Lote.ps1`; `C/Rotinas/01-CriarUsuariosLote/` | CSV/Excel; parâmetros `SomenteValidar`, `GerarModelo` e `AtualizarExistentes`; relatório CSV. Exportação da lista no modo integrado depende da escolha do operador. |
| T03 | `M/02 - Gerenciar Acessos de Projetos.ps1` | Projetos, usuários, grupos/listas; interface e logs de permissões. |
| T04 | `M/13 - Incluir Usuarios em Projetos PW Web por Planilha.ps1`; `P/incluir_usuarios_projetos_roles_pwdm_v2.py` | Planilha, seleção de projetos, funções e fluxo de navegador/API. |
| T05 | `P/gerenciar_participante_pwdm_connected_v2.py`, `gerenciar_participante_pwdm_v2.py`, `regras_v2.py`, `lote_usuarios_pwdm_v2.py` | Autenticação assistida; regras separadas em módulo; logs/resumos e testes locais de regras disponíveis. Testes não foram executados nesta documentação. |
| T06 | `M/04 - Gerenciar Usuarios Inativos.ps1` | Consulta e escolha de ação; classificação de risco crítico na Central. |
| T07 | `M/01 - Criar Projetos no ProjectWise.ps1`; `Scripts/ProjectWise/Funcoes/CriarProjeto.ps1` | Parâmetros/formulário ou planilha; criação e configuração de projetos. Há outra cópia de funções sob `M/Funcoes`. |
| T08 | `M/06 - Criar Buscas Salvas em Projetos.ps1` | Seleção de projetos e estados; `WhatIf`, substituição opcional de existentes e inclusão de subpastas no escopo das buscas. |
| T09 | `M/07 - Exportar Modelo de Criacao de Projetos ProjectWise.ps1`; `M/09 - Extrair Dados de Projetos.ps1` | Modelo CSV de cadastro; inventário técnico em Excel com propriedades dos projetos. |
| T10 | `A/ValidaDocumentos.ps1` | Busca salva de documentos recebidos; validações de número, arquivos correspondentes e revisão/versão. A chamada de validação “Previsto na LD” está comentada: não considerar essa regra ativa. |
| T11 | `A/CriaPastaDisciplina.ps1`, `A/CriaPastaDisciplinaUnidade.ps1` | Atributos de disciplina/fase/poder concedente; funções de cálculo e garantia da estrutura de destino. |
| T12 | `A/MovimentaDocumentosValidos.ps1`, `_v2.ps1`, `_v2_ajuste_ketra.ps1`, `_v3.ps1` | Variantes do mesmo processo; não contabilizar como quatro automações independentes. A Central utiliza `_v3.ps1`. |
| T13 | `A/MovimentaDocumentosImportados.ps1`, `A/CopiaDocumentos.ps1` | Regras de organização e distribuição documental. Confirmar encadeamento e agendamento em produção. |
| T14 | `A/GeraGRDsDocumentosComUnidade.ps1`; cópia na raiz e `ExecutarGeraGRDsDocumentosComUnidade.ps1` | Parâmetros `Executar`, `Simular` e `SomenteTestarLogin`. A Central usa a cópia de A. A cadência semanal está descrita no README, sem verificação do agendamento. |
| T15 | `A/DevolveDocumentosProjetista.ps1`; cópia na raiz | GRD de saída, numeração, cópia de documentos e atualização de atributos/estados. |
| T16 | `A/MoverXFDFParaSuperados.ps1` | Consultas de PDFs e XFDF; correspondência por caminhos e processamento de superados. |
| T17 | `A/checaNotificacoesTravadas.ps1` | Consulta de notificações e rotina de destravamento. |
| T18 | `M/10 - Validar e Sugerir Ajuste de Nome de Arquivos.ps1`; `Renomear_Arquivos_Revisao_R00.ps1` na raiz | Pasta local/rede, análise e CSV; a rotina R00 inclui etapa de renomeação. Analisar/sugerir e renomear são capacidades distintas. |
| T19 | `M/11 - Inventariar Arquivos de Projetos.ps1` | Parâmetros `IncluirVersoes`, `ArquivoExistente`, `Reconciliar`; Excel com documentos, resumo e informações de conferência. |
| T20 | `M/12 - Exportar_Areas_de_Transferencia_ProjectWise.ps1`; `M/Interface_Exportacao_ProjectWise.ps1` | Destino local e seleção de projetos; `SomenteListar` e `ModoConsole`. |
| T21 | `M/14 - Contar Documentos ProjectWise.ps1` | Caminho interno ou endereço PW; opção de versões. Nomes iguais em pastas distintas são contados separadamente. |
| T22 | `M/15 - Limpar_Areas_de_Transferencia_ProjectWise.ps1`; `M/Interface_Limpeza_ProjectWise.ps1` | `SomenteListar`, seleção de alvos, inventários e controles de exclusão. Ambos possuem alterações locais não commitadas na data da análise. |
| T23 | `I/ScriptsPowershell/mapeamentoArquivos.ps1`, `ImportacaoArquivos.ps1`, `RelatorioImportacao.ps1` | Planilha de projetos, arquivos de origem e CSVs de mapeamento; relatório consolida contagens e tamanhos. |
| T24 | `M/Cria estrutra de pastas(video_registro).ps1` | Pasta de origem e destino; replica estrutura sem arquivos. |

### Dependências e operação

As dependências variam por rotina. A Central documenta Windows PowerShell 5.1, `PWPS_DAB`, `ImportExcel` e, para fluxos PWDM, Python, pacotes do `requirements.txt` e Chromium do Playwright. Algumas rotinas tratam a exigência de execução em MTA do módulo ProjectWise. O protótipo Graph possui requisitos próprios de PowerShell 7 e Microsoft Graph.

Os arquivos `.bat`, painéis, funções compartilhadas, scripts de diagnóstico e testes dão suporte aos processos acima; não representam necessariamente novos benefícios de negócio independentes. As cópias na raiz e versões alternativas exigem identificação do arquivo realmente utilizado em produção.

## 6. Iniciativas ainda em desenvolvimento

| Iniciativa | Evidência atual | Como apresentar |
|---|---|---|
| Importação automática de Incoming no PWDM | `A/DM_Auto_Import/README.md` define fase 1 de diagnóstico, captura de requisições/respostas e navegação manual. | Diagnóstico disponível. Aceite e importação automática não devem ser apresentados como concluídos. |
| Importador direto SharePoint → ProjectWise via Graph | `I/PW-SharePoint-Importer/README.md` descreve estrutura e conexão inicial. | Protótipo de conexão. Download, conexão ProjectWise, importação e logs finais ainda não implementados conforme o README. |
| Evolução arquitetural do PWDM | `P/PLANO_EVOLUCAO_V2.md` descreve extração de módulos, interface, resiliência e funções avançadas. | Plano de evolução. Conferir cada entrega no código e na homologação; não contabilizar o plano inteiro como executado. |
| Melhorias locais de limpeza | Diff local de T22 em relação a `c499c35`. | Alterações em desenvolvimento, separadas da base versionada e da implantação confirmada. |

Os scripts de importação existentes em `I/ScriptsPowershell` e o novo importador via Graph são iniciativas distintas. A existência dos primeiros não comprova a conclusão do segundo.

## 7. Indicadores para comprovar os ganhos

Para transformar os benefícios esperados em resultados demonstrados, comparar lotes equivalentes, mantendo volume, tipo de atividade e critérios de sucesso. Separar tempo de trabalho humano do tempo total de processamento.

| Indicador | Cálculo sugerido | Fonte necessária |
|---|---|---|
| Esforço humano por lote | Minutos de preparação + execução assistida + conferência + correções. | Apontamento do operador antes e depois. |
| Redução do esforço (%) | `(esforço anterior − esforço atual) / esforço anterior × 100`, quando o anterior for maior que zero. | Amostras comparáveis. |
| Tempo total do processo | Horário da conclusão − horário do início. | Logs e registro do processo anterior. |
| Taxa de sucesso | Itens concluídos com sucesso / itens tentados × 100. | Resumo por execução; definir tratamento de ignorados e sem alteração. |
| Taxa de retrabalho | Itens que exigiram correção / itens processados × 100. | Registro de falhas e correções. |
| Horas mensais liberadas | Diferença de esforço humano por lote × lotes mensais / 60. | Tempos em minutos e volume mensal real. |
| Cobertura de uso | Projetos/concessões atendidos / projetos/concessões elegíveis × 100. | Inventário do escopo e execuções confirmadas. |

### Quadro para preenchimento pelos responsáveis

| Processo | Volume do lote | Esforço antes (min) | Esforço depois (min) | Falhas antes/depois | Data de implantação | Responsável/evidência |
|---|---|---|---|---|---|---|
| Criação de usuários | A levantar | A medir | A medir | A medir | A confirmar | A informar |
| Gestão de acessos/PWDM | A levantar | A medir | A medir | A medir | A confirmar | A informar |
| Validação e organização documental | A levantar | A medir | A medir | A medir | A confirmar | A informar |
| Geração de GRDs | A levantar | A medir | A medir | A medir | A confirmar | A informar |
| Inventário/exportação/limpeza | A levantar | A medir | A medir | A medir | A confirmar | A informar |

Não preencher campos desconhecidos com zero: zero representa uma medição, enquanto “a medir” identifica informação pendente.

## 8. Validação para apresentação institucional

Esta versão já permite apresentar o escopo e as capacidades identificadas. Para publicar resultados como “melhorias implantadas”, os responsáveis devem confirmar:

1. Quais rotinas estão em produção, em quais concessões/projetos e desde quando.
2. Como cada atividade era efetivamente realizada antes, inclusive automações anteriores já existentes.
3. Qual arquivo/versão está em uso e quais rotinas estão agendadas.
4. Quais evidências de execução e homologação sustentam o status de cada entrega.
5. Os volumes e tempos que permitam calcular ganhos efetivos.

**Texto-base para apresentação à gestão:** “O conjunto de automações ProjectWise reúne recursos para cadastro em lote, gestão de acessos, padronização de projetos, validação e organização de documentos, emissão de GRDs e inventário do acervo. As rotinas incorporam regras e controles que apoiam a execução padronizada e o acompanhamento dos resultados. A quantificação dos ganhos de produtividade e a abrangência da implantação serão consolidadas com os responsáveis a partir dos registros de operação.”
