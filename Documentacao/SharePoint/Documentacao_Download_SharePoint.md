# Documentação de uso — Download de arquivos do SharePoint

**Data:** 19/10/2026  
**Código:** `Scripts/SharePoint/sharepoint-downloader`  
**Modo documentado:** download da pasta completa (`--browser-all`).

## 1. Objetivo

O objetivo do projeto é realizar a **migração dos dados do SharePoint para o ProjectWise (PW)**. Para isso, os arquivos precisam primeiro ser baixados para armazenamento local e, em seguida, transferidos para o ProjectWise.

Na primeira leva da migração, ocorreram problemas recorrentes de **download incompleto dos dados**, comprometendo a disponibilidade dos documentos para a transferência ao PW. A aplicação foi criada para apoiar essa etapa com inventário dos arquivos, acompanhamento do progresso, registro de falhas e possibilidade de retomada, facilitando a identificação e o tratamento de pendências antes da importação.

O fluxo do projeto é: **SharePoint → download local nas máquinas disponibilizadas → conferência dos arquivos → transferência e validação no ProjectWise**. A aplicação documentada executa a etapa de download; a transferência ao PW é uma etapa posterior do processo de migração.

O processo copia o conteúdo atual dos arquivos, sem excluir ou alterar os documentos na origem.

## 2. Quantidade de arquivos e volume

O volume previsto para a migração é de **2 a 3 TB de dados**, conforme o escopo atualizado do projeto. Os downloads serão distribuídos entre as máquinas solicitadas para dividir a carga.

## 3. Tempo necessário

### Referência medida

A execução completa de teste registrada levou **1.620,35 segundos**, equivalentes a aproximadamente **27 minutos**, para copiar os 1.046 arquivos e 7,34 GB. O tempo é o intervalo interno do lote, abrangendo inventário e transferência; não inclui instalação e login inicial.

A taxa média equivalente do lote foi de aproximadamente **4,53 MB/s**, ou **36,2 Mbps**. Trata-se de uma referência observada em um lote específico, e não de uma velocidade garantida.

### Estimativa para o escopo de 2 a 3 TB

Mantendo a mesma taxa média do lote registrado em uma única máquina, a projeção seria de aproximadamente **122,7 a 184,1 horas**, ou **5,1 a 7,7 dias de execução contínua**. Em janelas de oito horas, seriam necessárias aproximadamente **16 a 24 janelas**. Essa extrapolação pressupõe comportamento semelhante dos arquivos, rede e armazenamento, e precisa ser revisada após um piloto representativo do escopo integral.

A distribuição dos lotes entre várias máquinas permite executar downloads em paralelo e pode reduzir o prazo total. A redução não é necessariamente proporcional à quantidade de máquinas, pois elas podem compartilhar a mesma conexão e estar sujeitas às mesmas limitações do SharePoint. O prazo do conjunto dependerá da distribuição dos lotes e da taxa agregada efetivamente obtida.

Como referência adicional de planejamento, a tabela abaixo utiliza **2 a 3 TB decimais** e uma hipótese de aproveitamento de 70% da banda reservada para o conjunto dos downloads:

| Banda reservada | Tempo estimado com 70% de aproveitamento |
|---|---:|
| 50 Mbps | 127,0 a 190,5 horas |
| 100 Mbps | 63,5 a 95,2 horas |
| 200 Mbps | 31,7 a 47,6 horas |
| 500 Mbps | 12,7 a 19,0 horas |

**Fórmula:** tempo em horas = volume em bytes × 8 ÷ (banda em Mbps × 1.000.000 × aproveitamento × 3.600).

## 4. Funcionamento do código

1. O operador inicia `Baixar Pasta Completa SharePoint.bat`, informa a URL da pasta e seleciona o destino.
2. Realiza o login corporativo e o MFA no Microsoft Edge, quando solicitado.
3. O programa consulta a API REST do SharePoint na sessão autenticada e percorre as subpastas e páginas de resultados.
4. Registra o inventário em SQLite, com caminho, tamanho e referência de versão disponível.
5. Transfere os arquivos em blocos, com três downloads simultâneos por padrão; a configuração aceita de um a oito.
6. Valida o tamanho transferido, registra os resultados e apresenta o resumo de baixados, existentes e falhas/conflitos.

Arquivos parciais identificados podem ser retomados quando a resposta do servidor permite. Arquivos existentes com tamanho compatível e sem divergência de versão conhecida são reaproveitados. Tamanho igual, sozinho, não comprova integridade do conteúdo. Conflitos são preservados para análise.

O inventário concluído é reutilizado nas execuções seguintes com a mesma origem e destino. Para procurar inclusões ou alterações na origem, utiliza-se `--refresh-inventory`. O inventário não impede que documentos mudem no SharePoint durante a cópia.

## 5. Recursos para execução e acompanhamento

- Máquinas disponibilizadas temporariamente para o projeto, com Windows e Microsoft Edge.
- Instalação temporária de Python compatível e das dependências da aplicação nas máquinas solicitadas, para distribuir a carga de download.
- Conta corporativa com acesso de leitura e download nas pastas selecionadas.
- Conexão de rede e destino com permissão de gravação e espaço livre suficiente.
- Armazenamento distribuído conforme os lotes atribuídos a cada máquina. Para **2 a 3 TB de dados**, considerar inicialmente **2,4 a 3,6 TB livres no conjunto dos destinos**, incluindo margem de planejamento de 20%, a revisar conforme o inventário, arquivos parciais e logs. Essa capacidade é agregada, e não uma exigência por máquina.
- Preferência por destino fora de pasta sincronizada pelo OneDrive, evitando sincronização concorrente da cópia.

O responsável deve acompanhar o progresso e os logs, manter o computador disponível durante a operação e atender eventuais solicitações de novo login. Para continuar após interrupção, deve utilizar a mesma URL e o mesmo destino e preservar o banco `.sharepoint_<identificador>.sqlite` e os arquivos parciais.

Os lotes devem ser divididos por pastas de origem, com registro da máquina responsável e do destino local de cada lote, evitando sobreposição e facilitando a conferência antes da transferência ao PW.

A conclusão deve ser conferida pelo resumo final e pela análise das falhas ou conflitos. A indicação de término do lote com pendências não equivale à cópia integral concluída.

Após a transferência dos dados e a validação da migração no ProjectWise, as máquinas serão **devolvidas à equipe de Infraestrutura e formatadas**. O uso do Python e do armazenamento local nessas máquinas será temporário, restrito à execução do projeto. Antes da devolução, os registros necessários à comprovação da migração devem ser preservados no local corporativo definido pelo responsável.

## 6. Resumo para apresentação

### Download contínuo por lista de projetos

A aplicação também permite ler uma planilha XLSX e baixar os projetos em sequência. A coluna E define a prioridade (1 primeiro) e a coluna F fornece o link real da pasta. Projetos com a mesma prioridade seguem a ordem das linhas. O operador utiliza o lançador `Baixar Projetos da Planilha.bat`, informa a planilha, seleciona o destino e autentica no Edge.

A sessão é reutilizada entre os projetos na mesma máquina, com novo login quando houver expiração ou exigência corporativa. Cada projeto possui destino separado por ano, lote, SH e identificador da origem. Falhas e links inválidos são registrados para tratamento, e os demais projetos continuam. O resumo JSON é atualizado ao longo da execução.

É possível selecionar linhas ou prioridades diferentes por máquina para dividir a carga. A prioridade é respeitada dentro da fila de cada equipamento; a distribuição entre máquinas precisa ser coordenada pelo responsável. Na planilha analisada, 74 projetos possuem links aceitos e dois dependem da correção das células F56 e F57.

### Texto para apresentação

> O projeto tem como objetivo migrar de 2 a 3 TB de dados do SharePoint para o ProjectWise. O download local é a etapa preparatória dessa transferência. A aplicação foi criada após problemas de downloads incompletos na primeira leva, para permitir acompanhar os arquivos e tratar pendências antes da importação no PW. Será necessária a instalação temporária de Python nas máquinas solicitadas para dividir a carga de download. Na execução de referência, foram baixados 1.046 arquivos e 7,34 GB em cerca de 27 minutos, sem falhas ao final. Pela mesma taxa, em uma única máquina, o download de 2 a 3 TB levaria aproximadamente 5,1 a 7,7 dias contínuos; o prazo com várias máquinas dependerá da taxa agregada obtida. A quantidade total de arquivos será confirmada por inventário. Após a transferência e validação no ProjectWise, as máquinas serão devolvidas à Infraestrutura e formatadas.

