# Solicitação de autorização para download de dados do SharePoint

**Destinatários:** Tecnologia da Informação e Segurança da Informação (SI)  
**Data:** 07/10/2026  
**Status:** Proposta para análise técnica e de segurança, sem autorização de execução em massa  
**Solicitante e área:** A definir  
**Volume informado:** Aproximadamente 3,3 TB, sujeito à confirmação por inventário

Solicitamos a avaliação de uma aplicação Python para copiar arquivos de uma pasta ou biblioteca do SharePoint Online para armazenamento corporativo aprovado. A primeira etapa será um teste controlado de autenticação, descoberta da pasta, listagem e download de um arquivo. A transferência completa dependerá da validação desse teste, do desenvolvimento dos controles adicionais e da autorização de TI e SI.

A finalidade de negócio, a classificação dos dados e a autorização do proprietário da informação devem ser registradas antes da execução. A cópia proposta abrange o conteúdo atual dos arquivos; histórico de versões, permissões, listas e demais metadados SharePoint não estão incluídos automaticamente. Portanto, a solução não deve ser apresentada como backup completo do SharePoint.

## 1 Dados que precisam ser definidos

| Informação | Valor ou responsável pela definição |
|---|---|
| Finalidade de negócio | A definir pelo solicitante e proprietário da informação |
| Nome do servidor | 01SSRV316TEMP |
| IP do servidor | 192.168.5.200 |
| FQDN do servidor | A confirmar por Infraestrutura |
| Ambiente | Homologação para o teste e produção para a execução completa, se autorizada |
| Usuário Microsoft 365 para o teste | t_peterson.caputti@ecorodovias.com.br |
| Conta Windows de execução | A definir; informar se será a mesma identidade do operador |
| Site, biblioteca e pasta de origem | A definir; informar URL e delimitação exata do escopo |
| Tenant ID e Client ID | A definir após registro aprovado no Microsoft Entra ID |
| Diretório e volume de destino | A definir por Infraestrutura e proprietário da informação |
| Classificação e retenção da cópia | A definir por proprietário da informação e SI |
| Janela de execução e banda autorizada | A definir por Redes, Infraestrutura e SI |
| Responsáveis pelo acompanhamento | A definir por solicitante e TI |

Não se deve substituir esses campos por nomes presumidos. O nome da máquina de desenvolvimento não identifica o servidor autorizado para a operação.

## 2 Estrutura da arquitetura

O fluxo proposto é: operador autorizado → aplicação no Windows → autenticação no Microsoft Entra ID → consultas HTTPS ao Microsoft Graph → download HTTPS do conteúdo → armazenamento corporativo aprovado. Logs e, na segunda etapa, um inventário persistente ficam no servidor executor.

O Entra ID emite tokens de acesso após autenticação e consentimento. O Graph fornece os identificadores do site, da biblioteca e dos arquivos, além de uma URL temporária de download. O conteúdo é transferido diretamente do serviço Microsoft para o destino local, sem plataforma externa intermediária. [Download pelo Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0).

| Componente | Responsabilidade | Situação |
|---|---|---|
| CLI Python e lançador BAT | Entrada de URL e destino; execução supervisionada | Implementado no POC |
| MSAL e MSAL Extensions | Login, reutilização de sessão e cache criptografado pelo Windows | Implementado no POC |
| Cliente Microsoft Graph | Consultas, paginação e tratamento de falhas temporárias | Implementado no POC |
| Resolução SharePoint | Identificação de site, bibliotecas e pasta pela URL | Implementado no POC |
| Transferência em streaming | Download de um arquivo sem carregar seu conteúdo inteiro em memória | Implementado no POC |
| Inventário e manifesto persistentes | Estado de cada arquivo, bytes, tentativas e resultado de validação | Proposto para a segunda etapa |
| Fila com concorrência limitada | Recursão, retomada, controle de velocidade e renovação de sessão | Proposto para a segunda etapa |
| Relatório de reconciliação | Sucessos, erros, divergências e arquivos alterados | Proposto para a segunda etapa |

O POC está em `Scripts/SharePoint/sharepoint-downloader`. Os lançadores existentes 02.1 e 05 foram considerados como referência para a execução de Python no Windows. Sua autenticação por automação de navegador não foi reutilizada.

## 3 Linguagem e componentes

Linguagem: Python 3.11 ou superior. Bibliotecas externas: `msal`, `msal-extensions` e `requests`. Componentes da biblioteca padrão: `pathlib`, `logging`, `dataclasses`, `json`, `argparse` e demais utilitários. Para a segunda etapa, propõe-se SQLite para o inventário e `hashlib` para SHA-256 local, sem necessidade de banco de dados externo.

A instalação utilizará ambiente virtual Python dedicado, dependências revisadas por TI e versões registradas para homologação. Não há necessidade de Selenium ou Playwright. O processo operacional deve executar sem privilégios administrativos; privilégios de instalação serão definidos pela política corporativa.

## 4 Requisitos para conexão com o SharePoint

1. Registro da aplicação no tenant corporativo do Microsoft Entra ID e consentimento conforme política da empresa.
2. Conta corporativa habilitada, com direito de leitura e download nos recursos de origem; restrições de compartilhamento, rótulos e políticas também precisam permitir a operação.
3. Saída HTTPS pela porta TCP 443 para autenticação Microsoft, `graph.microsoft.com`, o hostname SharePoint da empresa e os endpoints de conteúdo necessários.
4. DNS, horário do sistema, certificados e proxy corporativo funcionando. A aplicação deve manter a validação TLS habilitada e respeitar os controles de segurança da empresa.
5. Rede e firewall devem usar a relação oficial vigente de endpoints Microsoft 365. URLs de conteúdo podem apontar para hosts diferentes do hostname do site; liberar somente o site e o Graph pode ser insuficiente. Não se propõe uma lista estática de IPs nem liberação irrestrita da internet. [Endpoints oficiais Microsoft 365](https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges).
6. Permissão de gravação no destino, espaço disponível, políticas de acesso e proteção do servidor aprovadas por TI.

O retorno `http://localhost` do login interativo é um callback local da aplicação desktop. Não corresponde à publicação de um serviço HTTP acessível pela rede nem exige abertura de porta de entrada no perímetro. [Configuração MSAL Python](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/acquiring-tokens).

## 5 Recomendação de autenticação

### Teste supervisionado com o POC atual

Recomenda-se autenticação moderna delegada, com MSAL e login interativo no navegador Microsoft, respeitando MFA e Conditional Access. Registro single tenant, plataforma Mobile and desktop applications e redirect URI `http://localhost`. A aplicação não solicita nem armazena senha e não utiliza Client Secret.

O código atual solicita apenas a permissão Graph **Sites.Read.All delegated**, exigida pela descoberta de site utilizada no POC. Ela autoriza leitura dentro dos acessos do usuário conectado, mas não restringe a aplicação somente à pasta informada. O caminho configurado limita a operação do programa, não constitui uma restrição de autorização no Entra. A referência Microsoft indica AdminConsentRequired = No para essa permissão delegada; a política do tenant pode, ainda assim, exigir aprovação administrativa. [Permissão Sites.Read.All](https://learn.microsoft.com/en-us/graph/permissions-reference#sitesreadall).

Device Code existe como opção explícita, somente se aprovado e permitido pelo tenant. Não é alternativa para contornar MFA, consentimento ou Conditional Access. O cache de tokens usa DPAPI e fica protegido no perfil Windows do executor; tokens e URLs temporárias de download não devem aparecer em relatórios ou logs.

### Escopo restrito e execução automática futura

Para a transferência completa, recomenda-se avaliar **Sites.Selected**, com concessão explícita de leitura somente no site autorizado. O registro e o consentimento, isoladamente, não concedem acesso: também é necessário provisionar a permissão no recurso. Em modo delegado, os direitos efetivos continuam limitados pela interseção dos direitos do usuário e da aplicação. A mudança exige adaptação de escopos, resolução e homologação do código; não está implementada no POC. [Permissões selecionadas](https://learn.microsoft.com/en-us/graph/permissions-selected-overview).

Se TI exigir execução sem operador presente, recomenda-se uma identidade de aplicação separada, com Sites.Selected application e concessão read no site aprovado. A autenticação poderá utilizar certificado corporativo com chave privada protegida ou identidade gerenciada, quando a infraestrutura suportar esse mecanismo. Essa modalidade não usa usuário humano nem MFA interativo; exige controles próprios da identidade de workload, consentimento e aprovação específica de SI. Não se recomenda guardar senha de conta de serviço para automatizar o login atual.

## 6 Usuário que será utilizado

**UPN Microsoft 365 do operador do teste:** t_peterson.caputti@ecorodovias.com.br.  
**Conta Windows no servidor:** A definir.  
**Proprietário ou responsável pela identidade:** A definir.

O operador deverá ter apenas o acesso necessário à origem, sem papel administrativo para realizar o download. As ações administrativas de registro, consentimento e concessão de acesso deverão ser realizadas por TI com identidade autorizada. Se a operação final for automática, registrar também nome, Application ID, proprietário, forma de autenticação e validade da identidade de aplicação aprovada.

## 7 Instalação no servidor

**Nome do servidor:** 01SSRV316TEMP.  
**IP do servidor:** 192.168.5.200.  
**FQDN:** A confirmar por Infraestrutura.

Servidor informado pelo solicitante para a execução proposta. A instalação e a conectividade nesse servidor ainda precisam ser validadas por TI. O código e o ambiente virtual disponíveis atualmente pertencem à estação de desenvolvimento.

Propõe-se um Windows Server em versão homologada pela empresa, com Python compatível, diretório de aplicação dedicado e volume de dados separado. Dimensionamento inicial para avaliação: 4 vCPU e 8 GB de RAM; são parâmetros propostos, não requisitos mínimos medidos. O piloto deverá medir consumo e validar compatibilidade com proxy, proteção de endpoint e armazenamento.

Para 3,3 TB decimais de conteúdo, reservar preliminarmente **pelo menos 4,0 TB decimais livres** no volume de destino, incluindo aproximadamente 20% de margem. O número deverá ser recalculado pelo inventário, tamanho dos maiores arquivos, espaço de arquivos parciais, relatórios e crescimento. 4,0 TB correspondem a aproximadamente 3,64 TiB; verificar as unidades exibidas pelo armazenamento. Um disco nominal de 4 TB não garante 4 TB livres após formatação e outros usos.

TI deverá definir permissões NTFS, criptografia em repouso conforme classificação, EDR, monitoramento de disco e procedimento de retenção ou exclusão da cópia. Se o destino for um compartilhamento de rede, avaliar também SMB, suas permissões e a banda adicional servidor → compartilhamento. A configuração inicial pressupõe armazenamento diretamente conectado ao servidor.

## 8 Técnica utilizada para realizar o download

A técnica será cópia por APIs oficiais Microsoft, em modo somente leitura na origem. A aplicação consulta o Graph para identificar os arquivos e obtém `@microsoft.graph.downloadUrl`. Em seguida, transfere o conteúdo por HTTPS em blocos para um arquivo local temporário `.part`, sem encaminhar o token Graph ao host de conteúdo.

A URL de download é pré-autenticada e temporária; deve ser tratada como informação sensível e usada imediatamente. Não será registrada ou enviada a terceiros. A segunda etapa poderá usar HTTP Range para retomar transferências, obtendo nova URL quando necessário e validando a resposta 206 e o intervalo recebido antes de anexar bytes. Se o serviço responder com o conteúdo completo ou o arquivo tiver mudado, a transferência precisará ser reiniciada com controle. [Download e intervalos HTTP](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0).

Não haverá exclusão, movimentação ou alteração dos arquivos de origem como parte dessa operação. A transferência não garante uma fotografia transacional de uma biblioteca ativa; mudanças durante a execução precisam ser detectadas e reconciliadas.

## 9 Controle de integridade

**Disponível no POC:** download em streaming, arquivo parcial separado, comparação do total de bytes com o tamanho informado pelo Graph e promoção para o nome final apenas após essa comparação. Um arquivo final já existente não é sobrescrito. O POC não calcula hash, não valida ETag após a transferência e não retoma downloads.

**Obrigatório para a etapa completa, ainda a implementar e homologar:**

- Inventário persistente por Site ID, Drive ID e Item ID, com caminho, tamanho, ETag/cTag quando disponível, data de modificação, tamanho local, estado e tentativas.
- Verificação de identidade, tamanho e versão antes e depois da transferência. Arquivos alterados durante a cópia devem ficar pendentes de nova transferência, sem serem apresentados como íntegros.
- Cálculo de SHA-256 local e registro no manifesto para verificações posteriores. Esse hash local, isoladamente, não demonstra igualdade com a origem.
- Comparação com o hash remoto efetivamente fornecido, mediante cálculo local do mesmo algoritmo. Não presumir SHA-256 no Graph: a propriedade sha256Hash não é suportada. QuickXorHash pode apoiar a detecção de divergências, mas não é uma prova criptográfica de autenticidade. [Hashes do Graph](https://learn.microsoft.com/en-us/graph/api/resources/hashes?view=graph-rest-1.0).
- Reconciliação final por arquivo, quantidade e soma de bytes, com lista explícita de falhas, alterações e divergências. Tamanho e contagem iguais não substituem a verificação por arquivo.
- Mapeamento determinístico de nomes incompatíveis com Windows e caminhos longos, com detecção de colisões e preservação da correspondência origem → destino no manifesto.

Quando não houver hash remoto utilizável, o relatório deverá informar exatamente o nível de validação realizado. Arquivos críticos poderão requerer releitura ou validação adicional definida com SI. A conclusão dependerá de resolver as divergências ou documentar sua aceitação formal pelo responsável.

## 10 Estimativa de consumo de banda

Premissa de cálculo: **3,3 TB decimais = 3.300 GB = 26.400.000 megabits**. Trata-se do conteúdo transferido, sem histórico de versões. Fórmula do tempo ideal em horas: volume em bits dividido pela taxa em bits por segundo e por 3.600.

| Banda reservada | Taxa ideal | Tempo ideal contínuo | Cenário a 70% de eficiência |
|---|---|---|---|
| 50 Mbps | 6,25 MB/s | 146,7 h ou 6,1 dias | 209,5 h ou 8,7 dias |
| 100 Mbps | 12,5 MB/s | 73,3 h ou 3,1 dias | 104,8 h ou 4,4 dias |
| 200 Mbps | 25 MB/s | 36,7 h ou 1,5 dia | 52,4 h ou 2,2 dias |
| 500 Mbps | 62,5 MB/s | 14,7 h | 21,0 h |

Esses valores são cálculos de planejamento, não desempenho medido ou prazo garantido. A eficiência de 70% é uma hipótese, sujeita a quantidade e tamanho dos arquivos, latência, throttling, proxy, disco e concorrência com outras aplicações. Janelas parciais estendem o prazo: a 100 Mbps e 70% de eficiência, seriam aproximadamente 14 janelas de oito horas, sem incluir pausas adicionais.

Reservar preliminarmente **3,6 a 4,0 TB de tráfego de entrada** para conteúdo, protocolo e eventuais retransmissões, como margem de planejamento de 10% a 20%. Repetições integrais ou validações por releitura podem superar essa estimativa. A duração a 70% de eficiência não deve ser interpretada como acréscimo de 30% ao volume de dados.

Propõe-se piloto com medição de bytes, taxa e impacto no link. Redes deverá definir teto de banda e janela; o limitador no aplicativo e a concorrência controlada são funcionalidades futuras. O POC atual não possui limitador de banda. Controles de rede corporativos podem ser usados se disponíveis.

## 11 Estratégia de execução

1. **Preparar a solicitação:** completar servidor, identidade, finalidade, origem, destino e classificação; obter concordância do proprietário da informação.
2. **Revisar com TI:** avaliar código e dependências, registro Entra, rede, armazenamento, autenticadores e permissões propostas.
3. **Solicitar reunião com SI:** apresentar esta documentação e obter autorização explícita do teste, com escopo e condições registrados.
4. **Executar o POC autorizado:** autenticar, confirmar site e pasta, listar filhos diretos e baixar um arquivo pequeno. Registrar evidências e resultado sem divulgar tokens.
5. **Desenvolver e homologar a segunda etapa:** implementar inventário, recursão, retomada, integridade, renovação de autenticação, velocidade e relatórios. Validar interrupção e retomada, mudança de versão, nomes inválidos, colisões e falhas de rede/disco.
6. **Executar piloto representativo autorizado:** sugerem-se 10 a 50 GB, condicionados à aprovação, contendo arquivos pequenos e grandes. Medir banda, disco e tempo; ajustar recursos e controles. O POC atual, de um arquivo, não executa esse piloto em lote.
7. **Submeter plano de execução completa:** apresentar evidências e estimativas revisadas a TI e SI e confirmar a autorização para o escopo de 3,3 TB.
8. **Transferir por lotes monitorados:** respeitar janela, teto de banda e concorrência aprovados. Respeitar Retry-After em 429 e usar tentativas limitadas com espera progressiva para falhas temporárias. Não aumentar agressivamente a carga para superar throttling. [Throttling no Microsoft Graph](https://learn.microsoft.com/en-us/graph/throttling).
9. **Reconciliar e encerrar:** emitir manifesto e relatório de pendências, obter aceite e executar a política de retenção aprovada. Revogar acessos temporários e limpar cache e credenciais da aplicação quando a finalidade terminar, sob orientação de TI.

Critérios de interrupção: autorização revogada, impacto no ambiente produtivo, risco identificado por SI, indisponibilidade do destino, pouco espaço livre, bloqueio persistente de autenticação ou divergências de integridade sem tratamento. A origem deve permanecer preservada.

## 12 Encaminhamento e reunião com SI

Após o preenchimento dos campos pendentes e a análise preliminar de TI, solicitar reunião com Segurança da Informação para apresentar arquitetura, identidades, escopo de dados, controles de integridade, estimativas e estratégia de execução. Registrar os participantes, condições, responsáveis e decisão no chamado corporativo.

**Texto sugerido para a solicitação de reunião**

Solicito reunião com Segurança da Informação para avaliação e autorização do procedimento de cópia de aproximadamente 3,3 TB de arquivos do SharePoint Online para armazenamento corporativo aprovado. Encaminho a proposta técnica com arquitetura, requisitos de conexão, autenticação, controle de integridade, estimativa de banda e etapas de execução. Propomos iniciar com teste supervisionado de um arquivo e submeter os resultados antes da transferência completa. Na reunião, precisamos validar classificação e escopo dos dados, servidor e identidades, permissões Entra e SharePoint, proteção e retenção da cópia, janela de execução e critérios de autorização. A execução ocorrerá somente após os aceites aplicáveis.

Este documento prepara o encaminhamento. A reunião não foi solicitada ou agendada, e nenhuma liberação corporativa foi presumida.

## 13 Evidências disponíveis e pendências técnicas

O POC possui testes locais de resolução de URL e biblioteca, descoberta de subsite, paginação e retry simulados, download sem envio do bearer token ao host de conteúdo e proteção DPAPI. Os sete testes passaram na preparação inicial. Esses resultados não demonstram acesso real ao tenant, capacidade de transferência de 3,3 TB ou conformidade com políticas internas.

Pendências: preencher os dados da seção 1, confirmar consentimento e acesso efetivo, executar o teste real autorizado e desenvolver/homologar os componentes da segunda etapa. SI deverá avaliar a proposta em conjunto com as políticas internas; os controles descritos como futuros não devem ser considerados já disponíveis.
