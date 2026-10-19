# Acesso ARTESP

Arquivos cujo nome comeca com PL (sem distinguir maiusculas/minusculas) e cujo
numero antes da revisao e igual ou maior que 900 nao podem receber permissao.
Exemplos bloqueados: `PL-...-015-R00.pdf` e `DE-...-903-R00b.dwg`.
`DE-...-899-R00.pdf` e `DE-...-015-R900.pdf` nao sao bloqueados por essas regras.
O inventario mantem todos os documentos; `ArquivosExcluidos.csv` registra os
bloqueados e seus motivos. Somente elegiveis entram nas contagens/selecoes de
estados. A regra e revalidada no documento atual antes da escrita pela API.
As regras nao removem permissoes concedidas em execucoes anteriores.

Para remover concessoes anteriores desses arquivos, escolha **2 - reverter**
no inicio do script. Selecione disciplinas e workflow/estado(s), confira os
arquivos bloqueados, gere o plano e digite `LOTE`, depois `REVERTER LOTE`.
A reversao remove somente a entrada propria de ARTESP no workflow/estado atual.
Os demais participantes e seus direitos sao comparados e preservados. Projeto,
Unidade, camada Document, outros estados e versoes historicas nao sao alterados.
Nao restaura automaticamente toda a heranca nem desfaz outras mudancas anteriores.
Se a remocao faria voltar um acesso ARTESP herdado, a rotina nao remove e registra
`Nao revertido - ARTESP herdado; revisar origem`. Ela nao altera a pasta de origem.
Ausencia de ARTESP gera `Sem alteracao`. Os detalhes ficam no relatorio do lote.
Teste isolado: `Tests/Test-ReversaoARTESP.ps1`, sem escrita no ProjectWise.

Execute `18 - Preparar Acesso ARTESP.ps1` no Windows PowerShell. O script
reinicia em Windows PowerShell 64 bits/MTA quando necessario.

1. Selecione concessao e projeto. Confirme `APLICAR` se faltar ARTESP com Read.
2. Confirme a etapa de `2 - Unidade`: Read somente nesta pasta; descendentes podem herdar.
3. Navegue ate a etapa (por exemplo, Executivo). Selecione uma disciplina para navegar
   ou varias disciplinas pelos numeros separados por virgula ou ponto e virgula.
   Confirme as pastas escolhidas. A consulta inclui documentos atuais somente de
   cada pasta selecionada, sem subpastas. Workflow/estado selecionado vale para todas.
4. Selecione workflow/estado(s) pelos numeros, separados por virgula; confira o total.
5. Selecione todos os documentos ou escolha os numeros desejados. Confirme `SIM`.
6. Digite `LOTE` para preparar a aplicacao completa, `TESTE` para um arquivo ou Enter para encerrar.
7. Apos o planejamento de todos os documentos, digite `APLICAR LOTE` para executar.

Nos documentos, ARTESP recebe Read + File Read na ACL do workflow/estado atual.
Os participantes anteriores e suas mascaras sao preservados. Quando a ACL era
herdada, as regras atuais da pasta sao materializadas no documento. Essa ACL
propria deixa de acompanhar futuras mudancas de permissoes na origem.

Direitos diferentes ja existentes para ARTESP geram bloqueio, sem substituicao.
O lote para na primeira falha. Arquivos ja concluidos nao sao desfeitos e uma
falha de escrita pode deixar o arquivo em processamento com resultado parcial.
Revise o relatorio antes de repetir.

Os resultados ficam em `RelatoriosAcessoARTESP/<execucao>/Lote/ResultadoDocumentos.csv`.
Os documentos e os resultados identificam a pasta e seu ID. O arquivo
`ContagensPorPastaEstado.csv` detalha as quantidades por pasta, workflow e estado.
Cada documento tem plano com ACL propria/origem/esperada, metadados anteriores,
ACL depois e resultado. O plano aprovado e revalidado antes de aplicar.

As etapas de projeto e Unidade sao reais e independentes da confirmacao do lote.
Cancelar o lote nao desfaz essas etapas. O script nao modifica a camada Document,
o workflow, o estado ou as versoes historicas dos documentos.

Os scripts 19, 20 e 21 sao ferramentas de diagnostico/recuperacao dos testes.
Para uso normal, utilize o script 18 e mantenha a pasta `Funcoes` junto dele.

Validacao isolada: `Tests/Test-LoteARTESP.ps1`, sem conexao ao ProjectWise.

O planejamento reutiliza as permissoes por pasta e workflow/estado e identifica
ARTESP uma vez por lote. A consulta inicial de cada documento tambem e reutilizada.
Na aplicacao, as consultas do documento e da origem sao atualizadas para detectar
mudancas desde o planejamento; a comparacao completa depois da escrita permanece.

`ProgressoDocumentos.csv` registra cada transicao, incluindo o inicio da aplicacao,
por acrescentamento. `ResultadoDocumentos.csv` e consolidado ao final de cada fase,
no cancelamento e em falhas. Se houver encerramento abrupto, consulte o diario.
`TemposEtapas.csv` mede planejamento/aplicacao (sem incluir a espera pela confirmacao).
`TemposOperacoes.csv` mede as consultas instrumentadas e suas quantidades; nao
representa o tempo total de todas as chamadas nativas. O ganho real depende do PW.

Na aplicacao, documentos que ja possuem ARTESP e a ACL completa esperada entram
no caminho `Sem alteracao - ARTESP e ACL completa conferidos`. A rotina relê a
ACL inteira e a identidade, mas nao executa escrita nem duplica o planejamento.
O registro compacto fica em `ResultadoSemAlteracao.json`; o plano anterior
continua no diretorio do documento. Os demais arquivos mantêm a verificacao
completa antes e depois da escrita.

As operacoes `EscreverPermissao`, `ValidarACLAntes`, `ValidarACLDepois`,
`ValidarDocumentoAntes` e `ValidarDocumentoDepois` agora sao medidas separadamente.
Os documentos ja corretos usam `ValidarACLJaCorreta` e `ValidarDocumentoJaCorreto`.
Participantes sao reutilizados durante o planejamento e resolvidos novamente
na leitura atualizada da origem antes de aplicar; conversoes identicas de
mascaras sao reutilizadas durante o lote.

`Tests/Test-CacheARTESP.ps1` verifica o cache e o bloqueio de mudancas na origem
com consultas simuladas, usando apenas a conversao local de mascaras do PWPS_DAB.
