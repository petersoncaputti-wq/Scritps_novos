# Acesso ARTESP

Execute `18 - Preparar Acesso ARTESP.ps1` no Windows PowerShell. O script
reinicia em Windows PowerShell 64 bits/MTA quando necessario.

1. Selecione concessao e projeto. Confirme `APLICAR` se faltar ARTESP com Read.
2. Confirme a etapa de `2 - Unidade`: Read somente nesta pasta; descendentes podem herdar.
3. Navegue ate a disciplina. A consulta inclui documentos atuais somente dessa pasta.
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
Cada documento tem plano com ACL propria/origem/esperada, metadados anteriores,
ACL depois e resultado. O plano aprovado e revalidado antes de aplicar.

As etapas de projeto e Unidade sao reais e independentes da confirmacao do lote.
Cancelar o lote nao desfaz essas etapas. O script nao modifica a camada Document,
o workflow, o estado ou as versoes historicas dos documentos.

Os scripts 19, 20 e 21 sao ferramentas de diagnostico/recuperacao dos testes.
Para uso normal, utilize o script 18 e mantenha a pasta `Funcoes` junto dele.

Validacao isolada: `Tests/Test-LoteARTESP.ps1`, sem conexao ao ProjectWise.
