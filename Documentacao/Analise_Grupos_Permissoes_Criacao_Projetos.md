# Analise dos grupos e permissoes na criacao de projetos

Data: 13/10/2026. Analise estatica do codigo local; nao representa uma auditoria dos acessos efetivos no datasource.

## Fluxo efetivamente chamado pela Central

Painel.CriarProjetos.ps1 -> ExecucaoManualUsuario/01 - Criar Projetos no ProjectWise.ps1 -> ProjectWise/Funcoes/CriarProjeto.ps1. A copia em ExecucaoManualUsuario/Funcoes nao e carregada por esse fluxo.

CriarProjeto: pastas -> listas de usuarios -> grupos e vinculos -> seguranca de pastas/documentos.

## Grupos e vinculos exatos

S = Sigla Concessao; P = Nome Projeto. Os nomes abaixo preservam a grafia do codigo.

| Grupo | Listas das quais o grupo se torna membro | Linha fonte |
|---|---|---|
| Engenharia GPR | S-P-ENGENHARIA-GESTOR ENG<br>S-P-ENGENHARIA-CONSULTA-TODAS REV<br>S-P-UNIDADE-CONSULTA-TODAS REV<br>S-P-UNIDADE-GESTOR ENG<br>S-P-GRD-CONSULTA<br>S-P-LD-CONSULTA | 814 |
| S-P-PODER CONCEDENTE | S-P-UNIDADE-Poder Concedente | 832 |
| S-ENGENHARIA UNIDADE | S-P-UNIDADE-CONSULTA-TODAS REV<br>S-P-UNIDADE-Lider Unidade | 845 |
| S-P-ASSISTENTE ENG | S-P-ENGENHARIA-ASSISTENTE ENG<br>S-P-ENGENHARIA-CONSULTA-TODAS REV<br>S-P-UNIDADE-CONSULTA-TODAS REV<br>S-P-UNIDADE-Lider Unidade<br>S-P-GRD-EDICAO<br>S-P-LD-EDICAO | 859 |
| S-P-GESTOR ENG | S-P-ENGENHARIA-GESTOR ENG<br>S-P-ENGENHARIA-CONSULTA-TODAS REV<br>S-P-UNIDADE-CONSULTA-TODAS REV<br>S-P-UNIDADE-GESTOR ENG<br>S-P-GRD-CONSULTA<br>S-P-LD-CONSULTA | 877 |
| S-P-ESPECIALISTA | S-P-ENGENHARIA-ESPECIALISTA<br>S-P-ENGENHARIA-CONSULTA-TODAS REV<br>S-P-GRD-CONSULTA<br>S-P-LD-CONSULTA | 895 |
| S-P-PROJETISTA | S-P-GRD-CONSULTA | 911 |
| S-P-ENG CONSULTA ULTIMA REV | S-P-ENGENHARIA-CONSULTA | 924 |
| S-P-ENG CONSULTA TODAS REV | S-P-ENGENHARIA-CONSULTA-TODAS REV | 937 |
| S-P-CONSULTA ULTIMA REV | S-P-UNIDADE-CONSULTA | 950 |
| S-P-CONSULTA TODAS REV | S-P-UNIDADE-CONSULTA-TODAS REV | 963 |
| S-P-GESTOR UNIDADE | S-P-UNIDADE-CONSULTA-TODAS REV<br>S-P-UNIDADE-Lider Unidade | 976 |

## Estrutura das listas

Uma lista da concessao S contem a lista S-P. S-P contem ENGENHARIA, UNIDADE, LD e GRD. As listas de papeis sao membros da lista da area. CONSULTA-TODAS REV e membro de CONSULTA. O grupo e membro das listas indicadas acima; a rotina nao inclui pessoas nos grupos.

Sao 20 listas especificas por projeto e uma lista compartilhada da concessao (21 nomes verificados).

## Matriz completa de chamadas de seguranca

149 chamadas definidas: 140 nas funcoes chamadas pelo fluxo principal e 9 em DefinirAcessoXFDF, que nao e chamada pelo fluxo principal.

Codigos MemberAccess preservados literalmente para permitir comparacao com o modulo instalado. Esta tabela registra as chamadas do script; nao resolve heranca, combinacao de acessos, regras preexistentes nem permissoes de transicao dos workflows.

| Funcao/linha | Alvo | Tipo | Workflow | Estado | Membro | Acessos |
|---|---|---|---|---|---|---|
| DefinirAcessoRaizProjeto/443 | $nomePastaRaiz | Pastas |  |  | Engenharia GPR | r |
| DefinirAcessoRaizProjeto/444 | $nomePastaRaiz | Pastas |  |  | $siglaConcessao-ENGENHARIA UNIDADE | r |
| DefinirAcessoRaizProjeto/446 | $nomePastaRaiz | Pastas |  |  | Everyone | na |
| DefinirAcessoRaizProjeto/447 | $nomePastaRaiz | Pastas |  |  | $nomeUserList | r |
| DefinirAcessoRaizProjeto/449 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-ASSISTENTE ENG | r |
| DefinirAcessoRaizProjeto/450 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-ENG CONSULTA TODAS REV | r |
| DefinirAcessoRaizProjeto/451 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-ENG CONSULTA ULTIMA REV | r |
| DefinirAcessoRaizProjeto/452 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-ESPECIALISTA | r |
| DefinirAcessoRaizProjeto/453 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-GESTOR ENG | r |
| DefinirAcessoRaizProjeto/454 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-GESTOR UNIDADE | r |
| DefinirAcessoRaizProjeto/455 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-PROJETISTA | r |
| DefinirAcessoRaizProjeto/456 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-CONSULTA TODAS REV | r |
| DefinirAcessoRaizProjeto/457 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-CONSULTA ULTIMA REV | r |
| DefinirAcessoRaizProjeto/459 | $nomePastaRaiz | Documentos |  |  | Everyone | na |
| DefinirAcessoEngenharia/468 | $pastaEngenharia | Pastas |  |  | Everyone | na |
| DefinirAcessoEngenharia/469 | $pastaEngenharia | Pastas |  |  | $nomeUserListEngenharia | r |
| DefinirAcessoEngenharia/470 | $pastaEngenharia | Documentos |  |  | admin | fc |
| DefinirAcessoEngenharia/472 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em elaboracao | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/473 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em elaboracao | $nomeUserListEngenharia-ESPECIALISTA | c, r, w, cw, fr, fw |
| DefinirAcessoEngenharia/474 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em elaboracao | $nomeUserListEngenharia-ASSISTENTE ENG | c, r, w, cw, fr, fw |
| DefinirAcessoEngenharia/475 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em elaboracao | $nomeUserListEngenharia-GESTOR ENG | c, r, w, cw, fr, fw |
| DefinirAcessoEngenharia/476 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em elaboracao | Everyone | na |
| DefinirAcessoEngenharia/478 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise do Assistente | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/479 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise do Assistente | $nomeUserListEngenharia-ASSISTENTE ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/480 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise do Assistente | $nomeUserListEngenharia-GESTOR ENG | r, fr |
| DefinirAcessoEngenharia/481 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise do Assistente | Everyone | na |
| DefinirAcessoEngenharia/483 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise da Engenharia | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/484 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise da Engenharia | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/485 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise da Engenharia | Everyone | na |
| DefinirAcessoEngenharia/487 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise Especifica | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/488 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise Especifica | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/489 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise Especifica | $nomeUserListEngenharia-ESPECIALISTA | r, w, cw, fr |
| DefinirAcessoEngenharia/490 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em analise Especifica | Everyone | na |
| DefinirAcessoEngenharia/492 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em consolidacao pela Engenharia | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/493 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em consolidacao pela Engenharia | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/494 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em consolidacao pela Engenharia | Everyone | na |
| DefinirAcessoEngenharia/496 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Projetista | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/497 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Projetista | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/498 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Projetista | $nomeUserListEngenharia-PROJETISTA | r, fr |
| DefinirAcessoEngenharia/499 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Projetista | Everyone | na |
| DefinirAcessoEngenharia/501 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em revisao pelo Projetista | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/502 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em revisao pelo Projetista | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/503 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em revisao pelo Projetista | $nomeUserListEngenharia-PROJETISTA | r, fr |
| DefinirAcessoEngenharia/504 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em revisao pelo Projetista | Everyone | na |
| DefinirAcessoEngenharia/506 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/507 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/508 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade | Everyone | na |
| DefinirAcessoEngenharia/510 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade c/ Ressalvas | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/511 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade c/ Ressalvas | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/512 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade c/ Ressalvas | Everyone | na |
| DefinirAcessoEngenharia/514 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade c/ Pendencias do Especialista | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/515 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade c/ Pendencias do Especialista | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/516 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em Envio para Unidade c/ Pendencias do Especialista | Everyone | na |
| DefinirAcessoEngenharia/518 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/519 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/520 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade | Everyone | na |
| DefinirAcessoEngenharia/522 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade com Ressalvas | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/523 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade com Ressalvas | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/524 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade com Ressalvas | Everyone | na |
| DefinirAcessoEngenharia/526 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade e Projetista com Ressalvas | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/527 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade e Projetista com Ressalvas | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/528 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade e Projetista com Ressalvas | $nomeUserListEngenharia-PROJETISTA | r, fr |
| DefinirAcessoEngenharia/529 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade e Projetista com Ressalvas | Everyone | na |
| DefinirAcessoEngenharia/531 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias do Especialista | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/532 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias do Especialista | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/533 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias do Especialista | $nomeUserListEngenharia-ESPECIALISTA | r, w, cw, fr |
| DefinirAcessoEngenharia/534 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias do Especialista | Everyone | na |
| DefinirAcessoEngenharia/536 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias da Engenharia | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/537 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias da Engenharia | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/538 | $pastaEngenharia | Documentos | Workflow - Engenharia | Enviado para Unidade c/ Pendencias da Engenharia | Everyone | na |
| DefinirAcessoEngenharia/540 | $pastaEngenharia | Documentos | Workflow - Engenharia | Devolvido pela Unidade | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/541 | $pastaEngenharia | Documentos | Workflow - Engenharia | Devolvido pela Unidade | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/542 | $pastaEngenharia | Documentos | Workflow - Engenharia | Devolvido pela Unidade | Everyone | na |
| DefinirAcessoEngenharia/544 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em devolucao para Unidade | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/545 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em devolucao para Unidade | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/546 | $pastaEngenharia | Documentos | Workflow - Engenharia | Em devolucao para Unidade | Everyone | na |
| DefinirAcessoEngenharia/548 | $pastaEngenharia | Documentos | Workflow - Engenharia | Arquivo interno Engenharia | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/549 | $pastaEngenharia | Documentos | Workflow - Engenharia | Arquivo interno Engenharia | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/550 | $pastaEngenharia | Documentos | Workflow - Engenharia | Arquivo interno Engenharia | Everyone | na |
| DefinirAcessoEngenharia/552 | $pastaEngenharia | Documentos | Workflow - Engenharia | Concluido pela Engenharia | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/553 | $pastaEngenharia | Documentos | Workflow - Engenharia | Concluido pela Engenharia | $nomeUserListEngenharia-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoEngenharia/554 | $pastaEngenharia | Documentos | Workflow - Engenharia | Concluido pela Engenharia | Everyone | na |
| DefinirAcessoEngenharia/556 | $pastaEngenharia | Documentos | Workflow - Engenharia | Superado | $nomeUserListEngenharia-CONSULTA | r, fr |
| DefinirAcessoEngenharia/557 | $pastaEngenharia | Documentos | Workflow - Engenharia | Superado | Everyone | na |
| DefinirAcessoEngenharia/559 | $pastaEngenharia | Documentos | Workflow - Comentarios | Markup | $nomeUserListEngenharia-GESTOR ENG | c, r, w, cw, fr, fw |
| DefinirAcessoEngenharia/560 | $pastaEngenharia | Documentos | Workflow - Comentarios | Markup | $nomeUserListEngenharia-ESPECIALISTA | c, r, w, cw, fr, fw |
| DefinirAcessoEngenharia/561 | $pastaEngenharia | Documentos | Workflow - Comentarios | Markup | Everyone | na |
| DefinirAcessoEngenharia/563 | $pastaEngenharia | Documentos | Workflow - Engenharia - Superados | Superado | $nomeUserListEngenharia-CONSULTA-TODAS REV | r, fr |
| DefinirAcessoEngenharia/564 | $pastaEngenharia | Documentos | Workflow - Engenharia - Superados | Superado | Everyone | na |
| DefinirAcessoUnidade/573 | $pastaUnidade | Pastas |  |  | Everyone | na |
| DefinirAcessoUnidade/574 | $pastaUnidade | Pastas |  |  | $nomeUserListUnidade | r |
| DefinirAcessoUnidade/575 | $pastaUnidade | Documentos |  |  | admin | fc |
| DefinirAcessoUnidade/577 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Em elaboracao | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/578 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Em elaboracao | $nomeUserListUnidade-Lider Unidade | c, r, w, cw, fr, fw |
| DefinirAcessoUnidade/579 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Em elaboracao | Everyone | na |
| DefinirAcessoUnidade/581 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Emitido pela Engenharia | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/582 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Emitido pela Engenharia | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/583 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Emitido pela Engenharia | Everyone | na |
| DefinirAcessoUnidade/585 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Solicitado reanalise da Engenharia | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/586 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Solicitado reanalise da Engenharia | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/587 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Solicitado reanalise da Engenharia | $nomeUserListUnidade-GESTOR ENG | r, w, cw, fr |
| DefinirAcessoUnidade/588 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Solicitado reanalise da Engenharia | Everyone | na |
| DefinirAcessoUnidade/590 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/591 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/592 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente | $nomeUserListUnidade-Poder Concedente | r, fr |
| DefinirAcessoUnidade/593 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente | Everyone | na |
| DefinirAcessoUnidade/595 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Concluido pela Unidade | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/596 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Concluido pela Unidade | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/597 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Concluido pela Unidade | Everyone | na |
| DefinirAcessoUnidade/599 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Nova emissao sendo analisada pela Engenharia | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/600 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Nova emissao sendo analisada pela Engenharia | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/601 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Nova emissao sendo analisada pela Engenharia | Everyone | na |
| DefinirAcessoUnidade/603 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente - Nova emissao em analise Eng | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/604 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente - Nova emissao em analise Eng | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/605 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente - Nova emissao em analise Eng | $nomeUserListUnidade-Poder Concedente | r, fr |
| DefinirAcessoUnidade/606 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Enviado ao Poder Concedente - Nova emissao em analise Eng | Everyone | na |
| DefinirAcessoUnidade/608 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Concluido - Nova emissao em analise Eng | $nomeUserListUnidade-CONSULTA | r, fr |
| DefinirAcessoUnidade/609 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Concluido - Nova emissao em analise Eng | $nomeUserListUnidade-Lider Unidade | r, w, cw, fr |
| DefinirAcessoUnidade/610 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Concluido - Nova emissao em analise Eng | Everyone | na |
| DefinirAcessoUnidade/612 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Superado | $nomeUserListUnidade-CONSULTA-TODAS REV | r, fr |
| DefinirAcessoUnidade/613 | $pastaUnidade | Documentos | Workflow - Engenharia - Unidade | Superado | Everyone | na |
| DefinirAcessoUnidade/615 | $pastaUnidade | Documentos | Workflow - Comentarios | Markup | $nomeUserListUnidade-Lider Unidade | c, r, w, cw, fr, fw |
| DefinirAcessoUnidade/616 | $pastaUnidade | Documentos | Workflow - Comentarios | Markup | $nomeUserListUnidade-GESTOR ENG | c, r, w, cw, fr, fw |
| DefinirAcessoUnidade/617 | $pastaUnidade | Documentos | Workflow - Comentarios | Markup | Everyone | na |
| DefinirAcessoLD/626 | $nomePastaRaiz | Pastas |  |  | Engenharia GPR | fc |
| DefinirAcessoLD/627 | $nomePastaRaiz | Pastas |  |  | $siglaConcessao-ENGENHARIA UNIDADE | fc |
| DefinirAcessoLD/629 | $nomePastaRaiz | Pastas |  |  | Everyone | na |
| DefinirAcessoLD/630 | $nomePastaRaiz | Pastas |  |  | $nomeUserList | r |
| DefinirAcessoLD/631 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-EDICAO | r, c, w |
| DefinirAcessoLD/633 | $nomePastaRaiz | Documentos |  |  | Everyone | na |
| DefinirAcessoLD/634 | $nomePastaRaiz | Documentos |  |  | $nomeUserList | r, fr |
| DefinirAcessoLD/635 | $nomePastaRaiz | Documentos |  |  | $nomeUserList-EDICAO | c, d, r, w, cw, fr, fw |
| DefinirAcessoGRD/644 | $nomePastaRaiz | Pastas |  |  | Engenharia GPR | fc |
| DefinirAcessoGRD/645 | $nomePastaRaiz | Pastas |  |  | $siglaConcessao-ENGENHARIA UNIDADE | fc |
| DefinirAcessoGRD/647 | $nomePastaRaiz | Pastas |  |  | Everyone | na |
| DefinirAcessoGRD/648 | $nomePastaRaiz | Pastas |  |  | $nomeUserList | r |
| DefinirAcessoGRD/649 | $nomePastaRaiz | Pastas |  |  | $nomeUserList-EDICAO | r, c, w |
| DefinirAcessoGRD/651 | $nomePastaRaiz | Documentos |  |  | Everyone | na |
| DefinirAcessoGRD/652 | $nomePastaRaiz | Documentos |  |  | $nomeUserList | r, fr |
| DefinirAcessoGRD/653 | $nomePastaRaiz | Documentos |  |  | $nomeUserList-EDICAO | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/664 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-ASSISTENTE ENG | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/665 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-CONSULTA TODAS REV | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/666 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-CONSULTA ULTIMA REV | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/667 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-ENG CONSULTA TODAS REV | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/668 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-ENG CONSULTA ULTIMA REV | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/669 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-ESPECIALISTA | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/670 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-GESTOR ENG | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/671 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-GESTOR UNIDADE | c, d, r, w, cw, fr, fw |
| DefinirAcessoXFDF/672 | $pastasXFDF | Documentos |  |  | $siglaConcessao-$projeto-PROJETISTA | c, d, r, w, cw, fr, fw |

## Pontos encontrados e limites

1. PROJETISTA e ligado somente a GRD-CONSULTA; nao e ligado a ENGENHARIA-PROJETISTA. A lista existe e recebe leitura em tres estados, mas esse grupo nao chega a ela pelo fluxo de criacao.
2. Engenharia GPR e global e acumula vinculos a cada projeto. S-ENGENHARIA UNIDADE e compartilhado pela concessao e acumula vinculos aos projetos dela.
3. CONSULTA-TODAS REV e aninhada em CONSULTA; assim recebe as regras de consulta comum mais as regras especificas de revisoes antigas. Contudo, Workflow - Engenharia, estado Superado, concede r/fr tambem a CONSULTA comum. A denominacao ULTIMA REV nao garante exclusao de todos os superados nesse workflow. Workflow - Engenharia - Superados concede acesso apenas a CONSULTA-TODAS REV; na Unidade, Superado tambem concede somente a CONSULTA-TODAS REV.
4. DefinirAcessoXFDF existe, mas nao e chamada por DefineAcessosProjeto. Suas regras de edicao/exclusao para nove grupos nao sao executadas por essa rotina. As pastas xfdf__ e de disciplina sao criadas por outro fluxo.
5. Poder Concedente recebe acesso pela lista UNIDADE-Poder Concedente e pelas listas superiores; tem leitura explicita nos dois estados de envio ao Poder Concedente. Nao aparece entre os grupos com regra direta de pasta na raiz, mas ha regra para a lista do projeto.
6. Nomes de gestor, assistente e projetista na planilha viram propriedades do projeto. Nao sao resolvidos como usuarios e nao ha chamada para inclui-los nos grupos.
7. Grupos e listas existentes sao reutilizados por nome exato. A rotina tenta adicionar os vinculos novamente; nao remove membros/vinculos anteriores nem atualiza a descricao de grupo existente.
8. Falhas em Add-PWMemberToUserList sao capturadas e registradas como avisos; nao impedem o retorno de sucesso do projeto. Nao ha transacao, rollback ou verificacao final completa.
9. Projeto existente pula criacao de pastas, mas processa listas, grupos e seguranca novamente. Isso pode alterar acessos de projeto existente.
10. O nome do projeto e limitado a 28 caracteres pelo executor. Nao ha validacao equivalente para o nome final de cada grupo, nem verificacao de colisoes entre combinacoes sigla/projeto.
11. Nao ha limpeza explicita das regras de seguranca preexistentes. Portanto, o codigo sozinho nao demonstra os acessos efetivos de cada pessoa.
12. 0 - Documentos Gerais nao possui funcao especifica de seguranca no fluxo; seus acessos dependem das regras herdadas/preexistentes e do workflow. admin recebe fc documental explicito na Area de Trabalho e na Unidade.

## Recomendacoes para uma proxima alteracao

- Confirmar a politica de acesso do projetista antes de vincular o grupo a ENGENHARIA-PROJETISTA.
- Confirmar se ULTIMA REV deve perder acesso ao estado Superado do Workflow - Engenharia.
- Consolidar as falhas de vinculo/seguranca e apresentar resultado parcial, em vez de sucesso simples.
- Conferir os membros dos grupos compartilhados e seus vinculos antes de usa-los como acesso abrangente.
- Auditar as regras efetivas no datasource antes de concluir que alguem possui apenas os acessos previstos nesta matriz.

Nenhuma rotina funcional foi alterada nesta analise.
