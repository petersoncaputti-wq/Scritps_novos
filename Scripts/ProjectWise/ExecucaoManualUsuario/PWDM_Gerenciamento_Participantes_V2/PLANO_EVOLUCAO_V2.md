# Evolucao do Gerenciamento de Participantes PWDM

## Objetivo

Transformar o script atual em uma ferramenta mais segura, testavel e visual, reaproveitando as partes que ja funcionam: autenticacao manual via navegador, leitura do catalogo PWDM/Connect, cruzamento ProjectWise x PWDM, montagem de operacoes, aplicacao via endpoints e geracao de logs.

## Diretrizes

- Evoluir por fases pequenas, com teste antes de seguir para a proxima funcao.
- Separar regras de negocio da interface visual.
- Manter o fluxo atual executavel durante a transicao.
- Evitar automatizar a exclusao ate que o endpoint destrutivo esteja capturado, validado e coberto por uma etapa explicita de confirmacao.
- Tratar `session.json` como credencial sensivel e nunca depender dele em operacao real.

## Fase 1 - Base de testes e regras puras

Escopo:
- Cobrir regras que nao dependem de navegador: permissao, decisao de acao, descricao, mascaramento, extracao de IDs e comparacao de titulo/permissoes.
- Criar uma base para refatorar sem alterar comportamento.

Criterio de aceite:
- Testes unitarios passam localmente.
- Nenhuma chamada ao PWDM/Connect e feita durante os testes.

## Fase 2 - Extrair motor PWDM

Escopo:
- Separar funcoes em modulos por responsabilidade:
  - `config.py`: constantes e caminhos.
  - `regras.py`: validacoes, permissoes e decisao de acoes. Concluido parcialmente.
  - `projectwise.py`: leitura e selecao da arvore ProjectWise.
  - `pwdm_api.py`: endpoints, tokens, GET/POST e leitura de participantes.
  - `logs.py`: sanitizacao e exportacao JSON/CSV.
  - `cli.py`: fluxo atual em terminal.
- Manter `gerenciar_participante_pwdm.py` como entrada compativel.

Criterio de aceite:
- Fluxo atual continua funcionando pelo terminal.
- Testes da Fase 1 continuam passando.

Progresso:
- Criado `regras.py` com regras puras de permissao, acao efetiva, descricao de permissoes, normalizacao de texto, extracao de IDs de URL, mascara de e-mail e comparacoes simples.
- `gerenciar_participante_pwdm.py` passou a importar essas regras.
- Testes unitarios passaram a validar `regras.py` diretamente.

Proximo recorte sugerido:
- Extrair `logs.py` com sanitizacao, status de resultado e exportacao JSON/CSV.
- Depois extrair `pwdm_api.py`, que e mais sensivel por envolver navegador, token e endpoints.

## Fase 3 - Interface visual inicial

Escopo:
- Criar uma interface visual para entrada de dados e previa:
  - e-mail do usuario;
  - acao: incluir ou alterar;
  - titulo/cargo;
  - permissoes;
  - selecao de concessao/projetos;
  - painel de log;
  - previa das operacoes antes de aplicar.
- Usar uma abordagem simples para Windows, alinhada ao padrao visual do script `ProjectWise_Gestao_Acessos_FallbackInteligente.ps1`.

Criterio de aceite:
- Usuario consegue montar e revisar operacoes pela interface.
- Aplicacao continua exigindo confirmacao final antes de alterar PWDM.

## Fase 4 - Execucao assistida e resiliencia

Escopo:
- Melhorar mensagens de erro por projeto.
- Exibir resumo por status: atualizado, adicionado, sem alteracao, ignorado e erro.
- Permitir exportar a previa antes da execucao.
- Preservar logs sanitizados.

Criterio de aceite:
- Falha em um projeto nao impede o processamento dos demais.
- Resultado final mostra claramente o que foi aplicado e o que falhou.

## Fase 5 - Funcoes avancadas

Escopo:
- Inclusao em lote por planilha.
- Reprocessamento de falhas a partir do CSV/log.
- Exclusao somente apos diagnostico, validacao do endpoint e confirmacao reforcada.

Criterio de aceite:
- Cada funcao nova tem teste automatizado para a regra de decisao e teste manual documentado para a parte PWDM.
