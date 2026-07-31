# Plano de migração

## Regras

1. Os códigos em `ExecucaoManualUsuario` são referências originais e não serão alterados pela migração.
2. Cada rotina terá arquivos novos em sua subpasta dentro de `Rotinas`.
3. O catálogo continuará usando `modo: legado` até a versão integrada ser testada e aprovada.
4. Cada migração deve validar entradas, cancelamento, erros, resultado e geração de logs.
5. A troca para `modo: integrado` ocorrerá individualmente, mantendo possibilidade de retorno ao original.

## Ordem proposta

1. Criar usuários em lote
2. Gerenciar acessos de projetos
3. Incluir usuários em projetos PW Web
4. Gerenciar participantes PWDM
5. Gerenciar usuários inativos
6. Criar projetos no ProjectWise
7. Criar buscas salvas

O item “Criar projetos no ProjectWise” apareceu duas vezes na relação inicial e foi considerado uma única rotina.
