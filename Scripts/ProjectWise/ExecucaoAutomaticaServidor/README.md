# Rotinas automáticas do ProjectWise

## Autenticação

As rotinas automáticas não armazenam mais senha no código. Antes de executá-las,
configure `ECORODOVIAS_PW_PASSWORD` no contexto da conta de serviço usada pelo
Agendador de Tarefas.

Na configuração definitiva, restrinja a leitura da variável à conta de serviço
e não registre seu valor em scripts, argumentos, logs ou histórico. Como a senha
anterior estava versionada, ela deve ser trocada no ProjectWise.

As execuções abertas pela Central usam o login gráfico e não precisam dessa
variável.
