# Criar grupo ProjectWise e incluir participantes por Excel

Script: `17 - Criar Grupo e Incluir Participantes por Excel.ps1`.

Requisitos: Windows PowerShell, modulos `PWPS_DAB` e `ImportExcel`, acesso ao datasource e autorizacao para administrar grupos. Nao requer Microsoft Excel instalado. O script abre o login do ProjectWise.

A primeira linha da planilha deve conter `Email` ou `E-mail` (tambem aceita `Mail`). Uma pessoa por linha. A coluna `Nome` e opcional e informativa; a identificacao e feita pelo e-mail exato. Use `.xlsx` ou `.xlsm`. Sem `-NomeAba`, ImportExcel le a primeira aba.

Execucao interativa: execute o script e informe o nome do grupo; selecione a planilha e autentique-se no datasource correto.

```powershell
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1'
```

Gerar modelo (nao conecta ao ProjectWise e nao sobrescreve arquivo existente):

```powershell
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1' -GerarModelo -CaminhoExcel 'C:\Dados\Modelo_Grupo.xlsx'
```

Verificar os participantes sem criar grupo nem incluir membros (consulta o ProjectWise):

```powershell
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1' -NomeGrupo 'ARTESP' -CaminhoExcel 'C:\Dados\Participantes.xlsx' -SomenteValidar
```

Executar inclusao:

```powershell
& '.\17 - Criar Grupo e Incluir Participantes por Excel.ps1' -NomeGrupo 'ARTESP' -DescricaoGrupo 'Participantes ARTESP' -CaminhoExcel 'C:\Dados\Participantes.xlsx' -NomeAba 'Participantes'
```

O grupo existente e reutilizado sem remover membros nem alterar sua descricao. Um grupo novo so e criado quando houver um participante valido a incluir. Usuarios ausentes nao sao criados. E-mail ambiguo ou invalido gera erro na linha; o lote continua. E-mails repetidos sao registrados como duplicados. A inclusao e confirmada consultando os membros do grupo.

Relatorio CSV separado por ponto e virgula na pasta `Logs`, com linha, grupo, e-mail, login, status e mensagem. A pasta pode ser alterada por `-PastaRelatorios`. Erros iniciais (planilha, login ou consulta do grupo) interrompem a execucao; erros individuais constam no relatorio. Nao ha rollback: inclusoes ja confirmadas permanecem se ocorrer falha posterior.

O script nao atribui permissoes a pastas/documentos. Se o grupo reutilizado ja tiver permissoes, novos membros passam a receber os acessos associados a ele. Este script e independente e nao foi adicionado ao menu da Central.

Testes locais sem conexao:

```powershell
powershell -NoProfile -File '.\Tests\Test-CriarGrupoExcel.ps1'
```
