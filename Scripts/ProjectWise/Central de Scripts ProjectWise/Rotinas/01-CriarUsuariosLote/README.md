# Criar usuários em lote

Nova implementação para execução pela Central.

- `CriarUsuariosLote.Integrado.ps1`: cópia de trabalho desacoplada do console, com parâmetros próprios para a Central.
- `Painel.CriarUsuariosLote.ps1`: painel WinForms carregado dentro da janela principal.

No modo integrado, a lista não é exportada automaticamente. Ao concluir, a Central
exibe o resumo, pergunta se o usuário deseja exportar e, somente após resposta
afirmativa, solicita a pasta de destino.

O original permanece preservado em `ExecucaoManualUsuario\03 - Criar Usuarios ProjectWise em Lote.ps1`.

## Identificacao e nomes alternativos

O cadastro existente e identificado pelo e-mail exato, ignorando maiusculas e espacos nas extremidades. Quando o nome/login pertence a outro e-mail (ou o cadastro existente nao possui e-mail), uma janela solicita um nome alternativo. Nomes vazios ou ocupados repetem a solicitacao. Cancelar ou fechar a janela registra erro apenas naquela linha e o lote continua.

O cadastro da outra pessoa nao e atualizado nem recebe acessos, inclusive com "Atualizar usuarios existentes" marcado. O resultado mostra o login efetivamente utilizado e informa na mensagem quando ele difere do solicitado. E-mails encontrados em mais de um cadastro bloqueiam a linha para revisao; falhas de consulta tambem interrompem a linha.

O modo "Somente validar" verifica a planilha sem consultar o ProjectWise ou solicitar nomes alternativos.

Teste local sem conexao: `powershell -NoProfile -File .\Tests\Test-IdentificacaoUsuarios.ps1`.
