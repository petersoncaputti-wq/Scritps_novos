# Criar usuários em lote

Nova implementação para execução pela Central.

- `CriarUsuariosLote.Integrado.ps1`: cópia de trabalho desacoplada do console, com parâmetros próprios para a Central.
- `Painel.CriarUsuariosLote.ps1`: painel WinForms carregado dentro da janela principal.

No modo integrado, a lista não é exportada automaticamente. Ao concluir, a Central
exibe o resumo, pergunta se o usuário deseja exportar e, somente após resposta
afirmativa, solicita a pasta de destino.

O original permanece preservado em `ExecucaoManualUsuario\03 - Criar Usuarios ProjectWise em Lote.ps1`.
