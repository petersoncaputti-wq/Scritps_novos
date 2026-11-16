# Downloader Engemap Cloud

Python 3.10+ sem dependencias externas. Adaptacao independente do fluxo de streaming,
arquivo parcial e retomada Range/If-Range usado no downloader SharePoint.
O downloader antigo permanece intacto.

No PowerShell, nesta pasta:

```powershell
.\Executar-Engemap.ps1 -Listar
.\Executar-Engemap.ps1
.\Executar-Engemap.ps1 -Destino 'D:\Downloads\Engemap'
```

Sem `-Destino`, o executavel PowerShell abre uma janela para selecionar ou criar
a pasta de destino. Cancelar fecha o fluxo sem iniciar downloads. `-Listar` nao
abre o seletor. O parametro `-Destino` continua disponivel para automacao.

Digite a senha quando solicitada. Ela fica apenas em memoria; nao e salva em arquivos
ou passada na linha de comando. Use o mesmo destino para continuar depois.
Tambem e possivel executar `python downloader.py --destino "D:\Downloads\Engemap"`.
Para outro compartilhamento, informe `--url https://servidor/s/token`.

Preserva a hierarquia e os nomes originais validos no Windows, mantendo os vinculos
entre TIFF/TFW e SHP/DBF/SHX/PRJ. Nomes incompativeis ou longos sao normalizados
com sufixo; colisoes de nomes sao recusadas antes de baixar. O inventario mantem
os nomes remotos originais.
Cria inventario JSON, estado JSON, log e arquivos temporarios `.engemap.part`.
Confere tamanho e calcula SHA-256 local; em novas execucoes confere o hash antes
de pular arquivos concluidos. O hash nao e uma verificacao contra checksum remoto.
Retoma parciais somente com ETag forte e identidade compativel; se o servidor ignora
Range, reinicia o arquivo. Sem ETag nao retoma nem pula automaticamente.
Arquivos finais existentes divergentes sao preservados e reportados como falha;
mova-os ou use outro destino para baixar a nova versao.

Requer compartilhamento com permissao de download e acesso WebDAV publico.
Tenta o endpoint moderno e o legado. HTTP 401/403 pode indicar senha incorreta,
permissao insuficiente ou WebDAV publico desabilitado. Nao contorna essas restricoes.
Downloads sao sequenciais, com ate tres tentativas para falhas de rede.
Ctrl+C preserva os parciais. A trava `.engemap.lock` impede execucoes simultaneas;
se houver encerramento forcado, remova essa trava somente apos confirmar que nao
existe outra execucao no mesmo destino.

Testes locais: `python -m unittest -v test_downloader.py`.
O acesso real sem senha retornou HTTP 401; a validacao ponta a ponta com o
compartilhamento depende de o usuario informar a senha localmente.
