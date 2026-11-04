# Kartado — cadastro em lote, versao de teste

Python 3.10 ou superior. Instale a dependencia com `python -m pip install -r .\requirements.txt`. Esta versao le arquivos Excel .xlsx. Unidade e perfil sao unicos por lote.

Na primeira linha do Excel, use as colunas nome, sobrenome, email e login (uma coluna para cada campo). Exemplo:

```text
nome;sobrenome;email;login
Ana;Silva;ana.silva@example.com;ana.silva
```

Salve como Pasta de Trabalho do Excel (.xlsx), sem converter para CSV. Arquivos antigos .xls devem ser salvos como .xlsx. A leitura usa a aba ativa; para escolher outra, informe --aba "Nome da aba". Cabecalhos podem ter maiusculas e espacos nas extremidades; e-mail tambem e aceito. Linhas vazias sao ignoradas. Formulas nos campos de cadastro sao recusadas; cole como valores. Nao inclua senhas no arquivo.

Execute os comandos nesta pasta:

```powershell
# Valida arquivo local, sem conexao
python .\criar_usuarios.py --arquivo 'C:\caminho\usuarios.xlsx' --validar

# Apresenta previa e pede aprovacao antes de criar
python .\criar_usuarios.py --arquivo 'C:\caminho\usuarios.xlsx'

# Consulta e mostra somente a previa, sem oferecer criacao
python .\criar_usuarios.py --arquivo 'C:\caminho\usuarios.xlsx' --somente-previa
```

Alternativa interativa: `powershell -NoProfile -File .\Executar-Kartado.ps1`; apos a previa, digite `CRIAR` para aprovar ou qualquer outra resposta para cancelar. Nao e necessario usar --executar. Para somente consultar, use `-SomentePrevia`.

O iniciador abre uma janela visual para selecionar o arquivo .xlsx. Cancelar encerra sem conectar a API. Tambem pode executar `python .\criar_usuarios.py` para abrir a janela, ou `python .\criar_usuarios.py --validar` para selecionar e validar sem conexao. O parametro `--arquivo` continua disponivel para uso sem janela. O seletor usa tkinter/Tcl-Tk, normalmente incluido no instalador do Python para Windows.

Para iniciar pelo Explorador de Arquivos, abra `Abrir-Kartado.bat`. O iniciador mantem a janela aberta ate pressionar ENTER, inclusive em caso de erro. Falhas do Python geram `Logs/diagnostico_*.json` com o tipo de erro, mensagem e local da falha, sem variaveis locais ou credenciais. Para automacao sem pausa, use o iniciador PowerShell com `-SemPausa`.

O login e a senha da API sao solicitados a cada execucao. A senha fica oculta e nao e gravada. Use a senha atualizada da conta; a credencial exposta na conversa nao foi incorporada ao codigo.

O ambiente e fixado em `https://eco.api.kartado.com.br/`, conforme a captura real. A autenticacao Basic foi indicada pela documentacao, mas ainda precisa ser validada nesse ambiente. Se houver 401/403, confira com o Kartado a autenticacao e as permissoes da conta. Nao cole cookies de sessao no script.

O grupo e obtido de `relationships.companyGroup.data.id` da unidade quando disponivel. Caso contrario, informe manualmente o UUID confirmado para a unidade. Pode fornecer `--grupo UUID`. O grupo observado no cadastro de Ecovias Noroeste Paulista foi `6d356a3c-8adf-49f7-b545-528da6138f6f`; nao e aplicado automaticamente a outras unidades.

O POST reproduz o payload capturado: `createAccessRequest: true`, relacionamentos `company`, `companyGroup`, `permissions`. Um HTTP 201 confirma criacao, mas nao aprovacao/liberacao de acesso. Confira o usuario e a solicitacao no portal. O nome Administrador nao implica `isAdmin=true`; a consulta preserva `is_admin=false`, como no portal observado.

Antes de criar, o script consulta os usuarios da unidade selecionada com os filtros `company`, `only_company` e `sort=first_name`, conforme a requisicao observada no portal. Percorre a paginacao e ignora logins encontrados sem diferenciar maiusculas. Essa consulta nao e global: usuarios cadastrados em outras unidades podem nao aparecer. Nao altera usuarios existentes, nao concede acesso adicional a eles e nao deduplica por email (Kartado permite email compartilhado). Conflitos retornados pelo servidor sao registrados. A correcao dos filtros deve ser validada no ambiente real; nao comprova por si so que a autenticacao Basic tera acesso ao endpoint.

Cada resultado e persistido imediatamente em `Logs/resultado_*.jsonl`, com linha, login, unidade, perfil, status e ID criado. O relatorio contem identificadores de usuarios: mantenha-o em pasta de acesso adequado. Respostas completas e credenciais nao sao gravadas. Nao ha repeticao automatica de POST. Falhas de rede/JSON deixam o resultado incerto e interrompem o lote; confira no portal antes de repetir. Falhas HTTP 401, 403, 429 ou 5xx tambem interrompem o processamento. Linhas posteriores nao sao enviadas. Erros 4xx de cadastro sao registrados e o lote prossegue.

Para o primeiro teste, use apenas um usuario necessario, revise unidade, grupo e perfil na previa e verifique a aprovacao no portal antes de aumentar o lote. Esta implementacao foi verificada localmente; nenhuma conexao ou criacao real foi executada durante o desenvolvimento.

Testes locais:

```powershell
python -m unittest discover -s . -p 'test_*.py' -v
```


