# PW-SharePoint-Importer

Importador faseado para copiar arquivos de uma pasta do SharePoint para uma pasta do ProjectWise.

## Fase atual

Esta primeira fase cria somente:

- Estrutura inicial do projeto.
- Configuracao em `config/settings.json`.
- Funcao `Connect-GraphSharePoint`.
- `src/Main.ps1` chamando apenas a conexao com Microsoft Graph.

Ainda nao foram implementados download, conexao ProjectWise, importacao ou logs finais.

## Requisitos da fase 1

- PowerShell 7.
- Modulo `Microsoft.Graph`.

Para instalar ou atualizar o PowerShell 7 no Windows:

```powershell
winget install --id Microsoft.PowerShell --source winget
```

Para instalar o modulo:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

## Como configurar

Edite `config/settings.json` e configure os dados do tenant e os escopos do Microsoft Graph.

Exemplo:

```json
{
  "SharePoint": {
    "TenantUrl": "https://grupoecorodovias.sharepoint.com",
    "TenantId": "bf7db00b-6dcc-41f4-8148-8653bb3d6537",
    "Scopes": [
      "Sites.Read.All",
      "Files.Read.All"
    ]
  }
}
```

## Sobre permissoes do Microsoft Graph

O Microsoft Graph PowerShell abre login interativo usando o SDK oficial da Microsoft.
Mesmo assim, dependendo da politica da empresa, os escopos `Sites.Read.All` e `Files.Read.All`
podem exigir consentimento de administrador.

## Como testar no VS Code

1. Abra a pasta `PW-SharePoint-Importer` no VS Code.
2. Abra o terminal integrado.
3. Confirme que esta usando PowerShell 7:

```powershell
$PSVersionTable.PSVersion
```

4. Se o modulo `Microsoft.Graph` ainda nao estiver instalado, instale no terminal do PowerShell 7:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

5. Configure `TenantUrl`, `TenantId` e `Scopes` em `config/settings.json`.
6. Execute:

```powershell
.\src\Main.ps1
```

7. Faca login na janela interativa da Microsoft.

Resultado esperado:

- `Success: True`
- `Step: Connect-GraphSharePoint`
- `Account` preenchido com o usuario autenticado.

Se ocorrer erro, o retorno exibira:

- `Success: False`
- `Step`
- `Message`
- `Error.Message`
- `Error.Type`
