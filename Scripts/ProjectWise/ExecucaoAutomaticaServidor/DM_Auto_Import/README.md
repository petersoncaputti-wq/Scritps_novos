# DM_Auto_Import

Projeto Python para diagnosticar, validar e futuramente automatizar o fluxo de Incoming/Acknowledge/Import do ProjectWise Deliverables Management.

O projeto foi pensado para evoluir em fases, com execucao futura por agendamento em servidor Windows. A Fase 1 nao importa nada automaticamente: ela apenas abre o navegador, permite login e navegacao manual, captura requests/responses relevantes e salva um JSON de diagnostico.

## Objetivo

Automatizar de forma segura, auditavel e controlada o processamento de entregaveis Incoming no ProjectWise Deliverables Management para multiplos projetos/DMs.

Hoje os entregaveis chegam na tela Incoming e precisam ser aceitos/importados manualmente. O robo sera desenvolvido por etapas para identificar pendencias e processar somente entregaveis elegiveis, com controle de duplicidade e sem acoes destrutivas sem validacao explicita.

## Fases planejadas

- Fase 1: diagnostico de Incoming/import, sem automacao de aceite/importacao.
- Fase 2: automatizar 1 projeto de teste apos validacao dos endpoints.
- Fase 3: criar e validar a base `projetos_monitorados.json`.
- Fase 4: rodar em lote para 10 projetos.
- Fase 5: expandir para 100+ projetos.

## Estrutura

```text
DM_Auto_Import/
  README.md
  requirements.txt
  diagnostico_dm_incoming_import.py
  config.py
  Logs/
  data/
    projetos_monitorados.json
    controle_processamento.json
```

## Instalar dependencias

```powershell
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## Executar a Fase 1

```powershell
python diagnostico_dm_incoming_import.py
```

Por padrao, o script usa um perfil persistente do Chromium em `data/browser_profile`. Isso permite manter cookies/sessao entre testes, reduzindo a necessidade de login repetido. Nao salve essa pasta em repositorio e nao compartilhe esse conteudo.

Fluxo de uso:

1. Rode o script.
2. Faca login no navegador aberto pelo Playwright.
3. Abra um projeto/DM de teste.
4. Acesse Incoming/Submittals.
5. Abra um submittal pendente.
6. Execute manualmente o acknowledge/import.
7. Volte ao terminal e pressione ENTER.
8. O script salvara um JSON em `Logs`.
9. O JSON sera analisado depois para definir a Fase 2.

## Logs

Os logs sao salvos em `Logs` com timestamp no nome do arquivo.

O JSON inclui:

- timestamp;
- objetivo;
- observacoes;
- totalEventos;
- totalEventosProvavelmenteDM;
- eventosProvavelmenteDM;
- eventosCompletos.

O script mascara dados sensiveis quando detectados, incluindo:

- `authorization`;
- `cookie`;
- `set-cookie`;
- tokens;
- request verification tokens;
- e-mails.

GUIDs sao mantidos no log, pois serao usados para identificar `projectId`, `connectSpaceId` e `submittalId`.

## Cuidados

- Nao compartilhe logs sem revisar.
- Nao salve `session.json` em repositorio.
- Nao salve `data/browser_profile` em repositorio.
- Nao ative importacao automatica antes da validacao do endpoint.
- Nao automatize exclusao nem acoes destrutivas sem validacao explicita.
- Use `autoImportar: false` por padrao enquanto os endpoints e criterios de elegibilidade nao forem validados.
