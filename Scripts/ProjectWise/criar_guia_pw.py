from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUT = Path(__file__).parent / 'Documentacao'
OUT.mkdir(exist_ok=True)
d = Document()
s = d.sections[0]
s.page_width, s.page_height = Cm(21), Cm(29.7)
s.top_margin = s.bottom_margin = Cm(2)
s.left_margin = s.right_margin = Cm(2.2)
for name in ['Normal', 'Title', 'Subtitle', 'Heading 1', 'Heading 2']:
    st=d.styles[name]
    st.font.name='Calibri'
    st.font.color.rgb=RGBColor(0,0,0)
d.styles['Normal'].font.size=Pt(11)
d.styles['Normal'].paragraph_format.space_after=Pt(7)
d.styles['Normal'].paragraph_format.line_spacing=1.12
d.styles['Title'].font.size=Pt(26)
d.styles['Heading 1'].font.size=Pt(18)
d.styles['Heading 2'].font.size=Pt(12)
for name in ['Heading 1','Heading 2']:
    d.styles[name].paragraph_format.keep_with_next=True
    d.styles[name].paragraph_format.space_before=Pt(12)
    d.styles[name].paragraph_format.space_after=Pt(6)
def p(t): d.add_paragraph(t)
def h(t): d.add_heading(t,2)
def bullets(items):
    for t in items: d.add_paragraph(t,'List Bullet')
def chapter(t):
    d.add_page_break()
    d.add_heading(t,1)
def table(headers, rows):
    t=d.add_table(rows=1, cols=len(headers)); t.autofit=False
    for c,txt in zip(t.rows[0].cells,headers): c.text=txt
    rep=OxmlElement('w:tblHeader');t.rows[0]._tr.get_or_add_trPr().append(rep)
    for row in rows:
        for c,txt in zip(t.add_row().cells,row): c.text=txt
    for i,row in enumerate(t.rows):
        for c in row.cells:
            pr=c._tc.get_or_add_tcPr()
            sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'E7EDF3' if i==0 else ('F5F5F5' if i%2==0 else 'FFFFFF'));pr.append(sh)
            borders=OxmlElement('w:tcBorders')
            for edge in ['top','left','bottom','right']:
                e=OxmlElement('w:'+edge);e.set(qn('w:val'),'single');e.set(qn('w:sz'),'4');e.set(qn('w:color'),'D9D9D9');borders.append(e)
            pr.append(borders)
            for para in c.paragraphs:
                para.paragraph_format.space_after=Pt(6)
                para.paragraph_format.space_before=Pt(6)
                for r in para.runs: r.font.size=Pt(10);r.bold=(i==0)
    d.add_paragraph()

d.add_heading('Guia introdutório ao ProjectWise',0)
d.add_paragraph('Gestão e controle de documentos', 'Subtitle')
p('Material de apoio ao treinamento • 22 de outubro de 2026')
h('Objetivo e aplicação')
p('Este guia apresenta os conceitos e as operações básicas do ProjectWise, desde o acesso até o acompanhamento dos fluxos documentais. Explica também como as automações internas apoiam a organização, a validação e a distribuição de documentos.')
p('Ao concluir a leitura, o participante deverá conseguir localizar documentos, compreender o ciclo de edição, reconhecer as regras de organização e identificar a próxima ação em uma pendência documental.')
p('As telas, os comandos e as operações disponíveis variam conforme a versão, o acesso pelo ProjectWise Explorer ou Web e as permissões do usuário. Os procedimentos de cada projeto devem orientar a execução das atividades.')
h('Roteiro de leitura e treinamento')
table(['Etapa','Conteúdo','Tempo'],[
('1','Apresentação do ProjectWise','15 min'),('2','Acesso e navegação','20 min'),('3','Gestão de documentos','40 min'),('4','Pesquisa e localização','20 min'),('5','Governança e boas práticas','25 min'),('6','Fluxos documentais','20 min')])
p('Carga horária de referência: 140 minutos, equivalentes a 2h20, sem intervalo. Os tempos incluem explicações e demonstrações introdutórias.')

chapter('1 Apresentação do ProjectWise')
h('1.1 O que é o ProjectWise')
p('O ProjectWise, da Bentley Systems, é uma solução para gerenciar informações e documentos de projetos de infraestrutura. Ele permite organizar arquivos em um ambiente compartilhado, controlar acessos e acompanhar alterações durante o desenvolvimento dos trabalhos.')
p('Um documento reúne o arquivo e as informações utilizadas para identificá-lo e gerenciá-lo. Essas informações podem incluir descrição, disciplina, revisão, versão e situação no fluxo de trabalho. Assim, é possível compreender a qual projeto um desenho, relatório ou memorial pertence e em que etapa do processo se encontra.')
h('1.2 Objetivos e benefícios')
bullets(['Centralizar documentos em um ambiente compartilhado.','Organizar o conteúdo por projetos, pastas e disciplinas.','Controlar o acesso e as operações de edição conforme as responsabilidades.','Consultar versões e informações históricas.','Localizar documentos por pesquisas e atributos.','Acompanhar as etapas do fluxo documental.'])
p('Esses recursos ajudam a reduzir o uso de arquivos desatualizados, a duplicação desnecessária de documentos e o tempo gasto para localizar informações.')
h('1.3 Papel das automações internas')
p('Os scripts internos complementam o ProjectWise com regras específicas da organização. As rotinas consultadas incluem criação de projetos, gerenciamento de acessos, buscas salvas, validação de documentos, distribuição por disciplina e geração de GRD.')
p('O usuário precisa compreender quais informações deve preencher e como suas ações influenciam o processamento. Um atributo incorreto, por exemplo, pode impedir a identificação da pasta de destino.')
h('Exemplo de aplicação')
p('Ao receber uma nova entrega, a equipe registra os documentos, informa seus atributos e acompanha a situação de processamento. As verificações automáticas apoiam o controle documental; a análise técnica continua sob responsabilidade da equipe designada.')

chapter('2 Acesso e navegação')
h('2.1 Acesso ao ambiente')
p('O acesso exige uma conta habilitada e permissões compatíveis com a atividade. A utilização pode ocorrer pelo ProjectWise Explorer, instalado no computador, ou pela interface Web disponibilizada pela organização.')
p('No Explorer, pode ser necessário selecionar a fonte de dados, chamada datasource, antes da autenticação. Na interface Web, o acesso ocorre pelo endereço informado pela empresa. Após entrar, o usuário visualiza o conteúdo permitido para seu perfil.')
p('Se um projeto não aparecer, confira o ambiente selecionado e a concessão de acesso. Estar autenticado não significa ter autorização para consultar todo o conteúdo.')
h('2.2 Organização da interface')
p('A interface normalmente oferece uma área de navegação por projetos e pastas e uma área de listagem dos documentos. As colunas disponíveis podem apresentar nome, descrição, revisão, versão, estado e outras propriedades.')
bullets(['Confira o projeto e a pasta selecionados.','Leia o nome e a descrição do documento.','Verifique revisão, versão e estado documental.','Observe se existe uma reserva para edição.'])
h('2.3 Estrutura de projetos e pastas')
p('A estrutura segue o padrão da empresa e pode considerar empreendimento, contrato, disciplina, tipo documental ou etapa do processo. A localização do arquivo integra seu contexto de gestão.')
p('Nos scripts internos há rotinas de criação de projetos e pastas de disciplinas, além de processamento que utiliza atributos para definir destinos. Inserir um documento em local inadequado pode dificultar a consulta e interferir no fluxo esperado.')
h('2.4 Propriedades e atributos')
p('As propriedades apresentam informações sobre o documento. Os atributos são campos configurados para identificação, pesquisa e aplicação das regras internas, como disciplina, revisão e número documental.')
p('O nome do arquivo e os atributos devem ser coerentes. Divergências podem prejudicar a pesquisa, a validação e o encaminhamento automático.')
h('Prática sugerida')
p('Acesse um projeto de treinamento, localize a pasta de uma disciplina e consulte as propriedades de um documento. Identifique quais informações aparecem na listagem e quais estão disponíveis nas propriedades.')

chapter('3 Gestão de documentos')
h('3.1 Inclusão de documentos ou upload')
p('A inclusão consiste em inserir o arquivo no ProjectWise e registrar as informações necessárias ao gerenciamento. Antes de incluir, confirme o projeto, a pasta, a nomenclatura e a revisão correspondente à entrega.')
p('Verifique se o documento já existe. Conforme o procedimento interno, uma nova entrega pode exigir atualização ou criação de versão do registro existente. Os atributos obrigatórios precisam ser preenchidos com atenção, pois alimentam buscas, validações e encaminhamentos.')
h('3.2 Download e consulta')
p('O download disponibiliza uma cópia no computador. Essa cópia representa o conteúdo obtido naquele momento e pode ficar desatualizada se houver alterações posteriores no ambiente.')
p('Antes de utilizar um arquivo baixado, confira sua revisão e situação no ProjectWise. Obter uma cópia para consulta e reservar um documento para edição são operações com finalidades diferentes. O download, por si só, não deve ser interpretado como reserva ou autorização para edição.')
h('3.3 Abertura de documentos')
p('A abertura pode ocorrer em um visualizador ou em uma aplicação associada ao arquivo. O comportamento depende da integração e da operação escolhida. Em aplicações integradas, abrir para edição pode envolver a reserva do documento. Observe as mensagens para identificar se está consultando ou editando.')
h('3.4 Check-out ou reserva para edição')
p('O check-out reserva o documento para edição e, em geral, impede alterações simultâneas na mesma cópia gerenciada. Isso ajuda a evitar conflitos e perda de trabalho.')
p('Antes de reservar, confirme o documento e sua autorização para alterá-lo. Mantenha a reserva somente pelo tempo necessário, pois ela pode impedir que outras pessoas deem continuidade ao trabalho.')
h('3.5 Check-in ou devolução após a edição')
p('O check-in devolve o arquivo editado ao ProjectWise e libera a reserva correspondente. Salvar no aplicativo de edição não deve ser considerado, por si só, confirmação de devolução ao ambiente.')
for text in ['Localize e confira o documento.','Realize o check-out.','Edite e salve o arquivo.','Realize o check-in.','Confirme a conclusão e a liberação da reserva.']:
    d.add_paragraph(text,'List Number')
p('Quando disponível, utilize o comentário de devolução para registrar o motivo ou o resumo da alteração.')
h('3.6 Cancelamento de uma edição')
p('Se a edição não precisar ser mantida, pode existir uma operação para desfazer ou liberar o check-out. O nome e o comportamento variam conforme a interface. Confira o efeito sobre as alterações locais antes de executar, pois o trabalho pode ser descartado.')
h('3.7 Versão e revisão')
p('A versão está ligada ao controle das ocorrências do documento no sistema. A revisão identifica uma evolução ou emissão documental conforme o padrão adotado pela organização.')
p('Um check-in não significa necessariamente a emissão de nova revisão. Alterar a revisão no nome do arquivo também não garante que o controle de versões esteja correto. Os scripts internos comparam informações de revisão e versão com valores calculados, reforçando a necessidade de coerência entre arquivo e atributos.')

chapter('4 Pesquisa e localização de documentos')
h('4.1 Pesquisa por nome')
p('A pesquisa por nome é útil quando se conhece o código completo ou parte dele. Critérios mais precisos facilitam a identificação; buscas amplas podem retornar documentos de projetos, disciplinas ou revisões diferentes.')
p('Confira o escopo da consulta. Uma pesquisa limitada à pasta atual pode deixar de apresentar um documento localizado em outra parte da estrutura.')
h('4.2 Pesquisa por atributos')
p('Quando o nome não é conhecido, os atributos permitem pesquisar características cadastradas. Conforme a configuração, podem ser usados disciplina, revisão, número documental, estado e outros campos.')
p('Combinar disciplina e situação documental, por exemplo, pode ajudar a identificar os documentos que exigem atenção de uma equipe. A qualidade do resultado depende do preenchimento dos dados.')
h('4.3 Filtros e interpretação dos resultados')
p('Os filtros restringem a consulta. Antes de utilizar um resultado, confirme projeto, revisão e situação. O primeiro item encontrado não deve ser automaticamente considerado o documento aplicável à atividade.')
bullets(['Revise o escopo selecionado quando não houver resultados.','Confira a grafia e a identificação utilizadas.','Reduza a combinação de filtros para investigar a busca.','Verifique se possui permissão para consultar o conteúdo.'])
h('4.4 Buscas salvas')
p('As buscas salvas armazenam critérios para consultas recorrentes. Elas apoiam a localização de documentos e o acompanhamento de pendências.')
p('Uma busca salva é uma consulta reutilizável: seus resultados podem mudar conforme os documentos são incluídos, atualizados ou movimentados no fluxo. A rotina interna de criação de buscas contempla consultas por estado documental, incluindo situações relacionadas à análise e à validação.')
h('Prática sugerida')
p('Localize um documento pelo nome e depois por atributos. Consulte uma busca salva e identifique o critério responsável por selecionar os documentos apresentados.')

chapter('5 Governança e boas práticas')
h('5.1 Nomenclatura documental')
p('A nomenclatura padronizada permite identificar informações do documento de forma consistente. O validador interno contempla uma estrutura com campos como concessionária, rodovia e UF, localização, tipo de obra, tipo de projeto, classe, disciplina, sequência e revisão.')
p('A regra aplicável deve ser confirmada no padrão documental vigente para cada projeto. O validador pode identificar inconsistências e sugerir ajustes; a correção deve preservar a identificação do documento e sua coerência com os atributos.')
h('5.2 Organização documental')
p('Mantenha os documentos nas estruturas previstas, evitando pastas paralelas e duplicações sem finalidade definida. Uma organização consistente facilita consultas e automações.')
p('Para uma nova entrega, siga o procedimento aplicável de revisão, versão e substituição, preservando o histórico necessário.')
h('5.3 Permissões e responsabilidades')
p('As permissões determinam as ações disponíveis. Um usuário pode consultar um documento sem poder editá-lo ou alterar seu estado. O controle pode considerar usuários, grupos e regras dos objetos do ambiente.')
p('Há rotinas internas de gerenciamento de acessos e participantes. Solicite o acesso compatível com sua função e comunique impedimentos que afetem uma atividade autorizada.')
h('5.4 Rastreabilidade')
p('A rastreabilidade permite compreender o histórico documental. A trilha de auditoria pode registrar eventos das operações conforme as opções habilitadas no ambiente. Comentários claros, atributos corretos e uso adequado dos fluxos ajudam a tornar esse histórico compreensível.')
h('5.5 Cuidados na rotina')
bullets(['Confira projeto, pasta e documento antes de alterar.','Preencha atributos de acordo com a entrega.','Verifique revisão e situação antes de usar uma cópia local.','Devolva documentos ao concluir a edição.','Trate mensagens de validação antes de reenviar.','Preserve o histórico e registre informações úteis para a continuidade do trabalho.'])

chapter('6 Fluxos documentais')
h('6.1 Conceito e responsabilidades')
p('O fluxo representa as etapas pelas quais o documento passa e as responsabilidades associadas. Pode envolver recebimento, validação, análise, devolução, disponibilização e superação. Os estados e as transições disponíveis dependem da configuração do ambiente.')
h('6.2 Validação automática')
p('As rotinas internas verificam aspectos como número documental, revisão, versão e, conforme as condições previstas no código, presença de PDF associado ao arquivo editável. Ao identificar inconsistências, podem registrar uma mensagem e direcionar o documento para uma situação de não validação.')
p('Validação automática não comprova aprovação técnica. O documento pode atender às verificações programadas e ainda precisar da análise de engenharia ou de outra área responsável.')
h('6.3 Estados identificados nas rotinas internas')
table(['Estado','Interpretação no processamento'],[
('Nao Validado pelo Sistema','Foi identificada uma inconsistência nas verificações executadas.'),
('Validado pelo Sistema - A ser movido','Passou pelas verificações e aguarda o processamento seguinte.'),
('Validado pelo Sistema - Pasta destino nao localizada','A rotina não conseguiu localizar a pasta de destino.'),
('Validado pelo Sistema - Copiado p/ Disciplina','O processamento registrou a cópia para a disciplina.'),
('Superado','Foi tratado como anterior no contexto da rotina de substituição.')])
p('A utilização desses estados, suas transições e os responsáveis devem corresponder à configuração em operação.')
h('6.4 Distribuição e documentos superados')
p('A rotina de movimentação utiliza informações documentais para localizar o destino e executar o tratamento previsto. Também há lógica para encaminhar documentos anteriores à pasta Superados e aplicar o estado correspondente, conforme as condições do processamento.')
p('Atributos e revisão incorretos podem afetar o destino ou a relação com documentos anteriores. A consulta a superados deve considerar a finalidade histórica e a situação do documento vigente.')
h('6.5 Guia de Remessa de Documentos')
p('A GRD, usualmente denominada Guia de Remessa de Documentos, apoia a identificação dos documentos incluídos em uma remessa. Existe uma rotina interna específica para sua geração para unidade. Critérios de seleção, destinatários e condições de emissão seguem o procedimento correspondente.')
p('Ao consultar uma remessa, confira os documentos relacionados e as revisões entregues.')
h('6.6 Acompanhamento de pendências')
p('Pendências podem envolver cadastro, análise, arquivos em edição ou falhas de processamento. Localize o documento, confira o estado e a mensagem, identifique a correção e seu responsável, execute a ação permitida e confirme o resultado após a atualização ou o reprocessamento.')
p('Quando precisar de suporte, informe o projeto, a identificação do documento, o estado apresentado e a mensagem recebida.')
h('Exemplo de aplicação')
p('Uma entrega apresenta revisão incompatível com o valor esperado. O responsável consulta a mensagem e verifica se o problema está no arquivo, na identificação ou nos atributos. Após a correção e o encaminhamento previsto, o documento pode ser reprocessado. Se passar nas verificações, seguirá para a próxima etapa; a análise técnica permanece vinculada às regras e responsabilidades do fluxo.')

chapter('Referências e materiais de apoio')
h('Publicações da Bentley')
for title,url in [
('Working Remotely with ProjectWise','https://www.bentley.com/wp-content/uploads/document-Working-Remotely-With-ProjectWise-LTR-EN-0320.pdf'),
('Permissions in ProjectWise','https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0020684'),
('ProjectWise Glossary A','https://bentleysystems.service-now.com/community?id=kb_article_view&sysparm_article=KB0021937')]:
    p(title);p(url)
h('Rotinas internas relacionadas')
for t in [
'ExecucaoManualUsuario / 06 - Criar Buscas Salvas em Projetos.ps1',
'ExecucaoManualUsuario / 10 - Validar e Sugerir Ajuste de Nome de Arquivos.ps1',
'ExecucaoAutomaticaServidor / ValidaDocumentos.ps1',
'ExecucaoAutomaticaServidor / MovimentaDocumentosValidos_v3.ps1',
'ExecucaoAutomaticaServidor / GeraGRDsDocumentosComUnidade.ps1',
'Central de Scripts ProjectWise / README.md e documentação das rotinas de entrada de documentos, movimentação e geração de GRD']:
    d.add_paragraph(t,'List Bullet')
h('Preparação para o treinamento')
p('Disponibilize um projeto de treinamento com documentos de exemplo, contas com permissões compatíveis e buscas salvas. As demonstrações devem utilizar a interface adotada pela turma e seguir os procedimentos internos vigentes.')
d.core_properties.title='Guia introdutório ao ProjectWise'
d.core_properties.subject='Gestão e controle de documentos'
d.core_properties.author=''
dest=OUT/'Guia_Introdutorio_ProjectWise.docx'
d.save(dest)
print(dest)
