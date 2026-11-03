from pathlib import Path

root = Path(__file__).parent
src = (root / 'criar_guia_pw.py').read_text(encoding='utf-8')

def replace(old, new):
    global src
    assert old in src, old[:120]
    src = src.replace(old, new)

replace("Explica também como as automações internas apoiam a organização, a validação e a distribuição de documentos.", "O conteúdo é voltado à consulta e ao acompanhamento de documentos pela versão Web, considerando os procedimentos de recebimento e organização adotados pela empresa.")
replace("Ao concluir a leitura, o participante deverá conseguir localizar documentos, compreender o ciclo de edição, reconhecer as regras de organização e identificar a próxima ação em uma pendência documental.", "Ao concluir a leitura, o participante deverá conseguir acessar o ProjectWise Web, localizar e consultar documentos, compreender as formas de inclusão e reconhecer as regras de nomenclatura e o tratamento de pendências.")
replace("As telas, os comandos e as operações disponíveis variam conforme a versão, o acesso pelo ProjectWise Explorer ou Web e as permissões do usuário. Os procedimentos de cada projeto devem orientar a execução das atividades.", "Este guia considera exclusivamente o acesso pelo ProjectWise Web. As telas, as opções de visualização e os comandos disponíveis dependem da configuração do ambiente e das permissões do usuário.")
start = src.index("h('1.3 Papel das automações internas')")
end = src.index("h('Exemplo de aplicação')", start)
src = src[:start] + src[end:]
replace("Ao receber uma nova entrega, a equipe registra os documentos, informa seus atributos e acompanha a situação de processamento. As verificações automáticas apoiam o controle documental; a análise técnica continua sob responsabilidade da equipe designada.", "O recebimento de documentação ocorre por meio de GRDs aceitas pelos assistentes da engenharia. Após o processamento previsto, os usuários com acesso podem localizar os documentos e acompanhar sua situação pelo ProjectWise Web. A análise técnica permanece sob responsabilidade da equipe designada.")
replace("O acesso exige uma conta habilitada e permissões compatíveis com a atividade. A utilização pode ocorrer pelo ProjectWise Explorer, instalado no computador, ou pela interface Web disponibilizada pela organização.", "O acesso tratado neste guia é realizado exclusivamente pelo ProjectWise Web, utilizando o navegador e o endereço disponibilizado pela organização. É necessário possuir uma conta habilitada e permissões compatíveis com os projetos e as atividades a consultar.")
replace("No Explorer, pode ser necessário selecionar a fonte de dados, chamada datasource, antes da autenticação. Na interface Web, o acesso ocorre pelo endereço informado pela empresa. Após entrar, o usuário visualiza o conteúdo permitido para seu perfil.", "O ProjectWise também possui uma aplicação para computador, denominada ProjectWise Explorer. Entretanto, os procedimentos deste guia utilizam a versão Web. Acesse o endereço informado, conclua a autenticação solicitada e selecione o projeto ou a conexão disponibilizados para seu perfil.")
replace("A interface normalmente oferece uma área de navegação por projetos e pastas e uma área de listagem dos documentos. As colunas disponíveis podem apresentar nome, descrição, revisão, versão, estado e outras propriedades.", "No ProjectWise Web, a navegação permite acessar os projetos e as estruturas de pastas disponibilizados ao usuário. A listagem de documentos pode apresentar nome, descrição, revisão, versão, estado e outras propriedades, conforme a configuração do ambiente.")
replace("'Observe se existe uma reserva para edição.'", "'Confirme se o arquivo corresponde à entrega que deseja consultar.'")
start = src.index("p('A inclusão consiste")
end = src.index("h('3.2 Download", start)
src = src[:start] + """p('A inclusão da documentação ocorre por meio de GRDs aceitas pelos assistentes da engenharia. A GRD identifica os documentos encaminhados na remessa, e seu aceite integra o procedimento de recebimento adotado pela empresa.')
p('Também existe inclusão manual, realizada exclusivamente por pessoal autorizado, em situações específicas e conforme a demanda. Esse procedimento deve respeitar a definição do projeto, a pasta de destino, a nomenclatura e os atributos aplicáveis.')
p('Para os usuários deste guia, a atividade principal é consultar os documentos disponibilizados e acompanhar sua situação no ProjectWise Web. Caso um documento esperado não seja localizado, o usuário deve verificar o recebimento com o assistente da engenharia ou com o responsável pela inclusão.')
""" + src[end:]
replace("Antes de utilizar um arquivo baixado, confira sua revisão e situação no ProjectWise. Obter uma cópia para consulta e reservar um documento para edição são operações com finalidades diferentes. O download, por si só, não deve ser interpretado como reserva ou autorização para edição.", "Antes de utilizar um arquivo baixado pelo ProjectWise Web, confira sua revisão e situação. Uma cópia obtida para consulta não altera o documento do ambiente. Caso seja necessária uma alteração ou substituição, encaminhe a demanda ao responsável, conforme o procedimento interno.")
start = src.index("p('A abertura pode ocorrer")
end = src.index("h('3.7 Versão e revisão')", start)
src = src[:start] + """p('O ProjectWise Web pode permitir a visualização no navegador quando o formato e os recursos do ambiente forem compatíveis. Quando o arquivo for baixado e aberto no computador, é necessário ter instalado um programa adequado à leitura daquele formato e compatível com a versão do arquivo.')
p('O acesso ao documento no ProjectWise Web não substitui o programa utilizado para ler o arquivo no computador. Exemplos de aplicações são:')
table(['Formato', 'Exemplos de programas para leitura'], [
('.pdf', 'Adobe Acrobat Reader ou outro leitor de PDF compatível.'),
('.dwg', 'Autodesk DWG TrueView, AutoCAD ou outro visualizador compatível.'),
('.dgn', 'Bentley View, MicroStation ou outro leitor compatível com DGN.'),
('.docx', 'Microsoft Word ou aplicação compatível.'),
('.xlsx', 'Microsoft Excel ou aplicação compatível.')])
p('A abertura correta também pode depender de referências externas, fontes e outros arquivos associados. Se o documento não abrir ou apresentar informações incompletas, confira o programa, a compatibilidade e os arquivos de apoio necessários antes de solicitar suporte.')
""" + src[end:]
replace("h('3.7 Versão e revisão')", "h('3.4 Versão e revisão')")
replace("Um check-in não significa necessariamente a emissão de nova revisão. Alterar a revisão no nome do arquivo também não garante que o controle de versões esteja correto. Os scripts internos comparam informações de revisão e versão com valores calculados, reforçando a necessidade de coerência entre arquivo e atributos.", "Ao consultar um documento no ProjectWise Web, confira a revisão indicada no nome e nos atributos, bem como as informações de versão disponíveis. Alterar somente a indicação de revisão no nome do arquivo não garante a atualização correta do registro. Quando identificar divergências, comunique o responsável pela documentação.")
replace("A pesquisa por nome é útil quando se conhece o código completo ou parte dele. Critérios mais precisos facilitam a identificação; buscas amplas podem retornar documentos de projetos, disciplinas ou revisões diferentes.", "No ProjectWise Web, utilize o recurso de pesquisa disponibilizado no ambiente para localizar um documento pelo código completo ou por parte do nome. Critérios mais precisos facilitam a identificação; buscas amplas podem retornar documentos de projetos, disciplinas ou revisões diferentes.")
replace("Quando o nome não é conhecido, os atributos permitem pesquisar características cadastradas. Conforme a configuração, podem ser usados disciplina, revisão, número documental, estado e outros campos.", "Quando o nome não é conhecido, utilize os filtros e campos de pesquisa disponíveis no ProjectWise Web. Conforme a configuração da conexão, podem ser oferecidos critérios por disciplina, revisão, número documental, estado ou outros atributos. Utilize apenas as opções efetivamente apresentadas no ambiente.")
replace("As buscas salvas armazenam critérios para consultas recorrentes. Elas apoiam a localização de documentos e o acompanhamento de pendências.", "Quando disponibilizadas na conexão do ProjectWise Web, as buscas salvas permitem reutilizar critérios de consulta definidos no ambiente. Elas apoiam a localização de documentos e o acompanhamento de pendências.")
replace("Uma busca salva é uma consulta reutilizável: seus resultados podem mudar conforme os documentos são incluídos, atualizados ou movimentados no fluxo. A rotina interna de criação de buscas contempla consultas por estado documental, incluindo situações relacionadas à análise e à validação.", "Os resultados podem mudar conforme os documentos são incluídos, atualizados ou movimentados no fluxo. Se a busca desejada não estiver disponível na interface Web, utilize os recursos de pesquisa apresentados ou solicite orientação ao responsável pelo ambiente.")
replace("Localize um documento pelo nome e depois por atributos. Consulte uma busca salva e identifique o critério responsável por selecionar os documentos apresentados.", "No ProjectWise Web, localize um documento pelo nome, confirme a pasta, a revisão e o estado e utilize os filtros disponíveis para refinar a consulta. Se houver buscas salvas acessíveis, consulte uma delas e observe os documentos retornados.")
start=src.index("p('A nomenclatura padronizada")
end=src.index("h('5.2 Organização documental')",start)
src=src[:start]+"""p('A nomenclatura padronizada reúne campos que identificam a origem, a localização, o tipo e a revisão do documento. O exemplo real abaixo permite identificar cada parte do nome de um arquivo utilizado na documentação do projeto.')
p('ERM-465RJ-010+550-BSO-FUN-DE-F1-001-R00B.dwg')
table(['Parte', 'Campo', 'Significado no exemplo'], [
('ERM', 'Concessionária', 'Código que identifica a concessionária responsável pelo documento.'),
('465RJ', 'Rodovia e UF', 'Rodovia 465 no estado do Rio de Janeiro. O campo combina três dígitos e a sigla da UF.'),
('010+550', 'Localização', 'Ponto no km 10 mais 550 metros. A forma 010-020 indicaria um trecho entre os quilômetros 10 e 20.'),
('BSO', 'Tipo de obra', 'Código do tipo de obra ou intervenção, conforme a tabela de nomenclatura aplicável.'),
('FUN', 'Tipo de projeto', 'Código da categoria ou etapa do projeto, conforme o padrão adotado.'),
('DE', 'Classe documental', 'Código que identifica a classe ou o tipo de documento.'),
('F1', 'Disciplina', 'Código da disciplina técnica à qual o documento está vinculado.'),
('001', 'Sequência', 'Número sequencial de três dígitos utilizado para diferenciar documentos.'),
('R00B', 'Revisão', 'Identificação da revisão do documento, composta por R, os dígitos 00 e a letra B, preservando a grafia do arquivo real.'),
('.dwg', 'Extensão do arquivo', 'Indica um arquivo de desenho em formato DWG, que exige um programa compatível para abertura no computador. A extensão não integra os nove campos de identificação.')])
p('Os nove campos são separados por hífens. A localização pode conter um hífen adicional quando representar um trecho. Os significados específicos dos códigos de concessionária, tipo de obra, tipo de projeto, classe e disciplina devem seguir a tabela de nomenclatura vigente no projeto.')
p('Ao consultar um documento, compare o nome com os atributos registrados e com a revisão esperada para a entrega. Divergências devem ser encaminhadas ao responsável pela documentação para avaliação e correção.')
"""+src[end:]
replace("'Devolva documentos ao concluir a edição.'", "'Encaminhe solicitações de inclusão ou substituição ao pessoal autorizado.'")
replace("Pendências podem envolver cadastro, análise, arquivos em edição ou falhas de processamento.", "Pendências podem envolver cadastro, análise, recebimento ou falhas de processamento.")
replace("Existe uma rotina interna específica para sua geração para unidade. Critérios de seleção, destinatários e condições de emissão seguem o procedimento correspondente.", "No recebimento de documentação, as GRDs são aceitas pelos assistentes da engenharia, conforme o procedimento adotado pela empresa. A inclusão manual é restrita a pessoal autorizado e a situações específicas, conforme a demanda.")
start = src.index("chapter('Referências e materiais de apoio')")
end = src.index("d.core_properties.title", start)
src=src[:start]+src[end:]
replace("dest=OUT/'Guia_Introdutorio_ProjectWise.docx'", "dest=OUT/'Guia_ProjectWise_Web_Revisado.docx'")
exec(compile(src, str(root / 'criar_guia_pw.py'), 'exec'), {'__file__': str(root / 'criar_guia_pw.py'), '__name__':'__main__'})
