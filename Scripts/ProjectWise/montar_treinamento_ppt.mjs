import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {createHash} from 'node:crypto';
import {FileBlob,PresentationFile} from 'file:///C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
process.on('uncaughtException',e=>{console.error(e.message);console.error(String(e.stack).split('\n').slice(0,5).join('\n'));process.exit(1)});
process.on('unhandledRejection',e=>{console.error(e?.message??e);process.exit(1)});
const workspaceDir=process.cwd();
process.env.RUNTIME_NODE_MODULES='C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const SKILL_DIR='C:/Users/Peterson/.codex/plugins/cache/openai-primary-runtime/presentations/26.1020.11127/skills/presentations';
const BUILD=path.join(workspaceDir,'Documentacao/ppt_build');
const OUT=path.join(workspaceDir,'Documentacao/Apresentacoes');
await fs.mkdir(OUT,{recursive:true});
const {finalizePresentation}=await import(pathToFileURL(path.join(SKILL_DIR,'container_tools/artifact_tool_utils.mjs')).href);
const slides=[
 {type:'cover',title:'ProjectWise\nWeb',sub:'Gestão de documentos'},
 {title:'Roteiro do treinamento',table:[['Etapa','Conteúdo','Tempo'],['1','Apresentação','15 min'],['2','Acesso e navegação','20 min'],['3','Gestão de documentos','40 min'],['4','Pesquisa e localização','20 min'],['5','Governança e boas práticas','25 min'],['6','Fluxos documentais','20 min']], widths:[100,800,200], note:'Duração de referência: 140 minutos, equivalentes a 2h20, sem intervalo. Os tempos incluem explicações e demonstrações.'},
 {title:'O que é o ProjectWise',paras:['O ProjectWise gerencia documentos e informações de projetos de infraestrutura.','Cada documento reúne um arquivo e informações como descrição, disciplina, revisão e situação.','O ambiente compartilhado permite localizar documentos e acompanhar seu histórico.'],note:'Exemplos de documentos: desenhos, relatórios, memoriais e planilhas. O conteúdo deste treinamento considera exclusivamente o acesso pela versão Web.'},
 {title:'Benefícios na rotina',paras:['Documentos organizados por projeto, pasta e disciplina.','Acesso de acordo com as responsabilidades de cada participante.','Consulta às revisões, às versões e à situação documental.','Menos tempo para localizar arquivos e menor risco de utilizar uma cópia desatualizada.']},
 {type:'section',title:'Acesso e navegação',paras:['Etapa 2 do treinamento','Acesso exclusivamente pelo ProjectWise Web.']},
 {title:'Acesso pelo navegador',paras:['O usuário acessa o endereço disponibilizado pela organização e conclui a autenticação.','Depois, seleciona o projeto ou a conexão disponíveis para seu perfil.','A visualização do conteúdo depende das permissões concedidas.','Se um projeto não aparecer, confira o ambiente e solicite a verificação do acesso.'],note:'O ProjectWise Explorer é a aplicação para computador. Esta apresentação utiliza exclusivamente a versão Web. Não se pressupõe que todos os projetos estejam visíveis após a autenticação.'},
 {title:'Projetos, pastas e documentos',paras:['A estrutura de pastas organiza os documentos conforme o padrão do projeto.','Na listagem, confira o nome, a descrição, a revisão e o estado disponíveis.','Antes de utilizar um resultado, confirme se ele pertence ao projeto e à pasta corretos.'],note:'Demonstração: acessar um projeto de treinamento, navegar até uma disciplina e selecionar um documento. As colunas disponíveis variam conforme a configuração do ambiente.'},
 {title:'Propriedades e atributos',paras:['Os atributos registram informações utilizadas na identificação e na pesquisa.','Disciplina, revisão e número documental são exemplos de campos do cadastro.','O nome do arquivo e os atributos precisam ser coerentes.','Divergências podem prejudicar a localização e o processamento do documento.'],note:'Atividade: localizar nas propriedades de um documento a revisão e a disciplina. Comparar essas informações com o nome do arquivo.'},
 {type:'section',title:'Gestão de documentos',paras:['Etapa 3 do treinamento','Recebimento, consulta e identificação das entregas.']},
 {title:'Inclusão da documentação',paras:['A inclusão ocorre por meio de GRDs aceitas pelos assistentes da engenharia.','A GRD identifica os documentos encaminhados na remessa.','A inclusão manual cabe ao pessoal autorizado, em situações específicas e conforme a demanda.','Se uma entrega não aparecer, verifique o recebimento com o responsável.'],note:'A inclusão manual não é uma atividade geral dos participantes deste treinamento. Ela deve seguir a demanda e a autorização aplicáveis, respeitando pasta, nomenclatura e atributos.'},
 {title:'Download e consulta',paras:['O download disponibiliza uma cópia do documento no computador.','Essa cópia pode ficar desatualizada após uma nova entrega ou alteração no ambiente.','Antes de usar o arquivo, confira a revisão e a situação no ProjectWise Web.','Solicitações de alteração ou substituição seguem o procedimento interno.']},
 {title:'Programas para abrir os arquivos',table:[['Formato','Exemplos de programas compatíveis'],['PDF','Adobe Acrobat Reader ou outro leitor de PDF.'],['DWG','Autodesk DWG TrueView, AutoCAD ou outro leitor compatível.'],['DGN','Bentley View, MicroStation ou outro leitor compatível.'],['DOCX','Microsoft Word ou aplicação compatível.'],['XLSX','Microsoft Excel ou aplicação compatível.']],widths:[160,940],note:'O navegador pode oferecer pré-visualização conforme o formato e os recursos do ambiente. Para abrir uma cópia baixada no computador, é necessário instalar um programa compatível com o formato e com a versão do arquivo. Referências externas, fontes e outros arquivos de apoio também podem ser necessários.'},
 {title:'Versão e revisão',paras:['A versão está ligada ao controle das ocorrências de um documento no sistema.','A revisão identifica uma evolução ou emissão conforme o padrão documental.','Na consulta, compare o nome do arquivo com os atributos e a entrega esperada.','Uma divergência deve ser encaminhada ao responsável pela documentação.']},
 {type:'section',title:'Pesquisa e localização',paras:['Etapa 4 do treinamento','Consultas realizadas no ProjectWise Web.']},
 {title:'Pesquisa pelo nome',paras:['Utilize o recurso de pesquisa disponível no ProjectWise Web.','Informe o código completo ou uma parte conhecida do nome.','Confira o escopo da busca, pois ele pode limitar os resultados.','Valide projeto, revisão e situação antes de utilizar o documento.']},
 {title:'Filtros e atributos',paras:['Use os filtros e os campos efetivamente disponíveis na conexão Web.','Quando oferecidos, combine critérios como disciplina, revisão, número e estado.','Se não houver resultados, revise o escopo, a grafia e os filtros.','Verifique também se o usuário possui permissão de acesso ao conteúdo.'],note:'Não há pressuposto de que todos os recursos do Explorer estejam disponíveis na interface Web. A demonstração deve utilizar as opções apresentadas no ambiente da turma.'},
 {title:'Buscas salvas',paras:['Quando disponíveis na conexão Web, as buscas salvas reutilizam critérios de consulta.','Os resultados mudam conforme os documentos são incluídos ou atualizados.','Essas consultas podem apoiar o acompanhamento de pendências.','Se a busca desejada não aparecer, utilize a pesquisa disponível ou solicite orientação.']},
 {title:'Atividade de pesquisa',paras:['Localize um documento pelo nome no ProjectWise Web.','Confirme a pasta, a revisão e o estado apresentados.','Refine a consulta utilizando os filtros disponíveis.','Se houver uma busca salva acessível, identifique quais documentos ela retorna.'],note:'Utilizar um projeto de treinamento com conteúdo previamente selecionado. A atividade depende dos recursos e das permissões disponíveis para a turma.'},
 {type:'section',title:'Governança e boas práticas',paras:['Etapa 5 do treinamento','Nomenclatura, organização e responsabilidades.']},
 {title:'Exemplo real de nomenclatura',paras:['ERM-465RJ-010+550-BSO-FUN-DE-F1-001-R00B.dwg','O nome reúne nove campos de identificação separados por hífens.','A extensão indica o formato do arquivo.','A grafia da revisão R00B permanece igual à do documento real.'],note:'Exemplo real fornecido pelo usuário. Os próximos slides identificam cada campo. Os significados específicos de cada código seguem a tabela de nomenclatura vigente no projeto.'},
 {title:'Nomenclatura e identificação',table:[['Parte','Campo','Leitura'],['ERM','Concessionária','Código da concessionária responsável.'],['465RJ','Rodovia e UF','Rodovia 465, no Rio de Janeiro.'],['010+550','Localização','Km 10 mais 550 metros.'],['BSO','Tipo de obra','Código da obra ou intervenção.'],['FUN','Tipo de projeto','Código da categoria ou etapa do projeto.']],widths:[160,300,640],note:'Os significados específicos dos códigos ERM, BSO e FUN devem seguir a tabela de nomenclatura aplicável. A localização no formato 010-020 representaria um trecho entre os quilômetros 10 e 20.'},
 {title:'Nomenclatura e documento',table:[['Parte','Campo','Leitura'],['DE','Classe documental','Código da classe ou do tipo de documento.'],['F1','Disciplina','Código da disciplina técnica.'],['001','Sequência','Número sequencial com três dígitos.'],['R00B','Revisão','Identificação da revisão do documento.'],['.dwg','Extensão','Arquivo de desenho em formato DWG.']],widths:[160,300,640],note:'A extensão não integra os nove campos. Preservar R00B exatamente como consta no exemplo real, sem converter a letra para minúscula. Comparar nome, atributos e revisão esperada para a entrega.'},
 {title:'Organização documental',paras:['Os documentos seguem as pastas e as disciplinas previstas para o projeto.','Pastas paralelas e duplicações sem finalidade definida dificultam a consulta.','Novas entregas precisam seguir o procedimento de revisão, versão e substituição.','O histórico deve permanecer disponível conforme o controle documental adotado.']},
 {title:'Permissões e rastreabilidade',paras:['Consultar um documento não significa ter permissão para alterá-lo.','Cada participante deve possuir acesso compatível com sua função.','A trilha de auditoria registra eventos conforme a configuração do ambiente.','Atributos corretos e registros claros ajudam a compreender o histórico documental.']},
 {type:'section',title:'Fluxos documentais',paras:['Etapa 6 do treinamento','Estados, validação e acompanhamento de pendências.']},
 {title:'Validação automática',paras:['As verificações podem considerar número documental, revisão e versão.','Conforme a regra aplicável, também verificam a presença de PDF associado ao arquivo editável.','Quando há inconsistência, o processamento pode registrar uma mensagem.','A aprovação técnica continua dependendo da análise da equipe responsável.'],note:'Passar nas verificações automáticas não comprova a aprovação técnica do conteúdo. O usuário deve interpretar a mensagem e encaminhar a correção conforme a responsabilidade definida.'},
 {title:'Estados documentais',table:[['Estado identificado','Interpretação'],['Nao Validado pelo Sistema','O processamento identificou uma inconsistência.'],['Validado pelo Sistema - A ser movido','Aguarda o processamento seguinte.'],['Validado pelo Sistema - Pasta destino nao localizada','A rotina não localizou a pasta de destino.'],['Validado pelo Sistema - Copiado p/ Disciplina','O processamento registrou a cópia para a disciplina.'],['Superado','Documento tratado como anterior no fluxo.']],widths:[550,550],note:'Nomes identificados nas rotinas internas e reproduzidos como registrados. As transições e responsabilidades devem corresponder à configuração em operação.'},
 {title:'Distribuição e GRD',paras:['As informações do documento orientam sua distribuição para a disciplina.','O fluxo pode encaminhar documentos anteriores para a pasta Superados.','A GRD identifica os documentos e as revisões de uma remessa.','No recebimento, os assistentes da engenharia aceitam as GRDs conforme o procedimento.'],note:'A consulta a documentos superados deve considerar a finalidade histórica e a situação do documento vigente. A inclusão manual permanece restrita a pessoal autorizado e situações específicas conforme a demanda.'},
 {title:'Acompanhamento de pendências',paras:['Localize o documento e confira o estado e a mensagem disponível.','Identifique a correção necessária e o responsável pela ação.','Acompanhe o resultado após a atualização ou o reprocessamento.','Ao solicitar suporte, informe projeto, documento, estado e mensagem recebida.']},
 {title:'Caso para discussão',paras:['Uma entrega apresenta revisão diferente da esperada.','O responsável consulta a mensagem e compara arquivo, nome e atributos.','A correção segue o procedimento interno e a autorização aplicável.','Depois do reprocessamento, a equipe confere o resultado e a próxima etapa.'],note:'Perguntas para a turma: onde localizar o documento? Qual revisão aparece no nome? Quem deve avaliar a divergência? A aprovação técnica está confirmada apenas pela validação automática? Finalizar retomando as responsabilidades e as dúvidas dos participantes.'},
 {type:'end'}
];

for(const [key,input,output] of [
 ['riominas','Modelo PPT_Ecovias Rio Minas.pptx','ProjectWise_Web_Ecovias_Rio_Minas.pptx'],
 ['ecorodovias','Modelo PPT_EcoRodovias.pptx','ProjectWise_Web_EcoRodovias.pptx']]){
 const ref='C:/Users/Peterson/Downloads/'+input;
 const presentation=await PresentationFile.importPptx(await FileBlob.load(ref));
 const refs=[...presentation.slides.items];
 const built=[];
 const tableOwners=[];
 for(let i=0;i<slides.length;i++){
  const cfg=slides[i];
  const source=cfg.type==='cover'||cfg.type==='end'?refs[1]:cfg.type==='section'?refs[3]:refs[2];
  const slide=source.duplicate();built.push(slide);
  if(cfg.type==='cover'||cfg.type==='end'){
   const title=slide.shapes.items.find(s=>s.name==='CaixaDeTexto 30');
   const sub=slide.shapes.items.find(s=>s.name==='TextBox 13');
   title.text=cfg.type==='end'?'Dúvidas':cfg.title;sub.text=cfg.type==='end'?'ProjectWise Web':cfg.sub;
  }else if(cfg.type!=='end'){
   const title=slide.shapes.items.find(s=>s.name==='CaixaDeTexto 14');
   const body=slide.shapes.items.find(s=>s.name==='TextBox 13');
   title.text=cfg.title;
   title.position={left:42.91,top:87.99,width:1080,height:65};
   title.text.style={typeface:'Verdana',fontSize:42.67,bold:true,color:cfg.type==='section'?'#FFFFFF':'#1DA613',autoFit:'none',wrap:true,insets:0};
   body.position={left:59.66,top:190,width:1140,height:420};
   body.text=cfg.table?'':cfg.paras.join('\n\n');
   body.text.style={typeface:'Verdana',fontSize:26,color:cfg.type==='section'?'#FFFFFF':'#000000',autoFit:'none',wrap:true,verticalAlignment:'top',insets:0};
   if(cfg.table){
    tableOwners.push(i+1);
    const rows=cfg.table.length,columns=cfg.table[0].length;
    const table=slide.tables.add({rows,columns,left:59.66,top:185,width:1100,height:rows*65,columnWidths:cfg.widths,values:cfg.table});
    table.borders.assign({style:'solid',fill:'#D9E6D9',width:1});
    table.styleOptions={headerRow:true};
    for(let r=0;r<rows;r++)for(let c=0;c<columns;c++){
     const cell=table.getCell(r,c);
     cell.fill=r===0?'#056B28':r%2?'#FFFFFF':'#F0F7F0';
     cell.text.style={typeface:'Verdana',fontSize:22,color:r===0?'#FFFFFF':'#172317',bold:r===0,wrap:true,autoFit:'none',verticalAlignment:'middle',insets:{left:12,right:12,top:6,bottom:6}};
    }
   }
  }
  if(cfg.note)slide.speakerNotes.text=cfg.note;
 }
 const snapshot=(await presentation.inspect({kind:'slide',limit:200,maxChars:100000})).ndjson.trim().split('\n').map(line=>JSON.parse(line));
 const anchors=new Map(snapshot.filter(r=>r.kind==='slide').map(r=>[presentation.slides.items[r.slideIndex].id,r.id]));
 const ids=built.map(s=>anchors.get(s.id));
 presentation.slides.keep(ids);presentation.slides.reorder(ids);
 console.log(key,'built',presentation.slides.items.length);
 const draft=path.join(BUILD,key+'_candidate.pptx');
 await(await PresentationFile.exportPptx(presentation)).save(draft);
 for(let i=0;i<presentation.slides.items.length;i++){
  const png=await presentation.export({slide:presentation.slides.items[i],format:'png',scale:1});
  await fs.writeFile(path.join(BUILD,`${key}_final_${i+1}.png`),new Uint8Array(await png.arrayBuffer()));
 }
 await fs.writeFile(path.join(BUILD,key+'_final_inspect.txt'),(await presentation.inspect({kind:'slide,textbox,table,image',maxChars:200000})).ndjson);
 const final=path.join(OUT,output);
 const result=await finalizePresentation({workspaceDir,candidatePath:draft,finalPath:final,
 pythonExecutable:'C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
 integrityValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_package_integrity.py'),
 layoutValidatorPath:path.join(SKILL_DIR,'container_tools/inspect_presentation_layout_geometry.py'),
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit',...tableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],
 requiredNativeTableOwnerSlides:tableOwners,
 fontPolicy:{basis:'reference',families:['Verdana'],referencePath:ref,referenceSha256:createHash('sha256').update(await fs.readFile(ref)).digest('hex')},
 verifyArtifactToolImport:true,receiptPath:path.join(BUILD,output+'.validation.json')});
 console.log(key,'FINAL',final,JSON.stringify(result).slice(0,600));
}
