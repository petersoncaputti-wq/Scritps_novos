import fs from 'node:fs/promises';
import {createRequire} from 'node:module';
import {fileURLToPath,pathToFileURL} from 'node:url';
const require=createRequire('C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/package.json');
const {Workbook, SpreadsheetFile}=await import(pathToFileURL(require.resolve('@oai/artifact-tool')).href);
const inferred = 'Nome inferido do e-mail';
const rows = [
['Rogerio J. Oliveira','rogerio.j.oliveira@ecovias.com.br',inferred],
['Amanda Silva Barbosa',null,'E-mail não informado'],
['Monique Lima','monique.lima@ecovias.com.br',inferred],
['Rafael Santos De Freitas','rafael.freitas@ecovias.com.br','Nome informado na conversa'],
['Joao Otavio Naibo Turmena','joao.turmena@ecovias.com.br','Nome informado na conversa'],
['Levindo Junior','T_Levindo.Junior@ecovias.com.br',inferred],
['Paulo A. Silva','paulo.a.silva@ecoriominas.com.br',inferred],
['Rafael Vianna','T_Rafael.Vianna@ecorodovias.com.br',inferred],
['Juliana Guimaraes de Mello',null,'E-mail não informado'],
['Robson Juacaba','t_robson.juacaba@ecovias.com.br',inferred],
['Jessica Silva','t_jessica.silva@ecorodovias.com.br',inferred],
['Vanessa Siqueira Brito','vanessa.s.brito@ecovias.com.br','Nome informado na conversa'],
['Igor Santana','T_Igor.Santana@ecovias.com.br',inferred],
['William Vieira Emenegildo',null,'E-mail não informado'],
['Sergio Mesquita','t_sergio.mesquita@ecorodovias.com.br',inferred],
['Patrick Oliveira','t_patrick.oliveira@ecovias.com.br',inferred],
['Thairine Isabel','T_thairine.isabel@ecorodovias.com.br',inferred],
['Andreane N. Santos','andreane.n.santos@ecovias.com.br',inferred],
['Maria T. Lima','T_Maria.T.Lima@ecovias.com.br',inferred],
].sort((a,b)=>a[0].localeCompare(b[0],'pt-BR'));
if(rows.length!==19 || new Set(rows.filter(r=>r[1]).map(r=>r[1].toLowerCase())).size!==16) throw Error('Contagem incorreta');
const wb=Workbook.create();
const s=wb.worksheets.add('Contatos');
s.showGridLines=false;
s.getRange('A2').values=[['Contatos — ProjectWise']];
s.getRange('A3').values=[['Fonte: conversa fornecida pelo usuário.']];
s.getRange('A5:C5').values=[['Nome','E-mail','Observação']];
s.getRange('A6:C24').values=rows;
s.getRange('A2:C24').format.font={name:'Arial',size:11,color:'#243247'};
s.getRange('A2').format.font={name:'Arial',size:15,bold:true};
s.getRange('A3').format.font={name:'Arial',size:10,italic:true,color:'#64748B'};
s.getRange('A5:C5').format={fill:'#24476A',font:{name:'Arial',size:11,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',rowHeight:27};
s.getRange('A6:C24').format.rowHeight=24;
s.getRange('A1:A24').format.columnWidth=38;
s.getRange('B1:B24').format.columnWidth=49;
s.getRange('C1:C24').format.columnWidth=37;
for(let i=0;i<rows.length;i++) if(i%2===1)s.getRange(`A${i+6}:C${i+6}`).format.fill='#EFF4F8';
s.tables.add('A5:C24',true,'ContatosPW');
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'Contatos!A5:C24',include:'values',tableMaxRows:20,tableMaxCols:3,maxChars:4500})).ndjson);
const dir=new URL('.',import.meta.url);
const preview=await wb.render({sheetName:'Contatos',range:'A1:C24',scale:1.5,format:'png'});
await fs.writeFile(new URL('preview.png',dir),new Uint8Array(await preview.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(wb)).save(fileURLToPath(new URL('contatos_projectwise.xlsx',dir)));
