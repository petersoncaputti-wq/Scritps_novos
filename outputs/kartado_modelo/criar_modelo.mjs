import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';
const wb = Workbook.create();
const s = wb.worksheets.add('Usuarios');
s.showGridLines = false;
s.getRange('A1:D1').values = [['nome','sobrenome','email','login']];
s.getRange('A1:D101').format.font.name = 'Arial';
s.getRange('A1:D101').format.font.size = 11;
s.getRange('A1:D1').format.fill = '#164374';
s.getRange('A1:D1').format.font.color = '#FFFFFF';
s.getRange('A1:D1').format.font.bold = true;
s.getRange('A1:D1').format.rowHeight = 28;
s.getRange('A1:D1').format.horizontalAlignment = 'center';
s.getRange('A2:D101').format.fill = '#FFF7DE';
s.getRange('A2:D101').setNumberFormat('@');
s.getRange('A2:D101').format.rowHeight = 22;
s.getRange('A1:A101').format.columnWidth = 24;
s.getRange('B1:B101').format.columnWidth = 34;
s.getRange('C1:C101').format.columnWidth = 42;
s.getRange('D1:D101').format.columnWidth = 30;
s.freezePanes.freezeRows(1);
s.getRange('F1').values = [['Como preencher']];
s.getRange('F1').format.font.bold = true;
s.getRange('F1:F14').format.columnWidth = 93;
s.getRange('F2:F14').values = [
 ['Preencha a partir da linha 2, um usuário por linha.'],
 ['As quatro colunas são obrigatórias. Mantenha os cabeçalhos da linha 1.'],
 ['nome: primeiro nome, até 30 caracteres.'],
 ['sobrenome: demais nomes, até 150 caracteres.'],
 ['email: endereço válido, até 254 caracteres.'],
 ['login: até 150 caracteres, com letras, números ou @ . + - _'],
 ['Use logins únicos. Preencha com texto, sem fórmulas.'],
 ['Unidade e perfil serão selecionados no script para todo o lote.'],
 ['As células amarelas estão reservadas para preenchimento. Pode adicionar linhas.'],
 ['Salve em .xlsx e mantenha esta aba Usuarios ativa ao executar o script.'],
 ['Não inclua senhas. As linhas vazias serão ignoradas.'],
 ['Exemplo ilustrativo (fora da área importada):'],
 ['Ana | Silva | ana.silva@example.com | ana.silva'],
];
s.getRange('F1:F14').format.font.name = 'Arial';
s.getRange('F1:F14').format.font.size = 11;
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'Usuarios!A1:D4',include:'values,formulas',tableMaxRows:4,tableMaxCols:4})).ndjson);
const preview = await wb.render({sheetName:'Usuarios',range:'A1:F14',scale:1,format:'png'});
await fs.writeFile(new URL('./preview.png', import.meta.url), new Uint8Array(await preview.arrayBuffer()));
await (await SpreadsheetFile.exportXlsx(wb)).save(fileURLToPath(new URL('./Modelo_Usuarios_Kartado.xlsx',import.meta.url)));
