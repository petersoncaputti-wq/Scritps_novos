import fs from 'node:fs/promises';
import {FileBlob,PresentationFile} from 'file:///C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
process.on('unhandledRejection',e=>{console.error(e.message);process.exit(1)});
for(const [key,file] of [['riominas','ProjectWise_Web_Ecovias_Rio_Minas.pptx'],['ecorodovias','ProjectWise_Web_EcoRodovias.pptx']]){
const p=await PresentationFile.importPptx(await FileBlob.load('Documentacao/Apresentacoes/'+file));
const b=await p.export({slide:p.slides.items[30],format:'png',scale:1});
await fs.writeFile('Documentacao/ppt_build/'+key+'_reimport_31.png',new Uint8Array(await b.arrayBuffer()));console.log(key,'conferido');}
