import fs from 'node:fs/promises';
import path from 'node:path';
import {FileBlob, PresentationFile} from 'file:///C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
const dir=path.resolve('Documentacao/ppt_build');await fs.mkdir(dir,{recursive:true});
for(const [key,file] of [['riominas','Modelo PPT_Ecovias Rio Minas.pptx'],['ecorodovias','Modelo PPT_EcoRodovias.pptx']]){
 const p=await PresentationFile.importPptx(await FileBlob.load('C:/Users/Peterson/Downloads/'+file));
 await fs.writeFile(path.join(dir,key+'_inspect.txt'),(await p.inspect({kind:'slide,textbox,shape,image,layout',maxChars:150000})).ndjson);
 for(let i=0;i<p.slides.items.length;i++){
  const b=await p.export({slide:p.slides.items[i],format:'png',scale:.7});
  await fs.writeFile(path.join(dir,`${key}_ref_${i+1}.png`),new Uint8Array(await b.arrayBuffer()));
 }
 console.log(key,p.slides.items.length);
}
