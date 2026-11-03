import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {FileBlob,PresentationFile} from 'file:///C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';
process.on('unhandledRejection',e=>{console.error(e.message);process.exit(1)});
process.env.RUNTIME_NODE_MODULES='C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const root=process.cwd(), dir=path.join(root,'Documentacao/ppt_build');
const skill='C:/Users/Peterson/.codex/plugins/cache/openai-primary-runtime/presentations/26.1020.11127/skills/presentations';
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
for(const [key,file] of [['riominas','ProjectWise_Web_Ecovias_Rio_Minas'],['ecorodovias','ProjectWise_Web_EcoRodovias']]){
 const p=await PresentationFile.importPptx(await FileBlob.load('Documentacao/Apresentacoes/'+file+'.pptx'));
 const slide=p.slides.items[30];
 const title=slide.shapes.items.find(s=>s.name==='CaixaDeTexto 30');
 const sub=slide.shapes.items.find(s=>s.name==='TextBox 13');
 title.text='Dúvidas';title.text.style={typeface:'Verdana',fontSize:42.67,bold:true,color:'#FFFFFF',autoFit:'none',insets:0};title.bringToFront();
 sub.text='ProjectWise Web';sub.text.style={typeface:'Verdana',fontSize:24,color:'#72D516',autoFit:'none',insets:0};sub.bringToFront();
 const draft=path.join(dir,key+'_entrega.pptx');await(await PresentationFile.exportPptx(p)).save(draft);
 const img=await p.export({slide,format:'png',scale:1});await fs.writeFile(path.join(dir,key+'_entrega_31.png'),new Uint8Array(await img.arrayBuffer()));
 const out=path.join(root,'Documentacao/Apresentacoes/'+file+'_Treinamento.pptx');
 await finalizePresentation({workspaceDir:root,candidatePath:draft,finalPath:out,
 pythonExecutable:'C:/Users/Peterson/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
 integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit',...[2,12,21,22,27].flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:[2,12,21,22,27],
 verifyArtifactToolImport:true,receiptPath:path.join(dir,key+'_entrega_receipt.json')});console.log(out);
}
