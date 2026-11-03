from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from copy import deepcopy
from hashlib import sha256
import json
from lxml import etree as E

ROOT=Path(__file__).parent
OUT=ROOT/'Documentacao'
QA=OUT/'verificacao_timbrados'
QA.mkdir(exist_ok=True)
SOURCE=OUT/'Guia_ProjectWise_Web_Revisado.docx'
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS={'w':W}
def q(t): return '{'+W+'}'+t
def element(t, **attrs):
    el=E.Element(q(t))
    for k,v in attrs.items(): el.set(q(k),str(v))
    return el
def setchild(parent,tag,**attrs):
    old=parent.find(q(tag))
    if old is None: old=E.SubElement(parent,q(tag))
    for k,v in attrs.items(): old.set(q(k),str(v))
    return old
def xml(el): return E.tostring(el,encoding='UTF-8',xml_declaration=True,standalone=True)
def digest(b): return sha256(b).hexdigest()

with ZipFile(SOURCE) as z:
    source=E.fromstring(z.read('word/document.xml'))
sourcebody=source.find(q('body'))

templates=[
    (Path(r'C:\Users\Peterson\Downloads\Ecovias Rio Minas_Papel_Carta.docx'),'Guia_ProjectWise_Web_Ecovias_Rio_Minas.docx'),
    (Path(r'C:\Users\Peterson\Downloads\EcoRodovias_Papel_Carta_A4_20250227_Keenwork_AF.docx'),'Guia_ProjectWise_Web_EcoRodovias.docx')]
for template,name in templates:
    originalhash=digest(template.read_bytes())
    work=QA/template.stem
    work.mkdir(exist_ok=True)
    with ZipFile(template) as z:
        parts={i.filename:z.read(i.filename) for i in z.infolist()}
        infos=z.infolist()
    root=E.fromstring(parts['word/document.xml'])
    body=root.find(q('body'))
    section=deepcopy(body.find(q('sectPr')))
    geometry=E.tostring(section).decode()
    inventory={k:{'bytes':len(v),'sha256':digest(v),'role':'editable' if k in ['word/document.xml','word/styles.xml'] else 'preserve-only'} for k,v in parts.items()}
    (work/'package_inventory.json').write_text(json.dumps(inventory,ensure_ascii=False,indent=2),encoding='utf-8')
    (work/'artifact.md').write_text(
        f'# Contrato do papel timbrado\n\nReferência: {template}\nSHA256: {originalhash}\n'
        'Uma seção; página 11904 x 16836 twips; margens superior 1304, inferior 1440, esquerda e direita 1440 twips. '
        'Cabeçalho 584 e rodapé 834 twips. Sem troca de seção nem de cabeçalho por capítulo.\n\n'
        'Fonte do conteúdo de exemplo: Verdana; corpo 12 pt, título de exemplo 20 pt negrito. '
        'Hierarquia estendida para o guia: título 20 pt, capítulo 16 pt e subtópico 12 pt em negrito, todos pretos. '
        'Corpo 12 pt, entrelinha 1,12, espaço posterior 7 pt. Listas com marcador e recuo. '
        'Tabelas do guia ajustadas à largura útil de 9024 twips, mantendo bordas e cabeçalhos repetidos.\n\n'
        'Slot editável: word/document.xml / w:document/w:body, exceto w:sectPr. '
        'Substituir todos os parágrafos de Lorem ipsum pelo conteúdo integral do guia atualizado. '
        'Adicionar estilos próprios apenas para os novos papéis de texto; preservar os estilos existentes.\n\n'
        'Preservar literalmente as demais partes do pacote, inclusive cabeçalhos, rodapés, imagens PNG e EMF, '
        'relacionamentos, temas, controles, customXml e dados institucionais do modelo. '
        'Inventário completo em package_inventory.json. Sem campos novos ou novas dependências de mídia.\n\n'
        'Paginação visual não resolvida: renderizador LibreOffice indisponível neste ambiente. '
        'Validar estrutura, cobertura textual e preservação binária das partes institucionais.\n\n'
        f'Geometria preservada: {geometry}\n',encoding='utf-8')
    assert digest(template.read_bytes())==originalhash
    styles=E.fromstring(parts['word/styles.xml'])
    baseids={s.get(q('styleId')) for s in styles.findall(q('style'))}
    # Extend the template typography without altering any original style.
    roles={'Title':('PWGuiaTitulo',40,True), 'Subtitle':('PWGuiaSubtitulo',24,False),
           'Heading1':('PWGuiaCapitulo',32,True), 'Heading2':('PWGuiaTopico',24,True),
           'Normal':('PWGuiaCorpo',24,False),'ListBullet':('PWGuiaLista',24,False)}
    for role,(sid,size,bold) in roles.items():
        st=element('style',type='paragraph',styleId=sid)
        st.append(element('name',val=sid));st.append(element('basedOn',val='NormalWeb'))
        pr=E.SubElement(st,q('pPr'))
        setchild(pr,'spacing',before=0,after=140,line=269,lineRule='auto')
        if role in ['Title','Heading1','Heading2']:
            setchild(pr,'keepNext');setchild(pr,'spacing',before=200,after=120)
        if role in ['Heading1','Heading2']: setchild(pr,'outlineLvl',val=0 if role=='Heading1' else 1)
        if role=='ListBullet': setchild(pr,'ind',left=360,hanging=240)
        rp=E.SubElement(st,q('rPr'))
        setchild(rp,'rFonts',ascii='Verdana',hAnsi='Verdana',cs='Verdana')
        setchild(rp,'sz',val=size);setchild(rp,'szCs',val=size);setchild(rp,'color',val='000000');setchild(rp,'lang',val='pt-BR')
        if bold:setchild(rp,'b')
        styles.append(st)
    for child in list(body): body.remove(child)
    for child in sourcebody:
        if child.tag==q('sectPr'): continue
        new=deepcopy(child)
        for para in new.iter(q('p')):
            pp=para.find(q('pPr'))
            if pp is None: pp=E.Element(q('pPr'));para.insert(0,pp)
            ps=pp.find(q('pStyle'))
            role=ps.get(q('val')) if ps is not None else 'Normal'
            if role=='ListBullet':
                run=element('r');text=E.SubElement(run,q('t'));text.text='• '
                text.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
                para.insert(list(para).index(pp)+1,run)
            mapped=roles.get(role,roles['Normal'])[0]
            setchild(pp,'pStyle',val=mapped)
            for num in pp.findall(q('numPr')): pp.remove(num)
            for run in para.findall(q('r')):
                rp=run.find(q('rPr'))
                if rp is None: rp=E.Element(q('rPr'));run.insert(0,rp)
                setchild(rp,'rFonts',ascii='Verdana',hAnsi='Verdana',cs='Verdana')
                # Tables use 11 pt to accommodate their explanatory columns.
                if child.tag==q('tbl'):
                    setchild(rp,'sz',val=22);setchild(rp,'szCs',val=22)
                setchild(rp,'color',val='000000')
        if new.tag==q('tbl'):
            count=len(new.find(q('tr')).findall(q('tc')))
            if count==3:
                # Nomenclature fields need a wider explanation column.
                header=''.join(new.find(q('tr')).itertext())
                widths=[1250,2000,5774] if 'Significado' in header else [850,6674,1500]
            else: widths=[2200,6824]
            grid=new.find(q('tblGrid'))
            if grid is not None:
                for col,width in zip(grid,widths):col.set(q('w'),str(width))
            tp=new.find(q('tblPr'))
            if tp is None:tp=E.Element(q('tblPr'));new.insert(0,tp)
            setchild(tp,'tblW',w=9024,type='dxa')
            margins=setchild(tp,'tblCellMar')
            for side in ['top','bottom','left','right']:setchild(margins,side,w=85,type='dxa')
            for row in new.findall(q('tr')):
                trp=row.find(q('trPr'))
                if trp is None:trp=E.Element(q('trPr'));row.insert(0,trp)
                setchild(trp,'cantSplit')
                for cell,width in zip(row.findall(q('tc')),widths):
                    cp=cell.find(q('tcPr'))
                    setchild(cp,'tcW',w=width,type='dxa')
        body.append(new)
    body.append(section)
    parts['word/document.xml']=xml(root)
    parts['word/styles.xml']=xml(styles)
    final=OUT/name
    with ZipFile(final,'w',ZIP_DEFLATED) as z:
        for info in infos:z.writestr(info,parts[info.filename])
    # Package fidelity and complete content coverage.
    with ZipFile(template) as ref, ZipFile(final) as z:
        assert z.testzip() is None
        changed=[n for n in ref.namelist() if ref.read(n)!=z.read(n)]
        assert set(changed)=={'word/document.xml','word/styles.xml'},changed
        result=E.fromstring(z.read('word/document.xml'))
        assert E.tostring(result.find('.//w:sectPr',NS))==E.tostring(section)
        texts=result.xpath('//w:body//w:t/text()',namespaces=NS)
        assert 'ERM-465RJ-010+550-BSO-FUN-DE-F1-001-R00B.dwg' in texts
        assert 'Lorem ipsum' not in '\n'.join(texts)
        for text in sourcebody.xpath('.//w:t/text()',namespaces=NS):assert text in texts,text
        assert digest(template.read_bytes())==originalhash
        assert len(result.xpath('//w:sectPr',namespaces=NS))==1
    (work/'verification.json').write_text(json.dumps({'output':str(final),'changed_parts':changed,'all_other_parts_identical':True,'all_source_text_preserved':True,'template_unchanged':True},ensure_ascii=False,indent=2),encoding='utf-8')
    print(final)
