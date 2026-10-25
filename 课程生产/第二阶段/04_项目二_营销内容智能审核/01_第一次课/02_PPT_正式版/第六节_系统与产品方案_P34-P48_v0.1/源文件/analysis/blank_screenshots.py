from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import copy,json
p=Path(__file__).resolve().parent.parent
ppt=sorted((p/'exports').glob('*.pptx'))[-1];tmp=ppt.with_suffix('.tmp')
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
# User-requested scoped edit: clear source screenshots and retain empty editable frames.
count=0
with ZipFile(ppt) as src,ZipFile(tmp,'w') as dst:
 for item in src.infolist():
  data=src.read(item.filename)
  if item.filename.startswith('ppt/slides/') and item.filename.endswith('.xml') and '第六节 / 6.2 页面介绍'.encode() in data:
   root=E.fromstring(data)
   pics=root.xpath('//p:pic',namespaces=ns);assert len(pics)==1
   pic=pics[0];sp=E.Element('{'+ns['p']+'}sp')
   nv=E.SubElement(sp,'{'+ns['p']+'}nvSpPr')
   E.SubElement(nv,'{'+ns['p']+'}cNvPr',id='5',name='截图预留区（可删除后粘贴截图）')
   E.SubElement(nv,'{'+ns['p']+'}cNvSpPr');E.SubElement(nv,'{'+ns['p']+'}nvPr')
   pr=copy.deepcopy(pic.find('p:spPr',ns));sp.append(pr)
   fill=E.SubElement(pr,'{'+ns['a']+'}solidFill');E.SubElement(fill,'{'+ns['a']+'}srgbClr',val='FFFFFF')
   ln=E.SubElement(pr,'{'+ns['a']+'}ln',w='12700');f=E.SubElement(ln,'{'+ns['a']+'}solidFill');E.SubElement(f,'{'+ns['a']+'}srgbClr',val='B7C9C2');E.SubElement(ln,'{'+ns['a']+'}prstDash',val='dash')
   pic.getparent().replace(pic,sp);count+=1
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
  dst.writestr(item,data)
assert count==5;tmp.replace(ppt)
(p/'validation/screenshot_placeholders.json').write_text(json.dumps({'pages':'P39—P43','empty_frames':count,'old_screenshots_removed':True,'frame_px':{'x':48,'y':177,'width':848,'height':477}},ensure_ascii=False,indent=2))
print('Five screenshots cleared; editable empty frames preserved.')
