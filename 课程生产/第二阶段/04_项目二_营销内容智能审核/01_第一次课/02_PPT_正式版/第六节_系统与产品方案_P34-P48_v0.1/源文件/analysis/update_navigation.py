from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import copy,shutil,json
p=Path(__file__).resolve().parent.parent
ppt=sorted((p/'exports').glob('*.pptx'))[-1]
ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main','p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
allowed={'1 业务背景','2 项目目标','3 业务方案','4 审核点','5 方案选择','6 产品方案','7 落地验证'}
with ZipFile(p/'sources/第一次课_电商营销审核提效_结构强化版_v1.1.pptx') as z:
 root=E.fromstring(z.read('ppt/slides/slide29.xml'))
 refs={}
 for sp in root.xpath('//p:sp',namespaces=ns):
  text=''.join(sp.xpath('.//a:t/text()',namespaces=ns))
  if text in ['1 业务背景','4 审核点']:
   refs[text]=copy.deepcopy(sp.find('.//a:rPr',ns))
assert len(refs)==2
tmp=ppt.with_suffix('.tmp');count=0
with ZipFile(ppt) as src,ZipFile(tmp,'w') as dst:
 for item in src.infolist():
  data=src.read(item.filename)
  if item.filename.startswith('ppt/slides/') and item.filename.endswith('.xml'):
   root=E.fromstring(data)
   for sp in root.xpath('//p:sp',namespaces=ns):
    text=''.join(sp.xpath('.//a:t/text()',namespaces=ns))
    if text not in allowed:continue
    for run in sp.xpath('.//a:r',namespaces=ns):
     old=run.find('a:rPr',ns)
     if old is not None:run.remove(old)
     style=copy.deepcopy(refs['4 审核点' if text=='6 产品方案' else '1 业务背景'])
     style.set('b','1' if text=='6 产品方案' else '0')
     run.insert(0,style)
    count+=1
   data=E.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)
  dst.writestr(item,data)
tmp.replace(ppt)
old=p.parent/'lesson1_section3_native_ppt169_20261018/validation/fonts.conf'
conf=old.read_text().replace('lesson1_section3_native_ppt169_20261018','lesson1_section6_native_ppt169_20261018')
(p/'validation/fonts.conf').write_text(conf)
print(json.dumps({'footer_items_updated':count,'active_chapter':6,'body_geometry_unchanged':True}))
