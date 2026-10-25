from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json
p=Path('/Users/didi/Documents/Codex/2026-10-18/wo-l/work/ppt/lesson1_complete_review_ppt169_20261018')
f=p/'exports/第一次课_完整授课PPT_格式校正版_v2.0.pptx';parts=json.loads((p/'analysis/slide_order.json').read_text())
A='http://schemas.openxmlformats.org/drawingml/2006/main';ns={'a':A,'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}
with ZipFile(f) as z:ent={x:z.read(x) for x in z.namelist()}
def settext(sp,text):
 body=sp.find('p:txBody',ns);pars=body.findall('a:p',ns);orig=pars[0];style=copy.deepcopy(orig.find('.//a:rPr',ns));ppr=orig.find('a:pPr',ns)
 for c in pars:body.remove(c)
 para=E.SubElement(body,'{'+A+'}p')
 if ppr is not None:para.append(copy.deepcopy(ppr))
 for i,line in enumerate(text.split('\n')):
  if i:E.SubElement(para,'{'+A+'}br')
  run=E.SubElement(para,'{'+A+'}r')
  if style is not None:run.append(copy.deepcopy(style))
  E.SubElement(run,'{'+A+'}t').text=line
fixes=[]
for idx,part in enumerate(parts,1):
 r=E.fromstring(ent[part])
 if idx==1:
  bg=r.find('p:cSld/p:bg/p:bgPr/a:solidFill',ns)
  for x in list(bg):bg.remove(x)
  E.SubElement(bg,'{'+A+'}srgbClr',val='F7F8F4');fixes.append({'page':1,'fix':'封面保留浅底，恢复深色主标题对比度'})
 for sp in r.xpath('//p:sp',namespaces=ns):
  t=''.join(sp.xpath('.//a:t/text()',namespaces=ns));off=sp.find('p:spPr/a:xfrm/a:off',ns);ext=sp.find('p:spPr/a:xfrm/a:ext',ns)
  def ypos(v):off.set('y',str(v*9525))
  def width(v):ext.set('cx',str(v*9525))
  if idx==2 and t.startswith('180分钟'):ypos(630)
  if idx==3 and t.startswith('读原句 → 找依据'):
   settext(sp,'读原句 → 找依据\n核对、反馈、复核');fixes.append({'page':3,'fix':'缩短审核员动作行，消除与说明文字重叠'})
  if idx==5:
   if t.startswith('“本次活动到手1799元，'):
    settext(sp,'“到手1799元，\n人人同价！”');fixes.append({'page':5,'fix':'待审引句显式分两行，避免覆盖核对提示'})
   if t.startswith('2399元 → 1999元'):settext(sp,'2399 → 1999 → 1799元')
   if t.startswith('这里只演示价格核对'):ypos(633)
  if idx==6 and t.startswith('记录作者与审核员'):ypos(635)
  if idx==7 and t=='约15分钟':width(125);off.set('x',str(1100*9525))
  if idx==8 and t.startswith('这些是首版期望改善'):ypos(630)
  if idx==9:
   if t.startswith('资料或关联不齐'):
    settext(sp,'资料或关联不齐，先补齐或缩小试用范围；已查项无问题，不等于整篇审核完成。');width(1120)
   if t.startswith('系统未发现已覆盖项'):settext(sp,'')
  if idx==12:
   if t.startswith('下一节 · 业务方案'):ypos(590)
   if t.startswith('谁发起 →'):ypos(620);width(1100)
  if idx==40 and t.startswith('检查完成、人工待审'):settext(sp,'检查与人工审核，\n分别显示状态。')
  if idx==47 and t.startswith('资料冲突时带原句'):
   settext(sp,'资料冲突带依据交人确认；异常要有状态、动作和责任人。');fixes.append({'page':47,'fix':'底部收束语缩短，避开页脚'})
 ent[part]=E.tostring(r,encoding='UTF-8',xml_declaration=True,standalone=True)
with ZipFile(f,'w',ZIP_DEFLATED) as z:
 for name,data in ent.items():z.writestr(name,data)
(p/'analysis/visual_fixes.json').write_text(json.dumps({'fixes':fixes,'additional':'P02、P06、P08、P09、P12底部留白与说明位置调整，P07时长文字框扩宽，P40说明避免单字断行。'},ensure_ascii=False,indent=2))
print('Visual overlap, footer clearance and cover contrast fixes applied.')
