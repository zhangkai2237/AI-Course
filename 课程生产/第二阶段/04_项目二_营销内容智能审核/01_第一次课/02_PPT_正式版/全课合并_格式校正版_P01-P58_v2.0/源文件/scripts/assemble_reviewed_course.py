from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from collections import Counter
import posixpath,copy,json,re,hashlib
ROOT=Path('/Users/didi/Documents/Codex/2026-10-18/wo-l')
p=ROOT/'work/ppt/lesson1_complete_review_ppt169_20261018'
for d in ['analysis','sources','exports','validation']: (p/d).mkdir(parents=True,exist_ok=True)
files=[ROOT/'outputs'/x for x in ['第一次课_重做样稿_P01-P06_v0.1.pptx','第一次课_第二节_项目目标_P07-P12_v0.1.pptx','第一次课_第三节_业务方案_P13-P17_重点版_v0.2.pptx','第一次课_第四节_审核点定义管理与调用_P18-P26_v0.1.pptx','第一次课_第五节_方案选择与调用流程_P27-P33_v0.1.pptx','第一次课_第六节_系统与产品方案_P34-P48_v0.1.pptx']]
files.append(sorted((ROOT/'work/ppt/lesson1_section7_native_ppt169_20261018/exports').glob('*.pptx'))[-1])
A='http://schemas.openxmlformats.org/drawingml/2006/main';P='http://schemas.openxmlformats.org/presentationml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';PK='http://schemas.openxmlformats.org/package/2006/relationships';CT='http://schemas.openxmlformats.org/package/2006/content-types'
ns={'a':A,'p':P,'r':R}
def xml(b):return E.fromstring(b)
def dump(x):return E.tostring(x,encoding='UTF-8',xml_declaration=True,standalone=True)
def relpath(part):return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')
def owner(r):
 if r=='_rels/.rels':return ''
 return posixpath.join(posixpath.dirname(posixpath.dirname(r)),posixpath.basename(r)[:-5])
def target(part,t):return posixpath.normpath(posixpath.join(posixpath.dirname(part),t)) if not t.startswith('/') else t.lstrip('/')
def order(entries):
 pr=xml(entries['ppt/presentation.xml']);rr=xml(entries['ppt/_rels/presentation.xml.rels']);refs={r.get('Id'):target('ppt/presentation.xml',r.get('Target')) for r in rr}
 return [refs[x.get('{'+R+'}id')] for x in pr.find('p:sldIdLst',ns)]
def section(i):
 return next(s for a,b,s in [(1,6,1),(7,12,2),(13,17,3),(18,26,4),(27,33,5),(34,48,6),(49,58,7)] if a<=i<=b)
with ZipFile(files[0]) as z:out={x:z.read(x) for x in z.namelist() if not x.endswith('/')}
ct=xml(out['[Content_Types].xml']);pr=xml(out['ppt/presentation.xml']);rr=xml(out['ppt/_rels/presentation.xml.rels']);slideparts=order(out)
source_records=[]
for num,f in enumerate(files):
 with ZipFile(f) as z:ent={x:z.read(x) for x in z.namelist() if not x.endswith('/')}
 parts=order(ent);source_records.append({'path':str(f),'slides':len(parts),'sha256':hashlib.sha256(f.read_bytes()).hexdigest()})
 if num==0:continue
 # Import all PPT parts with unique names. Every relationship is rewritten to its mapped target.
 mapping={name:posixpath.join(posixpath.dirname(name),f'd{num}_'+posixpath.basename(name)) for name in ent if name.startswith('ppt/') and not name.endswith('.rels') and name!='ppt/presentation.xml'}
 for old,new in mapping.items():out[new]=ent[old]
 for old,new in mapping.items():
  rel=relpath(old)
  if rel not in ent:continue
  tree=xml(ent[rel])
  for relnode in tree:
   if relnode.get('TargetMode')=='External':continue
   orig=target(old,relnode.get('Target'))
   assert orig in mapping,(f.name,old,orig)
   relnode.set('Target',posixpath.relpath(mapping[orig],posixpath.dirname(new)))
  out[relpath(new)]=dump(tree)
 sct=xml(ent['[Content_Types].xml'])
 for node in sct:
  if E.QName(node).localname=='Default':
   same=[x for x in ct if x.get('Extension')==node.get('Extension')]
   if same:assert same[0].get('ContentType')==node.get('ContentType')
   else:ct.append(copy.deepcopy(node))
  else:
   old=node.get('PartName','').lstrip('/')
   if old in mapping:
    c=copy.deepcopy(node);c.set('PartName','/'+mapping[old]);ct.append(c)
 # Register imported masters in the presentation root, while keeping original layout/theme graph.
 for listtag,childtag,typ in [('sldMasterIdLst','sldMasterId','slideMaster'),('notesMasterIdLst','notesMasterId','notesMaster')]:
  newlist=pr.find('p:'+listtag,ns)
  if newlist is None:
   newlist=E.Element('{'+P+'}'+listtag)
   # Presentation schema: slide masters, notes masters, handout masters, slide list.
   idx=1 if listtag=='notesMasterIdLst' else 0;pr.insert(idx,newlist)
  masterparts=[v for k,v in mapping.items() if re.fullmatch(r'ppt/'+('slideMasters/slideMaster' if typ=='slideMaster' else 'notesMasters/notesMaster')+r'\d+\.xml',k)]
  for mp in masterparts:
   rid='rId'+str(max([int(x.get('Id')[3:]) for x in rr if re.fullmatch(r'rId\d+',x.get('Id',''))]+[0])+1)
   E.SubElement(rr,'{'+PK+'}Relationship',Id=rid,Type=R+'/'+typ,Target=posixpath.relpath(mp,'ppt'))
   attrs={'{'+R+'}id':rid}
   if typ=='slideMaster':attrs['id']=str(2147483648+num*10+len(newlist))
   E.SubElement(newlist,'{'+P+'}'+childtag,**attrs)
 for part in parts:
  rid='rId'+str(max([int(x.get('Id')[3:]) for x in rr if re.fullmatch(r'rId\d+',x.get('Id',''))]+[0])+1)
  E.SubElement(rr,'{'+PK+'}Relationship',Id=rid,Type=R+'/slide',Target=posixpath.relpath(mapping[part],'ppt'))
  E.SubElement(pr.find('p:sldIdLst',ns),'{'+P+'}sldId',id=str(256+len(slideparts)),**{'{'+R+'}id':rid})
  slideparts.append(mapping[part])
assert len(slideparts)==58
# Canonical green palette and editable footer from the confirmed native deck.
with ZipFile(files[3]) as z:
 theme=xml(z.read('ppt/theme/theme1.xml'));palette=theme.find('.//a:clrScheme',ns)
 # Pick a content slide with full navigation, independent of OOXML part number.
 for fn in z.namelist():
  if fn.startswith('ppt/slides/') and fn.endswith('.xml'):
   r=xml(z.read(fn))
   if '6 产品方案' in ''.join(r.xpath('//a:t/text()',namespaces=ns)):
    footer=next(g for g in r.find('p:cSld/p:spTree',ns) if E.QName(g).localname=='grpSp' and (g.find('p:grpSpPr/a:xfrm/a:off',ns) is not None) and int(g.find('p:grpSpPr/a:xfrm/a:off',ns).get('y'))==6286500)
    break
for name,data in list(out.items()):
 if name.startswith('ppt/theme/') and name.endswith('.xml'):
  r=xml(data);old=r.find('.//a:clrScheme',ns)
  if old is not None:old.getparent().replace(old,copy.deepcopy(palette))
  out[name]=dump(r)
chapters={1,7,13,18,27,34,49};audit=[]
for idx,part in enumerate(slideparts,1):
 r=xml(out[part]);tree=r.find('p:cSld/p:spTree',ns);sec=section(idx)
 # Unify first 12 content titles with native template's 42px (31.5pt) title hierarchy.
 if idx<=12 and idx not in {1,7}:
  for sp in r.xpath('//p:sp',namespaces=ns):
   off=sp.find('p:spPr/a:xfrm/a:off',ns);texts=sp.xpath('.//a:t/text()',namespaces=ns)
   if off is None or not texts:continue
   y=int(off.get('y'))/9525;sizes=[int(x) for x in sp.xpath('.//a:rPr/@sz',namespaces=ns)]
   if sizes and max(sizes)>=2400 and 65<=y<=130:
    for rp in sp.xpath('.//a:rPr',namespaces=ns):rp.set('sz','3150');rp.set('b','1')
 # Footer format: standard green navigation on content slides; chapter slides keep dark treatment.
 if idx not in chapters:
  removed=[]
  for sp in list(r.xpath('//p:sp | //p:pic | //p:cxnSp',namespaces=ns)):
   off=sp.find('p:spPr/a:xfrm/a:off',ns)
   if off is not None and int(off.get('y'))>=6286500:
    removed+=sp.xpath('.//a:t/text()',namespaces=ns);sp.getparent().remove(sp)
  new=copy.deepcopy(footer)
  currentmax=max([int(x) for x in r.xpath('//p:cNvPr/@id',namespaces=ns)]+[1])
  for j,x in enumerate(new.xpath('.//p:cNvPr',namespaces=ns),currentmax+1):x.set('id',str(j))
  navlabels=['1 业务背景','2 项目目标','3 业务方案','4 审核体系','5 方案选择','6 产品方案','7 落地验证']
  navsh=[sp for sp in new.xpath('.//p:sp',namespaces=ns) if re.match(r'^[1-7] ',''.join(sp.xpath('.//a:t/text()',namespaces=ns)))]
  assert len(navsh)==7
  for n,sp in enumerate(navsh,1):
   txt=sp.find('.//a:t',ns);txt.text=navlabels[n-1]
   for extra in sp.xpath('.//a:t',namespaces=ns)[1:]:extra.text=''
   ext=sp.find('p:spPr/a:xfrm/a:ext',ns);ext.set('cx',str(110*9525))
   for rp in sp.xpath('.//a:rPr',namespaces=ns):
    rp.set('b','1' if n==sec else '0');fill=rp.find('a:solidFill',ns)
    if fill is not None:rp.remove(fill)
    fill=E.SubElement(rp,'{'+A+'}solidFill');E.SubElement(fill,'{'+A+'}srgbClr',val='126B63' if n==sec else '617673')
  # The footer's page label is the rightmost text-bearing shape.
  pg=max([sp for sp in new.xpath('.//p:sp',namespaces=ns) if sp.xpath('.//a:t',namespaces=ns)],key=lambda sp:int(sp.find('p:spPr/a:xfrm/a:off',ns).get('x')))
  pg.find('.//a:t',ns).text=f'{idx:02}'
  tree.append(new);audit.append({'page':idx,'footer':'7项导航；当前章节高亮','replaced_footer':removed})
 else:
  for t in r.xpath('//a:t',namespaces=ns):
   if re.fullmatch(r'\d{1,2}( / 70)?',t.text or '') and int((t.text or '').split()[0])==idx:t.text=f'{idx:02}'
  # Match dark chapter background on earlier blue chapter/cover.
  if idx in {1,7}:
   bg=r.find('p:cSld/p:bg/p:bgPr',ns)
   if bg is not None:
    for fill in list(bg):bg.remove(fill)
    fill=E.SubElement(bg,'{'+A+'}solidFill');E.SubElement(fill,'{'+A+'}srgbClr',val='17302F');E.SubElement(bg,'{'+A+'}effectLst')
 out[part]=dump(r)
out['ppt/presentation.xml']=dump(pr);out['ppt/_rels/presentation.xml.rels']=dump(rr);out['[Content_Types].xml']=dump(ct)
# Check complete relationship graph and unique native shape ids before writing.
for name,b in out.items():
 if name.endswith('.rels'):
  own=owner(name)
  for rel in xml(b):
   if rel.get('TargetMode')!='External':assert target(own,rel.get('Target')) in out,(name,rel.get('Target'))
for part in slideparts:
 r=xml(out[part]);ids=r.xpath('//p:cNvPr/@id',namespaces=ns);assert len(ids)==len(set(ids)),part
path=p/'exports/第一次课_完整授课PPT_格式校正版_v2.0.pptx'
with ZipFile(path,'w',ZIP_DEFLATED) as z:
 for k,v in out.items():z.writestr(k,v)
(p/'analysis/source_manifest.json').write_text(json.dumps(source_records,ensure_ascii=False,indent=2))
(p/'analysis/format_changes.json').write_text(json.dumps({'slides':58,'palette':'结构强化版绿色配色','fonts':'保留Microsoft YaHei主题字体，统一前12页内容标题为31.5pt','navigation':'内容页统一7项章节导航、当前节高亮、连续两位页码；章节页保留深色设计','screenshot_blanks':'P39—P43保留5处空白','details':audit},ensure_ascii=False,indent=2))
(p/'analysis/slide_order.json').write_text(json.dumps(slideparts,ensure_ascii=False,indent=2))
conf=(ROOT/'work/ppt/lesson1_section7_native_ppt169_20261018/validation/fonts.conf').read_text().replace('lesson1_section7_native_ppt169_20261018','lesson1_complete_review_ppt169_20261018')
(p/'validation/fonts.conf').write_text(conf)
print(path)
