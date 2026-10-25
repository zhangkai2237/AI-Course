from pathlib import Path
from PIL import Image
from pypdf import PdfReader
from zipfile import ZipFile
import json,shutil,html,re
p=Path(__file__).resolve().parent.parent
root=Path('/Users/didi/Documents/Codex/2026-10-18/wo-l');out=root/'outputs'
base='第一次课_第三节_业务方案_P13-P17_重点版_v0.2'
ppt=sorted((p/'exports').glob('*.pptx'))[-1]
pdf=next((p/'validation/concise_render').glob('*.pdf'))
assert len(PdfReader(pdf).pages)==5
with ZipFile(ppt) as z:
 slides=[f for f in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+.xml',f)]
 notes=[f for f in z.namelist() if re.fullmatch(r'ppt/notesSlides/notesSlide\d+.xml',f)]
 assert len(slides)==len(notes)==5
 assert sum(z.read(f).count(b'<a:t>') for f in slides)>70
shutil.copy2(ppt,out/(base+'.pptx'));shutil.copy2(pdf,out/(base+'_浏览版.pdf'))
titles=['第三节标题页','三个环节总览','事前自检','事中审核','事后巡检']
names=[];montage=Image.new('RGB',(1296,1112),'#D3DFDA')
for i,t in enumerate(titles):
 src=p/'validation/concise_render'/f'page-{i+1}.png'
 name=f'第一次课_第三节_P{i+13}_{t}_重点版_v0.2.png'
 shutil.copy2(src,out/name);names.append(name)
 im=Image.open(src).convert('RGB').resize((640,360))
 montage.paste(im,((i%2)*656,(i//2)*376))
montage_name='第一次课_第三节_五页重点版预览_v0.2.png';montage.save(out/montage_name)
figures=''.join(f'<figure id="p{i+13}"><figcaption>P{i+13} · {t}</figcaption><img src="{html.escape(names[i],quote=True)}" alt="{t}"></figure>' for i,t in enumerate(titles))
page='''<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>第三节·业务方案：自检、审核与巡检</title><style>body{margin:0;background:#edf2ef;color:#203d37;font:18px/1.7 -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif}main{max-width:1280px;margin:36px auto;padding:0 24px}h1{font-size:32px;margin:0}a{color:#176d61}nav{display:flex;gap:20px;flex-wrap:wrap;margin:18px 0}figure{margin:28px 0 42px}figcaption{font-size:24px;font-weight:700;margin:12px 0}img{display:block;width:100%;height:auto;box-shadow:0 6px 28px #172c4d15}</style><main><h1>第一次课 · 第三节：业务方案</h1><p>5页重点版，连续页码P13—P17。章节标题、三环节总览，以及自检、审核、巡检各一页。每页说明谁在什么时候做、主要负责什么，以及结果怎样交接。</p><p>自检提前准备；审核正式确认；巡检跟进实际更新。页面复用结构强化版原生模板，保留可编辑文字和图形，5页均含教师备注。</p><nav><a href="'''+base+'''.pptx">下载可编辑PPT</a><a href="'''+base+'''_浏览版.pdf">浏览版PDF</a><a href="第一次课_第二节_项目目标_P07-P12_v0.1_预览.html">返回第二节</a><a href="第一次课_完整老师讲义_常规审核版_v1.0.html#chapter3">对应完整讲义</a></nav>'''+figures+'''<p>画面由实际PPTX渲染。品牌、人物和案例均为教学模拟。讲解细节保留在教师备注中。</p></main></html>'''
(out/(base+'_预览.html')).write_text(page)
meta={'base':base,'files':[base+'.pptx',base+'_浏览版.pdf',base+'_预览.html',montage_name]+names,'slides':5,'notes':5,'scope':'本节仅解释自检、审核和巡检三个环节的主要职责及衔接。'}
(p/'analysis/deliverables_concise.json').write_text(json.dumps(meta,ensure_ascii=False,indent=2))
print(json.dumps({'files':len(meta['files']),'slides':5,'native_text_and_shapes':True,'notes':5},ensure_ascii=False))
