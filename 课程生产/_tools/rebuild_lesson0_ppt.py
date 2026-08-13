from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
import shutil, tempfile, re, sys

src, dst = map(Path, sys.argv[1:3])
slides = [
('课程说明会与学习地图','第 00 课｜先看清地图，再开始学习'),
('你现在最大的困惑是什么','方向不清｜技术焦虑｜没有项目｜求职表达｜工具碎片'),
('今天完成四件事','课程初心｜三阶段路径｜透明约定｜第一次场景思考'),
('为什么做这门课','初级产品经理需要一条从知识、项目到求职的完整转型路径'),
('我观察到的两种课程偏向','更偏技术原理｜更偏工具操作\n这是基于有限课程样本的个人观察'),
('AI 产品经理不是 AI 工具产品经理','工具能力很重要，但不能替代业务问题、产品判断和结果验证'),
('工具仍然重要','研究｜写作｜搭建 Demo｜表达方案｜提高效率'),
('课程试图连接两端','理解产品经理够用的技术\n始终从业务问题和产品结果出发'),
('三阶段是一条成果链','系统学习 → 项目实践 → 求职转化'),
('第一阶段：建立体系','基础认知｜产品形态｜方案设计｜评估治理｜成本立项'),
('第一阶段热门内容地图','Prompt｜Workflow｜RAG｜Agent｜MCP｜Vibe Coding｜评估｜ROI'),
('第二阶段：完成项目','1 个公共项目 + 2 个方向项目 + 1 个主项目深化'),
('为什么采用小组和不同项目','多个行业案例互学｜共享资料和 Bad Case｜每个人保留独立产出'),
('项目必须真实标注','课程项目｜个人项目｜业务改造方案\n不伪造上线和业务数据'),
('第三阶段：求职转化','岗位定位｜简历｜作品集｜项目讲述｜面试｜投递复盘'),
('三阶段分别解决什么','你懂不懂｜你做没做｜别人能否看懂和相信'),
('一份透明的课程说明','AI 变化快｜课程会更新｜讲师也在持续学习'),
('什么保持稳定，什么可能变化','稳定：阶段目标、能力主线、核心交付\n变化：工具、案例、部分顺序和扩展内容'),
('更新不是随意改课','说明调整原因｜影响范围｜新旧版本｜材料同步'),
('课程如何保证质量','来源追溯｜实际验证｜专业审核｜公开勘误｜持续更新'),
('学习不是听完','课前问题 → 课中判断 → 课后重做 → 反馈 → 修订'),
('为什么必须动手','眼睛会、脑子会，不等于手会'),
('模板的正确用法','第一遍防遗漏｜第二遍按项目调整｜第三遍形成自己的方法'),
('两类表达有什么不同','“我搭了一个客服机器人”\n“我解决查询慢和口径不一，并设计了验证流程”'),
('AI 产品经理的价值','把 AI 能力转化为可用、可控、可评估、可持续的产品'),
('全课程知识地图','基础 → 场景与 PRD → Workflow/RAG/Agent → 原型 → 评估治理 → 成本立项'),
('四个热门概念，一句话理解','Prompt 管输入｜Workflow 管流程｜RAG 管知识｜Agent 管任务执行'),
('贯穿场景：本地生活商家助手','商家规则多、查询慢、回答口径不一\n基于公开信息与教学假设'),
('第一次场景思考四问','谁遇到什么问题？｜现在怎样解决？｜AI 帮助什么？｜如何验证？'),
('互评只问三个问题','场景具体吗？｜AI 必要吗？｜结果可验证吗？'),
('今天带走四个结论','工具不等于产品｜三阶段递进｜动态更新但质量透明｜从真实问题开始'),
('作业与下一课','能力与学习目标｜第一次场景思考卡\n下一课：AI 产品经理岗位认知与转型路径'),
]

P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'a':A,'p':P}
ET.register_namespace('p',P); ET.register_namespace('a',A); ET.register_namespace('r',R)

def q(ns,tag): return '{%s}%s'%(ns,tag)

def add_text_shape(sp_tree, shape_id, name, text, x, y, cx, cy, size, bold=False, color='17365D'):
    sp=ET.SubElement(sp_tree,q(P,'sp'))
    nv=ET.SubElement(sp,q(P,'nvSpPr'))
    ET.SubElement(nv,q(P,'cNvPr'),{'id':str(shape_id),'name':name})
    ET.SubElement(nv,q(P,'cNvSpPr'),{'txBox':'1'})
    ET.SubElement(nv,q(P,'nvPr'))
    sppr=ET.SubElement(sp,q(P,'spPr'))
    xfrm=ET.SubElement(sppr,q(A,'xfrm'))
    ET.SubElement(xfrm,q(A,'off'),{'x':str(x),'y':str(y)})
    ET.SubElement(xfrm,q(A,'ext'),{'cx':str(cx),'cy':str(cy)})
    geom=ET.SubElement(sppr,q(A,'prstGeom'),{'prst':'rect'}); ET.SubElement(geom,q(A,'avLst'))
    ET.SubElement(sppr,q(A,'noFill')); ln=ET.SubElement(sppr,q(A,'ln')); ET.SubElement(ln,q(A,'noFill'))
    tx=ET.SubElement(sp,q(P,'txBody'))
    ET.SubElement(tx,q(A,'bodyPr'),{'wrap':'square','anchor':'ctr','lIns':'0','rIns':'0','tIns':'0','bIns':'0'})
    ET.SubElement(tx,q(A,'lstStyle'))
    for idx,line in enumerate(text.split('\n')):
        para=ET.SubElement(tx,q(A,'p'))
        r=ET.SubElement(para,q(A,'r'))
        attrs={'lang':'zh-CN','sz':str(size)}
        if bold: attrs['b']='1'
        rp=ET.SubElement(r,q(A,'rPr'),attrs)
        fill=ET.SubElement(rp,q(A,'solidFill')); ET.SubElement(fill,q(A,'srgbClr'),{'val':color})
        ET.SubElement(r,q(A,'t')).text=line
        ET.SubElement(para,q(A,'endParaRPr'),{'lang':'zh-CN','sz':str(size)})
    return sp
with tempfile.TemporaryDirectory() as td:
    root = Path(td)
    with ZipFile(src) as z: z.extractall(root)
    slide_files = sorted((root/'ppt/slides').glob('slide*.xml'), key=lambda p:int(re.search(r'\d+',p.stem).group()))
    if len(slide_files) < len(slides): raise RuntimeError('template has too few slides')
    for i,(title,body) in enumerate(slides):
        p=slide_files[i]; tree=ET.parse(p); root_el=tree.getroot()
        sp_tree=root_el.find('.//p:spTree',NS)
        # Remove all old visible objects while retaining the mandatory group properties.
        for child in list(sp_tree):
            if child.tag not in (q(P,'nvGrpSpPr'),q(P,'grpSpPr')): sp_tree.remove(child)
        add_text_shape(sp_tree,2,'标题',title,700000,600000,10800000,950000,3000,True,'17365D')
        add_text_shape(sp_tree,3,'正文',body,850000,1800000,10400000,3000000,2100,False,'222222')
        add_text_shape(sp_tree,4,'页码',f'{i+1:02d}',10800000,6100000,700000,350000,1100,False,'7F8C8D')
        tree.write(p,encoding='UTF-8',xml_declaration=True)
    # Remove unused template slides from the presentation list and relationships.
    pns={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    relns={'pr':'http://schemas.openxmlformats.org/package/2006/relationships'}
    pres_path=root/'ppt/presentation.xml'; pres_tree=ET.parse(pres_path)
    sld_list=pres_tree.find('.//p:sldIdLst',pns); sld_ids=list(sld_list)
    remove_rids=[]
    for node in sld_ids[len(slides):]:
        remove_rids.append(node.attrib['{'+pns['r']+'}id']); sld_list.remove(node)
    pres_tree.write(pres_path,encoding='UTF-8',xml_declaration=True)
    rel_path=root/'ppt/_rels/presentation.xml.rels'; rel_tree=ET.parse(rel_path)
    rel_root=rel_tree.getroot()
    for node in list(rel_root):
        if node.attrib.get('Id') in remove_rids: rel_root.remove(node)
    rel_tree.write(rel_path,encoding='UTF-8',xml_declaration=True)
    for p in slide_files[len(slides):]:
        p.unlink(missing_ok=True)
        (p.parent/'_rels'/(p.name+'.rels')).unlink(missing_ok=True)
    with ZipFile(dst,'w',ZIP_DEFLATED) as z:
        for p in root.rglob('*'):
            if p.is_file(): z.write(p,p.relative_to(root))
