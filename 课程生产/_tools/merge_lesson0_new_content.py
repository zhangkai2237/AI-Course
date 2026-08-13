from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED
import re
import xml.etree.ElementTree as ET

ROOT = Path('/Users/didi/Project/AI课程/课程生产/第一阶段/第00课_课程说明会学习地图与第一次AI场景思考_完整课程包_v2')
TARGET = ROOT / '06_正式PPT.pptx'
SOURCE = Path('/tmp/lesson0_new_content.pptx')

NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
ET.register_namespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
ET.register_namespace('p', NS_P)
ET.register_namespace('r', NS_R)

with TemporaryDirectory() as td:
    td = Path(td)
    base_dir, src_dir = td/'base', td/'src'
    with ZipFile(TARGET) as z: z.extractall(base_dir)
    with ZipFile(SOURCE) as z: z.extractall(src_dir)

    pres_file = base_dir/'ppt/presentation.xml'
    rel_file = base_dir/'ppt/_rels/presentation.xml.rels'
    pres_tree, rel_tree = ET.parse(pres_file), ET.parse(rel_file)
    pres_root, rel_root = pres_tree.getroot(), rel_tree.getroot()
    sld_list = pres_root.find(f'{{{NS_P}}}sldIdLst')

    rel_by_id = {x.attrib['Id']: x for x in list(rel_root)}
    # 删除原第27、29、30、32页，保留04/05/06章节页
    old = list(sld_list)
    for idx in (31,29,28,26):
        node = old[idx]
        rid = node.attrib[f'{{{NS_R}}}id']
        sld_list.remove(node)
        if rid in rel_by_id: rel_root.remove(rel_by_id[rid])

    slide_files = list((base_dir/'ppt/slides').glob('slide*.xml'))
    next_slide = max(int(re.search(r'(\d+)', x.stem).group(1)) for x in slide_files) + 1
    next_sid = max(int(x.attrib['id']) for x in list(sld_list)) + 1
    next_rid = max(int(re.search(r'(\d+)$', x.attrib['Id']).group(1)) for x in list(rel_root) if re.search(r'(\d+)$', x.attrib['Id'])) + 1
    source_slides = sorted((src_dir/'ppt/slides').glob('slide*.xml'), key=lambda x:int(re.search(r'(\d+)',x.stem).group(1)))
    positions = list(range(26,35)) + list(range(36,44)) + list(range(45,49))

    for src, pos in zip(source_slides, positions):
        src_no = int(re.search(r'(\d+)', src.stem).group(1))
        dst_no = next_slide; next_slide += 1
        (base_dir/f'ppt/slides/slide{dst_no}.xml').write_bytes(src.read_bytes())
        src_rel = src_dir/f'ppt/slides/_rels/slide{src_no}.xml.rels'
        if src_rel.exists():
            dst_rel = base_dir/f'ppt/slides/_rels/slide{dst_no}.xml.rels'
            dst_rel.parent.mkdir(parents=True, exist_ok=True)
            dst_rel.write_bytes(src_rel.read_bytes())
        rid = f'rId{next_rid}'; next_rid += 1
        ET.SubElement(rel_root, f'{{{NS_REL}}}Relationship', {
            'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
            'Target': f'slides/slide{dst_no}.xml'
        })
        sid = ET.Element(f'{{{NS_P}}}sldId', {'id': str(next_sid), f'{{{NS_R}}}id': rid}); next_sid += 1
        nodes = list(sld_list)
        if pos >= len(nodes): sld_list.append(sid)
        else: sld_list.insert(pos, sid)

    pres_tree.write(pres_file, encoding='UTF-8', xml_declaration=True)
    rel_tree.write(rel_file, encoding='UTF-8', xml_declaration=True)

    ct_file = base_dir/'[Content_Types].xml'
    ct_tree = ET.parse(ct_file); ct_root = ct_tree.getroot()
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    existing = {x.attrib.get('PartName') for x in ct_root}
    for no in range(33, next_slide):
        part = f'/ppt/slides/slide{no}.xml'
        if part not in existing:
            ET.SubElement(ct_root, f'{{{ct_ns}}}Override', {'PartName':part,'ContentType':'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'})
    ET.register_namespace('', ct_ns)
    ct_tree.write(ct_file, encoding='UTF-8', xml_declaration=True)

    out = td/'merged.pptx'
    with ZipFile(out, 'w', ZIP_DEFLATED) as z:
        for f in base_dir.rglob('*'):
            if f.is_file(): z.write(f, f.relative_to(base_dir).as_posix())
    TARGET.write_bytes(out.read_bytes())
    print(TARGET)
