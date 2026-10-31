# -*- coding: utf-8 -*-
"""把 Visio 导出的 EMF 以**矢量**形式插入 docx（python-docx 原生不支持 EMF）。

用法：
    from word_emf import add_emf
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_emf(p, "fig3-1.emf", w_cm=15.0, h_cm=10.2)

注意：
  · w_cm / h_cm 必须等于该图在 Visio 里的页面尺寸，才能保证图内字号 1:1；
  · 默认四边各裁 1.4%（CROP）以去掉 Visio EMF 导出的灰色发丝边框，
    同时把外框尺寸放大 1/(1-2*CROP)，裁边后图形仍为原尺寸；
  · 若图内元素贴边，请把 Visio 页面留白调到 ≥0.35cm，否则会被裁掉。
"""
from docx.shared import Cm
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from docx.opc.constants import RELATIONSHIP_TYPE as RT

CROP = 0.014          # 四边裁剪比例


def add_emf(paragraph, emf_path, w_cm, h_cm, crop=CROP, keep_aspect=True):
    blob = open(emf_path, "rb").read()
    pkg = paragraph.part.package
    idx = len([p for p in pkg.iter_parts() if str(p.partname).endswith(".emf")]) + 1
    part = Part(PackURI("/word/media/fig_%d.emf" % idx), "image/x-emf", blob, pkg)
    rid = paragraph.part.relate_to(part, RT.IMAGE)

    k = 1.0 / (1.0 - 2 * crop) if crop else 1.0
    cx = int(round(w_cm * 360000 * k))      # 1cm = 360000 EMU
    cy = int(round(h_cm * 360000 * k))
    src = ('<a:srcRect l="%d" t="%d" r="%d" b="%d"/>'
           % tuple(int(crop * 100000) for _ in range(4))) if crop else ""
    sid = idx + 900
    xml = (
        '<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="%d" name="Figure %d"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic><pic:nvPicPr><pic:cNvPr id="%d" name="fig.emf"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/>%s<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    ) % (cx, cy, sid, sid, sid, rid, src, cx, cy)
    from lxml import etree
    paragraph._p.append(etree.fromstring(xml))
    return cx, cy


def build_master_vsdx(app, entries, out_path, title_page=None):
    """把多张图汇总成一个多页 .vsdx（便于交付给需要继续编辑的人）。

    entries: [(页面名, 无参函数 → Draw)]，函数内部创建 Draw 即可；
             需要配合 visio_engine.set_master(doc) 使用。
    """
    import visio_engine as ve
    doc = app.Documents.Add("")
    ve.set_master(doc)
    for name, fn in entries:
        d = fn()
        d.page.Name = name
    ve.set_master(None)
    if title_page:
        try:
            pg = doc.Pages.Item(1)
            pg.Name = "图集说明"
            sh = pg.DrawRectangle(0.6, 0.6, 7.67, 11.09)
            sh.Text = title_page
            sh.Cells("LinePattern").FormulaU = "0"
            sh.Cells("FillPattern").FormulaU = "0"
            sh.Cells("Char.Font").FormulaU = '"宋体"'
            sh.Cells("Char.Size").FormulaU = "11 pt"
            sh.Cells("Para.HorzAlign").FormulaU = "0"
            sh.Cells("VerticalAlign").FormulaU = "0"
        except Exception:
            pass
    doc.SaveAs(out_path)
    return doc
