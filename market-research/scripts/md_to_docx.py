# -*- coding: utf-8 -*-
"""市场调研报告 Markdown → Word(.docx) 转换脚本（market-research skill 自带）。

用法:
    python md_to_docx.py 报告.md                 # 输出同目录同名 .docx
    python md_to_docx.py 报告.md -o 输出.docx    # 指定输出路径

依赖: python-docx（首次使用: python -m pip install python-docx）

转换规则:
    - 正文引用编号 [n](#ref-n) → Word 内部超链接，跳转到"参考来源"条目上的书签
      （Word 书签名不允许连字符，转换时 ref-n 映射为书签 ref_n）
    - 裸 URL 与 [文本](http...) → 可点击外部超链接
    - 标题黑体、正文宋体小四、1.3 倍行距、正文首行缩进 2 字符
    - 表格 9pt、首行底纹、跨页重复表头、行不拆分、固定列宽（按内容权重分配）
"""
import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

# ---------- 版式常量 ----------
F_BODY = "宋体"          # 正文中文字体
F_HEAD = "黑体"          # 标题中文字体
F_ASCII = "Times New Roman"
C_HEAD = "1F3864"        # 标题深蓝
C_LINK = "0563C1"        # 超链接蓝
C_TEXT = "000000"
C_QUOTE = "595959"
SHADE_HEADER = "F2F2F2"  # 表头底纹
PAGE_W = 11906           # A4 宽（twips）
MARGIN_L, MARGIN_R = 1701, 1417
USABLE_W = PAGE_W - MARGIN_L - MARGIN_R
BODY_SZ, CELL_SZ, REF_SZ, SMALL_SZ = 24, 18, 21, 18  # 半磅：12/9/10.5/9pt

TOKEN_RE = re.compile(
    r'<a id="(?P<anchor>[^"]+)"></a>'
    r'|\[(?P<ltext>[^\]]+)\]\((?P<lurl>[^)\s]+)\)'
    r'|\*\*(?P<btext>[^*]+)\*\*'
    r'|`(?P<code>[^`]+)`'
)
BARE_URL_RE = re.compile(r'https?://[^\s<>"\'（）【】]+')
EMOJI_MAP = {"⚠️": "[!]", "⚠": "[!]"}


def sanitize(text):
    """替换 Word 里易变成豆腐块的字符（emoji 等）。"""
    for k, v in EMOJI_MAP.items():
        text = text.replace(k, v)
    return "".join(ch for ch in text if not (0xFE00 <= ord(ch) <= 0xFE0F or ord(ch) >= 0x1F000))


def _plain_segs(s, bold=False):
    """纯文本切段；裸 URL 自动转外部链接段。"""
    s = s.replace("\\*", "*")
    segs, pos = [], 0
    for m in BARE_URL_RE.finditer(s):
        if m.start() > pos:
            segs.append({"text": s[pos:m.start()], "bold": bold})
        url = m.group(0).rstrip(".,;、）)】。,;:")
        segs.append({"text": url, "bold": bold, "link": ("external", url)})
        tail = m.group(0)[len(url):]
        if tail:
            segs.append({"text": tail, "bold": bold})
        pos = m.end()
    if pos < len(s):
        segs.append({"text": s[pos:], "bold": bold})
    return segs


def parse_inline(s):
    """行内 Markdown → run 段列表。段: {text, bold, anchor, link}"""
    s = sanitize(s)
    segs, pos = [], 0
    for m in TOKEN_RE.finditer(s):
        if m.start() > pos:
            segs.extend(_plain_segs(s[pos:m.start()]))
        if m.group("anchor") is not None:
            segs.append({"text": "", "anchor": m.group("anchor")})
        elif m.group("lurl") is not None:
            url = m.group("lurl")
            kind = "internal" if url.startswith("#") else "external"
            segs.append({"text": m.group("ltext"), "link": (kind, url.lstrip("#"))})
        elif m.group("btext") is not None:
            segs.extend(_plain_segs(m.group("btext"), bold=True))
        elif m.group("code") is not None:
            segs.append({"text": m.group("code")})
        pos = m.end()
    if pos < len(s):
        segs.extend(_plain_segs(s[pos:]))
    return segs


def _rpr(size=None, bold=False, color=C_TEXT, east=F_BODY, ascii_f=F_ASCII):
    rPr = OxmlElement("w:rPr")
    rF = OxmlElement("w:rFonts")
    for k, v in (("w:ascii", ascii_f), ("w:hAnsi", ascii_f), ("w:eastAsia", east)):
        rF.set(qn(k), v)
    rPr.append(rF)
    if bold:
        rPr.append(OxmlElement("w:b"))
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rPr.append(c)
    if size:
        for tag in ("w:sz", "w:szCs"):
            e = OxmlElement(tag)
            e.set(qn("w:val"), str(size))
            rPr.append(e)
    return rPr


def _link_run(text, target, internal, size=None):
    """构造 w:hyperlink 元素；internal 时 target 为书签名，否则为关系 rId。"""
    h = OxmlElement("w:hyperlink")
    if internal:
        h.set(qn("w:anchor"), target.replace("-", "_"))
    else:
        h.set(qn("r:id"), target)
    r = OxmlElement("w:r")
    r.append(_rpr(size, color=C_LINK))
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    h.append(r)
    return h


def fill_paragraph(p, segs, size=None, east=F_BODY, bold_all=False):
    """把 run 段写入段落，返回段落内的书签名列表。

    size=None 时字号继承段落样式；bold_all=True 时所有普通 run 加粗（表头用）。
    """
    bookmarks = []
    for seg in segs:
        if seg.get("anchor"):
            bookmarks.append(seg["anchor"].replace("-", "_"))
            continue
        if seg.get("link"):
            kind, target = seg["link"]
            internal = kind == "internal"
            if not internal:
                target = p.part.relate_to(target, RT.HYPERLINK, is_external=True)
            p._p.append(_link_run(seg["text"], target, internal, size=size))
            continue
        run = p.add_run(seg["text"])
        if size:
            run.font.size = Pt(size / 2)
        if bold_all or seg.get("bold"):
            run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(C_TEXT)
        run.font.name = F_ASCII
        run._element.rPr.rFonts.set(qn("w:eastAsia"), east)
    return bookmarks


def add_bookmark(p, name, bm_id):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bm_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bm_id))
    p._p.get_or_add_pPr().addnext(start)
    p._p.append(end)


def first_line_indent_2char(p):
    """正文首行缩进 2 字符（firstLineChars=200，兼容旧版同时写 480 twips）。"""
    ind = OxmlElement("w:ind")
    ind.set(qn("w:firstLine"), "480")
    ind.set(qn("w:firstLineChars"), "200")
    p._p.get_or_add_pPr().append(ind)


def hanging_indent(p, twips=480):
    pf = p.paragraph_format
    pf.left_indent = Twips(twips)
    pf.first_line_indent = Twips(-twips)


def bottom_border(p, color="D9D9D9"):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single")
    b.set(qn("w:sz"), "6")
    b.set(qn("w:color"), color)
    pBdr.append(b)
    pPr.append(pBdr)


def style_east(st, east):
    st.element.rPr.rFonts.set(qn("w:eastAsia"), east)


def setup_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = F_ASCII
    normal.font.size = Pt(BODY_SZ / 2)
    normal.font.color.rgb = RGBColor.from_string(C_TEXT)
    style_east(normal, F_BODY)
    normal.paragraph_format.line_spacing = 1.3
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    title = doc.styles["Title"]
    title.font.name = F_ASCII
    title.font.size = Pt(20)
    title.font.bold = True
    title.font.color.rgb = RGBColor.from_string(C_HEAD)
    style_east(title, F_HEAD)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.line_spacing = 1.3
    title.paragraph_format.space_after = Pt(10)

    for name, size, before, after in (
        ("Heading 1", 15, 14, 6),
        ("Heading 2", 13, 10, 4),
        ("Heading 3", 12, 8, 4),
    ):
        st = doc.styles[name]
        st.font.name = F_ASCII
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor.from_string(C_HEAD)
        style_east(st, F_HEAD)
        st.paragraph_format.line_spacing = 1.3
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)


def build_footer(doc):
    sec = doc.sections[0]
    sec.top_margin = Twips(1440)
    sec.bottom_margin = Twips(1440)
    sec.left_margin = Twips(MARGIN_L)
    sec.right_margin = Twips(MARGIN_R)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = fp.add_run("第 ")
    r1.font.size = Pt(SMALL_SZ / 2)
    r1.font.name = F_ASCII
    r1._element.rPr.rFonts.set(qn("w:eastAsia"), F_BODY)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE \\* MERGEFORMAT ")
    fr = OxmlElement("w:r")
    fr.append(_rpr(SMALL_SZ, color="666666"))
    ft = OxmlElement("w:t")
    ft.text = "1"
    fr.append(ft)
    fld.append(fr)
    fp._p.append(fld)
    r2 = fp.add_run(" 页")
    r2.font.size = Pt(SMALL_SZ / 2)
    r2.font.name = F_ASCII
    r2._element.rPr.rFonts.set(qn("w:eastAsia"), F_BODY)


def cell_plain(s):
    """表格单元格纯文本（用于列宽权重估算）。"""
    s = re.sub(r'<a id="[^"]+"></a>', "", s)
    s = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", s)
    return s.replace("**", "").replace("`", "")


def vislen(s):
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in s)


def build_table(doc, rows):
    header, data = rows[0], rows[2:]
    ncols = len(header)
    weights = [
        max(vislen(cell_plain(r[c])) for r in rows if c < len(r)) for c in range(ncols)
    ]
    weights = [max(w, 4) for w in weights]
    fracs = [max(w / sum(weights), 0.08) for w in weights]
    fracs = [f / sum(fracs) for f in fracs]
    widths = [int(f * USABLE_W) for f in fracs]

    t = doc.add_table(rows=len(data) + 1, cols=ncols)
    t.style = doc.styles["Table Grid"]
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    tblPr = t._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    mar = OxmlElement("w:tblCellMar")
    for side, v in (("top", 60), ("left", 100), ("bottom", 60), ("right", 100)):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(v))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    look = tblPr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(lay)
        look.addprevious(mar)
    else:
        tblPr.append(lay)
        tblPr.append(mar)

    for ri, row in enumerate([header] + data):
        tr = t.rows[ri]
        trPr = tr._tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr._tr.insert(0, trPr)
        trPr.append(OxmlElement("w:cantSplit"))
        if ri == 0:
            trPr.append(OxmlElement("w:tblHeader"))
        for ci in range(ncols):
            cell = tr.cells[ci]
            cell.width = Twips(widths[ci])
            if ri == 0:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), SHADE_HEADER)
                tcPr.append(shd)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.3
            fill_paragraph(
                p,
                parse_inline(row[ci] if ci < len(row) else ""),
                size=CELL_SZ,
                bold_all=(ri == 0),
            )
    return len(data) + 1


def convert(md_path, out_path):
    lines = Path(md_path).read_text(encoding="utf-8").splitlines()
    doc = Document()
    setup_styles(doc)
    build_footer(doc)

    stats = {"paras": 0, "tables": 0, "table_rows": 0, "bookmarks": 0}
    bm_id = 100
    i, title_done = 0, False
    in_ref_zone = False  # 进入"参考来源"章节后条目字号缩小
    n = len(lines)

    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        # 表格块（首行 | 开头，第二行为分隔行）
        if s.startswith("|") and i + 1 < n and re.match(r"^\|[\s:\-|]+\|$", lines[i + 1].strip()):
            block = []
            while i < n and lines[i].strip().startswith("|"):
                block.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            stats["tables"] += 1
            stats["table_rows"] += build_table(doc, block)
            continue
        # 标题
        m = re.match(r"^(#{1,4})\s+(.*)$", s)
        if m:
            level, text = len(m.group(1)), m.group(2)
            if level == 1 and not title_done:
                p = doc.add_paragraph(style="Title")
                fill_paragraph(p, parse_inline(text), east=F_HEAD)
                doc.core_properties.title = text.replace("**", "")
                doc.core_properties.author = "Z.ai"
                doc.core_properties.subject = "市场调研报告"
                title_done = True
                i += 1
                continue
            in_ref_zone = "参考来源" in text
            p = doc.add_paragraph(style={2: "Heading 1", 3: "Heading 2", 4: "Heading 3"}[min(level, 4)])
            fill_paragraph(p, parse_inline(text), east=F_HEAD)
            for run in p.runs:
                run.font.color.rgb = RGBColor.from_string(C_HEAD)
                run.font.bold = True
            stats["paras"] += 1
            i += 1
            continue
        # 分隔线
        if re.match(r"^-{3,}$", s):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(8)
            bottom_border(p)
            i += 1
            continue
        # 引用块
        if s.startswith(">"):
            buf = []
            while i < n and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip("> ").strip())
                i += 1
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Twips(360)
            fill_paragraph(p, parse_inline(" ".join(buf)), size=SMALL_SZ + 3)
            for run in p.runs:
                run.font.color.rgb = RGBColor.from_string(C_QUOTE)
            stats["paras"] += 1
            continue
        # 无序列表
        if re.match(r"^[-*]\s+", s):
            p = doc.add_paragraph()
            hanging_indent(p)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            fill_paragraph(p, parse_inline("• " + re.sub(r"^[-*]\s+", "", s)))
            stats["paras"] += 1
            i += 1
            continue
        # 有序列表（含参考来源条目，写入书签）
        m = re.match(r"^(\d+)\.\s+(.*)$", s)
        if m:
            p = doc.add_paragraph()
            hanging_indent(p)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            size = REF_SZ if in_ref_zone else BODY_SZ
            bms = fill_paragraph(p, parse_inline(f"{m.group(1)}. {m.group(2)}"), size=size)
            for name in bms:
                bm_id += 1
                add_bookmark(p, name, bm_id)
                stats["bookmarks"] += 1
            stats["paras"] += 1
            i += 1
            continue
        # 普通段落（连续非空行合并）
        buf = [s]
        i += 1
        while i < n:
            nxt = lines[i].strip()
            if (
                not nxt
                or nxt.startswith(("#", "|", ">"))
                or re.match(r"^[-*]\s+", nxt)
                or re.match(r"^\d+\.\s+", nxt)
                or re.match(r"^-{3,}$", nxt)
            ):
                break
            buf.append(nxt)
            i += 1
        p = doc.add_paragraph()
        first_line_indent_2char(p)
        fill_paragraph(p, parse_inline("".join(buf)))
        stats["paras"] += 1

    doc.save(out_path)
    print(
        f"OK: {out_path}\n"
        f"  段落 {stats['paras']} | 表格 {stats['tables']}（{stats['table_rows']} 行）"
        f" | 书签 {stats['bookmarks']}"
    )


def main():
    ap = argparse.ArgumentParser(description="市场调研报告 md → docx")
    ap.add_argument("input", help="输入 .md 文件")
    ap.add_argument("-o", "--output", help="输出 .docx 路径（默认同名）")
    args = ap.parse_args()
    src = Path(args.input)
    if not src.exists():
        sys.exit(f"输入文件不存在: {src}")
    out = Path(args.output) if args.output else src.with_suffix(".docx")
    convert(src, out)


if __name__ == "__main__":
    main()
