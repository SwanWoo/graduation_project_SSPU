#!/usr/bin/env python3
"""
将 Markdown 内容注入到 Word 毕业论文模板中。
策略：保留原始 document.xml 的完整结构（封面、声明、摘要、目录、分节符、页眉页脚等），
只替换正文区域的段落内容（从 P71 开始到最后 sectPr 之前），用原始段落的 XML 格式作为模版。
"""

import re
import copy
import xml.etree.ElementTree as ET
import os
import zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def unescape_markdown(text):
    """去除 Markdown 转义符。"""
    text = text.replace(r'\*\*', '**')
    text = text.replace(r'\*', '*')
    text = text.replace(r'\-\>', '->')
    text = text.replace(r'\-', '-')
    text = text.replace(r'\(', '(')
    text = text.replace(r'\)', ')')
    text = text.replace(r'\#', '#')
    text = text.replace(r'\_', '_')
    text = text.replace(r'\\', '\\')
    return text


def parse_markdown(md_text):
    """解析 Markdown 文本为结构化块列表。"""
    lines = md_text.split('\n')
    blocks = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.strip() == '':
            i += 1
            continue
        if re.match(r'^---+\s*$', line.strip()):
            blocks.append({'type': 'hr'})
            i += 1
            continue
        heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if heading_match:
            level = len(heading_match.group(1))
            text = unescape_markdown(heading_match.group(2).strip())
            blocks.append({'type': 'heading', 'level': level, 'text': text})
            i += 1
            continue
        bullet_match = re.match(r'^(\s*)[-*]\s+(.+)$', line)
        if bullet_match:
            indent = len(bullet_match.group(1))
            text = unescape_markdown(bullet_match.group(2).strip())
            blocks.append({'type': 'bullet', 'text': text, 'indent': indent})
            i += 1
            continue
        num_match = re.match(r'^(\s*)\d+\.\s+(.+)$', line)
        if num_match:
            indent = len(num_match.group(1))
            text = unescape_markdown(num_match.group(2).strip())
            blocks.append({'type': 'numbered', 'text': text, 'indent': indent})
            i += 1
            continue
        # Regular paragraph
        para_lines = [line]
        i += 1
        while i < len(lines):
            next_line = lines[i]
            if next_line.strip() == '':
                i += 1
                break
            if (re.match(r'^#{1,6}\s', next_line) or
                re.match(r'^[-*]\s', next_line) or
                re.match(r'^\d+\.\s', next_line) or
                re.match(r'^---+\s*$', next_line.strip())):
                break
            para_lines.append(next_line)
            i += 1
        blocks.append({'type': 'paragraph', 'text': unescape_markdown(' '.join(para_lines))})
    return blocks


def parse_inline(text):
    """解析行内格式，返回 run 列表。"""
    runs = []
    pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)'
    last_end = 0
    for m in re.finditer(pattern, text):
        if m.start() > last_end:
            runs.append({'text': text[last_end:m.start()], 'bold': False, 'italic': False})
        if m.group(2):
            runs.append({'text': m.group(2), 'bold': True, 'italic': True})
        elif m.group(3):
            runs.append({'text': m.group(3), 'bold': True, 'italic': False})
        elif m.group(4):
            runs.append({'text': m.group(4), 'bold': False, 'italic': True})
        last_end = m.end()
    if last_end < len(text):
        runs.append({'text': text[last_end:], 'bold': False, 'italic': False})
    return runs


def make_run_elem(text, bold=False, italic=False):
    """创建 w:r 元素（沿用模板的命名空间）。"""
    r = ET.SubElement(ET.Element('dummy'), f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    if bold:
        ET.SubElement(rpr, f'{{{W}}}b')
    if italic:
        ET.SubElement(rpr, f'{{{W}}}i')
        ET.SubElement(rpr, f'{{{W}}}iCs')
    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def register_namespaces():
    """注册所有需要的命名空间。"""
    namespaces = {
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'o': 'urn:schemas-microsoft-com:office:office',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'v': 'urn:schemas-microsoft-com:vml',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    ET.register_namespace('', W)


def clone_paragraph_style(source_p):
    """
    从源段落克隆 pPr（段落属性），创建一个新的空段落。
    保留源段落的样式、编号、缩进等格式。
    """
    new_p = ET.Element(f'{{{W}}}p')
    source_ppr = source_p.find(f'{{{W}}}pPr')
    if source_ppr is not None:
        new_ppr = copy.deepcopy(source_ppr)
        # 移除 sectPr（如果有），它属于分节标记不属于段落样式
        sect = new_ppr.find(f'{{{W}}}sectPr')
        if sect is not None:
            new_ppr.remove(sect)
        new_p.append(new_ppr)
    return new_p


def set_paragraph_style(new_p, style_id, num_ilvl=None, num_id=None):
    """设置段落的 pStyle 和 numPr。"""
    ppr = new_p.find(f'{{{W}}}pPr')
    if ppr is None:
        ppr = ET.SubElement(new_p, f'{{{W}}}pPr')

    # 设置 pStyle
    existing_style = ppr.find(f'{{{W}}}pStyle')
    if existing_style is not None:
        ppr.remove(existing_style)
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set('val', style_id)

    # 设置 numPr
    if num_ilvl is not None:
        existing_num = ppr.find(f'{{{W}}}numPr')
        if existing_num is not None:
            ppr.remove(existing_num)
        numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
        ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
        ilvl.set('val', str(num_ilvl))
        numid = ET.SubElement(numpr, f'{{{W}}}numId')
        numid.set('val', str(num_id if num_id else 1))
    else:
        existing_num = ppr.find(f'{{{W}}}numPr')
        if existing_num is not None:
            ppr.remove(existing_num)


def set_paragraph_indent(new_p, first_line_chars=None, first_line=None, left=None):
    """设置段落缩进。"""
    ppr = new_p.find(f'{{{W}}}pPr')
    if ppr is None:
        ppr = ET.SubElement(new_p, f'{{{W}}}pPr')
    existing_ind = ppr.find(f'{{{W}}}ind')
    if existing_ind is not None:
        ppr.remove(existing_ind)
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    if first_line_chars is not None:
        ind.set(f'{{{W}}}firstLineChars', str(first_line_chars))
    if first_line is not None:
        ind.set(f'{{{W}}}firstLine', str(first_line))
    if left is not None:
        ind.set(f'{{{W}}}left', str(left))


def make_text_run(text, bold=False, italic=False, font_ascii=None, font_east_asia=None, hint=None):
    """创建一个文本 run，使用模板的字体风格。"""
    r = ET.Element(f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')

    has_fonts = False
    if font_ascii or font_east_asia or hint:
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        has_fonts = True
        if font_ascii:
            fonts.set(f'{{{W}}}ascii', font_ascii)
            fonts.set(f'{{{W}}}hAnsi', font_ascii)
        if font_east_asia:
            fonts.set(f'{{{W}}}eastAsia', font_east_asia)
        if hint:
            fonts.set(f'{{{W}}}hint', hint)

    if bold:
        b = ET.SubElement(rpr, f'{{{W}}}b')
    if italic:
        it = ET.SubElement(rpr, f'{{{W}}}i')
        itCs = ET.SubElement(rpr, f'{{{W}}}iCs')

    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def add_text_to_paragraph(p, text, bold=False, italic=False):
    """将文本添加到段落中，解析行内格式。"""
    inline_runs = parse_inline(text)
    for run_info in inline_runs:
        if not run_info['text']:
            continue
        r = make_text_run(
            run_info['text'],
            bold=run_info['bold'] or bold,
            italic=run_info['italic'] or italic,
            font_ascii='Times New Roman',
            hint='eastAsia',
        )
        p.append(r)


_bookmark_counter = 0

def next_bookmark_name():
    global _bookmark_counter
    _bookmark_counter += 1
    return f'_Toc{999999999 + _bookmark_counter}'


def build_heading(block, body_children, bookmark_name=None):
    """
    根据模板格式构建标题段落。
    H1 -> style '11' (1级), 居中, 三号黑体, numPr ilvl=0 numId=1
    H2 -> style 'ad' (二级), ilvl=1 numId=1
    H3 -> style 'a' (三级), 自动编号
    如果提供 bookmark_name，在段落首尾插入 bookmarkStart/bookmarkEnd。
    """
    level = block['level']
    text = block['text']

    if level == 1:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set('val', '11')
        numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
        ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
        ilvl.set('val', '0')
        numid = ET.SubElement(numpr, f'{{{W}}}numId')
        numid.set('val', '1')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set('val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set('val', '0')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}left', '0')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

    elif level == 2:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set('val', 'ad')
        numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
        ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
        ilvl.set('val', '1')
        numid = ET.SubElement(numpr, f'{{{W}}}numId')
        numid.set('val', '1')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set('val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set('val', '0')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}left', '0')
        ind.set(f'{{{W}}}firstLine', '0')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

    else:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set('val', 'a')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set('val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set('val', '0')
        spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
        spacing.set(f'{{{W}}}line', '360')
        spacing.set(f'{{{W}}}lineRule', 'auto')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}left', '0')
        ind.set(f'{{{W}}}firstLine', '0')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

    add_text_to_paragraph(p, text, bold=False)

    # 插入 bookmark
    if bookmark_name:
        bm_start = ET.SubElement(p, f'{{{W}}}bookmarkStart')
        bm_start.set(f'{{{W}}}id', '0')
        bm_start.set(f'{{{W}}}name', bookmark_name)
        bm_end = ET.SubElement(p, f'{{{W}}}bookmarkEnd')
        bm_end.set(f'{{{W}}}id', '0')

    return p


def build_body_paragraph(block):
    """
    构建正文段落。
    参考 P74/P100: style=aa, 首行缩进 firstLineChars=200 firstLine=480
    """
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set('val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set('val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set('val', '0')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLineChars', '200')
    ind.set(f'{{{W}}}firstLine', '480')
    add_text_to_paragraph(p, text)
    return p


def build_bullet_paragraph(block):
    """
    构建列表项段落。
    参考 P74: style=aa, 用 bullet numPr
    """
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set('val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set('val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set('val', '0')
    numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
    ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
    ilvl.set('val', '0')
    numid = ET.SubElement(numpr, f'{{{W}}}numId')
    numid.set('val', '2')  # bullet numbering
    add_text_to_paragraph(p, text)
    return p


def build_numbered_paragraph(block):
    """
    构建编号列表段落。
    参考 numbering abstractNumId=5 (numId=1): decimal numbering
    """
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set('val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set('val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set('val', '0')
    numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
    ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
    ilvl.set('val', '0')
    numid = ET.SubElement(numpr, f'{{{W}}}numId')
    numid.set('val', '1')  # decimal numbering
    add_text_to_paragraph(p, text)
    return p


def build_hr_paragraph(body_children):
    """
    构建分隔线段落。参考 P72（空行居中段落）。
    """
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set('val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set('val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set('val', '0')
    # 添加底部边框作为分隔线
    pBdr = ET.SubElement(ppr, f'{{{W}}}pBdr')
    bottom = ET.SubElement(pBdr, f'{{{W}}}bottom')
    bottom.set('val', 'single')
    bottom.set('sz', '6')
    bottom.set('space', '1')
    bottom.set('color', 'auto')
    return p


def build_toc_entry(number, title, toc_level, bookmark_name='_Toc1000000000'):
    """
    构建目录条目段落。
    toc_level: 1, 2, or 3
    模板格式参考：
      Level 1 (P47): style=12, 有 tab pos=630 和 right tab pos=8302 leader=dot
      Level 2 (P48): style=20, tab pos=1260, ind firstLine=480
      Level 3 (P56): style=30, tab pos=1710, ind firstLine=960
    每个条目结构：编号r + tab + 标题r + tab + PAGEREF field (begin, instrText, separate, 页码t, end)
    """
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')

    # 根据级别设置 style、tabs、ind
    if toc_level == 1:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set('val', '12')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab1 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab1.set('val', 'left')
        tab1.set('pos', '630')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set('val', 'right')
        tab2.set('leader', 'dot')
        tab2.set('pos', '8302')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set('val', '21')
        szCs = ET.SubElement(rpr_in_ppr, f'{{{W}}}szCs')
        szCs.set('val', '22')
    elif toc_level == 2:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set('val', '20')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab1 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab1.set('val', 'left')
        tab1.set('pos', '1260')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set('val', 'right')
        tab2.set('leader', 'dot')
        tab2.set('pos', '8302')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLine', '480')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        kern = ET.SubElement(rpr_in_ppr, f'{{{W}}}kern')
        kern.set('val', '2')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set('val', '21')
    else:  # level 3
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set('val', '30')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab1 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab1.set('val', 'left')
        tab1.set('pos', '1710')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set('val', 'right')
        tab2.set('leader', 'dot')
        tab2.set('pos', '8302')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLine', '960')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        kern = ET.SubElement(rpr_in_ppr, f'{{{W}}}kern')
        kern.set('val', '2')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set('val', '21')

    # 构建 run 的辅助函数
    def _r(text_content=None, rpr_extra=None):
        """创建一个 run。"""
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
        no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
        if text_content is not None:
            t = ET.SubElement(r, f'{{{W}}}t')
            t.text = text_content
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        return r

    def _r_ea(text_content):
        """创建一个 eastAsia hint run。"""
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
        no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
        if toc_level >= 2:
            kern = ET.SubElement(rpr, f'{{{W}}}kern')
            kern.set('val', '2')
            sz = ET.SubElement(rpr, f'{{{W}}}sz')
            sz.set('val', '21')
        # tab
        tab = ET.SubElement(r, f'{{{W}}}tab')
        return r

    # 编号 run
    if number:
        _r(number)

    # tab (级别2/3在 ea run 里，级别1单独)
    if toc_level == 1:
        tab_r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(tab_r, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        if toc_level == 1:
            b = ET.SubElement(rpr, f'{{{W}}}b')
            szCs = ET.SubElement(rpr, f'{{{W}}}szCs')
            szCs.set('val', '28')
        no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
        ET.SubElement(tab_r, f'{{{W}}}tab')
    else:
        _r_ea(None)

    # 标题 run
    _r(title)

    # tab before page number (triggers dot leader from right tab in pPr)
    tab_run = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(tab_run, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    ET.SubElement(tab_run, f'{{{W}}}tab')

    # PAGEREF field: begin
    r_begin = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_begin, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    fld_begin = ET.SubElement(r_begin, f'{{{W}}}fldChar')
    fld_begin.set(f'{{{W}}}fldCharType', 'begin')

    # PAGEREF instrText
    r_instr = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_instr, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    instr = ET.SubElement(r_instr, f'{{{W}}}instrText')
    instr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    instr.text = f' PAGEREF {bookmark_name} \\h '

    # empty run between instrText and separate
    ET.SubElement(p, f'{{{W}}}r')

    # PAGEREF field: separate
    r_sep = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_sep, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    fld_sep = ET.SubElement(r_sep, f'{{{W}}}fldChar')
    fld_sep.set(f'{{{W}}}fldCharType', 'separate')

    # 页码显示（占位，Word 打开后更新域会自动更新）
    r_page = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_page, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    t = ET.SubElement(r_page, f'{{{W}}}t')
    t.text = '1'

    # PAGEREF field: end
    r_end = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r_end, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    fld_end = ET.SubElement(r_end, f'{{{W}}}fldChar')
    fld_end.set(f'{{{W}}}fldCharType', 'end')

    return p


def strip_markdown_formatting(text):
    """去除 Markdown 格式标记（加粗、斜体等），只保留纯文本。"""
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', text)  # bold+italic
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)       # bold
    text = re.sub(r'\*(.+?)\*', r'\1', text)            # italic
    text = re.sub(r'`(.+?)`', r'\1', text)              # inline code
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)     # links
    return text.strip()


def build_toc_entries(blocks):
    """
    从 Markdown blocks 中提取标题，生成目录条目列表。
    使用多级编号计数器跟踪 H1/H2/H3 编号。
    标题文本去除 Markdown 格式标记。
    为每个标题分配唯一 bookmark 名称。
    """
    entries = []
    h1_count = 0
    h2_count = 0
    h3_count = 0

    for block in blocks:
        if block['type'] != 'heading':
            continue
        level = block['level']
        if level > 3:
            level = 3  # 最多三级目录

        if level == 1:
            h1_count += 1
            h2_count = 0
            h3_count = 0
            number = str(h1_count)
            toc_level = 1
        elif level == 2:
            h2_count += 1
            h3_count = 0
            number = f'{h1_count}.{h2_count}'
            toc_level = 2
        else:
            h3_count += 1
            number = f'{h1_count}.{h2_count}.{h3_count}'
            toc_level = 3

        clean_title = strip_markdown_formatting(block['text'])
        bookmark = next_bookmark_name()
        block['_bookmark'] = bookmark
        entries.append({
            'number': number,
            'title': clean_title,
            'toc_level': toc_level,
            'bookmark': bookmark,
        })

    return entries


def main():
    register_namespaces()

    template_dir = '/Users/swan/rootos/10_Action/Project/毕设'
    orig_xml_path = os.path.join(template_dir, 'word', 'document.xml')

    # Read markdown
    with open(os.path.join(template_dir, '作文模版.md'), 'r', encoding='utf-8') as f:
        md_text = f.read()

    # Parse markdown
    blocks = parse_markdown(md_text)

    # Build TOC entries from headings
    toc_entries = build_toc_entries(blocks)

    # Parse original document.xml
    tree = ET.parse(orig_xml_path)
    root = tree.getroot()
    body = root.find(f'{{{W}}}body')
    original_children = list(body)

    # === Phase 1: Identify key positions in original document ===
    # P0-P13: Cover
    # P14: Section break (sectPr in pPr) - end of cover
    # P15-P28: Declaration
    # P29: empty
    # P30: Section break (sectPr in pPr) - end of declaration
    # P31-P44: Chinese & English abstract
    # P45: Section break (sectPr in pPr) - end of abstract
    # P46: TOC title
    # P47-P60: TOC entries (14 items)
    # P61: TOC field end (fldChar end)
    # P62: "目录末尾插入分节符..."
    # P63-P69: empty paragraphs
    # P70: Section break (sectPr in pPr) - end of TOC, start of body
    # P71-P169: Body content
    # P170 (last_sectPr): Final sectPr

    # Find the final sectPr
    last_sectPr = body.find(f'{{{W}}}sectPr')

    # Find all inline sectPr paragraphs by index
    sect_break_indices = []
    for i, child in enumerate(original_children):
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == 'p':
            ppr = child.find(f'{{{W}}}pPr')
            if ppr is not None and ppr.find(f'{{{W}}}sectPr') is not None:
                sect_break_indices.append(i)

    # sect_break_indices should be [14, 30, 45, 70]
    # P70 is the TOC->body section break

    # === Phase 2: Replace TOC entries ===
    # Remove original P47-P60 (14 TOC entries)
    toc_remove = [original_children[i] for i in range(47, 61)]
    for elem in toc_remove:
        body.remove(elem)

    # Re-read current children to get correct insertion point
    current_children = list(body)

    # Find P46 (TOC title) in current tree - it's the element that was originally at index 46
    toc_title = original_children[46]
    ref_idx = current_children.index(toc_title)

    # Insert new TOC entries after the TOC title
    for j, entry in enumerate(toc_entries):
        toc_p = build_toc_entry(entry['number'], entry['title'], entry['toc_level'], entry['bookmark'])
        body.insert(ref_idx + 1 + j, toc_p)

    # === Phase 3: Replace body content ===
    # Re-read current children
    current_children = list(body)

    # Find P70 (the TOC->body section break) in current tree
    body_sect_break = original_children[70]
    body_sect_idx = current_children.index(body_sect_break)

    # Remove everything after body_sect_break (but not last_sectPr)
    to_remove = []
    for child in list(body):
        if child is last_sectPr:
            continue
        idx = list(body).index(child)
        if idx > body_sect_idx:
            to_remove.append(child)

    for elem in to_remove:
        body.remove(elem)

    # Append new body paragraphs
    for block in blocks:
        btype = block['type']
        if btype == 'heading':
            p = build_heading(block, original_children, bookmark_name=block.get('_bookmark'))
        elif btype == 'paragraph':
            p = build_body_paragraph(block)
        elif btype == 'bullet':
            p = build_bullet_paragraph(block)
        elif btype == 'numbered':
            p = build_numbered_paragraph(block)
        elif btype == 'hr':
            p = build_hr_paragraph(original_children)
        else:
            continue
        body.append(p)

    # Ensure last_sectPr is at the very end
    if last_sectPr is not None:
        if last_sectPr in list(body):
            body.remove(last_sectPr)
        body.append(last_sectPr)

    # Write new document.xml
    output_xml = os.path.join(template_dir, 'word', 'document_new.xml')
    tree.write(output_xml, encoding='UTF-8', xml_declaration=True)

    # Create docx
    output_docx = os.path.join(template_dir, '六级作文高分模版.docx')
    skip_extensions = {'.md', '.py', '.docx'}

    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root_dir, dirs, files in os.walk(template_dir):
            dirs[:] = [d for d in dirs if d not in ('.git', '.claude')]
            for file in files:
                if file.startswith('.DS_Store'):
                    continue
                if any(file.endswith(ext) for ext in skip_extensions):
                    continue
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, template_dir)
                if arcname == 'word/document.xml':
                    zf.write(output_xml, 'word/document.xml')
                elif arcname == 'word/document_new.xml':
                    continue
                else:
                    zf.write(file_path, arcname)

    # Clean up temp file
    os.remove(output_xml)

    print(f'Generated: {output_docx}')
    print(f'Converted {len(blocks)} blocks from markdown')
    print()
    print('Preserved template structure:')
    print('  - Cover page (P0-P13)')
    print('  - Section break (P14)')
    print('  - Declaration (P15-P28)')
    print('  - Section break (P30)')
    print('  - Chinese abstract (P31-P40)')
    print('  - English abstract (P41-P44)')
    print('  - Section break (P45)')
    print('  - Table of Contents (P46-P62)')
    print('  - Section break (P70)')
    print('  - [REPLACED] Body content with markdown')
    print('  - Final sectPr preserved')


if __name__ == '__main__':
    main()
