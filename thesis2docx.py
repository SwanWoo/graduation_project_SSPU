#!/usr/bin/env python3
"""
将 final_paper.md 毕业论文 Markdown 转换为 Word 文档。
基于 md2docx.py 架构，扩展支持：封面注入、表格、图片、公式、代码块等。
模板结构（段落索引 P0-P159）：
  P0-P13:  封面页
  P14:     分节符（封面→声明）
  P15-P28: 声明页
  P30:     分节符（声明→摘要）
  P31:     中文标题
  P32:     "摘要"标签
  P33-P35: 中文摘要正文
  P36:     中文关键词
  P37:     关键词格式说明（需清空）
  P38:     分页符（中→英摘要）
  P39:     英文标题
  P40:     "ABSTRACT"标签
  P41:     英文摘要正文
  P42:     英文关键词
  P43:     分节符（摘要→目录）
  P44:     目录标题
  P45-P58: 目录条目
  P68:     分节符（目录→正文）
  P69+:    正文内容
"""

import re
import copy
import io
import os
import sys
import zipfile
import urllib.request
import tempfile
import shutil
import subprocess
import xml.etree.ElementTree as ET
from latex2mathml import converter

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
XML_NS = '{http://www.w3.org/XML/1998/namespace}'
ML_NS = '{http://www.w3.org/1998/Math/MathML}'

UPRIGHT_FUNCTIONS = {
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp', 'lim', 'min', 'max',
    'sup', 'inf', 'det', 'dim', 'mod', 'gcd', 'deg', 'arg', 'hom',
    'ker', 'Im', 'Re', 'Pr',
}


# ============================================================
# 1. 命名空间注册
# ============================================================

def register_namespaces():
    """注册所有需要的命名空间。注意：同一 URI 只保留最后一个注册的前缀。"""
    namespaces = {
        'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'o': 'urn:schemas-microsoft-com:office:office',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'v': 'urn:schemas-microsoft-com:vml',
        'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'w10': 'urn:schemas-microsoft-com:office:word',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
        'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
        'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
        'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    }
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    # r: 前缀用于关系引用（图片 embed 等），必须在最后注册
    ET.register_namespace('r', R_NS)
    # w: 作为主命名空间前缀（而非默认命名空间，避免与 w:document 根元素冲突）
    ET.register_namespace('w', W)


# ============================================================
# 1.5 LaTeX → OMML 转换（行内公式）
# ============================================================

def _mathml_children(elem):
    """获取 MathML 元素的直接子元素列表（跳过注释等非元素节点）。"""
    return [c for c in elem if isinstance(c.tag, str)]


def _convert_mathml_node(src, dest_parent):
    """递归地将 MathML 元素转换为 OMML 元素，追加到 dest_parent。"""
    tag = src.tag
    local = tag.replace(ML_NS, '') if tag.startswith('{') else tag

    if local == 'math' or local == 'mrow':
        for child in _mathml_children(src):
            _convert_mathml_node(child, dest_parent)
        return

    if local in ('mi', 'mn', 'mo', 'mtext'):
        r = ET.SubElement(dest_parent, f'{{{M_NS}}}r')
        text = (src.text or '').strip() or src.text or ''
        t = ET.SubElement(r, f'{{{M_NS}}}t')
        t.text = text
        t.set(f'{XML_NS}space', 'preserve')
        # 函数名 upright 样式
        if local == 'mi' and text in UPRIGHT_FUNCTIONS:
            rpr = ET.SubElement(r, f'{{{M_NS}}}rPr')
            sty = ET.SubElement(rpr, f'{{{M_NS}}}sty')
            sty.set(f'{{{M_NS}}}val', 'p')
        return

    if local == 'mfrac':
        f = ET.SubElement(dest_parent, f'{{{M_NS}}}f')
        children = _mathml_children(src)
        if len(children) >= 2:
            num = ET.SubElement(f, f'{{{M_NS}}}num')
            _convert_mathml_node(children[0], num)
            den = ET.SubElement(f, f'{{{M_NS}}}den')
            _convert_mathml_node(children[1], den)
        return

    if local == 'msup':
        s = ET.SubElement(dest_parent, f'{{{M_NS}}}sSup')
        children = _mathml_children(src)
        if len(children) >= 2:
            e = ET.SubElement(s, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            sup = ET.SubElement(s, f'{{{M_NS}}}sup')
            _convert_mathml_node(children[1], sup)
        return

    if local == 'msub':
        s = ET.SubElement(dest_parent, f'{{{M_NS}}}sSub')
        children = _mathml_children(src)
        if len(children) >= 2:
            e = ET.SubElement(s, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            sub = ET.SubElement(s, f'{{{M_NS}}}sub')
            _convert_mathml_node(children[1], sub)
        return

    if local == 'msubsup':
        s = ET.SubElement(dest_parent, f'{{{M_NS}}}sSubSup')
        children = _mathml_children(src)
        if len(children) >= 3:
            e = ET.SubElement(s, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            sub = ET.SubElement(s, f'{{{M_NS}}}sub')
            _convert_mathml_node(children[1], sub)
            sup = ET.SubElement(s, f'{{{M_NS}}}sup')
            _convert_mathml_node(children[2], sup)
        return

    if local == 'msqrt':
        rad = ET.SubElement(dest_parent, f'{{{M_NS}}}rad')
        deg = ET.SubElement(rad, f'{{{M_NS}}}deg')
        e = ET.SubElement(rad, f'{{{M_NS}}}e')
        for child in _mathml_children(src):
            _convert_mathml_node(child, e)
        return

    if local == 'mroot':
        rad = ET.SubElement(dest_parent, f'{{{M_NS}}}rad')
        children = _mathml_children(src)
        # mroot: [base, index] → OMML: m:e=base, m:deg=index
        deg = ET.SubElement(rad, f'{{{M_NS}}}deg')
        e = ET.SubElement(rad, f'{{{M_NS}}}e')
        if len(children) >= 1:
            _convert_mathml_node(children[0], e)
        if len(children) >= 2:
            _convert_mathml_node(children[1], deg)
        return

    if local == 'mover':
        children = _mathml_children(src)
        if len(children) >= 2:
            # 检查是否为重音符号（第二个子元素是单个字符的 mo）
            acc_char = ''
            acc_src = children[1]
            acc_tag = acc_src.tag.replace(ML_NS, '') if acc_src.tag.startswith('{') else acc_src.tag
            if acc_tag == 'mo' and acc_src.text and len(acc_src.text.strip()) <= 2:
                acc_char = acc_src.text.strip()
            if acc_char:
                acc = ET.SubElement(dest_parent, f'{{{M_NS}}}acc')
                accPr = ET.SubElement(acc, f'{{{M_NS}}}accPr')
                chr_el = ET.SubElement(accPr, f'{{{M_NS}}}chr')
                chr_el.set(f'{{{M_NS}}}val', acc_char)
                e = ET.SubElement(acc, f'{{{M_NS}}}e')
                _convert_mathml_node(children[0], e)
            else:
                limU = ET.SubElement(dest_parent, f'{{{M_NS}}}limUpp')
                e = ET.SubElement(limU, f'{{{M_NS}}}e')
                _convert_mathml_node(children[0], e)
                lim = ET.SubElement(limU, f'{{{M_NS}}}lim')
                _convert_mathml_node(children[1], lim)
        return

    if local == 'munder':
        children = _mathml_children(src)
        if len(children) >= 2:
            limL = ET.SubElement(dest_parent, f'{{{M_NS}}}limLow')
            e = ET.SubElement(limL, f'{{{M_NS}}}e')
            _convert_mathml_node(children[0], e)
            lim = ET.SubElement(limL, f'{{{M_NS}}}lim')
            _convert_mathml_node(children[1], lim)
        return

    if local == 'munderover':
        children = _mathml_children(src)
        if len(children) >= 3:
            # 尝试检测 nary (求和/积分/乘积)
            nary_chars = {'∑': '∑', '∏': '∏', '∫': '∫', '⋃': '⋃', '⋂': '⋂'}
            base_text = ''
            base_src = children[0]
            base_tag = base_src.tag.replace(ML_NS, '') if base_src.tag.startswith('{') else base_src.tag
            if base_tag == 'mo' and base_src.text:
                base_text = base_src.text.strip()
            if base_text in nary_chars:
                nary = ET.SubElement(dest_parent, f'{{{M_NS}}}nary')
                naryPr = ET.SubElement(nary, f'{{{M_NS}}}naryPr')
                chr_el = ET.SubElement(naryPr, f'{{{M_NS}}}chr')
                chr_el.set(f'{{{M_NS}}}val', nary_chars[base_text])
                limLoc = ET.SubElement(naryPr, f'{{{M_NS}}}limLoc')
                limLoc.set(f'{{{M_NS}}}val', 'subSup')
                sub = ET.SubElement(nary, f'{{{M_NS}}}sub')
                _convert_mathml_node(children[1], sub)
                sup = ET.SubElement(nary, f'{{{M_NS}}}sup')
                _convert_mathml_node(children[2], sup)
                e = ET.SubElement(nary, f'{{{M_NS}}}e')
            else:
                # 普通上下标
                nary = ET.SubElement(dest_parent, f'{{{M_NS}}}nary')
                naryPr = ET.SubElement(nary, f'{{{M_NS}}}naryPr')
                limLoc = ET.SubElement(naryPr, f'{{{M_NS}}}limLoc')
                limLoc.set(f'{{{M_NS}}}val', 'subSup')
                sub = ET.SubElement(nary, f'{{{M_NS}}}sub')
                _convert_mathml_node(children[1], sub)
                sup = ET.SubElement(nary, f'{{{M_NS}}}sup')
                _convert_mathml_node(children[2], sup)
                e = ET.SubElement(nary, f'{{{M_NS}}}e')
                _convert_mathml_node(children[0], e)
        return

    if local == 'mtable':
        m = ET.SubElement(dest_parent, f'{{{M_NS}}}m')
        for child in _mathml_children(src):
            _convert_mathml_node(child, m)
        return

    if local == 'mtr':
        mr = ET.SubElement(dest_parent, f'{{{M_NS}}}mr')
        for child in _mathml_children(src):
            _convert_mathml_node(child, mr)
        return

    if local == 'mtd':
        e = ET.SubElement(dest_parent, f'{{{M_NS}}}e')
        for child in _mathml_children(src):
            _convert_mathml_node(child, e)
        return

    if local == 'mo' and (src.get('fence') == 'true' or src.text in ('(', ')', '[', ']', '{', '}', '|', '‖')):
        # 括号元素
        _convert_mathml_node.__wrapped__(src, dest_parent) if False else None
        r = ET.SubElement(dest_parent, f'{{{M_NS}}}r')
        t = ET.SubElement(r, f'{{{M_NS}}}t')
        t.text = (src.text or '')
        t.set(f'{XML_NS}space', 'preserve')
        return

    # 未知元素：递归处理子元素
    for child in _mathml_children(src):
        _convert_mathml_node(child, dest_parent)


def latex_to_omml(latex_str):
    """将 LaTeX 行内公式转换为 OMML <m:oMath> 元素。失败返回 None。"""
    try:
        latex_str = latex_str.strip()
        if not latex_str:
            return None
        mathml_str = converter.convert(latex_str)
        mathml_elem = ET.fromstring(mathml_str)
        omath = ET.Element(f'{{{M_NS}}}oMath')
        _convert_mathml_node(mathml_elem, omath)
        return omath
    except Exception as e:
        print(f'  警告: 行内公式转换失败: {e}', file=sys.stderr)
        return None


# ============================================================
# 2. Markdown 解析
# ============================================================

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


def parse_cover_table(md_text):
    """从 Markdown 开头提取封面信息表格。"""
    # 找到 </div> 之前的表格
    match = re.search(r'\|.*?\*\*题目\*\*.*?\|(.*?)\|', md_text, re.DOTALL)
    if not match:
        return {}

    cover = {}
    # 匹配所有 | **key** | value | 行
    lines = md_text.split('\n')
    in_table = False
    for line in lines:
        line = line.strip()
        if line.startswith('|') and '**' in line:
            in_table = True
            # 提取 key 和 value
            cells = [c.strip() for c in line.split('|')]
            cells = [c for c in cells if c]  # 去空
            if len(cells) >= 2:
                key = re.sub(r'\*\*', '', cells[0]).strip()
                value = re.sub(r'\*\*', '', cells[1]).strip()
                cover[key] = value
        elif in_table and not line.startswith('|'):
            break
        # 跳过分隔行
        if re.match(r'^\|[-:|]+\|$', line):
            continue
    return cover


def parse_markdown(md_text):
    """解析 Markdown 文本为结构化块列表。扩展支持表格、代码块、公式、图片等。"""
    lines = md_text.split('\n')
    blocks = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # 跳过空行
        if line.strip() == '':
            i += 1
            continue

        # 跳过 HTML div 和居中标签（页面分隔等）
        if re.match(r'^\s*<div\s', line.strip()):
            i += 1
            continue
        if re.match(r'^\s*</div>', line.strip()):
            i += 1
            continue

        # 水平分隔线 — 直接跳过，不在 docx 中生成
        if re.match(r'^---+\s*$', line.strip()):
            i += 1
            continue

        # 标题
        heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if heading_match:
            level = len(heading_match.group(1))
            text = unescape_markdown(heading_match.group(2).strip())
            blocks.append({'type': 'heading', 'level': level, 'text': text})
            i += 1
            continue

        # 代码块
        if re.match(r'^```', line.strip()):
            lang_match = re.match(r'^```(\w*)', line.strip())
            lang = lang_match.group(1) if lang_match else ''
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # 跳过结束的 ```
            blocks.append({'type': 'codeblock', 'lang': lang, 'code': '\n'.join(code_lines)})
            continue

        # 公式 $$...$$（独占一行）
        if line.strip().startswith('$$'):
            formula_lines = [line.strip()[2:]]
            if not formula_lines[0].endswith('$$'):
                i += 1
                while i < len(lines):
                    if lines[i].strip().endswith('$$'):
                        formula_lines.append(lines[i].strip()[:-2])
                        break
                    formula_lines.append(lines[i].strip())
                    i += 1
            else:
                formula_lines[0] = formula_lines[0][:-2]
            i += 1
            blocks.append({'type': 'formula', 'latex': '\n'.join(formula_lines).strip()})
            continue

        # 表格
        if line.strip().startswith('|') and i + 1 < len(lines) and re.match(r'^\|[-:|]+\|$', lines[i + 1].strip()):
            headers = [c.strip() for c in line.strip().split('|')]
            headers = [c for c in headers if c]
            i += 2  # 跳过表头和分隔行
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().split('|')]
                cells = [c for c in cells if c]
                rows.append(cells)
                i += 1
            blocks.append({'type': 'table', 'headers': headers, 'rows': rows})
            continue

        # 引用块
        if re.match(r'^>\s*(.*)$', line):
            quote_lines = []
            while i < len(lines) and re.match(r'^>\s*(.*)$', lines[i]):
                quote_lines.append(re.sub(r'^>\s*', '', lines[i]))
                i += 1
            blocks.append({'type': 'blockquote', 'text': unescape_markdown(' '.join(quote_lines))})
            continue

        # 图片行 ![alt](url)
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', line.strip())
        if img_match:
            blocks.append({'type': 'image', 'alt': img_match.group(1), 'url': img_match.group(2).strip()})
            i += 1
            continue

        # 无序列表
        bullet_match = re.match(r'^(\s*)[-*]\s+(.+)$', line)
        if bullet_match:
            indent = len(bullet_match.group(1))
            text = unescape_markdown(bullet_match.group(2).strip())
            blocks.append({'type': 'bullet', 'text': text, 'indent': indent})
            i += 1
            continue

        # 有序列表
        num_match = re.match(r'^(\s*)\d+\.\s+(.+)$', line)
        if num_match:
            indent = len(num_match.group(1))
            text = unescape_markdown(num_match.group(2).strip())
            blocks.append({'type': 'numbered', 'text': text, 'indent': indent})
            i += 1
            continue

        # 普通段落
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
                re.match(r'^---+\s*$', next_line.strip()) or
                re.match(r'^```', next_line.strip()) or
                re.match(r'^\$\$', next_line.strip()) or
                re.match(r'^!\[', next_line.strip()) or
                re.match(r'^>\s', next_line) or
                (next_line.strip().startswith('|') and i + 1 < len(lines) and re.match(r'^\|[-:|]+\|$', lines[i + 1].strip()))):
                break
            para_lines.append(next_line)
            i += 1
        blocks.append({'type': 'paragraph', 'text': unescape_markdown(' '.join(para_lines))})

    return blocks


def parse_inline(text):
    """解析行内格式，返回 run 列表。"""
    runs = []
    pattern = r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`)'
    last_end = 0
    for m in re.finditer(pattern, text):
        if m.start() > last_end:
            runs.append({'text': text[last_end:m.start()], 'bold': False, 'italic': False, 'code': False})
        if m.group(2):
            runs.append({'text': m.group(2), 'bold': True, 'italic': True, 'code': False})
        elif m.group(3):
            runs.append({'text': m.group(3), 'bold': True, 'italic': False, 'code': False})
        elif m.group(4):
            runs.append({'text': m.group(4), 'bold': False, 'italic': True, 'code': False})
        elif m.group(5):
            runs.append({'text': m.group(5), 'bold': False, 'italic': False, 'code': True})
        last_end = m.end()
    if last_end < len(text):
        runs.append({'text': text[last_end:], 'bold': False, 'italic': False, 'code': False})
    return runs


def strip_markdown_formatting(text):
    """去除 Markdown 格式标记，只保留纯文本。"""
    text = re.sub(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', r'\1', text)
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'`(.+?)`', r'\1', text)
    text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
    return text.strip()


# ============================================================
# 3. XML 辅助函数
# ============================================================

def make_text_run(text, bold=False, italic=False, code=False, font_ascii=None, font_east_asia=None, hint=None, sz=None, unbold=False):
    """创建一个文本 run。"""
    r = ET.Element(f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')

    has_fonts = False
    if font_ascii or font_east_asia or hint or code:
        fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
        has_fonts = True
        if code:
            fonts.set(f'{{{W}}}ascii', 'Courier New')
            fonts.set(f'{{{W}}}hAnsi', 'Courier New')
        elif font_ascii:
            fonts.set(f'{{{W}}}ascii', font_ascii)
            fonts.set(f'{{{W}}}hAnsi', font_ascii)
        if font_east_asia:
            fonts.set(f'{{{W}}}eastAsia', font_east_asia)
        if hint:
            fonts.set(f'{{{W}}}hint', hint)

    if bold:
        ET.SubElement(rpr, f'{{{W}}}b')
    elif unbold:
        # 显式关闭加粗（覆盖样式继承的 bold）
        b = ET.SubElement(rpr, f'{{{W}}}b')
        b.set(f'{{{W}}}val', '0')
    if italic:
        ET.SubElement(rpr, f'{{{W}}}i')
        ET.SubElement(rpr, f'{{{W}}}iCs')

    if sz:
        s = ET.SubElement(rpr, f'{{{W}}}sz')
        s.set(f'{{{W}}}val', str(sz))
        scs = ET.SubElement(rpr, f'{{{W}}}szCs')
        scs.set(f'{{{W}}}val', str(sz))

    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set(f'{XML_NS}space', 'preserve')
    return r


def add_text_to_paragraph(p, text, bold=False, italic=False, font_ascii=None):
    """将文本添加到段落中，解析行内格式和行内公式 $...$。"""
    # 按 $...$ 分割（排除 $$），交替处理文本和公式
    math_pattern = r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)'
    last_end = 0

    for m in re.finditer(math_pattern, text):
        # 公式前的文本
        text_before = text[last_end:m.start()]
        if text_before:
            _add_text_runs(p, text_before, bold, italic, font_ascii)

        # 行内公式
        latex = m.group(1)
        omml = latex_to_omml(latex)
        if omml is not None:
            p.append(omml)
        else:
            # 回退：斜体 Cambria Math 文本
            r = make_text_run(f'${latex}$', font_ascii='Cambria Math', font_east_asia='Cambria Math', italic=True)
            p.append(r)

        last_end = m.end()

    # 最后一段文本
    remaining = text[last_end:]
    if remaining:
        _add_text_runs(p, remaining, bold, italic, font_ascii)


def _add_text_runs(p, text, bold=False, italic=False, font_ascii=None):
    """将纯文本（不含公式）按行内格式解析后添加为 runs。"""
    inline_runs = parse_inline(text)
    for run_info in inline_runs:
        if not run_info['text']:
            continue
        r = make_text_run(
            run_info['text'],
            bold=run_info['bold'] or bold,
            italic=run_info['italic'] or italic,
            code=run_info.get('code', False),
            font_ascii=run_info.get('code', False) and 'Courier New' or font_ascii or 'Times New Roman',
            hint='eastAsia',
        )
        p.append(r)


# ============================================================
# 4. 段落构建函数
# ============================================================

def build_heading(block, bookmark_name=None, bookmark_id=None):
    """构建标题段落。H1→style 11, H2→style ad, H3→style a。
    标题文本已包含编号（如 "1 绪论"），不再使用 Word 自动编号避免重复。"""
    level = block['level']
    text = block['text']

    if level == 1:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', '11')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
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
        pstyle.set(f'{{{W}}}val', 'ad')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
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
        pstyle.set(f'{{{W}}}val', 'a')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
        spacing.set(f'{{{W}}}line', '360')
        spacing.set(f'{{{W}}}lineRule', 'auto')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}left', '0')
        ind.set(f'{{{W}}}firstLine', '0')
        # 样式 'a' 自带 numId=1 自动编号，显式取消避免编号重复
        numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
        numid = ET.SubElement(numpr, f'{{{W}}}numId')
        numid.set(f'{{{W}}}val', '0')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')

    add_text_to_paragraph(p, text, bold=False)

    if bookmark_name:
        bm_id = bookmark_id if bookmark_id is not None else 0
        bm_start = ET.SubElement(p, f'{{{W}}}bookmarkStart')
        bm_start.set(f'{{{W}}}id', str(bm_id))
        bm_start.set(f'{{{W}}}name', bookmark_name)
        bm_end = ET.SubElement(p, f'{{{W}}}bookmarkEnd')
        bm_end.set(f'{{{W}}}id', str(bm_id))

    return p


def build_body_paragraph(block):
    """构建正文段落（style aa, 首行缩进）。表格标题用 af4 样式。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')

    # 表格标题（如 "表2-1 xxx" 或 "续表2-1 xxx"）使用图注样式
    if re.match(r'^(续)?表\d', text):
        pstyle.set(f'{{{W}}}val', 'af4')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        add_text_to_paragraph(p, text)
        return p

    pstyle.set(f'{{{W}}}val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLineChars', '200')
    ind.set(f'{{{W}}}firstLine', '480')
    add_text_to_paragraph(p, text)
    return p


def build_bullet_paragraph(block):
    """构建无序列表段落。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
    ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
    ilvl.set(f'{{{W}}}val', '0')
    numid = ET.SubElement(numpr, f'{{{W}}}numId')
    numid.set(f'{{{W}}}val', '2')
    add_text_to_paragraph(p, text)
    return p


def build_numbered_paragraph(block):
    """构建编号列表段落。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    numpr = ET.SubElement(ppr, f'{{{W}}}numPr')
    ilvl = ET.SubElement(numpr, f'{{{W}}}ilvl')
    ilvl.set(f'{{{W}}}val', '0')
    numid = ET.SubElement(numpr, f'{{{W}}}numId')
    numid.set(f'{{{W}}}val', '1')
    add_text_to_paragraph(p, text)
    return p


def build_blockquote(block):
    """构建引用块段落（图注等，style af4, 居中）。"""
    text = block['text']
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'af4')
    add_text_to_paragraph(p, text)
    return p


def build_codeblock(block):
    """构建代码块（等宽字体，无缩进）。"""
    code = block['code']
    lines = code.split('\n')

    # 每行创建一个独立段落
    paragraphs = []
    for line in lines:
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'aa')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLineChars', '0')
        ind.set(f'{{{W}}}firstLine', '0')

        r = make_text_run(line or ' ', font_ascii='Courier New', font_east_asia='Courier New')
        p.append(r)
        paragraphs.append(p)

    return paragraphs  # 返回段落列表


def _render_formula_with_xelatex(latex, tag_text=None):
    """用 xelatex 将 LaTeX 公式编译为 PNG 图片。"""
    formula_text = latex.strip().strip('$').strip()
    # 去掉 \tag{} 编号（单独处理）
    if tag_text:
        formula_text = re.sub(r'\\tag\{[^}]+\}', '', formula_text).strip()

    # 生成完整的 LaTeX 文档
    # standalone preview 自动裁剪，scalebox 控制公式大小（2.0 ≈ 约四号）
    doc = r"\documentclass[preview,border=2pt]{standalone}" + "\n"
    doc += r"\usepackage{amsmath,amssymb}" + "\n"
    doc += r"\usepackage{unicode-math}" + "\n"
    doc += r"\usepackage{graphicx}" + "\n"
    doc += r"\begin{document}" + "\n"
    doc += r"\scalebox{2.0}{$\displaystyle " + formula_text + r"$}" + "\n"
    doc += r"\end{document}"

    # 在临时目录中编译
    tmpdir = tempfile.mkdtemp()
    tex_path = os.path.join(tmpdir, 'formula.tex')
    with open(tex_path, 'w') as f:
        f.write(doc)

    # xelatex 编译
    result = subprocess.run(
        ['xelatex', '-interaction=nonstopmode', '-output-directory', tmpdir, tex_path],
        capture_output=True, timeout=30
    )

    # 检查 pdf 是否生成
    pdf_path = os.path.join(tmpdir, 'formula.pdf')
    if not os.path.exists(pdf_path):
        raise RuntimeError(f'xelatex 编译失败: {result.stderr.decode()}')

    # pdf 转 png（使用 PyMuPDF 以 600 DPI 渲染，保证字体清晰锐利）
    import pymupdf
    pdfdoc = pymupdf.open(pdf_path)
    page = pdfdoc[0]
    zoom = 600 / 72  # 600 DPI
    mat = pymupdf.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_bytes = pix.tobytes('png')
    pdfdoc.close()
    return img_bytes


def build_formula_paragraph(block, media_files_dict=None, rels_counter=None):
    """构建公式段落：将 LaTeX 渲染为 PNG 图片并嵌入 docx。"""
    latex = block['latex']
    try:
        # 解析 \tag{...} 编号
        tag_match = re.search(r'\\tag\{([^}]+)\}', latex)
        tag_text = None
        if tag_match:
            tag_text = tag_match.group(1)

        img_bytes = _render_formula_with_xelatex(latex, tag_text)

        # 保存到 word/media/
        img_filename = f'formula_{hash(latex) & 0xFFFFFFFF:08x}.png'
        media_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media', img_filename)
        with open(media_path, 'wb') as f:
            f.write(img_bytes)

        if media_files_dict is not None:
            media_files_dict[f'word/media/{img_filename}'] = f'word/media/{img_filename}'

        # 公式图片按实际像素比例显示（600 DPI 渲染）
        # EMU = pixels / DPI * 72 * 9525
        w_px = int.from_bytes(img_bytes[16:20], 'big')
        h_px = int.from_bytes(img_bytes[20:24], 'big')
        cx = round(w_px / 600 * 72 * 9525)
        cy = round(h_px / 600 * 72 * 9525)

        r_id = None
        if media_files_dict is not None and '_rels_entries' in rels_counter:
            rels_counter['count'] += 1
            r_id = f'rIdImage{rels_counter["count"]}'
            rels_counter['_rels_entries'].append((r_id, f'media/{img_filename}'))
        else:
            r_id = 'rIdImage1'

        # 构建段落：公式图片居中，编号（如有）右对齐，用 tab 分隔
        PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
        WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
        A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        p_adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        p_adj.set(f'{{{W}}}val', '0')
        p_snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        p_snap.set(f'{{{W}}}val', '0')
        ET.SubElement(ppr, f'{{{W}}}keepNext')
        ET.SubElement(ppr, f'{{{W}}}keepLines')

        if tag_text:
            # 有编号：居中 tab + 图片 + 右对齐 tab + 编号
            tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
            tab_center = ET.SubElement(tabs, f'{{{W}}}tab')
            tab_center.set(f'{{{W}}}val', 'center')
            tab_center.set(f'{{{W}}}pos', '4536')  # 页面中间约 8cm
            tab_center.set(f'{{{W}}}leader', 'none')
            tab_right = ET.SubElement(tabs, f'{{{W}}}tab')
            tab_right.set(f'{{{W}}}val', 'right')
            tab_right.set(f'{{{W}}}pos', '9072')  # 右边距约 16cm
            tab_right.set(f'{{{W}}}leader', 'none')

            # 居中 tab run
            r_tab = ET.SubElement(p, f'{{{W}}}r')
            ET.SubElement(r_tab, f'{{{W}}}tab')
        else:
            # 无编号：直接居中
            p_jc = ET.SubElement(ppr, f'{{{W}}}jc')
            p_jc.set(f'{{{W}}}val', 'center')

        # 公式图片 run
        r = ET.SubElement(p, f'{{{W}}}r')
        rpr = ET.SubElement(r, f'{{{W}}}rPr')
        ET.SubElement(rpr, f'{{{W}}}noProof')

        drawing = ET.SubElement(r, f'{{{W}}}drawing')
        inline = ET.SubElement(drawing, f'{{{WP_NS}}}inline')
        ET.SubElement(inline, f'{{{WP_NS}}}extent', attrib={'cx': str(cx), 'cy': str(cy)})
        ET.SubElement(inline, f'{{{WP_NS}}}effectExtent', attrib={'l': '0', 't': '0', 'r': '0', 'b': '0'})
        ET.SubElement(inline, f'{{{WP_NS}}}docPr', attrib={'id': '1', 'name': 'Picture'})
        cNvGFP = ET.SubElement(inline, f'{{{PIC_NS}}}cNvGraphicFramePr')
        graphic = ET.SubElement(inline, f'{{{A_NS}}}graphic')
        graphicData = ET.SubElement(graphic, f'{{{A_NS}}}graphicData',
                                    attrib={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})
        pic = ET.SubElement(graphicData, f'{{{PIC_NS}}}pic')
        nvPicPr = ET.SubElement(pic, f'{{{PIC_NS}}}nvPicPr')
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPr', attrib={'id': '0', 'name': 'Picture'})
        ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPicPr')
        blipFill = ET.SubElement(pic, f'{{{PIC_NS}}}blipFill')
        ET.SubElement(blipFill, f'{{{A_NS}}}blip',
                      attrib={f'{{{R_NS}}}embed': r_id})
        stretch = ET.SubElement(blipFill, f'{{{A_NS}}}stretch')
        ET.SubElement(stretch, f'{{{A_NS}}}fillRect')
        spPr = ET.SubElement(pic, f'{{{PIC_NS}}}spPr')
        xfrm = ET.SubElement(spPr, f'{{{A_NS}}}xfrm')
        ET.SubElement(xfrm, f'{{{A_NS}}}off', attrib={'x': '0', 'y': '0'})
        ET.SubElement(xfrm, f'{{{A_NS}}}ext', attrib={'cx': str(cx), 'cy': str(cy)})
        prstGeom = ET.SubElement(spPr, f'{{{A_NS}}}prstGeom')
        prstGeom.set('prst', 'rect')
        ET.SubElement(prstGeom, f'{{{A_NS}}}avLst')

        if tag_text:
            # tab 字符 run
            r_tab = ET.SubElement(p, f'{{{W}}}r')
            r_tab_text = ET.SubElement(r_tab, f'{{{W}}}tab')

            # 编号 run
            r2 = make_text_run(f'（{tag_text}）', font_east_asia='黑体', sz=21)
            p.append(r2)

        return [p]

    except Exception as e:
        print(f'  警告: LaTeX 公式渲染失败: {e}', file=sys.stderr)
        # 回退：居中显示 LaTeX 文本
        p = ET.Element(f'{{{W}}}p')
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', 'aa')
        jc = ET.SubElement(ppr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', 'center')
        adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
        adj.set(f'{{{W}}}val', '0')
        snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
        snap.set(f'{{{W}}}val', '0')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLineChars', '0')
        ind.set(f'{{{W}}}firstLine', '0')
        r = make_text_run(latex, font_ascii='Cambria Math', font_east_asia='Cambria Math', italic=True)
        p.append(r)
        return [p]


def build_image_block(block, media_files_dict=None, rels_counter=None):
    """下载图片并嵌入 docx。返回段落列表。"""
    url = block.get('url', '').strip()
    alt = block.get('alt', '')

    if not url:
        return [build_image_placeholder(block)]

    try:
        # 下载图片
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        response = urllib.request.urlopen(req, timeout=15)
        img_bytes = response.read()
        content_type = response.headers.get('Content-Type', '')

        # 确定文件扩展名
        if 'png' in content_type:
            ext = 'png'
        elif 'gif' in content_type:
            ext = 'gif'
        elif 'bmp' in content_type:
            ext = 'bmp'
        else:
            ext = 'jpeg'

        img_filename = f'image_{hash(url) & 0xFFFFFFFF:08x}.{ext}'
        media_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'word', 'media', img_filename)
        with open(media_path, 'wb') as f:
            f.write(img_bytes)

        if media_files_dict is not None:
            media_files_dict[f'word/media/{img_filename}'] = f'word/media/{img_filename}'

        r_id, cx, cy = _make_image_rel(media_files_dict, rels_counter, img_filename, img_bytes, 450, 300)

        p = _build_image_paragraph(r_id, cx, cy)
        result = [p]

        # 添加图注段落（模板 P98: style=af4, jc=center, rFonts={ascii:'黑体', hAnsi:'黑体'}, szCs=21 五号）
        if alt:
            p_cap = ET.Element(f'{{{W}}}p')
            p_cap_ppr = ET.SubElement(p_cap, f'{{{W}}}pPr')
            p_cap_style = ET.SubElement(p_cap_ppr, f'{{{W}}}pStyle')
            p_cap_style.set(f'{{{W}}}val', 'af4')
            p_cap_jc = ET.SubElement(p_cap_ppr, f'{{{W}}}jc')
            p_cap_jc.set(f'{{{W}}}val', 'center')
            p_cap_adj = ET.SubElement(p_cap_ppr, f'{{{W}}}adjustRightInd')
            p_cap_adj.set(f'{{{W}}}val', '0')
            p_cap_snap = ET.SubElement(p_cap_ppr, f'{{{W}}}snapToGrid')
            p_cap_snap.set(f'{{{W}}}val', '0')
            cap_run = ET.SubElement(p_cap, f'{{{W}}}r')
            cap_rpr = ET.SubElement(cap_run, f'{{{W}}}rPr')
            cap_fonts = ET.SubElement(cap_rpr, f'{{{W}}}rFonts')
            cap_fonts.set(f'{{{W}}}ascii', '黑体')
            cap_fonts.set(f'{{{W}}}hAnsi', '黑体')
            cap_scs = ET.SubElement(cap_rpr, f'{{{W}}}szCs')
            cap_scs.set(f'{{{W}}}val', '21')
            cap_t = ET.SubElement(cap_run, f'{{{W}}}t')
            cap_t.text = alt
            cap_t.set(f'{XML_NS}space', 'preserve')
            result.append(p_cap)

        return result

    except Exception as e:
        print(f'  警告: 图片下载失败 ({url}): {e}', file=sys.stderr)
        return [build_image_placeholder(block)]


def _make_image_rel(media_files_dict, rels_counter, img_filename, img_bytes, default_cx, default_cy, no_scale=False):
    """为图片创建关系 ID 并计算尺寸。返回 (r_id, cx_emu, cy_emu)。
    no_scale=True 时直接使用 default_cx/default_cy（单位 pt），不做 max_w_emu 缩放。"""
    if rels_counter is not None:
        rels_counter['count'] += 1
        r_id = f'rIdImage{rels_counter["count"]}'
    else:
        r_id = 'rIdImage1'

    # 记录 r_id 和文件名的映射到 rels_entries（Target 相对于 word/_rels/ 目录）
    if rels_counter is not None and '_rels_entries' in rels_counter:
        rels_counter['_rels_entries'].append((r_id, f'media/{img_filename}'))

    # 尝试获取图片尺寸，统一按页面可用宽度等比缩放
    cx, cy = default_cx * 9525, default_cy * 9525  # pt → EMU
    if not no_scale:
        max_w_emu = 415 * 9525  # 页面可用宽度 ≈ 14.66cm
    try:
        if img_bytes[:8] == b'\x89PNG\r\n\x1a\n':
            # PNG
            w = int.from_bytes(img_bytes[16:20], 'big')
            h = int.from_bytes(img_bytes[20:24], 'big')
            if w > 0 and h > 0:
                scale = min(1.0, max_w_emu / (w * 9525))
                cx = int(max_w_emu)
                cy = int(h * 9525 * scale)
        elif img_bytes[:2] == b'\xff\xd8':
            # JPEG
            i = 2
            while i < len(img_bytes) - 1:
                if img_bytes[i] != 0xFF:
                    break
                marker = img_bytes[i + 1]
                if marker == 0xC0 or marker == 0xC2:
                    h = int.from_bytes(img_bytes[i + 5:i + 7], 'big')
                    w = int.from_bytes(img_bytes[i + 7:i + 9], 'big')
                    if w > 0 and h > 0:
                        scale = min(1.0, max_w_emu / (w * 9525))
                        cx = int(max_w_emu)
                        cy = int(h * 9525 * scale)
                    break
                elif marker == 0xD9 or marker == 0xDA:
                    break
                else:
                    length = int.from_bytes(img_bytes[i + 2:i + 4], 'big')
                    i += 2 + length
    except Exception:
        pass

    return r_id, cx, cy


def _build_image_paragraph(r_id, cx, cy):
    """构建包含嵌入式图片的居中段落。"""
    PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    jc = ET.SubElement(ppr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ET.SubElement(ppr, f'{{{W}}}keepNext')
    ET.SubElement(ppr, f'{{{W}}}keepLines')

    r = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    ET.SubElement(rpr, f'{{{W}}}noProof')

    drawing = ET.SubElement(r, f'{{{W}}}drawing')

    # wp:inline
    WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
    A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    inline = ET.SubElement(drawing, f'{{{WP_NS}}}inline')
    ET.SubElement(inline, f'{{{WP_NS}}}extent', attrib={'cx': str(cx), 'cy': str(cy)})
    ET.SubElement(inline, f'{{{WP_NS}}}effectExtent', attrib={'l': '0', 't': '0', 'r': '0', 'b': '0'})
    docPr = ET.SubElement(inline, f'{{{WP_NS}}}docPr', attrib={'id': '1', 'name': 'Picture'})

    # cNvGraphicFramePr
    cNvGFP = ET.SubElement(inline, f'{{{PIC_NS}}}cNvGraphicFramePr')

    # graphic
    graphic = ET.SubElement(inline, f'{{{A_NS}}}graphic')
    graphicData = ET.SubElement(graphic, f'{{{A_NS}}}graphicData',
                                attrib={'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'})

    # pic:pic
    pic = ET.SubElement(graphicData, f'{{{PIC_NS}}}pic')
    nvPicPr = ET.SubElement(pic, f'{{{PIC_NS}}}nvPicPr')
    cNvPr = ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPr', attrib={'id': '0', 'name': 'Picture'})
    cNvPicPr = ET.SubElement(nvPicPr, f'{{{PIC_NS}}}cNvPicPr')

    blipFill = ET.SubElement(pic, f'{{{PIC_NS}}}blipFill')
    blip = ET.SubElement(blipFill, f'{{{A_NS}}}blip',
                          attrib={f'{{{R_NS}}}embed': r_id})
    stretch = ET.SubElement(blipFill, f'{{{A_NS}}}stretch')
    ET.SubElement(stretch, f'{{{A_NS}}}fillRect')

    spPr = ET.SubElement(pic, f'{{{PIC_NS}}}spPr')
    xfrm = ET.SubElement(spPr, f'{{{A_NS}}}xfrm')
    ET.SubElement(xfrm, f'{{{A_NS}}}off', attrib={'x': '0', 'y': '0'})
    ET.SubElement(xfrm, f'{{{A_NS}}}ext', attrib={'cx': str(cx), 'cy': str(cy)})
    prstGeom = ET.SubElement(spPr, f'{{{A_NS}}}prstGeom', attrib={'prst': 'rect'})
    ET.SubElement(prstGeom, f'{{{A_NS}}}avLst')

    return p


def build_image_placeholder(block):
    """构建图片占位段落（显示图片说明文字，居中）。"""
    alt = block.get('alt', block.get('url', ''))
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'af4')
    add_text_to_paragraph(p, f'[{alt}]')
    return p


def build_table(block):
    """构建 Word 表格 XML。"""
    headers = block['headers']
    rows = block['rows']
    num_cols = len(headers)

    # 均分列宽
    total_width = 8312  # 页面可用宽度 (pgSz w - left margin - right margin)
    col_width = total_width // num_cols

    tbl = ET.Element(f'{{{W}}}tbl')

    # 表格属性
    tblPr = ET.SubElement(tbl, f'{{{W}}}tblPr')
    tblW = ET.SubElement(tblPr, f'{{{W}}}tblW')
    tblW.set(f'{{{W}}}w', '0')
    tblW.set(f'{{{W}}}type', 'auto')
    jc = ET.SubElement(tblPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')
    tblBorders = ET.SubElement(tblPr, f'{{{W}}}tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = ET.SubElement(tblBorders, f'{{{W}}}{border_name}')
        b.set(f'{{{W}}}val', 'single')
        b.set('sz', '4')
        b.set('space', '0')
        b.set('color', 'auto')
    tblLayout = ET.SubElement(tblPr, f'{{{W}}}tblLayout')
    tblLayout.set(f'{{{W}}}type', 'fixed')

    # 列定义
    tblGrid = ET.SubElement(tbl, f'{{{W}}}tblGrid')
    for _ in range(num_cols):
        gc = ET.SubElement(tblGrid, f'{{{W}}}gridCol')
        gc.set(f'{{{W}}}w', str(col_width))

    # 构建行的辅助函数
    def _make_row(cells_data, is_header=False):
        tr = ET.SubElement(tbl, f'{{{W}}}tr')
        if is_header:
            trPr = ET.SubElement(tr, f'{{{W}}}trPr')
            ET.SubElement(trPr, f'{{{W}}}tblHeader')
        for cell_text in cells_data:
            tc = ET.SubElement(tr, f'{{{W}}}tc')
            tcPr = ET.SubElement(tc, f'{{{W}}}tcPr')
            tcW = ET.SubElement(tcPr, f'{{{W}}}tcW')
            tcW.set(f'{{{W}}}w', str(col_width))
            tcW.set(f'{{{W}}}type', 'dxa')
            vAlign = ET.SubElement(tcPr, f'{{{W}}}vAlign')
            vAlign.set(f'{{{W}}}val', 'center')

            p = ET.SubElement(tc, f'{{{W}}}p')
            ppr = ET.SubElement(p, f'{{{W}}}pPr')
            pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
            pstyle.set(f'{{{W}}}val', 'aa')
            adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
            adj.set(f'{{{W}}}val', '0')
            snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
            snap.set(f'{{{W}}}val', '0')
            ind = ET.SubElement(ppr, f'{{{W}}}ind')
            ind.set(f'{{{W}}}firstLineChars', '0')
            ind.set(f'{{{W}}}firstLine', '0')

            if is_header:
                # 表头加粗
                r = make_text_run(cell_text, bold=True, font_ascii='Times New Roman')
                p.append(r)
            else:
                add_text_to_paragraph(p, cell_text)

    # 表头
    _make_row(headers, is_header=True)
    # 数据行
    for row in rows:
        # 补齐列数
        while len(row) < num_cols:
            row.append('')
        _make_row(row)

    return tbl


def build_hr_paragraph():
    """构建分隔线段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    pBdr = ET.SubElement(ppr, f'{{{W}}}pBdr')
    bottom = ET.SubElement(pBdr, f'{{{W}}}bottom')
    bottom.set(f'{{{W}}}val', 'single')
    bottom.set('sz', '6')
    bottom.set('space', '1')
    bottom.set('color', 'auto')
    return p


def build_page_break_paragraph():
    """构建分页符段落。"""
    p = ET.Element(f'{{{W}}}p')
    r = ET.SubElement(p, f'{{{W}}}r')
    br = ET.SubElement(r, f'{{{W}}}br')
    br.set(f'{{{W}}}type', 'page')
    return p


# ============================================================
# 5. 目录生成
# ============================================================

_bookmark_counter = 0


def next_bookmark_name():
    global _bookmark_counter
    _bookmark_counter += 1
    return f'_Toc{999999999 + _bookmark_counter}'


def build_toc_entry(number, title, toc_level, bookmark_name='_Toc1000000000'):
    """构建目录条目段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')

    if toc_level == 1:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', '12')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab1 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab1.set(f'{{{W}}}val', 'left')
        tab1.set(f'{{{W}}}pos', '630')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set(f'{{{W}}}val', 'right')
        tab2.set('leader', 'dot')
        tab2.set(f'{{{W}}}pos', '8302')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '21')
        szCs = ET.SubElement(rpr_in_ppr, f'{{{W}}}szCs')
        szCs.set(f'{{{W}}}val', '22')
    elif toc_level == 2:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', '20')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab1 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab1.set(f'{{{W}}}val', 'left')
        tab1.set(f'{{{W}}}pos', '1260')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set(f'{{{W}}}val', 'right')
        tab2.set('leader', 'dot')
        tab2.set(f'{{{W}}}pos', '8302')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLine', '480')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        kern = ET.SubElement(rpr_in_ppr, f'{{{W}}}kern')
        kern.set(f'{{{W}}}val', '2')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '21')
    else:
        pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
        pstyle.set(f'{{{W}}}val', '30')
        tabs = ET.SubElement(ppr, f'{{{W}}}tabs')
        tab1 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab1.set(f'{{{W}}}val', 'left')
        tab1.set(f'{{{W}}}pos', '1710')
        tab2 = ET.SubElement(tabs, f'{{{W}}}tab')
        tab2.set(f'{{{W}}}val', 'right')
        tab2.set('leader', 'dot')
        tab2.set(f'{{{W}}}pos', '8302')
        ind = ET.SubElement(ppr, f'{{{W}}}ind')
        ind.set(f'{{{W}}}firstLine', '960')
        rpr_in_ppr = ET.SubElement(ppr, f'{{{W}}}rPr')
        fonts = ET.SubElement(rpr_in_ppr, f'{{{W}}}rFonts')
        fonts.set(f'{{{W}}}ascii', 'Times New Roman')
        fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
        fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
        no_proof = ET.SubElement(rpr_in_ppr, f'{{{W}}}noProof')
        kern = ET.SubElement(rpr_in_ppr, f'{{{W}}}kern')
        kern.set(f'{{{W}}}val', '2')
        sz = ET.SubElement(rpr_in_ppr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '21')

    def _r(text_content=None):
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
            t.set(f'{XML_NS}space', 'preserve')
        return r

    # 编号 run
    if number:
        _r(number)

    # tab
    tab_r = ET.SubElement(p, f'{{{W}}}r')
    rpr = ET.SubElement(tab_r, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    fonts.set(f'{{{W}}}ascii', 'Times New Roman')
    fonts.set(f'{{{W}}}eastAsiaTheme', 'minorEastAsia')
    if toc_level == 1:
        b = ET.SubElement(rpr, f'{{{W}}}b')
        szCs = ET.SubElement(rpr, f'{{{W}}}szCs')
        szCs.set(f'{{{W}}}val', '28')
    no_proof = ET.SubElement(rpr, f'{{{W}}}noProof')
    ET.SubElement(tab_r, f'{{{W}}}tab')

    # 标题 run
    _r(title)

    # tab before page number
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
    instr.set(f'{XML_NS}space', 'preserve')
    instr.text = f' PAGEREF {bookmark_name} \\h '

    # empty run
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

    # 页码占位
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


def build_toc_entries(blocks):
    """从 Markdown blocks 中提取标题，生成目录条目列表。"""
    global _bookmark_counter
    entries = []
    h1_count = 0
    h2_count = 0
    h3_count = 0

    for block in blocks:
        if block['type'] != 'heading':
            continue
        level = block['level']
        if level > 3:
            level = 3

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
        # 从标题中分离编号和标题文字（标题可能已包含编号如 "1 绪论"）
        title_without_number = clean_title
        if level == 1:
            m = re.match(r'^(\d+)\s+(.+)$', clean_title)
            if m:
                title_without_number = m.group(2)
        elif level == 2:
            m = re.match(r'^(\d+\.\d+)\s+(.+)$', clean_title)
            if m:
                title_without_number = m.group(2)
        elif level == 3:
            m = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)$', clean_title)
            if m:
                title_without_number = m.group(2)

        bookmark = next_bookmark_name()
        block['_bookmark'] = bookmark
        block['_bookmark_id'] = _bookmark_counter
        entries.append({
            'number': number,
            'title': title_without_number,
            'toc_level': toc_level,
            'bookmark': bookmark,
        })

    return entries


# ============================================================
# 6. Bookmark 管理
# ============================================================

def find_max_bookmark_id(body):
    """扫描 body 中所有 bookmarkStart 的最大 id。"""
    max_id = -1
    for bm in body.iter(f'{{{W}}}bookmarkStart'):
        bid = bm.get(f'{{{W}}}id')
        if bid is not None:
            try:
                max_id = max(max_id, int(bid))
            except ValueError:
                pass
    return max_id


def remove_old_bookmarks(body):
    """移除整个 body 中所有 _Toc bookmarkStart 及配对的 bookmarkEnd。"""
    parent_map = {c: p for p in body.iter() for c in p}

    for bm in list(body.iter(f'{{{W}}}bookmarkStart')):
        name = bm.get(f'{{{W}}}name', '')
        if name.startswith('_Toc'):
            parent = parent_map.get(bm)
            if parent is not None:
                parent.remove(bm)

    parent_map = {c: p for p in body.iter() for c in p}

    for bm in list(body.iter(f'{{{W}}}bookmarkEnd')):
        bid = bm.get(f'{{{W}}}id', '')
        has_start = any(b.get(f'{{{W}}}id') == bid for b in body.iter(f'{{{W}}}bookmarkStart'))
        if not has_start:
            parent = parent_map.get(bm)
            if parent is not None:
                parent.remove(bm)


# ============================================================
# 7. 封面注入
# ============================================================

def _find_label_end_run(paragraph, label_patterns):
    """
    在段落中找到标签结束的位置。
    label_patterns: 标签文字列表（如 ['题    目', '学    号']）
    返回标签结束后的第一个 run 索引。
    """
    runs = paragraph.findall(f'{{{W}}}r')
    accumulated = ''
    for i, r in enumerate(runs):
        t = r.find(f'{{{W}}}t')
        if t is not None and t.text:
            accumulated += t.text
        # 检查累积文本是否匹配某个标签
        for pat in label_patterns:
            if pat in accumulated:
                return i + 1
    return 1  # 默认跳过第一个 run


def _make_cover_run(text, is_label=False, english_value=False):
    """创建封面字段 run，严格对照模板格式。
    模板 P4-P13 的 run 格式：
      - 标签 run：rFonts={hAnsi:'黑体'}, b={val:'0'}, sz={val:'44'}, szCs={val:'44'}, 无下划线
      - 值 run（中文）：同标签 + u={val:'single'}
      - P5 label：rFonts={asciiTheme:'minorHAnsi', hAnsiTheme:'minorHAnsi', cstheme:'minorHAnsi'}, sz=44
      - P5 value：同上 + sz=28 + u={val:'single'}，尾部空格 sz=24 + u={val:'single'}
    is_label=True: 标签 run，无下划线
    is_label=False: 值 run，有下划线
    english_value=True: P5 英文题目，用 theme 字体
    """
    r = ET.Element(f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    b = ET.SubElement(rpr, f'{{{W}}}b')
    b.set(f'{{{W}}}val', '0')

    if english_value:
        # P5 英文题目：theme font minorHAnsi (Calibri)
        fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
        fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
        fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
        if is_label:
            s = ET.SubElement(rpr, f'{{{W}}}sz')
            s.set(f'{{{W}}}val', '44')
        else:
            s = ET.SubElement(rpr, f'{{{W}}}sz')
            s.set(f'{{{W}}}val', '28')
    else:
        # 中文封面：hAnsi=黑体
        fonts.set(f'{{{W}}}hAnsi', '黑体')
        s = ET.SubElement(rpr, f'{{{W}}}sz')
        s.set(f'{{{W}}}val', '44')

    if not is_label:
        # 值 run 有下划线
        u = ET.SubElement(rpr, f'{{{W}}}u')
        u.set(f'{{{W}}}val', 'single')

    scs = ET.SubElement(rpr, f'{{{W}}}szCs')
    scs.set(f'{{{W}}}val', '44')
    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = text
    t.set(f'{XML_NS}space', 'preserve')
    return r


def _make_cover_trailing_spaces(english_value=False):
    """创建封面值后面的填充空格 run（带下划线），对照模板格式延伸到行尾。
    模板中尾部空格 run：sz=24 (中文) 或 sz=24 (英文), szCs=44, u=single"""
    r = ET.Element(f'{{{W}}}r')
    rpr = ET.SubElement(r, f'{{{W}}}rPr')
    fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
    b = ET.SubElement(rpr, f'{{{W}}}b')
    b.set(f'{{{W}}}val', '0')
    u = ET.SubElement(rpr, f'{{{W}}}u')
    u.set(f'{{{W}}}val', 'single')

    if english_value:
        fonts.set(f'{{{W}}}asciiTheme', 'minorHAnsi')
        fonts.set(f'{{{W}}}hAnsiTheme', 'minorHAnsi')
        fonts.set(f'{{{W}}}cstheme', 'minorHAnsi')
    else:
        fonts.set(f'{{{W}}}hAnsi', '黑体')

    s = ET.SubElement(rpr, f'{{{W}}}sz')
    s.set(f'{{{W}}}val', '24')
    scs = ET.SubElement(rpr, f'{{{W}}}szCs')
    scs.set(f'{{{W}}}val', '44')
    t = ET.SubElement(r, f'{{{W}}}t')
    t.text = '                    '  # 模板中约 10-20 个空格
    t.set(f'{XML_NS}space', 'preserve')
    return r


def inject_cover_data(paragraphs, cover_data):
    """
    将 final_paper.md 中的封面信息注入到模板封面段落（P4-P13）。
    对含有 AlternateContent/dropdown 的段落（P7, P10, P13）清除所有 runs 重建；
    其他段落保留标签 runs，替换后续内容。
    """
    # 段落标签文字和对应字段
    field_map = {
        4: ('题目', '题    目：'),
        5: ('英文题目', '英文题目：'),
        6: ('学号', '学    号：'),
        7: ('姓名', '姓    名：'),
        8: ('班级', '班    级：'),
        9: ('专业', '专    业：'),
        10: ('学部(院)', '学部(院)：'),
        11: ('入学时间', '入学时间：'),
        12: ('指导教师', '指导教师：'),
        13: ('日期', '日    期：'),
    }

    # 含 AlternateContent 提示文本的段落，需要完全重建
    rebuild_paragraphs = {7, 10, 13}

    for p_idx, (key, label) in field_map.items():
        if key not in cover_data:
            continue
        value = cover_data[key]
        if p_idx >= len(paragraphs):
            continue

        p = paragraphs[p_idx]
        runs = p.findall(f'{{{W}}}r')

        if p_idx in rebuild_paragraphs:
            # 移除所有 runs，保留 pPr
            for r in list(runs):
                p.remove(r)
            # 重建：标签（无下划线） + 值（有下划线） + 尾部空格（有下划线）
            label_run = _make_cover_run(label, is_label=True)
            p.append(label_run)
            value_run = _make_cover_run(value, english_value=(p_idx == 5))
            p.append(value_run)
            trailing = _make_cover_trailing_spaces(english_value=(p_idx == 5))
            p.append(trailing)
        else:
            # 常规处理：找冒号位置，保留标签，替换后续内容
            label_end = 0
            accumulated = ''
            for i, r in enumerate(runs):
                # 搜索所有子元素中的 w:t（包括 AlternateContent 内部）
                text = ''
                for t in r.iter(f'{{{W}}}t'):
                    if t.text:
                        text += t.text
                accumulated += text
                if '：' in accumulated:
                    label_end = i + 1
                    break

            runs_to_remove = runs[label_end:]
            for r in runs_to_remove:
                p.remove(r)

            new_run = _make_cover_run(value, english_value=(p_idx == 5))
            p.append(new_run)
            trailing = _make_cover_trailing_spaces(english_value=(p_idx == 5))
            p.append(trailing)



# ============================================================
# 8. 摘要替换
# ============================================================

def replace_abstracts(paragraphs, blocks, parent_body=None):
    """
    替换中英文摘要段落。
    模板结构（基于 document.xml 实际段落索引）：
      P31: 中文标题（含 AlternateContent 格式说明 → 需清理）
      P32: "摘要"标签（含格式说明 → 需清理）
      P33-P35: 中文摘要正文（含格式说明 → 需清理并替换）
      P36-P37: bookmarkEnd 元素（跳过）
      P38: 中文关键词（"关键词：网络爬虫；股票预警；WEB挖掘"）
      P39: 中文关键词格式说明（需清空或移除内容）
      P40: 分页符段落
      P41: 英文标题（含格式说明 → 需清理）
      P42: "ABSTRACT"标签（含格式说明 → 需清理）
      P43: 英文摘要正文
      P44: 英文关键词（含格式说明 → 需清理）
    """
    # === 清理 P31 格式说明文字（保留模板 pPr，清除 AlternateContent 和旧 runs） ===
    _clean_paragraph_keep_ppr(paragraphs[31])
    # 重新添加标题 run（模板格式：居中，黑体，sz=36，已有 pPr）
    title_run = make_text_run('基于网络爬虫的股票信息预警系统的设计与实现',
                               font_east_asia='黑体', sz=36)
    paragraphs[31].append(title_run)

    # === 清理 P32 格式说明文字，只保留"摘要" ===
    _clean_paragraph_keep_ppr(paragraphs[32])
    title_run2 = make_text_run('摘要', font_east_asia='黑体', bold=True,
                                sz=32)
    paragraphs[32].append(title_run2)

    # === 清理 P37（关键词格式说明段落）— 清空内容 ===
    _clean_paragraph_keep_ppr(paragraphs[37])

    # === 清理 P38（分页符段落）— 移除格式说明文字，保留分页符 ===
    _clean_paragraph_keep_ppr(paragraphs[38])
    # 重新添加分页符 run
    r_pb = ET.SubElement(paragraphs[38], f'{{{W}}}r')
    br = ET.SubElement(r_pb, f'{{{W}}}br')
    br.set(f'{{{W}}}type', 'page')

    # === 清理 P39（英文标题）===
    _clean_paragraph_keep_ppr(paragraphs[39])
    # 小二(18pt=36半点)、Times New Roman、居中、加粗、大写
    _add_jc_to_paragraph(paragraphs[39], 'center')
    en_title_run = make_text_run('DESIGN AND IMPLEMENTATION OF STOCK INFORMATION EARLY WARNING SYSTEM BASED ON WEB CRAWLER',
                                  font_ascii='Times New Roman', sz=36, bold=True)
    paragraphs[39].append(en_title_run)

    # === 清理 P40（ABSTRACT 标签）===
    _clean_paragraph_keep_ppr(paragraphs[40])
    # 三号(16pt=32半点)、Times New Roman、居中、加粗、大写
    _add_jc_to_paragraph(paragraphs[40], 'center')
    abstract_run = make_text_run('ABSTRACT', font_ascii='Times New Roman', sz=32, bold=True)
    paragraphs[40].append(abstract_run)

    # === 清理 P42（英文关键词格式说明）===
    _clean_paragraph_keep_ppr(paragraphs[42])

    # 找到中文摘要正文块（"摘要"标题之后，"ABSTRACT"标题之前的内容块）
    cn_abstract_blocks = []
    en_abstract_blocks = []
    in_cn_abstract = False
    in_en_abstract = False

    for block in blocks:
        if block['type'] == 'heading':
            text = strip_markdown_formatting(block['text'])
            if '摘要' in text and 'ABSTRACT' not in text.upper():
                in_cn_abstract = True
                in_en_abstract = False
                continue
            elif 'ABSTRACT' in text.upper():
                in_cn_abstract = False
                in_en_abstract = True
                continue
            else:
                in_cn_abstract = False
                in_en_abstract = False
                continue

        if block['type'] == 'hr':
            in_cn_abstract = False
            in_en_abstract = False
            continue

        if block['type'] in ('paragraph', 'bullet', 'numbered'):
            # 排除关键词段落（单独处理）
            text = block.get('text', '')
            if '关键词' in text or 'key words' in text.lower() or 'keywords' in text.lower():
                continue
            if in_cn_abstract:
                cn_abstract_blocks.append(block)
            elif in_en_abstract:
                en_abstract_blocks.append(block)

    # 替换中文摘要正文（P33-P35 → 使用摘要专用格式）
    _replace_abstract_body(paragraphs, 33, 35, cn_abstract_blocks, chinese=True,
                           parent_body=parent_body, insert_before_idx=36)

    # 替换英文摘要正文（P41 → 合并多段为一个段落，避免插入破坏后续索引）
    # 模板 P41 格式：rFonts={ascii:'Times New Roman', eastAsia:'BatangChe', hAnsi:'Times New Roman'},
    #   b={val:'0'}, bCs={val:'0'}, sz={val:'24'}
    if en_abstract_blocks:
        _clean_paragraph_keep_ppr(paragraphs[41])
        # 将所有英文摘要正文合并为一段
        merged_text = ' '.join(b['text'] for b in en_abstract_blocks)
        inline_runs = parse_inline(merged_text)
        for run_info in inline_runs:
            if not run_info['text']:
                continue
            r = ET.Element(f'{{{W}}}r')
            rpr = ET.SubElement(r, f'{{{W}}}rPr')
            fonts = ET.SubElement(rpr, f'{{{W}}}rFonts')
            fonts.set(f'{{{W}}}ascii', 'Times New Roman')
            fonts.set(f'{{{W}}}eastAsia', 'BatangChe')
            fonts.set(f'{{{W}}}hAnsi', 'Times New Roman')
            b_el = ET.SubElement(rpr, f'{{{W}}}b')
            b_el.set(f'{{{W}}}val', '0')
            bcs = ET.SubElement(rpr, f'{{{W}}}bCs')
            bcs.set(f'{{{W}}}val', '0')
            s = ET.SubElement(rpr, f'{{{W}}}sz')
            s.set(f'{{{W}}}val', '24')
            t_el = ET.SubElement(r, f'{{{W}}}t')
            t_el.text = run_info['text']
            t_el.set(f'{XML_NS}space', 'preserve')
            paragraphs[41].append(r)

    # 替换中文关键词（P36）
    cn_keywords = extract_keywords(cn_abstract_blocks, blocks, chinese=True)
    if cn_keywords:
        replace_keyword_paragraph(paragraphs, 36, cn_keywords, chinese=True)

    # 替换英文关键词（P42）
    en_keywords = extract_keywords(en_abstract_blocks, blocks, chinese=False)
    if en_keywords:
        replace_keyword_paragraph(paragraphs, 42, en_keywords, chinese=False)


def _clean_paragraph_keep_ppr(p):
    """清除段落中所有子元素（保留 pPr），包括 runs、AlternateContent 等。
    同时清除 pPr 中不合法的 rPr 子元素（rPr 只能在 r 里面）。"""
    ppr = p.find(f'{{{W}}}pPr')
    children = list(p)
    for child in children:
        if child is not ppr:
            p.remove(child)
    if ppr is not None:
        # 移除 pPr 中残留的 rPr（模板遗留问题）
        for rpr in ppr.findall(f'{{{W}}}rPr'):
            ppr.remove(rpr)


def _add_jc_to_paragraph(p, jc_val):
    """给段落添加或更新 jc（对齐方式）。"""
    ppr = p.find(f'{{{W}}}pPr')
    if ppr is None:
        ppr = ET.SubElement(p, f'{{{W}}}pPr')
    existing_jc = ppr.find(f'{{{W}}}jc')
    if existing_jc is not None:
        existing_jc.set(f'{{{W}}}val', jc_val)
    else:
        jc = ET.SubElement(ppr, f'{{{W}}}jc')
        jc.set(f'{{{W}}}val', jc_val)


def _replace_abstract_body(paragraphs, start_idx, end_idx, content_blocks, chinese=True,
                           parent_body=None, insert_before_idx=None):
    """替换摘要正文段落，使用模板中的摘要格式（firstLineChars=200, firstLine=480, sz=24）。
    insert_before_idx: 溢出段落插入到此索引的段落之前。"""
    if not content_blocks:
        return

    for i in range(start_idx, end_idx + 1):
        if i < len(paragraphs):
            _clean_paragraph_keep_ppr(paragraphs[i])

    # 将内容写入第一个段落
    p = paragraphs[start_idx]
    first_block = content_blocks[0]
    inline_runs = parse_inline(first_block['text'])
    for run_info in inline_runs:
        if not run_info['text']:
            continue
        if chinese:
            r = make_text_run(run_info['text'],
                              bold=run_info['bold'],
                              italic=run_info['italic'],
                              code=run_info.get('code', False),
                              font_ascii=run_info.get('code', False) and 'Courier New' or 'Times New Roman',
                              sz=24)
        else:
            r = make_text_run(run_info['text'],
                              bold=run_info['bold'],
                              italic=run_info['italic'],
                              code=run_info.get('code', False),
                              font_ascii=run_info.get('code', False) and 'Courier New' or 'Times New Roman',
                              sz=24)
        p.append(r)

    # 后续段落（2~N）写入 P34, P35...
    for block_i, block in enumerate(content_blocks[1:], start=1):
        para_idx = start_idx + block_i
        if para_idx <= end_idx and para_idx < len(paragraphs):
            p = paragraphs[para_idx]
        else:
            # 超出模板段落数，创建新段落插入到 body 中
            p = ET.Element(f'{{{W}}}p')
            ppr = ET.SubElement(p, f'{{{W}}}pPr')
            if not chinese:
                # 英文摘要使用 af3 样式
                pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
                pstyle.set(f'{{{W}}}val', 'af3')
            adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
            adj.set(f'{{{W}}}val', '0')
            snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
            snap.set(f'{{{W}}}val', '0')
            spacing = ET.SubElement(ppr, f'{{{W}}}spacing')
            spacing.set(f'{{{W}}}line', '360')
            spacing.set(f'{{{W}}}lineRule', 'auto')
            ind = ET.SubElement(ppr, f'{{{W}}}ind')
            ind.set(f'{{{W}}}firstLineChars', '200')
            ind.set(f'{{{W}}}firstLine', '480')
            # 插入到关键词段落之前（避免打乱后续段落顺序）
            if parent_body is not None and insert_before_idx is not None:
                before_p = paragraphs[insert_before_idx]
                idx = list(parent_body).index(before_p)
                parent_body.insert(idx, p)
            else:
                paragraphs.append(p)
                continue

        # 添加 runs（使用 pPr 中已有的格式，不再重新设置）
        inline_runs = parse_inline(block['text'])
        for run_info in inline_runs:
            if not run_info['text']:
                continue
            r = make_text_run(run_info['text'],
                              bold=run_info['bold'],
                              italic=run_info['italic'],
                              sz=24)
            p.append(r)


def body_paragraph_builder(block):
    """将 block 转换为正文段落。"""
    p = ET.Element(f'{{{W}}}p')
    ppr = ET.SubElement(p, f'{{{W}}}pPr')
    pstyle = ET.SubElement(ppr, f'{{{W}}}pStyle')
    pstyle.set(f'{{{W}}}val', 'aa')
    adj = ET.SubElement(ppr, f'{{{W}}}adjustRightInd')
    adj.set(f'{{{W}}}val', '0')
    snap = ET.SubElement(ppr, f'{{{W}}}snapToGrid')
    snap.set(f'{{{W}}}val', '0')
    ind = ET.SubElement(ppr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLineChars', '200')
    ind.set(f'{{{W}}}firstLine', '480')
    add_text_to_paragraph(p, block['text'])
    return p


def replace_paragraphs_range(paragraphs, start_idx, end_idx, content_blocks, builder_fn):
    """替换段落范围内的内容。用新块替换 P[start_idx] 到 P[end_idx] 的子 runs。"""
    if not content_blocks:
        return

    # 收集所有要替换的段落
    parents = []
    for i in range(start_idx, end_idx + 1):
        if i < len(paragraphs):
            parents.append(paragraphs[i])

    if not parents:
        return

    # 替换第一个段落的内容
    p = parents[0]
    # 移除所有 runs
    for r in list(p.findall(f'{{{W}}}r')):
        p.remove(r)
    # 添加新内容的 runs
    new_p = builder_fn(content_blocks[0])
    for r in new_p.findall(f'{{{W}}}r'):
        p.append(copy.deepcopy(r))

    # 清空多余的段落
    for i in range(1, len(parents)):
        p_empty = parents[i]
        for r in list(p_empty.findall(f'{{{W}}}r')):
            p_empty.remove(r)


def extract_keywords(abstract_blocks, all_blocks, chinese=True):
    """从所有 blocks 中提取关键词。"""
    for block in all_blocks:
        if block['type'] != 'paragraph':
            continue
        text = block['text']
        if chinese and '关键词' in text:
            # 提取 "关键词：xxx；yyy" 中的内容
            match = re.search(r'关键词[：:]\s*(.+)', text)
            if match:
                raw = match.group(1).strip().rstrip('；;')
                cleaned = strip_markdown_formatting(raw)
                # 清除可能残留的前后 ** 标记
                cleaned = cleaned.strip('*').strip()
                return cleaned
        elif not chinese and ('Key words' in text or 'Keywords' in text.lower()):
            match = re.search(r'(?:Key\s*words|Keywords)[：:]\s*(.+)', text, re.IGNORECASE)
            if match:
                raw = match.group(1).strip().rstrip(';')
                cleaned = strip_markdown_formatting(raw)
                cleaned = cleaned.strip('*').strip()
                return cleaned
    return None


def replace_keyword_paragraph(paragraphs, p_idx, keywords_text, chinese=True):
    """替换关键词段落。如果段落已被清理则重建完整内容，否则替换冒号后的内容。"""
    if p_idx >= len(paragraphs):
        return

    p = paragraphs[p_idx]
    runs = p.findall(f'{{{W}}}r')

    # 清除所有 runs 并重建
    for r in list(runs):
        p.remove(r)

    label = '关键词：' if chinese else 'Key words: '
    label_run = make_text_run(label, font_ascii='Times New Roman',
                               font_east_asia='黑体' if chinese else None,
                               bold=not chinese, sz=24)
    p.append(label_run)

    # 用分号分隔的关键词列表
    if chinese:
        kw_parts = re.split(r'[；;]', keywords_text)
    else:
        kw_parts = re.split(r'[;,]', keywords_text)

    for i, kw in enumerate(kw_parts):
        kw = kw.strip()
        if not kw:
            continue
        if i > 0:
            sep_r = make_text_run('；' if chinese else '; ', font_ascii='Times New Roman', sz=24,
                                  unbold=(not chinese))
            p.append(sep_r)
        new_r = make_text_run(kw, font_ascii='Times New Roman', sz=24, unbold=(not chinese))
        p.append(new_r)


# ============================================================
# 9. 正文段落提取（区分标题和正文）
# ============================================================

def extract_body_blocks(blocks):
    """
    从 blocks 中提取正文部分（"1 绪论"开始到末尾）。
    跳过封面、声明、摘要等前置部分。
    """
    body_blocks = []
    in_body = False

    for block in blocks:
        if block['type'] == 'heading':
            text = strip_markdown_formatting(block['text'])
            # 正文从 "1 绪论" 或第一个带数字编号的 H1 开始
            if re.match(r'^1\s', text) or re.match(r'^第[一二三四五六七八九十]', text):
                in_body = True
            # 也检测 "结论"、"致谢"、"参考文献"、"附录" 等 H1
            if text in ('结论', '致谢', '参考文献', '附录'):
                in_body = True

        if in_body:
            body_blocks.append(block)

    return body_blocks


# ============================================================
# 10. 主函数
# ============================================================

import argparse

def main():
    global _bookmark_counter
    register_namespaces()

    parser = argparse.ArgumentParser(description='将 Markdown 毕业论文转换为 Word 文档')
    parser.add_argument('input', nargs='?', default='final_paper.md',
                        help='输入的 Markdown 文件路径 (默认: final_paper.md)')
    args = parser.parse_args()

    project_dir = os.path.dirname(os.path.abspath(__file__))
    template_dir = os.path.join(project_dir, 'template')
    source_xml = os.path.join(template_dir, 'word', 'document.xml')
    output_docx = os.path.join(project_dir, '毕业论文_生成.docx')
    md_file = os.path.abspath(args.input)

    # 读取 Markdown
    with open(md_file, 'r', encoding='utf-8') as f:
        md_text = f.read()

    # 解析 Markdown
    blocks = parse_markdown(md_text)

    # 提取封面信息
    cover_data = parse_cover_table(md_text)
    print(f'封面信息: {cover_data}')

    # 解析 document.xml（从解压的毕业论文模板 XML 中读取）
    tree = ET.parse(source_xml)
    root = tree.getroot()
    body = root.find(f'{{{W}}}body')

    paragraphs = body.findall(f'{{{W}}}p')
    print(f'模板段落总数: {len(paragraphs)}')

    # 清除旧 _Toc bookmarks
    remove_old_bookmarks(body)
    max_existing_id = find_max_bookmark_id(body)
    _bookmark_counter = max_existing_id + 1
    print(f'Bookmark 起始 ID: {_bookmark_counter}')

    # === 1. 注入封面信息 ===
    inject_cover_data(paragraphs, cover_data)
    print('封面信息已注入')

    # === 2. 替换摘要（含清理格式说明文字，修正段落索引） ===
    current_children = list(body)
    replace_abstracts(paragraphs, blocks, parent_body=body)
    print('摘要已替换')

    # === 3. 生成 TOC 条目 ===
    body_blocks = extract_body_blocks(blocks)
    toc_entries = build_toc_entries(body_blocks)
    print(f'TOC 条目数: {len(toc_entries)}')

    # === 4. 替换 TOC 标题和条目 (P44=TOC标题, P45+=TOC条目) ===
    # 清理 TOC 标题段落 (P44)：只保留 "目录" 文字，清除格式说明
    toc_title_p = paragraphs[44]
    _clean_paragraph_keep_ppr(toc_title_p)
    new_title_r = make_text_run('目录', bold=True, font_east_asia='黑体', sz=36)
    toc_title_p.append(new_title_r)

    # 重新获取当前 children
    current_children = list(body)
    toc_title_idx = current_children.index(toc_title_p)

    # 移除旧的 TOC 条目（TOC标题之后到正文分节符之前的所有元素）
    toc_remove = []
    body_sect_break_p = None
    for i in range(toc_title_idx + 1, len(current_children)):
        elem = current_children[i]
        tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag == 'sectPr' or tag == 'p':
            # 检查是否是正文分节符（P68 附近）
            if tag == 'p':
                ppr = elem.find(f'{{{W}}}pPr')
                if ppr is not None:
                    sect = ppr.find(f'{{{W}}}sectPr')
                    if sect is not None:
                        body_sect_break_p = elem
                        break
        if elem is not toc_title_p:
            toc_remove.append(elem)

    for elem in toc_remove:
        if elem in list(body):
            body.remove(elem)

    # 重新获取 children
    current_children = list(body)
    toc_title_idx = current_children.index(toc_title_p)

    # 插入新 TOC 条目
    for j, entry in enumerate(toc_entries):
        toc_p = build_toc_entry(entry['number'], entry['title'], entry['toc_level'], entry['bookmark'])
        body.insert(toc_title_idx + 1 + j, toc_p)

    print('TOC 已替换')

    # === 5. 替换正文内容 ===
    # 找到正文分节符（最后一个 sectPr 段落中的嵌套分节符）
    # 在 P68 附近查找包含 sectPr 的段落
    current_children = list(body)
    body_sect_idx = len(current_children) - 1
    for i, child in enumerate(current_children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            ppr = child.find(f'{{{W}}}pPr')
            if ppr is not None:
                sect = ppr.find(f'{{{W}}}sectPr')
                if sect is not None:
                    # 检查 sectPr 的 type 是否为 nextPage
                    sect_type = sect.get(f'{{{W}}}type', '')
                    # 取最后一个分节符作为正文开始点
                    body_sect_idx = list(body).index(child)

    # 找到最后一个 body > sectPr
    last_sectPr = body.find(f'{{{W}}}sectPr')

    # 移除正文分节符之后、last_sectPr 之前的所有元素
    current_children = list(body)
    to_remove = []
    for i, child in enumerate(current_children):
        if child is last_sectPr:
            continue
        if i > body_sect_idx:
            to_remove.append(child)

    for elem in to_remove:
        body.remove(elem)

    # 用于跟踪新增的图片/公式媒体文件和关系
    media_files_dict = {}
    rels_counter = {'count': 0, '_rels_entries': []}

    # 插入新的正文内容
    first_h1_done = False
    for block in body_blocks:
        btype = block['type']
        if btype == 'heading':
            # H1 之前插入分页符（第一个 H1 不需要，前面已有分节符）
            if block['level'] == 1 and first_h1_done:
                p_break = build_page_break_paragraph()
                body.append(p_break)
            if block['level'] == 1:
                first_h1_done = True
            p = build_heading(block,
                              bookmark_name=block.get('_bookmark'),
                              bookmark_id=block.get('_bookmark_id'))
            body.append(p)
            # 一级标题后下空一行（模板 P89: style=aa, jc=center）
            if block['level'] == 1:
                p_blank = ET.Element(f'{{{W}}}p')
                p_blank_ppr = ET.SubElement(p_blank, f'{{{W}}}pPr')
                p_blank_style = ET.SubElement(p_blank_ppr, f'{{{W}}}pStyle')
                p_blank_style.set(f'{{{W}}}val', 'aa')
                p_blank_adj = ET.SubElement(p_blank_ppr, f'{{{W}}}adjustRightInd')
                p_blank_adj.set(f'{{{W}}}val', '0')
                p_blank_snap = ET.SubElement(p_blank_ppr, f'{{{W}}}snapToGrid')
                p_blank_snap.set(f'{{{W}}}val', '0')
                p_blank_jc = ET.SubElement(p_blank_ppr, f'{{{W}}}jc')
                p_blank_jc.set(f'{{{W}}}val', 'center')
                body.append(p_blank)
        elif btype == 'paragraph':
            p = build_body_paragraph(block)
            body.append(p)
        elif btype == 'bullet':
            p = build_bullet_paragraph(block)
            body.append(p)
        elif btype == 'numbered':
            p = build_numbered_paragraph(block)
            body.append(p)
        elif btype == 'blockquote':
            p = build_blockquote(block)
            body.append(p)
        elif btype == 'codeblock':
            code_paras = build_codeblock(block)
            if isinstance(code_paras, list):
                for cp in code_paras:
                    body.append(cp)
            else:
                body.append(code_paras)
        elif btype == 'formula':
            # 公式渲染为图片，返回段落列表
            formula_paras = build_formula_paragraph(block, media_files_dict, rels_counter)
            if isinstance(formula_paras, list):
                for fp in formula_paras:
                    body.append(fp)
            else:
                body.append(formula_paras)
        elif btype == 'image':
            # 下载图片并嵌入
            image_paras = build_image_block(block, media_files_dict, rels_counter)
            if isinstance(image_paras, list):
                for ip in image_paras:
                    body.append(ip)
            else:
                body.append(image_paras)
        elif btype == 'table':
            tbl = build_table(block)
            body.append(tbl)
        # hr blocks 已在 parse_markdown 中跳过
        else:
            continue

    # 确保 last_sectPr 在最后
    if last_sectPr is not None:
        if last_sectPr in list(body):
            body.remove(last_sectPr)
        body.append(last_sectPr)

    print(f'正文内容已替换 ({len(body_blocks)} 个块)')

    # === 6. 输出 docx ===
    output_xml = os.path.join(template_dir, 'word', 'document_new.xml')
    tree.write(output_xml, encoding='UTF-8', xml_declaration=True)

    # 从解压的文件重新构建 docx ZIP 容器
    docx_files = {
        '[Content_Types].xml': '[Content_Types].xml',
        '_rels/.rels': '_rels/.rels',
        'customXml/item1.xml': 'customXml/item1.xml',
        'customXml/itemProps1.xml': 'customXml/itemProps1.xml',
        'customXml/_rels/item1.xml.rels': 'customXml/_rels/item1.xml.rels',
        'docProps/app.xml': 'docProps/app.xml',
        'docProps/core.xml': 'docProps/core.xml',
        'docProps/custom.xml': 'docProps/custom.xml',
        'word/document.xml': output_xml,
        'word/_rels/document.xml.rels': 'word/_rels/document.xml.rels',
        'word/endnotes.xml': 'word/endnotes.xml',
        'word/fontTable.xml': 'word/fontTable.xml',
        'word/footer1.xml': 'word/footer1.xml',
        'word/footer2.xml': 'word/footer2.xml',
        'word/footer3.xml': 'word/footer3.xml',
        'word/footnotes.xml': 'word/footnotes.xml',
        'word/header1.xml': 'word/header1.xml',
        'word/header2.xml': 'word/header2.xml',
        'word/header3.xml': 'word/header3.xml',
        'word/header4.xml': 'word/header4.xml',
        'word/header5.xml': 'word/header5.xml',
        'word/header6.xml': 'word/header6.xml',
        'word/header7.xml': 'word/header7.xml',
        'word/_rels/header3.xml.rels': 'word/_rels/header3.xml.rels',
        'word/_rels/header5.xml.rels': 'word/_rels/header5.xml.rels',
        'word/_rels/header6.xml.rels': 'word/_rels/header6.xml.rels',
        'word/_rels/header7.xml.rels': 'word/_rels/header7.xml.rels',
        'word/media/image1.jpeg': 'word/media/image1.jpeg',
        'word/media/image2.png': 'word/media/image2.png',
        'word/media/image3.png': 'word/media/image3.png',
        'word/media/image4.emf': 'word/media/image4.emf',
        'word/media/image5.png': 'word/media/image5.png',
        'word/media/image6.wmf': 'word/media/image6.wmf',
        'word/numbering.xml': 'word/numbering.xml',
        'word/settings.xml': 'word/settings.xml',
        'word/styles.xml': 'word/styles.xml',
        'word/theme/theme1.xml': 'word/theme/theme1.xml',
        'word/webSettings.xml': 'word/webSettings.xml',
        'word/webextensions/taskpanes.xml': 'word/webextensions/taskpanes.xml',
        'word/webextensions/_rels/taskpanes.xml.rels': 'word/webextensions/_rels/taskpanes.xml.rels',
        'word/webextensions/webextension1.xml': 'word/webextensions/webextension1.xml',
        'word/embeddings/oleObject1.bin': 'word/embeddings/oleObject1.bin',
        'word/embeddings/oleObject2.bin': 'word/embeddings/oleObject2.bin',
    }

    # 添加动态生成的图片/公式媒体文件
    for zip_path, local_path in media_files_dict.items():
        if zip_path not in docx_files:
            docx_files[zip_path] = local_path

    # 生成更新后的 document.xml.rels（添加新图片关系）
    _update_rels_file(template_dir, media_files_dict, rels_counter)

    with zipfile.ZipFile(output_docx, 'w', zipfile.ZIP_DEFLATED) as zf_out:
        for zip_path, local_path in docx_files.items():
            full_path = os.path.join(template_dir, local_path)
            if os.path.exists(full_path):
                data = open(full_path, 'rb').read()
                if zip_path == 'word/settings.xml':
                    try:
                        settings_str = data.decode('utf-8')
                        if 'updateFields' not in settings_str:
                            settings_str = settings_str.replace(
                                '</w:settings>',
                                '<w:updateFields w:val="true"/></w:settings>'
                            )
                        data = settings_str.encode('utf-8')
                    except Exception:
                        pass
                zf_out.writestr(zip_path, data)

    # 清理临时文件
    if os.path.exists(output_xml):
        os.remove(output_xml)

    print(f'\n生成文件: {output_docx}')
    print(f'转换了 {len(body_blocks)} 个正文块')
    print(f'生成了 {len(toc_entries)} 个目录条目')


def _update_rels_file(template_dir, media_files_dict, rels_counter):
    """更新 document.xml.rels，为新增的图片添加关系条目。"""
    rels_path = os.path.join(template_dir, 'word', '_rels', 'document.xml.rels')
    if not os.path.exists(rels_path):
        return

    try:
        with open(rels_path, 'r', encoding='utf-8') as f:
            rels_content = f.read()

        # 移除旧的 rIdImage* 条目（每次重新生成）
        import re as _re
        rels_content = _re.sub(r'<Relationship\s+Id="rIdImage\d+"[^>]*/>\s*', '', rels_content)

        # 重新添加所有图片关系
        rels_entries = rels_counter.get('_rels_entries', [])
        for r_id, target in rels_entries:
            rel_entry = (f'<Relationship Id="{r_id}" '
                         f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                         f'Target="{target}"/>')
            rels_content = rels_content.replace('</Relationships>', rel_entry + '</Relationships>')

        with open(rels_path, 'w', encoding='utf-8') as f:
            f.write(rels_content)

    except Exception as e:
        print(f'  警告: 更新 rels 文件失败: {e}', file=sys.stderr)


if __name__ == '__main__':
    main()
