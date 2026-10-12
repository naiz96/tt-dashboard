# -*- coding: utf-8 -*-
"""零依赖 xlsx 读取器（仅标准库 zipfile + xml.etree）。

按工作表名读取，返回表头与数据行。
"""
import re
import zipfile
import xml.etree.ElementTree as ET

_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_REL_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
_COL_RE = re.compile(r'([A-Z]+)')


def _col_index(col):
    """列字母 -> 0 基索引，如 A->0, AA->26。"""
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n - 1


def _read_shared_strings(z):
    try:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    return [''.join(t.text or '' for t in si.iter(_NS + 't')) for si in root]


def _cell_value(c, shared):
    """单元格 -> 字符串值。inlineStr / sharedString / 数值统一转 str。"""
    t = c.get('t')
    v = c.find(_NS + 'v')
    if t == 'inlineStr':
        is_el = c.find(_NS + 'is')
        if is_el is not None:
            return ''.join(node.text or '' for node in is_el.iter(_NS + 't'))
        return ''
    if t == 's':
        if v is not None and v.text:
            idx = int(v.text)
            return shared[idx] if idx < len(shared) else ''
        return ''
    # 数值 / 布尔 / 公式结果等，保留原始文本形式
    return v.text if v is not None and v.text else ''


def _sheet_pairs(z):
    """[(工作表名, zip 内路径)]，按工作簿里的顺序。

    不能写死 xl/worksheets/sheet1.xml：工作表的文件名不一定是 sheet1，
    甚至可能没有这张（TT 后台加工过的多 sheet 文件就是 sheet2 起）。
    """
    try:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    except KeyError:
        return []
    rid2target = {r.get('Id'): r.get('Target') for r in rels}
    pairs = []
    for sh in wb.iter(_NS + 'sheet'):
        target = rid2target.get(sh.get(_REL_NS + 'id'))
        if target:
            pairs.append((sh.get('name') or '', _norm_target(target)))
    return pairs


def _norm_target(target):
    """关系里的 Target -> zip 内路径（可能相对 xl/，也可能带前导斜杠）。"""
    t = target.lstrip('/')
    return t if t.startswith('xl/') else 'xl/' + t


def sheet_names(path):
    """xlsx 里的工作表名，按工作簿顺序。"""
    with zipfile.ZipFile(path) as z:
        return [name for name, _t in _sheet_pairs(z)]


def read_xlsx(path, sheet=None):
    """读取 xlsx 的一张工作表（缺省第一张）。

    返回 (headers, rows)：
    - headers: 第一行各单元格字符串，按列展开到最大列
    - rows: 其余行，每行与 headers 等长，空单元格为 ''
    """
    with zipfile.ZipFile(path) as z:
        shared = _read_shared_strings(z)
        pairs = _sheet_pairs(z)
        if not pairs:
            target = 'xl/worksheets/sheet1.xml'
        elif sheet is None:
            target = pairs[0][1]
        else:
            names = dict(pairs)
            if sheet not in names:
                raise ValueError('xlsx 里没有工作表"%s"（有：%s）'
                                 % (sheet, '、'.join(n for n, _t in pairs)))
            target = names[sheet]
        root = ET.fromstring(z.read(target))
    parsed = {}  # 行号 -> {列索引: 值}
    max_col = -1
    for row_el in root.iter(_NS + 'row'):
        rn = int(row_el.get('r'))
        cells = {}
        for c in row_el.iter(_NS + 'c'):
            ref = c.get('r')
            if not ref:
                continue
            col_i = _col_index(_COL_RE.match(ref).group(1))
            cells[col_i] = _cell_value(c, shared)
            if col_i > max_col:
                max_col = col_i
        parsed[rn] = cells
    if not parsed:
        return [], []
    rn_list = sorted(parsed)
    width = max_col + 1
    headers = [parsed[rn_list[0]].get(i, '') for i in range(width)]
    rows = [[parsed[rn].get(i, '') for i in range(width)] for rn in rn_list[1:]]
    return headers, rows
