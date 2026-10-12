# -*- coding: utf-8 -*-
"""零依赖 xlsx 写入器：生成最小 Office Open XML 工作簿。

字符串写为 inlineStr 单元格，float/int 写为数值单元格，
自动按内容估算列宽。
"""
import zipfile

_XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
_REL_DOC_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def _esc(s):
    return (s.replace('&', '&amp;').replace('<', '&lt;')
             .replace('>', '&gt;').replace('"', '&quot;'))


def _col_name(i):
    """0 基列索引 -> 列字母。"""
    name = ''
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        name = chr(65 + r) + name
    return name


def _cell_xml(ref, value):
    if value is None or value == '':
        return ''
    if isinstance(value, bool):
        return '<c r="%s" t="b"><v>%d</v></c>' % (ref, int(value))
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value == int(value):
            v = str(int(value))
        else:
            v = repr(value)
        return '<c r="%s"><v>%s</v></c>' % (ref, v)
    return '<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, _esc(str(value)))


def _display_width(s):
    """显示宽度：非 ASCII 计 2，其余计 1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(s))


def _auto_width(column_values, header):
    w = max([_display_width(v) for v in column_values] + [_display_width(header) or 0])
    return max(8, min(w + 2, 42))


def write_xlsx(path, headers, rows):
    """生成 xlsx。

    headers: list[str]
    rows: list[list]，值与 headers 对齐，支持 str / float / int / None
    """
    n_cols = len(headers)
    widths = []
    for i in range(n_cols):
        col_vals = [row[i] if i < len(row) else '' for row in rows]
        widths.append(_auto_width(col_vals, headers[i]))

    sheet_rows = []
    cells = ''.join(_cell_xml('%s1' % _col_name(i), h) for i, h in enumerate(headers))
    sheet_rows.append('<row r="1">%s</row>' % cells)
    for ri, row in enumerate(rows, start=2):
        cells = ''.join(_cell_xml('%s%d' % (_col_name(i), ri), v)
                        for i, v in enumerate(row))
        sheet_rows.append('<row r="%d">%s</row>' % (ri, cells))

    cols = ''.join('<col min="%d" max="%d" width="%.1f" customWidth="1"/>'
                   % (i + 1, i + 1, w) for i, w in enumerate(widths))
    sheet = (_XML_HEADER
             + '<worksheet xmlns="%s" xmlns:r="%s">'
               '<cols>%s</cols><sheetData>%s</sheetData></worksheet>'
             % (_MAIN_NS, _REL_DOC_NS, cols, ''.join(sheet_rows)))

    workbook = (_XML_HEADER
                + '<workbook xmlns="%s" xmlns:r="%s">'
                  '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
                  '</workbook>' % (_MAIN_NS, _REL_DOC_NS))
    wb_rels = (_XML_HEADER
               + '<Relationships xmlns="%s">'
                 '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>'
                 '</Relationships>' % (_REL_NS, _REL_DOC_NS))
    root_rels = (_XML_HEADER
                 + '<Relationships xmlns="%s">'
                   '<Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>' % (_REL_NS, _REL_DOC_NS))
    content_types = (_XML_HEADER
                     + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                       '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                       '<Default Extension="xml" ContentType="application/xml"/>'
                       '<Override PartName="/xl/workbook.xml" '
                       'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                       '<Override PartName="/xl/worksheets/sheet1.xml" '
                       'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                       '</Types>')

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', workbook)
        z.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        z.writestr('xl/worksheets/sheet1.xml', sheet)
