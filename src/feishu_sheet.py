# -*- coding: utf-8 -*-
"""直连飞书电子表格取数：把在线表格读成与 xlsx 一样的网格，交给 clean.py 清洗。

浏览器不能直接调飞书（open.feishu.cn 没有 CORS 头，app_secret 更不能进网页），
所以取数只能走服务端：本机（main.py / serve.py 的 /api/data）和腾讯云函数各用一份，
都调这里的 fetch()/build_payload()，清洗口径与本地 xlsx 完全一致。

config.json 里这样配（三段都填好后把 source 改成 "feishu"）：
  "source": "feishu",
  "feishu": {"app_id": "cli_xxx", "app_secret": "xxx", "spreadsheet_token": "xxx"}
spreadsheet_token 直接填表格的浏览器地址整条链接即可，两种都认：
  云文档里的表：https://xxx.feishu.cn/sheets/<spreadsheet_token>?sheet=xxx
  知识库里的表：https://xxx.feishu.cn/wiki/<知识库节点编号>?sheet=xxx
知识库链接里没有表格 token（那串是知识库节点编号），resolve_token 先换出真实 token。

使用前提（一次性，见 部署到腾讯云.txt）：建企业自建应用、开通电子表格读取权限、
把这张表格加该应用为协作者；表在知识库里的还要开通"查看知识库"权限、并把应用共享
进该知识库。表格里要有 日数据汇总 / 分国家汇总 / 分国家时间统计 三张工作表
（与现在的 xlsx 一样），缺哪张就按"缺表"处理。
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import clean
import feishu_api as fa
import make_html

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')

# 单次读取窗口：26 列 x 1000 行（飞书读范围有上限，保守分片；数据量本来就小）
COL_CHUNK = 26
ROW_CHUNK = 1000

# 三张表：fetch 的键 -> (工作表名, 必需列名)
SHEETS = (
    ('daily', clean.DAILY_SHEET, clean.DAILY_NEED),
    ('country', clean.COUNTRY_SHEET, clean.COUNTRY_COLS),
    ('country_time', clean.COUNTRY_TIME_SHEET, clean.COUNTRY_COLS + ['日期']),
)


def source_of(cfg):
    """config.json 的 source 字段：'feishu' 直连飞书；'xlsx'（或空）读本地 data/。"""
    s = str(cfg.get('source') or 'xlsx').strip().lower()
    if s not in ('xlsx', 'feishu'):
        raise RuntimeError('config.json 的 source 只能写 "xlsx" 或 "feishu"（现在写的是 %r）'
                           % cfg.get('source'))
    return s


def creds_of(cfg):
    """从 config 里取飞书数据源三件套，返回 (app_id, app_secret, spreadsheet_token)。

    spreadsheet_token 原样返回（整条链接或纯 token 都行，换算在 fetch 里做）。
    缺了哪项直接报缺什么，别让人对着 401/404 猜。
    """
    f = cfg.get('feishu') or {}
    miss = [k for k in ('app_id', 'app_secret', 'spreadsheet_token') if not f.get(k)]
    if miss:
        raise RuntimeError('config.json 的 feishu 段还缺 %s（填法见 部署到腾讯云.txt）'
                           % '、'.join(miss))
    return f['app_id'], f['app_secret'], f['spreadsheet_token']


def _link_kind(raw):
    """判断填进来的是哪种东西，返回 ('wiki', 知识库节点编号) 或 ('token', 表格 token)。

    整条链接（/sheets/... 或 /wiki/...）与纯 spreadsheet_token 都认。
    """
    s = str(raw or '').strip()
    for kind, marker in (('wiki', '/wiki/'), ('token', '/sheets/')):
        i = s.find(marker)
        if i >= 0:
            return kind, s[i + len(marker):].split('?')[0].split('/')[0].strip()
    return 'token', s


def resolve_token(app_id, app_secret, raw):
    """把填写处原样粘的"表格链接或 token"换算成 spreadsheet_token。

    知识库里的表格链接是 /wiki/<节点编号>，编号不是表格 token——先调 wiki 的
    get_node 换出真实 token（obj_token）再用。前提：应用开通"查看知识库"权限、
    并被共享进该知识库（见 部署到腾讯云.txt，报错提示里也会再说一遍）。
    """
    kind, tok = _link_kind(raw)
    if not tok:
        raise RuntimeError('表格链接/token 是空的（config.json 的 feishu.spreadsheet_token '
                           '或云函数的 SHEET_TOKEN）')
    if kind != 'wiki':
        return tok
    try:
        node = fa.wiki_get_node(app_id, app_secret, tok).get('node') or {}
    except RuntimeError as e:
        raise RuntimeError('%s\n提示：知识库里的表格要先把应用共享进该知识库，'
                           '并给应用开"查看知识库"权限（见 部署到腾讯云.txt）' % e)
    if str(node.get('obj_type')) != 'sheet':
        raise RuntimeError('这条知识库链接指向的是 %s，不是电子表格（要贴表格那条链接）'
                           % (node.get('obj_type') or '未知类型'))
    obj = node.get('obj_token')
    if not obj:
        raise RuntimeError('知识库接口没有返回表格 token：%s' % node)
    return obj


def get_sheets(app_id, app_secret, spreadsheet_token):
    """列出工作表，返回 {表名: {'sheet_id', 'rows', 'cols'}}。"""
    data = fa.call(app_id, app_secret, 'GET',
                   '/sheets/v3/spreadsheets/%s/sheets/query' % spreadsheet_token)
    out = {}
    for s in data.get('sheets') or []:
        grid = s.get('grid_properties') or {}
        out[s.get('title')] = {'sheet_id': s['sheet_id'],
                               'rows': int(grid.get('row_count') or 0),
                               'cols': int(grid.get('column_count') or 0)}
    return out


def _col_letter(n):
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = ''
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def read_grid(app_id, app_secret, spreadsheet_token, meta):
    """读整张工作表，返回网格 list[list]（空格子统一成 ''）。

    按 COL_CHUNK x ROW_CHUNK 的窗口分片读；读到某一窗口整块为空就认为到了尾部、停读
    （表格数据行是连续的，尾部只有空行）。UnformattedValue：数字给 number、日期给
    Excel 序列号，_to_float/_to_day 两种都认。
    """
    out = []
    for c0 in range(0, max(1, meta['cols']), COL_CHUNK):
        c1 = min(c0 + COL_CHUNK, meta['cols'])
        r = 1
        while r <= max(1, meta['rows']):
            r1 = min(r + ROW_CHUNK - 1, meta['rows'])
            rng = '%s!%s%d:%s%d' % (meta['sheet_id'], _col_letter(c0), r,
                                    _col_letter(c1 - 1), r1)
            data = fa.call(app_id, app_secret, 'GET',
                           '/sheets/v2/spreadsheets/%s/values/%s?valueRenderOption=UnformattedValue'
                           % (spreadsheet_token, rng))
            values = (data.get('valueRange') or {}).get('values') or []
            if not any(v not in (None, '') for row in values for v in row):
                break  # 这一窗口整块为空，后面不会再有数据了
            for i, row in enumerate(values):
                ri = r - 1 + i
                while len(out) <= ri:
                    out.append([])
                line = out[ri]
                if len(line) < c0:
                    line.extend([''] * (c0 - len(line)))
                row = ['' if v is None else v for v in row]
                line[c0:c1] = row + [''] * (c1 - c0 - len(row))
            r = r1 + 1
    while out and not any(v not in (None, '') for v in out[-1]):
        out.pop()  # 去掉尾部全空行
    return out


def fetch(app_id, app_secret, spreadsheet_token):
    """读三张表，返回 {'daily'/'country'/'country_time': (headers, rows) 或 None}。

    spreadsheet_token 原样传填写处的链接/token 即可，resolve_token 先换算成真 token。
    表头行的定位与 xlsx 路径共用 clean.locate_table；某张表不在（或列不齐）时对应项为
    None，行为与本地缺文件一致（国家页隐藏 / 不做时间段筛选）。
    """
    spreadsheet_token = resolve_token(app_id, app_secret, spreadsheet_token)
    meta = get_sheets(app_id, app_secret, spreadsheet_token)
    out = {'daily': None, 'country': None, 'country_time': None}
    for key, name, need in SHEETS:
        if name not in meta:
            print('[!] 飞书表格里没有"%s"工作表，相关板块按缺表处理' % name)
            continue
        grid = read_grid(app_id, app_secret, spreadsheet_token, meta[name])
        try:
            out[key] = clean.locate_table(grid, need, name)
        except RuntimeError as e:
            print('[!] 跳过 %s（%s）' % (name, e))
    return out


def build_payload(app_id, app_secret, spreadsheet_token):
    """取数 + 清洗 + 组装看板数据载荷（main / serve / 云函数三处共用同一份口径）。"""
    src = fetch(app_id, app_secret, spreadsheet_token)
    dashboard_rows, robot_rows, _cols = clean.clean(src['daily'])
    country_rows = clean.clean_country(src['country'])
    country_daily_rows = clean.clean_country_daily(src['country_time'])
    return make_html.payload(dashboard_rows, robot_rows, country_rows, country_daily_rows)


if __name__ == '__main__':
    # 只取数 + 打印行数口径，方便和本地 xlsx 跑出来的对一眼；不发任何消息
    with open(CONFIG_PATH, encoding='utf-8') as f:
        cfg = json.load(f)
    app_id, app_secret, sheet_token = creds_of(cfg)
    kind, _tok = _link_kind(sheet_token)
    print('[>] 表格: %s' % ('知识库链接（自动换算真实 token）' if kind == 'wiki' else '普通链接/token'))
    src = fetch(app_id, app_secret, sheet_token)
    for key, name, _need in SHEETS:
        table = src[key]
        print('[%s] %s: %s' % ('OK' if table else '!', name,
                               '%d 行 x %d 列' % (len(table[1]), len(table[0]))
                               if table else '没读到（表格里没有或列不齐）'))
    if src['daily']:
        dashboard_rows, robot_rows, _cols = clean.clean(src['daily'])
        country_rows = clean.clean_country(src['country'])
        country_daily_rows = clean.clean_country_daily(src['country_time'])
        print('[OK] 清洗后: 明细 %d 行, 机器人 %d 行, 国家 %d 行, 国家日数据 %d 行'
              % (len(dashboard_rows), len(robot_rows), len(country_rows), len(country_daily_rows)))
        print('[OK] 消耗合计 %.2f' % sum(r['消耗'] for r in robot_rows))
    else:
        print('[X] 日数据汇总 没读到，先去 部署到腾讯云.txt 检查表格和应用权限')
