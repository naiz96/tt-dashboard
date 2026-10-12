# -*- coding: utf-8 -*-
"""清洗《投放日数据记录表》，产出看板表（明细）与机器人表（按天x优化师聚合）。
另清洗国家汇总（可选），产出国家/地区汇总行。

- 输入：data/ 下的《投放日数据记录表》（TT 数据加工过的多 sheet xlsx，含
  日数据汇总 / 分国家汇总 等工作表）。日数据汇总**全部读入按天合并**，看板因此
  保留历史，所以 data/ 就是数据唯一存放处、别删老文件；分国家汇总取文件名日期最新的
  那一份（一份就是一段完整时段的汇总）。老式的 TT 单表导出不再读取
- 日数据汇总 的粒度是 优化师 x 日期 x 国家组（人工维护，没有账户/系列维度），
  消耗/广告收入/展示/点击/转化 直接取数；首日ROI 和 CPA/CPM/CTR/CVR 一律由
  合计值重算，不读源表里那几列（源表基本没填）
- 输出：本地 output/看板表_YYYYMMDD.xlsx、output/机器人表_YYYYMMDD.xlsx
- 也能不读本地文件：clean()/clean_country()/clean_country_daily() 都接受外部传进来的
  (headers, rows) 表（飞书电子表格直连时由 feishu_sheet.py 提供），清洗口径完全相同
"""
import datetime
import glob
import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlsx_read import read_xlsx, sheet_names
from xlsx_write import write_xlsx

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
OUT_DIR = os.path.join(BASE_DIR, 'output')

DAILY_SHEET = '日数据汇总'
COUNTRY_SHEET = '分国家汇总'
COUNTRY_TIME_SHEET = '分国家时间统计'

# 日数据汇总 里要读的列；CPA/CPM/CTR/CVR/备注 列存在但基本是空的，不读
DAILY_NEED = ['优化师', '日期', '国家地区', '消耗', '广告收入', '展示', '点击', '转化']
DAILY_NUM = ['消耗', '广告收入', '展示', '点击', '转化']

# 看板表列（明细，含全部历史日期）
DASHBOARD_COLS = ['优化师', '按天', '国家地区', '消耗', '广告收入', '首日ROI',
                  '展示', '点击', '转化', 'CPA', 'CPM', 'CTR', 'CVR']
# 机器人表输出列（账户维度没了，只剩 优化师 x 天）
ROBOT_COLS = ['优化师', '按天', '消耗', '首日ROI', '单次转化成本', '转化量',
              'CPM', 'CTR', 'CVR']
# 国家表输出列，与 分国家汇总 的列名一一对应（原样搬运；源表的 CPA/CPM/CTR/CVR 是空的）
COUNTRY_COLS = ['国家地区', '消耗', 'D0 广告收入', 'D0 ROI', '展示', '点击', '转化',
                '视频播放量', '平均视频播放时长', '视频完播率',
                '播放 75% 次数', '播放 50% 次数', '播放 25% 次数']
COUNTRY_NUMERIC = set(COUNTRY_COLS) - {'国家地区'}

# Excel 日期序列号（1900 系列）起点：46287 = 2026-09-22
_EPOCH = datetime.date(1899, 12, 30)


def _name_date(path):
    """从文件名里取日期，用作排序键，取不到返回空串。

    国家汇总 是"截至某日"的时段汇总、没有日期列，只能靠文件名判断哪份更新
    （Amber-投放日数据记录表261008.xlsx）。这样文件被复制/改名、修改时间变了也
    不会选错。支持 20261008 / 261008 / 1008 三种写法，只写月日的用文件修改年份
    补足，跨年也排得对。
    """
    nums = re.findall(r'\d+', os.path.basename(path))
    if not nums:
        return ''
    n = max(nums, key=len)
    if len(n) == 8:
        return n
    if len(n) == 6:
        return '20' + n
    if len(n) == 4:
        return '%04d%s' % (time.localtime(os.path.getmtime(path)).tm_year, n)
    return ''


def _to_float(s):
    try:
        return float(s)
    except (TypeError, ValueError):
        return 0.0


def _to_day(v):
    """日期单元格 -> 'YYYY-MM-DD'。兼容文本日期与 Excel 序列号，认不出返回 None。

    飞书电子表格传回来的日期是序列号数字（float），先 str 一下再解析。
    """
    v = str(v or '').strip()
    if re.match(r'^\d{4}-\d{1,2}-\d{1,2}$', v):
        y, m, d = v.split('-')
        return '%04d-%02d-%02d' % (int(y), int(m), int(d))
    try:
        return (_EPOCH + datetime.timedelta(days=int(float(v)))).isoformat()
    except (TypeError, ValueError):
        return None


def locate_table(grid, need, where=''):
    """在网格里定位表头行（含全部 need 列名的那一行），返回 (headers, rows)。

    表头不一定在第一行：分国家汇总 的头一行是"数据截止：1007"。
    xlsx 与飞书电子表格共用这一条（飞书那边把 values 读成同样的网格）。
    """
    for i, row in enumerate(grid):
        if all(c in row for c in need):
            return row, grid[i + 1:]
    raise RuntimeError('工作表"%s"里没找到含 %s 的表头行' % (where, '、'.join(need)))


def _read_table(path, sheet, need):
    """读本地 xlsx 的工作表并定位表头行，返回 (headers, rows)。"""
    first, rest = read_xlsx(path, sheet)
    return locate_table([first] + rest, need, sheet)


def input_files():
    """data/ 下的《投放日数据记录表》，按修改时间从旧到新。

    只认含 日数据汇总 工作表的文件，老式 TT 单表导出打一行提示后跳过。
    Excel 打开表格时会生成 ~$xxx.xlsx 锁文件，glob 也匹配得到，但读它直接
    PermissionError（文件被 Excel 独占），所以先剔掉。
    """
    files = [f for f in glob.glob(os.path.join(DATA_DIR, '*.xlsx'))
             if not os.path.basename(f).startswith('~$')]
    if not files:
        raise FileNotFoundError('data/ 目录下没有 xlsx 文件，请先把《投放日数据记录表》放进去')
    want = []
    for f in files:
        if DAILY_SHEET in sheet_names(f):
            want.append(f)
        else:
            print('[!] 跳过 %s（不是《投放日数据记录表》，老格式不再读取）' % os.path.basename(f))
    if not want:
        raise FileNotFoundError('data/ 目录下没有《投放日数据记录表》（含"%s"工作表的 xlsx）'
                                % DAILY_SHEET)
    return sorted(want, key=os.path.getmtime)


def _read_daily():
    """读 data/ 下所有《投放日数据记录表》的 日数据汇总，按天合并，返回 (headers, rows)。"""
    parsed = []  # (最大数据日期, 修改时间, 名字, headers, rows)
    for p in input_files():
        headers, rows = _read_table(p, DAILY_SHEET, DAILY_NEED)
        di = headers.index('日期')
        days = [d for d in (_to_day(r[di]) for r in rows) if d]
        parsed.append((max(days) if days else '', os.path.getmtime(p),
                       os.path.basename(p), headers, rows))
    return _merge_daily(parsed)


def _merge_daily(parsed):
    """按天合并多份 日数据汇总（本地多文件与飞书直连共用），返回 (headers, rows)。

    同一天可能出现在多份文件里（后来的人在旧表上接着填），以数据日期更新的那份为准，
    整天替换而不是逐行去重——同一天一个优化师本来就可能有多行（按国家组拆开的），
    逐行去重会把真实数据丢掉。
    rows 已统一到第一份可用文件的表头顺序，源表缺的列填 None。
    """
    # 数据日期靠后的更可信（文件被复制过、时间戳变了也不受影响）；同一天导两次再用修改时间兜底
    parsed.sort(key=lambda it: (it[0], it[1]))

    base = parsed[0][3]
    base_day = base.index('日期')
    base_opt = base.index('优化师')
    per_day = {}
    day_src = {}
    dup_days = set()
    unfilled = 0
    for _maxday, _mtime, name, headers, rows in parsed:
        pick = [headers.index(c) if c in headers else -1 for c in base]
        di, oi = pick[base_day], pick[base_opt]
        by_day = {}
        for r in rows:
            day = _to_day(r[di])
            if not day:
                if (r[oi] or '').strip() not in ('', '合计', '总计'):
                    unfilled += 1  # 表里开了头没填完的行（比如只填了优化师没填日期）
                continue  # 合计行、尾部空行
            by_day.setdefault(day, []).append([r[i] if i >= 0 else None for i in pick])
        for day, day_rows in by_day.items():
            if day in day_src:
                dup_days.add(day)
            day_src[day] = name
            per_day[day] = day_rows  # 整天替换，不是逐行合并
    if unfilled:
        print('[!] %d 行只填了优化师、没填日期（源表里还没填完的行），已跳过' % unfilled)
    if dup_days:  # 上面是整段覆盖，这里补一句提示，方便发现几份文件的范围是不是重叠了
        print('[!] %d 天在多份文件里重复（%s .. %s），已取数据最新的那份'
              % (len(dup_days), min(dup_days), max(dup_days)))

    merged = []
    for day in sorted(per_day):
        merged.extend(per_day[day])
    print('[>] 读到 %d 份《投放日数据记录表》，合并后 %d 天 %d 行'
          % (len(parsed), len(per_day), len(merged)))
    return base, merged


def clean(daily=None):
    """清洗 + 聚合。

    daily: (headers, rows) 形式的 日数据汇总（飞书直连时传进来），缺省从 data/ 的
    xlsx 读并按天合并。
    返回 (dashboard_rows, robot_rows, dashboard_cols)，均为 list[dict] / list[str]。
    """
    headers, rows = daily if daily is not None else _read_daily()
    col_idx = {h: i for i, h in enumerate(headers)}
    ci = {c: col_idx[c] for c in DAILY_NEED}

    dashboard_rows = []
    robot_agg = {}  # (优化师, 按天) -> 累加值
    for r in rows:  # 日期不合法的行已在 _read_daily 里滤掉
        opt = (r[ci['优化师']] or '').strip()
        day = _to_day(r[ci['日期']])
        if not opt or opt in ('合计', '总计'):
            continue
        v = {c: _to_float(r[ci[c]]) for c in DAILY_NUM}
        spend, rev = v['消耗'], v['广告收入']
        impr, clk, conv = v['展示'], v['点击'], v['转化']

        # 看板表行：逐行派生比率（与机器人表的合计口径不同，别混）
        dashboard_rows.append({
            '优化师': opt, '按天': day, '国家地区': (r[ci['国家地区']] or '').strip(),
            '消耗': spend, '广告收入': rev,
            '首日ROI': round(rev / spend, 4) if spend > 0 else 0.0,
            '展示': impr, '点击': clk, '转化': conv,
            'CPA': round(spend / conv, 2) if conv > 0 else 0.0,
            'CPM': round(spend / impr * 1000, 2) if impr > 0 else 0.0,
            'CTR': round(clk / impr, 4) if impr > 0 else 0.0,
            'CVR': round(conv / clk, 4) if clk > 0 else 0.0,
        })

        # 机器人表聚合：同一优化师同一天的各国家组相加
        a = robot_agg.setdefault((opt, day), {'消耗': 0.0, '收入': 0.0, '转化量': 0.0,
                                              '展示量': 0.0, '点击量': 0.0})
        a['消耗'] += spend
        a['收入'] += rev
        a['转化量'] += conv
        a['展示量'] += impr
        a['点击量'] += clk

    robot_rows = []
    for (opt, day), a in sorted(robot_agg.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        spend, conv, impr, clk = a['消耗'], a['转化量'], a['展示量'], a['点击量']
        robot_rows.append({
            '优化师': opt, '按天': day,
            '消耗': round(spend, 2),
            '首日ROI': round(a['收入'] / spend, 4) if spend > 0 else 0.0,
            '单次转化成本': round(spend / conv, 2) if conv > 0 else 0.0,
            '转化量': int(conv),
            'CPM': round(spend / impr * 1000, 2) if impr > 0 else 0.0,
            'CTR': round(clk / impr, 4) if impr > 0 else 0.0,
            'CVR': round(conv / clk, 4) if clk > 0 else 0.0,
            # 展示量/点击量/广告收入供群消息聚合小计/总计用（收入要显示、ROI 要按 Σ收入/Σ消耗 算），
            # 不写入 xlsx
            '展示量': round(impr, 2),
            '点击量': round(clk, 2),
            '广告收入': round(a['收入'], 2),
        })
    return dashboard_rows, robot_rows, DASHBOARD_COLS


def _pick_latest(sheet, need):
    """本地 xlsx：按文件名日期从新到旧，取第一份含该工作表、列齐的文件。

    国家汇总/时间统计 都是"一份就是完整一段"，不做合并。都读不出来返回 None。
    文件名日期优先：老文件被复制过来、修改时间变新了也不会顶掉新数据。
    """
    files = sorted(input_files(),
                   key=lambda p: (_name_date(p), os.path.getmtime(p)), reverse=True)
    for path in files:  # 从新到旧，第一份能读出来的就用它
        try:
            return _read_table(path, sheet, need)
        except (ValueError, RuntimeError) as e:
            print('[!] 跳过 %s（%s）' % (os.path.basename(path), e))
    return None


def clean_country(table=None):
    """清洗国家汇总，产出国家/地区汇总行。

    table: (headers, rows) 形式的 分国家汇总（飞书直连时传进来），缺省按文件名日期取
    data/ 里最新的一份。分国家汇总 本身就是国家维度（没有账户、没有日期），原样搬过来，
    不再自己算 ROAS/CPM/CTR/CVR。没有可用表时返回 []。
    """
    if table is None:
        table = _pick_latest(COUNTRY_SHEET, COUNTRY_COLS)
        if table is None:
            print('[!] 没有可用的国家汇总（分国家汇总），看板不显示国家数据板块')
            return []
    headers, rows = table

    col_idx = {h: i for i, h in enumerate(headers)}
    ci = col_idx['国家地区']
    country_rows = []
    for r in rows:  # 源表已按消耗降序，保持原顺序
        name = (r[ci] or '').strip()
        if not name or name in ('合计', '总计') or name.startswith('总计'):
            continue  # "合计"行与空行
        row = {}
        for c in COUNTRY_COLS:
            v = r[col_idx[c]]
            row[c] = _to_float(v) if c in COUNTRY_NUMERIC else v
        country_rows.append(row)
    return country_rows


def clean_country_daily(table=None):
    """读 分国家时间统计，产出"一天一国一行"的国家数据，供国家页按时间段筛选。

    table: (headers, rows) 形式的分国家时间统计（飞书直连时传进来），缺省按文件名日期
    取 data/ 里最新的一份。列名与 分国家汇总 完全一样，只是多一列日期；日期列只在每天
    第一行写了值（后面的行留空，Excel 里像合并单元格的效果），往下补全。
    没有可用表时返回 []，国家页就只显示全时段。
    """
    if table is None:
        table = _pick_latest(COUNTRY_TIME_SHEET, COUNTRY_COLS + ['日期'])
        if table is None:
            print('[!] 没有可用的分国家时间统计，国家页不做时间段筛选')
            return []
    headers, rows = table

    col_idx = {h: i for i, h in enumerate(headers)}
    ci, di = col_idx['国家地区'], col_idx['日期']
    out, cur = [], None
    for r in rows:
        day = _to_day(r[di])
        if day:
            cur = day
        name = (r[ci] or '').strip()
        if not name or name in ('合计', '总计') or name.startswith('总计') or not cur:
            continue  # "合计"行、空行，以及还没出现日期之前的行
        row = {'按天': cur}
        for c in COUNTRY_COLS:
            v = r[col_idx[c]]
            row[c] = _to_float(v) if c in COUNTRY_NUMERIC else v
        out.append(row)
    return out


def data_tag(robot_rows):
    """数据里最大的日期，YYYYMMDD，备份文件命名用。"""
    dates = sorted({r['按天'] for r in robot_rows})
    if not dates:
        raise RuntimeError('清洗后没有任何数据行')
    return dates[-1].replace('-', '')


def save_backups(dashboard_rows, robot_rows, dashboard_cols):
    """保存本地 xlsx 备份，文件名带数据中最大日期。"""
    os.makedirs(OUT_DIR, exist_ok=True)
    tag = data_tag(robot_rows)
    dash_path = os.path.join(OUT_DIR, '看板表_%s.xlsx' % tag)
    robot_path = os.path.join(OUT_DIR, '机器人表_%s.xlsx' % tag)
    write_xlsx(dash_path, dashboard_cols,
               [[r[c] for c in dashboard_cols] for r in dashboard_rows])
    write_xlsx(robot_path, ROBOT_COLS,
               [[r[c] for c in ROBOT_COLS] for r in robot_rows])
    return dash_path, robot_path


def save_country_backup(country_rows, tag):
    """保存国家表本地备份；无数据返回 None。

    国家汇总没有日期，文件名沿用投放数据的最大日期（与看板表对齐）。
    """
    if not country_rows:
        return None
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, '国家表_%s.xlsx' % tag)
    write_xlsx(path, COUNTRY_COLS, [[r[c] for c in COUNTRY_COLS] for r in country_rows])
    return path


if __name__ == '__main__':
    dash, robot, cols = clean()
    dash_path, robot_path = save_backups(dash, robot, cols)
    print('[OK] 清洗完成: 看板表 %d 行, 机器人表 %d 行' % (len(dash), len(robot)))
    print('    %s' % dash_path)
    print('    %s' % robot_path)
    country = clean_country()
    if country:
        path = save_country_backup(country, data_tag(robot))
        print('[OK] 国家表 %d 行' % len(country))
        print('    %s' % path)
    else:
        print('[!] data/ 下没有可用的国家汇总，跳过')
    daily = clean_country_daily()
    if daily:
        print('[OK] 分国家时间统计 %d 行（国家页按天/周/月筛选用）' % len(daily))
