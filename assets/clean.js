// src/clean.py 的 JS 版：《投放日数据记录表》 -> 看板要的三张表。
//
// 口径和 Python 版逐条对齐（按天整段合并、比率一律由合计值重算、round 位数、
// 行序、int(转化量) 截断），改这里就要同步改 src/clean.py。
// 依赖同目录的 xlsx_read.js（root.TTSheet）。
(function (root) {
  'use strict';

  const DAILY_SHEET = '日数据汇总';
  const COUNTRY_SHEET = '分国家汇总';
  const COUNTRY_TIME_SHEET = '分国家时间统计';

  // 日数据汇总 里要读的列；CPA/CPM/CTR/CVR/备注 列存在但基本是空的，不读
  const DAILY_NEED = ['优化师', '日期', '国家地区', '消耗', '广告收入', '展示', '点击', '转化'];
  const DAILY_NUM = ['消耗', '广告收入', '展示', '点击', '转化'];

  const DASHBOARD_COLS = ['优化师', '按天', '国家地区', '消耗', '广告收入', '首日ROI',
                          '展示', '点击', '转化', 'CPA', 'CPM', 'CTR', 'CVR'];
  const COUNTRY_COLS = ['国家地区', '消耗', 'D0 广告收入', 'D0 ROI', '展示', '点击', '转化',
                        '视频播放量', '平均视频播放时长', '视频完播率',
                        '播放 75% 次数', '播放 50% 次数', '播放 25% 次数'];
  const COUNTRY_NUMERIC = COUNTRY_COLS.filter(c => c !== '国家地区');

  // Excel 日期序列号（1900 系列）起点
  const EPOCH_UTC = Date.UTC(1899, 11, 30);

  function toFloat(v) {
    const s = String(v == null ? '' : v).trim();
    if (!s) return 0;
    const n = Number(s);
    return isNaN(n) ? 0 : n;
  }

  function round(v, n) {
    const m = Math.pow(10, n);
    return Math.round(v * m) / m;
  }

  function text(v) {
    return String(v == null ? '' : v).trim();
  }

  // 日期单元格 -> 'YYYY-MM-DD'。兼容文本日期与 Excel 序列号，认不出返回 null
  function toDay(v) {
    const s = String(v == null ? '' : v).trim();
    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) {
      const p = s.split('-');
      return p[0].padStart(4, '0') + '-' + p[1].padStart(2, '0') + '-' + p[2].padStart(2, '0');
    }
    const n = Number(s);
    if (!s || isNaN(n)) return null;
    return new Date(EPOCH_UTC + Math.trunc(n) * 86400000).toISOString().slice(0, 10);
  }

  // 从文件名里取日期，用作排序键，取不到返回空串。
  // 支持 20261008 / 261008 / 1008 三种写法，只写月日的用文件修改年份补足
  function nameDate(name, mtime) {
    const nums = String(name == null ? '' : name).match(/\d+/g) || [];
    if (!nums.length) return '';
    let n = nums[0];
    for (const x of nums) if (x.length > n.length) n = x;
    if (n.length === 8) return n;
    if (n.length === 6) return '20' + n;
    if (n.length === 4) return String(new Date(mtime || 0).getFullYear()) + n;
    return '';
  }

  // 读工作表并定位表头行（含全部 need 列名的那一行）。
  // 表头不一定在第一行：分国家汇总 的头一行是"数据截止：1007"
  async function readTable(bytes, sheet, need) {
    const t = await root.TTSheet.readXlsx(bytes, sheet);
    const grid = [t.headers].concat(t.rows);
    for (let i = 0; i < grid.length; i++) {
      if (need.every(c => grid[i].indexOf(c) >= 0)) {
        return { headers: grid[i], rows: grid.slice(i + 1) };
      }
    }
    throw new Error('工作表"' + sheet + '"里没找到含 ' + need.join('、') + ' 的表头行');
  }

  // 读所有文件的 日数据汇总，按天合并。
  // 同一天可能出现在多份文件里，以数据日期更新的那份为准，整天替换而不是逐行去重
  async function readDaily(files) {
    const parsed = [];
    for (const f of files) {
      const t = await readTable(f.bytes, DAILY_SHEET, DAILY_NEED);
      const di = t.headers.indexOf('日期');
      const days = t.rows.map(r => toDay(r[di])).filter(Boolean).sort();
      parsed.push({
        maxday: days.length ? days[days.length - 1] : '',
        mtime: f.lastModified || 0, name: f.name,
        headers: t.headers, rows: t.rows,
      });
    }
    // 数据日期靠后的更可信；同一天导两次再用修改时间兜底
    parsed.sort((a, b) => (a.maxday === b.maxday ? a.mtime - b.mtime
                                                 : (a.maxday < b.maxday ? -1 : 1)));

    const base = parsed[0].headers;
    const baseDay = base.indexOf('日期'), baseOpt = base.indexOf('优化师');
    const perDay = {}, daySrc = {}, dupDays = new Set();
    let unfilled = 0;
    for (const it of parsed) {
      const pick = base.map(c => it.headers.indexOf(c));
      const di = pick[baseDay], oi = pick[baseOpt];
      const byDay = {};
      for (const r of it.rows) {
        const day = toDay(r[di]);
        if (!day) {
          const opt = text(r[oi]);   // 表里开了头没填完的行（比如只填了优化师没填日期）
          if (opt && opt !== '合计' && opt !== '总计') unfilled++;
          continue;                  // 合计行、尾部空行
        }
        if (!byDay[day]) byDay[day] = [];
        byDay[day].push(pick.map(i => (i >= 0 ? r[i] : null)));
      }
      for (const day of Object.keys(byDay)) {   // 整天替换，不是逐行合并
        if (daySrc[day] !== undefined) dupDays.add(day);
        daySrc[day] = it.name;
        perDay[day] = byDay[day];
      }
    }
    const merged = [];
    for (const day of Object.keys(perDay).sort()) merged.push(...perDay[day]);
    return {
      headers: base, rows: merged,
      meta: {files: parsed.length, days: Object.keys(perDay).length, rows: merged.length,
             unfilled: unfilled, dupDays: [...dupDays].sort()},
    };
  }

  function clean(daily) {
    const headers = daily.headers;
    const ci = {};
    DAILY_NEED.forEach(c => { ci[c] = headers.indexOf(c); });

    const detail = [], agg = {};
    for (const r of daily.rows) {
      const opt = text(r[ci['优化师']]);
      const day = toDay(r[ci['日期']]);
      if (!opt || opt === '合计' || opt === '总计') continue;
      const v = {};
      DAILY_NUM.forEach(c => { v[c] = toFloat(r[ci[c]]); });
      const spend = v['消耗'], rev = v['广告收入'];
      const impr = v['展示'], clk = v['点击'], conv = v['转化'];

      // 看板表行：逐行派生比率
      detail.push({
        '优化师': opt, '按天': day, '国家地区': text(r[ci['国家地区']]),
        '消耗': spend, '广告收入': rev,
        '首日ROI': spend > 0 ? round(rev / spend, 4) : 0,
        '展示': impr, '点击': clk, '转化': conv,
        'CPA': conv > 0 ? round(spend / conv, 2) : 0,
        'CPM': impr > 0 ? round(spend / impr * 1000, 2) : 0,
        'CTR': impr > 0 ? round(clk / impr, 4) : 0,
        'CVR': clk > 0 ? round(conv / clk, 4) : 0,
      });

      // 机器人表聚合：同一优化师同一天的各国家组相加
      const key = opt + '\u0000' + day;
      const a = agg[key] || (agg[key] = {'优化师': opt, '按天': day,
                                         '消耗': 0, '收入': 0, '转化量': 0, '展示量': 0, '点击量': 0});
      a['消耗'] += spend;
      a['收入'] += rev;
      a['转化量'] += conv;
      a['展示量'] += impr;
      a['点击量'] += clk;
    }

    const robot = Object.keys(agg).map(k => agg[k]).sort((x, y) => {
      if (x['按天'] !== y['按天']) return x['按天'] < y['按天'] ? -1 : 1;
      return x['优化师'] < y['优化师'] ? -1 : (x['优化师'] > y['优化师'] ? 1 : 0);
    }).map(a => {
      const spend = a['消耗'], conv = a['转化量'], impr = a['展示量'], clk = a['点击量'];
      return {
        '优化师': a['优化师'], '按天': a['按天'],
        '消耗': round(spend, 2),
        '首日ROI': spend > 0 ? round(a['收入'] / spend, 4) : 0,
        '单次转化成本': conv > 0 ? round(spend / conv, 2) : 0,
        '转化量': Math.trunc(conv),
        'CPM': impr > 0 ? round(spend / impr * 1000, 2) : 0,
        'CTR': impr > 0 ? round(clk / impr, 4) : 0,
        'CVR': clk > 0 ? round(conv / clk, 4) : 0,
        // 展示量/点击量/广告收入供前端重算 CPM/CTR/CVR 与群消息聚合用，页面上不显示
        '展示量': round(impr, 2),
        '点击量': round(clk, 2),
        '广告收入': round(a['收入'], 2),
      };
    });
    return {detail: detail, robot: robot};
  }

  // 分国家汇总 原样搬运（一国一行、没有日期），按文件名日期取最新的一份可用文件
  async function cleanCountry(files) {
    const sorted = files.slice().sort((a, b) => {
      const da = nameDate(a.name, a.lastModified), db = nameDate(b.name, b.lastModified);
      if (da !== db) return da < db ? 1 : -1;
      return (b.lastModified || 0) - (a.lastModified || 0);
    });
    let table = null;
    for (const f of sorted) {          // 从新到旧，第一份能读出来的就用它
      try { table = await readTable(f.bytes, COUNTRY_SHEET, COUNTRY_COLS); break; }
      catch (e) { /* 这份没有国家汇总，换下一份 */ }
    }
    if (!table) return [];
    const colIdx = {};
    table.headers.forEach((h, i) => { colIdx[h] = i; });
    const ci = colIdx['国家地区'];
    const out = [];
    for (const r of table.rows) {      // 源表已按消耗降序，保持原顺序
      const name = text(r[ci]);
      if (!name || name === '合计' || name === '总计' || name.indexOf('总计') === 0) continue;
      const row = {};
      for (const c of COUNTRY_COLS) {
        row[c] = COUNTRY_NUMERIC.indexOf(c) >= 0 ? toFloat(r[colIdx[c]]) : r[colIdx[c]];
      }
      out.push(row);
    }
    return out;
  }

  // 分国家时间统计：一天一国一行（列名与 分国家汇总 一样，只是多一列日期）。
  // 日期列只在每天第一行有值（Excel 里像合并单元格），往下补全。
  async function cleanCountryDaily(files) {
    const sorted = files.slice().sort((a, b) => {
      const da = nameDate(a.name, a.lastModified), db = nameDate(b.name, b.lastModified);
      if (da !== db) return da < db ? 1 : -1;
      return (b.lastModified || 0) - (a.lastModified || 0);
    });
    let table = null;
    for (const f of sorted) {
      try { table = await readTable(f.bytes, COUNTRY_TIME_SHEET, COUNTRY_COLS.concat(['日期'])); break; }
      catch (e) { /* 这份没有时间统计，换下一份 */ }
    }
    if (!table) return [];
    const colIdx = {};
    table.headers.forEach((h, i) => { colIdx[h] = i; });
    const ci = colIdx['国家地区'], di = colIdx['日期'];
    const out = [];
    let cur = '';
    for (const r of table.rows) {
      const d = toDay(r[di]);
      if (d) cur = d;
      const name = text(r[ci]);
      if (!name || name === '合计' || name === '总计' || name.indexOf('总计') === 0 || !cur) continue;
      const row = {'按天': cur};
      for (const c of COUNTRY_COLS) {
        row[c] = COUNTRY_NUMERIC.indexOf(c) >= 0 ? toFloat(r[colIdx[c]]) : r[colIdx[c]];
      }
      out.push(row);
    }
    return out;
  }

  // 入口：files = [{name, lastModified, bytes}]（拖进来的 xlsx，可以多个）
  // 返回与 make_html.generate() 的 payload 同构的一份数据
  async function buildDashboard(files) {
    if (!root.TTSheet) throw new Error('xlsx_read.js 没加载');
    const daily = await readDaily(files);
    const tables = clean(daily);
    const country = await cleanCountry(files);
    const countryDaily = await cleanCountryDaily(files);
    // 明细表只用 优化师 x 日期 两层，国家地区 不做展示列
    const detailCols = DASHBOARD_COLS.filter(c => c !== '国家地区' && tables.detail.some(r => c in r));
    return {robot: tables.robot, detail: tables.detail, detailCols: detailCols,
            country: country, countryDaily: countryDaily, meta: daily.meta};
  }

  root.TTClean = {
    buildDashboard: buildDashboard,
    toDay: toDay, toFloat: toFloat, nameDate: nameDate,
    DAILY_SHEET: DAILY_SHEET, COUNTRY_SHEET: COUNTRY_SHEET,
    COUNTRY_TIME_SHEET: COUNTRY_TIME_SHEET,
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
