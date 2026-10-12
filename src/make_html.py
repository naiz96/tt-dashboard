# -*- coding: utf-8 -*-
"""生成 HTML 看板（dashboard.html）。

自包含单文件：echarts（assets/echarts.min.js）和读表/清洗脚本（assets/xlsx_read.js、
assets/clean.js）全部内联，数据内嵌 <script> JSON 块，不引用任何外部文件——
双击即渲染，发给同事就这一个 html，同事也能把《投放日数据记录表》拖进来换数据。
配色遵循 dataviz 参考调色板（浅色模式）。

页面交互：
- 数据来源：打开时显示生成看板时的数据；页头「选择表格」或把 xlsx 拖到页面上，
  由内联 JS 现场解析（口径与 src/clean.py 一致），换成本地那份表格的数据
- 板块切换：总数据 / 国家数据 / 单人数据（单人页可选优化师，图表表格联动；
  国家页按国家/地区汇总，明细表点表头排序）
- 每个板块顶部可选时间段：按天 / 按周 / 按月（周以周一为起点），默认最新周期；
  KPI 卡与柱状图取选中周期，折线图每个点一个周期，图表标题标注当前周期。
  国家页多一个「全部」按钮（默认），显示 分国家汇总 那份全时段数据；
  选到某天/某周/某月时改用 分国家时间统计 聚合，比率列由合计值重算、
  时长与完播率按视频播放量加权
- 明细表为层级表：优化师 > 日期 两层，优化师可折叠；前 2 列（优化师/按天）
  横向滚动时冻结；指标列可勾选显示或隐藏。源表按国家组拆开的多行在这里按天合并显示
- 表格表头点击排序（数值/文本分别排序，ASC/DESC 切换；分组行按其聚合值排序）
- 明细表支持按日期筛选，父级行随之重算
- 父级行聚合口径见脚本内 AGG（与 clean.py 一致：求和项直接相加，比率一律由合计值重算）
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_PATH = os.path.join(BASE_DIR, 'dashboard.html')
ASSETS_DIR = os.path.join(BASE_DIR, 'assets')

# 看板明细表展示列（不含 国家地区：明细表只用 优化师 x 日期 两层，国家维度在国家页看）
_DETAIL_COLS = [
    ('优化师', 1), ('按天', 2),
    ('消耗', 1), ('广告收入', 1), ('首日ROI', 1), ('展示', 1), ('点击', 1), ('转化', 1),
    ('CPA', 1), ('CPM', 1), ('CTR', 1), ('CVR', 1),
]

_HTML = u'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>TT 投放看板</title>
<style>
:root {
  --page: #f9f9f7;
  --surface: #fcfcfb;
  --ink: #0b0b0b;
  --ink-2: #52514e;
  --muted: #898781;
  --grid: #e1e0d9;
  --axis: #c3c2b7;
  --border: rgba(11,11,11,0.10);
  --series-1: #2a78d6;
  --series-2: #eb6834;
  --series-3: #1baf7a;
  --series-4: #eda100;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  background: var(--page);
  color: var(--ink);
  padding: 24px 28px 48px;
}
header { display: flex; align-items: baseline; gap: 16px; margin-bottom: 16px; }
header h1 { font-size: 20px; font-weight: 600; }
header .meta { color: var(--ink-2); font-size: 13px; }
.tabs { display: flex; gap: 8px; margin-bottom: 18px; }
.tab-btn {
  font-family: inherit; font-size: 14px; font-weight: 500;
  padding: 7px 18px; border-radius: 6px; cursor: pointer;
  border: 1px solid var(--border); background: var(--surface); color: var(--ink-2);
}
.tab-btn:hover { color: var(--ink); }
.tab-btn.active { background: var(--series-1); border-color: var(--series-1); color: #fff; }
.view { display: none; }
.view.active { display: block; }
.person-bar, .period-bar { display: flex; align-items: center; gap: 8px; margin-bottom: 18px; }
.person-label, .period-label { font-size: 13px; color: var(--ink-2); }
.person-btn, .period-btn {
  font-family: inherit; font-size: 13px; font-weight: 500;
  padding: 5px 14px; border-radius: 999px; cursor: pointer;
  border: 1px solid var(--border); background: var(--surface); color: var(--ink-2);
}
.person-btn:hover, .period-btn:hover { color: var(--ink); }
.person-btn.active, .period-btn.active { background: var(--series-1); border-color: var(--series-1); color: #fff; }
.kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 18px; }
/* 国家页有 6 张卡：3 列正好两行，用 4 列会剩两个挂在第二行左边 */
.kpi-row.cols-3 { grid-template-columns: repeat(3, 1fr); }
.kpi-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 16px 18px 14px;
}
.kpi-card .label { font-size: 12px; color: var(--muted); margin-bottom: 6px; }
.kpi-card .value { font-size: 26px; font-weight: 600; letter-spacing: -0.01em; }
.kpi-card .value small { font-size: 14px; font-weight: 400; color: var(--ink-2); margin-left: 2px; }
.chart-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 18px; }
.chart-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 16px 18px 10px;
}
.chart-card h3 { font-size: 13px; font-weight: 600; color: var(--ink-2); margin-bottom: 8px; }
.chart { width: 100%; height: 280px; }
.table-card {
  background: var(--surface);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 16px 18px 12px;
  margin-bottom: 18px;
}
.table-card h3 { font-size: 13px; font-weight: 600; color: var(--ink-2); margin-bottom: 10px; }
.toolbar { display: flex; align-items: center; gap: 10px; margin-bottom: 10px; }
.filter-select {
  font-family: inherit; font-size: 13px; color: var(--ink);
  padding: 6px 8px;
  border: 1px solid var(--border); border-radius: 6px;
  background: var(--surface); outline: none;
}
.filter-select:focus { border-color: var(--series-1); }
.count { font-size: 12px; color: var(--muted); }
.table-wrap { overflow-x: auto; }
/* 国家汇总表 13 列，窄屏要横向滚，把国家名钉在左边免得滚过去就不知道在看哪国 */
#c-table th:first-child, #c-table td:first-child {
  position: sticky; left: 0; background: var(--surface); z-index: 1;
}
table { border-collapse: collapse; width: 100%; font-size: 13px; }
th, td { text-align: right; padding: 7px 10px; border-bottom: 1px solid var(--grid); white-space: nowrap; }
th { color: var(--muted); font-weight: 500; border-bottom: 1px solid var(--axis); cursor: pointer; user-select: none; }
th:hover { color: var(--ink-2); }
th.sorted { color: var(--ink); }
th.sorted.asc::after { content: ' ▲'; font-size: 10px; }
th.sorted.desc::after { content: ' ▼'; font-size: 10px; }
th:not(.num), td:not(.num) { text-align: left; }
td.num, th.num { font-variant-numeric: tabular-nums; }
/* 层级明细表：固定列宽 + 冻结前 3 列 + 指标列显隐 */
.tree-tbl { table-layout: fixed; border-collapse: separate; border-spacing: 0; }
/* 背景必须给到 th：冻结列靠它挡住从下面滚过去的表头，透明的话两层文字会叠在一起 */
.tree-tbl td, .tree-tbl th { background: var(--surface); }
.tree-tbl .fix { position: sticky; z-index: 2; overflow: hidden; text-overflow: ellipsis; }
.tree-tbl .fix-end { box-shadow: 3px 0 6px -3px rgba(11,11,11,0.18); }
.tree-tbl td[data-key] { cursor: pointer; user-select: none; }
.tree-tbl tr.g1 > td { background: #f1f1ee; }
.tree-tbl tr.g2 > td { background: #f7f7f4; }
.tree-tbl tr.g1 > td:first-child { font-weight: 600; }
.tree-tbl tr.g2 > td.lvl2 { font-weight: 500; }
.tree-tbl tr.leaf > td { color: var(--ink-2); }
.tree-tbl td.lvl2 { padding-left: 22px; }
.caret { display: inline-block; width: 13px; color: var(--muted); font-size: 11px; }
.colpick { position: relative; }
.col-btn {
  font-family: inherit; font-size: 13px; color: var(--ink-2);
  padding: 6px 12px; cursor: pointer;
  border: 1px solid var(--border); border-radius: 6px; background: var(--surface);
}
.col-btn:hover { color: var(--ink); }
.col-panel {
  position: absolute; top: 34px; left: 0; z-index: 10;
  width: 272px; max-height: 320px; overflow-y: auto;
  padding: 10px 12px;
  background: var(--surface); border: 1px solid var(--border); border-radius: 8px;
  box-shadow: 0 6px 20px rgba(11,11,11,0.12);
}
.col-panel[hidden] { display: none; }
.col-panel label {
  display: flex; align-items: center; gap: 6px;
  font-size: 13px; color: var(--ink-2); padding: 1px 0; cursor: pointer;
}
.col-panel-head {
  display: flex; gap: 8px; padding-bottom: 6px; margin-bottom: 4px;
  border-bottom: 1px solid var(--grid);
}
.col-panel-head button {
  font-family: inherit; font-size: 12px; color: var(--ink-2);
  padding: 2px 8px; cursor: pointer;
  border: 1px solid var(--border); border-radius: 4px; background: var(--surface);
}
.col-panel-head button:hover { color: var(--ink); }
@media (max-width: 1100px) {
  .kpi-row, .kpi-row.cols-3 { grid-template-columns: repeat(2, 1fr); }
  .chart-grid { grid-template-columns: 1fr; }
}
.src-bar { margin-left: auto; display: flex; align-items: center; gap: 10px; font-size: 13px; color: var(--muted); }
.src-bar button {
  font-family: inherit; font-size: 13px; color: var(--ink-2); cursor: pointer;
  padding: 4px 10px; border: 1px solid var(--border); border-radius: 4px; background: var(--surface);
}
.src-bar button:hover { color: var(--ink); }
.msg { display: none; padding: 9px 16px; font-size: 13px; background: #fdf1ec; color: #a3320f;
       border: 1px solid var(--border); border-radius: 4px; margin-bottom: 16px; }
.msg.show { display: block; }
#drop-mask {
  display: none; position: fixed; inset: 0; z-index: 9; text-align: center;
  align-items: center; justify-content: center; font-size: 18px; color: var(--ink-2);
  background: rgba(249,249,247,0.88); border: 2px dashed var(--axis);
}
#drop-mask.show { display: flex; }
</style>
<script>__LIB__</script>
</head>
<body>
<header>
  <h1>TT 投放看板</h1>
  <span class="meta" id="meta"></span>
  <span class="src-bar">
    <span id="src-label"></span>
    <button id="pick-btn">选择表格</button>
    <input type="file" id="file-input" accept=".xlsx" multiple hidden>
  </span>
</header>
<div class="msg" id="msg"></div>
<div id="drop-mask">松开鼠标，导入表格</div>
<div class="tabs">
  <button class="tab-btn active" data-view="view-total">总数据</button>
  <button class="tab-btn" data-view="view-country">国家数据</button>
  <button class="tab-btn" data-view="view-person">单人数据</button>
</div>

<div id="view-total" class="view active">
  <div class="period-bar" id="t-period">
    <span class="period-label">时间</span>
    <button class="period-btn active" data-gran="day">按天</button>
    <button class="period-btn" data-gran="week">按周</button>
    <button class="period-btn" data-gran="month">按月</button>
    <select class="filter-select period-filter"></select>
  </div>
  <div class="kpi-row">
    <div class="kpi-card"><div class="label">消耗</div><div class="value" id="kpi-spend"></div></div>
    <div class="kpi-card"><div class="label">首日ROI</div><div class="value" id="kpi-roi"></div></div>
    <div class="kpi-card"><div class="label">单次转化成本</div><div class="value" id="kpi-cpa"></div></div>
    <div class="kpi-card"><div class="label">转化量</div><div class="value" id="kpi-conv"></div></div>
  </div>
  <div class="table-card"><h3>看板明细表</h3><div id="detail-box"></div></div>
  <div class="chart-grid">
    <div class="chart-card"><h3 id="t-ttl-spend">按天消耗</h3><div class="chart" id="chart-spend"></div></div>
    <div class="chart-card"><h3 id="t-ttl-roi">按天首日ROI</h3><div class="chart" id="chart-roi"></div></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><h3 id="t-ttl-opt">优化师消耗对比</h3><div class="chart" id="chart-opt"></div></div>
    <div class="chart-card"><h3 id="t-ttl-opt-conv">优化师转化量对比</h3><div class="chart" id="chart-opt-conv"></div></div>
  </div>
</div>

<div id="view-country" class="view">
  <div class="period-bar" id="c-period">
    <span class="period-label">时间</span>
    <button class="period-btn active" data-gran="all">全部</button>
    <button class="period-btn" data-gran="day">按天</button>
    <button class="period-btn" data-gran="week">按周</button>
    <button class="period-btn" data-gran="month">按月</button>
    <select class="filter-select period-filter"></select>
  </div>
  <div class="kpi-row cols-3">
    <div class="kpi-card"><div class="label">消耗</div><div class="value" id="c-kpi-spend"></div></div>
    <div class="kpi-card"><div class="label">D0 广告收入</div><div class="value" id="c-kpi-rev"></div></div>
    <div class="kpi-card"><div class="label">首日ROI</div><div class="value" id="c-kpi-roi"></div></div>
    <div class="kpi-card"><div class="label">展示量</div><div class="value" id="c-kpi-impr"></div></div>
    <div class="kpi-card"><div class="label">点击量</div><div class="value" id="c-kpi-clk"></div></div>
    <div class="kpi-card"><div class="label">转化量</div><div class="value" id="c-kpi-conv"></div></div>
  </div>
  <div class="table-card">
    <h3 id="c-ttl-table">国家/地区汇总</h3>
    <div class="toolbar">
      <span class="count" id="c-count"></span>
    </div>
    <div class="table-wrap"><table id="c-table"></table></div>
  </div>
  <div class="chart-grid">
    <div class="chart-card"><h3 id="c-ttl-spend">国家/地区 消耗对比</h3><div class="chart" id="c-chart-spend-by-country"></div></div>
    <div class="chart-card"><h3 id="c-ttl-conv">国家/地区 转化量对比</h3><div class="chart" id="c-chart-conv-by-country"></div></div>
  </div>
</div>

<div id="view-person" class="view">
  <div class="person-bar">
    <span class="person-label">优化师</span>
    <span id="person-btns"></span>
  </div>
  <div class="period-bar" id="p-period">
    <span class="period-label">时间</span>
    <button class="period-btn active" data-gran="day">按天</button>
    <button class="period-btn" data-gran="week">按周</button>
    <button class="period-btn" data-gran="month">按月</button>
    <select class="filter-select period-filter"></select>
  </div>
  <div class="kpi-row">
    <div class="kpi-card"><div class="label">消耗</div><div class="value" id="p-kpi-spend"></div></div>
    <div class="kpi-card"><div class="label">首日ROI</div><div class="value" id="p-kpi-roi"></div></div>
    <div class="kpi-card"><div class="label">单次转化成本</div><div class="value" id="p-kpi-cpa"></div></div>
    <div class="kpi-card"><div class="label">转化量</div><div class="value" id="p-kpi-conv"></div></div>
  </div>
  <div class="table-card"><h3>看板明细表</h3><div id="p-detail-box"></div></div>
  <div class="chart-grid">
    <div class="chart-card"><h3 id="p-ttl-spend">按天消耗</h3><div class="chart" id="p-chart-spend"></div></div>
    <div class="chart-card"><h3 id="p-ttl-roi">按天首日ROI</h3><div class="chart" id="p-chart-roi"></div></div>
  </div>
</div>
<script>
const DATA = __DATA__;
const INK = '#0b0b0b', INK2 = '#52514e', MUTED = '#898781',
      GRID = '#e1e0d9', AXIS = '#c3c2b7', S1 = '#2a78d6';

const $ = id => document.getElementById(id);
const num = v => { const n = Number(v); return isNaN(n) ? 0 : n; };
const pct = v => (num(v) * 100).toFixed(2) + '%';
const money = v => num(v).toLocaleString('zh-CN', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const int0 = v => num(v).toLocaleString('zh-CN');

// 当前展示的数据：打开页面时装内置的那份，同事拖表格进来就换成导入的（见文件末尾 boot）
let robot = [], detail = [], detailCols = [], country = [], countryDaily = [], dates = [], persons = [];

function agg(rows) {
  const spend = rows.reduce((s, r) => s + num(r['消耗']), 0);
  const conv = rows.reduce((s, r) => s + num(r['转化量']), 0);
  const roiW = rows.reduce((s, r) => s + num(r['消耗']) * num(r['首日ROI']), 0);
  return { spend: spend, conv: conv, roi: spend > 0 ? roiW / spend : 0, cpa: conv > 0 ? spend / conv : 0 };
}

const baseAxis = {
  axisLine: {lineStyle: {color: AXIS}},
  axisTick: {show: false},
  axisLabel: {color: MUTED},
  splitLine: {lineStyle: {color: GRID}}
};
const baseTooltip = {
  backgroundColor: '#fcfcfb',
  borderColor: GRID,
  textStyle: {color: INK, fontSize: 12},
  valueFormatter: v => (v === null || v === undefined || isNaN(v)) ? '-' : Number(v).toLocaleString('zh-CN')
};
const lineStyle = {width: 2, color: S1};

function catAxis(data) { return Object.assign({type: 'category', data: data}, baseAxis); }
function valAxis(fmt) {
  return Object.assign({type: 'value'}, baseAxis, fmt ? {axisLabel: {color: MUTED, formatter: fmt}} : {});
}

// isPct: 数值按百分比展示（ROI 图）
function lineOpt(xData, yData, isPct, area) {
  const tipFmt = isPct ? v => (num(v) * 100).toFixed(2) + '%' : null;
  const axisFmt = isPct ? v => (v * 100).toFixed(0) + '%' : null;
  return {
    grid: {left: 60, right: 20, top: 16, bottom: 28},
    tooltip: Object.assign({trigger: 'axis', axisPointer: {type: 'cross'}}, baseTooltip,
      tipFmt ? {valueFormatter: tipFmt} : {}),
    xAxis: catAxis(xData),
    yAxis: valAxis(axisFmt),
    series: [{
      type: 'line', data: yData, lineStyle: lineStyle, itemStyle: {color: S1},
      symbol: 'circle', symbolSize: 8,
      areaStyle: area ? {color: 'rgba(42,120,214,0.06)'} : undefined,
      emphasis: {focus: 'series'}
    }]
  };
}

function barOpt(names, values, labelFmt, horizontal) {
  const label = {show: true, color: INK2, fontSize: 11,
                 formatter: p => (labelFmt || money)(p.value)};
  if (horizontal) {
    // 横向柱状图：名称完整显示在左侧（containLabel 自动留白不截断），按值降序
    const pairs = names.map((n, i) => ({n: n, v: values[i]}))
                       .sort((a, b) => b.v - a.v);
    return {
      grid: {left: 10, right: 64, top: 10, bottom: 20, containLabel: true},
      tooltip: Object.assign({trigger: 'axis', axisPointer: {type: 'shadow'}}, baseTooltip),
      xAxis: valAxis(),
      yAxis: Object.assign({type: 'category', inverse: true,
        data: pairs.map(p => p.n)}, baseAxis,
        {axisLabel: {color: MUTED, fontSize: 11}, splitLine: {show: false}}),
      series: [{
        type: 'bar', data: pairs.map(p => p.v), barMaxWidth: 22,
        itemStyle: {color: S1, borderRadius: [0, 4, 4, 0]},
        label: Object.assign({position: 'right'}, label)
      }]
    };
  }
  return {
    grid: {left: 60, right: 20, top: 28, bottom: 40},
    tooltip: Object.assign({trigger: 'axis', axisPointer: {type: 'shadow'}}, baseTooltip),
    xAxis: Object.assign(catAxis(names), {
      axisLabel: {color: MUTED, interval: 0, rotate: names.length > 4 ? 28 : 0, fontSize: 11}
    }),
    yAxis: valAxis(),
    series: [{
      type: 'bar', data: values, barMaxWidth: 48,
      itemStyle: {color: S1, borderRadius: [4, 4, 0, 0]},
      label: Object.assign({position: 'top'}, label)
    }]
  };
}

function initChart(id) {
  const chart = echarts.init($(id));
  window.addEventListener('resize', () => chart.resize());
  return chart;
}

// ===== 层级明细表 =====
// 优化师 > 日期 两层：同一天按国家组拆开的多行在这里就合并掉了，国家维度看国家页
// 冻结列（横向滚动时固定）及其宽度，累加得出各列的 sticky left 偏移
const ID_COLS = ['优化师', '按天'];
const FIX_W = {'优化师': 84, '按天': 110};

// 父级行聚合口径（与 clean.py / notify.py 一致）
// ['sum'] 求和；['w', x] 按 x 加权；['r', 分子, 分母, k] = Σ分子/Σ分母*k
// 比率一律由合计值重算，不能对叶子行的比率再平均
const AGG = {
  '消耗': ['sum'],
  '广告收入': ['sum'],
  '展示': ['sum'],
  '点击': ['sum'],
  '转化': ['sum'],
  '首日ROI': ['w', '消耗'],
  'CPA': ['r', '消耗', '转化', 1],
  'CPM': ['r', '消耗', '展示', 1000],
  'CTR': ['r', '点击', '展示', 1],
  'CVR': ['r', '转化', '点击', 1],
};

// 各数值列的展示小数位（父级行与叶子行同一列用同一位数，保证列对齐）
const DEC = {
  '消耗': 2, '广告收入': 2, '首日ROI': 3, '展示': 0, '点击': 0, '转化': 0,
  'CPA': 2, 'CPM': 2, 'CTR': 4, 'CVR': 4,
};

function sumCol(rows, col) { return rows.reduce((s, r) => s + num(r[col]), 0); }

// 分组行聚合值；返回 null = 该列汇总不出来
function aggVal(col, rows) {
  const rule = AGG[col];
  if (!rule || !rows.length) return null;
  if (rule[0] === 'sum') return sumCol(rows, col);
  const den = sumCol(rows, rule[2] || rule[1]);
  if (!(den > 0)) return null;
  if (rule[0] === 'w') return rows.reduce((s, r) => s + num(r[rule[1]]) * num(r[col]), 0) / den;
  return sumCol(rows, rule[1]) / den * rule[3];
}

// 指标列宽度：按表头字宽估算（13px 字号，中文按 1.75 个西文字符）
function colWidth(name) {
  let w = 0;
  for (const ch of name) w += ch.charCodeAt(0) > 255 ? 1.75 : 1;
  return Math.max(76, Math.round(w * 7.6) + 22);
}

function dataTable(boxId, cols, rows, textCols) {
  const box = $(boxId);
  const idCols = ID_COLS.filter(c => cols.indexOf(c) >= 0);
  const metricCols = cols.filter(c => idCols.indexOf(c) < 0);
  // expanded 记录"用户手动展开过"的组，集合为空 = 全部收拢（首次进入只显示优化师层级）
  const hidden = new Set(), expanded = new Set();
  const fixLeft = {};
  let fx = 0;
  idCols.forEach(c => { fixLeft[c] = fx; fx += FIX_W[c] || 0; });
  const fixEnd = idCols[idCols.length - 1];
  let sortKey = null, sortDir = 1;

  box.innerHTML = '<div class="toolbar">' +
    '<select class="filter-select date-filter"><option value="">全部日期</option></select>' +
    '<div class="colpick"><button type="button" class="col-btn">显示列</button>' +
    '<div class="col-panel" hidden><div class="col-panel-head">' +
    '<button type="button" class="col-all">全选</button>' +
    '<button type="button" class="col-none">全不选</button></div>' +
    metricCols.map(c => '<label><input type="checkbox" checked data-col="' + c + '">' + c + '</label>').join('') +
    '</div></div>' +
    '<span class="count"></span></div>' +
    '<div class="table-wrap"><table class="tree-tbl"></table></div>';
  const table = box.querySelector('table');
  const dateSel = box.querySelector('.date-filter');
  const count = box.querySelector('.count');
  const panel = box.querySelector('.col-panel');

  const isNumCol = c => !textCols.has(c);
  const visible = () => {
    let data = rows;
    if (dateSel.value) data = data.filter(r => r['按天'] === dateSel.value);
    return data;
  };

  // 过滤后的行 -> 优化师组 -> 日期组（日期组的 leaves 是源表按国家组拆开的那些行，只用于聚合）
  function buildTree(leaves) {
    const byOpt = new Map();
    leaves.forEach(r => {
      const opt = r['优化师'];
      if (!byOpt.has(opt)) byOpt.set(opt, new Map());
      const days = byOpt.get(opt);
      const dk = r['按天'];
      if (!days.has(dk)) days.set(dk, {key: 'day:' + opt + '|' + dk, label: dk, leaves: []});
      days.get(dk).leaves.push(r);
    });
    return [...byOpt].map(([opt, days]) => {
      const list = [...days.values()];
      return {key: 'opt:' + opt, label: opt, days: list,
              leaves: list.reduce((all, d) => all.concat(d.leaves), [])};
    });
  }

  // 同级排序用值：分组行取聚合值，叶子行取自身值；无法汇总的排最后
  function sortValue(key, grp, own) {
    if (!isNumCol(key)) return own ? String(own[key] === undefined ? '' : own[key]) : '';
    const v = own ? num(own[key]) : aggVal(key, grp);
    if (v === null) return sortDir > 0 ? Infinity : -Infinity;
    return v;
  }

  function sortTree(tree) {
    if (!sortKey) return;
    const cmp = (a, b) => (typeof a === 'number' ? a - b : a.localeCompare(b, 'zh')) * sortDir;
    tree.sort((g1, g2) => cmp(sortValue(sortKey, g1.leaves), sortValue(sortKey, g2.leaves)));
    tree.forEach(g => {
      g.days.sort((a1, a2) => cmp(sortValue(sortKey, a1.leaves), sortValue(sortKey, a2.leaves)));
      g.days.forEach(a => a.leaves.sort((r1, r2) =>
        cmp(sortValue(sortKey, null, r1), sortValue(sortKey, null, r2))));
    });
  }

  function cellCls(col) {
    const parts = [];
    if (fixLeft[col] !== undefined) parts.push('fix');
    if (col === fixEnd) parts.push('fix-end');
    if (isNumCol(col)) parts.push('num');
    return parts.join(' ');
  }

  function td(col, attr, html, extra) {
    const cls = cellCls(col) + (extra ? ' ' + extra : '');
    const st = fixLeft[col] === undefined ? '' : ' style="left:' + fixLeft[col] + 'px"';
    return '<td' + (cls ? ' class="' + cls + '"' : '') + st + (attr || '') + '>' + html + '</td>';
  }

  // 指标单元格：分组行用聚合值，叶子行用原值；数值统一按列定小数位，行行对齐
  function metricTd(col, grp, own) {
    const v = own ? (own[col] === undefined || own[col] === null ? null : own[col]) : aggVal(col, grp);
    if (v === null || v === '') return td(col, '', '-');
    if (!isNumCol(col)) return td(col, '', String(v));
    return td(col, '', num(v).toFixed(DEC[col] === undefined ? 2 : DEC[col]));
  }

  // mcols 只含指标列：冻结列已由 rowspan 占位，各层级行不得再输出，否则整表右移
  function optRow(g, span, mcols) {
    const open = expanded.has(g.key);
    let s = td(idCols[0], ' rowspan="' + span + '" data-key="' + g.key + '"',
               '<span class="caret">' + (open ? '▾' : '▸') + '</span>' + g.label);
    for (let i = 1; i < idCols.length; i++) s += td(idCols[i], '', '');
    return '<tr class="g1">' + s + mcols.map(c => metricTd(c, g.leaves)).join('') + '</tr>';
  }

  // 日期组行：值是该天各国家组行的合计
  function dayRow(a, mcols) {
    return '<tr class="g2">' + td(idCols[1], '', a.label, 'lvl2') +
           mcols.map(c => metricTd(c, a.leaves)).join('') + '</tr>';
  }

  function bodyRows(tree, mcols) {
    const out = [];
    tree.forEach(g => {
      const days = expanded.has(g.key) ? g.days : [];
      out.push(optRow(g, 1 + days.length, mcols));
      days.forEach(a => out.push(dayRow(a, mcols)));
    });
    return out.join('');
  }

  function render() {
    // 日期选项随数据更新（切换优化师时）
    const opts = [...new Set(rows.map(r => r['按天']))].sort();
    const cur = dateSel.value;
    dateSel.innerHTML = '<option value="">全部日期</option>' +
      opts.map(d => '<option' + (d === cur ? ' selected' : '') + '>' + d + '</option>').join('');
    if (cur && opts.indexOf(cur) < 0) dateSel.value = '';

    const leaves = visible();
    const disp = idCols.concat(metricCols.filter(c => !hidden.has(c)));
    const mcols = disp.filter(c => idCols.indexOf(c) < 0);
    const tree = buildTree(leaves);
    sortTree(tree);

    let total = 0;
    const colgroup = disp.map(c => {
      const w = fixLeft[c] === undefined ? colWidth(c) : FIX_W[c];
      total += w;
      return '<col style="width:' + w + 'px">';
    }).join('');
    table.style.width = 'max(100%, ' + total + 'px)';
    const head = '<thead><tr>' + disp.map(c => {
      const st = fixLeft[c] === undefined ? '' : ' style="left:' + fixLeft[c] + 'px"';
      const cls = cellCls(c) + (sortKey === c ? (sortDir > 0 ? ' sorted asc' : ' sorted desc') : '');
      return '<th class="' + cls + '" data-col="' + c + '"' + st + '>' + c + '</th>';
    }).join('') + '</tr></thead>';

    table.innerHTML = '<colgroup>' + colgroup + '</colgroup>' + head +
                      '<tbody>' + bodyRows(tree, mcols) + '</tbody>';
    count.textContent = '共 ' + tree.reduce((n, g) => n + g.days.length, 0) + ' 条（优化师 x 日期）';
  }

  // 折叠/展开（点分组行的名称格）与表头排序
  table.addEventListener('click', e => {
    const grp = e.target.closest('td[data-key]');
    if (grp) {
      const k = grp.getAttribute('data-key');
      if (expanded.has(k)) expanded.delete(k); else expanded.add(k);
      render();
      return;
    }
    const th = e.target.closest('th[data-col]');
    if (!th) return;
    const c = th.getAttribute('data-col');
    if (sortKey === c) sortDir = -sortDir;
    else { sortKey = c; sortDir = isNumCol(c) ? -1 : 1; }
    render();
  });

  function syncChecks() {
    panel.querySelectorAll('input[data-col]').forEach(b => {
      b.checked = !hidden.has(b.getAttribute('data-col'));
    });
  }

  dateSel.addEventListener('change', render);
  box.querySelector('.col-btn').addEventListener('click', () => { panel.hidden = !panel.hidden; });
  panel.addEventListener('change', e => {
    const c = e.target.getAttribute('data-col');
    if (!c) return;
    if (e.target.checked) hidden.delete(c); else hidden.add(c);
    render();
  });
  panel.querySelector('.col-all').addEventListener('click', () => { hidden.clear(); syncChecks(); render(); });
  panel.querySelector('.col-none').addEventListener('click', () => {
    metricCols.forEach(c => hidden.add(c));
    syncChecks();
    render();
  });
  render();
  return {
    setRows: function (newRows) {
      rows = newRows;
      sortKey = null;
      sortDir = 1;
      expanded.clear();
      render();
    }
  };
}

const detailTextCols = new Set(['优化师', '按天']);

function setKpi(prefix, k) {
  $(prefix + '-spend').innerHTML = money(k.spend);
  $(prefix + '-roi').innerHTML = pct(k.roi);
  $(prefix + '-cpa').innerHTML = money(k.cpa);
  $(prefix + '-conv').textContent = int0(k.conv);
}

// ===== 时间段（按天 / 按周 / 按月）=====
// 按周/按月 = 把数据按周/月汇总：KPI 卡与柱状图取选中周期，折线图一个点一个周期。
// 周以周一为起点；周期名用桶内实际有数据的天数区间，不满整周/整月时如实显示。
const UNIT = {day: '天', week: '周', month: '月'};

function fmtDate(d) {
  const p = n => (n < 10 ? '0' : '') + n;
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

function bucketOf(day, gran) {
  if (gran === 'day') return day;
  if (gran === 'month') return day.slice(0, 7);
  const d = new Date(day + 'T00:00:00');
  d.setDate(d.getDate() - (d.getDay() + 6) % 7);
  return fmtDate(d);
}

// 数据里出现过的周期，按时间升序；每个周期带桶内实际有的天（用于命名）
function bucketsOf(rows, gran) {
  const m = new Map();
  rows.forEach(r => {
    const k = bucketOf(r['按天'], gran);
    if (!m.has(k)) m.set(k, []);
    if (m.get(k).indexOf(r['按天']) < 0) m.get(k).push(r['按天']);
  });
  return [...m.keys()].sort().map(k => ({key: k, days: m.get(k).sort()}));
}

function bucketLabel(b, gran) {
  if (gran === 'month') return b.key;
  return b.days[0] === b.days[b.days.length - 1] ? b.days[0]
       : b.days[0] + ' ~ ' + b.days[b.days.length - 1];
}

// 每个板块一个时间段控件：粒度按钮 + 周期下拉，默认选中最新周期；改动后回调重画该板块。
// 调用返回的 pick(rows)：rows 为该板块的全量行，返回
// {gran, unit, label, rows: 选中周期内的行, points: 各周期 [{label, rows}]}
function makePeriodBar(barId, redraw) {
  const bar = $(barId), sel = bar.querySelector('.period-filter');
  const btns = [...bar.querySelectorAll('.period-btn')];
  let gran = 'day', key = '';
  btns.forEach(b => b.addEventListener('click', () => {
    gran = b.getAttribute('data-gran');
    key = '';   // 换粒度后回到该粒度的最新周期
    btns.forEach(x => x.classList.toggle('active', x === b));
    redraw();
  }));
  sel.addEventListener('change', () => { key = sel.value; redraw(); });

  return function (rows) {
    const buckets = bucketsOf(rows, gran);
    if (!buckets.some(b => b.key === key)) key = buckets.length ? buckets[buckets.length - 1].key : '';
    sel.innerHTML = buckets.map(b =>
      '<option value="' + b.key + '">' + bucketLabel(b, gran) + '</option>').join('');
    sel.value = key;

    const inBucket = b => rows.filter(r => bucketOf(r['按天'], gran) === b.key);
    const cur = buckets.filter(b => b.key === key)[0];
    return {
      gran: gran, unit: UNIT[gran],
      label: cur ? bucketLabel(cur, gran) : '',
      rows: cur ? inBucket(cur) : [],
      points: buckets.map(b => ({label: bucketLabel(b, gran), rows: inBucket(b)}))
    };
  };
}

// 折线图取点：每个周期一个值
function spendPoints(points) {
  return points.map(p => Math.round(p.rows.reduce((s, r) => s + num(r['消耗']), 0) * 100) / 100);
}
function roiPoints(points) {
  return points.map(p => Math.round(agg(p.rows).roi * 10000) / 10000);
}

// ===== 总数据 =====
let totalRendered = false, tBar = null, tCharts = null, tTable = null;
function renderTotal() {
  if (!tBar) tBar = makePeriodBar('t-period', renderTotal);
  const p = tBar(robot);
  setKpi('kpi', agg(p.rows));

  if (!tCharts) {
    tCharts = {
      spend: initChart('chart-spend'), roi: initChart('chart-roi'),
      opt: initChart('chart-opt'), optConv: initChart('chart-opt-conv')
    };
  }
  const xLabels = p.points.map(x => x.label);
  $('t-ttl-spend').textContent = '按' + p.unit + '消耗';
  $('t-ttl-roi').textContent = '按' + p.unit + '首日ROI';
  tCharts.spend.setOption(lineOpt(xLabels, spendPoints(p.points)), true);
  tCharts.roi.setOption(lineOpt(xLabels, roiPoints(p.points), true, true), true);

  const byOpt = {}, byOptConv = {};
  p.rows.forEach(r => {
    byOpt[r['优化师']] = (byOpt[r['优化师']] || 0) + num(r['消耗']);
    byOptConv[r['优化师']] = (byOptConv[r['优化师']] || 0) + num(r['转化量']);
  });
  $('t-ttl-opt').textContent = '优化师消耗对比（' + p.label + '）';
  $('t-ttl-opt-conv').textContent = '优化师转化量对比（' + p.label + '）';
  tCharts.opt.setOption(barOpt(Object.keys(byOpt), Object.values(byOpt)), true);
  tCharts.optConv.setOption(barOpt(Object.keys(byOptConv), Object.values(byOptConv), int0), true);

  // 明细表是 优化师 x 日期 的全量明细（自带日期筛选），不随时间段切换重画
  if (!tTable) tTable = dataTable('detail-box', detailCols, detail, detailTextCols);
}

// ===== 单人数据 =====
let pCharts = null, pTable = null, pBar = null;
let currentPerson = '';

function renderPerson(name) {
  const switched = name !== currentPerson;
  currentPerson = name;
  document.querySelectorAll('#person-btns .person-btn').forEach(b =>
    b.classList.toggle('active', b.textContent === name));
  const prows = robot.filter(r => r['优化师'] === name);
  const drows = detail.filter(r => r['优化师'] === name);

  if (!pBar) pBar = makePeriodBar('p-period', () => renderPerson(currentPerson));
  const p = pBar(prows);
  setKpi('p-kpi', agg(p.rows));

  if (!pCharts) {
    pCharts = {
      spend: initChart('p-chart-spend'), roi: initChart('p-chart-roi')
    };
  }
  const xLabels = p.points.map(x => x.label);
  $('p-ttl-spend').textContent = '按' + p.unit + '消耗';
  $('p-ttl-roi').textContent = '按' + p.unit + '首日ROI';
  pCharts.spend.setOption(lineOpt(xLabels, spendPoints(p.points)), true);
  pCharts.roi.setOption(lineOpt(xLabels, roiPoints(p.points), true, true), true);

  // 明细表只在切换优化师时换数据，切时间段不重画，免得清掉展开/排序状态
  if (!pTable) pTable = dataTable('p-detail-box', detailCols, drows, detailTextCols);
  else if (switched) pTable.setRows(drows);
}

// ===== 国家数据 =====
// "全部"时段用 分国家汇总（一国一行的全时段汇总，原样展示，不做二次聚合）；
// 选到某天/某周/某月时改用 分国家时间统计（一天一国一行）聚合。
// 表列：[列名, 取值（排序用，数值列取成 number）, 展示]
const C_COLS = [
  ['国家地区', r => r['国家地区'], v => v],
  ['消耗', r => num(r['消耗']), money],
  ['D0 广告收入', r => num(r['D0 广告收入']), money],
  ['D0 ROI', r => num(r['D0 ROI']), pct],
  ['展示', r => num(r['展示']), int0],
  ['点击', r => num(r['点击']), int0],
  ['转化', r => num(r['转化']), int0],
  ['视频播放量', r => num(r['视频播放量']), int0],
  ['平均视频播放时长', r => num(r['平均视频播放时长']), v => v.toFixed(2)],
  // 源表这列是比率（0.1247），单元格带百分比格式、Excel 里显示 12.47%，所以过一遍 pct
  ['视频完播率', r => num(r['视频完播率']), pct],
  ['播放 75% 次数', r => num(r['播放 75% 次数']), int0],
  ['播放 50% 次数', r => num(r['播放 50% 次数']), int0],
  ['播放 25% 次数', r => num(r['播放 25% 次数']), int0]
];

// 求和列；比率列一律由合计值重算，时长/完播率按视频播放量加权
// （选"按天"时一天只有一行，算出来就是源表原值）
const C_SUM = ['消耗', 'D0 广告收入', '展示', '点击', '转化',
               '视频播放量', '播放 75% 次数', '播放 50% 次数', '播放 25% 次数'];

function countryAgg(rows) {
  const byCty = new Map();
  rows.forEach(r => {
    const k = r['国家地区'];
    if (!byCty.has(k)) byCty.set(k, []);
    byCty.get(k).push(r);
  });
  return [...byCty].map(([name, list]) => {
    const sum = c => list.reduce((s, r) => s + num(r[c]), 0);
    const play = sum('视频播放量');
    const wavg = c => {
      if (play > 0) return list.reduce((s, r) => s + num(r[c]) * num(r['视频播放量']), 0) / play;
      return list.length ? list.reduce((s, r) => s + num(r[c]), 0) / list.length : 0;
    };
    const row = {'国家地区': name};
    C_SUM.forEach(c => { row[c] = sum(c); });
    row['D0 ROI'] = row['消耗'] > 0 ? row['D0 广告收入'] / row['消耗'] : 0;
    row['平均视频播放时长'] = wavg('平均视频播放时长');
    row['视频完播率'] = wavg('视频完播率');
    return row;
  });
}

let cSort = 1, cDir = -1;   // 默认按消耗降序
let cRows = [];             // 当前时间段要展示的国家行

// KPI 卡的汇总值：各行直接相加，首日ROI 用 D0 收入/消耗（与源表列口径一致）
function countrySum() {
  const sum = c => cRows.reduce((s, r) => s + num(r[c]), 0);
  const spend = sum('消耗'), rev = sum('D0 广告收入');
  return {spend: spend, rev: rev, roi: spend > 0 ? rev / spend : 0,
          impr: sum('展示'), clk: sum('点击'), conv: sum('转化')};
}

function countryTable() {
  const col = C_COLS[cSort];
  const list = cRows.slice().sort((x, y) => {
    const a = col[1](x), b = col[1](y);
    return (typeof a === 'number' ? a - b : String(a).localeCompare(String(b), 'zh')) * cDir;
  });

  $('c-table').innerHTML =
    '<thead><tr>' + C_COLS.map((c, i) => {
      const cls = (i === 0 ? '' : 'num') +
                  (i === cSort ? (cDir > 0 ? ' sorted asc' : ' sorted desc') : '');
      return '<th class="' + cls.trim() + '" data-col="' + i + '">' + c[0] + '</th>';
    }).join('') + '</tr></thead><tbody>' +
    list.map(r => '<tr>' + C_COLS.map((c, i) =>
      '<td' + (i === 0 ? '' : ' class="num"') + '>' + c[2](c[1](r)) + '</td>'
    ).join('') + '</tr>').join('') + '</tbody>';
  $('c-count').textContent = list.length + ' 个国家/地区';
}

$('c-table').addEventListener('click', e => {
  const th = e.target.closest('th[data-col]');
  if (!th) return;
  const i = Number(th.getAttribute('data-col'));
  if (i === cSort) cDir = -cDir; else { cSort = i; cDir = i === 0 ? 1 : -1; }
  countryTable();
});

// 国家页时间段控件：全部（分国家汇总 那份它自己的全时段汇总）/ 按天 / 按周 / 按月。
// rows 为 分国家时间统计 的"一天一国一行"；返回 {label, rows}，rows 为 null = 全部时段
function makeCountryBar(barId, redraw) {
  const bar = $(barId), sel = bar.querySelector('.period-filter');
  const btns = [...bar.querySelectorAll('.period-btn')];
  let gran = 'all', key = '';
  btns.forEach(b => b.addEventListener('click', () => {
    gran = b.getAttribute('data-gran');
    key = '';   // 换粒度后回到该粒度的最新周期
    btns.forEach(x => x.classList.toggle('active', x === b));
    redraw();
  }));
  sel.addEventListener('change', () => { key = sel.value; redraw(); });

  return function (rows) {
    if (gran === 'all' || !rows.length) {   // 没有日数据（老表）时一律按全时段显示
      sel.innerHTML = '<option value="">全部时段</option>';
      return {label: '', rows: null};
    }
    const buckets = bucketsOf(rows, gran);
    if (!buckets.some(b => b.key === key)) key = buckets.length ? buckets[buckets.length - 1].key : '';
    sel.innerHTML = buckets.map(b =>
      '<option value="' + b.key + '">' + bucketLabel(b, gran) + '</option>').join('');
    sel.value = key;
    const cur = buckets.filter(b => b.key === key)[0];
    return {label: cur ? bucketLabel(cur, gran) : '',
            rows: cur ? rows.filter(r => bucketOf(r['按天'], gran) === cur.key) : []};
  };
}

let countryRendered = false, cCharts = null, cBar = null;
function renderCountry() {
  if (!cBar) cBar = makeCountryBar('c-period', renderCountry);
  // 表格里没有 分国家时间统计 时（老表）不显示时间段控件
  $('c-period').style.display = countryDaily.length ? '' : 'none';
  const period = cBar(countryDaily);
  cRows = period.rows === null ? country : countryAgg(period.rows);
  const suffix = period.label ? '（' + period.label + '）' : '';

  const k = countrySum();
  $('c-kpi-spend').innerHTML = money(k.spend);
  $('c-kpi-rev').innerHTML = money(k.rev);
  $('c-kpi-roi').innerHTML = pct(k.roi);
  $('c-kpi-impr').innerHTML = int0(k.impr);
  $('c-kpi-clk').innerHTML = int0(k.clk);
  $('c-kpi-conv').innerHTML = int0(k.conv);
  $('c-ttl-table').textContent = '国家/地区汇总' + suffix;
  $('c-ttl-spend').textContent = '国家/地区 消耗对比' + suffix;
  $('c-ttl-conv').textContent = '国家/地区 转化量对比' + suffix;

  if (!cCharts) {
    cCharts = {
      bySpend: initChart('c-chart-spend-by-country'),
      byConv: initChart('c-chart-conv-by-country')
    };
  }
  // 20 多国挤在默认 280px 里看不清，图高按国家数撑开（换时间段国家数会变，每次都要调）
  const h = Math.max(280, cRows.length * 26 + 30) + 'px';
  $('c-chart-spend-by-country').style.height = h;
  $('c-chart-conv-by-country').style.height = h;

  cCharts.bySpend.setOption(barOpt(
    cRows.map(r => r['国家地区']), cRows.map(r => num(r['消耗'])), money, true), true);
  cCharts.byConv.setOption(barOpt(
    cRows.map(r => r['国家地区']), cRows.map(r => num(r['转化'])), int0, true), true);

  countryTable();
}

// ===== 板块切换 =====
let personRendered = false, currentView = 'view-total';
function activateTab(viewId) {
  currentView = viewId;
  document.querySelectorAll('.tab-btn').forEach(b =>
    b.classList.toggle('active', b.getAttribute('data-view') === viewId));
  document.querySelectorAll('.view').forEach(v =>
    v.classList.toggle('active', v.id === viewId));
  if (viewId === 'view-total') {
    if (!totalRendered) { totalRendered = true; renderTotal(); }
  } else if (viewId === 'view-country') {
    if (!countryRendered) { countryRendered = true; renderCountry(); }
  } else if (!personRendered) {
    personRendered = true;
    renderPerson(currentPerson);
  }
}
document.querySelectorAll('.tab-btn').forEach(b =>
  b.addEventListener('click', () => activateTab(b.getAttribute('data-view'))));

// ===== 装数据：内置的那份 / 同事拖进来的那份，都走这里 =====
function boot(data, label) {
  robot = data.robot || [];
  detail = data.detail || [];
  detailCols = data.detailCols || [];
  country = data.country || [];
  countryDaily = data.countryDaily || [];
  dates = [...new Set(robot.map(r => r['按天']))].sort();
  persons = [...new Set(robot.map(r => r['优化师']))].sort();
  currentPerson = persons[0] || '';

  // 图和时间段控件复用（重复 init 会在同一个 div 上叠实例、重复绑事件会触发两次），
  // 表格和"已渲染"标志重置，让三个板块都按新数据重画
  totalRendered = personRendered = countryRendered = false;
  tTable = pTable = null;

  $('person-btns').innerHTML = persons.map(p =>
    '<button class="person-btn' + (p === currentPerson ? ' active' : '') + '">' + p + '</button>').join('');
  document.querySelectorAll('#person-btns .person-btn').forEach(b =>
    b.addEventListener('click', () => renderPerson(b.textContent)));
  // 没有国家表时不显示国家板块（没放国家汇总的表也能正常看看板）
  document.querySelector('.tab-btn[data-view="view-country"]').style.display = country.length ? '' : 'none';
  if (!country.length && currentView === 'view-country') currentView = 'view-total';

  $('src-label').textContent = label;
  $('meta').textContent = robot.length
    ? '数据日期: ' + dates.join(' ~ ') + '  |  更新于 ' + new Date().toLocaleString('zh-CN')
    : '还没有数据';
  activateTab(currentView);
}

// ===== 导入《投放日数据记录表》=====
function tip(text) {
  const box = $('msg');
  box.textContent = text || '';
  box.classList.toggle('show', !!text);
}

async function importFiles(fileList) {
  const files = [];
  for (const f of fileList) {
    files.push({name: f.name, lastModified: f.lastModified,
                bytes: new Uint8Array(await f.arrayBuffer())});
  }
  if (!files.length) return;
  try {
    const data = await TTClean.buildDashboard(files);
    if (!data.robot.length) {
      tip('这份表格里没读到投放数据，请确认拖进来的是《投放日数据记录表》。');
      return;
    }
    const extra = [];
    if (data.meta.unfilled) extra.push(data.meta.unfilled + ' 行只填了优化师、没填日期，已跳过');
    if (data.meta.dupDays.length) extra.push(data.meta.dupDays.length + ' 天在多份文件里重复，取了数据日期更新的那份');
    tip(extra.join('；'));
    boot(data, '数据来源: ' + files.map(f => f.name).join('、')
              + '（' + data.meta.days + ' 天 ' + data.meta.rows + ' 行）');
  } catch (e) {
    tip('导入失败: ' + ((e && e.message) || e));
  }
}

$('pick-btn').addEventListener('click', () => $('file-input').click());
$('file-input').addEventListener('change', e => {
  importFiles(e.target.files);
  e.target.value = '';   // 清掉，下次选同一个文件也触发
});
let dragDepth = 0;
window.addEventListener('dragenter', e => {
  e.preventDefault();
  dragDepth++;
  $('drop-mask').classList.add('show');
});
window.addEventListener('dragover', e => e.preventDefault());
window.addEventListener('dragleave', () => {
  if (--dragDepth <= 0) { dragDepth = 0; $('drop-mask').classList.remove('show'); }
});
window.addEventListener('drop', e => {
  e.preventDefault();
  dragDepth = 0;
  $('drop-mask').classList.remove('show');
  importFiles(e.dataTransfer.files);
});

// 打开页面先显示内置的那份数据（生成看板时的快照），随时可以拖表格进来换
boot(DATA, '数据来源: 内置数据');
// 再试着拉实时数据（飞书直连）：成功就换成实时（刷新页面即最新）；失败（离线、口令不对、
// 没配飞书）静默停在内置快照。file:// 双击打开时走本机共享的地址，共享在跑也能刷到最新。
const API = location.protocol === 'file:' ? 'http://127.0.0.1:8000/api/data'
                                          : 'api/data' + location.search;
fetch(API, {cache: 'no-store'})
  .then(r => (r.ok ? r.json() : null))
  .then(d => { if (d && d.robot && d.robot.length) boot(d, '数据来源: 飞书实时'); })
  .catch(() => {});
</script>
</body>
</html>
'''


def _js(name):
    """把 assets/ 下的 JS 读进来塞进 <script>。

    JS 里万一出现 </script> 会提前闭合脚本块，替换成 <\\/script>（在字符串/正则里等价）。
    """
    with open(os.path.join(ASSETS_DIR, name), encoding='utf-8') as f:
        return f.read().replace('</script>', '<\\/script>')


def payload(dashboard_rows, robot_rows, country_rows=None, country_daily_rows=None):
    """看板数据载荷：内嵌进 html 的 DATA 与 /api/data 返回的是同一结构。

    country_daily_rows 是 分国家时间统计 的"一天一国一行"，供国家页按天/周/月筛选。
    """
    detail_cols = [c for c, _w in _DETAIL_COLS if any(c in r for r in dashboard_rows)]
    return {
        'robot': robot_rows,
        'detail': dashboard_rows,
        'detailCols': detail_cols,
        'country': country_rows or [],
        'countryDaily': country_daily_rows or [],
    }


def generate(dashboard_rows, robot_rows, country_rows=None, country_daily_rows=None):
    """生成 dashboard.html。行均为 dict（字段名 -> 值），与飞书读回结构一致。

    echarts 读表/清洗两个 JS 全部内联，生成物不引用任何外部文件：发给同事就是这一个
    html，双击即看；同事拖《投放日数据记录表》进来，由内联的 JS 自己解析。
    打开页面时会再试一次 /api/data（飞书直连），拿不到就用内嵌的这份快照。
    """
    data = payload(dashboard_rows, robot_rows, country_rows, country_daily_rows)
    lib = '\n;\n'.join(_js(n) for n in ('echarts.min.js', 'xlsx_read.js', 'clean.js'))
    html = (_HTML.replace('__DATA__',
                          json.dumps(data, ensure_ascii=False).replace('</', '<\\/'))
                 .replace('__LIB__', lib))
    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)
    return HTML_PATH


if __name__ == '__main__':
    import clean
    dash, robot, _cols = clean.clean()
    path = generate(dash, robot, clean.clean_country(), clean.clean_country_daily())
    print('[OK] 看板已生成: %s' % path)
