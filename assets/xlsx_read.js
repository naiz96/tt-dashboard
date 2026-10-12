// src/xlsx_read.py 的 JS 版：零依赖读 xlsx。
//
// 自己扫 zip 中央目录 + DecompressionStream('deflate-raw') 解压 + 正则扫 XML
// （不用 DOMParser：这样同一份代码在浏览器和 Node 里都能跑，好和 Python 结果对账）。
// 浏览器里按普通脚本引入（make_html.py 会把它内联进看板），
// Node 里 require 后取 globalThis.TTSheet。
(function (root) {
  'use strict';

  function decodeXml(s) {
    if (s.indexOf('&') < 0) return s;
    return s
      .replace(/&#x([0-9a-fA-F]+);/g, (m, h) => String.fromCharCode(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (m, d) => String.fromCharCode(parseInt(d, 10)))
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }

  // 取标签属性：属性顺序不固定，只能按名字找
  function attr(tag, name) {
    const m = tag.match(new RegExp('[\\s]' + name + '\\s*=\\s*"([^"]*)"'));
    return m ? decodeXml(m[1]) : null;
  }

  // 列字母 -> 0 基索引，如 A->0, AA->26
  function colIndex(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  const utf8 = new TextDecoder('utf-8');

  // ===== zip =====
  // 只读中央目录拿条目表，用到哪条解哪条
  function zipEntries(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    const u16 = p => dv.getUint16(p, true);
    const u32 = p => dv.getUint32(p, true);
    let eocd = -1;
    const floor = Math.max(0, buf.length - 65557);
    for (let p = buf.length - 22; p >= floor; p--) {
      if (u32(p) === 0x06054b50) { eocd = p; break; }
    }
    if (eocd < 0) throw new Error('这个文件不是 xlsx（没找到 zip 结尾标记）');
    const count = u16(eocd + 10);
    let p = u32(eocd + 16);
    const out = {};
    for (let i = 0; i < count; i++) {
      if (u32(p) !== 0x02014b50) break;
      const nameLen = u16(p + 28);
      out[utf8.decode(buf.subarray(p + 46, p + 46 + nameLen))] = {
        method: u16(p + 10),
        compSize: u32(p + 20),
        localOff: u32(p + 42),
      };
      p += 46 + nameLen + u16(p + 30) + u16(p + 32);
    }
    return out;
  }

  async function entryBytes(buf, ent) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    const u16 = p => dv.getUint16(p, true);
    const u32 = p => dv.getUint32(p, true);
    const p = ent.localOff;
    if (u32(p) !== 0x04034b50) throw new Error('xlsx 内部结构损坏（zip 本地头不对）');
    // 本地头的额外区长度和中央目录可能不同，得按本地头自己算数据起点
    const start = p + 30 + u16(p + 26) + u16(p + 28);
    const data = buf.subarray(start, start + ent.compSize);
    if (ent.method === 0) return data;
    if (ent.method !== 8) throw new Error('xlsx 用了不支持的压缩方式: ' + ent.method);
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('浏览器太旧，不支持解压 xlsx（请用 Chrome / Edge 打开）');
    }
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  // ===== xlsx =====
  // Target -> zip 内路径（可能相对 xl/，也可能带前导斜杠）
  function normTarget(target) {
    const t = target.replace(/^\/+/, '');
    return t.indexOf('xl/') === 0 ? t : 'xl/' + t;
  }

  // [(工作表名, zip 内路径)]，按工作簿里的顺序。
  // 不能写死 xl/worksheets/sheet1.xml：工作表的文件名不一定是 sheet1，
  // 甚至可能没有这张（TT 后台加工过的多 sheet 文件就是 sheet2 起）。
  async function sheetPairs(buf) {
    const ents = zipEntries(buf);
    const wb = ents['xl/workbook.xml'], rels = ents['xl/_rels/workbook.xml.rels'];
    if (!wb || !rels) return [];
    const rid2target = {};
    utf8.decode(await entryBytes(buf, rels)).replace(/<Relationship\b[^>]*>/g, tag => {
      const id = attr(tag, 'Id'), target = attr(tag, 'Target');
      if (id && target) rid2target[id] = normTarget(target);
      return '';
    });
    const pairs = [];
    utf8.decode(await entryBytes(buf, wb)).replace(/<sheet\b[^>]*>/g, tag => {
      const name = attr(tag, 'name'), rid = attr(tag, 'r:id');
      if (name && rid && rid2target[rid]) pairs.push([name, rid2target[rid]]);
      return '';
    });
    return pairs;
  }

  async function sheetNames(buf) {
    return (await sheetPairs(buf)).map(p => p[0]);
  }

  // sharedStrings：<si> 下所有 <t> 拼起来；<si/> 这种空条目也要占一个位
  async function readShared(buf, ents) {
    const ent = ents['xl/sharedStrings.xml'];
    if (!ent) return [];
    const xml = utf8.decode(await entryBytes(buf, ent));
    const out = [];
    const re = /<si\b([^>]*?)(\/>|>([\s\S]*?)<\/si>)/g;
    let m;
    while ((m = re.exec(xml))) {
      if (m[2] === '/>') { out.push(''); continue; }
      let text = '';
      (m[3].match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).forEach(t => {
        const inner = t.replace(/^<t\b[^>]*>/, '').replace(/<\/t>$/, '');
        text += decodeXml(inner);
      });
      out.push(text);
    }
    return out;
  }

  // 一行里的单元格 -> {列索引: 值}。
  // <c r="A1" s="1" t="s"><v>1</v></c>：t="s" 查 sharedStrings，t="inlineStr" 取 <is>，
  // 其余（数值/布尔/公式结果）取 <v> 原文，空单元格没有 <v>。
  function rowCells(rowTag, body, shared) {
    const cells = {};
    let maxCol = -1;
    const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
    let m;
    while ((m = re.exec(body))) {
      const ref = attr(m[1], 'r');
      if (!ref) continue;
      const letters = ref.match(/^[A-Z]+/);
      if (!letters) continue;
      const i = colIndex(letters[0]);
      const t = attr(m[1], 't');
      const inner = m[2] === '/>' ? '' : m[3];
      let val = '';
      if (t === 'inlineStr') {
        (inner.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || []).forEach(x => {
          val += decodeXml(x.replace(/^<t\b[^>]*>/, '').replace(/<\/t>$/, ''));
        });
      } else {
        const v = inner.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
        if (v) val = t === 's' ? (shared[parseInt(decodeXml(v[1]), 10)] || '') : decodeXml(v[1]);
      }
      cells[i] = val;
      if (i > maxCol) maxCol = i;
    }
    return { cells, maxCol };
  }

  function parseSheet(xml, shared) {
    const parsed = {};
    let maxCol = -1;
    const re = /<row\b([^>]*?)(\/>|>([\s\S]*?)<\/row>)/g;
    let m, lastRow = 0;
    while ((m = re.exec(xml))) {
      const rAttr = attr(m[1], 'r');
      const rn = rAttr ? parseInt(rAttr, 10) : lastRow + 1;
      lastRow = rn;
      const got = rowCells(m[1], m[2] === '/>' ? '' : m[3], shared);
      parsed[rn] = got.cells;
      if (got.maxCol > maxCol) maxCol = got.maxCol;
    }
    const rnList = Object.keys(parsed).map(Number).sort((a, b) => a - b);
    if (!rnList.length) return { headers: [], rows: [] };
    const width = maxCol + 1;
    const expand = rn => Array.from({ length: width }, (_, i) => parsed[rn][i] || '');
    return { headers: expand(rnList[0]), rows: rnList.slice(1).map(expand) };
  }

  // 读一张工作表，缺省第一张。返回 {headers, rows}
  async function readXlsx(buf, sheet) {
    const ents = zipEntries(buf);
    const pairs = await sheetPairs(buf);
    let target;
    if (!pairs.length) {
      target = 'xl/worksheets/sheet1.xml';
    } else if (!sheet) {
      target = pairs[0][1];
    } else {
      const hit = pairs.filter(p => p[0] === sheet)[0];
      if (!hit) {
        throw new Error('xlsx 里没有工作表"' + sheet + '"（有：' + pairs.map(p => p[0]).join('、') + '）');
      }
      target = hit[1];
    }
    if (!ents[target]) throw new Error('xlsx 里没找到工作表文件 ' + target);
    const shared = await readShared(buf, ents);
    return parseSheet(utf8.decode(await entryBytes(buf, ents[target])), shared);
  }

  root.TTSheet = { sheetNames: sheetNames, readXlsx: readXlsx };
})(typeof globalThis !== 'undefined' ? globalThis : this);
