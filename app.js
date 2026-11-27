'use strict';
/* 起重机性能表通用校验 + ST 代码生成
   通用原则：工具打开时什么都不知道；上传 Excel 后才解析、校验、生成。
   数据模型：{ rawScopes:{表名:[幅度点]}, sheets:{ 表名:{ PRG:{label,titleRow,ampCol,rowMap,meta,orig:{Full,Half,Fast},filled:{Full,Half,Fast},extraGroups} } } }
   rows: { '13': {'3.0':100000,...} } 缺失填 0，文本格原样保留字符串
   组别：每块的臂长按“递减处分段”自动分出多组；第 1 组→全伸数组，第 2 组→快伸数组，第 3 组→Fast 数组，其余组程序无对应数组，只记录、不出代码。 */
const LS_KEY = 'jt_perf_upload_v2';   // localStorage：上次上传的 Excel 解析结果（v1 旧结构已废弃，不再读）
const CFG_KEY = 'jt_perf_cfg_v1';     // localStorage：程序配置
const PROG_KEY = 'jt_prog_v1';        // localStorage：上次上传的程序解析结果（POU 表数据）
const DEFAULT_CFG = { arrFull: 'Full_Weight', arrHalf: 'Half_Weight', arrFast: 'Fast_Weight', arrArm: 'Arm_Length_Weight_Point', arrScope: 'Scope_Weight_Point', guard: 'NOT TeleOut_Fast_Control_LMI', hy: 'HY', cap: 30, prefix: {} };

function loadCfg() {
  try {
    const raw = typeof localStorage !== 'undefined' && localStorage.getItem(CFG_KEY);
    if (raw) { const o = JSON.parse(raw); return Object.assign({}, DEFAULT_CFG, o, { prefix: Object.assign({}, DEFAULT_CFG.prefix, o.prefix || {}) }); }
  } catch (e) { /* 读不到就用默认 */ }
  return Object.assign({}, DEFAULT_CFG, { prefix: {} });
}
function saveCfg() { try { localStorage.setItem(CFG_KEY, JSON.stringify(state.cfg)); } catch (e) { /* 存不下不影响使用 */ } }

let state = { data: null, dataLabel: '', sheet: null, prg: null, src: 'orig', grp: 'Full', boom: null, fileMeta: null, edits: {}, cfg: loadCfg(), program: null };
let issues = [];   // 显示的问题列表 = 趋势校验 + 程序对比（合并后）
let issuesTrend = []; // 只含趋势校验的问题（不含程序对比）；issues = mergeProgramDiffs(issuesTrend)
let issuesTrendOrig = []; // 上传 Excel 时的趋势问题（撤销在线修改时恢复；与 dataOrig 对应）
let issuesOrig = null;  // 载入/上传时的原始校验结果（撤销在线修改时恢复）
let dataOrig = null;    // 载入/上传时的数据深拷贝（撤销在线修改时恢复）

/* ---------------- 工具 ---------------- */
const $ = id => document.getElementById(id);
const rk = r => Number(r).toFixed(1); // 幅度键：'3.0'、'10.0'（JS 的 String(3.0) 是 '3'，不能直接用）
const fmtR = x => Number.isInteger(x) ? x.toFixed(1) : String(x); // 幅度显示：6 -> '6.0'
const fmtB = b => String(b).replace(/\.0$/, '');
const fmt1 = x => Number(x).toFixed(1); // 代码里的臂长数值：13.0、17.8（保留一位小数）
const cleanBoom = b => (typeof b === 'number' ? Math.round(b * 10) / 10 : b);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
function block() { return state.data.sheets[state.sheet][state.prg]; }
function grpData() { return block()[state.src][state.grp]; }
function colLetter(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; }
/* PRG 显示名：前缀 + 工况键（后移后缀用配置里的 hy） */
function prgKeyOf(meta) { return meta.letter + 'C' + String(Math.round(meta.cw)).padStart(2, '0') + (meta.back > 0 ? 'HY' : ''); }
function displayPrg(sheet, prgObj) {
  const meta = prgObj.meta;
  return (state.cfg.prefix[sheet] || '') + meta.letter + 'C' + String(Math.round(meta.cw)).padStart(2, '0') + (meta.back > 0 ? state.cfg.hy : '');
}
/* 本表快伸组第 0 行对应的臂长 = 全伸组第一个臂长（程序约定：Half_Weight[0] 与 Full 第 0 行同值） */
function halfZero() { return String(block()[state.src]['Full']['booms'][0]); }
/* 本表幅度点：Excel 主序列（按 IEweightCol 顺延补齐）+ 本块的附加幅度行，一起升序排 */
function sheetScope() {
  const raw = (state.data.rawScopes && state.data.rawScopes[state.sheet]) || [];
  const cap = parseInt(state.cfg.cap, 10);
  if (!cap || cap <= 0 || raw.length < 2) return raw.slice();
  const out = raw.slice();
  const step = out[out.length - 1] - out[out.length - 2];
  if (!(step > 0)) return out;
  let guard = 0;
  while (out.length < cap + 1 && guard++ < 50) out.push(Math.round((out[out.length - 1] + step) * 10) / 10);
  return out;
}
function prgScope() {
  const base = sheetScope();
  const b = block();
  const extras = new Set();
  ['Full', 'Half', 'Fast'].forEach(g => {
    const ex = b[state.src][g].extra || {};
    Object.keys(ex).forEach(k => extras.add(parseFloat(k)));
  });
  const all = base.concat([...extras]).sort((x, y) => x - y);
  return [...new Set(all.map(x => Math.round(x * 10) / 10))];
}

/* ---------------- Excel 解析 ----------------
   grid: sheet_to_json(header:1) 的二维数组，0 起下标。
   每块：幅度列含 '幅度(m)' 的行是表头行；标题行 = 表头行-1；臂长行 = 表头行+1（有序号行时 +2）。
   横向多个复本，步长 24 列：copy1 原表、copy2 右表、copy3 填充版。 */
function rowTitle(grid, row, c0, c1) {
  if (!grid[row]) return null;
  for (let c = c0; c <= c1; c++) {
    const v = grid[row][c];
    if (typeof v === 'string' && v.trim()) return v;
  }
  return null;
}
function parseBlock(grid, ampCol, headRow, off, extraOverride) {
  const titleRow = headRow - 1;
  // 标题位置不固定：全伸块在 C/D 列，半伸块在 A/B 列，取本复本列区内第一个非空字符串
  const title = rowTitle(grid, titleRow, ampCol + off, ampCol + off + 23);
  const extra = extraOverride !== undefined ? extraOverride : (grid[headRow + 1] && grid[headRow + 1][ampCol + 1 + off] === 0);
  const boomRow = extra ? headRow + 2 : headRow + 1;
  // 臂长列数不固定（新表 Fast 组 7 根、共 22 列；旧表 21 列）：沿臂长行数到第一个非数字格，最多 24 列（一个复本的步长）。
  // 21 列写死会丢最后一根臂长（如 HC38 的 Fast 46m）。
  let nBoom = 0;
  while (nBoom < 24) {
    const v = grid[boomRow] ? grid[boomRow][ampCol + 1 + off + nBoom] : null;
    if (typeof v !== 'number') break;
    nBoom++;
  }
  if (!nBoom) nBoom = 21;
  const booms = [];
  for (let i = 0; i < nBoom; i++) booms.push(grid[boomRow] ? grid[boomRow][ampCol + 1 + off + i] : null);
  const rows = []; // {radius, vals[nBoom], sheetRow}
  let r = boomRow + 1;
  while (r < grid.length) {
    const a = grid[r] ? grid[r][ampCol + off] : null;
    if (typeof a === 'string' && a.trim() === '倍率') break;
    if (a === null || a === undefined || a === '') { r++; continue; }
    const vals = [];
    for (let i = 0; i < nBoom; i++) vals.push(grid[r] ? grid[r][ampCol + 1 + off + i] : null);
    rows.push({ radius: a, vals, sheetRow: r });
    r++;
  }
  return { title, titleRow, extra, boomRow, booms, rows };
}
/* 臂长按递减处分段：13→50.7、17.7→41.2、22.5→46 是三组（每组内递增，组间回落）。
   返回每组在 21 列里的列下标；空列（null）跳过。 */
function splitGroups(booms) {
  const entries = booms.map((b, i) => ({ b: cleanBoom(b), i })).filter(e => typeof e.b === 'number');
  const groups = [];
  let cur = [];
  entries.forEach(e => {
    if (cur.length && e.b < cur[cur.length - 1].b) { groups.push(cur); cur = []; }
    cur.push(e);
  });
  if (cur.length) groups.push(cur);
  return groups.map(g => g.map(e => e.i));
}
function toGroups(copy, rawScope, groupIdxs) {
  const groups = {};
  [['Full', 0], ['Half', 1], ['Fast', 2]].forEach(([g, gi]) => {
    const idxs = groupIdxs[gi] || [];
    const blist = idxs.map(i => cleanBoom(copy.booms[i]));
    const rows = {};
    idxs.forEach((ci, k) => {
      const vals = {};
      rawScope.forEach(rad => {
        const hit = copy.rows.find(rr => rr.radius === rad);
        let v = hit ? hit.vals[ci] : 0;
        if (v === null || v === undefined) v = 0;
        vals[rk(rad)] = v;
      });
      rows[String(blist[k])] = vals;
    });
    const extra = {};
    copy.rows.forEach(rr => {
      if (typeof rr.radius === 'number' && !rawScope.includes(rr.radius)) {
        extra[rk(rr.radius)] = {};
        idxs.forEach((ci, k) => {
          const v = rr.vals[ci];
          extra[rk(rr.radius)][String(blist[k])] = (v === null || v === undefined) ? 0 : v;
        });
      }
    });
    groups[g] = { booms: blist, rows, extra };
  });
  const extraGroups = groupIdxs.slice(3).map(idxs => idxs.map(i => cleanBoom(copy.booms[i])));
  return { groups, extraGroups };
}
/* 幅度重复修正：恰好一个重复值、主序列里恰好缺一个值、且缺失值是重复值在主序列里的前一项时，把第一次出现的重复行改标为缺失值 */
function fixDupRadius(copy, rawScope) {
  const radii = copy.rows.map(r => r.radius).filter(x => typeof x === 'number');
  const seen = {}, dups = [];
  radii.forEach(x => { if (seen[x]) dups.push(x); seen[x] = true; });
  if (dups.length !== 1) return null;
  const dup = dups[0];
  const missing = rawScope.filter(x => !seen[x]);
  if (missing.length === 1 && rawScope.indexOf(missing[0]) === rawScope.indexOf(dup) - 1) {
    const first = copy.rows.find(r => r.radius === dup);
    if (first) first.radius = missing[0];
    return { dup, fixed: missing[0] };
  }
  return null;
}
/* 一格两个数（如 "2100\n800"）：按同列随幅度递减的趋势取数。
   候选数里恰好一个能接上（小于上一行、大于下一行）的取它，并记一条 note；
   否则不动（继续按文本处理）。就地改写 copy.rows，调用点与 fixDupRadius 相同。 */
function fixMultiNumCells(copy, ampCol) {
  const notes = [];
  if (!copy || !copy.rows || !copy.booms) return notes;
  for (let ci = 0; ci < copy.booms.length; ci++) {
    const boom = cleanBoom(copy.booms[ci]);
    if (typeof boom !== 'number') continue;
    for (let i = 0; i < copy.rows.length; i++) {
      const raw = copy.rows[i].vals[ci];
      if (typeof raw !== 'string') continue;
      const nums = (raw.match(/\d+(?:\.\d+)?/g) || []).map(Number);
      if (nums.length < 2) continue;
      let prev = null, next = null;
      for (let k = i - 1; k >= 0; k--) { const v = copy.rows[k].vals[ci]; if (typeof v === 'number' && v !== 0) { prev = v; break; } }
      for (let k = i + 1; k < copy.rows.length; k++) { const v = copy.rows[k].vals[ci]; if (typeof v === 'number' && v !== 0) { next = v; break; } }
      const fits = nums.filter(n => (prev === null || n < prev) && (next === null || n > next));
      if (fits.length !== 1) continue;
      copy.rows[i].vals[ci] = fits[0];
      notes.push({ boom, radius: copy.rows[i].radius, raw, picked: fits[0], cell: colLetter(ampCol + 1 + ci) + (copy.boomRow + 2 + i) });
    }
  }
  return notes;
}
function titleInfo(leftTitle, rightTitle) {
  const t = String(leftTitle || '') + ' ' + String(rightTitle || '');
  const spanM = String(leftTitle || '').match(/跨距([\d.]+)m/);
  const span = spanM ? parseFloat(spanM[1]) : null;
  const cwM = String(rightTitle || leftTitle || '').match(/(\d+(?:\.\d+)?)吨配重/);
  const cw = cwM ? parseFloat(cwM[1]) : null;
  const backM = t.match(/后移(\d+)mm/);
  const back = backM ? parseInt(backM[1], 10) : null;
  return { span, cw, back };
}
function parseWorkbook(wb) {
  const data = { rawScopes: {}, sheets: {} };
  const found = [];
  const rawBlocks = []; // {sheet, c1, c2, c3, titleRow, ampCol}
  wb.SheetNames.forEach(name => {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null, raw: true });
    // 找幅度列：第一处 '幅度(m)'
    let ampCol = -1;
    outer: for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < (grid[r] || []).length; c++) {
        if (grid[r][c] === '幅度(m)') { ampCol = c; break outer; }
      }
    }
    if (ampCol < 0) return;
    const heads = [];
    for (let r = 0; r < grid.length; r++) if (grid[r] && grid[r][ampCol] === '幅度(m)') heads.push(r);
    heads.forEach(hr => {
      const c1 = parseBlock(grid, ampCol, hr, 0);
      const c2 = parseBlock(grid, ampCol, hr, 24, c1.extra);
      const c3 = parseBlock(grid, ampCol, hr, 48, c1.extra);
      rawBlocks.push({ sheet: name, c1, c2, c3, titleRow: c1.titleRow + 1, ampCol });
    });
  });
  if (!rawBlocks.length) return { data, found };
  // 全文件跨距排大小：最大的=全伸 F，最小的=半伸 H（标题里的“全伸/半伸”文字不可靠，以跨距数字为准）
  const spans = [...new Set(rawBlocks.map(b => titleInfo(b.c1.title, b.c2.title).span).filter(x => x !== null))].sort((a, b) => b - a);
  const rankOf = {};
  if (spans.length >= 2) { rankOf[spans[0]] = 'F'; rankOf[spans[spans.length - 1]] = 'H'; }
  // 每张表的主幅度序列 = 各块幅度列表里出现最多的那一个（2.5m 这种附加行只在个别块出现，不进主序列）
  const bySheet = {};
  rawBlocks.forEach(b => { (bySheet[b.sheet] = bySheet[b.sheet] || []).push(b); });
  Object.keys(bySheet).forEach(sheet => {
    const counts = {};
    bySheet[sheet].forEach(b => {
      const key = b.c1.rows.map(r => r.radius).filter(x => typeof x === 'number').sort((x, y) => x - y).join(',');
      counts[key] = (counts[key] || 0) + 1;
    });
    let best = null, bestN = 0;
    Object.keys(counts).forEach(k => { if (counts[k] > bestN || (counts[k] === bestN && best !== null && k.split(',').length > best.split(',').length)) { best = k; bestN = counts[k]; } });
    data.rawScopes[sheet] = best ? best.split(',').map(Number) : [];
  });
  rawBlocks.forEach(b => {
    const info = titleInfo(b.c1.title, b.c2.title);
    if (info.span === null || info.cw === null) return;
    const letter = rankOf[info.span] || 'F';
    const meta = { letter, cw: info.cw, back: info.back === null ? 0 : info.back };
    const prgKey = prgKeyOf(meta);
    const rawScope = data.rawScopes[b.sheet];
    const fix = fixDupRadius(b.c1, rawScope);
    const multiNotes = fixMultiNumCells(b.c1, b.ampCol);
    const groupIdxs = splitGroups(b.c1.booms);
    const g1 = toGroups(b.c1, rawScope, groupIdxs);
    const g3 = toGroups(b.c3, rawScope, groupIdxs);
    const rowMap = {};
    b.c1.rows.forEach(r => { if (typeof r.radius === 'number') rowMap[rk(r.radius)] = r.sheetRow; });
    const sheetData = data.sheets[b.sheet] = data.sheets[b.sheet] || {};
    sheetData[prgKey] = {
      label: (letter === 'F' ? '全伸 ' : '半伸 ') + info.cw + 't 后移' + (info.back === null ? '—' : info.back),
      titleRow: b.titleRow, ampCol: b.ampCol, rowMap, meta,
      groupStarts: { Full: (groupIdxs[0] || [])[0] || 0, Half: (groupIdxs[1] || [])[0] || 0, Fast: (groupIdxs[2] || [])[0] || 0 },
      extraGroups: g1.extraGroups,
      orig: g1.groups, filled: g3.groups,
    };
    found.push({ sheet: b.sheet, prg: prgKey, c1: b.c1, c2: b.c2, c3: b.c3, fix, multiNotes, titleRow: b.titleRow, span: info.span, cw: info.cw, back: meta.back, ampCol: b.ampCol, letter, spanRankSize: spans.length, groupIdxs, rawScope });
  });
  return { data, found };
}

/* ---------------- 校验（只按通用规则） ---------------- */
/* 本列正常步长：相邻两行递减值的 75 分位（下限 1000），用于判断尖峰（升、降都远大于正常步长） */
function colDecScale(vals) {
  const ds = [];
  for (let i = 1; i < vals.length; i++) {
    if (typeof vals[i] === 'number' && typeof vals[i - 1] === 'number' && vals[i - 1] !== 0 && vals[i] !== 0 && vals[i - 1] > vals[i]) ds.push(vals[i - 1] - vals[i]);
  }
  if (!ds.length) return 1000;
  ds.sort((a, b) => a - b);
  return Math.max(1000, ds[Math.floor(ds.length * 0.75)] || 1000);
}

function validate(found) {
  const list = [];
  const add = (level, sheet, prg, cell, text) => list.push({ level, sheet, prg, cell, text });
  // 填充版系数：先算全文件最常见的比值，再逐块比
  const ratioCounts = {};
  const blockTop = new Map();
  found.forEach(f => {
    const fac = {};
    f.c3.rows.forEach(r3 => {
      const r1 = f.c1.rows.find(r => r.radius === r3.radius);
      if (!r1) return;
      r3.vals.forEach((v, i) => {
        const o = r1.vals[i];
        if (typeof v === 'number' && typeof o === 'number' && o > 0 && v > 0) {
          const k = Math.round(v / o * 1000) / 1000;
          fac[k] = (fac[k] || 0) + 1;
        }
      });
    });
    const top = Object.entries(fac).sort((a, b) => b[1] - a[1])[0];
    if (top) { blockTop.set(f, { ratio: parseFloat(top[0]), n: top[1] }); ratioCounts[top[0]] = (ratioCounts[top[0]] || 0) + top[1]; }
  });
  let modeRatio = null, modeN = 0;
  Object.keys(ratioCounts).forEach(k => { if (ratioCounts[k] > modeN) { modeN = ratioCounts[k]; modeRatio = parseFloat(k); } });
  // 左表标题配重：按右表配重分组，组内众数是惯例（如左 13t/右 14t 是整文件惯例）；偏离组内众数的才报
  const leftByRight = {};
  found.forEach(f => {
    const lcw = (String(f.c1.title || '').match(/(\d+(?:\.\d+)?)吨配重/) || [])[1];
    const rcw = (String(f.c2.title || '').match(/(\d+(?:\.\d+)?)吨配重/) || [])[1];
    if (lcw && rcw) { (leftByRight[rcw] = leftByRight[rcw] || {})[lcw] = ((leftByRight[rcw] || {})[lcw] || 0) + 1; }
  });
  const leftMode = {};
  Object.keys(leftByRight).forEach(rcw => {
    let best = null, bestN = 0;
    Object.keys(leftByRight[rcw]).forEach(l => { if (leftByRight[rcw][l] > bestN) { best = l; bestN = leftByRight[rcw][l]; } });
    leftMode[rcw] = best;
  });
  found.forEach(f => {
    const tag = f.sheet + ' · ' + f.prg + '（R' + f.titleRow + '）';
    const c1 = f.c1;
    // 标题文字与跨距不符（标题写全伸，但跨距是文件里最小的，实为半伸；反之亦然）
    const titleStr = String(c1.title || '');
    if (f.spanRankSize >= 2) {
      const kw = titleStr.indexOf('半伸') >= 0 ? 'H' : titleStr.indexOf('全伸') >= 0 ? 'F' : null;
      if (kw && kw !== f.letter) add('warn', f.sheet, f.prg, colLetter(f.ampCol + 2) + f.titleRow, tag + ' 标题写“' + (kw === 'F' ? '全伸' : '半伸') + '”，但跨距 ' + f.span + 'm 是文件里' + (f.letter === 'H' ? '最小' : '最大') + '的，应为“' + (f.letter === 'F' ? '全伸' : '半伸') + '”');
    }
    // 左表标题配重偏离同系列惯例（PRG 按右表取）
    const lcw = (String(c1.title || '').match(/(\d+(?:\.\d+)?)吨配重/) || [])[1];
    const rcw = (String(f.c2.title || '').match(/(\d+(?:\.\d+)?)吨配重/) || [])[1];
    if (lcw && rcw && leftMode[rcw] && lcw !== leftMode[rcw]) add('warn', f.sheet, f.prg, '标题行', tag + ' 左表标题配重 ' + lcw + 't，与同系列其他块的 ' + leftMode[rcw] + 't 不一致（右表 ' + rcw + 't，PRG 按右表取）');
    // 配重分量和
    const cwM = titleStr.match(/（基本配重0吨\+([^）]+)/);
    const tonM = titleStr.match(/(\d+(?:\.\d+)?)吨配重/);
    if (cwM && tonM) {
      const sum = cwM[1].split('+').reduce((s, x) => s + (parseFloat(x) || 0), 0);
      if (Math.abs(sum - parseFloat(tonM[1])) > 0.01) add('warn', f.sheet, f.prg, '标题行', tag + ' 配重标题 ' + tonM[1] + 't ≠ 分量和 ' + sum + 't');
    }
    // 一格两个数：已按同列递减趋势取过数（只作提示，不是差异）
    (f.multiNotes || []).forEach(n => add('warn', f.sheet, f.prg, n.cell, tag + ' 臂长 ' + n.boom + 'm、幅度 ' + fmtR(n.radius) + 'm 一格两个数 ' + JSON.stringify(n.raw) + '，已按同列递减趋势取 ' + n.picked + '，如与纸质表不符请改回'));
    // 幅度重复 / 缺失
    if (f.fix) add('warn', f.sheet, f.prg, 'B列', tag + ' 幅度 ' + fmtR(f.fix.dup) + 'm 出现两次、' + fmtR(f.fix.fixed) + 'm 缺失，已按趋势把第一行 ' + fmtR(f.fix.dup) + ' 当作 ' + fmtR(f.fix.fixed) + ' 处理');
    else {
      const radii = c1.rows.map(r => r.radius).filter(x => typeof x === 'number');
      const seen = {};
      radii.forEach(x => { if (seen[x]) add('warn', f.sheet, f.prg, 'B列', tag + ' 幅度 ' + x + 'm 重复'); seen[x] = true; });
      const missing = f.rawScope.filter(x => !seen[x]);
      missing.forEach(x => add('warn', f.sheet, f.prg, 'B列', tag + ' 幅度 ' + x + 'm 缺失（主序列里有，表里没有）'));
    }
    // 序号行
    if (c1.extra) add('warn', f.sheet, f.prg, '序号行', tag + ' 多出一行 0,1,2…序号行，全表仅个别块有，格式不一致');
    // 附加幅度行（不在主序列内的幅度，如 2.5m）
    const extraRadii = c1.rows.map(r => r.radius).filter(x => typeof x === 'number' && !f.rawScope.includes(x));
    if (extraRadii.length) add('info', f.sheet, f.prg, '', tag + ' 有 ' + extraRadii.map(fmtR).join('/') + 'm 幅度行（不在主幅度序列内），ST 会作为附加列生成');
    // 逐列检查（只查进程序的三组：全伸组 + 快伸组 + Fast 组）
    const colIdxs = (f.groupIdxs[0] || []).concat(f.groupIdxs[1] || [], f.groupIdxs[2] || []);
    colIdxs.forEach(ci => {
      const boom = cleanBoom(c1.booms[ci]);
      let prev = null, prevR = null, prevCell = null, prevPrev = null, prevI = -1;
      const colVals = c1.rows.map(r => ({ rad: r.radius, v: r.vals[ci] }));
      const present = colVals.filter(x => x.v !== null && x.v !== undefined);
      const minI = present.length ? colVals.indexOf(present[0]) : -1;
      const maxI = present.length ? colVals.indexOf(present[present.length - 1]) : -1;
      const scale = colDecScale(colVals.map(x => x.v));
      colVals.forEach((x, i) => {
        const cell = colLetter(f.ampCol + 1 + ci) + (c1.boomRow + 2 + i);
        if (typeof x.v === 'string') {
          let extra = '';
          const r2 = f.c2.rows.find(r => r.radius === x.rad);
          if (r2 && r2.vals[ci] !== null && r2.vals[ci] !== undefined) extra = '，右表同格为 ' + JSON.stringify(r2.vals[ci]);
          add('warn', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + x.rad + 'm 一格内有文本 ' + JSON.stringify(x.v) + '，不是数字' + extra + '，请核对纸质表');
          return;
        }
        if (x.v === null || x.v === undefined) {
          if (minI >= 0 && i > minI && i < maxI) add('warn', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + x.rad + 'm 空白空洞');
          return;
        }
        if (x.v % 100 !== 0) add('warn', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + x.rad + 'm = ' + x.v + '，不是 100 的倍数（请核对纸质表）');
        // 孤立低点：上一格比它上下两格都低（错数常是孤立一格，邻格是好的锚点）；但若本格是尖峰（升、降都远大于本列正常步长），则上一格不是低点
        let prevIsDip = false;
        if (prev !== null && prevPrev !== null && i === prevI + 1) {
          const dip = Math.min(prevPrev, x.v) - prev;
          if (dip > 0) {
            const nextV = (i + 1 < colVals.length) ? colVals[i + 1].v : null;
            const spike = typeof nextV === 'number' && nextV !== 0 && (x.v - prev) > 3 * scale && (x.v - nextV) > 3 * scale;
            if (typeof nextV === 'number' && nextV !== 0 && !spike) {
              prevIsDip = true;
              if (dip > 1000) add('warn', f.sheet, f.prg, prevCell, tag + ' 臂长 ' + boom + 'm、幅度 ' + prevR + 'm = ' + prev + '，比上下两格（' + prevPrev + '/' + x.v + '）都低 ' + dip + '，疑似孤立错数，请核对纸质表');
              else add('warn', f.sheet, f.prg, prevCell, tag + ' 臂长 ' + boom + 'm、幅度 ' + prevR + 'm = ' + prev + '，比上下两格（' + prevPrev + '/' + x.v + '）都低 ' + dip + '，疑似孤立错数，请核对纸质表');
            }
          }
        }
        if (prev !== null && x.v > prev && !prevIsDip) add('warn', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + prevR + '→' + x.rad + 'm 上升 ' + prev + '→' + x.v + '，违反递减规律');
        prevPrev = prev; prev = x.v; prevR = x.rad; prevCell = cell; prevI = i;
      });
    });
    // 首行（最小幅度）检查：同幅度下臂长越长载荷应越小；首行值与次行值比值异常（只查原表）
    if (c1.rows.length >= 2) {
      const r0 = c1.rows[0].radius, r1 = c1.rows[1].radius;
      const groups = [
        { name: 'Full', idxs: f.groupIdxs[0] || [] },
        { name: 'Half', idxs: ((f.groupIdxs[0] && f.groupIdxs[0][0] !== undefined) ? [f.groupIdxs[0][0]] : []).concat(f.groupIdxs[1] || []) },
        { name: 'Fast', idxs: f.groupIdxs[2] || [] },
      ];
      groups.forEach(g => {
        let prev = null, prevB = null;
        g.idxs.forEach(ci => {
          const v = c1.rows[0].vals[ci];
          const boom = cleanBoom(c1.booms[ci]);
          if (typeof v !== 'number' || v === 0) return;
          if (prev !== null && v > prev) {
            const cell = colLetter(f.ampCol + 1 + ci) + (c1.boomRow + 2);
            if (v >= 2 * prev) add('warn', f.sheet, f.prg, cell, tag + ' 同幅度 ' + fmtR(r0) + 'm 下，臂长 ' + boom + 'm = ' + v + '，比更短臂长 ' + prevB + 'm 的 ' + prev + ' 高出 ' + (v / prev).toFixed(1) + ' 倍，臂长越长载荷应越小，疑似多打了 0，请核对纸质表');
            else add('warn', f.sheet, f.prg, cell, tag + ' 同幅度 ' + fmtR(r0) + 'm 下，臂长 ' + boom + 'm = ' + v + ' 略高于更短臂长 ' + prevB + 'm 的 ' + prev + '，请核对纸质表');
          }
          prev = v; prevB = boom;
        });
      });
      colIdxs.forEach(ci => {
        const v0 = c1.rows[0].vals[ci], v1 = c1.rows[1].vals[ci];
        if (typeof v0 !== 'number' || v0 === 0 || typeof v1 !== 'number' || v1 <= 0 || v0 <= 2 * v1) return;
        const boom = cleanBoom(c1.booms[ci]);
        const cell = colLetter(f.ampCol + 1 + ci) + (c1.boomRow + 2);
        const dupI = list.findIndex(x => x.sheet === f.sheet && x.prg === f.prg && x.cell === cell && x.text.indexOf('同幅度 ') >= 0);
        if (dupI >= 0) list.splice(dupI, 1); // 比值已报的格，不再重复报跨臂长问题
        add('warn', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + fmtR(r0) + 'm = ' + v0 + '，是下一行 ' + fmtR(r1) + 'm（' + v1 + '）的 ' + (v0 / v1).toFixed(1) + ' 倍，疑似多打了 0，请核对纸质表');
      });
    }
    // 右表与原表比对（记录前 3 处明细；只比进程序的两组）
    let diffN = 0; const diffEx = [];
    f.c2.rows.forEach(r2 => {
      const r1 = c1.rows.find(r => r.radius === r2.radius);
      if (!r1) return;
      colIdxs.forEach(i => {
        const v = r2.vals[i], o = r1.vals[i];
        if (v !== null && v !== undefined && o !== null && o !== undefined && v !== o) {
          diffN++;
          if (diffEx.length < 3) diffEx.push('臂长 ' + cleanBoom(c1.booms[i]) + 'm、幅度 ' + r2.radius + 'm：原表 ' + JSON.stringify(o) + '、右表 ' + JSON.stringify(v));
        }
      });
    });
    if (diffN) add('warn', f.sheet, f.prg, '', tag + ' 右表与原表重叠格有 ' + diffN + ' 处数值不同（' + diffEx.join('；') + '）');
    // 右表臂长网格（只有右表真的有臂长时才比；单复本文件右表全是空，不报）
    const g1 = c1.booms.slice(0, 3).map(cleanBoom).join('/');
    const g2 = f.c2.booms.slice(0, 3).map(cleanBoom).join('/');
    const hasC2 = f.c2.booms && f.c2.booms.some(b => b !== null && b !== undefined);
    if (hasC2 && g1 !== g2) add('warn', f.sheet, f.prg, '臂长行', tag + ' 右表臂长网格与左表不同：左 ' + g1 + '，右 ' + g2 + '（短约 0.1m）');
    // 填充版系数与全文件主要比值不一致
    const bt = blockTop.get(f);
    if (bt && modeRatio !== null && Math.abs(bt.ratio - modeRatio) > 0.05) add('warn', f.sheet, f.prg, '', tag + ' 填充版/原表比值主要为 ' + bt.ratio + '（全文件主要为 ' + modeRatio + '），共 ' + bt.n + ' 格');
  });
  return list;
}

/* ---------------- ST 生成 ---------------- */
function boomList() {
  const g = grpData();
  if (state.grp === 'Half') return [halfZero()].concat(g.booms.map(String));
  return g.booms.map(String);
}
function boomRow(boom) {
  const g = grpData();
  if (state.grp === 'Half' && String(boom) === halfZero()) return block()[state.src]['Full']['rows'][halfZero()];
  return g.rows[String(boom)];
}
function boomIndex(boom) {
  if (state.grp === 'Full') return grpData().booms.map(String).indexOf(String(boom));
  if (state.grp === 'Half') return String(boom) === halfZero() ? 0 : grpData().booms.map(String).indexOf(String(boom)) + 1;
  return grpData().booms.map(String).indexOf(String(boom));
}
function arrName() { return state.grp === 'Full' ? state.cfg.arrFull : state.grp === 'Half' ? state.cfg.arrHalf : state.cfg.arrFast; }
function cellVal(grp, boom, rad) {
  const g = block()[state.src][grp];
  const key = rk(rad);
  const isHalfZero = grp === 'Half' && String(boom) === halfZero();
  const tg = isHalfZero ? block()[state.src]['Full'] : g;
  const ex = tg.extra && tg.extra[key];
  if (ex) { const v = ex[String(boom)]; return (v === null || v === undefined) ? 0 : v; }
  const row = tg.rows[String(boom)];
  const v = row ? row[key] : 0;
  return (v === null || v === undefined) ? 0 : v;
}
function setCell(grp, boom, rad, v) {
  const g = block()[state.src][grp];
  const key = rk(rad);
  const isHalfZero = grp === 'Half' && String(boom) === halfZero();
  const tg = isHalfZero ? block()[state.src]['Full'] : g;
  if (tg.extra && tg.extra[key]) { tg.extra[key][String(boom)] = v; return; }
  const row = tg.rows[String(boom)];
  if (row) row[key] = v;
}
function excelCellFor(grp, boom, rad) {
  const b = block();
  const bl = b[state.src][grp].booms.map(String);
  let ci = bl.indexOf(String(boom));
  if (grp === 'Half') ci = String(boom) === halfZero() ? 0 : b.groupStarts.Half + ci;
  if (grp === 'Fast') ci = b.groupStarts.Fast + ci;
  const col = colLetter(b.ampCol + 1 + ci);
  const row = b.rowMap ? b.rowMap[rk(rad)] : null;
  return col + (row !== null && row !== undefined ? row + 1 : '?');
}
/* 重算某一臂长列的校验问题（与 validate() 同一套文本格式；只对原表重算） */
function columnIssues(grp, boom) {
  const out = [];
  const tag = state.sheet + ' · ' + state.prg + '（R' + block().titleRow + '）';
  const bm = cleanBoom(parseFloat(boom));
  const radii = prgScope();
  const colVals = radii.map(r => ({ rad: r, v: cellVal(grp, boom, r) }));
  const present = colVals.filter(x => typeof x.v === 'number' && x.v !== 0);
  const minI = present.length ? colVals.indexOf(present[0]) : -1;
  const maxI = present.length ? colVals.indexOf(present[present.length - 1]) : -1;
  const scale = colDecScale(colVals.map(x => x.v));
  let prev = null, prevR = null, prevCell = null, prevPrev = null, prevI = -1;
  colVals.forEach((x, i) => {
    const cell = excelCellFor(grp, boom, x.rad);
    if (typeof x.v === 'string') {
      out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + x.rad + 'm 一格内有文本 ' + JSON.stringify(x.v) + '，不是数字，请核对纸质表' });
      return;
    }
    if ((x.v === 0 || x.v === null || x.v === undefined) && minI >= 0 && i > minI && i < maxI) {
      out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + x.rad + 'm 空白空洞' });
      return;
    }
    if (typeof x.v !== 'number') return;
    if (x.v % 100 !== 0) out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + x.rad + 'm = ' + x.v + '，不是 100 的倍数（请核对纸质表）' });
    let prevIsDip = false;
    if (prev !== null && prevPrev !== null && i === prevI + 1) {
      const dip = Math.min(prevPrev, x.v) - prev;
      if (dip > 0) {
        const nextV = (i + 1 < colVals.length) ? colVals[i + 1].v : null;
        const spike = typeof nextV === 'number' && nextV !== 0 && (x.v - prev) > 3 * scale && (x.v - nextV) > 3 * scale;
        if (typeof nextV === 'number' && nextV !== 0 && !spike) {
          prevIsDip = true;
          if (dip > 1000) out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell: prevCell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + prevR + 'm = ' + prev + '，比上下两格（' + prevPrev + '/' + x.v + '）都低 ' + dip + '，疑似孤立错数，请核对纸质表' });
          else out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell: prevCell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + prevR + 'm = ' + prev + '，比上下两格（' + prevPrev + '/' + x.v + '）都低 ' + dip + '，疑似孤立错数，请核对纸质表' });
        }
      }
    }
    if (prev !== null && x.v > prev && !prevIsDip) out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + prevR + '→' + x.rad + 'm 上升 ' + prev + '→' + x.v + '，违反递减规律' });
    prevPrev = prev; prev = x.v; prevR = x.rad; prevCell = cell; prevI = i;
  });
  return out;
}
/* 首行（最小幅度）检查（与 validate() 同一套规则；只对原表重算） */
function firstRowIssues(grp) {
  const out = [];
  const tag = state.sheet + ' · ' + state.prg + '（R' + block().titleRow + '）';
  const scope = sheetScope();
  if (scope.length < 2) return out;
  const r0 = scope[0], r1 = scope[1];
  const b = block()[state.src][grp];
  const booms = grp === 'Half' ? [halfZero()].concat(b.booms.map(String)) : b.booms.map(String);
  const ratioHit = {};
  booms.forEach(bm => {
    if (grp === 'Half' && bm === halfZero()) return; // 13m 列是全伸组的，已在 Full 里查过，不重复报
    const v0 = cellVal(grp, bm, r0), v1 = cellVal(grp, bm, r1);
    if (typeof v0 === 'number' && v0 !== 0 && typeof v1 === 'number' && v1 > 0 && v0 > 2 * v1) ratioHit[bm] = { v0, v1 };
  });
  let prev = null, prevB = null;
  booms.forEach(bm => {
    const v = cellVal(grp, bm, r0);
    const boom = cleanBoom(parseFloat(bm));
    if (typeof v !== 'number' || v === 0) return;
    if (ratioHit[bm]) {
      out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell: excelCellFor(grp, bm, r0), text: tag + ' 臂长 ' + boom + 'm、幅度 ' + fmtR(r0) + 'm = ' + ratioHit[bm].v0 + '，是下一行 ' + fmtR(r1) + 'm（' + ratioHit[bm].v1 + '）的 ' + (ratioHit[bm].v0 / ratioHit[bm].v1).toFixed(1) + ' 倍，疑似多打了 0，请核对纸质表' });
    } else if (prev !== null && v > prev) {
      if (v >= 2 * prev) out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell: excelCellFor(grp, bm, r0), text: tag + ' 同幅度 ' + fmtR(r0) + 'm 下，臂长 ' + boom + 'm = ' + v + '，比更短臂长 ' + prevB + 'm 的 ' + prev + ' 高出 ' + (v / prev).toFixed(1) + ' 倍，臂长越长载荷应越小，疑似多打了 0，请核对纸质表' });
      else out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell: excelCellFor(grp, bm, r0), text: tag + ' 同幅度 ' + fmtR(r0) + 'm 下，臂长 ' + boom + 'm = ' + v + ' 略高于更短臂长 ' + prevB + 'm 的 ' + prev + '，请核对纸质表' });
    }
    prev = v; prevB = boom;
  });
  return out;
}
function applyEdit(grp, boom, rad, newVal) {
  const oldVal = cellVal(grp, boom, rad);
  if (oldVal === newVal) return;
  setCell(grp, boom, rad, newVal);
  const key = [state.sheet, state.prg, state.src, grp, String(boom), String(rad)].join('|');
  state.edits[key] = { grp, boom: String(boom), rad, old: oldVal, new: newVal };
  if (state.src === 'orig') {
    const needle = '臂长 ' + cleanBoom(parseFloat(boom)) + 'm、幅度 ';
    issuesTrend = issuesTrend.filter(i => !(i.sheet === state.sheet && i.prg === state.prg && (i.text.indexOf(needle) >= 0 || i.text.indexOf('同幅度 ') >= 0 || i.text.indexOf('是下一行 ') >= 0)));
    ['Full', 'Half', 'Fast'].forEach(g => {
      const bl = block()[state.src][g].booms.map(String);
      if (bl.includes(String(boom)) || (g === 'Half' && String(boom) === halfZero())) issuesTrend = issuesTrend.concat(columnIssues(g, boom));
    });
    ['Full', 'Half', 'Fast'].forEach(g => { issuesTrend = issuesTrend.concat(firstRowIssues(g)); });
    issues = mergeProgramDiffs(issuesTrend); // 改过的格与程序的对比跟着重算
  }
  renderValidation(); renderOut(); renderEdits(); drawErrorImage();
  computeCircleTables(); renderCircleJump();
}
function undoEdits() {
  if (!dataOrig) return;
  state.data = JSON.parse(JSON.stringify(dataOrig));
  state.edits = {};
  issuesTrend = JSON.parse(JSON.stringify(issuesTrendOrig));
  issues = mergeProgramDiffs(issuesTrend);
  issuesOrig = JSON.parse(JSON.stringify(issues));
  renderValidation(); renderAll(); renderEdits(); drawErrorImage();
  computeCircleTables(); renderCircleJump();
}
function renderEdits() {
  const el = $('editPanel'); if (!el) return;
  const keys = Object.keys(state.edits);
  if (!keys.length) { el.innerHTML = ''; return; }
  let html = '<div class="info">✏️ 在线修改 ' + keys.length + ' 格（已记入当前数据，ST 代码与圈图已同步更新；重新上传 Excel 会清空这些修改）：<br>';
  keys.forEach(k => { const e = state.edits[k]; html += '· ' + e.grp + ' ' + fmtB(parseFloat(e.boom)) + 'm / ' + e.rad + 'm：' + JSON.stringify(e.old) + ' → <b>' + e.new + '</b><br>'; });
  if (state.src !== 'orig') html += '（当前是填充版：校验结果按原表算，改填充版不触发重算，请切回原表核对）';
  html += '</div><div class="rowbtns"><button id="undoEdits">↩ 撤销全部在线修改</button></div>';
  el.innerHTML = html;
  $('undoEdits').onclick = undoEdits;
}
function stLinesFor(grp, boom) {
  const saved = state.grp; state.grp = grp;
  const idx = boomIndex(boom);
  const arr = arrName();
  state.grp = saved;
  const lines = ['\t(*' + fmt1(parseFloat(boom)) + 'm*)'];
  prgScope().forEach((rad, j) => {
    const v = cellVal(grp, boom, rad);
    if (typeof v === 'string') { lines.push('\t' + arr + '[' + idx + ',' + j + ']:= 0 ;\t(* 原表此处为 ' + JSON.stringify(v) + '，待核对，暂填0 *)'); return; }
    lines.push('\t' + arr + '[' + idx + ',' + j + ']:= ' + v + ' ;');
  });
  return lines.join('\n');
}
/* 代码面板的 HTML 版：与程序不同的行红底标出并注明程序值（只影响显示；复制按钮仍用纯文本 stLinesFor，粘贴进 CoDeSys 是干净代码） */
function stLinesHtml(grp, boom) {
  const saved = state.grp; state.grp = grp;
  const idx = boomIndex(boom);
  const arr = arrName();
  state.grp = saved;
  const scope = prgScope();
  const showDiff = state.program && state.src === 'orig';
  let nDiff = 0;
  const html = ['\t(*' + fmt1(parseFloat(boom)) + 'm*)'].concat(scope.map((rad, j) => {
    const v = cellVal(grp, boom, rad);
    let ln;
    if (typeof v === 'string') ln = '\t' + arr + '[' + idx + ',' + j + ']:= 0 ;\t(* 原表此处为 ' + JSON.stringify(v) + '，待核对，暂填0 *)';
    else ln = '\t' + arr + '[' + idx + ',' + j + ']:= ' + v + ' ;';
    if (showDiff && typeof v === 'number') {
      const pv = programValue(grp, boom, rad);
      if (typeof pv === 'number' && pv !== v) {
        nDiff++;
        return '<span class="codeline-diff">' + esc(ln) + '  /* 程序 ' + pv + ' */</span>';
      }
    }
    return esc(ln);
  }));
  return { html: html.join('\n'), nDiff };
}
function fullPRGText() {
  const b = block();
  const cfg = state.cfg;
  const fullBooms = b[state.src]['Full']['booms'];
  const halfBooms = b[state.src]['Half']['booms'];
  const f0 = halfZero();
  const guard = (cfg.guard || '').trim() || 'TRUE';
  const scope = prgScope();
  const rawScope = sheetScope();
  const extraRadii = scope.filter(r => !rawScope.includes(r));
  let t = '(* PRG: ' + displayPrg(state.sheet, b) + ' · ' + b.label + ' *)\n';
  t += '(*臂长点*)\nIF ' + guard + ' THEN\n';
  fullBooms.forEach((bm, i) => { t += '\t' + cfg.arrArm + '[' + i + ']:= ' + fmt1(bm) + ';\n'; });
  t += '(*快伸工况*)\nELSE\n\t' + cfg.arrArm + '[0]:= ' + fmt1(parseFloat(f0)) + ';\n';
  halfBooms.forEach((bm, i) => { t += '\t' + cfg.arrArm + '[' + (i + 1) + ']:= ' + fmt1(bm) + ';\n'; });
  t += 'END_IF;\n\n(*幅度点*' + (extraRadii.length ? '（附加幅度点：' + extraRadii.map(fmtR).join('/') + 'm）' : '') + '*)\n';
  scope.forEach((rad, i) => { t += '\t' + cfg.arrScope + '[' + i + ']:= ' + rad.toFixed(2) + ';\n'; });
  t += '\n(*支腿' + (b.meta.letter === 'F' ? '全伸' : '半伸') + '工况*)\n';
  fullBooms.forEach(bm => { t += stLinesFor('Full', String(bm)) + '\n\n'; });
  scope.forEach((rad, j) => { t += '\t' + cfg.arrFull + '[9,' + j + ']:= 0 ;\n'; });
  t += '\n(*快伸工况性能表*)\n';
  [f0].concat(halfBooms.map(String)).forEach(bm => { t += stLinesFor('Half', bm) + '\n\n'; });
  const halfCount = halfBooms.length + 1;
  for (let r = halfCount; r <= 9; r++) { scope.forEach((rad, j) => { t += '\t' + cfg.arrHalf + '[' + r + ',' + j + ']:= 0 ;\n'; }); t += '\n'; }
  const fastBooms = b[state.src]['Fast'] ? b[state.src]['Fast']['booms'] : [];
  if (fastBooms.length) {
    t += '\n(*Fast 性能表*)\n';
    fastBooms.forEach(bm => { t += stLinesFor('Fast', String(bm)) + '\n\n'; });
    const fastCount = fastBooms.length;
    for (let r = fastCount; r <= 9; r++) { scope.forEach((rad, j) => { t += '\t' + cfg.arrFast + '[' + r + ',' + j + ']:= 0 ;\n'; }); t += '\n'; }
  }
  t += 'RETURN;\n';
  return t;
}

/* ---------------- 动态查表（插值） ---------------- */
function lookupValue(grpName, boom, rad) {
  const g = block()[state.src][grpName];
  const booms = grpName === 'Half' ? [parseFloat(halfZero())].concat(g.booms.map(Number)) : g.booms.map(Number);
  const rows = {};
  booms.forEach(b => { rows[String(b)] = (grpName === 'Half' && String(b) === halfZero()) ? block()[state.src]['Full']['rows'][halfZero()] : g.rows[String(b)]; });
  const scope = sheetScope();
  const clamp = (x, lo, hi, name) => {
    if (x < lo) return { v: lo, out: name + '小于表内最小值 ' + lo + '，按边界取值' };
    if (x > hi) return { v: hi, out: name + '大于表内最大值 ' + hi + '，按边界取值' };
    return { v: x, out: null };
  };
  const cb = clamp(boom, booms[0], booms[booms.length - 1], '臂长');
  const cr = clamp(rad, scope[0], scope[scope.length - 1], '幅度');
  const notes = [cb.out, cr.out].filter(Boolean);
  const b = cb.v, r = cr.v;
  const b0 = [...booms].reverse().find(x => x <= b);
  const b1 = booms.find(x => x >= b);
  const r0 = [...scope].reverse().find(x => x <= r);
  const r1 = scope.find(x => x >= r);
  const valAt = (bb, rr) => { const v = rows[String(bb)][rk(rr)]; return (typeof v === 'number') ? v : NaN; };
  const v00 = valAt(b0, r0), v01 = valAt(b0, r1), v10 = valAt(b1, r0), v11 = valAt(b1, r1);
  if ([v00, v01, v10, v11].some(isNaN)) return { error: '参与插值的四个表点中有非数字/异常格，请先核对该表', corners: { b0, b1, r0, r1, v00, v01, v10, v11 } };
  const fr = (r1 === r0) ? 0 : (r - r0) / (r1 - r0);
  const fb = (b1 === b0) ? 0 : (b - b0) / (b1 - b0);
  const vB0 = v00 + (v01 - v00) * fr;
  const vB1 = v10 + (v11 - v10) * fr;
  const rated = vB0 + (vB1 - vB0) * fb;
  return { rated: Math.round(rated), corners: { b0, b1, r0, r1, v00, v01, v10, v11 }, notes };
}

/* ---------------- 渲染 ---------------- */
function renderSheets() {
  const el = $('sheetBtns'); el.innerHTML = '';
  if (!state.data) return;
  Object.keys(state.data.sheets).forEach(s => {
    const b = document.createElement('button');
    b.textContent = s; if (s === state.sheet) b.className = 'on';
    b.onclick = () => { state.sheet = s; if (!state.data.sheets[s][state.prg]) state.prg = Object.keys(state.data.sheets[s])[0]; state.boom = null; renderAll(); };
    el.appendChild(b);
  });
}
function renderPrgs() {
  const el = $('prgBtns'); el.innerHTML = '';
  if (!state.data) return;
  const names = Object.keys(state.data.sheets[state.sheet] || {});
  const order = names.slice().sort((a, b) => {
    const ma = state.data.sheets[state.sheet][a].meta, mb = state.data.sheets[state.sheet][b].meta;
    return (mb.cw - ma.cw) || (ma.letter === 'F' ? -1 : 1) - (mb.letter === 'F' ? -1 : 1) || (mb.back - ma.back);
  });
  order.forEach(p => {
    const b = document.createElement('button');
    b.textContent = displayPrg(state.sheet, state.data.sheets[state.sheet][p]) + '（' + state.data.sheets[state.sheet][p].label + '）';
    if (p === state.prg) b.className = 'on';
    b.onclick = () => { state.prg = p; state.boom = null; renderAll(); };
    el.appendChild(b);
  });
}
function renderInfo() {
  if (!state.data) { $('blockInfo').innerHTML = '尚未上传 Excel，请先在第 1 节上传文件。'; const w0 = $('blockWarn'); if (w0) w0.style.display = 'none'; return; }
  const b = block();
  $('blockInfo').innerHTML = '当前：<b>' + state.sheet + ' · ' + displayPrg(state.sheet, b) + '</b>（' + b.label + '）｜数据源：<b>' + (state.src === 'orig' ? '原表' : '力限填充版') + '</b>｜Excel 标题行 R' + b.titleRow + '｜数据：' + esc(state.dataLabel);
  const w = $('blockWarn');
  const warns = issues.filter(i => i.prg === state.prg && i.sheet === state.sheet && i.level !== 'info').map(i => (i.level === 'err' ? '🔴 ' : '🟡 ') + esc(i.text));
  const notes = [];
  if (b.extraGroups && b.extraGroups.length) notes.push('本块另有第 ' + (b.extraGroups.map((g, i) => (i + 4) + ' 组（' + g.map(fmtB).join('/') + 'm）').join('、')) + '，程序里没有对应数组，未生成 ST 代码。');
  const cap = parseInt(state.cfg.cap, 10);
  const scopeLen = sheetScope().length;
  if (cap > 0 && scopeLen > cap + 1) {
    const unused = sheetScope().slice(cap + 1);
    notes.push('幅度点共 ' + scopeLen + ' 个，程序插值只用到前 ' + (cap + 1) + ' 个（IEweightCol=' + cap + '），' + unused.map(fmtR).join('/') + 'm 这几列程序里用不上。');
  }
  const all = warns.concat(notes);
  if (all.length) { w.style.display = 'block'; w.innerHTML = (warns.length ? '⚠️ 本表已知问题：<br>· ' + warns.join('<br>· ') + '<br>' : '') + (notes.length ? 'ℹ️ ' + notes.join('<br>ℹ️ ') : ''); }
  else w.style.display = 'none';
}
function renderBooms() {
  const el = $('boomBtns'); el.innerHTML = '';
  if (!state.data) return;
  boomList().forEach(b => {
    const btn = document.createElement('button');
    btn.textContent = fmtB(parseFloat(b)) + 'm（' + arrName() + '[' + boomIndex(b) + ']）';
    if (state.boom === b) btn.className = 'on';
    btn.onclick = () => { state.boom = b; renderOut(); renderBooms(); };
    el.appendChild(btn);
  });
}
function renderOut() {
  const out = $('out'), tbl = $('checkTable'), w = $('warnOut');
  if (!state.data) { out.textContent = '请先上传 Excel。'; tbl.innerHTML = ''; w.innerHTML = ''; return; }
  if (state.boom === null) { out.textContent = '请先选臂长。'; tbl.innerHTML = ''; w.innerHTML = ''; return; }
  const shown = stLinesHtml(state.grp, state.boom);
  out.innerHTML = shown.html;
  let html = '<tr><th>幅度(m)</th><th class="num">Excel 值（点格子可改）</th><th class="num">进程序</th><th>备注</th></tr>';
  prgScope().forEach(rad => {
    const v = cellVal(state.grp, state.boom, rad);
    const iss = issues.find(i => i.sheet === state.sheet && i.prg === state.prg && i.text.indexOf('臂长 ' + cleanBoom(parseFloat(state.boom)) + 'm、幅度 ' + rad + 'm') >= 0);
    const ek = [state.sheet, state.prg, state.src, state.grp, String(state.boom), String(rad)].join('|');
    const ed = state.edits[ek];
    const cls = typeof v === 'string' ? 'bad' : iss ? (iss.level === 'err' ? 'bad' : 'warncell') : (v === 0 ? 'zero' : '');
    html += '<tr><td>' + rad + '</td><td class="num ' + cls + (ed ? ' edited' : '') + '"><input class="celledit" type="text" inputmode="decimal" value="' + esc(typeof v === 'string' ? v : v) + '" data-grp="' + state.grp + '" data-boom="' + esc(String(state.boom)) + '" data-rad="' + rad + '"></td><td class="num">' + (typeof v === 'string' ? 0 : v) + '</td><td>' + (iss ? esc(iss.text.split('】').pop() || iss.text) : '') + (ed ? ' ✏️已改' : '') + '</td></tr>';
  });
  tbl.innerHTML = html;
  w.innerHTML = shown.nDiff ? '<span class="badge" style="background:#fde2e2;border-color:#f5a3a3;color:#c00">本臂长与程序有 ' + shown.nDiff + ' 处不同（红底行），复制后粘贴到 CoDeSys 即覆盖为 Excel 的值</span>' : '';
}
/* 分段复制：当前 PRG 每个臂长一段，各自带复制按钮 */
function renderSegments() {
  const el = $('segBtns'); if (!el) return;
  if (!state.data) { el.innerHTML = ''; return; }
  const b = block();
  let html = '';
  [['Full', '全伸组'], ['Half', '快伸组'], ['Fast', 'Fast组']].forEach(([g, label]) => {
    const booms = g === 'Half' ? [halfZero()].concat(b[state.src][g].booms.map(String)) : b[state.src][g].booms.map(String);
    booms.forEach(bm => { html += '<button data-grp="' + g + '" data-boom="' + bm + '">' + label + ' ' + fmtB(parseFloat(bm)) + 'm 📋</button>'; });
  });
  el.innerHTML = html;
}
/* 圈图：当前选中表（PRG + 组别）画成表格图，错误红圈、待确认黄圈、改过的蓝点 */
function cellIssueLevel(grp, boom, rad) {
  const v = cellVal(grp, boom, rad);
  // 程序对比（主校验）：原表里本格与程序的值不同，直接标红（文本格也算差异）
  if (state.program && state.src === 'orig') {
    const pv = programValue(grp, boom, rad);
    if (pv !== null && ((typeof v === 'number' && v !== pv) || typeof v === 'string')) return 'err';
  }
  if (typeof v === 'string') return 'warn'; // 文本格：没程序对比时只是提示
  const vals = prgScope().map(r => cellVal(grp, boom, r));
  const idx = prgScope().indexOf(rad);
  const scale = colDecScale(vals);
  // 孤立低点：本格比上下两格都低（邻格是好的锚点，孤立一格错数）；但若下一格是尖峰（升、降都远大于本列正常步长），则本格不是低点
  if (typeof v === 'number' && v !== 0 && idx > 0 && idx < vals.length - 1) {
    const above = vals[idx - 1], below = vals[idx + 1];
    if (typeof above === 'number' && above !== 0 && typeof below === 'number' && below !== 0) {
      const dip = Math.min(above, below) - v;
      if (dip > 0) {
        const d2 = vals[idx + 2];
        if (typeof d2 === 'number' && d2 !== 0) {
          const spike = (below - v) > 3 * scale && (below - d2) > 3 * scale;
          if (!spike) return 'warn'; // 孤立低点只是提示
        }
      }
    }
  }
  let prev = null, prevJ = -1;
  for (let i = 0; i < idx; i++) { if (typeof vals[i] === 'number' && vals[i] !== 0) { prev = vals[i]; prevJ = i; } }
  if (typeof v === 'number' && v !== 0 && prev !== null && v > prev) {
    // 上一格若是孤立低点，本格的“上升”是假的（低点本身会报）；若本格是尖峰，上升是真的
    let prevIsDip = false;
    if (prevJ > 0) {
      const pp = vals[prevJ - 1];
      if (typeof pp === 'number' && pp !== 0 && prev < Math.min(pp, v)) {
        const d2 = vals[idx + 1];
        if (typeof d2 === 'number' && d2 !== 0) {
          const spike = (v - prev) > 3 * scale && (v - d2) > 3 * scale;
          if (!spike) prevIsDip = true;
        }
      }
    }
    if (!prevIsDip) return 'warn'; // 上升违反递减只是提示
  }
  if (typeof v === 'number' && v !== 0 && v % 100 !== 0) return 'warn';
  if (v === 0) {
    let hasBefore = false, hasAfter = false;
    for (let i = 0; i < idx; i++) { if (typeof vals[i] === 'number' && vals[i] !== 0) hasBefore = true; }
    for (let i = idx + 1; i < vals.length; i++) { if (typeof vals[i] === 'number' && vals[i] !== 0) hasAfter = true; }
    if (hasBefore && hasAfter) return 'warn'; // 空白空洞只是提示
  }
  // 首行检查（与 validate() 同规则）：最小幅度行
  const scope0 = sheetScope();
  if (scope0.length >= 2 && rad === scope0[0] && typeof v === 'number' && v !== 0) {
    const r1 = scope0[1];
    const v1 = cellVal(grp, boom, r1);
    if (typeof v1 === 'number' && v1 > 0 && v > 2 * v1) return 'warn'; // 首行疑似多打 0 只是提示
    const b = block()[state.src][grp];
    const booms = grp === 'Half' ? [halfZero()].concat(b.booms.map(String)) : b.booms.map(String);
    const idxB = booms.indexOf(String(boom));
    if (idxB > 0) {
      for (let j = idxB - 1; j >= 0; j--) {
        const pv = cellVal(grp, booms[j], rad);
        if (typeof pv === 'number' && pv !== 0) {
          if (v > pv) return 'warn'; // 同幅度臂越长值越大只是提示
          break;
        }
      }
    }
  }
  return null;
}
/* 扫一遍全部表×2 组，记下哪些表有圈（供跳转栏用；只在数据变化时算一次） */
function computeCircleTables() {
  if (!state.data) { state.circleTables = []; return; }
  const saved = { sheet: state.sheet, prg: state.prg, grp: state.grp };
  const out = [];
  Object.keys(state.data.sheets).forEach(sh => {
    Object.keys(state.data.sheets[sh]).forEach(prg => {
      ['Full', 'Half', 'Fast'].forEach(g => {
        state.sheet = sh; state.prg = prg; state.grp = g;
        const b = block()[state.src][g];
        const booms = g === 'Half' ? [halfZero()].concat(b.booms.map(String)) : b.booms.map(String);
        let err = 0, warn = 0;
        booms.forEach(bm => prgScope().forEach(rad => {
          const lv = cellIssueLevel(g, bm, rad);
          if (lv === 'err') err++;
          else if (lv === 'warn') warn++;
        }));
        if (err || warn) out.push({ sheet: sh, prg, grp: g, err, warn });
      });
    });
  });
  state.sheet = saved.sheet; state.prg = saved.prg; state.grp = saved.grp;
  out.sort((a, b) => (b.err - a.err) || (b.warn - a.warn));  // 错得多的排前面
  state.circleTables = out;
}
/* 当前表没圈就跳到第一张错表（打开工具 / 重新上传时用；当前表本身有圈就不动） */
function jumpToFirstErrorTable() {
  const list = state.circleTables || [];
  if (!list.length) return;
  const cur = list.findIndex(t => t.sheet === state.sheet && t.prg === state.prg && t.grp === state.grp);
  if (cur >= 0) return;
  const t = list[0];
  state.sheet = t.sheet; state.prg = t.prg; state.grp = t.grp; state.boom = null;
}
/* 只在错表之间翻：dir=1 下一张，dir=-1 上一张，到头绕回 */
function stepErrorTable(dir) {
  const list = state.circleTables || [];
  if (!list.length) return;
  let i = list.findIndex(t => t.sheet === state.sheet && t.prg === state.prg && t.grp === state.grp);
  if (i < 0) i = dir > 0 ? -1 : 0;
  i = (i + dir + list.length) % list.length;
  const t = list[i];
  state.sheet = t.sheet; state.prg = t.prg; state.grp = t.grp; state.boom = null;
  document.querySelectorAll('#grpBtns button').forEach(b => b.classList.toggle('on', b.dataset.grp === t.grp));
  renderAll();
  const cv = $('errCanvas');
  if (cv && cv.scrollIntoView) cv.scrollIntoView({ block: 'nearest' });
}
function renderCircleJump() {
  const el = $('circleJump'); if (!el) return;
  if (!state.data) { el.innerHTML = ''; return; }
  const list = state.circleTables || [];
  const total = Object.keys(state.data.sheets).reduce((s, sh) => s + Object.keys(state.data.sheets[sh]).length, 0) * 3;
  if (!list.length) { el.innerHTML = '<span class="sub">全部 ' + total + ' 个表都没有错误格。</span>'; return; }
  const gName = g => ({ Full: '全伸组', Half: '快伸组', Fast: 'Fast组' })[g] || g;
  el.innerHTML = '<span class="sub">有圈的表（点一下跳过去看圈）：</span>' + list.map((t, i) =>
    '<button data-i="' + i + '">' + t.sheet.replace('模式', '') + ' ' + displayPrg(t.sheet, state.data.sheets[t.sheet][t.prg]) + ' ' + gName(t.grp) + ' 🔴' + t.err + ' 🟡' + t.warn + '</button>').join('');
}
function drawErrorImage() {
  const cv = $('errCanvas'); if (!cv) return;
  const ctx = cv.getContext && cv.getContext('2d'); if (!ctx) return;
  if (!state.data) {
    cv.width = 640; cv.height = 90;
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 640, 90);
    ctx.fillStyle = '#888'; ctx.font = '14px sans-serif'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
    ctx.fillText('尚未上传 Excel，圈图在这里出。', 12, 45);
    const cap0 = $('imgCount'); if (cap0) cap0.textContent = '尚未上传 Excel。';
    const nav0 = $('errNavLabel'); if (nav0) nav0.textContent = '';
    return;
  }
  const b = block();
  const g = b[state.src][state.grp];
  const booms = state.grp === 'Half' ? [halfZero()].concat(g.booms.map(String)) : g.booms.map(String);
  const radii = prgScope();
  let nErr = 0, nWarn = 0;
  const radW = 62, colW = 96, rowH = 26, headH = 30, titleH = 48, legH = 32;
  const W = radW + booms.length * colW, H = titleH + headH + radii.length * rowH + legH;
  cv.width = W; cv.height = H;
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
  ctx.fillStyle = '#111'; ctx.font = 'bold 16px sans-serif'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
  ctx.fillText(state.sheet + ' · ' + displayPrg(state.sheet, b) + '（' + b.label + '）· ' + ({ Full: '全伸组', Half: '快伸组', Fast: 'Fast组' })[state.grp] + ' · ' + (state.src === 'orig' ? '原表' : '填充版'), 10, 18);
  ctx.fillStyle = '#666'; ctx.font = '12px sans-serif';
  ctx.fillText('数据：' + state.dataLabel + '　' + new Date().toLocaleDateString(), 10, 38);
  ctx.font = '13px sans-serif'; ctx.textAlign = 'center';
  ctx.fillStyle = '#f5f5f5'; ctx.fillRect(0, titleH, W, headH);
  ctx.fillStyle = '#333'; ctx.fillText('幅度(m)', radW / 2, titleH + headH / 2);
  booms.forEach((bm, i) => { ctx.fillText(fmtB(parseFloat(bm)) + 'm', radW + i * colW + colW / 2, titleH + headH / 2); });
  radii.forEach((rad, j) => {
    const y = titleH + headH + j * rowH;
    if (j % 2 === 0) { ctx.fillStyle = '#fafafa'; ctx.fillRect(0, y, W, rowH); }
    ctx.fillStyle = '#333'; ctx.fillText(String(rad), radW / 2, y + rowH / 2);
    booms.forEach((bm, i) => {
      const v = cellVal(state.grp, bm, rad);
      const cx = radW + i * colW + colW / 2, cy = y + rowH / 2;
      const lvl = cellIssueLevel(state.grp, bm, rad);
      if (lvl === 'err') nErr++;
      else if (lvl === 'warn') nWarn++;
      const ek = [state.sheet, state.prg, state.src, state.grp, String(bm), String(rad)].join('|');
      ctx.fillStyle = typeof v === 'string' ? '#c00' : v === 0 ? '#bbb' : '#222';
      ctx.fillText(typeof v === 'string' ? v : v, cx, cy);
      if (lvl === 'err') { ctx.strokeStyle = '#d00'; ctx.lineWidth = 2.5; ctx.beginPath(); ctx.ellipse(cx, cy, 36, 11, 0, 0, Math.PI * 2); ctx.stroke(); }
      else if (lvl === 'warn') { ctx.strokeStyle = '#d9a400'; ctx.lineWidth = 2; ctx.beginPath(); ctx.ellipse(cx, cy, 36, 11, 0, 0, Math.PI * 2); ctx.stroke(); }
      if (state.edits[ek]) { ctx.fillStyle = '#1677ff'; ctx.beginPath(); ctx.arc(cx + 38, cy - 9, 4, 0, Math.PI * 2); ctx.fill(); }
    });
  });
  ctx.strokeStyle = '#ddd'; ctx.lineWidth = 1;
  for (let j = 0; j <= radii.length; j++) { const y = titleH + headH + j * rowH; ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(W, y); ctx.stroke(); }
  for (let i = 0; i <= booms.length; i++) { const x = radW + i * colW; ctx.beginPath(); ctx.moveTo(x, titleH); ctx.lineTo(x, titleH + headH + radii.length * rowH); ctx.stroke(); }
  ctx.beginPath(); ctx.moveTo(radW, titleH); ctx.lineTo(radW, titleH + headH + radii.length * rowH); ctx.stroke();
  const ly = H - legH / 2;
  ctx.textAlign = 'left'; ctx.font = '12px sans-serif'; ctx.fillStyle = '#333';
  ctx.fillText('🔴 红圈=确认错误　🟡 黄圈=待确认　🔵 蓝点=在线改过的值', 10, ly);
  const cap = $('imgCount');
  if (cap) cap.textContent = (nErr || nWarn)
    ? '本表共 ' + (nErr + nWarn) + ' 处圈：🔴 确认错误 ' + nErr + ' 个，🟡 待确认 ' + nWarn + ' 个。'
    : '本表没有错误格 ✅（干净，0 圈是正常的；有圈的表见下面跳转）';
  const nav = $('errNavLabel');
  if (nav) {
    const list = state.circleTables || [];
    const i = list.findIndex(t => t.sheet === state.sheet && t.prg === state.prg && t.grp === state.grp);
    nav.textContent = list.length ? (i >= 0 ? '第 ' + (i + 1) + ' / ' + list.length + ' 张错表' : '当前表没错（共 ' + list.length + ' 张错表）') : '全部表都干净';
  }
}
function renderValidation() {
  if (!state.data) { $('validSummary').innerHTML = '尚未上传 Excel，请先在第 1 节上传文件。'; $('validList').innerHTML = ''; return; }
  const errs = issues.filter(i => i.level === 'err').length;
  const warns = issues.filter(i => i.level === 'warn').length;
  const infos = issues.filter(i => i.level === 'info').length;
  $('validSummary').innerHTML = '校验完成：🔴 确认错误 <b>' + errs + '</b> 处，🟡 待确认 <b>' + warns + '</b> 处，ℹ️ 提示 <b>' + infos + '</b> 处。红色先改，黄色核对纸质表。';
  if (!issues.length) { $('validList').innerHTML = '<div class="info">未发现问题。</div>'; return; }
  let html = '<table><tr><th>级别</th><th>表 / 工况</th><th>位置</th><th>问题</th></tr>';
  issues.slice(0, 1000).forEach(i => {
    html += '<tr class="lvl-' + i.level + '"><td>' + (i.level === 'err' ? '🔴 错误' : i.level === 'warn' ? '🟡 待确认' : 'ℹ️ 提示') + '</td><td>' + esc(i.sheet + ' · ' + i.prg) + '</td><td>' + esc(i.cell || '—') + '</td><td>' + esc(i.text) + '</td></tr>';
  });
  $('validList').innerHTML = html + '</table>' + (issues.length > 1000 ? '<div class="info">共 ' + issues.length + ' 条，仅显示前 1000 条；圈图里全部标红，分表差异见第 1 节的程序对比。</div>' : '');
}
/* 配置区：每张表一个 PRG 前缀输入框（上传后出现；改完立即反映在工况名和代码里） */
function renderPrefixBox() {
  const el = $('prefixBox'); if (!el) return;
  if (!state.data) { el.innerHTML = ''; return; }
  let html = '';
  Object.keys(state.data.sheets).forEach(sh => {
    html += '<label>“' + esc(sh) + '”的 PRG 前缀<input type="text" data-sheet="' + esc(sh) + '" value="' + esc(state.cfg.prefix[sh] || '') + '" placeholder="如 JTC100" style="width:130px"></label>';
  });
  el.innerHTML = html;
}
function syncCfgInputs() {
  const m = { cfgArrFull: 'arrFull', cfgArrHalf: 'arrHalf', cfgArrFast: 'arrFast', cfgArrArm: 'arrArm', cfgArrScope: 'arrScope', cfgGuard: 'guard', cfgHy: 'hy', cfgCap: 'cap' };
  Object.keys(m).forEach(id => { const el = $(id); if (el) el.value = state.cfg[m[id]]; });
}
function renderAll() { renderSheets(); renderPrgs(); renderInfo(); renderBooms(); renderOut(); renderSegments(); renderEdits(); drawErrorImage(); }

/* ---------------- 上传 ---------------- */
function saveUpload(fileName) {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify({ name: fileName, time: Date.now(), data: state.data, issues, issuesTrend }));
  } catch (e) { /* 存不下（如浏览器禁了本地存储）不影响使用，只是下次打开恢复不了 */ }
}
function useUploaded(wb, fileName, fileSize) {
  try {
    const prevKey = i => [i.sheet, i.prg, i.cell, i.text].join('|');
    const prevIssues = issues.slice();
    const { data, found } = parseWorkbook(wb);
    const nPrg = Object.values(data.sheets).reduce((s, sh) => s + Object.keys(sh).length, 0);
    if (!nPrg) { $('fileInfo').innerHTML = '<b>解析失败：</b>在 ' + esc(fileName) + ' 里没找到含“幅度(m)”的性能表。请确认文件是起重机性能表原文件。'; return; }
    state.data = data;
    dataOrig = JSON.parse(JSON.stringify(data));
    state.edits = {};
    state.fileMeta = { name: fileName, time: Date.now(), size: fileSize || 0 };
    state.dataLabel = '已上传：' + fileName;
    $('dataBadge').textContent = '已上传文件';
    issuesTrend = validate(found);
    issuesTrendOrig = JSON.parse(JSON.stringify(issuesTrend));
    if (state.program) rebuildProgramMap(); // Excel 换了，程序 POU 的对应关系要重算
    issues = mergeProgramDiffs(issuesTrend);
    issuesOrig = JSON.parse(JSON.stringify(issues));
    const errs = issues.filter(i => i.level === 'err').length;
    const prevErrs = prevIssues.filter(i => i.level === 'err').length;
    const fixed = prevIssues.filter(i => !issues.some(j => prevKey(j) === prevKey(i))).length;
    let delta = '';
    if (prevIssues.length) delta = '<br>与上次比：确认错误 ' + prevErrs + ' → ' + errs + ' 处' + (fixed ? '，已修复 ' + fixed + ' 处' : '') + '。';
    $('fileInfo').innerHTML = '已载入 <b>' + esc(fileName) + '</b>（' + new Date().toLocaleString() + (fileSize ? '，' + Math.round(fileSize / 1024) + ' KB' : '') + '）：解析出 ' + Object.keys(data.sheets).length + ' 张表、' + nPrg + ' 个工况块。PRG 与块的对应按标题自动识别（跨距大的=全伸、小的=半伸；配重取右表；后移>0加后缀）。已记下这次上传，下次打开页面自动恢复，不用重新传。' + delta;
    saveUpload(fileName);
    computeCircleTables(); renderCircleJump();
    jumpToFirstErrorTable();
    if (!data.sheets[state.sheet]) state.sheet = Object.keys(data.sheets)[0];
    if (!data.sheets[state.sheet][state.prg]) state.prg = Object.keys(data.sheets[state.sheet])[0];
    state.boom = null;
    renderPrefixBox();
    renderProgInfo();
    renderValidation(); renderAll();
  } catch (e) {
    $('fileInfo').innerHTML = '<b>解析出错：</b>' + esc(e.message);
  }
}
function handleFile(file) {
  if (typeof XLSX === 'undefined') { $('fileInfo').innerHTML = '<b>解析库未加载：</b>本地 xlsx.full.min.js 缺失或未加载成功，请确认它与本页面在同一目录下，再刷新页面。'; return; }
  const rd = new FileReader();
  rd.onload = e => { try { useUploaded(XLSX.read(e.target.result, { type: 'array' }), file.name, file.size); } catch (err) { $('fileInfo').innerHTML = '<b>读取失败：</b>' + esc(err.message); } };
  rd.readAsArrayBuffer(file);
}

/* ---------------- 程序文件上传：Excel 里的值 vs 程序里写的值，逐格对比，差异标红 ----------------
   两种格式都认：CoDeSys 导出的 .export（TextBlob + Name 元素定归属——实现体不一定在声明体后面，
   不能按 PROGRAM 头切段）和 PLCopenXML .xml（<pou name="Working_Table_...">）。
   数组名取配置里的 arrFull/arrHalf/arrArm/arrScope（与 ST 生成保持一致）。
   对齐方式：按（臂长，幅度）数值对齐——程序用自己的 Arm/Scope 数组找行列号，Excel 用 cellVal；
   程序的 2.5m 列与 Excel 的 2.5m 列对上，不会错位。 */
function parseProgram(text) {
  const escRe = s => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const cfg = state.cfg;
  const reFull = new RegExp(escRe(cfg.arrFull) + '\\[(\\d+),(\\d+)\\]:=\\s*(-?\\d+)\\s*;', 'g');
  const reHalf = new RegExp(escRe(cfg.arrHalf) + '\\[(\\d+),(\\d+)\\]:=\\s*(-?\\d+)\\s*;', 'g');
  const reFast = new RegExp(escRe(cfg.arrFast) + '\\[(\\d+),(\\d+)\\]:=\\s*(-?\\d+)\\s*;', 'g');
  const reArm = new RegExp(escRe(cfg.arrArm) + '\\[(\\d+)\\]:=\\s*([\\d.]+)\\s*;', 'g');
  const reScope = new RegExp(escRe(cfg.arrScope) + '\\[(\\d+)\\]:=\\s*([\\d.]+)\\s*;', 'g');
  const hasFull = new RegExp(escRe(cfg.arrFull) + '\\[\\d+,\\d+\\]:=\\s*\\d+');
  function parsePouBody(body) {
    // Arm 赋值按出现顺序分块（索引归零处分块）：第一块=全伸臂长，第二块=快伸臂长
    const blocks = []; let cur = {}, last = -1;
    [...body.matchAll(reArm)].forEach(m => {
      const i = parseInt(m[1], 10), v = parseFloat(m[2]);
      if (i <= last && Object.keys(cur).length) { blocks.push(cur); cur = {}; }
      cur[i] = v; last = i;
    });
    if (Object.keys(cur).length) blocks.push(cur);
    const scope = {};
    [...body.matchAll(reScope)].forEach(m => { scope[parseInt(m[1], 10)] = parseFloat(m[2]); });
    const full = {}, half = {}, fast = {};
    [...body.matchAll(reFull)].forEach(m => { full[m[1] + ',' + m[2]] = parseInt(m[3], 10); });
    [...body.matchAll(reHalf)].forEach(m => { half[m[1] + ',' + m[2]] = parseInt(m[3], 10); });
    [...body.matchAll(reFast)].forEach(m => { fast[m[1] + ',' + m[2]] = parseInt(m[3], 10); });
    // 臂长块按内容认组，不按出现顺序（兼容两分支和三分支）：
    // 全伸=最长的一块；快伸=以 13.0 开头的其余块；Fast=不以 13.0 开头的块。
    // V001.8 两分支：全伸 9 根、快伸 7 根；V001.9 三分支：快伸 7 根、Fast 7 根、全伸 9 根。
    let fullArm = {}, halfArm = {}, fastArm = {};
    if (blocks.length) {
      let fi = 0;
      blocks.forEach((b, i) => { if (Object.keys(b).length > Object.keys(blocks[fi]).length) fi = i; });
      fullArm = blocks[fi] || {};
      blocks.forEach((b, i) => {
        if (i === fi) return;
        const first = b[0];
        if (first !== undefined && Math.abs(first - 13.0) <= 0.15) { if (!Object.keys(halfArm).length) halfArm = b; }
        else if (!Object.keys(fastArm).length) fastArm = b;
      });
      if (!Object.keys(halfArm).length) halfArm = fullArm; // 只有一块时：快伸沿用全伸
    }
    return { fullArm, halfArm, fastArm, scope, full, half, fast };
  }
  const pous = {};
  // .export：TextBlob 里找含赋值的那块，归属取最近的上一个 Name 元素
  const names = [];
  const nameRe = /<Single[^>]*Name="Name"[^>]*>([^<]*)<\/Single>/g;
  let m;
  while ((m = nameRe.exec(text))) names.push({ pos: m.index, name: m[1] });
  if (names.length) {
    const blobRe = /<Single[^>]*Name="TextBlobForSerialisation"[^>]*>([\s\S]*?)<\/Single>/g;
    while ((m = blobRe.exec(text))) {
      const body = m[1];
      if (!hasFull.test(body)) continue;
      let nm = null;
      for (let i = names.length - 1; i >= 0; i--) { if (names[i].pos < m.index) { nm = names[i].name; break; } }
      if (!nm || !/^Working_Table_/.test(nm)) continue;
      pous[nm] = parsePouBody(body);
    }
  }
  // PLCopenXML：<pou name="Working_Table_...">…</pou>
  if (!Object.keys(pous).length) {
    const parts = text.split(/<pou name="(Working_Table_[A-Za-z0-9]+)"/);
    for (let i = 1; i < parts.length - 1; i += 2) {
      const body = parts[i + 1].split('</pou>')[0];
      if (hasFull.test(body)) pous[parts[i]] = parsePouBody(body);
    }
  }
  if (!Object.keys(pous).length) throw new Error('未识别的程序文件格式：没找到 Working_Table_ 工况表（请上传 CoDeSys 导出的 .export 或 PLCopenXML .xml，并确认第 3 节的数组名与程序里一致）');
  return pous;
}
/* 程序 POU 与 Excel 块的对应：精确匹配（配置前缀 + 工况键）优先；后备按工况键后缀匹配，
   臂长表最贴合的胜出（五臂/六臂同名工况靠臂长区分，正确匹配 9/9、错表约 4/9）。 */
/* 在 {值:下标} 索引里找最接近的下标（容差 0.15）：Excel 臂长 12.9 与程序臂长 13.0 差 0.1m 时仍能对应上；幅度点间隔 0.5m，不会误配 */
function nearestIdx(idx, v, tol) {
  tol = (tol === undefined) ? 0.15 : tol;
  if (idx[v] !== undefined) return idx[v];
  let best = undefined, bestD = Infinity;
  Object.keys(idx).forEach(k => {
    const d = Math.abs(parseFloat(k) - v);
    if (d < bestD) { bestD = d; best = idx[k]; }
  });
  return (bestD <= tol) ? best : undefined;
}
/* 程序里实际有的配重（从 POU 名里取 F/H 后面 C 的数字，如 JTC101FC21→21；不能直接搜 C\d+，会误取 JTC100 里的 100） */
function progCounterweights(pous) {
  const s = new Set();
  Object.keys(pous).forEach(n => {
    const m = n.slice('Working_Table_'.length).match(/[FH]C(\d+)/);
    if (m) s.add(parseInt(m[1], 10));
  });
  return s;
}
function rebuildProgramMap() {
  if (!state.program || !state.data) return;
  const map = {};
  const missing = [];
  const pous = state.program.pous;
  const cwSet = progCounterweights(pous);
  Object.keys(state.data.sheets).forEach(sh => {
    Object.keys(state.data.sheets[sh]).forEach(key => {
      const blk = state.data.sheets[sh][key];
      const disp = displayPrg(sh, blk);
      let found = pous['Working_Table_' + disp] ? 'Working_Table_' + disp : null;
      if (!found) {
        // 兜底：单复本文件只有左表配重（20/13），程序用右表配重（21/14）——取程序里最接近的配重再试
        let nearCw = null, nearD = Infinity;
        cwSet.forEach(v => { const d = Math.abs(v - blk.meta.cw); if (d < nearD) { nearD = d; nearCw = v; } });
        const nearKey = (nearCw !== null && nearCw !== Math.round(blk.meta.cw))
          ? blk.meta.letter + 'C' + String(nearCw).padStart(2, '0') + (blk.meta.back > 0 ? 'HY' : '')
          : null;
        const cands = Object.keys(pous).filter(n => {
          const rest = n.slice('Working_Table_'.length);
          return rest === key || rest.endsWith(key) || (nearKey && (rest === nearKey || rest.endsWith(nearKey)));
        });
        let best = null, bestScore = -1;
        const booms = blk.orig.Full.booms.map(b => cleanBoom(b));
        const need = Math.max(3, Math.ceil(booms.length * 0.6));
        cands.forEach(n => {
          const arms = Object.values(pous[n].fullArm || {});
          let score = 0;
          booms.forEach(b => { if (arms.some(a => Math.abs(cleanBoom(a) - b) <= 0.15)) score++; });
          if (score > bestScore) { bestScore = score; best = n; }
        });
        if (best && bestScore >= need) found = best;
      }
      map[sh + '|' + key] = found || '';
      if (!found) missing.push(sh + '|' + key);
    });
  });
  state.program.map = map;
  state.program.missingPrg = missing;
  const used = new Set(Object.values(map).filter(Boolean));
  state.program.unmatched = Object.keys(pous).filter(n => !used.has(n)).map(n => n.slice('Working_Table_'.length));
  state.program._gridCache = {};
}
function programPouFor(sh, key) {
  if (!state.program || !state.data) return null;
  const hit = (state.program.map || {})[sh + '|' + key];
  return hit ? hit : null;
}
/* 某（表，工况，组）的程序网格：臂长/幅度→行列号 + 值表（带缓存） */
function programGridFor(sh, key, grp) {
  if (!state.program) return null;
  const cacheKey = sh + '|' + key + '|' + grp;
  const cache = state.program._gridCache || (state.program._gridCache = {});
  if (cache[cacheKey] !== undefined) return cache[cacheKey];
  const pouName = programPouFor(sh, key);
  if (!pouName) { cache[cacheKey] = null; return null; }
  const P = state.program.pous[pouName];
  if (grp === 'Fast') {
    // 程序里没有 Fast_Weight 代码，或只有部分行（不全）：不参与对比，用户按提示生成代码块补进程序。
    // “全”的标准：Fast_Weight 的行数 ≥ Fast 臂长数（如 HC38 应有 6 行，只贴了第 0 行就不算全）。
    if (!P.fast || !Object.keys(P.fast).length) { cache[cacheKey] = null; return null; }
    const fastRows = new Set(Object.keys(P.fast).map(k => k.split(',')[0]));
    const fastArmCount = Object.keys(P.fastArm || {}).length;
    if (fastArmCount && fastRows.size < fastArmCount) { cache[cacheKey] = null; return null; }
  }
  const arm = grp === 'Full' ? P.fullArm : grp === 'Half' ? P.halfArm : P.fastArm;
  const armIdx = {};
  Object.keys(arm).forEach(i => { armIdx[cleanBoom(arm[i])] = parseInt(i, 10); });
  const scopeIdx = {};
  Object.keys(P.scope).forEach(i => { scopeIdx[cleanBoom(P.scope[i])] = parseInt(i, 10); });
  const grid = { armIdx, scopeIdx, arr: grp === 'Full' ? P.full : grp === 'Half' ? P.half : P.fast };
  cache[cacheKey] = grid;
  return grid;
}
/* 当前选中格的程序值（无对应返回 null）；cellIssueLevel 用 */
function programValue(grp, boom, rad) {
  if (!state.program || !state.data || state.src !== 'orig') return null;
  if (grp === 'Half' && String(boom) === halfZero()) return null; // 快伸 13m 行没有对应的 Excel 列（13m 列属于全伸组），不参与程序对比
  const grid = programGridFor(state.sheet, state.prg, grp);
  if (!grid) return null;
  const r = nearestIdx(grid.armIdx, cleanBoom(parseFloat(boom)));
  const c = nearestIdx(grid.scopeIdx, cleanBoom(rad));
  if (r === undefined || c === undefined) return null;
  const v = grid.arr[r + ',' + c];
  return (typeof v === 'number') ? v : 0;
}
/* 全部程序对比问题（只比原表）：每格 Excel 值（cellVal，含 2.5m 附加列）vs 程序值 */
function programDiffIssues() {
  const out = [];
  if (!state.program || !state.data) return out;
  const saved = { sheet: state.sheet, prg: state.prg, grp: state.grp, src: state.src };
  state.src = 'orig';
  const noCounter = {};
  Object.keys(state.data.sheets).forEach(sh => {
    Object.keys(state.data.sheets[sh]).forEach(key => {
      state.sheet = sh; state.prg = key;
      const blk = state.data.sheets[sh][key];
      const tag = sh + ' · ' + key + '（R' + blk.titleRow + '）';
      ['Full', 'Half', 'Fast'].forEach(grp => {
        state.grp = grp;
        const grid = programGridFor(sh, key, grp);
        if (!grid) return;
        const booms = blk.orig[grp].booms.map(String); // Half 的 booms 不含 halfZero（快伸 13m 行不参与程序对比，13m 列属于全伸组，已在 Full 里比过）
        const radii = prgScope();
        booms.forEach(bm => {
          const r = nearestIdx(grid.armIdx, cleanBoom(parseFloat(bm)));
          const bShow = cleanBoom(parseFloat(bm));
          radii.forEach(rad => {
            const c = nearestIdx(grid.scopeIdx, cleanBoom(rad));
            const ev = cellVal(grp, bm, rad);
            if (r === undefined || c === undefined) {
              if (typeof ev === 'number' && ev !== 0) noCounter[sh + '|' + key] = (noCounter[sh + '|' + key] || 0) + 1;
              return;
            }
            const pv = grid.arr[r + ',' + c];
            const p = (typeof pv === 'number') ? pv : 0;
            if (typeof ev === 'string') {
              out.push({ level: 'err', kind: 'prog', sheet: sh, prg: key, cell: excelCellFor(grp, bm, rad), text: tag + ' 臂长 ' + bShow + 'm、幅度 ' + rad + 'm：Excel 为文本 ' + JSON.stringify(ev) + '，程序为 ' + p + '，请核对纸质表' });
            } else if (p !== ev) {
              out.push({ level: 'err', kind: 'prog', sheet: sh, prg: key, cell: excelCellFor(grp, bm, rad), text: tag + ' 臂长 ' + bShow + 'm、幅度 ' + rad + 'm：程序 ' + p + '，Excel ' + ev + '，请核对纸质表' });
            }
          });
        });
      });
    });
  });
  Object.keys(noCounter).forEach(k => {
    const [sh, key] = k.split('|');
    out.push({ level: 'info', kind: 'prog', sheet: sh, prg: key, cell: '', text: sh + ' · ' + key + '（R' + state.data.sheets[sh][key].titleRow + '）有 ' + noCounter[k] + ' 格在程序里没有对应臂长/幅度，未参与对比' });
  });
  (state.program.unmatched || []).forEach(n => out.push({ level: 'info', kind: 'prog', sheet: '', prg: n, cell: '', text: '程序里有工况表 ' + n + '，Excel 里没有对应的块，未参与对比' }));
  (state.program.missingPrg || []).forEach(k => {
    const [sh, key] = k.split('|');
    out.push({ level: 'info', kind: 'prog', sheet: sh, prg: key, cell: '', text: sh + ' · ' + key + ' 在 Excel 里有，程序里没有对应的工况表，未参与对比' });
  });
  state.sheet = saved.sheet; state.prg = saved.prg; state.grp = saved.grp; state.src = saved.src;
  return out;
}
/* 趋势问题 + 程序对比问题合并：差异问题已含程序值，取代同格的“一格内有文本 / 空白空洞”趋势问题 */
function mergeProgramDiffs(trend) {
  const out = trend.slice();
  if (!state.program || !state.data) return out;
  const diffs = programDiffIssues();
  if (!diffs.length) return out;
  const keys = new Set(diffs.map(d => {
    const m = d.text.match(/臂长 ([\d.]+)m、幅度 ([\d.]+)m/);
    return m ? d.sheet + '|' + d.prg + '|' + m[1] + '|' + m[2] : '';
  }));
  const kept = out.filter(i => {
    if (i.text.indexOf('一格内有文本') < 0 && i.text.indexOf('空白空洞') < 0) return true;
    const m = i.text.match(/臂长 ([\d.]+)m、幅度 ([\d.]+)m/);
    if (!m) return true;
    return !keys.has(i.sheet + '|' + i.prg + '|' + m[1] + '|' + m[2]);
  });
  return diffs.concat(kept); // 差异优先级最高，放最前面；趋势作为提示，跟在后面
}
function saveProg() {
  try {
    if (!state.program) { localStorage.removeItem(PROG_KEY); return; }
    const p = state.program;
    localStorage.setItem(PROG_KEY, JSON.stringify({ name: p.name, time: p.time, size: p.size, format: p.format, pous: p.pous }));
  } catch (e) { /* 存不下不影响使用 */ }
}
function useProgram(text, fileName, fileSize) {
  const pous = parseProgram(text);
  const nPou = Object.keys(pous).length;
  const fmt = text.indexOf('TextBlobForSerialisation') >= 0 ? 'CoDeSys 导出（.export）' : 'PLCopenXML（.xml）';
  state.program = { name: fileName, time: Date.now(), size: fileSize || 0, format: fmt, pous, map: {}, missingPrg: [], unmatched: [] };
  rebuildProgramMap();
  issues = mergeProgramDiffs(issuesTrend);
  issuesOrig = JSON.parse(JSON.stringify(issues));
  saveProg();
  renderProgInfo();
  renderValidation(); renderAll();
  computeCircleTables(); renderCircleJump();
  jumpToFirstErrorTable();
}
function handleProgFile(file) {
  const rd = new FileReader();
  rd.onload = e => { try { useProgram(e.target.result, file.name, file.size); } catch (err) { $('progInfo').innerHTML = '<b>读取失败：</b>' + esc(err.message); } };
  rd.onerror = () => { $('progInfo').innerHTML = '<b>读取失败：</b>无法读取该文件。'; };
  rd.readAsText(file);
}
function clearProgram() {
  state.program = null;
  try { localStorage.removeItem(PROG_KEY); } catch (e) {}
  issues = issuesTrend.slice();
  issuesOrig = JSON.parse(JSON.stringify(issues));
  renderProgInfo();
  renderValidation(); renderAll();
  computeCircleTables(); renderCircleJump();
}
function renderProgInfo() {
  const el = $('progInfo'); if (!el) return;
  const p = state.program;
  if (!p) { el.innerHTML = '尚未上传程序文件。上传后逐格对比 Excel 里的值与程序里写的值，差异标红（圈图红圈 + 校验列表红行）。'; return; }
  const nPou = Object.keys(p.pous).length;
  if (!state.data) {
    el.innerHTML = '已载入程序 <b>' + esc(p.name) + '</b>（' + new Date(p.time).toLocaleString() + '，' + esc(p.format) + '，' + nPou + ' 个工况表）。尚未上传 Excel，传了 Excel 才开始对比。';
    return;
  }
  const diffs = issues.filter(i => i.kind === 'prog' && i.level === 'err');
  const byPrg = {};
  diffs.forEach(i => { const k = i.sheet + '|' + i.prg; byPrg[k] = (byPrg[k] || 0) + 1; });
  const parts = Object.entries(byPrg).sort((a, b) => b[1] - a[1]).map(([k, n]) => {
    const [sh, key] = k.split('|');
    const blk = state.data.sheets[sh] && state.data.sheets[sh][key];
    return (blk ? sh.replace('模式', '') + ' ' + displayPrg(sh, blk) : sh + ' ' + key) + ' ' + n;
  });
  let html = '已载入程序 <b>' + esc(p.name) + '</b>（' + esc(p.format) + '，' + nPou + ' 个工况表）：与 Excel 共 <b>' + diffs.length + '</b> 处不同，已标红。';
  // Fast 状态分三档：没有 / 有但不全（只贴了部分行）/ 全。只“全”的表才参与对比，其余只生成、不对比。
  const fastComplete = P => {
    if (!P.fast || !Object.keys(P.fast).length) return false;
    const rows = new Set(Object.keys(P.fast).map(k => k.split(',')[0]));
    const n = Object.keys(P.fastArm || {}).length;
    return !n || rows.size >= n;
  };
  const nFastAny = Object.values(p.pous).filter(P => P.fast && Object.keys(P.fast).length).length;
  const nFastFull = Object.values(p.pous).filter(fastComplete).length;
  if (nFastAny === 0) html += '<br>⚠️ 程序里<b>没有 Fast_Weight 代码</b>：Fast 组只生成、不对比。请在第 3 节复制整表 ST（已含 Fast_Weight 代码块）补进程序，补完重新上传程序文件即可对比。';
  else if (nFastFull < nFastAny) html += '<br>⚠️ 程序里有 Fast_Weight 代码（' + nFastAny + ' 个表），但只有 ' + nFastFull + ' 个表的行数贴全了：不全的表 Fast 组只生成、不对比，贴全后重新上传程序文件即可对比。';
  else html += '<br>✅ 程序里有 Fast_Weight 代码（' + nFastFull + ' 个表，均已贴全），Fast 组已参与对比。';
  // 不全的表：逐个列出缺哪些臂长（按程序 Fast 臂长表核对），让用户知道具体要补哪几行
  const incomplete = Object.entries(p.pous).filter(([nm, P]) => {
    if (!P.fast || !Object.keys(P.fast).length) return false;
    const rows = new Set(Object.keys(P.fast).map(k => k.split(',')[0]));
    const n = Object.keys(P.fastArm || {}).length;
    return n && rows.size < n;
  }).map(([nm, P]) => {
    const rows = new Set(Object.keys(P.fast).map(k => k.split(',')[0]));
    const missing = Object.keys(P.fastArm).filter(i => !rows.has(String(i))).map(i => cleanBoom(P.fastArm[i]));
    return nm.slice('Working_Table_'.length) + ' 缺 ' + missing.join('m、') + 'm';
  });
  if (incomplete.length) html += '<br>缺臂长的表：' + esc(incomplete.join('；')) + '。请在第 3 节复制对应表的 Fast_Weight 代码块补进程序。';
  if (parts.length) html += '<br>分表：' + esc(parts.join(' · ')) + '。';
  if ((p.missingPrg || []).length) html += '<br>Excel 有、程序没有的表：' + esc(p.missingPrg.map(k => { const [sh, key] = k.split('|'); return sh.replace('模式', '') + ' ' + key; }).join('、')) + '。';
  if ((p.unmatched || []).length) html += '<br>程序有、Excel 没有的表：' + esc(p.unmatched.join('、')) + '。';
  el.innerHTML = html;
}

/* ---------------- 事件 ---------------- */
$('srcBtns').addEventListener('click', e => {
  const s = e.target.dataset.src; if (!s) return;
  state.src = s; state.boom = null;
  document.querySelectorAll('#srcBtns button').forEach(b => b.classList.toggle('on', b.dataset.src === s));
  renderAll();
});
$('grpBtns').addEventListener('click', e => {
  const g = e.target.dataset.grp; if (!g) return;
  state.grp = g; state.boom = null;
  document.querySelectorAll('#grpBtns button').forEach(b => b.classList.toggle('on', b.dataset.grp === g));
  renderAll();
});
function copyText(txt, btn) {
  const label = btn.textContent;
  const done = () => { btn.textContent = '✓ 已复制'; btn.classList.add('done'); setTimeout(() => { btn.classList.remove('done'); btn.textContent = label; }, 1500); };
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(txt).then(done).catch(() => fallback());
  else fallback();
  function fallback() { const ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); done(); } catch (e) { } document.body.removeChild(ta); }
}
$('copyOne').onclick = e => { if (!state.data) { alert('请先上传 Excel'); return; } if (state.boom === null) { alert('请先选一个臂长'); return; } copyText($('out').textContent, e.target); };
$('copyAll').onclick = e => { if (!state.data) { alert('请先上传 Excel'); return; } copyText(fullPRGText(), e.target); };
$('lkGo').onclick = () => {
  if (!state.data) { $('lkOut').innerHTML = '<div class="err">请先上传 Excel。</div>'; return; }
  const grp = $('lkGrp').value;
  const boom = parseFloat($('lkBoom').value), rad = parseFloat($('lkRad').value);
  const load = parseFloat($('lkLoad').value);
  if (isNaN(boom) || isNaN(rad)) { $('lkOut').innerHTML = '<div class="err">请填写臂长和幅度。</div>'; return; }
  const saved = state.grp; state.grp = grp;
  const r = lookupValue(grp, boom, rad);
  state.grp = saved;
  if (r.error) { $('lkOut').innerHTML = '<div class="err">' + esc(r.error) + '</div>'; return; }
  let html = '<div class="info">当前工况 <b>' + state.sheet + ' · ' + displayPrg(state.sheet, block()) + '</b>（' + ({ Full: '全伸组', Half: '快伸组', Fast: 'Fast组' })[grp] + '）｜臂长 ' + boom + 'm、幅度 ' + rad + 'm<br>' +
    '插值点：臂长 ' + r.corners.b0 + '～' + r.corners.b1 + 'm，幅度 ' + r.corners.r0 + '～' + r.corners.r1 + 'm；四角值 ' + r.corners.v00 + ' / ' + r.corners.v01 + ' / ' + r.corners.v10 + ' / ' + r.corners.v11 + ' kg<br>' +
    '额定载荷：<span class="big">' + r.rated.toLocaleString() + ' kg</span>';
  if (r.notes.length) html += '<br>⚠️ ' + esc(r.notes.join('；'));
  if (!isNaN(load) && r.rated > 0) {
    const rate = load / r.rated * 100;
    const st = rate > 100 ? '🔴 超载' : rate >= 90 ? '🟡 预警（≥90%）' : '🟢 正常';
    html += '<br>实际吊重 ' + load.toLocaleString() + ' kg，载荷率 <span class="big">' + rate.toFixed(1) + '%</span> ' + st;
  } else if (!isNaN(load)) html += '<br>额定载荷为 0（超出作业范围），不能起吊。';
  $('lkOut').innerHTML = html + '</div>';
};
$('checkTable').addEventListener('change', e => {
  const inp = e.target;
  if (!inp.classList || !inp.classList.contains('celledit')) return;
  const raw = inp.value.trim();
  const num = Number(raw);
  const newVal = (raw !== '' && !isNaN(num)) ? num : raw;
  applyEdit(inp.dataset.grp, inp.dataset.boom, parseFloat(inp.dataset.rad), newVal);
});
$('segBtns').addEventListener('click', e => {
  const btn = e.target.closest ? e.target.closest('button') : e.target;
  if (!btn || !btn.dataset.boom) return;
  copyText(stLinesFor(btn.dataset.grp, btn.dataset.boom), btn);
});
$('dlImg').onclick = () => {
  const cv = $('errCanvas');
  if (!cv || !cv.toBlob) { alert('当前浏览器不支持导出图片'); return; }
  cv.toBlob(bl => {
    if (!bl) return;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(bl);
    a.download = state.sheet + '_' + displayPrg(state.sheet, block()) + '_' + state.grp + '_圈图.png';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  });
};
$('circleJump').addEventListener('click', e => {
  const btn = e.target.closest ? e.target.closest('button') : e.target;
  if (!btn || btn.dataset.i === undefined) return;
  const t = (state.circleTables || [])[parseInt(btn.dataset.i, 10)];
  if (!t) return;
  state.sheet = t.sheet; state.prg = t.prg; state.grp = t.grp; state.boom = null;
  document.querySelectorAll('#grpBtns button').forEach(b => b.classList.toggle('on', b.dataset.grp === t.grp));
  renderAll();
  const cv = $('errCanvas');
  if (cv && cv.scrollIntoView) cv.scrollIntoView({ block: 'nearest' });
});
$('prevErr').onclick = () => stepErrorTable(-1);
$('nextErr').onclick = () => stepErrorTable(1);
$('resetData').onclick = () => {
  try { localStorage.removeItem(LS_KEY); } catch (e) {}
  location.reload();
};
/* 配置输入：改完保存并立即重渲染（工况名、代码、圈图都跟着变） */
/* 配置改了（前缀/数组名/后缀/上限）→ 程序 POU 的对应关系可能变，重算对比 */
function refreshAfterCfg() {
  if (!state.data) return;
  renderAll();
  if (state.program) {
    rebuildProgramMap();
    issues = mergeProgramDiffs(issuesTrend);
    issuesOrig = JSON.parse(JSON.stringify(issues));
    renderProgInfo();
    renderValidation();
    computeCircleTables(); renderCircleJump();
  }
}
[['cfgArrFull', 'arrFull'], ['cfgArrHalf', 'arrHalf'], ['cfgArrFast', 'arrFast'], ['cfgArrArm', 'arrArm'], ['cfgArrScope', 'arrScope'], ['cfgGuard', 'guard'], ['cfgHy', 'hy'], ['cfgCap', 'cap']].forEach(([id, key]) => {
  const el = $(id);
  if (!el) return;
  el.addEventListener('input', () => {
    state.cfg[key] = key === 'cap' ? (parseInt(el.value, 10) || 0) : el.value;
    saveCfg();
    refreshAfterCfg();
  });
});
$('prefixBox').addEventListener('input', e => {
  const inp = e.target;
  if (!inp.dataset || inp.dataset.sheet === undefined) return;
  state.cfg.prefix[inp.dataset.sheet] = inp.value;
  saveCfg();
  refreshAfterCfg();
});
$('resetCfg').onclick = () => {
  state.cfg = Object.assign({}, DEFAULT_CFG, { prefix: {} });
  saveCfg(); syncCfgInputs(); renderPrefixBox();
  refreshAfterCfg();
};
const drop = $('drop');
drop.addEventListener('click', () => $('fileInput').click());
$('fileInput').addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });

/* 程序文件上传（第 1 节第二个框） */
const progDrop = $('progDrop');
progDrop.addEventListener('click', () => $('progInput').click());
$('progInput').addEventListener('change', e => { if (e.target.files[0]) handleProgFile(e.target.files[0]); });
['dragover', 'dragenter'].forEach(ev => progDrop.addEventListener(ev, e => { e.preventDefault(); progDrop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => progDrop.addEventListener(ev, e => { e.preventDefault(); progDrop.classList.remove('over'); }));
progDrop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleProgFile(f); });
$('clearProg').onclick = () => clearProgram();

/* ---------------- 启动：先恢复上次上传的 Excel；没有就是空状态，等用户上传 ---------------- */
(function init() {
  syncCfgInputs();
  try {
    const saved = typeof localStorage !== 'undefined' && localStorage.getItem(LS_KEY);
    if (saved) {
      const obj = JSON.parse(saved);
      if (obj && obj.data && obj.data.sheets) {
        state.data = obj.data;
        dataOrig = JSON.parse(JSON.stringify(obj.data));
        state.fileMeta = { name: obj.name, time: obj.time };
        state.dataLabel = '已上传：' + obj.name + '（已恢复）';
        $('dataBadge').textContent = '已恢复上传';
        issues = obj.issues || [];
        // 趋势问题单独存一份（合并后的列表里，被程序差异取代的 3 个趋势问题已删掉，清除程序时要靠这份恢复）
        issuesTrend = obj.issuesTrend || issues.filter(i => i.kind !== 'prog');
        issuesTrendOrig = JSON.parse(JSON.stringify(issuesTrend));
        // 恢复上次上传的程序（若有），并按当前程序重算对比；程序没恢复就去掉已存的对比问题
        let progRestored = false;
        try {
          const psaved = typeof localStorage !== 'undefined' && localStorage.getItem(PROG_KEY);
          if (psaved) {
            const pobj = JSON.parse(psaved);
            if (pobj && pobj.pous && Object.keys(pobj.pous).length) { state.program = pobj; progRestored = true; }
          }
        } catch (e) { /* 程序恢复失败不影响 Excel */ }
        if (progRestored) { rebuildProgramMap(); issues = mergeProgramDiffs(issuesTrend); }
        else issues = issuesTrend.slice();
        issuesOrig = JSON.parse(JSON.stringify(issues));
        $('fileInfo').innerHTML = '已恢复上次上传的 <b>' + esc(obj.name) + '</b>（' + new Date(obj.time).toLocaleString() + '）。改完 Excel 重新上传即可替换；点“清除上次上传”回到空状态。';
        computeCircleTables(); renderCircleJump();
        jumpToFirstErrorTable();
        if (!state.data.sheets[state.sheet]) state.sheet = Object.keys(state.data.sheets)[0];
        if (!state.data.sheets[state.sheet][state.prg]) state.prg = Object.keys(state.data.sheets[state.sheet])[0];
        renderPrefixBox();
        renderProgInfo();
        renderValidation(); renderAll();
        return;
      }
    }
  } catch (e) { /* 恢复失败则进空状态 */ }
  // 没恢复到上传：空状态，不显示任何校验结果
  issues = [];
  issuesTrend = [];
  issuesTrendOrig = [];
  issuesOrig = [];
  dataOrig = null;
  state.circleTables = [];
  renderProgInfo();
  renderValidation(); renderAll();
})();
