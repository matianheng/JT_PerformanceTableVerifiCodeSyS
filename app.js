'use strict';
/* JTC100G6 性能表动态校验 + ST 代码生成
   数据模型：{ scope:[35个幅度点], sheets:{ 表名:{ PRG:{label,titleRow,orig:{Full:{booms,rows},Half,Third},filled:{...}} } } }
   rows: { '13': {'3.0':100000,...} } 缺失填 0，文本格原样保留字符串 */
const SCOPE = [3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 24.0, 26.0, 28.0, 30.0, 32.0, 34.0, 36.0, 38.0, 40.0, 42.0, 44.0, 46.0, 48.0, 50.0, 52.0];
const PRG_ORDER = ['FC38HY', 'FC38', 'HC38HY', 'HC38', 'FC27HY', 'FC27', 'HC27HY', 'HC27', 'FC21HY', 'FC21', 'HC21HY', 'HC21', 'FC14HY', 'FC14', 'HC14HY', 'HC14', 'FC00', 'HC00'];
const PRG_LABEL = { 'FC38HY': '全伸 38t 后移800', 'FC38': '全伸 38t 后移0', 'HC38HY': '半伸 38t 后移800', 'HC38': '半伸 38t 后移0', 'FC27HY': '全伸 27t 后移800', 'FC27': '全伸 27t 后移0', 'HC27HY': '半伸 27t 后移800', 'HC27': '半伸 27t 后移0', 'FC21HY': '全伸 20.5t 后移800', 'FC21': '全伸 20.5t 后移0', 'HC21HY': '半伸 20.5t 后移800', 'HC21': '半伸 20.5t 后移0', 'FC14HY': '全伸 14t 后移800', 'FC14': '全伸 14t 后移0', 'HC14HY': '半伸 14t 后移800', 'HC14': '半伸 14t 后移0', 'FC00': '全伸 0t', 'HC00': '半伸 0t' };

let state = { data: null, dataLabel: '内置数据', sheet: '六臂模式', prg: 'FC14', src: 'orig', grp: 'Full', boom: null };
let issues = [];   // {level:'err'|'warn'|'info', sheet, prg, cell, text}

/* 内置数据以紧凑数组形式分块存放在 data/data_*.js（window.__DP），此处展开为 app 用的键值模型 */
function expandCompact(parts) {
  if (!parts) return null;
  const data = { scope: SCOPE.slice(), sheets: {} };
  const sheetNames = { '5': '五臂模式', '6': '六臂模式' };
  ['5a', '5b', '6a', '6b'].forEach(tag => {
    const prgs = parts[tag]; if (!prgs) return;
    const sheet = sheetNames[tag[0]];
    const src = tag[1] === 'a' ? 'orig' : 'filled';
    data.sheets[sheet] = data.sheets[sheet] || {};
    Object.keys(prgs).forEach(prg => {
      const p = prgs[prg];
      data.sheets[sheet][prg] = data.sheets[sheet][prg] || { label: p.label, titleRow: p.titleRow, orig: null, filled: null };
      const groups = {};
      Object.keys(p[src]).forEach(g => {
        const gd = p[src][g];
        const rows = {};
        gd.booms.forEach((b, i) => {
          const vals = {};
          SCOPE.forEach((rad, j) => { vals[Number(rad).toFixed(1)] = gd.rows[i][j]; });
          rows[String(b)] = vals;
        });
        groups[g] = { booms: gd.booms, rows };
      });
      data.sheets[sheet][prg][src] = groups;
    });
  });
  return data;
}
state.data = (typeof window !== 'undefined' && window.__DP) ? expandCompact(window.__DP) : null;

/* ---------------- 工具 ---------------- */
const $ = id => document.getElementById(id);
const rk = r => Number(r).toFixed(1); // 幅度键：'3.0'、'10.0'（与 data.js 一致；JS 的 String(3.0) 是 '3'，不能直接用）
const fmtR = x => Number.isInteger(x) ? x.toFixed(1) : String(x); // 幅度显示：6 -> '6.0'
const fmtB = b => String(b).replace(/\.0$/, '');
const cleanBoom = b => (typeof b === 'number' ? Math.round(b * 10) / 10 : b);
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
function block() { return state.data.sheets[state.sheet][state.prg]; }
function grpData() { return block()[state.src][state.grp]; }
function colLetter(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; }

/* ---------------- Excel 解析 ----------------
   grid: sheet_to_json(header:1) 的二维数组，0 起下标。
   每块：幅度列含 '幅度(m)' 的行是表头行；标题行 = 表头行-1；臂长行 = 表头行+1（有序号行时 +2）。
   横向 6 个复本，步长 24 列：copy1 原表、copy2 右表、copy3 填充版。 */
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
  const booms = [];
  for (let i = 0; i < 21; i++) booms.push(grid[boomRow] ? grid[boomRow][ampCol + 1 + off + i] : null);
  const rows = []; // {radius, vals[21], sheetRow}
  let r = boomRow + 1;
  while (r < grid.length) {
    const a = grid[r] ? grid[r][ampCol + off] : null;
    if (typeof a === 'string' && a.trim() === '倍率') break;
    if (a === null || a === undefined || a === '') { r++; continue; }
    const vals = [];
    for (let i = 0; i < 21; i++) vals.push(grid[r] ? grid[r][ampCol + 1 + off + i] : null);
    rows.push({ radius: a, vals, sheetRow: r });
    r++;
  }
  return { title, titleRow, extra, boomRow, booms, rows };
}

function toGroups(copy) {
  const groups = {};
  [['Full', 0, 9], ['Half', 9, 15], ['Third', 15, 21]].forEach(([g, a, b]) => {
    const blist = copy.booms.slice(a, b).map(cleanBoom);
    const rows = {};
    copy.booms.slice(a, b).forEach((bRaw, k) => {
      const vals = {};
      SCOPE.forEach(rad => {
        const hit = copy.rows.find(rr => rr.radius === rad);
        let v = hit ? hit.vals[a + k] : 0;
        if (v === null || v === undefined) v = 0;
        vals[rk(rad)] = v;
      });
      rows[String(blist[k])] = vals;
    });
    groups[g] = { booms: blist, rows };
  });
  return groups;
}

/* 幅度重复修正：恰好一个重复值、一个缺失值且缺失 = 重复 - 0.5 时，把第一次出现的重复行改标为缺失值 */
function fixDupRadius(copy) {
  const radii = copy.rows.map(r => r.radius).filter(x => typeof x === 'number');
  const seen = {}, dups = [];
  radii.forEach(x => { if (seen[x]) dups.push(x); seen[x] = true; });
  if (dups.length !== 1) return null;
  const dup = dups[0];
  const maxR = Math.max(...radii);
  const expected = SCOPE.filter(x => x <= maxR);
  const missing = expected.filter(x => !seen[x]);
  if (missing.length === 1 && Math.abs(missing[0] - (dup - 0.5)) < 1e-9) {
    const first = copy.rows.find(r => r.radius === dup);
    if (first) first.radius = missing[0];
    return { dup, fixed: missing[0] };
  }
  return null;
}

function parseTitleForPrg(leftTitle, rightTitle) {
  const t = String(leftTitle || '') + ' ' + String(rightTitle || '');
  const spanM = String(leftTitle || '').match(/跨距([\d.]+)m/);
  const span = spanM ? parseFloat(spanM[1]) : null;
  const out = span === 8.4 ? 'F' : span === 5.5 ? 'H' : null;
  const cwM = String(rightTitle || leftTitle || '').match(/(\d+(?:\.\d+)?)吨配重/);
  const cw = cwM ? parseFloat(cwM[1]) : null;
  const num = cw === 38 ? '38' : cw === 27 ? '27' : cw === 20.5 ? '21' : cw === 14 ? '14' : cw === 0 ? '00' : null;
  const backM = t.match(/后移(\d+)mm/);
  const back = backM ? parseInt(backM[1], 10) : null;
  const hy = back === 800 ? 'HY' : '';
  if (!out || !num) return null;
  return { prg: out + 'C' + num + hy, span, cw, back };
}

function parseWorkbook(wb) {
  const data = { scope: SCOPE.slice(), sheets: {} };
  const found = [];
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
    const sheetData = {};
    heads.forEach(hr => {
      const c1 = parseBlock(grid, ampCol, hr, 0);
      const c2 = parseBlock(grid, ampCol, hr, 24, c1.extra);
      const c3 = parseBlock(grid, ampCol, hr, 48, c1.extra);
      const m = parseTitleForPrg(c1.title, c2.title);
      if (!m) return;
      const fix = fixDupRadius(c1);
      sheetData[m.prg] = {
        label: (m.span === 8.4 ? '全伸 ' : '半伸 ') + m.cw + 't 后移' + (m.back === null ? '—' : m.back),
        titleRow: c1.titleRow + 1,
        orig: toGroups(c1), filled: toGroups(c3),
      };
      found.push({ sheet: name, prg: m.prg, c1, c2, c3, fix, titleRow: c1.titleRow + 1, span: m.span, cw: m.cw, back: m.back, ampCol });
    });
    if (Object.keys(sheetData).length) data.sheets[name] = sheetData;
  });
  return { data, found };
}

/* ---------------- 校验 ---------------- */
function validate(found) {
  const list = [];
  const add = (level, sheet, prg, cell, text) => list.push({ level, sheet, prg, cell, text });
  found.forEach(f => {
    const tag = f.sheet + ' · ' + f.prg + '（R' + f.titleRow + '）';
    const c1 = f.c1;
    // 标题文字与跨距不符
    const titleStr = String(c1.title || '');
    if (f.span === 5.5 && titleStr.indexOf('支腿全伸') === 0) add('err', f.sheet, f.prg, colLetter(f.ampCol + 2) + f.titleRow, tag + ' 半伸表标题误写“支腿全伸侧方跨距5.5m”，应为“支腿半伸”');
    if (f.span === 8.4 && titleStr.indexOf('支腿半伸') === 0) add('err', f.sheet, f.prg, '标题行', tag + ' 全伸表标题误写“支腿半伸”');
    // 配重分量和
    const cwM = titleStr.match(/（基本配重0吨\+([^）]+)/);
    const tonM = titleStr.match(/(\d+(?:\.\d+)?)吨配重/);
    if (cwM && tonM) {
      const sum = cwM[1].split('+').reduce((s, x) => s + (parseFloat(x) || 0), 0);
      if (Math.abs(sum - parseFloat(tonM[1])) > 0.01) add('err', f.sheet, f.prg, '标题行', tag + ' 配重标题 ' + tonM[1] + 't ≠ 分量和 ' + sum + 't');
    }
    // 幅度重复 / 缺失
    if (f.fix) add('warn', f.sheet, f.prg, 'B列', tag + ' 幅度 ' + fmtR(f.fix.dup) + 'm 出现两次、' + fmtR(f.fix.fixed) + 'm 缺失，已按趋势把第一行 ' + fmtR(f.fix.dup) + ' 当作 ' + fmtR(f.fix.fixed) + ' 处理');
    else {
      const radii = c1.rows.map(r => r.radius).filter(x => typeof x === 'number');
      const seen = {};
      radii.forEach(x => { if (seen[x]) add('err', f.sheet, f.prg, 'B列', tag + ' 幅度 ' + x + 'm 重复'); seen[x] = true; });
    }
    // 序号行
    if (c1.extra) add('warn', f.sheet, f.prg, '序号行', tag + ' 多出一行 0,1,2…序号行，全表仅个别块有，格式不一致');
    // 2.5m 行
    if (c1.rows.some(r => r.radius === 2.5)) add('info', f.sheet, f.prg, '', tag + ' 有 2.5m 幅度行，程序幅度点从 3.0 起，该行无处可填');
    // 逐列检查
    for (let ci = 0; ci < 21; ci++) {
      const boom = cleanBoom(c1.booms[ci]);
      let prev = null, prevR = null;
      const colVals = c1.rows.map(r => ({ rad: r.radius, v: r.vals[ci] }));
      const present = colVals.filter(x => x.v !== null && x.v !== undefined);
      const minI = present.length ? colVals.indexOf(present[0]) : -1;
      const maxI = present.length ? colVals.indexOf(present[present.length - 1]) : -1;
      colVals.forEach((x, i) => {
        const cell = colLetter(f.ampCol + 1 + ci) + (c1.boomRow + 2 + i);
        if (typeof x.v === 'string') {
          let extra = '';
          const r2 = f.c2.rows.find(r => r.radius === x.rad);
          if (r2 && r2.vals[ci] !== null && r2.vals[ci] !== undefined) extra = '，右表同格为 ' + JSON.stringify(r2.vals[ci]);
          add('err', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + x.rad + 'm 一格内有文本 ' + JSON.stringify(x.v) + '，不是数字' + extra + '，请核对纸质表');
          return;
        }
        if (x.v === null || x.v === undefined) {
          if (minI >= 0 && i > minI && i < maxI) add('err', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + x.rad + 'm 空白空洞');
          return;
        }
        if (x.v % 100 !== 0) add('warn', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + x.rad + 'm = ' + x.v + '，不是 100 的倍数（请核对纸质表）');
        if (prev !== null && x.v > prev) add('err', f.sheet, f.prg, cell, tag + ' 臂长 ' + boom + 'm、幅度 ' + prevR + '→' + x.rad + 'm 上升 ' + prev + '→' + x.v + '，违反递减规律');
        prev = x.v; prevR = x.rad;
      });
    }
    // 右表与原表比对（记录前 3 处明细）
    let diffN = 0; const diffEx = [];
    f.c2.rows.forEach(r2 => {
      const r1 = c1.rows.find(r => r.radius === r2.radius);
      if (!r1) return;
      r2.vals.forEach((v, i) => {
        const o = r1.vals[i];
        if (v !== null && v !== undefined && o !== null && o !== undefined && v !== o) {
          diffN++;
          if (diffEx.length < 3) diffEx.push('臂长 ' + cleanBoom(c1.booms[i]) + 'm、幅度 ' + r2.radius + 'm：原表 ' + JSON.stringify(o) + '、右表 ' + JSON.stringify(v));
        }
      });
    });
    if (diffN) add('err', f.sheet, f.prg, '', tag + ' 右表与原表重叠格有 ' + diffN + ' 处数值不同（' + diffEx.join('；') + '）');
    // 右表臂长网格
    const g1 = c1.booms.slice(0, 3).map(cleanBoom).join('/');
    const g2 = f.c2.booms.slice(0, 3).map(cleanBoom).join('/');
    if (g1 !== g2) add('warn', f.sheet, f.prg, '臂长行', tag + ' 右表臂长网格与左表不同：左 ' + g1 + '，右 ' + g2 + '（短约 0.1m）');
    // 填充版系数
    const fac = {};
    f.c3.rows.forEach(r3 => {
      const r1 = c1.rows.find(r => r.radius === r3.radius);
      if (!r1) return;
      r3.vals.forEach((v, i) => {
        const o = r1.vals[i];
        if (typeof v === 'number' && typeof o === 'number' && o > 0 && v > 0) {
          const k = Math.round(v / o * 1000) / 1000;
          fac[k] = (fac[k] || 0) + 1;
        }
      });
    });
    const top = Object.entries(fac).sort((a, b) => b[1] - a[1]).slice(0, 3);
    if (top.length && Math.abs(top[0][0] - 1.33) > 0.001) add('warn', f.sheet, f.prg, '', tag + ' 填充版/原表比值主要为 ' + top[0][0] + '（不是 1.33），共 ' + top[0][1] + ' 格');
  });
  return list;
}

/* ---------------- ST 生成 ---------------- */
function boomList() {
  const g = grpData();
  if (state.grp === 'Half') return ['13'].concat(g.booms.map(String));
  return g.booms.map(String);
}
function boomRow(boom) {
  const g = grpData();
  if (state.grp === 'Half' && String(boom) === '13') return block()[state.src]['Full']['rows']['13'];
  return g.rows[String(boom)];
}
function boomIndex(boom) {
  if (state.grp === 'Full') return grpData().booms.map(String).indexOf(String(boom));
  if (state.grp === 'Half') return String(boom) === '13' ? 0 : grpData().booms.map(String).indexOf(String(boom)) + 1;
  return grpData().booms.map(String).indexOf(String(boom));
}
function arrName() { return state.grp === 'Full' ? 'Full_Weight' : state.grp === 'Half' ? 'Half_Weight' : 'Third_Weight'; }
function stLinesFor(grp, boom) {
  const saved = state.grp; state.grp = grp;
  const idx = boomIndex(boom);
  const arr = arrName();
  const row = boomRow(boom);
  state.grp = saved;
  const lines = ['\t(*' + fmtB(parseFloat(boom)) + 'm*)'];
  SCOPE.forEach((rad, j) => {
    const v = row[rk(rad)];
    if (typeof v === 'string') { lines.push('\t' + arr + '[' + idx + ',' + j + ']:= 0 ;\t(* 原表此处为 ' + JSON.stringify(v) + '，待核对，暂填0 *)'); return; }
    lines.push('\t' + arr + '[' + idx + ',' + j + ']:= ' + v + ' ;');
  });
  return lines.join('\n');
}
function fullPRGText() {
  const b = block();
  const fullBooms = b[state.src]['Full']['booms'];
  const halfBooms = b[state.src]['Half']['booms'];
  const thirdBooms = b[state.src]['Third']['booms'];
  let t = '(*臂长点*)\nIF NOT TeleOut_Fast_Control_LMI THEN\n';
  fullBooms.forEach((bm, i) => { t += '\tArm_Length_Weight_Point[' + i + ']:= ' + fmtB(bm) + ';\n'; });
  t += '(*快伸工况*)\nELSE\n\tArm_Length_Weight_Point[0]:= 13.0;\n';
  halfBooms.forEach((bm, i) => { t += '\tArm_Length_Weight_Point[' + (i + 1) + ']:= ' + fmtB(bm) + ';\n'; });
  t += 'END_IF;\n\n(*幅度点*)\n';
  SCOPE.forEach((rad, i) => { t += '\tScope_Weight_Point[' + i + ']:= ' + rad.toFixed(2) + ';\n'; });
  t += '\n(*支腿' + (state.prg[0] === 'F' ? '全伸' : '半伸') + '工况*)\n';
  fullBooms.forEach(bm => { t += stLinesFor('Full', String(bm)) + '\n\n'; });
  SCOPE.forEach((rad, j) => { t += '\tFull_Weight[9,' + j + ']:= 0 ;\n'; });
  t += '\n(*快伸工况性能表*)\n';
  ['13'].concat(halfBooms.map(String)).forEach(bm => { t += stLinesFor('Half', bm) + '\n\n'; });
  const halfCount = halfBooms.length + 1;
  for (let r = halfCount; r <= 9; r++) { SCOPE.forEach((rad, j) => { t += '\tHalf_Weight[' + r + ',' + j + ']:= 0 ;\n'; }); t += '\n'; }
  t += '(*第三组性能表（数组名 Third_Weight 为占位，请按实际改）*)\n';
  thirdBooms.forEach(bm => { t += stLinesFor('Third', String(bm)) + '\n\n'; });
  t += 'RETURN;\n';
  return t;
}

/* ---------------- 动态查表（插值） ---------------- */
function lookupValue(grpName, boom, rad) {
  const g = block()[state.src][grpName];
  const booms = grpName === 'Half' ? [13].concat(g.booms.map(Number)) : g.booms.map(Number);
  const rows = {};
  booms.forEach(b => { rows[String(b)] = (grpName === 'Half' && b === 13) ? block()[state.src]['Full']['rows']['13'] : g.rows[String(b)]; });
  const clamp = (x, lo, hi, name) => {
    if (x < lo) return { v: lo, out: name + '小于表内最小值 ' + lo + '，按边界取值' };
    if (x > hi) return { v: hi, out: name + '大于表内最大值 ' + hi + '，按边界取值' };
    return { v: x, out: null };
  };
  const cb = clamp(boom, booms[0], booms[booms.length - 1], '臂长');
  const cr = clamp(rad, SCOPE[0], SCOPE[SCOPE.length - 1], '幅度');
  const notes = [cb.out, cr.out].filter(Boolean);
  const b = cb.v, r = cr.v;
  const b0 = [...booms].reverse().find(x => x <= b);
  const b1 = booms.find(x => x >= b);
  const r0 = [...SCOPE].reverse().find(x => x <= r);
  const r1 = SCOPE.find(x => x >= r);
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
  Object.keys(state.data.sheets).forEach(s => {
    const b = document.createElement('button');
    b.textContent = s; if (s === state.sheet) b.className = 'on';
    b.onclick = () => { state.sheet = s; if (!state.data.sheets[s][state.prg]) state.prg = Object.keys(state.data.sheets[s])[0]; state.boom = null; renderAll(); };
    el.appendChild(b);
  });
}
function renderPrgs() {
  const el = $('prgBtns'); el.innerHTML = '';
  const names = Object.keys(state.data.sheets[state.sheet] || {});
  const order = PRG_ORDER.filter(p => names.includes(p)).concat(names.filter(p => !PRG_ORDER.includes(p)));
  order.forEach(p => {
    const b = document.createElement('button');
    b.textContent = p + '（' + (PRG_LABEL[p] || state.data.sheets[state.sheet][p].label) + '）';
    if (p === state.prg) b.className = 'on';
    b.onclick = () => { state.prg = p; state.boom = null; renderAll(); };
    el.appendChild(b);
  });
}
function renderInfo() {
  const b = block();
  $('blockInfo').innerHTML = '当前：<b>' + state.sheet + ' · ' + state.prg + '</b>（' + (PRG_LABEL[state.prg] || b.label) + '）｜数据源：<b>' + (state.src === 'orig' ? '原表' : '力限填充版') + '</b>｜Excel 标题行 R' + b.titleRow + '｜数据：' + esc(state.dataLabel);
  const w = $('blockWarn');
  const warns = issues.filter(i => i.prg === state.prg && i.sheet === state.sheet && i.level !== 'info').map(i => (i.level === 'err' ? '🔴 ' : '🟡 ') + esc(i.text));
  if (warns.length) { w.style.display = 'block'; w.innerHTML = '⚠️ 本表已知问题：<br>· ' + warns.join('<br>· '); }
  else w.style.display = 'none';
}
function renderBooms() {
  const el = $('boomBtns'); el.innerHTML = '';
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
  if (state.boom === null) { out.textContent = '请先选臂长。'; tbl.innerHTML = ''; w.innerHTML = ''; return; }
  out.textContent = stLinesFor(state.grp, state.boom);
  const row = boomRow(state.boom);
  let html = '<tr><th>幅度(m)</th><th class="num">Excel 值</th><th class="num">进程序</th><th>备注</th></tr>';
  SCOPE.forEach(rad => {
    const v = row[rk(rad)];
    const iss = issues.find(i => i.sheet === state.sheet && i.prg === state.prg && i.text.indexOf('臂长 ' + cleanBoom(parseFloat(state.boom)) + 'm、幅度 ' + rad + 'm') >= 0);
    const cls = typeof v === 'string' ? 'bad' : iss ? (iss.level === 'err' ? 'bad' : 'warncell') : (v === 0 ? 'zero' : '');
    html += '<tr><td>' + rad + '</td><td class="num ' + cls + '">' + esc(typeof v === 'string' ? v : v) + '</td><td class="num">' + (typeof v === 'string' ? 0 : v) + '</td><td>' + (iss ? esc(iss.text.split('】').pop() || iss.text) : '') + '</td></tr>';
  });
  tbl.innerHTML = html;
  w.innerHTML = '';
}
function renderValidation() {
  const errs = issues.filter(i => i.level === 'err').length;
  const warns = issues.filter(i => i.level === 'warn').length;
  const infos = issues.filter(i => i.level === 'info').length;
  $('validSummary').innerHTML = '校验完成：🔴 确认错误 <b>' + errs + '</b> 处，🟡 待确认 <b>' + warns + '</b> 处，ℹ️ 提示 <b>' + infos + '</b> 处。红色先改，黄色核对纸质表。';
  if (!issues.length) { $('validList').innerHTML = '<div class="info">未发现问题。</div>'; return; }
  let html = '<table><tr><th>级别</th><th>表 / 工况</th><th>位置</th><th>问题</th></tr>';
  issues.forEach(i => {
    html += '<tr class="lvl-' + i.level + '"><td>' + (i.level === 'err' ? '🔴 错误' : i.level === 'warn' ? '🟡 待确认' : 'ℹ️ 提示') + '</td><td>' + esc(i.sheet + ' · ' + i.prg) + '</td><td>' + esc(i.cell || '—') + '</td><td>' + esc(i.text) + '</td></tr>';
  });
  $('validList').innerHTML = html + '</table>';
}
function renderAll() { renderSheets(); renderPrgs(); renderInfo(); renderBooms(); renderOut(); }

/* ---------------- 上传 ---------------- */
function useUploaded(wb, fileName) {
  try {
    const { data, found } = parseWorkbook(wb);
    const nPrg = Object.values(data.sheets).reduce((s, sh) => s + Object.keys(sh).length, 0);
    if (!nPrg) { $('fileInfo').innerHTML = '<b>解析失败：</b>在 ' + esc(fileName) + ' 里没找到含“幅度(m)”的性能表。请确认文件是 JTC100G6 性能表原文件。'; return; }
    state.data = data;
    state.dataLabel = '已上传：' + fileName;
    $('dataBadge').textContent = '已上传文件';
    $('fileInfo').innerHTML = '已载入 <b>' + esc(fileName) + '</b>：解析出 ' + Object.keys(data.sheets).length + ' 张表、' + nPrg + ' 个工况块。PRG 与块的对应按标题自动识别（跨距 8.4m=全伸、5.5m=半伸）。';
    issues = validate(found);
    if (!data.sheets[state.sheet]) state.sheet = Object.keys(data.sheets)[0];
    if (!data.sheets[state.sheet][state.prg]) state.prg = Object.keys(data.sheets[state.sheet])[0];
    state.boom = null;
    renderValidation(); renderAll();
  } catch (e) {
    $('fileInfo').innerHTML = '<b>解析出错：</b>' + esc(e.message);
  }
}
function handleFile(file) {
  if (typeof XLSX === 'undefined') { $('fileInfo').innerHTML = '<b>解析库未加载：</b>当前离线或 CDN 不可达，无法解析上传的 Excel。内置数据仍可用；联网后刷新页面再上传。'; return; }
  const rd = new FileReader();
  rd.onload = e => { try { useUploaded(XLSX.read(e.target.result, { type: 'array' }), file.name); } catch (err) { $('fileInfo').innerHTML = '<b>读取失败：</b>' + esc(err.message); } };
  rd.readAsArrayBuffer(file);
}

/* ---------------- 事件 ---------------- */
$('srcBtns').addEventListener('click', e => {
  const s = e.target.dataset.src; if (!s) return;
  state.src = s; state.boom = null;
  document.querySelectorAll('#srcBtns button').forEach(b => b.classList.toggle('on', b.dataset.src === s));
  renderBooms(); renderOut(); renderInfo();
});
$('grpBtns').addEventListener('click', e => {
  const g = e.target.dataset.grp; if (!g) return;
  state.grp = g; state.boom = null;
  document.querySelectorAll('#grpBtns button').forEach(b => b.classList.toggle('on', b.dataset.grp === g));
  renderBooms(); renderOut();
});
function copyText(txt, btn) {
  const label = btn.textContent;
  const done = () => { btn.textContent = '✓ 已复制'; btn.classList.add('done'); setTimeout(() => { btn.classList.remove('done'); btn.textContent = label; }, 1500); };
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(txt).then(done).catch(() => fallback());
  else fallback();
  function fallback() { const ta = document.createElement('textarea'); ta.value = txt; document.body.appendChild(ta); ta.select(); try { document.execCommand('copy'); done(); } catch (e) { } document.body.removeChild(ta); }
}
$('copyOne').onclick = e => { if (state.boom === null) { alert('请先选一个臂长'); return; } copyText($('out').textContent, e.target); };
$('copyAll').onclick = e => { copyText(fullPRGText(), e.target); };
$('lkGo').onclick = () => {
  const grp = $('lkGrp').value;
  const boom = parseFloat($('lkBoom').value), rad = parseFloat($('lkRad').value);
  const load = parseFloat($('lkLoad').value);
  if (isNaN(boom) || isNaN(rad)) { $('lkOut').innerHTML = '<div class="err">请填写臂长和幅度。</div>'; return; }
  const saved = state.grp; state.grp = grp;
  const savedSrc = state.src;
  const r = lookupValue(grp, boom, rad);
  state.grp = saved; state.src = savedSrc;
  if (r.error) { $('lkOut').innerHTML = '<div class="err">' + esc(r.error) + '</div>'; return; }
  let html = '<div class="info">当前工况 <b>' + state.sheet + ' · ' + state.prg + '</b>（' + (grp === 'Full' ? '全伸组' : grp === 'Half' ? '快伸组' : '第三组') + '）｜臂长 ' + boom + 'm、幅度 ' + rad + 'm<br>' +
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
const drop = $('drop');
drop.addEventListener('click', () => $('fileInput').click());
$('fileInput').addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });

/* ---------------- 启动：内置数据用预置校验结果，上传新文件后改为实时校验 ---------------- */
(function init() {
  if (!state.data) { $('validSummary').textContent = '内置数据未加载。'; return; }
  issues = (typeof BUILTIN_ISSUES !== 'undefined' ? BUILTIN_ISSUES : []);
  renderValidation(); renderAll();
})();
