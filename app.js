'use strict';
/* JTC100G6 性能表动态校验 + ST 代码生成
   数据模型：{ scope:[35个幅度点], sheets:{ 表名:{ PRG:{label,titleRow,orig:{Full:{booms,rows},Half,Third},filled:{...}} } } }
   rows: { '13': {'3.0':100000,...} } 缺失填 0，文本格原样保留字符串 */
const SCOPE = [3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0, 9.0, 10.0, 11.0, 12.0, 14.0, 16.0, 18.0, 20.0, 22.0, 24.0, 26.0, 28.0, 30.0, 32.0, 34.0, 36.0, 38.0, 40.0, 42.0, 44.0, 46.0, 48.0, 50.0, 52.0];
const PRG_ORDER = ['FC38HY', 'FC38', 'HC38HY', 'HC38', 'FC27HY', 'FC27', 'HC27HY', 'HC27', 'FC21HY', 'FC21', 'HC21HY', 'HC21', 'FC14HY', 'FC14', 'HC14HY', 'HC14', 'FC00', 'HC00'];
const PRG_LABEL = { 'FC38HY': '全伸 38t 后移800', 'FC38': '全伸 38t 后移0', 'HC38HY': '半伸 38t 后移800', 'HC38': '半伸 38t 后移0', 'FC27HY': '全伸 27t 后移800', 'FC27': '全伸 27t 后移0', 'HC27HY': '半伸 27t 后移800', 'HC27': '半伸 27t 后移0', 'FC21HY': '全伸 20.5t 后移800', 'FC21': '全伸 20.5t 后移0', 'HC21HY': '半伸 20.5t 后移800', 'HC21': '半伸 20.5t 后移0', 'FC14HY': '全伸 14t 后移800', 'FC14': '全伸 14t 后移0', 'HC14HY': '半伸 14t 后移800', 'HC14': '半伸 14t 后移0', 'FC00': '全伸 0t', 'HC00': '半伸 0t' };

let state = { data: null, dataLabel: '内置数据', sheet: '六臂模式', prg: 'FC14', src: 'orig', grp: 'Full', boom: null, fileMeta: null, edits: {} };
let issues = [];   // {level:'err'|'warn'|'info', sheet, prg, cell, text}
let issuesOrig = null;  // 载入/上传时的原始校验结果（撤销在线修改时恢复）
let dataOrig = null;    // 载入/上传时的数据深拷贝（撤销在线修改时恢复）
const LS_KEY = 'jt_perf_upload_v1';  // localStorage：上次上传的 Excel 解析结果

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
        const extra = {};
        if (gd.extra) Object.keys(gd.extra).forEach(rad => {
          extra[rad] = {};
          gd.booms.forEach((b, i) => { extra[rad][String(b)] = gd.extra[rad][i]; });
        });
        groups[g] = { booms: gd.booms, rows, extra };
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
    const extra = {};
    copy.rows.forEach(rr => {
      if (typeof rr.radius === 'number' && !SCOPE.includes(rr.radius)) {
        extra[rk(rr.radius)] = {};
        copy.booms.slice(a, b).forEach((bRaw, k) => {
          const v = rr.vals[a + k];
          extra[rk(rr.radius)][String(blist[k])] = (v === null || v === undefined) ? 0 : v;
        });
      }
    });
    groups[g] = { booms: blist, rows, extra };
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
  // 只认跨距数字区分：>=8m 为全伸 F，<=6m 为半伸 H（Excel 里是 8.4/5.5；8.8 也按全伸认）
  const cwM = String(rightTitle || leftTitle || '').match(/(\d+(?:\.\d+)?)吨配重/);
  const cw = cwM ? parseFloat(cwM[1]) : null;
  const num = cw === 38 ? '38' : cw === 27 ? '27' : cw === 20.5 ? '21' : cw === 14 ? '14' : cw === 0 ? '00' : null;
  const backM = t.match(/后移(\d+)mm/);
  const back = backM ? parseInt(backM[1], 10) : null;
  const hy = back === 800 ? 'HY' : '';
  const out = span === null ? null : span >= 8 ? 'F' : span <= 6 ? 'H' : null;
  if (!out || !num) return null;
  return { prg: out + 'C' + num + hy, span, cw, back, out };
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
        label: (m.out === 'F' ? '全伸 ' : '半伸 ') + m.cw + 't 后移' + (m.back === null ? '—' : m.back),
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
    if (f.span !== null && f.span <= 6 && titleStr.indexOf('支腿全伸') === 0) add('err', f.sheet, f.prg, colLetter(f.ampCol + 2) + f.titleRow, tag + ' 半伸表标题误写“支腿全伸侧方跨距5.5m”，应为“支腿半伸”');
    if (f.span !== null && f.span >= 8 && titleStr.indexOf('支腿半伸') === 0) add('err', f.sheet, f.prg, '标题行', tag + ' 全伸表标题误写“支腿半伸”');
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
    if (c1.rows.some(r => r.radius === 2.5)) add('info', f.sheet, f.prg, '', tag + ' 有 2.5m 幅度行，ST 按程序幅度点生成（含 2.5m 列，仅 FC38HY）');
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
/* 程序里各 PRG 的幅度点（与 .export / PLCopenXML 导出一致）：五臂 3.0~44.0（31 点），六臂 3.0~52.0（35 点），FC38HY 首块多一个 2.5m 点 */
function prgScope() {
  const base = state.sheet === '五臂模式' ? SCOPE.slice(0, 31) : SCOPE.slice();
  return state.prg === 'FC38HY' ? [2.5].concat(base) : base;
}
function setCell(grp, boom, rad, v) {
  const g = block()[state.src][grp];
  if (rad === 2.5) {
    if (grp === 'Half' && String(boom) === '13') {
      const exF = block()[state.src]['Full'].extra || (block()[state.src]['Full'].extra = {});
      (exF['2.5'] || (exF['2.5'] = {}))['13'] = v;
    } else {
      const ex = g.extra || (g.extra = {});
      (ex['2.5'] || (ex['2.5'] = {}))[String(boom)] = v;
    }
    return;
  }
  const row = (grp === 'Half' && String(boom) === '13') ? block()[state.src]['Full']['rows']['13'] : g.rows[String(boom)];
  if (row) row[rk(rad)] = v;
}
function excelCellFor(grp, boom, rad) {
  const b = block();
  const bl = b[state.src][grp].booms.map(String);
  let ci = bl.indexOf(String(boom));
  if (grp === 'Half') ci = String(boom) === '13' ? 0 : 9 + ci;
  if (grp === 'Third') ci = 15 + ci;
  const ampCol0 = state.sheet === '五臂模式' ? 0 : 1;
  const col = colLetter(ampCol0 + 1 + ci);
  const has25 = !!(b[state.src]['Full'].extra && b[state.src]['Full'].extra['2.5']);
  const rowIdx = rad === 2.5 ? 0 : SCOPE.indexOf(rad) + (has25 ? 1 : 0);
  return col + (b.titleRow + 3 + rowIdx);
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
  let prev = null, prevR = null;
  colVals.forEach((x, i) => {
    const cell = excelCellFor(grp, boom, x.rad);
    if (typeof x.v === 'string') {
      out.push({ level: 'err', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + x.rad + 'm 一格内有文本 ' + JSON.stringify(x.v) + '，不是数字，请核对纸质表' });
      return;
    }
    if ((x.v === 0 || x.v === null || x.v === undefined) && minI >= 0 && i > minI && i < maxI) {
      out.push({ level: 'err', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + x.rad + 'm 空白空洞' });
      return;
    }
    if (typeof x.v !== 'number') return;
    if (x.v % 100 !== 0) out.push({ level: 'warn', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + x.rad + 'm = ' + x.v + '，不是 100 的倍数（请核对纸质表）' });
    if (prev !== null && x.v > prev) out.push({ level: 'err', sheet: state.sheet, prg: state.prg, cell, text: tag + ' 臂长 ' + bm + 'm、幅度 ' + prevR + '→' + x.rad + 'm 上升 ' + prev + '→' + x.v + '，违反递减规律' });
    prev = x.v; prevR = x.rad;
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
    issues = issues.filter(i => !(i.sheet === state.sheet && i.prg === state.prg && i.text.indexOf(needle) >= 0));
    ['Full', 'Half', 'Third'].forEach(g => {
      const bl = block()[state.src][g].booms.map(String);
      if (bl.includes(String(boom)) || (g === 'Half' && String(boom) === '13')) issues = issues.concat(columnIssues(g, boom));
    });
  }
  renderValidation(); renderOut(); renderEdits(); drawErrorImage();
}
function undoEdits() {
  if (!dataOrig) return;
  state.data = JSON.parse(JSON.stringify(dataOrig));
  state.edits = {};
  issues = issuesOrig ? JSON.parse(JSON.stringify(issuesOrig)) : [];
  renderValidation(); renderAll(); renderEdits(); drawErrorImage();
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
function cellVal(grp, boom, rad) {
  const g = block()[state.src][grp];
  if (rad === 2.5) {
    let ex = (g.extra && g.extra['2.5']) || {};
    let v = ex[String(boom)];
    if (grp === 'Half' && String(boom) === '13') {
      const exF = (block()[state.src]['Full'].extra && block()[state.src]['Full'].extra['2.5']) || {};
      v = exF['13'];
    }
    return (v === null || v === undefined) ? 0 : v;
  }
  const row = (grp === 'Half' && String(boom) === '13') ? block()[state.src]['Full']['rows']['13'] : g.rows[String(boom)];
  const v = row ? row[rk(rad)] : 0;
  return (v === null || v === undefined) ? 0 : v;
}
function stLinesFor(grp, boom) {
  const saved = state.grp; state.grp = grp;
  const idx = boomIndex(boom);
  const arr = arrName();
  const row = boomRow(boom);
  state.grp = saved;
  const lines = ['\t(*' + fmtB(parseFloat(boom)) + 'm*)'];
  prgScope().forEach((rad, j) => {
    const v = cellVal(grp, boom, rad);
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
  t += 'END_IF;\n\n(*幅度点*' + (state.prg === 'FC38HY' ? '（含 2.5m）' : '') + '*)\n';
  prgScope().forEach((rad, i) => { t += '\tScope_Weight_Point[' + i + ']:= ' + rad.toFixed(2) + ';\n'; });
  t += '\n(*支腿' + (state.prg[0] === 'F' ? '全伸' : '半伸') + '工况*)\n';
  fullBooms.forEach(bm => { t += stLinesFor('Full', String(bm)) + '\n\n'; });
  prgScope().forEach((rad, j) => { t += '\tFull_Weight[9,' + j + ']:= 0 ;\n'; });
  t += '\n(*快伸工况性能表*)\n';
  ['13'].concat(halfBooms.map(String)).forEach(bm => { t += stLinesFor('Half', bm) + '\n\n'; });
  const halfCount = halfBooms.length + 1;
  for (let r = halfCount; r <= 9; r++) { prgScope().forEach((rad, j) => { t += '\tHalf_Weight[' + r + ',' + j + ']:= 0 ;\n'; }); t += '\n'; }
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
  w.innerHTML = '';
}
/* 分段复制：当前 PRG 每个臂长一段，各自带复制按钮 */
function renderSegments() {
  const el = $('segBtns'); if (!el) return;
  const b = block();
  let html = '';
  [['Full', '全伸组'], ['Half', '快伸组'], ['Third', '第三组']].forEach(([g, label]) => {
    const booms = g === 'Half' ? ['13'].concat(b[state.src][g].booms.map(String)) : b[state.src][g].booms.map(String);
    booms.forEach(bm => { html += '<button data-grp="' + g + '" data-boom="' + bm + '">' + label + ' ' + fmtB(parseFloat(bm)) + 'm 📋</button>'; });
  });
  el.innerHTML = html;
}
/* 圈图：当前选中表（PRG + 组别）画成表格图，错误红圈、待确认黄圈、改过的蓝点 */
function cellIssueLevel(grp, boom, rad) {
  const v = cellVal(grp, boom, rad);
  if (typeof v === 'string') return 'err';
  const vals = prgScope().map(r => cellVal(grp, boom, r));
  const idx = prgScope().indexOf(rad);
  let prev = null;
  for (let i = 0; i < idx; i++) { if (typeof vals[i] === 'number' && vals[i] !== 0) prev = vals[i]; }
  if (typeof v === 'number' && v !== 0 && prev !== null && v > prev) return 'err';
  if (typeof v === 'number' && v !== 0 && v % 100 !== 0) return 'warn';
  if (v === 0) {
    let hasBefore = false, hasAfter = false;
    for (let i = 0; i < idx; i++) { if (typeof vals[i] === 'number' && vals[i] !== 0) hasBefore = true; }
    for (let i = idx + 1; i < vals.length; i++) { if (typeof vals[i] === 'number' && vals[i] !== 0) hasAfter = true; }
    if (hasBefore && hasAfter) return 'err';
  }
  return null;
}
function drawErrorImage() {
  const cv = $('errCanvas'); if (!cv) return;
  const ctx = cv.getContext && cv.getContext('2d'); if (!ctx) return;
  const b = block();
  const g = b[state.src][state.grp];
  const booms = state.grp === 'Half' ? ['13'].concat(g.booms.map(String)) : g.booms.map(String);
  const radii = prgScope();
  const radW = 62, colW = 96, rowH = 26, headH = 30, titleH = 48, legH = 32;
  const W = radW + booms.length * colW, H = titleH + headH + radii.length * rowH + legH;
  cv.width = W; cv.height = H;
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, W, H);
  ctx.fillStyle = '#111'; ctx.font = 'bold 16px sans-serif'; ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
  ctx.fillText(state.sheet + ' · ' + state.prg + '（' + (PRG_LABEL[state.prg] || b.label) + '）· ' + (state.grp === 'Full' ? '全伸组' : state.grp === 'Half' ? '快伸组' : '第三组') + ' · ' + (state.src === 'orig' ? '原表' : '填充版'), 10, 18);
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
function renderAll() { renderSheets(); renderPrgs(); renderInfo(); renderBooms(); renderOut(); renderSegments(); renderEdits(); drawErrorImage(); }

/* ---------------- 上传 ---------------- */
function saveUpload(fileName) {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify({ name: fileName, time: Date.now(), data: state.data, issues }));
  } catch (e) { /* 存不下（如浏览器禁了本地存储）不影响使用，只是下次打开恢复不了 */ }
}
function useUploaded(wb, fileName, fileSize) {
  try {
    const prevKey = i => [i.sheet, i.prg, i.cell, i.text].join('|');
    const prevIssues = issues.slice();
    const { data, found } = parseWorkbook(wb);
    const nPrg = Object.values(data.sheets).reduce((s, sh) => s + Object.keys(sh).length, 0);
    if (!nPrg) { $('fileInfo').innerHTML = '<b>解析失败：</b>在 ' + esc(fileName) + ' 里没找到含“幅度(m)”的性能表。请确认文件是 JTC100G6 性能表原文件。'; return; }
    state.data = data;
    dataOrig = JSON.parse(JSON.stringify(data));
    state.edits = {};
    state.fileMeta = { name: fileName, time: Date.now(), size: fileSize || 0 };
    state.dataLabel = '已上传：' + fileName;
    $('dataBadge').textContent = '已上传文件';
    issues = validate(found);
    issuesOrig = JSON.parse(JSON.stringify(issues));
    const errs = issues.filter(i => i.level === 'err').length;
    const prevErrs = prevIssues.filter(i => i.level === 'err').length;
    const prevSet = new Set(prevIssues.map(prevKey));
    const fixed = prevIssues.filter(i => !issues.some(j => prevKey(j) === prevKey(i))).length;
    let delta = '';
    if (prevIssues.length && (state.fileMeta || prevIssues.length)) delta = '<br>与上次比：确认错误 ' + prevErrs + ' → ' + errs + ' 处' + (fixed ? '，已修复 ' + fixed + ' 处' : '') + '。';
    $('fileInfo').innerHTML = '已载入 <b>' + esc(fileName) + '</b>（' + new Date().toLocaleString() + (fileSize ? '，' + Math.round(fileSize / 1024) + ' KB' : '') + '）：解析出 ' + Object.keys(data.sheets).length + ' 张表、' + nPrg + ' 个工况块。PRG 与块的对应按标题自动识别（跨距 8.4m=全伸、5.5m=半伸）。已记下这次上传，下次打开页面自动恢复，不用重新传。' + delta;
    saveUpload(fileName);
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
  rd.onload = e => { try { useUploaded(XLSX.read(e.target.result, { type: 'array' }), file.name, file.size); } catch (err) { $('fileInfo').innerHTML = '<b>读取失败：</b>' + esc(err.message); } };
  rd.readAsArrayBuffer(file);
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
    a.download = state.sheet + '_' + state.prg + '_' + state.grp + '_圈图.png';
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  });
};
$('resetData').onclick = () => {
  try { localStorage.removeItem(LS_KEY); } catch (e) {}
  location.reload();
};
const drop = $('drop');
drop.addEventListener('click', () => $('fileInput').click());
$('fileInput').addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0]); });
['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });

/* ---------------- 启动：先恢复上次上传的 Excel；没有则用内置数据 ---------------- */
(function init() {
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
        issuesOrig = JSON.parse(JSON.stringify(issues));
        $('fileInfo').innerHTML = '已恢复上次上传的 <b>' + esc(obj.name) + '</b>（' + new Date(obj.time).toLocaleString() + '）。改完 Excel 重新上传即可替换；点“恢复内置数据”可切回 2026-11-23 内置数据。';
        renderValidation(); renderAll();
        return;
      }
    }
  } catch (e) { /* 恢复失败则用内置数据 */ }
  if (!state.data) { $('validSummary').textContent = '内置数据未加载。'; return; }
  issues = (typeof BUILTIN_ISSUES !== 'undefined' ? BUILTIN_ISSUES : []);
  issuesOrig = JSON.parse(JSON.stringify(issues));
  dataOrig = JSON.parse(JSON.stringify(state.data));
  renderValidation(); renderAll();
})();
