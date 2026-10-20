/* 국방예산 대시보드 — 엑셀/CSV를 읽어 분야별·신규사업별로 집계 */
(function () {
  'use strict';

  const PALETTE = ['#1B3A66', '#F28C28', '#159A8C', '#7A5AC8', '#C2457A', '#2F80C4', '#B88A1B', '#4C9A2A', '#8E5A2B', '#5B6B82'];
  const $ = (s) => document.querySelector(s);
  const state = { rows: [], fields: [], fileName: '', sheetNames: [], wb: null, sel: null, sort: {}, zoom: false };

  /* ---------- 숫자·형식 ---------- */
  const nf = new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 2 });
  function num(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/[,\s원억조%]/g, '').replace(/^△/, '-').replace(/^▲/, '');
    if (s === '' || s === '-' ) return null;
    const n = Number(s);
    return isFinite(n) ? n : null;
  }
  // 억원 → "1조 2,345억" / "565억"
  function eok(v) {
    if (v == null || isNaN(v)) return '—';
    const sign = v < 0 ? '−' : '';
    const a = Math.abs(v);
    if (a >= 10000) {
      const jo = Math.floor(a / 10000), rest = Math.round(a - jo * 10000);
      return sign + jo + '조' + (rest ? ' ' + nf.format(rest) + '억' : '');
    }
    return sign + nf.format(a >= 100 ? Math.round(a) : Math.round(a * 100) / 100) + '억';
  }
  function short(v) { // 축 라벨용
    if (v == null) return '';
    const a = Math.abs(v);
    if (a >= 10000) return (v / 10000).toFixed(a >= 100000 ? 0 : 1) + '조';
    return nf.format(Math.round(v)) + '억';
  }
  function pct(r) {
    if (r == null || !isFinite(r)) return '—';
    return (r > 0 ? '+' : r < 0 ? '−' : '') + Math.abs(r * 100).toFixed(1) + '%';
  }
  const cls = (r) => (r > 0 ? 'up' : r < 0 ? 'dn' : '');
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ---------- 열 인식 ---------- */
  const KEYS = [
    ['no', /^(연번|번호|no\.?|순번)$/i],
    ['y27', /(2027|27년|\(b\)|정부안|당해|금년|요구)/i],
    ['y26', /(2026|26년|\(a\)|전년|기정|본예산)/i],
    ['rate', /(증가율|증감률|%)/],
    ['diff', /(증감|차이|b-a)/i],
    ['field', /(분야|영역|부문|중점)/],
    ['name', /(사업|지표|항목|과제)/],
    ['unit', /(단위|구분)/],
    ['note', /(내용|근거|비고|설명|메모)/],
  ];
  function mapHeader(row) {
    const map = {}; let score = 0;
    row.forEach((h, i) => {
      if (h == null) return;
      const s = String(h).trim();
      for (const [k, re] of KEYS) {
        if (map[k] == null && re.test(s)) { map[k] = i; score++; break; }
      }
    });
    return { map, score };
  }
  function findTable(aoa) {
    let best = null;
    for (let i = 0; i < Math.min(10, aoa.length); i++) {
      const { map, score } = mapHeader(aoa[i] || []);
      if (map.name != null && map.y27 != null && (!best || score > best.score)) best = { i, map, score };
    }
    return best;
  }

  function parseRows(aoa) {
    const t = findTable(aoa);
    if (!t) throw new Error('열 제목을 찾지 못했습니다. 최소한 「사업명」과 「2027년(당해)」 열이 있어야 합니다.');
    const m = t.map, out = [];
    let lastField = '';
    for (let r = t.i + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const name = row[m.name];
      if (name == null || String(name).trim() === '') continue;
      const nm = String(name).trim();
      if (/^(합계|소계|총계|계)$/.test(nm) || /^※/.test(nm)) continue;
      let field = m.field != null && row[m.field] != null && String(row[m.field]).trim() !== '' ? String(row[m.field]).trim() : lastField;
      lastField = field;
      const a = m.y26 != null ? num(row[m.y26]) : null;
      const b = num(row[m.y27]);
      const rateCell = m.rate != null ? row[m.rate] : null;
      const markedNew = (rateCell != null && /신규|신설/.test(String(rateCell)));
      const isNew = markedNew || ((a == null || a === 0) && b != null && b > 0);
      const unit = m.unit != null && row[m.unit] != null ? String(row[m.unit]).trim() : '억원';
      out.push({
        no: m.no != null ? row[m.no] : out.length + 1,
        field: field || '(분야 없음)',
        name: nm,
        unit,
        money: /억|원|백만|천만/.test(unit) || unit === '',
        a, b,
        diff: a != null && b != null ? b - a : null,
        rate: a ? (b - a) / a : null,
        isNew,
        note: m.note != null && row[m.note] != null ? String(row[m.note]) : '',
      });
    }
    if (!out.length) throw new Error('데이터 행이 없습니다.');
    // 단위가 백만원/천만원이면 억원으로 환산
    out.forEach((d) => {
      const k = /백만/.test(d.unit) ? 0.01 : /천만/.test(d.unit) ? 0.1 : (/^원$/.test(d.unit) ? 1e-8 : 1);
      if (k !== 1) { ['a', 'b', 'diff'].forEach((f) => { if (d[f] != null) d[f] *= k; }); d.unit = '억원'; }
    });
    return { rows: out, missing: ['field', 'y26'].filter((k) => m[k] == null) };
  }

  /* ---------- 파일 읽기 ---------- */
  function decodeCsv(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
    catch (e) { return new TextDecoder('euc-kr').decode(buf); }
  }
  function readFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const buf = new Uint8Array(reader.result);
        let wb;
        if (/\.csv$/i.test(file.name)) wb = XLSX.read(decodeCsv(buf).replace(/^﻿/, ''), { type: 'string', raw: false });
        else wb = XLSX.read(buf, { type: 'array' });
        loadWorkbook(wb, file.name);
      } catch (err) { showAlert('파일을 읽지 못했습니다: ' + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }
  function sheetAoa(wb, name) {
    return XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
  }
  function loadWorkbook(wb, fileName) {
    state.wb = wb; state.fileName = fileName;
    // 표를 인식할 수 있는 시트 중 첫 번째
    const usable = wb.SheetNames.filter((n) => findTable(sheetAoa(wb, n)));
    if (!usable.length) { showAlert('어느 시트에서도 열 제목을 찾지 못했습니다. 「분야 · 사업명 · 2026년 · 2027년」 같은 제목 행이 필요합니다.'); return; }
    const sel = $('#sheet');
    sel.innerHTML = usable.map((n) => `<option>${esc(n)}</option>`).join('');
    sel.hidden = usable.length < 2;
    useSheet(usable[0]);
  }
  function useSheet(name) {
    try {
      const { rows, missing } = parseRows(sheetAoa(state.wb, name));
      setData(rows, state.fileName, name, missing);
    } catch (err) { showAlert(err.message); }
  }
  function loadSample() {
    const s = window.SAMPLE_DATA;
    if (!s) return;
    try {
      const { rows, missing } = parseRows(s.rows);
      $('#sheet').hidden = true;
      setData(rows, s.name, s.sheet, missing);
    } catch (err) { showAlert(err.message); }
  }

  function showAlert(msg) { const a = $('#alert'); a.textContent = msg; a.hidden = !msg; }

  function setData(rows, fileName, sheet, missing) {
    state.rows = rows; state.sel = null;
    const names = [...new Set(rows.map((r) => r.field))];
    state.fields = names.map((f, i) => {
      const items = rows.filter((r) => r.field === f);
      const money = items.filter((r) => r.money);
      const a = sum(money, 'a'), b = sum(money, 'b');
      return { field: f, color: PALETTE[i % PALETTE.length], items, a, b, diff: b - a, rate: a ? (b - a) / a : null,
        nNew: items.filter((r) => r.isNew).length, bNew: sum(money.filter((r) => r.isNew), 'b') };
    });
    const msgs = [];
    if (missing.includes('field')) msgs.push('「분야」 열이 없어 모든 사업을 한 분야로 묶었습니다.');
    if (missing.includes('y26')) msgs.push('「2026년(전년)」 열이 없어 증감·신규 판정이 제한됩니다.');
    const nonMoney = rows.filter((r) => !r.money).length;
    if (nonMoney) msgs.push(`금액이 아닌 단위(명·% 등) ${nonMoney}건은 합계에서 빼고 목록에만 표시합니다.`);
    showAlert(msgs.join(' '));
    $('#meta').textContent = `${fileName} · 시트 「${sheet}」 · ${rows.length}개 사업 · ${state.fields.length}개 분야`;
    renderAll();
  }
  const sum = (arr, k) => arr.reduce((s, r) => s + (r[k] || 0), 0);
  const fieldColor = (f) => (state.fields.find((x) => x.field === f) || {}).color || '#5B6B82';

  /* ---------- 렌더 ---------- */
  function renderAll() {
    renderKpis(); renderFieldCharts(); renderFieldCards(); renderFieldDetail();
    renderNew(); renderAllTable(); fillFilters();
  }

  function renderKpis() {
    const money = state.rows.filter((r) => r.money);
    const a = sum(money, 'a'), b = sum(money, 'b');
    const nw = money.filter((r) => r.isNew);
    const top = [...state.fields].filter((f) => f.rate != null).sort((x, y) => y.rate - x.rate)[0];
    const k = [
      { l: '당해 예산 합계', v: eok(b), d: `전년 ${eok(a)}`, c: '#1B3A66' },
      { l: '전년 대비 증감', v: `<span class="${cls(b - a)}">${pct(a ? (b - a) / a : null)}</span>`, d: `${b - a >= 0 ? '+' : ''}${eok(b - a)}`, c: '#159A8C' },
      { l: '신규사업', v: `${nw.length}<small style="font-size:.55em">개</small>`, d: `신규 예산 ${eok(sum(nw, 'b'))}`, c: '#F28C28' },
      { l: '최고 증가 분야', v: top ? `<span class="up">${pct(top.rate)}</span>` : '—', d: top ? esc(top.field) : '', c: '#7A5AC8' },
    ];
    $('#kpis').innerHTML = k.map((x) => `<div class="kpi" style="--c:${x.c}"><div class="l">${x.l}</div><div class="v">${x.v}</div><div class="d">${x.d}</div></div>`).join('');
  }

  /* SVG helpers */
  const SVGNS = 'http://www.w3.org/2000/svg';
  function el(tag, attrs, parent) { const e = document.createElementNS(SVGNS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (parent) parent.appendChild(e); return e; }
  function textW(svg, s, size, weight) { const t = el('text', { 'font-size': size, 'font-weight': weight || 400, x: -9999, y: -9999 }, svg); t.textContent = s; const w = t.getComputedTextLength(); t.remove(); return w; }
  function fit(svg, t, max) { let s = t.textContent; while (t.getComputedTextLength() > max && s.length > 1) { s = s.slice(0, -1); t.textContent = s + '…'; } }
  const tip = $('#tip');
  function showTip(ev, title, lines) {
    tip.innerHTML = `<div class="t">${esc(title)}</div>` + lines.map((l) => `<div>${esc(l)}</div>`).join('');
    tip.hidden = false;
    const w = tip.offsetWidth, h = tip.offsetHeight;
    let x = ev.clientX + 14, y = ev.clientY + 14;
    if (x + w > innerWidth - 8) x = ev.clientX - w - 14;
    if (y + h > innerHeight - 8) y = ev.clientY - h - 14;
    tip.style.left = x + 'px'; tip.style.top = y + 'px';
  }
  const hideTip = () => { tip.hidden = true; };

  /* 분야별 금액: 전년/당해 이중 막대 */
  function renderFieldCharts() {
    const box = $('#c-field-amt'); box.innerHTML = '';
    const F = state.fields;
    if (!F.length) return;
    const W = Math.max(320, box.clientWidth || 560), rowH = 44, top = 4, H = top + F.length * rowH + 24;
    const svg = el('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, box);
    const lw = Math.min(W * 0.4, Math.max(...F.map((f) => textW(svg, f.field, 13))) + 14);
    const vw = Math.max(...F.map((f) => textW(svg, short(f.b) + ' ≫', 12, 700))) + 12;
    const vals = F.map((f) => Math.max(f.a, f.b)).sort((p, q) => q - p);
    const fullMax = vals[0] || 1;
    const max = state.zoom && vals.length > 1 && vals[1] > 0 ? vals[1] * 1.15 : fullMax;
    const x = (v) => lw + (W - lw - vw) * (Math.min(v, max) / max);
    const cut = (v) => v > max * 1.0001;
    // grid
    [0, .25, .5, .75, 1].forEach((p) => {
      const gx = lw + (W - lw - vw) * p;
      el('line', { x1: gx, x2: gx, y1: top, y2: H - 20, stroke: 'var(--line)' }, svg);
      const t = el('text', { x: gx, y: H - 4, 'font-size': 11, 'text-anchor': p === 0 ? 'start' : 'middle', class: 'mu' }, svg); t.textContent = short(max * p);
    });
    F.forEach((f, i) => {
      const y = top + i * rowH;
      const t = el('text', { x: lw - 10, y: y + rowH / 2, 'font-size': 13, 'text-anchor': 'end', 'dominant-baseline': 'middle', 'font-weight': 600 }, svg); t.textContent = f.field; fit(svg, t, lw - 14);
      const g = el('g', { style: 'cursor:pointer' }, svg);
      el('rect', { x: lw, y: y + 7, width: Math.max(1, x(f.a) - lw), height: 12, rx: 4, fill: 'var(--a)' }, g);
      el('rect', { x: lw, y: y + 21, width: Math.max(1, x(f.b) - lw), height: 16, rx: 4, fill: f.color }, g);
      [[f.a, 7, 12], [f.b, 21, 16]].forEach(([val, yy, hh]) => { if (cut(val)) { const bx = x(val) - 14; el('path', { d: `M${bx},${y + yy + hh + 2} l5,${-(hh + 4)} M${bx + 5},${y + yy + hh + 2} l5,${-(hh + 4)}`, stroke: 'var(--panel)', 'stroke-width': 3 }, g); } });
      const v = el('text', { x: x(f.b) + 6, y: y + 29, 'font-size': 12, 'font-weight': 700, 'dominant-baseline': 'middle' }, g); v.textContent = short(f.b) + (cut(f.b) ? ' ≫' : '');
      el('rect', { x: 0, y, width: W, height: rowH, fill: 'transparent' }, g);
      g.addEventListener('mousemove', (ev) => showTip(ev, f.field, [`전년 ${eok(f.a)}`, `당해 ${eok(f.b)} (${pct(f.rate)})`, `사업 ${f.items.length}개 · 신규 ${f.nNew}개`]));
      g.addEventListener('mouseleave', hideTip);
      g.addEventListener('click', () => selectField(f.field));
    });

    // 증가율
    const rb = $('#c-field-rate'); rb.innerHTML = '';
    const R = F.filter((f) => f.rate != null).sort((p, q) => q.rate - p.rate);
    if (!R.length) { rb.innerHTML = '<div class="empty">전년 값이 없어 증가율을 계산할 수 없습니다.</div>'; return; }
    const W2 = Math.max(320, rb.clientWidth || 560), rh = 34, H2 = R.length * rh + 8;
    const s2 = el('svg', { viewBox: `0 0 ${W2} ${H2}`, height: H2 }, rb);
    const lw2 = Math.min(W2 * 0.4, Math.max(...R.map((f) => textW(s2, f.field, 13))) + 14);
    const vw2 = Math.max(...R.map((f) => textW(s2, pct(f.rate), 12, 700))) + 12;
    const mn = Math.min(0, ...R.map((f) => f.rate)), mx = Math.max(0, ...R.map((f) => f.rate)) || 1;
    const x2 = (v) => lw2 + (W2 - lw2 - vw2) * ((v - mn) / (mx - mn || 1));
    el('line', { x1: x2(0), x2: x2(0), y1: 0, y2: H2, stroke: 'var(--mu)', 'stroke-opacity': .4 }, s2);
    R.forEach((f, i) => {
      const y = i * rh + 4;
      const t = el('text', { x: lw2 - 10, y: y + rh / 2, 'font-size': 13, 'text-anchor': 'end', 'dominant-baseline': 'middle' }, s2); t.textContent = f.field; fit(s2, t, lw2 - 14);
      const x0 = Math.min(x2(0), x2(f.rate)), w = Math.abs(x2(f.rate) - x2(0));
      const r = el('rect', { x: x0, y: y + 8, width: Math.max(2, w), height: rh - 16, rx: 4, fill: f.rate < 0 ? 'var(--red)' : f.color, style: 'cursor:pointer' }, s2);
      r.addEventListener('mousemove', (ev) => showTip(ev, f.field, [`${eok(f.a)} → ${eok(f.b)}`, `증감 ${f.diff >= 0 ? '+' : ''}${eok(f.diff)} (${pct(f.rate)})`]));
      r.addEventListener('mouseleave', hideTip);
      r.addEventListener('click', () => selectField(f.field));
      const v = el('text', { x: f.rate >= 0 ? x2(f.rate) + 6 : x2(0) + 6, y: y + rh / 2, 'font-size': 12, 'font-weight': 700, 'dominant-baseline': 'middle', fill: f.rate < 0 ? 'var(--red)' : 'var(--ink)' }, s2); v.textContent = pct(f.rate);
    });
  }

  function renderFieldCards() {
    $('#fcards').innerHTML = state.fields.map((f) => `
      <button class="fc ${state.sel === f.field ? 'on' : ''}" style="--c:${f.color}" data-field="${esc(f.field)}">
        <span class="n">${esc(f.field)}</span>
        <span class="v">${eok(f.b)}</span>
        <span class="d">전년 ${eok(f.a)} · <span class="${cls(f.rate)}">${pct(f.rate)}</span></span>
        <span class="chips"><span class="chip">사업 ${f.items.length}</span>${f.nNew ? `<span class="chip new">신규 ${f.nNew} · ${eok(f.bNew)}</span>` : ''}</span>
      </button>`).join('');
    $('#fcards').querySelectorAll('.fc').forEach((b) => b.addEventListener('click', () => selectField(b.dataset.field)));
  }
  function selectField(f) {
    state.sel = state.sel === f ? null : f;
    renderFieldCards(); renderFieldDetail();
    if (state.sel) $('#field-detail').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
  function renderFieldDetail() {
    const box = $('#c-field-items'); box.innerHTML = '';
    const f = state.fields.find((x) => x.field === state.sel);
    $('#fd-clear').hidden = !f;
    if (!f) { $('#fd-title').textContent = '분야를 선택하세요'; box.innerHTML = '<div class="empty">위의 분야 카드나 막대를 누르면 해당 분야의 사업별 전년·당해 예산을 비교합니다.</div>'; return; }
    $('#fd-title').textContent = `${f.field} — 사업별 비교`;
    const items = f.items.filter((r) => r.money && r.b != null).sort((p, q) => q.b - p.b);
    const W = Math.max(320, box.clientWidth || 900), rh = 46, H = items.length * rh + 6;
    const svg = el('svg', { viewBox: `0 0 ${W} ${H}`, height: H }, box);
    const lw = Math.min(W * 0.42, Math.max(...items.map((r) => textW(svg, r.name, 13))) + 14);
    const lab = (r) => `${short(r.b)}  ${r.isNew ? '신규' : pct(r.rate)}`;
    const vw = Math.max(...items.map((r) => textW(svg, lab(r), 12, 700))) + 14;
    const max = Math.max(...items.map((r) => Math.max(r.a || 0, r.b))) || 1;
    const x = (v) => lw + (W - lw - vw) * ((v || 0) / max);
    items.forEach((r, i) => {
      const y = i * rh;
      const t = el('text', { x: lw - 10, y: y + rh / 2, 'font-size': 13, 'text-anchor': 'end', 'dominant-baseline': 'middle' }, svg); t.textContent = r.name; fit(svg, t, lw - 14);
      const g = el('g', {}, svg);
      el('rect', { x: lw, y: y + 9, width: Math.max(1, x(r.a) - lw), height: 10, rx: 3, fill: 'var(--a)' }, g);
      el('rect', { x: lw, y: y + 21, width: Math.max(1, x(r.b) - lw), height: 16, rx: 4, fill: r.isNew ? 'var(--or)' : f.color }, g);
      const v = el('text', { x: x(r.b) + 6, y: y + 29, 'font-size': 12, 'font-weight': 700, 'dominant-baseline': 'middle', fill: r.isNew ? 'var(--or)' : 'var(--ink)' }, g); v.textContent = lab(r);
      el('rect', { x: 0, y, width: W, height: rh, fill: 'transparent' }, g);
      g.addEventListener('mousemove', (ev) => showTip(ev, r.name, [`전년 ${eok(r.a)} → 당해 ${eok(r.b)}`, r.isNew ? '신규사업' : `증감 ${pct(r.rate)}`, r.note].filter(Boolean)));
      g.addEventListener('mouseleave', hideTip);
    });
  }

  /* 신규사업 */
  function renderNew() {
    const nw = state.rows.filter((r) => r.isNew && r.money && r.b != null).sort((p, q) => q.b - p.b);
    const rank = $('#c-new-rank');
    if (!nw.length) { rank.innerHTML = '<div class="empty">신규사업이 없습니다.</div>'; $('#c-new-donut').innerHTML = ''; $('#new-legend').innerHTML = ''; $('#t-new').innerHTML = ''; return; }
    const total = sum(nw, 'b'), max = nw[0].b || 1;
    rank.innerHTML = `<div class="rank">${nw.map((r, i) => `
      <div class="rr" data-i="${i}"><div class="k">${i + 1}</div>
        <div class="nm"><i style="background:${fieldColor(r.field)}"></i><span title="${esc(r.name)}">${esc(r.name)}</span></div>
        <div class="tr"><div class="fl" style="width:${Math.max(1.5, r.b / max * 100)}%;background:${fieldColor(r.field)}"></div></div>
        <div class="vl"><b>${eok(r.b)}</b><small>${(r.b / total * 100).toFixed(1)}%</small></div></div>`).join('')}</div>`;
    rank.querySelectorAll('.rr').forEach((row) => {
      const r = nw[+row.dataset.i];
      row.addEventListener('mousemove', (ev) => showTip(ev, r.field, [r.name, `당해 ${eok(r.b)} (신규 합계의 ${(r.b / total * 100).toFixed(1)}%)`, r.note].filter(Boolean)));
      row.addEventListener('mouseleave', hideTip);
    });

    // donut by field
    const byF = state.fields.map((f) => ({ f: f.field, c: f.color, v: sum(nw.filter((r) => r.field === f.field), 'b'), n: nw.filter((r) => r.field === f.field).length })).filter((d) => d.v > 0).sort((p, q) => q.v - p.v);
    const box = $('#c-new-donut'); box.innerHTML = '';
    const S = 220, R = 100, r0 = 62;
    const svg = el('svg', { viewBox: `0 0 ${S} ${S}`, height: S }, box);
    let ang = -Math.PI / 2;
    byF.forEach((d) => {
      const a1 = ang, a2 = ang + (d.v / total) * Math.PI * 2; ang = a2;
      const large = a2 - a1 > Math.PI ? 1 : 0;
      const p = (rad, a) => [S / 2 + rad * Math.cos(a), S / 2 + rad * Math.sin(a)];
      const [x1, y1] = p(R, a1), [x2, y2] = p(R, a2 - 0.0001), [x3, y3] = p(r0, a2 - 0.0001), [x4, y4] = p(r0, a1);
      const path = el('path', { d: `M${x1},${y1}A${R},${R} 0 ${large} 1 ${x2},${y2}L${x3},${y3}A${r0},${r0} 0 ${large} 0 ${x4},${y4}Z`, fill: d.c, stroke: 'var(--panel)', 'stroke-width': 2 }, svg);
      path.addEventListener('mousemove', (ev) => showTip(ev, d.f, [`신규 ${d.n}개 · ${eok(d.v)}`, `${(d.v / total * 100).toFixed(1)}%`]));
      path.addEventListener('mouseleave', hideTip);
    });
    const t1 = el('text', { x: S / 2, y: S / 2 - 8, 'text-anchor': 'middle', 'font-size': 12, class: 'mu' }, svg); t1.textContent = `신규 ${nw.length}개`;
    const t2 = el('text', { x: S / 2, y: S / 2 + 16, 'text-anchor': 'middle', 'font-size': 18, 'font-weight': 800 }, svg); t2.textContent = short(total);
    $('#new-legend').innerHTML = byF.map((d) => `<div><i style="background:${d.c}"></i><span>${esc(d.f)}</span><b>${eok(d.v)} · ${(d.v / total * 100).toFixed(1)}%</b></div>`).join('');

    table('#t-new', state.rows.filter((r) => r.isNew), [
      ['no', '연번', 'num'], ['field', '분야'], ['name', '사업명'], ['b', '당해 예산', 'num', eok], ['note', '내용', 'desc'],
    ], 'new');
  }

  /* 표 (정렬) */
  function table(sel, rows, cols, key) {
    const st = state.sort[key] || { k: null, dir: 1 };
    let data = [...rows];
    if (st.k) data.sort((p, q) => {
      const a = p[st.k], b = q[st.k];
      if (a == null) return 1; if (b == null) return -1;
      return (typeof a === 'number' && typeof b === 'number' ? a - b : String(a).localeCompare(String(b), 'ko')) * st.dir;
    });
    const t = $(sel);
    t.innerHTML = `<thead><tr>${cols.map(([k, h, c]) => `<th data-k="${k}" class="${c === 'num' ? 'num' : ''}" ${st.k === k ? `aria-sort="${st.dir > 0 ? 'ascending' : 'descending'}"` : ''}>${h}</th>`).join('')}</tr></thead>
      <tbody>${data.length ? data.map((r) => `<tr>${cols.map(([k, , c, f]) => {
        let v = r[k];
        if (k === 'rate') return `<td class="num ${r.isNew ? 'hl' : cls(v)}">${r.isNew ? '신규' : pct(v)}</td>`;
        if (k === 'field') return `<td><span class="chip" style="background:${fieldColor(v)}22;color:inherit">${esc(v)}</span></td>`;
        if (k === 'name') return `<td><b>${esc(v)}</b>${r.isNew && key === 'all' ? ' <span class="chip new">신규</span>' : ''}</td>`;
        return `<td class="${c || ''}">${f ? (r.money ? f(v) : (v == null ? '—' : nf.format(v) + ' ' + esc(r.unit))) : esc(v ?? '')}</td>`;
      }).join('')}</tr>`).join('') : `<tr><td colspan="${cols.length}" class="empty">조건에 맞는 사업이 없습니다.</td></tr>`}</tbody>`;
    t.querySelectorAll('th').forEach((th) => th.addEventListener('click', () => {
      const k = th.dataset.k;
      state.sort[key] = { k, dir: st.k === k ? -st.dir : (['a', 'b', 'diff', 'rate'].includes(k) ? -1 : 1) };
      key === 'all' ? renderAllTable() : renderNew();
    }));
  }

  function fillFilters() {
    const s = $('#f-field'), cur = s.value;
    s.innerHTML = '<option value="">전체 분야</option>' + state.fields.map((f) => `<option>${esc(f.field)}</option>`).join('');
    if (state.fields.some((f) => f.field === cur)) s.value = cur;
  }
  function renderAllTable() {
    const q = $('#q').value.trim().toLowerCase(), ff = $('#f-field').value, kind = $('#f-kind').value;
    const rows = state.rows.filter((r) =>
      (!ff || r.field === ff) && (!kind || (kind === 'new') === r.isNew) &&
      (!q || [r.name, r.field, r.note].join(' ').toLowerCase().includes(q)));
    table('#t-all', rows, [
      ['no', '연번', 'num'], ['field', '분야'], ['name', '사업명'], ['a', '전년', 'num', eok], ['b', '당해', 'num', eok],
      ['diff', '증감', 'num', (v) => (v > 0 ? '+' : '') + eok(v)], ['rate', '증가율', 'num'], ['note', '내용', 'desc'],
    ], 'all');
  }

  /* 집계 CSV 저장 */
  function exportCsv() {
    if (!state.rows.length) return;
    const lines = [['구분', '분야', '사업명', '단위', '전년', '당해', '증감', '증가율', '신규여부']];
    state.fields.forEach((f) => lines.push(['분야합계', f.field, '', '억원', f.a, f.b, f.diff, f.rate == null ? '' : (f.rate * 100).toFixed(1) + '%', `신규 ${f.nNew}개`]));
    state.rows.forEach((r) => lines.push(['사업', r.field, r.name, r.unit, r.a ?? '', r.b ?? '', r.diff ?? '', r.isNew ? '신규' : (r.rate == null ? '' : (r.rate * 100).toFixed(1) + '%'), r.isNew ? 'Y' : '']));
    const csv = '﻿' + lines.map((l) => l.map((v) => /[",\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : v).join(',')).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = (state.fileName.replace(/\.[^.]+$/, '').replace(/\s*\(예시\)/, '') || '예산') + '_집계.csv';
    a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  /* ---------- 이벤트 ---------- */
  const drop = $('#drop'), file = $('#file');
  drop.addEventListener('click', () => file.click());
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); file.click(); } });
  file.addEventListener('change', () => { if (file.files[0]) readFile(file.files[0]); file.value = ''; });
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) readFile(f); });
  // 페이지 어디에 놓아도 읽기
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => { e.preventDefault(); const f = e.dataTransfer && e.dataTransfer.files[0]; if (f && !drop.contains(e.target)) readFile(f); });
  $('#sheet').addEventListener('change', (e) => useSheet(e.target.value));
  document.querySelectorAll('#zoom button').forEach((btn) => btn.addEventListener('click', () => {
    state.zoom = btn.dataset.z === '1';
    document.querySelectorAll('#zoom button').forEach((x) => x.classList.toggle('on', x === btn));
    renderFieldCharts();
  }));
  $('#btn-sample').addEventListener('click', loadSample);
  $('#btn-export').addEventListener('click', exportCsv);
  $('#btn-print').addEventListener('click', () => window.print());
  $('#fd-clear').addEventListener('click', () => selectField(state.sel));
  ['input', 'change'].forEach((t) => { $('#q').addEventListener(t, renderAllTable); });
  $('#f-field').addEventListener('change', renderAllTable);
  $('#f-kind').addEventListener('change', renderAllTable);
  document.querySelectorAll('.tab').forEach((b) => b.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((x) => { x.classList.toggle('on', x === b); x.setAttribute('aria-selected', x === b); });
    ['field', 'new', 'all'].forEach((k) => { $('#tab-' + k).hidden = k !== b.dataset.tab; });
    if (state.rows.length) renderAll(); // 숨김 상태에서 그린 차트 폭 보정
  }));
  let rt; window.addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(() => state.rows.length && renderAll(), 150); });

  loadSample();
})();
