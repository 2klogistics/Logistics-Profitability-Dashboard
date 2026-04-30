const PAGES = [
  { id: 'trend', title: 'สรุปภาพรวมและดัชนีชี้วัดผลประกอบการหลัก' },
  { id: 'ranking', title: 'การจัดลำดับเส้นทางตามผลตอบแทนสุทธิ' },
  { id: 'customer', title: 'อัตราผลตอบแทนและส่วนต่างกำไรรายลูกค้า' },
  { id: 'ownout', title: 'สัดส่วนการใช้รถบริษัทและรถรับจ้างภายนอก' },
  { id: 'loss', title: 'ประสิทธิภาพของกลุ่มเที่ยววิ่งที่มีส่วนต่างขาดทุน' },
  { id: 'vehicle', title: 'วิเคราะห์ประสิทธิภาพประเภทรถ' },
  { id: 'driver', title: 'ประเมินผลงานพนักงานขับรถ' },
  { id: 'comparison', title: 'การวิเคราะห์ส่วนต่างและความผันผวนของผลประกอบการรายเดือน' },
  { id: 'fraud', title: 'ตรวจสอบความโปร่งใสและประสิทธิภาพการดำเนินงาน' }
];
const MONTHS = ['January', 'February', 'March', 'April'];
const MTH = { January: 'ม.ค.', February: 'ก.พ.', March: 'มี.ค.', April: 'เม.ย.' };
const COLORS = ['#3b82f6', '#6366f1', '#22c55e', '#f59e0b', '#ef4444', '#06b6d4', '#ec4899', '#8b5cf6', '#14b8a6', '#f97316'];
let DATA = null, currentPage = 0;

// Helpers
const fmt = n => n == null ? '-' : Number(n).toLocaleString('th-TH');
const fmtB = n => n == null ? '-' : Number(n).toLocaleString('th-TH', { minimumFractionDigits: 0 });
const fmtP = n => n == null ? '-' : Number(n).toFixed(2) + '%';
const pc = n => n >= 0 ? 'positive' : 'negative';
const tag = (t, c) => `<span class="tag tag-${c}">${t}</span>`;
const esc = s => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// Sortable+Searchable table engine
let tableStates = {};
function mkTable(id, cols, data, opts = {}) {
  let s = tableStates[id];
  if (!s || !opts._restore) {
    s = tableStates[id] = { col: opts.defaultSort || 0, asc: opts.defaultAsc !== false, filter: '', page: 0, perPage: 50, _cols: cols, _data: data };
  } else {
    s._cols = cols; s._data = data;
  }
  function render() {
    const wasFocused = document.activeElement?.id === id + '_q';
    const caretPos = wasFocused ? document.activeElement.selectionStart : 0;
    const q = s.filter.toLowerCase();
    let fd = s._data.filter(r => !q || cols.some((_, i) => String(r[i] || '').toLowerCase().includes(q)));
    fd.sort((a, b) => {
      const av = a[s.col], bv = b[s.col];
      const n = typeof av === 'number' && typeof bv === 'number';
      const r = n ? (av - bv) : String(av || '').localeCompare(String(bv || ''));
      return s.asc ? r : -r;
    });
    const total = fd.length, pages = Math.ceil(total / s.perPage), start = s.page * s.perPage;
    const pd = fd.slice(start, start + s.perPage);
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = `
      <div style="display:flex;gap:10px;align-items:center;margin-bottom:12px;flex-wrap:wrap">
        <input id="${id}_q" type="text" placeholder="ค้นหา..." value="${esc(s.filter)}"
          style="flex:1;min-width:180px;padding:8px 12px;background:var(--surface);border:1px solid var(--border);border-radius:8px;color:var(--text);font-size:13px">
        <span style="font-size:12px;color:var(--muted)">${fmt(total)} รายการ | หน้า ${s.page + 1}/${pages || 1}</span>
        <div style="display:flex;gap:4px">
          <button onclick="tblPage('${id}',-1)" style="${btnSt}">‹</button>
          <button onclick="tblPage('${id}',1)" style="${btnSt}">›</button>
        </div>
        <select id="${id}_pp" style="padding:6px 8px;background:var(--surface);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:12px">
          ${[25, 50, 100, 250, 500].map(v => `<option value="${v}"${v === s.perPage ? ' selected' : ''}>${v} แถว</option>`).join('')}
        </select>
      </div>
      <div class="table-wrap"><table><thead><tr>${cols.map((c, i) => `<th onclick="tblSort('${id}',${i})" style="cursor:pointer;user-select:none">${esc(c)}${s.col === i ? (s.asc ? ' ▲' : ' ▼') : ''}</th>`).join('')}</tr></thead>
      <tbody>${pd.map((r, ri) => `<tr>${r.map((v, ci) => {
      if (typeof v === 'number') { return `<td class="${v < 0 ? 'negative' : v > 0 ? 'positive' : ''}">${fmt(v)}</td>` }
      if (typeof v === 'string' && v.startsWith('<')) { return `<td>${v}</td>` }
      return `<td>${esc(v)}</td>`
    }).join('')}</tr>`).join('')}</tbody></table></div>`;
    const inp = document.getElementById(id + '_q');
    if (inp) {
      inp.addEventListener('input', e => { s.filter = e.target.value; s.page = 0; render(); });
      if (wasFocused) { inp.focus(); try { inp.setSelectionRange(caretPos, caretPos); } catch (e) { } }
    }
    const pp = document.getElementById(id + '_pp');
    if (pp) pp.addEventListener('change', e => { s.perPage = +e.target.value; s.page = 0; render(); });
  }
  render();
}
const btnSt = 'padding:4px 10px;background:var(--card);border:1px solid var(--border);border-radius:6px;color:var(--text);cursor:pointer;font-size:13px';
function tblSort(id, col) {
  const s = tableStates[id]; if (!s) return;
  if (s.col === col) s.asc = !s.asc; else { s.col = col; s.asc = false; }
  s.page = 0; mkTable(id, s._cols, s._data, { _restore: true });
}
function tblPage(id, d) {
  const s = tableStates[id]; if (!s) return;
  const q = s.filter.toLowerCase();
  const fd = s._data.filter(r => !q || s._cols.some((_, i) => String(r[i] || '').toLowerCase().includes(q)));
  const pages = Math.ceil(fd.length / s.perPage);
  s.page = Math.max(0, Math.min(s.page + d, pages - 1));
  mkTable(id, s._cols, s._data, { _restore: true });
}


// KPI card
const kpi = (label, val, cls = '', sub = '') => `<div class="kpi"><div class="kpi-label">${label}</div><div class="kpi-value ${cls}">${val}</div>${sub ? `<div class="kpi-sub">${sub}</div>` : ''}</div>`;

// Professional bar chart — 3-col layout: label | bar | value (value always visible outside bar)
function barChart(items, getLabel, getW, getVal, getColor, getSub, hideW, whiteVal) {
  return `<div style="padding:8px 0">${items.map((it, i) => {
    const w = Math.max(1, getW(it, i)).toFixed(1);
    const color = getColor ? getColor(it, i) : COLORS[i % 10];
    const label = String(getLabel(it, i));
    const val = getVal(it, i);
    const sub = getSub ? getSub(it, i) : '';
    const valColor = whiteVal ? '#ffffff' : color;
    return `<div style="display:grid;grid-template-columns:150px 1fr 150px;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.03)">
      <div style="text-align:right;line-height:1.4">
        <div style="font-size:12px;color:var(--text);white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(label)}">${esc(label.length > 22 ? label.substring(0, 21) + '\u2026' : label)}</div>
        ${sub ? `<div style="font-size:10px;color:var(--muted)">${sub}</div>` : ''}
      </div>
      <div style="background:rgba(255,255,255,0.05);border-radius:6px;height:26px;overflow:hidden">
        <div style="width:${w}%;height:100%;background:linear-gradient(90deg,${color}66,${color});border-radius:6px;box-shadow:0 2px 10px ${color}2a"></div>
      </div>
      <div style="font-size:12px;font-weight:700;color:${valColor};white-space:nowrap;padding-left:4px">${val}</div>
    </div>`;
  }).join('')}</div>`;
}

// Page builders
function buildTrend(d) {
  const s = d.summary;
  let h = `<div class="kpi-grid">${kpi('จำนวนเที่ยวทั้งหมด', fmt(s.totalTrips), 'accent')}${kpi('ราคารับรวม', fmt(s.totalRevenue) + ' THB', 'green')}${kpi('ส่วนต่างรวม', fmt(s.totalMargin) + ' THB', 'green')}${kpi('กำไร % เฉลี่ย', fmtP(s.avgMarginPct), '')}</div>`;
  const rows = d.routeTrend.map(r => {
    const tot = MONTHS.reduce((a, m) => a + (r.months[m]?.trips || 0), 0);
    return [r.customer, r.vtype || '-', r.route, tot,
    ...MONTHS.flatMap(m => [r.months[m]?.trips || 0, r.months[m]?.margin || 0])
    ];
  });
  const cols = ['ลูกค้า', 'ประเภทรถ', 'เส้นทาง (Route)', 'จำนวนเที่ยวรวม', ...MONTHS.flatMap(m => [MTH[m] + ' (เที่ยว)', MTH[m] + ' (ส่วนต่าง)'])];
  h += `<div class="table-card"><div class="table-card-header"><h3>สรุปผลการดำเนินงานรายเส้นทางและแนวโน้มส่วนต่างกำไรประจำเดือน</h3></div><div class="table-wrap" id="t_trend"></div></div>`;
  setTimeout(() => mkTable('t_trend', cols, rows, { defaultSort: 3, defaultAsc: false }), 0);
  return h;
}

function buildRanking(d) {
  const rk = d.routeRanking;
  const mkRows = arr => arr.map(r => [r.customer, r.route, r.desc || '-', r.trips, r.margin, r.avgMargin, r.pct, r.loss]);
  const cols = ['ลูกค้า', 'เส้นทาง (Route)', 'ชื่อเส้นทาง', 'จำนวนเที่ยว', 'ส่วนต่างรวม', 'ส่วนต่างเฉลี่ย/เที่ยว', 'กำไร %', 'เที่ยวขาดทุน'];
  let h = `<div class="kpi-grid">${kpi('เส้นทางดีสุด', fmtB(rk.top[0]?.margin) + ' THB', 'green', rk.top[0]?.route || '')}${kpi('เส้นทางขาดทุนสูงสุด', fmtB(rk.bottom[0]?.margin) + ' THB', 'red', rk.bottom[0]?.route || '')}</div>`;
  h += `<div class="table-card"><div class="table-card-header"><h3>จัดอันดับเส้นทางที่มีส่วนต่างเป็นกำไร</h3></div><div id="t_top"></div></div>`;
  h += `<div class="table-card" style="margin-top:24px"><div class="table-card-header"><h3>จัดอันดับเส้นทางที่มีส่วนต่างเป็นขาดทุน</h3></div><div id="t_bot"></div></div>`;
  setTimeout(() => { mkTable('t_top', cols, mkRows(rk.top), { defaultSort: 4, defaultAsc: false }); mkTable('t_bot', cols, mkRows(rk.bottom), { defaultSort: 4, defaultAsc: true }); }, 0);
  return h;
}

function buildCustomer(d) {
  const cp = d.customerProfit;
  const s = d.summary;
  let h = `<div class="kpi-grid">
    ${kpi('จำนวนลูกค้า', cp.length, 'accent')}
    ${kpi('ราคารับรวม', fmt(s.totalRevenue) + ' THB', 'green')}
    ${kpi('ส่วนต่างรวม', fmt(s.totalMargin) + ' THB', 'green')}
    ${kpi('กำไร % เฉลี่ย', fmtP(s.avgMarginPct), '')}
  </div>`;
  const maxR = Math.max(...cp.map(c => c.recv));
  h += `<div class="chart-card"><h3>รายได้และส่วนต่างกำไรแยกตามลูกค้า</h3>
    ${barChart(cp, c => c.name, c => c.recv / maxR * 100, c => fmt(c.recv) + ' THB', null, c => fmt(c.trips) + ' เที่ยว', true, true)}
  </div>`;
  // separate margin chart
  const maxM = Math.max(...cp.filter(c => c.margin > 0).map(c => c.margin), 1);
  h += `<div class="chart-card"><h3>ส่วนต่างกำไรสุทธิแยกตามลูกค้า</h3>
    ${barChart(cp, c => c.name, c => c.margin <= 0 ? 1 : c.margin / maxM * 100, c => fmt(c.margin) + ' THB', c => c.margin >= 0 ? COLORS[2] : '#ef4444', null, true)}
  </div>`;

  const rows = cp.map(c => [c.name, c.trips, c.recv, c.margin, c.pct, c.loss, c.oil, ...MONTHS.flatMap(m => [c.months[m]?.trips || 0, c.months[m]?.margin || 0])]);
  const cols = ['ลูกค้า', 'จำนวนเที่ยว', 'ราคารับ', 'ส่วนต่าง', 'กำไร %', 'เที่ยวขาดทุน', 'จ่ายสำรองน้ำมัน', ...MONTHS.flatMap(m => [MTH[m] + ' (เที่ยว)', MTH[m] + ' (ส่วนต่าง)'])];
  h += `<div class="table-card"><div class="table-card-header"><h3>ภาพรวมผลประกอบการและเสถียรภาพรายได้จำแนกตามรายลูกค้า</h3></div><div id="t_cust"></div></div>`;
  setTimeout(() => mkTable('t_cust', cols, rows, { defaultSort: 3, defaultAsc: false }), 0);
  return h;
}
function buildOwnOut(d) {
  const oo = d.ownVsOutsource, co = oo.company, ou = oo.outsource;
  const tot = co.trips + ou.trips;
  const coP = (co.trips / tot * 100).toFixed(1), ouP = (ou.trips / tot * 100).toFixed(1);
  let h = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px">
      <div class="kpi">
        <div class="kpi-label" style="color:var(--accent)">รถบริษัท</div>
        <div class="kpi-value">${fmt(co.trips)} <span style="font-size:13px;font-weight:400;color:var(--muted)">เที่ยว</span></div>
        <div style="margin:10px 0 4px;background:rgba(255,255,255,.06);border-radius:5px;height:8px;overflow:hidden"><div style="width:${coP}%;height:100%;background:linear-gradient(90deg,#3b82f677,#3b82f6);border-radius:5px"></div></div>
        <div style="font-size:11px;color:var(--muted);margin-bottom:10px">${coP}% ของเที่ยวทั้งหมด</div>
        <div style="font-size:12px;color:var(--muted)">ราคารับ: <b style="color:var(--text)">${fmt(co.recv)} THB</b></div>
        <div style="font-size:12px;margin-top:4px;color:${co.margin >= 0 ? 'var(--green)' : 'var(--red)'}">ส่วนต่าง: <b>${fmt(co.margin)} THB</b> (${fmtP(co.pct)})</div>
      </div>
      <div class="kpi">
        <div class="kpi-label" style="color:var(--accent2)">รถจ้างภายนอก</div>
        <div class="kpi-value">${fmt(ou.trips)} <span style="font-size:13px;font-weight:400;color:var(--muted)">เที่ยว</span></div>
        <div style="margin:10px 0 4px;background:rgba(255,255,255,.06);border-radius:5px;height:8px;overflow:hidden"><div style="width:${ouP}%;height:100%;background:linear-gradient(90deg,#6366f177,#6366f1);border-radius:5px"></div></div>
        <div style="font-size:11px;color:var(--muted);margin-bottom:10px">${ouP}% ของเที่ยวทั้งหมด</div>
        <div style="font-size:12px;color:var(--muted)">ราคารับ: <b style="color:var(--text)">${fmt(ou.recv)} THB</b></div>
        <div style="font-size:12px;margin-top:4px;color:${ou.margin >= 0 ? 'var(--green)' : 'var(--red)'}">ส่วนต่าง: <b>${fmt(ou.margin)} THB</b> (${fmtP(ou.pct)})</div>
      </div>
    </div>`;

  const mkR = arr => arr.map(r => [r.route, r.trips, r.margin]);
  const cols = ['เส้นทาง (Route)', 'จำนวนเที่ยว', 'ส่วนต่าง'];
  ['company', 'outsource'].forEach(k => {
    const label = k === 'company' ? 'รถบริษัท' : 'รถจ้างภายนอก';
    h += `<div class="table-card"><div class="table-card-header"><h3>${label}</h3></div><div id="t_oo_${k}"></div></div>`;
  });
  setTimeout(() => {
    mkTable('t_oo_company', cols, mkR(co.topRoutes), { defaultSort: 1, defaultAsc: false });
    mkTable('t_oo_outsource', cols, mkR(ou.topRoutes), { defaultSort: 1, defaultAsc: false });
  }, 0);
  return h;
}


function buildLoss(d) {
  const lt = d.lossTrip;
  if (!lt) { return '<div class="kpi"><div class="kpi-value red">No loss data</div></div>'; }

  // byMonth bar chart — safe access
  const validMonths = MONTHS.filter(m => lt.byMonth && lt.byMonth[m]);
  const maxL = validMonths.length > 0 ? Math.max(...validMonths.map(m => Math.abs(lt.byMonth[m].loss || 0))) : 1;
  let monthBar = '';
  if (validMonths.length > 0) {
    const maxC = Math.max(...validMonths.map(m => lt.byMonth[m].count || 0), 1);
    monthBar = `<div class="chart-card"><h3>จำนวนเที่ยวขาดทุนแยกเป็นรายเดือน</h3><div style="padding:8px 0">`;
    validMonths.forEach(m => {
      const bm = lt.byMonth[m];
      const wC = Math.max(2, (bm.count || 0) / maxC * 100).toFixed(1);
      const wL = Math.max(2, (Math.abs(bm.loss || 0) / maxL * 100)).toFixed(1);
      monthBar += `<div style="margin-bottom:16px;padding-bottom:16px;border-bottom:1px solid rgba(255,255,255,0.05)">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <span style="font-size:13px;font-weight:700;color:var(--text)">${MTH[m] || m}</span>
          <div style="display:flex;gap:20px">
            <span style="font-size:12px;color:var(--muted)">เที่ยว: <b style="color:var(--text)">${bm.count}</b></span>
            <span style="font-size:12px;color:#ef4444;font-weight:600">มูลค่า: <b>${fmt(bm.loss)} THB</b></span>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:80px 1fr;gap:8px;align-items:center">
          <div style="font-size:10px;color:var(--muted);text-align:right">จำนวนเที่ยว</div>
          <div style="background:rgba(255,255,255,0.05);border-radius:5px;height:22px;overflow:hidden">
            <div style="width:${wC}%;height:100%;background:linear-gradient(90deg,#f59e0b77,#f59e0b);border-radius:5px;display:flex;align-items:center;padding-left:8px;box-shadow:0 2px 8px #f59e0b33">
              <span style="font-size:11px;font-weight:700;color:#fff">${bm.count} เที่ยว</span>
            </div>
          </div>
          <div style="font-size:10px;color:var(--muted);text-align:right">มูลค่าขาดทุน</div>
          <div style="background:rgba(255,255,255,0.05);border-radius:5px;height:22px;overflow:hidden">
            <div style="width:${wL}%;height:100%;background:linear-gradient(90deg,#b9191977,#b91919);border-radius:5px;display:flex;align-items:center;padding-left:8px;box-shadow:0 2px 8px #b9191933">
              <span style="font-size:11px;font-weight:700;color:#fff">${fmt(bm.loss)} THB</span>
            </div>
          </div>
        </div>
      </div>`;
    });
    monthBar += `</div></div>`;
  }

  // KPI summary
  let h = `<div class="kpi-grid">
    ${kpi('เที่ยวขาดทุน', fmt(lt.total), 'red', 'จาก ' + fmt(lt.totalTrips) + ' เที่ยว')}
    ${kpi('อัตราขาดทุน', fmtP(lt.lossPct), 'red', '')}
    ${kpi('มูลค่าขาดทุนรวม', fmt(lt.totalLoss) + ' THB', 'red', '')}
    ${kpi('ขาดทุนเฉลี่ย/เที่ยว', lt.total > 0 ? fmt(Math.round((lt.totalLoss || 0) / lt.total)) + ' THB' : '-', 'red', '')}
  </div>`;

  h += monthBar;

  // byCustomer table — handle array or object
  const custArr = Array.isArray(lt.byCustomer) ? lt.byCustomer :
    lt.byCustomer ? Object.entries(lt.byCustomer).map(([k, v]) => ({ name: k, count: v.count, loss: v.loss })) : [];
  const custRows = custArr.map(c => [c.name || '-', c.count || 0, c.loss || 0]);

  // byRoute table
  const routeArr = Array.isArray(lt.byRoute) ? lt.byRoute :
    lt.byRoute ? Object.entries(lt.byRoute).map(([k, v]) => ({ name: k, count: v.count, loss: v.loss })) : [];
  const routeRows = routeArr.map(r => [r.name || '-', r.count || 0, r.loss || 0]);

  h += `<div class="table-card" style="margin-bottom:24px"><div class="table-card-header"><h3>จำนวนเที่ยวที่ขาดทุนแยกตามรายลูกค้า</h3></div><div id="t_lc"></div></div>`;
  h += `<div class="table-card"><div class="table-card-header"><h3>จำนวนเที่ยวที่ขาดทุนแยกตามเส้นทาง</h3></div><div id="t_lr"></div></div>`;

  setTimeout(() => {
    if (document.getElementById('t_lc'))
      mkTable('t_lc', ['ลูกค้า', 'เที่ยวขาดทุน', 'มูลค่าขาดทุน (THB)'], custRows, { defaultSort: 2, defaultAsc: true });
    if (document.getElementById('t_lr'))
      mkTable('t_lr', ['เส้นทาง (Route)', 'เที่ยวขาดทุน', 'มูลค่าขาดทุน (THB)'], routeRows, { defaultSort: 1, defaultAsc: false });
  }, 0);
  return h;
}


function buildVehicle(d) {
  const vt = d.vehicleType;
  const maxT = Math.max(...vt.map(v => v.trips));
  const maxM = Math.max(...vt.filter(v => v.margin > 0).map(v => v.margin), 1);
  let h = `<div class="chart-card"><h3>สัดส่วนเที่ยวแยกตามประเภทรถ</h3>
    ${barChart(vt, v => v.type, v => v.trips / maxT * 100, v => fmt(v.trips) + ' เที่ยว', (v, i) => COLORS[i % 10], v => 'สัดส่วน ' + v.share + '%', true, true)}
  </div>`;
  h += `<div class="chart-card"><h3>ส่วนต่างเฉลี่ยสุทธิ/เที่ยว แยกตามประเภทรถ</h3>
    ${barChart(vt, v => v.type, v => v.avgMargin <= 0 ? 1 : v.avgMargin / Math.max(...vt.map(x => x.avgMargin || 0), 1) * 100, v => fmt(v.avgMargin) + ' บาท/เที่ยว', (v, i) => v.avgMargin >= 0 ? COLORS[i % 10] : '#ef4444', v => 'ราคารับ ' + fmt(v.avgRecv) + ' THB/เที่ยว', true, true)}
  </div>`;
  const rows = vt.map(v => [v.type, v.trips, v.share, v.recv, v.margin, v.avgRecv, v.avgMargin, v.pct, v.loss]);
  const cols = ['ประเภทรถ', 'จำนวนเที่ยว', 'สัดส่วน %', 'ราคารับรวม', 'ส่วนต่างรวม', 'ราคารับเฉลี่ย/เที่ยว', 'ส่วนต่างเฉลี่ย/เที่ยว', 'กำไร %', 'จำนวนเที่ยวที่ขาดทุน'];
  h += `<div class="table-card"><div class="table-card-header"><h3>ประสิทธิภาพแยกตามประเภทรถ</h3></div><div id="t_vt"></div></div>`;
  setTimeout(() => mkTable('t_vt', cols, rows, { defaultSort: 1, defaultAsc: false }), 0);
  return h;
}

function buildDriver(d) {
  const dp = d.driverPerf;
  const compCnt = dp.filter(x => x.isCompany).length;
  let h = `<div class="kpi-grid">${kpi('พขร. ทั้งหมด', fmt(dp.length), 'accent')}${kpi('พขร. บริษัท', fmt(compCnt), 'green')}</div>`;
  const rows = dp.map((dr, i) => [i + 1, dr.name, dr.isCompany ? tag('รถบริษัท', 'blue') : tag('รถจ้างภายนอก', 'orange'), dr.trips, dr.margin, dr.pct, dr.loss, dr.mainRoute]);
  const cols = ['#', 'ชื่อพขร.', 'ประเภท', 'จำนวนเที่ยว', 'ส่วนต่างรวม', 'กำไร %', 'เที่ยวขาดทุน', 'เส้นทางหลัก'];
  h += `<div class="table-card"><div class="table-card-header"><h3>ดัชนีชี้วัดผลงานพนักงานขับรถรายบุคคล</h3></div><div id="t_dp"></div></div>`;
  setTimeout(() => mkTable('t_dp', cols, rows, { defaultSort: 3, defaultAsc: false }), 0);
  return h;
}


function buildComparison(d) {
  const mc = d.monthComparison;
  if (!mc || !mc.length) return '<div class="kpi"><div class="kpi-value red">No data — re-run analyze_all.ps1</div></div>';

  const SS = 'padding:6px 10px;background:var(--surface);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:13px';
  const customers = [...new Set(mc.map(r => r.customer))].sort();
  const vtypes = [...new Set(mc.map(r => r.vtype || '-'))].sort();

  // Filter bar
  let h = `<div style="background:var(--card);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:20px">
    <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:flex-end">
      <div><div style="font-size:11px;color:var(--muted);margin-bottom:4px">ลูกค้า</div>
        <select id="cmp_cust" onchange="cmpFilter()" style="${SS}">
          <option value="">ทั้งหมด</option>
          ${customers.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('')}
        </select></div>
      <div><div style="font-size:11px;color:var(--muted);margin-bottom:4px">ประเภทรถ</div>
        <select id="cmp_vtype" onchange="cmpFilter()" style="${SS}">
          <option value="">ทั้งหมด</option>
          ${vtypes.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('')}
        </select></div>
      <div style="flex:1;min-width:180px"><div style="font-size:11px;color:var(--muted);margin-bottom:4px">ค้นหา (Route / พขร. / ทะเบียน / ผู้รับโอน)</div>
        <input id="cmp_q" type="text" placeholder="พิมพ์เพื่อค้นหา..." oninput="cmpFilter()" style="${SS};width:100%;box-sizing:border-box"></div>
      <div><div style="font-size:11px;color:var(--muted);margin-bottom:4px">เรียงตาม</div>
        <select id="cmp_sort" onchange="cmpFilter()" style="${SS}">
          <option value="marginVariance">กำไร/ขาดทุน</option>
          <option value="recvVariance">ราคารับ</option>
          <option value="totalTrips">จำนวนเที่ยว</option>
        </select></div>
      <div style="align-self:center"><span id="cmp_count" style="font-size:12px;color:var(--muted);white-space:nowrap"></span></div>
    </div>
    <div style="margin-top:14px;padding-top:14px;border-top:1px solid var(--border)">
      <div style="font-size:11px;color:var(--muted);margin-bottom:8px">เลือกเดือนที่ต้องการเปรียบเทียบ (เลือกได้ตั้งแต่ 1 เดือนขึ้นไป)</div>
      <div style="display:flex;gap:16px;flex-wrap:wrap;align-items:center">
        ${MONTHS.map(m => `
          <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;color:var(--text);user-select:none">
            <input type="checkbox" id="cmp_m_${m}" onchange="cmpFilter()" checked
              style="width:15px;height:15px;accent-color:var(--accent);cursor:pointer">
            ${MTH[m]}
          </label>`).join('')}
        <div style="width:1px;height:20px;background:var(--border)"></div>
        <button onclick="cmpSelectMonths('all')" style="${btnSt};font-size:11px">ทั้งหมด</button>
        <button onclick="cmpSelectMonths('none')" style="${btnSt};font-size:11px">ล้าง</button>
        <span style="font-size:10px;color:var(--muted)">• เลือกเดือนเดียว = ดูข้อมูลเดือนนั้น</span>
      </div>
    </div>
    <div style="margin-top:14px;padding-top:14px;border-top:1px solid var(--border)">
      <div style="font-size:10px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px">★ Quick Filters</div>
      <div style="display:flex;gap:10px;flex-wrap:wrap">
        <label id="cmp_qf_oil_lbl" style="display:inline-flex;align-items:center;gap:8px;cursor:pointer;padding:7px 14px;background:rgba(245,158,11,.07);border:1px solid rgba(245,158,11,.2);border-radius:8px;font-size:12px;color:#fff;user-select:none;transition:all .2s" onmouseenter="this.style.background='rgba(245,158,11,.15)'" onmouseleave="var cb=document.getElementById('cmp_q_oil');this.style.background=cb.checked?'rgba(245,158,11,.22)':'rgba(245,158,11,.07)'">
          <input type="checkbox" id="cmp_q_oil" onchange="window._cmpPage=0;cmpFilter();var lbl=document.getElementById('cmp_qf_oil_lbl');lbl.style.background=this.checked?'rgba(245,158,11,.22)':'rgba(245,158,11,.07)';lbl.style.borderColor=this.checked?'#f59e0b':'rgba(245,158,11,.2)';" style="width:15px;height:15px;accent-color:#f59e0b;cursor:pointer;flex-shrink:0">
          <span> สำรองน้ำมัน &gt; 50% ของราคาจ่าย</span>
        </label>
      </div>
    </div>
  </div>
  <div id="cmp_pager"></div>
  <div id="cmp_table"></div>`;

  window._cmpData = mc;
  window._cmpPage = 0;
  window._cmpPerPage = 15;
  setTimeout(() => cmpFilter(), 0);
  return h;
}

function cmpSelectMonths(mode) {
  MONTHS.forEach(m => {
    const cb = document.getElementById('cmp_m_' + m);
    if (cb) cb.checked = (mode === 'all');
  });
  cmpFilter();
}

function cmpFilter() {
  const mc = window._cmpData;
  if (!mc) return;
  const cust = document.getElementById('cmp_cust')?.value || '';
  const vt = document.getElementById('cmp_vtype')?.value || '';
  const q = (document.getElementById('cmp_q')?.value || '').toLowerCase();
  const sortBy = document.getElementById('cmp_sort')?.value || 'marginVariance';
  const onlyHighOil = document.getElementById('cmp_q_oil')?.checked || false;

  // Determine selected months from checkboxes
  const selMonths = MONTHS.filter(m => document.getElementById('cmp_m_' + m)?.checked);
  const isSingle = selMonths.length === 1;

  let filtered = mc.filter(r => {
    if (cust && r.customer !== cust) return false;
    if (vt && (r.vtype || '-') !== vt) return false;
    if (q && ![(r.route || ''), (r.driver || ''), (r.contractor || ''), (r.plate || '')].some(s => s.toLowerCase().includes(q))) return false;
    // Must have data in ALL selected months
    if (selMonths.length === 0) return false;
    if (!selMonths.every(m => r.months && r.months[m])) return false;
    // Quick filter: any selected month where avgPay > 0 AND avgOil > avgPay * 0.5
    if (onlyHighOil) {
      const hasHighOil = selMonths.some(m => {
        const mv = r.months?.[m];
        return mv && mv.avgPay > 0 && mv.avgOil > mv.avgPay * 0.5;
      });
      if (!hasHighOil) return false;
    }
    return true;
  }).sort((a, b) => {
    if (isSingle) {
      const m = selMonths[0];
      return (b.months[m]?.avgMargin || 0) - (a.months[m]?.avgMargin || 0);
    }
    return (b[sortBy] || 0) - (a[sortBy] || 0);
  });

  const total = filtered.length;
  const perPage = window._cmpPerPage || 15;
  const maxPage = Math.max(0, Math.ceil(total / perPage) - 1);
  const page = Math.min(window._cmpPage || 0, maxPage);
  window._cmpPage = page;
  const pd = filtered.slice(page * perPage, (page + 1) * perPage);
  window._cmpPd = pd;
  const pages = Math.max(1, Math.ceil(total / perPage));

  const cntEl = document.getElementById('cmp_count');
  if (cntEl) cntEl.textContent = `พบ ${total} กลุ่ม | หน้า ${page + 1}/${pages}`;

  // Pager
  const pager = document.getElementById('cmp_pager');
  if (pager) {
    pager.innerHTML = `<div style="display:flex;gap:8px;align-items:center;margin-bottom:16px">
      <button onclick="window._cmpPage=Math.max(0,(window._cmpPage||0)-1);cmpFilter()" style="${btnSt}">‹ ก่อนหน้า</button>
      <button onclick="window._cmpPage=Math.min(${maxPage},(window._cmpPage||0)+1);cmpFilter()" style="${btnSt}">ถัดไป ›</button>
      <select onchange="window._cmpPerPage=+this.value;window._cmpPage=0;cmpFilter()" style="padding:4px 8px;background:var(--surface);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:12px">
        ${[10, 15, 25, 50].map(v => `<option value="${v}"${v === perPage ? ' selected' : ''}>${v} กลุ่ม/หน้า</option>`).join('')}
      </select>
    </div>`;
  }

  const tbl = document.getElementById('cmp_table');
  if (!tbl) return;
  if (!pd.length) { tbl.innerHTML = '<div style="padding:48px;text-align:center;color:var(--muted)">ไม่พบข้อมูลที่ตรงกัน</div>'; return; }

  const allM = ['January', 'February', 'March', 'April'];
  const MTH_TH = { January: 'มกราคม', February: 'กุมภาพันธ์', March: 'มีนาคม', April: 'เมษายน' };
  const MTH_SH = { January: 'ม.ค.', February: 'ก.พ.', March: 'มี.ค.', April: 'เม.ย.' };

  let html = '';
  pd.forEach((r, ri) => {
    const mKeys = selMonths.filter(m => r.months && r.months[m]);
    const margins = mKeys.map(m => r.months[m].avgMargin);
    const oils = mKeys.map(m => r.months[m].avgOil);
    const recvs = mKeys.map(m => r.months[m].avgRecv);
    const maxMg = Math.max(...margins), minMg = Math.min(...margins);
    const maxOil = Math.max(...oils), minOil = Math.min(...oils);
    const maxRecv = Math.max(...recvs), minRecv = Math.min(...recvs);
    const bestM = mKeys[margins.indexOf(maxMg)];
    const worstM = minMg !== maxMg ? mKeys[margins.indexOf(minMg)] : null;

    // — Card header (group info)
    html += `<div style="background:var(--card);border:1px solid var(--border);border-radius:12px;margin-bottom:20px;overflow:hidden">
      <div style="background:var(--surface);padding:14px 18px;border-bottom:1px solid var(--border)">
        <div style="display:flex;flex-wrap:wrap;gap:16px;align-items:flex-start">
          <div style="flex:2;min-width:160px">
            <div style="font-size:10px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em">ลูกค้า | ประเภทรถ</div>
            <div style="font-size:14px;font-weight:700;color:var(--accent)">${esc(r.customer)}</div>
            <div style="font-size:12px;color:var(--muted)">${esc(r.vtype || '-')}</div>
          </div>
          <div style="flex:3;min-width:180px">
            <div style="font-size:10px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em">เส้นทาง (Route)</div>
            <div style="font-size:13px;font-weight:600;word-break:break-all">${esc(r.route)}</div>
            <div style="font-size:11px;color:var(--muted)">${esc(r.routeDesc || '')}</div>
          </div>
          <div style="flex:2;min-width:140px">
            <div style="font-size:10px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em">พขร. | ทะเบียน</div>
            <div style="font-size:13px;font-weight:600">${esc(r.driver || '-')}</div>
            <div style="font-size:12px;color:var(--muted)">${esc(r.plate || '-')}</div>
          </div>
      
          <div style="min-width:110px;text-align:right">
            <div style="font-size:10px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em">${mKeys.length} เดือน</div>
          </div>
        </div>
      </div>

      <div style="overflow-x:auto;padding:0 4px 4px">
        <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:600px">
          <thead>
            <tr style="background:rgba(59,130,246,.07)">
              <th style="padding:10px 14px;text-align:left;font-size:11px;font-weight:600;color:var(--muted);white-space:nowrap">เดือน</th>
              <th style="padding:10px 14px;text-align:center;font-size:11px;font-weight:600;color:var(--muted)">จำนวนเที่ยว</th>
              <th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--muted)">ราคารับ / เที่ยว</th>
              <th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--muted)">ราคาจ่าย / เที่ยว</th>
              <th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--orange)">สำรองน้ำมัน / เที่ยว</th>
              <th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--muted)">ส่วนต่าง / เที่ยว</th>
              <th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--muted)">กำไร %</th>
              <th style="padding:10px 14px;text-align:right;font-size:11px;font-weight:600;color:var(--red)">เทียบกับเดือนดีสุด</th>
            </tr>
          </thead>
          <tbody>`;

    mKeys.forEach(m => {
      const mv = r.months[m];
      const isBest = m === bestM;
      const isWorst = m === worstM;
      const diffMg = mv.avgMargin - maxMg;
      const profitColor = mv.avgProfit >= 0 ? '#22c55e' : '#ef4444';
      const rowBg = isBest ? 'rgba(34,197,94,.10)' : isWorst ? 'rgba(239,68,68,.08)' : '';
      const badge = isBest
        ? `<span style="font-size:10px;background:rgba(34,197,94,.25);color:#22c55e;padding:1px 7px;border-radius:10px;font-weight:600">▲ ดีสุด</span>`
        : isWorst
          ? `<span style="font-size:10px;background:rgba(239,68,68,.2);color:#ef4444;padding:1px 7px;border-radius:10px;font-weight:600">▼ แย่สุด</span>` : '';
      const companyBadge = mv.avgPay === 0
        ? `<span style="font-size:10px;background:rgba(99,102,241,.2);color:#818cf8;padding:1px 7px;border-radius:10px;font-weight:600;white-space:nowrap">รถบริษัท</span>`
        : '';
      const oilHi = mv.avgOil === maxOil && maxOil !== minOil;
      const recvHi = mv.avgRecv === maxRecv && maxRecv !== minRecv;
      const recvLo = mv.avgRecv === minRecv && maxRecv !== minRecv;

      const diffStyle = diffMg < 0 ? 'color:#ef4444;font-weight:600' : 'color:var(--muted)';
      html += `<tr style="border-top:1px solid var(--border);background:${rowBg};cursor:pointer" title="คลิกเพื่อดูรายละเอียด" onclick="cmpDrilldown(${ri},'${m}')">
        <td style="padding:11px 14px;font-weight:700">
          <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">
            <span style="white-space:nowrap">${MTH_TH[m] || m}</span>
            ${badge}${companyBadge}
          </div>
        </td>
        <td style="padding:11px 14px;text-align:center;color:var(--muted)">${mv.trips}</td>
        <td style="padding:11px 14px;text-align:right;${recvHi ? 'color:#22c55e;font-weight:600' : recvLo ? 'color:#ef4444;font-weight:600' : ''}">${fmt(mv.avgRecv)}</td>
        <td style="padding:11px 14px;text-align:right">${mv.avgPay > 0 ? fmt(mv.avgPay) : '<span style="color:var(--muted)">-</span>'}</td>
        <td style="padding:11px 14px;text-align:right;${mv.avgOil > 0 && mv.avgPay > 0 && mv.avgOil > mv.avgPay * 0.5 ? 'color:var(--orange);font-weight:700' : oilHi ? 'color:var(--orange);font-weight:700' : ''}">${mv.avgOil > 0 ? fmt(mv.avgOil) : '-'}</td>
        <td style="padding:11px 14px;text-align:right;font-weight:600;color:${mv.avgMargin >= 0 ? '#22c55e' : '#ef4444'}">${fmt(mv.avgMargin)}</td>
        <td style="padding:11px 14px;text-align:right;color:${profitColor};font-weight:600">${mv.avgProfit}%</td>
        <td style="padding:11px 14px;text-align:right;${diffStyle}">${diffMg < 0 ? fmt(diffMg) : '—'}</td>
      </tr>`;

    });

    html += `</tbody></table></div></div>`;
  });

  tbl.innerHTML = html;
}

function cmpDrilldown(ri, month) {
  const r = window._cmpPd?.[ri];
  if (!r) return;
  const mv = r.months?.[month];
  if (!mv) return;
  const MTH_TH = { January: 'มกราคม', February: 'กุมภาพันธ์', March: 'มีนาคม', April: 'เมษายน' };
  document.getElementById('cmp_modal')?.remove();
  const modal = document.createElement('div');
  modal.id = 'cmp_modal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(6px)';
  modal.onclick = e => { if (e.target === modal) modal.remove(); };

  // Build trip rows — use real records if available, else generate from avg×trips
  let rows = [];
  if (mv.records && mv.records.length) {
    rows = mv.records.map(rec => ({
      recv: rec.recv || mv.avgRecv,
      pay: rec.pay || mv.avgPay,
      oil: rec.oil != null ? rec.oil : mv.avgOil,
      margin: rec.margin || mv.avgMargin,
      pct: rec.pct != null ? rec.pct : mv.avgProfit
    }));
  } else {
    for (let i = 0; i < mv.trips; i++) {
      rows.push({ recv: mv.avgRecv, pay: mv.avgPay, oil: mv.avgOil, margin: mv.avgMargin, pct: mv.avgProfit });
    }
  }

  const totRcv = rows.reduce((s, r) => s + r.recv, 0);
  const totPay = rows.reduce((s, r) => s + r.pay, 0);
  const totOil = rows.reduce((s, r) => s + r.oil, 0);
  const totMg = rows.reduce((s, r) => s + r.margin, 0);
  const totPct = totRcv > 0 ? (totMg / totRcv * 100).toFixed(2) : '0';
  const mgColor = totMg >= 0 ? '#22c55e' : '#ef4444';

  const TH = (label, align = 'right') =>
    `<th style="padding:10px 14px;text-align:${align};font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;white-space:nowrap;border-bottom:2px solid var(--border)">${label}</th>`;

  const tripRows = rows.map((row, i) => {
    const rColor = row.margin >= 0 ? '#22c55e' : '#ef4444';
    const rowBg = (i % 2 === 0) ? '' : 'background:rgba(255,255,255,.025)';
    return `<tr style="${rowBg};transition:background .15s" onmouseover="this.style.background='rgba(59,130,246,.12)'" onmouseout="this.style.background='${i % 2 === 0 ? '' : 'rgba(255,255,255,.025)'}'" >
      <td style="padding:9px 14px;text-align:center;color:var(--muted);font-size:12px;font-weight:600">${i + 1}</td>
      <td style="padding:9px 14px;text-align:right;font-weight:500">${fmt(row.recv)}</td>
      <td style="padding:9px 14px;text-align:right">${row.pay > 0 ? fmt(row.pay) : '<span style="color:var(--muted)">-</span>'}</td>
      <td style="padding:9px 14px;text-align:right;color:${row.oil > 0 ? 'var(--orange)' : 'var(--muted)'}">${row.oil > 0 ? fmt(row.oil) : '-'}</td>
      <td style="padding:9px 14px;text-align:right;font-weight:700;color:${rColor}">${fmt(row.margin)}</td>
      <td style="padding:9px 14px;text-align:right;font-weight:600;color:${rColor}">${row.pct}%</td>
    </tr>`;
  }).join('');

  const body = `
    <div style="overflow-x:auto;border-radius:10px;border:1px solid var(--border)">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:520px">
        <thead>
          <tr style="background:linear-gradient(135deg,rgba(59,130,246,.15),rgba(99,102,241,.1))">
            ${TH('#', 'center')}
            ${TH('ราคารับ / เที่ยว')}
            ${TH('ราคาจ่าย / เที่ยว')}
            ${TH('สำรองน้ำมัน / เที่ยว')}
            ${TH('ส่วนต่าง / เที่ยว')}
            ${TH('กำไร %')}
          </tr>
        </thead>
        <tbody>${tripRows}</tbody>
        <tfoot>
          <tr style="border-top:2px solid var(--accent);background:linear-gradient(135deg,rgba(59,130,246,.12),rgba(99,102,241,.08))">
            <td style="padding:11px 14px;font-weight:800;color:var(--accent);font-size:13px">รวม ${rows.length} เที่ยว</td>
            <td style="padding:11px 14px;text-align:right;font-weight:800;color:var(--text)">${fmt(Math.round(totRcv))}</td>
            <td style="padding:11px 14px;text-align:right;font-weight:700;color:var(--text)">${totPay > 0 ? fmt(Math.round(totPay)) : '-'}</td>
            <td style="padding:11px 14px;text-align:right;font-weight:700;color:${totOil > 0 ? 'var(--orange)' : 'var(--muted)'}">${totOil > 0 ? fmt(Math.round(totOil)) : '-'}</td>
            <td style="padding:11px 14px;text-align:right;font-weight:800;color:${mgColor}">${fmt(Math.round(totMg))}</td>
            <td style="padding:11px 14px;text-align:right;font-weight:800;color:${mgColor}">${totPct}%</td>
          </tr>
        </tfoot>
      </table>
    </div>`;

  modal.innerHTML = `
    <div style="background:var(--card);border:1px solid var(--border);border-radius:16px;max-width:860px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 24px 64px rgba(0,0,0,.6)">
      <!-- sticky header -->
      <div style="background:var(--surface);padding:18px 22px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;z-index:2;backdrop-filter:blur(10px)">
        <div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
            <span style="font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)">${MTH_TH[month] || month}</span>
            <span style="width:4px;height:4px;border-radius:50%;background:var(--accent);display:inline-block"></span>
            <span style="font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)">รายละเอียด ${rows.length} เที่ยว</span>
          </div>
          <div style="font-size:16px;font-weight:800;color:var(--accent)">${esc(r.customer)} <span style="font-weight:400;color:var(--muted)">·</span> ${esc(r.vtype || '-')}</div>
          <div style="font-size:12px;color:var(--text);margin-top:3px;opacity:.8">${esc(r.route)}</div>
          ${r.routeDesc ? `<div style="font-size:11px;color:var(--muted)">${esc(r.routeDesc)}</div>` : ''}
          <div style="font-size:11px;color:var(--muted);margin-top:2px">${r.driver ? esc(r.driver) : ''}${r.plate ? ' &nbsp;·&nbsp; <span style="font-family:monospace">' + esc(r.plate) + '</span>' : ''}</div>
        </div>
        <button onclick="document.getElementById('cmp_modal').remove()"
          style="flex-shrink:0;background:rgba(255,255,255,.06);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:5px 13px;cursor:pointer;font-size:18px;line-height:1;margin-left:16px;transition:background .2s"
          onmouseover="this.style.background='rgba(239,68,68,.2)'" onmouseout="this.style.background='rgba(255,255,255,.06)'">×</button>
      </div>
      <div style="padding:18px 22px">${body}</div>
    </div>`;
  document.body.appendChild(modal);
}



function initNav() {
  const nav = document.getElementById('navList');
  nav.innerHTML = PAGES.map((p, i) => `<div class="nav-item${i === 0 ? ' active' : ''}" data-idx="${i}"><span class="nav-num">${i + 1}</span>${p.title}</div>`).join('');
  nav.addEventListener('click', e => {
    const item = e.target.closest('.nav-item');
    if (!item) return;
    nav.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    item.classList.add('active');
    showPage(+item.dataset.idx);
  });
}

function showPage(idx) {
  currentPage = idx;
  const num = String(idx + 1).padStart(2, '0');
  document.getElementById('pageTitle').innerHTML = `
    <div style="display:flex;align-items:center;gap:20px;padding:4px 0;">
      <div style="width:56px;height:56px;border-radius:14px;background:linear-gradient(135deg, rgba(59,130,246,0.15) 0%, rgba(139,92,246,0.05) 100%);border:1px solid rgba(59,130,246,0.2);display:flex;flex-direction:column;align-items:center;justify-content:center;box-shadow:0 4px 15px rgba(0,0,0,0.05);position:relative;overflow:hidden;">
        <div style="position:absolute;top:0;left:0;width:100%;height:3px;background:linear-gradient(90deg, #3b82f6, #8b5cf6);"></div>
        <span style="font-size:10px;font-weight:800;color:#3b82f6;letter-spacing:1px;margin-bottom:2px;text-transform:uppercase;">Part</span>
        <span style="font-size:22px;font-weight:800;color:var(--text);line-height:1;">${num}</span>
      </div>
      <div style="display:flex;flex-direction:column;justify-content:center;">
        <div style="font-size:24px;font-weight:700;color:var(--text);letter-spacing:-0.3px;line-height:1.2;">${PAGES[idx].title}</div>
      </div>
    </div>`;
  document.getElementById('pageBadge').textContent = `${idx + 1} / ${PAGES.length}`;
  const c = document.getElementById('content');
  const builders = [buildTrend, buildRanking, buildCustomer, buildOwnOut, buildLoss, buildVehicle, buildDriver, buildComparison, buildFraudAudit];
  c.innerHTML = builders[idx](DATA);
  c.scrollTop = 0;
}

function init() {
  if (typeof DATA_JSON === 'undefined') {
    document.getElementById('content').innerHTML = '<div class="kpi"><div class="kpi-value red">data.js not found</div></div>';
    return;
  }
  DATA = DATA_JSON;
  initNav();
  showPage(0);
}
init();

function buildFraudAudit(d) {
  if (typeof BENCHMARK_DATA === 'undefined') return '<div class="kpi-grid"><h3>ไม่พบข้อมูล benchmark_data.js</h3></div>';

  const SS = 'padding:6px 10px;background:var(--surface);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:13px';
  const customers = [...new Set(BENCHMARK_DATA.map(r => r.customer))].sort();
  const vtypes = [...new Set(BENCHMARK_DATA.map(r => r.vtype || '-'))].sort();

  const mthsOrder = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const months = [...new Set(BENCHMARK_DATA.map(r => r.month))].sort((a, b) => mthsOrder.indexOf(a) - mthsOrder.indexOf(b));
  const MTH_TH = { January: 'มกราคม', February: 'กุมภาพันธ์', March: 'มีนาคม', April: 'เมษายน', May: 'พฤษภาคม', June: 'มิถุนายน', July: 'กรกฎาคม', August: 'สิงหาคม', September: 'กันยายน', October: 'ตุลาคม', November: 'พฤศจิกายน', December: 'ธันวาคม' };

  let h = `<div style="background:var(--card);border:1px solid var(--border);border-radius:12px;padding:16px;margin-bottom:20px">
    <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:flex-end">
      <div><div style="font-size:11px;color:var(--muted);margin-bottom:4px">เดือน</div>
        <select id="bm_month" onchange="window._bmPage=0;bmFilter()" style="${SS}">
          <option value="">ทั้งหมด</option>
          ${months.map(m => `<option value="${m}">${MTH_TH[m] || m}</option>`).join('')}
        </select></div>
      <div><div style="font-size:11px;color:var(--muted);margin-bottom:4px">ลูกค้า</div>
        <select id="bm_cust" onchange="window._bmPage=0;bmFilter()" style="${SS}">
          <option value="">ทั้งหมด</option>
          ${customers.map(c => `<option value="${esc(c)}">${esc(c)}</option>`).join('')}
        </select></div>
      <div><div style="font-size:11px;color:var(--muted);margin-bottom:4px">ประเภทรถ</div>
        <select id="bm_vtype" onchange="window._bmPage=0;bmFilter()" style="${SS}">
          <option value="">ทั้งหมด</option>
          ${vtypes.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('')}
        </select></div>
      <div style="flex:1;min-width:180px"><div style="font-size:11px;color:var(--muted);margin-bottom:4px">ค้นหาเส้นทาง (Route / พขร.)</div>
        <input id="bm_q" type="text" placeholder="พิมพ์เพื่อค้นหา..." oninput="window._bmPage=0;bmFilter()" style="${SS};width:100%;box-sizing:border-box"></div>
      <div style="align-self:center"><span id="bm_count" style="font-size:12px;color:var(--muted);white-space:nowrap"></span></div>
    </div>
  </div>
  <div id="bm_pager"></div>
  <div id="bm_table"></div>`;

  window._bmPage = 0;
  window._bmPerPage = 10;
  setTimeout(() => bmFilter(), 0);
  return h;
}

function bmFilter() {
  const data = BENCHMARK_DATA;
  if (!data) return;
  const mo = document.getElementById('bm_month')?.value || '';
  const cust = document.getElementById('bm_cust')?.value || '';
  const vt = document.getElementById('bm_vtype')?.value || '';
  const q = (document.getElementById('bm_q')?.value || '').toLowerCase();

  let filtered = data.filter(r => {
    if (mo && r.month !== mo) return false;
    if (cust && r.customer !== cust) return false;
    if (vt && (r.vtype || '-') !== vt) return false;
    if (q) {
      const matchR = [(r.route || ''), (r.routeDesc || '')].some(s => s.toLowerCase().includes(q));
      const matchD = r.drivers.some(d => [(d.driver || ''), (d.plate || '')].some(s => s.toLowerCase().includes(q)));
      if (!matchR && !matchD) return false;
    }
    return true;
  });

  const total = filtered.length;
  const perPage = window._bmPerPage || 10;
  const maxPage = Math.max(0, Math.ceil(total / perPage) - 1);
  const page = Math.min(window._bmPage || 0, maxPage);
  window._bmPage = page;
  const pd = filtered.slice(page * perPage, (page + 1) * perPage);
  window._bmPd = pd;
  const pages = Math.max(1, Math.ceil(total / perPage));

  const cntEl = document.getElementById('bm_count');
  if (cntEl) cntEl.textContent = `พบ ${total} เส้นทาง | หน้า ${page + 1}/${pages}`;

  const btnSt = 'padding:4px 10px;background:var(--surface);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:12px;cursor:pointer';
  const pager = document.getElementById('bm_pager');
  if (pager) {
    pager.innerHTML = `<div style="display:flex;gap:8px;align-items:center;margin-bottom:16px">
      <button onclick="window._bmPage=Math.max(0,(window._bmPage||0)-1);bmFilter()" style="${btnSt}">‹ ก่อนหน้า</button>
      <button onclick="window._bmPage=Math.min(${maxPage},(window._bmPage||0)+1);bmFilter()" style="${btnSt}">ถัดไป ›</button>
      <select onchange="window._bmPerPage=+this.value;window._bmPage=0;bmFilter()" style="padding:4px 8px;background:var(--surface);border:1px solid var(--border);border-radius:6px;color:var(--text);font-size:12px">
        ${[10, 20, 50].map(v => `<option value="${v}"${v === perPage ? ' selected' : ''}>${v} เส้นทาง/หน้า</option>`).join('')}
      </select>
    </div>`;
  }

  const tbl = document.getElementById('bm_table');
  if (!tbl) return;
  if (!pd.length) { tbl.innerHTML = '<div style="padding:48px;text-align:center;color:var(--muted)">ไม่พบข้อมูลที่ตรงกัน</div>'; return; }

  const MTH_TH = { January: 'มกราคม', February: 'กุมภาพันธ์', March: 'มีนาคม', April: 'เมษายน' };

  let html = '';
  pd.forEach((r, ri) => {
    html += `<div style="background:var(--card);border:1px solid var(--border);border-radius:12px;margin-bottom:24px;overflow:hidden">
      <div style="background:var(--surface);padding:14px 18px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
        <div>
          <div style="font-size:10px;text-transform:uppercase;color:var(--muted);letter-spacing:.05em">ลูกค้า | ประเภทรถ</div>
          <div style="font-size:14px;font-weight:700;color:var(--accent)">${esc(r.customer)} <span style="font-weight:400;color:var(--muted)">·</span> ${esc(r.vtype || '-')}</div>
          <div style="font-size:13px;font-weight:600;margin-top:4px">${esc(r.route)}</div>
          <div style="font-size:11px;color:var(--muted)">${esc(r.routeDesc || '')}</div>
        </div>
        <div style="text-align:right">
          ${r.drivers.length === 1 ? `<div style="font-size:11px;font-weight:600;color:var(--orange);margin-bottom:6px"></div>` : ''}
          <div style="font-size:12px;font-weight:700;color:var(--accent);margin-bottom:4px;padding:4px 8px;background:rgba(59,130,246,0.1);border-radius:4px;display:inline-block;">เดือน ${MTH_TH[r.month] || r.month}</div>
          <div style="font-size:11px;color:var(--muted)">ค่าเฉลี่ยของเส้นทางนี้ (รวม ${r.totalTrips} เที่ยว)</div>
          <div style="display:flex;gap:12px;margin-top:4px;font-size:12px">
            <div><span style="color:var(--muted)">ราคาจ่าย:</span> <b style="color:var(--text)">${fmt(r.avgPay)}</b></div>
            <div><span style="color:var(--muted)">สำรองน้ำมัน:</span> <b style="color:var(--text)">${fmt(r.avgOil)}</b></div>
          </div>
        </div>
      </div>
      
      <div style="overflow-x:auto">
        <table style="width:100%;border-collapse:collapse;font-size:12px;min-width:650px">
          <thead>
            <tr style="background:rgba(255,255,255,.03)">
              <th style="padding:10px 18px;text-align:left;color:var(--muted);font-weight:600">พขร. | ทะเบียน</th>
              <th style="padding:10px 12px;text-align:center;color:var(--muted);font-weight:600">จำนวนเที่ยว</th>
              <th style="padding:10px 12px;text-align:right;color:var(--muted);font-weight:600">เฉลี่ยราคารับ/เที่ยว</th>
              <th style="padding:10px 12px;text-align:right;color:var(--muted);font-weight:600">ราคาจ่ายสูงสุด/เที่ยว</th>
              <th style="padding:10px 12px;text-align:right;color:var(--muted);font-weight:600">เฉลี่ยราคาจ่าย/เที่ยว</th>
              <th style="padding:10px 12px;text-align:right;color:var(--muted);font-weight:600">ราคาจ่ายต่ำสุด/เที่ยว</th>
              <th style="padding:10px 12px;text-align:right;color:var(--muted);font-weight:600">เฉลี่ยสำรองน้ำมัน/เที่ยว</th>
              <th style="padding:10px 18px;text-align:right;color:var(--muted);font-weight:600">เฉลี่ยส่วนต่าง/เที่ยว</th>
            </tr>
          </thead>
          <tbody>`;

    r.drivers.sort((a, b) => {
      const getScore = (d) => {
        let score = 0;
        if (d.avgPay > 0 && d.avgOil > d.avgPay * 0.5) score += 100;
        if (d.avgPay > (r.avgPay * 1.05)) score += 50;
        if (d.avgOil > (r.avgOil * 1.10)) score += 50;
        if (d.avgMargin < 0) score += 10;
        return score;
      };
      return getScore(b) - getScore(a);
    }).forEach((d, di) => {
      const isPayHigh = d.avgPay > (r.avgPay * 1.05);
      const isOilHigh = d.avgOil > (r.avgOil * 1.10);
      const isOilOver50 = (d.avgPay > 0 && d.avgOil > d.avgPay * 0.5);
      const isMarginNegative = d.avgMargin < 0;

      const payStyle = isPayHigh ? 'color:var(--red);font-weight:700' : '';
      let oilStyle = isOilHigh ? 'color:var(--orange);font-weight:700' : '';
      if (isOilOver50) oilStyle = 'color:var(--red);font-weight:700'; // Override with stricter style
      const mgStyle = !isMarginNegative ? 'color:var(--green)' : 'color:var(--red);font-weight:700';

      let oilContent = fmt(d.avgOil);
      if (isOilOver50) {
        oilContent += `<br><div style="font-size:10px;background:rgba(239,68,68,0.1);color:var(--red);padding:3px 6px;border-radius:4px;margin-top:4px;display:inline-block;font-weight:600;border:1px solid rgba(239,68,68,0.2);white-space:nowrap">❗️ เกินเกณฑ์ 50% (${((d.avgOil / d.avgPay) * 100).toFixed(0)}%)</div>`;
      } else if (isOilHigh) {
        oilContent += `<br><div style="font-size:10px;background:rgba(249,115,22,0.1);color:var(--orange);padding:3px 6px;border-radius:4px;margin-top:4px;display:inline-block;font-weight:600;border:1px solid rgba(249,115,22,0.2);white-space:nowrap">สูงกว่าค่าเฉลี่ยเส้นทาง</div>`;
      }

      const dotColor = (isPayHigh || isOilHigh || isOilOver50 || isMarginNegative) ? 'var(--red)' : 'transparent';

      html += `<tr style="border-top:1px solid rgba(255,255,255,.05);cursor:pointer;transition:background .2s" onclick="bmDrilldown(${ri},${di})" onmouseover="this.style.background='rgba(59,130,246,.12)'" onmouseout="this.style.background=''">
        <td style="padding:10px 18px;font-weight:600">
          <div style="display:flex;align-items:center;gap:6px">
            <span style="width:6px;height:6px;border-radius:50%;background:${dotColor}"></span>
            <div>
              <div style="color:var(--text)">${esc(d.driver || '-')}</div>
              <div style="font-size:10px;color:var(--muted);font-weight:400;margin-top:2px">${esc(d.plate || '-')}</div>
            </div>
          </div>
        </td>
        <td style="padding:10px 12px;text-align:center">${d.totalTrips}</td>
        <td style="padding:10px 12px;text-align:right">${fmt(d.avgPay + d.avgMargin)}</td>
        <td style="padding:10px 12px;text-align:right;color:var(--red)">${d.maxPay >= 0 ? fmt(d.maxPay) : '-'}</td>
        <td style="padding:10px 12px;text-align:right;${payStyle}">${fmt(d.avgPay)}</td>
        <td style="padding:10px 12px;text-align:right;color:var(--green)">${d.minPay !== 999999999 ? fmt(d.minPay) : '-'}</td>
        <td style="padding:10px 12px;text-align:right;${oilStyle}">${oilContent}</td>
        <td style="padding:10px 18px;text-align:right;${mgStyle}">${fmt(d.avgMargin)}</td>
      </tr>`;
    });

    html += `</tbody></table></div></div>`;
  });

  tbl.innerHTML = html;
}

function bmDrilldown(ri, di) {
  const r = window._bmPd?.[ri];
  const d = r?.drivers?.[di];
  if (!d || !d.records) return;

  const MTH_TH = { January: 'มกราคม', February: 'กุมภาพันธ์', March: 'มีนาคม', April: 'เมษายน' };
  document.getElementById('bm_modal')?.remove();
  const modal = document.createElement('div');
  modal.id = 'bm_modal';
  modal.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;backdrop-filter:blur(6px)';
  modal.onclick = e => { if (e.target === modal) modal.remove(); };

  const TH = (label, align = 'right') => `<th style="padding:10px 14px;text-align:${align};font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;border-bottom:2px solid var(--border)">${label}</th>`;

  const tag = (text, color) => `<span style="font-size:10px;padding:2px 6px;border-radius:4px;background:var(--${color});color:#fff;margin-right:4px;white-space:nowrap">${text}</span>`;

  const tripRows = d.records.map((row, i) => {
    const rColor = row.margin >= 0 ? '#22c55e' : '#ef4444';
    const rowBg = (i % 2 === 0) ? '' : 'background:rgba(255,255,255,.025)';

    const isPayHigh = row.pay > (r.avgPay * 1.05);
    const isOilHigh = row.oil > (r.avgOil * 1.10);
    const pStyle = isPayHigh ? 'color:var(--red);font-weight:700' : '';
    const oStyle = isOilHigh ? 'color:var(--orange);font-weight:700' : '';

    let tags = [];
    if (isPayHigh) tags.push(tag('ราคาจ่ายแพงกว่า พขร.ท่านอื่น', 'red'));
    if (isOilHigh) tags.push(tag('น้ำมันแพงกว่า พขร.ท่านอื่น', 'orange'));
    if (row.margin < 0) tags.push(tag('ขาดทุน', 'red'));
    if (row.oil > row.pay * 0.5 && row.pay > 0) tags.push(tag('สำรองน้ำมัน >50%', 'orange'));

    return `<tr style="${rowBg};transition:background .15s" onmouseover="this.style.background='rgba(59,130,246,.12)'" onmouseout="this.style.background='${i % 2 === 0 ? '' : 'rgba(255,255,255,.025)'}'" >
      <td style="padding:9px 14px;text-align:center;color:var(--muted);font-size:12px;font-weight:600">${row.date}</td>
      <td style="padding:9px 14px;text-align:right;font-weight:500">${fmt(row.recv)}</td>
      <td style="padding:9px 14px;text-align:right;${pStyle}">${row.pay > 0 ? fmt(row.pay) : '<span style="color:var(--muted)">-</span>'}</td>
      <td style="padding:9px 14px;text-align:right;${oStyle}">${row.oil > 0 ? fmt(row.oil) : '-'}</td>
      <td style="padding:9px 14px;text-align:right;font-weight:700;color:${rColor}">${fmt(row.margin)}</td>
      <td style="padding:9px 14px;text-align:right">${esc(row.payee || '-')}</td>
      <td style="padding:9px 14px;text-align:left">${tags.join(' ')}</td>
    </tr>`;
  }).join('');

  const body = `
    <div style="overflow-x:auto;border-radius:10px;border:1px solid var(--border)">
      <table style="width:100%;border-collapse:collapse;font-size:13px;min-width:750px">
        <thead>
          <tr style="background:linear-gradient(135deg,rgba(59,130,246,.1),rgba(99,102,241,.05))">
            ${TH('วันที่', 'center')}
            ${TH('ราคารับ')}
            ${TH('ราคาจ่าย')}
            ${TH('สำรองน้ำมัน')}
            ${TH('ส่วนต่าง')}
            ${TH('ผู้รับโอน')}
            ${TH('สาเหตุความผิดปกติ', 'left')}
          </tr>
        </thead>
        <tbody>${tripRows}</tbody>
      </table>
    </div>`;

  modal.innerHTML = `
    <div style="background:var(--card);border:1px solid var(--border);border-radius:16px;max-width:960px;width:100%;max-height:90vh;overflow:auto;box-shadow:0 24px 64px rgba(0,0,0,.6)">
      <div style="background:var(--surface);padding:18px 22px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;z-index:2;backdrop-filter:blur(10px)">
        <div>
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px">
            <span style="font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent)">รายละเอียดการวิ่งเดือน ${MTH_TH[r.month] || r.month}</span>
            <span style="width:4px;height:4px;border-radius:50%;background:var(--accent);display:inline-block"></span>
            <span style="font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted)">รวม ${d.totalTrips} เที่ยว</span>
          </div>
          <div style="font-size:16px;font-weight:800;color:var(--text)">${esc(d.driver)} <span style="font-weight:400;color:var(--muted)">·</span> ${esc(d.plate || '-')}</div>
          <div style="font-size:12px;color:var(--muted);margin-top:3px">${esc(r.customer)} · ${esc(r.route)}</div>
        </div>
        <button onclick="document.getElementById('bm_modal').remove()"
          style="flex-shrink:0;background:rgba(255,255,255,.06);border:1px solid var(--border);border-radius:8px;color:var(--text);padding:5px 13px;cursor:pointer;font-size:18px;line-height:1;margin-left:16px;transition:background .2s"
          onmouseover="this.style.background='rgba(239,68,68,.2)'" onmouseout="this.style.background='rgba(255,255,255,.06)'">×</button>
      </div>
      <div style="padding:18px 22px">${body}</div>
    </div>`;
  document.body.appendChild(modal);
}