const $ = (id) => document.getElementById(id);

let rawBank = [];
let rawBook = [];
let bank = [];
let book = [];
let matched = [];
let review = [];
let pending = [];
let currentTab = 'matched';
let chartInstances = {};

const STOP = new Set([
  'PAGO', 'PAGOS', 'A', 'PROVE', 'PROVEEDOR', 'NOMIN', 'NOMINA',
  'TRANSFERENCIA', 'TRANSF', 'ABONO', 'CTA', 'SUC', 'VIRTUAL',
  'IMPTO', 'GOBIERNO', 'COBRO', 'RECAUDO', 'SAS', 'LTDA', 'SA',
  'S', 'DE', 'DEL', 'LA', 'LAS', 'LOS', 'EL', 'Y', 'EN', 'POR',
  'PARA', 'BANCO', 'CUENTA', 'SERVICIO', 'SERVICIOS'
]);

const money = (n) =>
  new Intl.NumberFormat('es-CO', {
    style: 'currency',
    currency: 'COP',
    maximumFractionDigits: 2
  }).format((n || 0) / 100);

const cents = (v) => Math.round((Number(v) || 0) * 100);

function norm(s) {
  return String(s ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/Ñ/g, 'N')
    .replace(/ñ/g, 'n')
    .toUpperCase()
    .replace(/[^A-Z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseDate(v) {
  if (v instanceof Date && !isNaN(v)) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (typeof v === 'number' && window.XLSX) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(d.y, d.m - 1, d.d);
  }
  if (v == null || v === '') return null;

  const s = String(v).trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  }

  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let y = +m[3];
    if (y < 100) y += 2000;
    return new Date(y, +m[2] - 1, +m[1]);
  }

  const d = new Date(s);
  return isNaN(d) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function dateText(d) {
  if (!d) return '';
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

function days(a, b) {
  return Math.round((a - b) / 86400000);
}

function monthSame(a, b) {
  return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth();
}

function amount(v) {
  if (typeof v === 'number') return cents(v);
  let s = String(v ?? '').trim();
  if (!s) return 0;
  s = s.replace(/\s/g, '').replace(/\$/g, '');
  if (s.includes(',') && s.includes('.')) {
    s = s.lastIndexOf(',') > s.lastIndexOf('.')
      ? s.replace(/\./g, '').replace(',', '.')
      : s.replace(/,/g, '');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  return cents(Number(s) || 0);
}

function meaningful(text, companyTokens = []) {
  return norm(text)
    .split(' ')
    .filter((t) => t.length > 1 && !STOP.has(t) && !companyTokens.includes(t));
}

function detectCompanyTokens() {
  const co = norm($('company').value);
  return co.split(' ').filter((t) => t.length > 2 && !STOP.has(t));
}

function nameMatches(a, b, companyTokens) {
  const A = meaningful(a, companyTokens);
  const B = meaningful(b, companyTokens);
  if (!A.length || !B.length) return false;

  const short = A.length <= B.length ? A : B;
  const long = A.length <= B.length ? B : A;

  const ok = short.every((x) => long.some((y) => x === y || x.startsWith(y) || y.startsWith(x)));
  if (!ok) return false;
  if (short.length === 1 && short[0].length < 4) return false;
  if (short.length === 1 && (STOP.has(short[0]) || companyTokens.includes(short[0]))) return false;
  return true;
}

function readCSV(text) {
  const firstLine = text.split(/\r?\n/)[0] || '';
  const delim = (firstLine.match(/;/g) || []).length > (firstLine.match(/,/g) || []).length ? ';' : ',';
  const rows = [];
  let row = [];
  let cell = '';
  let q = false;

  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"' && q && text[i + 1] === '"') {
      cell += '"';
      i++;
    } else if (c === '"') {
      q = !q;
    } else if (c === delim && !q) {
      row.push(cell);
      cell = '';
    } else if ((c === '\n' || c === '\r') && !q) {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(cell);
      if (row.some((x) => String(x).trim())) rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += c;
    }
  }
  row.push(cell);
  if (row.some((x) => String(x).trim())) rows.push(row);
  return rows;
}

async function workbookRows(file) {
  const buf = await file.arrayBuffer();
  if (/\.csv$|\.txt$/i.test(file.name)) {
    return readCSV(new TextDecoder().decode(buf));
  }
  if (!window.XLSX) {
    throw Error('No se pudo cargar la librería Excel. Conecta a internet y vuelve a abrir la aplicación.');
  }
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: null,
    raw: true
  });
}

function findHeader(headers, options) {
  const hs = headers.map((h) => norm(h));
  for (const o of options) {
    const i = hs.findIndex((h) => h === norm(o));
    if (i >= 0) return i;
  }
  return -1;
}

function bankRowsFrom(rows) {
  return rows
    .map((r, i) => {
      if (!r || r.length < 8) return null;
      const d = parseDate(r[3]);
      const v = amount(r[5]);
      const desc = String(r[7] ?? '').trim();
      if (!d || (!desc && !v)) return null;
      return {
        id: 'B' + i,
        side: 'Banco',
        date: d,
        desc,
        beneficiary: '',
        value: v,
        used: false,
        sourceRow: i + 1
      };
    })
    .filter(Boolean);
}

function bookRowsFrom(rows) {
  if (!rows.length) throw Error('El libro auxiliar está vacío.');
  const hi = rows.findIndex((r) =>
    r.some((c) => ['FECHA', 'DETALLE', 'NOMBRE BENEFICIARIO', 'DEBITO', 'CREDITO'].includes(norm(c)))
  );
  if (hi < 0) {
    throw Error('No encuentro encabezados FECHA, DETALLE, NOMBRE BENEFICIARIO, DÉBITO y CRÉDITO en el libro auxiliar.');
  }

  const h = rows[hi];
  const fi = findHeader(h, ['FECHA']);
  const di = findHeader(h, ['DETALLE', 'CONCEPTO']);
  const ni = findHeader(h, ['NOMBRE BENEFICIARIO', 'BENEFICIARIO', 'TERCERO']);
  const de = findHeader(h, ['DEBITO', 'DÉBITO']);
  const cr = findHeader(h, ['CREDITO', 'CRÉDITO']);

  if (fi < 0 || de < 0 || cr < 0) {
    throw Error('Faltan columnas requeridas en el libro auxiliar.');
  }

  return rows
    .slice(hi + 1)
    .map((r, i) => {
      const d = parseDate(r[fi]);
      const desc = String(di >= 0 ? (r[di] ?? '') : '');
      const ben = String(ni >= 0 ? (r[ni] ?? '') : '');
      const v = amount(r[de]) - amount(r[cr]);
      if (!d || (!desc && !ben && !v)) return null;
      return {
        id: 'C' + (hi + i + 1),
        side: 'Contabilidad',
        date: d,
        desc,
        beneficiary: ben,
        value: v,
        used: false,
        sourceRow: hi + i + 2
      };
    })
    .filter(Boolean);
}

function filterByDateRange(items) {
  const from = parseDate($('fromDate').value);
  const to = parseDate($('toDate').value);
  return items
    .filter((item) => {
      if (from && item.date < from) return false;
      if (to && item.date > to) return false;
      return true;
    })
    .map((item) => ({ ...item, used: false }));
}

function addMatch(b, c, rule, detail = '') {
  b.used = true;
  c.used = true;
  matched.push({
    bank: b,
    book: c,
    rule,
    detail,
    diff: Math.abs(b.value - c.value)
  });
}

function exactPass() {
  for (const b of bank.filter((x) => !x.used)) {
    const c = book.find((x) => !x.used && x.value === b.value && days(b.date, x.date) === 0);
    if (c) addMatch(b, c, 'Fecha y valor exactos');
  }
}

function nearPass() {
  for (const b of bank.filter((x) => !x.used)) {
    const cs = book
      .filter((c) => !c.used && c.value === b.value && Math.abs(days(b.date, c.date)) <= 3)
      .sort((a, c) => Math.abs(days(b.date, a.date)) - Math.abs(days(b.date, c.date)));
    if (cs[0]) addMatch(b, cs[0], 'Valor exacto · fecha ±3 días');
  }
}

function namePass(companyTokens) {
  for (const b of bank.filter((x) => !x.used)) {
    const cs = book
      .filter(
        (c) =>
          !c.used &&
          c.value === b.value &&
          Math.abs(days(b.date, c.date)) <= 15 &&
          nameMatches(b.desc, c.beneficiary + ' ' + c.desc, companyTokens)
      )
      .sort((a, c) => Math.abs(days(b.date, a.date)) - Math.abs(days(b.date, c.date)));
    if (cs[0]) addMatch(b, cs[0], 'Valor exacto · nombre ±15 días');
  }
}

function groupKeyFor(b, c, companyTokens) {
  if (c.beneficiary && norm(c.beneficiary).length > 2) return 'BEN:' + norm(c.beneficiary);
  if (nameMatches(b.desc, c.beneficiary + ' ' + c.desc, companyTokens)) {
    return 'NOM:' + meaningful(c.beneficiary + ' ' + c.desc, companyTokens).sort().join('|');
  }
  const shared = meaningful(b.desc, companyTokens).filter((t) =>
    meaningful(c.desc + ' ' + c.beneficiary, companyTokens).some(
      (u) => t === u || t.startsWith(u) || u.startsWith(t)
    )
  );
  if (shared.length) return 'TXT:' + shared.sort().join('|');
  return 'ANY';
}

function windowOk(group, target) {
  return group.every((x) => Math.abs(days(x.date, target.date)) <= 3 || monthSame(x.date, target.date));
}

function findSubset(items, target, limit = 12) {
  const xs = items.slice(0, limit);
  let found = null;

  function dfs(i, sum, chosen) {
    if (sum === target && chosen.length >= 2) {
      found = chosen.slice();
      return true;
    }
    if (sum >= target || i >= xs.length || found) return false;
    for (let j = i; j < xs.length; j++) {
      if (sum + Math.abs(xs[j].value) > target) continue;
      if (dfs(j + 1, sum + Math.abs(xs[j].value), chosen.concat(xs[j]))) return true;
    }
    return false;
  }

  dfs(0, 0, []);
  return found;
}

function trySumDirection(singles, targets, companyTokens, direction) {
  for (const target of targets.filter((x) => !x.used)) {
    const candidates = singles.filter(
      (x) =>
        !x.used &&
        Math.sign(x.value) === Math.sign(target.value) &&
        Math.abs(x.value) < Math.abs(target.value) &&
        windowOk([x], target)
    );
    if (candidates.length < 2) continue;

    const groups = new Map();
    for (const x of candidates) {
      const key =
        direction === 'bankToBook'
          ? groupKeyFor(x, target, companyTokens)
          : groupKeyFor(target, x, companyTokens);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(x);
    }

    let chosen = null;
    let usedKey = '';
    const keys = [...groups.keys()]
      .filter((k) => k !== 'ANY')
      .sort((a, b) => {
        const rank = (k) => (k.startsWith('BEN:') ? 0 : k.startsWith('NOM:') ? 1 : 2);
        return rank(a) - rank(b);
      });
    keys.push('ANY');

    for (const key of keys) {
      const group = groups.get(key) || [];
      if (group.length < 2) continue;

      const byDay = new Map();
      for (const x of group) {
        const dk = dateText(x.date);
        if (!byDay.has(dk)) byDay.set(dk, []);
        byDay.get(dk).push(x);
      }

      const dayItems = [...byDay.entries()]
        .map(([day, arr]) => ({
          day,
          date: arr[0].date,
          value: arr.reduce((s, x) => s + Math.abs(x.value), 0),
          items: arr
        }))
        .filter((x) => windowOk(x.items, target));

      const daySubset = findSubset(
        dayItems.map((x) => ({ ...x, value: Math.sign(target.value) * x.value })),
        Math.abs(target.value),
        8
      );

      if (daySubset) {
        chosen = daySubset.flatMap((x) => x.items);
        usedKey = key;
        break;
      }

      const subset = findSubset(
        group
          .filter((x) => Math.abs(days(x.date, target.date)) <= 3 || monthSame(x.date, target.date))
          .map((x) => ({ ...x, value: Math.sign(target.value) * Math.abs(x.value) })),
        Math.abs(target.value),
        12
      );

      if (subset) {
        chosen = subset;
        usedKey = key;
        break;
      }
    }

    if (
      chosen &&
      chosen.length > 1 &&
      chosen.every((x) => !x.used) &&
      chosen.reduce((s, x) => s + Math.abs(x.value), 0) === Math.abs(target.value)
    ) {
      chosen.forEach((x) => (x.used = true));
      target.used = true;
      const b = direction === 'bankToBook' ? chosen : [target];
      const c = direction === 'bankToBook' ? [target] : chosen;
      matched.push({
        bank: b,
        book: c,
        rule: 'Cruce por suma',
        detail: `Grupo: ${usedKey === 'ANY' ? 'movimientos cercanos' : usedKey.replace(/^(BEN|NOM|TXT):/, '')}`,
        diff: 0
      });
    }
  }
}

function sumPass(companyTokens) {
  trySumDirection(
    bank.filter((x) => !x.used),
    book.filter((x) => !x.used),
    companyTokens,
    'bankToBook'
  );
  trySumDirection(
    bank.filter((x) => !x.used),
    book.filter((x) => !x.used),
    companyTokens,
    'bookToBank'
  );
}

function reviewPass(companyTokens) {
  for (const b of bank.filter((x) => !x.used)) {
    const cs = book
      .filter(
        (c) =>
          !c.used &&
          Math.abs(b.value - c.value) <= 10000000 &&
          days(b.date, c.date) === 0 &&
          nameMatches(b.desc, c.beneficiary + ' ' + c.desc, companyTokens)
      )
      .sort((a, c) => Math.abs(b.value - a.value) - Math.abs(b.value - c.value));

    if (cs.length) {
      const c = cs[0];
      b.used = true;
      c.used = true;
      review.push({
        bank: b,
        book: c,
        diff: Math.abs(b.value - c.value),
        rule: 'Misma fecha y nombre; valor distinto'
      });
    }
  }
}

function runReconciliation() {
  matched = [];
  review = [];
  bank = filterByDateRange(rawBank);
  book = filterByDateRange(rawBook);

  const ct = detectCompanyTokens();
  exactPass();
  nearPass();
  namePass(ct);
  sumPass(ct);
  reviewPass(ct);

  pending = [
    ...bank.filter((x) => !x.used).map((x) => ({ side: 'Banco', item: x })),
    ...book.filter((x) => !x.used).map((x) => ({ side: 'Contabilidad', item: x }))
  ];

  render();
}

function esc(s) {
  return String(s ?? '').replace(
    /[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])
  );
}

function render() {
  const total = bank.length + book.length;
  const matchedItems = matched.reduce(
    (n, m) => n + (Array.isArray(m.bank) ? m.bank.length : 1) + (Array.isArray(m.book) ? m.book.length : 1),
    0
  );
  const pct = total ? Math.round((matchedItems / total) * 100) : 0;

  $('bankCount').textContent = bank.length;
  $('bookCount').textContent = book.length;
  $('matchedCount').textContent = matchedItems;
  $('pendingCount').textContent = pending.length;
  $('percent').textContent = pct + '%';
  $('progressBar').style.width = pct + '%';
  $('nMatched').textContent = matched.length;
  $('nReview').textContent = review.length;
  $('nPending').textContent = pending.length;

  const hasDates = $('fromDate').value || $('toDate').value;
  $('periodTag').textContent = hasDates
    ? `${$('fromDate').value || 'Inicio'} — ${$('toDate').value || 'Fin'}`
    : total
    ? 'Periodo cargado'
    : 'Pendiente';

  showTable();
  renderCharts(matchedItems);
}

function renderCharts(matchedItems) {
  const ids = ['statusChart', 'rulesChart', 'pendingChart'];
  if (!window.Chart) {
    ids.forEach((id) => {
      const c = $(id);
      if (c) c.parentElement.innerHTML = '<div class="chart-empty">Conecta a internet para cargar las gráficas.</div>';
    });
    return;
  }

  const make = (id, config) => {
    const canvas = $(id);
    if (!canvas) return;
    if (chartInstances[id]) chartInstances[id].destroy();
    chartInstances[id] = new Chart(canvas, config);
  };

  const pendingBank = pending.filter((x) => x.side === 'Banco').length;
  const pendingBook = pending.filter((x) => x.side === 'Contabilidad').length;

  make('statusChart', {
    type: 'doughnut',
    data: {
      labels: ['Conciliados', 'Pendientes', 'Por revisar'],
      datasets: [
        {
          data: [matchedItems, pending.length, review.length * 2],
          backgroundColor: ['#2475e8', '#f59e0b', '#e05b68'],
          borderColor: '#ffffff',
          borderWidth: 4,
          hoverOffset: 7
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '68%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: { usePointStyle: true, pointStyle: 'circle', padding: 18, font: { size: 11 } }
        },
        tooltip: { callbacks: { label: (c) => `${c.label}: ${c.raw} movimientos` } }
      }
    }
  });

  const ruleCounts = {};
  matched.forEach((m) => {
    const rule = m.rule || 'Otro';
    ruleCounts[rule] =
      (ruleCounts[rule] || 0) +
      (Array.isArray(m.bank) ? m.bank.length : 1) +
      (Array.isArray(m.book) ? m.book.length : 1);
  });

  const ruleLabels = Object.keys(ruleCounts);
  const ruleValues = ruleLabels.map((k) => ruleCounts[k]);

  make('rulesChart', {
    type: 'bar',
    data: {
      labels: ruleLabels,
      datasets: [
        {
          label: 'Movimientos conciliados',
          data: ruleValues,
          backgroundColor: ['#2868d7', '#1e9a8a', '#6b8de3', '#f0a23b', '#8b72d9'],
          borderRadius: 7,
          borderSkipped: false,
          maxBarThickness: 42
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (c) => `${c.raw} movimientos` } }
      },
      scales: {
        x: { beginAtZero: true, grid: { color: '#edf1f6' }, ticks: { precision: 0, font: { size: 10 } } },
        y: { grid: { display: false }, ticks: { font: { size: 10 } } }
      }
    }
  });

  make('pendingChart', {
    type: 'bar',
    data: {
      labels: ['Extracto bancario', 'Libro auxiliar'],
      datasets: [
        {
          label: 'Pendientes',
          data: [pendingBank, pendingBook],
          backgroundColor: ['#4f86e8', '#6a78c8'],
          borderRadius: 8,
          barThickness: 35
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (c) => `${c.raw} movimientos pendientes` } }
      },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 10 } }, grid: { color: '#edf1f6' } },
        x: { grid: { display: false }, ticks: { font: { size: 11 } } }
      }
    }
  });
}

function tableHtml(rows, tab) {
  if (!rows.length) return '<div class="empty">No hay registros en esta sección.</div>';

  const headers =
    tab === 'matched'
      ? ['Regla', 'Fecha banco', 'Descripción banco', 'Valor banco', 'Fecha contable', 'Beneficiario / detalle', 'Valor contable', 'Diferencia']
      : tab === 'review'
      ? ['Estado', 'Fecha', 'Banco / descripción', 'Valor banco', 'Contabilidad / tercero', 'Valor contable', 'Diferencia']
      : ['Origen', 'Fecha', 'Concepto / descripción', 'Valor', 'Observación'];

  let html = '<table><thead><tr>' + headers.map((h) => `<th>${h}</th>`).join('') + '</tr></thead><tbody>';

  for (const r of rows) {
    if (tab === 'matched') {
      const bs = Array.isArray(r.bank) ? r.bank : [r.bank];
      const cs = Array.isArray(r.book) ? r.book : [r.book];
      const cols = [
        `<span class="pill">${esc(r.rule)}</span>`,
        bs.map((x) => dateText(x.date)).join('<br>'),
        bs.map((x) => esc(x.desc)).join('<br>'),
        bs.map((x) => money(x.value)).join('<br>'),
        cs.map((x) => dateText(x.date)).join('<br>'),
        cs.map((x) => esc((x.beneficiary + ' ' + x.desc).trim())).join('<br>'),
        cs.map((x) => money(x.value)).join('<br>'),
        money(r.diff)
      ];
      html += '<tr>' + cols.map((v) => `<td>${v}</td>`).join('') + '</tr>';
    } else if (tab === 'review') {
      const cols = [
        '<span class="pill warn">Por revisar</span>',
        dateText(r.bank.date),
        esc(r.bank.desc),
        money(r.bank.value),
        esc((r.book.beneficiary + ' ' + r.book.desc).trim()),
        money(r.book.value),
        money(r.diff)
      ];
      html += '<tr>' + cols.map((v) => `<td>${v}</td>`).join('') + '</tr>';
    } else {
      const cols = [
        esc(r.side),
        dateText(r.item.date),
        esc((r.item.beneficiary + ' ' + r.item.desc).trim()),
        money(r.item.value),
        'No se encontró cruce válido'
      ];
      html += '<tr>' + cols.map((v) => `<td>${v}</td>`).join('') + '</tr>';
    }
  }

  return html + '</tbody></table>';
}

function showTable() {
  let rows = currentTab === 'matched' ? matched : currentTab === 'review' ? review : pending;
  const q = norm($('search').value);
  if (q) rows = rows.filter((r) => norm(JSON.stringify(r)).includes(q));
  $('tableHolder').innerHTML = tableHtml(rows, currentTab);
  $('rowCount').textContent = `${rows.length} registros`;
}

async function readInput(input, kind) {
  const f = $(input).files[0];
  if (!f) throw Error('Selecciona el archivo de ' + (kind === 'bank' ? 'extracto bancario' : 'libro auxiliar') + '.');
  const rows = await workbookRows(f);
  return kind === 'bank' ? bankRowsFrom(rows) : bookRowsFrom(rows);
}

$('analyzeBtn').addEventListener('click', async () => {
  try {
    $('status').textContent = 'Leyendo y cruzando movimientos…';
    $('status').className = 'status';
    rawBank = await readInput('bankFile', 'bank');
    rawBook = await readInput('bookFile', 'book');
    runReconciliation();
    $('status').textContent = `Análisis listo: ${matched.length} cruces, ${review.length} por revisar y ${pending.length} pendientes.`;
    $('status').className = 'status success';
  } catch (e) {
    $('status').textContent = e.message;
    $('status').className = 'status error';
  }
});

$('exampleBtn').addEventListener('click', () => {
  rawBank = [
    { id: 'B1', side: 'Banco', date: new Date(2026, 8, 3), desc: 'PAGO PROVE MARTHA LUCIA QUI', beneficiary: '', value: 20000000, used: false },
    { id: 'B2', side: 'Banco', date: new Date(2026, 8, 5), desc: 'COMISION BANCARIA', beneficiary: '', value: -200000, used: false },
    { id: 'B3', side: 'Banco', date: new Date(2026, 8, 8), desc: 'PAGO NOMINA ANA', beneficiary: '', value: 5000000, used: false },
    { id: 'B4', side: 'Banco', date: new Date(2026, 8, 8), desc: 'PAGO NOMINA LUIS', beneficiary: '', value: 4000000, used: false }
  ];
  rawBook = [
    { id: 'C1', side: 'Contabilidad', date: new Date(2026, 8, 3), desc: 'Pago proveedor', beneficiary: 'QUINTERO ARBOLEDA MARTHA LUCIA', value: 20000000, used: false },
    { id: 'C2', side: 'Contabilidad', date: new Date(2026, 8, 30), desc: 'Comisiones bancarias', beneficiary: '', value: -200000, used: false },
    { id: 'C3', side: 'Contabilidad', date: new Date(2026, 8, 8), desc: 'Nómina mensual', beneficiary: 'NOMINA', value: 9000000, used: false },
    { id: 'C4', side: 'Contabilidad', date: new Date(2026, 8, 12), desc: 'Pago proveedor', beneficiary: 'Otro tercero', value: 1234500, used: false }
  ];
  $('company').value = 'Empresa de práctica';
  $('fromDate').value = '';
  $('toDate').value = '';
  runReconciliation();
  $('status').textContent = 'Ejemplo cargado. Revisa los resultados y luego selecciona tus archivos.';
  $('status').className = 'status success';
});

$('clearBtn').addEventListener('click', () => {
  rawBank = [];
  rawBook = [];
  bank = [];
  book = [];
  matched = [];
  review = [];
  pending = [];
  $('bankFile').value = '';
  $('bookFile').value = '';
  $('fromDate').value = '';
  $('toDate').value = '';
  $('search').value = '';
  $('status').textContent = 'Selecciona ambos archivos o carga el ejemplo.';
  $('status').className = 'status';
  render();
  $('tableHolder').innerHTML = '<div class="empty">Carga los archivos y pulsa Analizar conciliación.</div>';
});

document.querySelectorAll('.tab').forEach((btn) =>
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((x) => x.classList.remove('active'));
    btn.classList.add('active');
    currentTab = btn.dataset.tab;
    showTable();
  })
);

$('search').addEventListener('input', showTable);
$('fromDate').addEventListener('change', () => {
  if (rawBank.length || rawBook.length) runReconciliation();
});
$('toDate').addEventListener('change', () => {
  if (rawBank.length || rawBook.length) runReconciliation();
});

function exportRows() {
  const out = [
    ['Tipo', 'Regla/Estado', 'Fecha banco', 'Descripción banco', 'Valor banco', 'Fecha contable', 'Beneficiario', 'Detalle contable', 'Valor contable', 'Diferencia']
  ];
  for (const m of matched) {
    const bs = Array.isArray(m.bank) ? m.bank : [m.bank];
    const cs = Array.isArray(m.book) ? m.book : [m.book];
    out.push([
      'Conciliado',
      m.rule,
      bs.map((x) => dateText(x.date)).join(' | '),
      bs.map((x) => x.desc).join(' | '),
      bs.reduce((s, x) => s + x.value, 0) / 100,
      cs.map((x) => dateText(x.date)).join(' | '),
      cs.map((x) => x.beneficiary).join(' | '),
      cs.map((x) => x.desc).join(' | '),
      cs.reduce((s, x) => s + x.value, 0) / 100,
      m.diff / 100
    ]);
  }
  for (const r of review) {
    out.push([
      'Por revisar',
      r.rule,
      dateText(r.bank.date),
      r.bank.desc,
      r.bank.value / 100,
      dateText(r.book.date),
      r.book.beneficiary,
      r.book.desc,
      r.book.value / 100,
      r.diff / 100
    ]);
  }
  for (const r of pending) {
    out.push([
      'Pendiente',
      r.side,
      dateText(r.item.date),
      r.item.desc,
      r.item.value / 100,
      '',
      r.item.beneficiary,
      '',
      '',
      ''
    ]);
  }
  return out;
}

$('excelBtn').addEventListener('click', () => {
  const rows = exportRows();
  const csv = rows
    .map((r) => r.map((v) => '"' + String(v ?? '').replace(/"/g, '""') + '"').join(';'))
    .join('\r\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'Informe_Conciliacion.csv';
  a.click();
  URL.revokeObjectURL(a.href);
});

$('pdfBtn').addEventListener('click', () => window.print());

render();

