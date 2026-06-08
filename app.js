// Plugin global registrado antes dos graficos
Chart.register({
  id: 'barLabels',
  afterDraw(chart) {
    if (!chart.config.options.plugins.barLabels) return;
    const ctx = chart.ctx;
    chart.data.datasets.forEach((dataset, i) => {
      const meta = chart.getDatasetMeta(i);
      meta.data.forEach((bar, idx) => {
        const val = dataset.data[idx];
        if (val === undefined || val === null) return;
        ctx.save();
        ctx.font = '600 12px "DM Sans", sans-serif';
        ctx.fillStyle = '#e8eaf0';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'middle';
        ctx.fillText(val.toLocaleString('pt-BR'), bar.x + 6, bar.y);
        ctx.restore();
      });
    });
  }
});

let allClients = [], allLicensees = [];
let currentPage = 1;
const PAGE_SIZE = 12;
let chartTop, chartTipo, chartTime;
let sortQtd = null; // null = data desc, 'desc' = maior→menor, 'asc' = menor→maior

const COLORS = ['#06b6d4','#10b981','#3b7bff','#ec4899','#34d399','#a78bfa','#fb923c','#38bdf8'];

// LOADING OVERLAY
function showLoading(msg) {
  let el = document.getElementById('loading-overlay');
  if (!el) {
    el = document.createElement('div');
    el.id = 'loading-overlay';
    el.innerHTML = `
      <div class="loading-box">
        <div class="loading-spinner"></div>
        <div class="loading-msg" id="loading-msg"></div>
      </div>`;
    document.body.appendChild(el);
  }
  document.getElementById('loading-msg').textContent = msg || 'Lendo planilha...';
  el.style.display = 'flex';
}

function hideLoading() {
  const el = document.getElementById('loading-overlay');
  if (el) el.style.display = 'none';
}

// FILE INPUT
document.getElementById('file-input').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('upload-label').textContent = file.name;
  showLoading('Lendo planilha "' + file.name + '"...');
  const reader = new FileReader();
  reader.onload = function(ev) {
    setTimeout(() => {
      try {
        const wb = XLSX.read(ev.target.result, { type: 'array', cellDates: true });
        parseWorkbook(wb);
        hideLoading();
        showToast('Planilha carregada com sucesso!', 'success');
      } catch(err) {
        hideLoading();
        showToast('Erro ao ler planilha: ' + err.message, 'error');
      }
    }, 50);
  };
  reader.readAsArrayBuffer(file);
});

// DRAG & DROP
const uploadDiv = document.querySelector('.upload-area');
uploadDiv.addEventListener('dragover', e => { e.preventDefault(); uploadDiv.style.borderColor = 'var(--accent)'; });
uploadDiv.addEventListener('dragleave', () => { uploadDiv.style.borderColor = ''; });
uploadDiv.addEventListener('drop', e => {
  e.preventDefault();
  uploadDiv.style.borderColor = '';
  const file = e.dataTransfer.files[0];
  if (file) {
    document.getElementById('file-input').files = e.dataTransfer.files;
    document.getElementById('file-input').dispatchEvent(new Event('change'));
  }
});

// Converte letra de coluna (A, B, ..., Z, AA...) para índice 0-based
function colIndex(letter) {
  letter = letter.toUpperCase();
  let n = 0;
  for (let i = 0; i < letter.length; i++) {
    n = n * 26 + (letter.charCodeAt(i) - 64);
  }
  return n - 1;
}

// Converte número serial do Excel para Date
function excelSerialToDate(serial) {
  // Excel serial: dias desde 1900-01-01 (com bug do ano bissexto 1900)
  const utc_days = Math.floor(serial - 25569);
  const utc_value = utc_days * 86400 * 1000;
  return new Date(utc_value);
}

// Lê célula da sheet por coluna (letra) e linha (1-based)
function cellVal(sheet, col, row) {
  const addr = col.toUpperCase() + row;
  const cell = sheet[addr];
  if (!cell) return '';
  // Já é Date
  if (cell.t === 'd' || cell.v instanceof Date) return cell.v;
  // Número serial de data
  if (cell.t === 'n' && cell.z && (
    String(cell.z).includes('yy') ||
    String(cell.z).includes('mm') ||
    String(cell.z).includes('dd') ||
    String(cell.z).includes('d/') ||
    String(cell.z).includes('m/')
  )) {
    return excelSerialToDate(cell.v);
  }
  // Texto no formato YYYY-MM-DD
  if (cell.t === 's') {
    const m = String(cell.v).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  return cell.v !== undefined ? cell.v : '';
}

function readDataAtivo(cell) {
  if (!cell) return '';
  if (cell.t === 'd' || cell.v instanceof Date) return cell.v;
  if (cell.t === 'n') return excelSerialToDate(cell.v);
  // texto: "2025-01-23" ou qualquer formato
  return String(cell.v).trim();
}

function parseWorkbook(wb) {
  allClients = [];
  allLicensees = [];
  const names = wb.SheetNames.map(s => s.toLowerCase().trim());

  // Única aba: Relatório
  const relIdx = names.findIndex(s =>
    s.includes('relatório') || s.includes('relatorio') || s.includes('relat')
  );
  const relSheet = wb.Sheets[wb.SheetNames[relIdx >= 0 ? relIdx : 0]];

  if (!relSheet) { buildDashboard(); return; }

  const range = XLSX.utils.decode_range(relSheet['!ref'] || 'A1');

  // --- CLIENTES: colunas A-H, header linha 4, dados linha 5+ ---
  // A=ID cliente, B=Nome, C=E-mail, D=Data Cadastro,
  // E=ID licenciado, F=Nome licenciado, G=Tipo licença, H=Status
  for (let row = 5; row <= range.e.r + 1; row++) {
    const id   = cellVal(relSheet, 'A', row);
    const nome = cellVal(relSheet, 'B', row);
    if (!id && !nome) continue;
    allClients.push({
      id,
      nome,
      email:        cellVal(relSheet, 'C', row),
      data:         cellVal(relSheet, 'D', row),
      idLicenciado: cellVal(relSheet, 'E', row),
      licenciado:   cellVal(relSheet, 'F', row),
      tipo:         cellVal(relSheet, 'G', row),
      status:       cellVal(relSheet, 'H', row)
    });
  }

  // --- LICENCIADOS: colunas J-N, header linha 4, dados linha 5+ ---
  // J=ID licenciado, K=Nome, L=Tipo de licença, M=Qtd. Clientes, N=Data Ativo
  const licMap = {};
  for (let row = 5; row <= range.e.r + 1; row++) {
    const id = cellVal(relSheet, 'J', row);
    if (!id) continue;
    if (licMap[String(id)]) continue; // já registrou esse licenciado
    licMap[String(id)] = {
      id,
      nome:      cellVal(relSheet, 'K', row),
      tipo:      cellVal(relSheet, 'L', row),
      qtd:       Number(cellVal(relSheet, 'M', row)) || 0,
      dataAtivo: readDataAtivo(relSheet['N' + row])
    };
  }
  allLicensees = Object.values(licMap);

  buildDashboard();
}

function buildDashboard() {
  document.getElementById('upload-area').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';

  const total = allClients.length;
  const licCount = allLicensees.length || new Set(allClients.map(c => c.licenciado)).size;

  // Top licenciado = maior Qtd. Clientes da aba Total
  const topLic = allLicensees.length
    ? [...allLicensees].sort((a, b) => b.qtd - a.qtd)[0]
    : null;

  document.getElementById('topbar-sub').textContent =
    `${total.toLocaleString('pt-BR')} clientes · ${licCount} licenciados`;

  document.getElementById('metrics-row').innerHTML = `
    <div class="metric blue">
      <div class="metric-label">Total de clientes</div>
      <div class="metric-value">${total.toLocaleString('pt-BR')}</div>
      <div class="metric-sub">registros na planilha</div>
      <div class="metric-icon">👥</div>
    </div>
    <div class="metric amber">
      <div class="metric-label">Licenciados</div>
      <div class="metric-value">${licCount}</div>
      <div class="metric-sub">cadastrados</div>
      <div class="metric-icon">🏷️</div>
    </div>
    ${topLic ? `
    <div class="metric teal">
      <div class="metric-label">Top licenciado</div>
      <div class="metric-value" style="font-size:18px;line-height:1.3;">${topLic.nome}</div>
      <div class="metric-sub">${topLic.qtd.toLocaleString('pt-BR')} clientes</div>
      <div class="metric-icon">🥇</div>
    </div>` : ''}
  `;

  buildChartTop();
  buildChartTipo();
  buildChartTime();
  buildFilters();
  renderTable();
}

function buildChartTop() {
  const data = [...allLicensees].sort((a, b) => b.qtd - a.qtd).slice(0, 10);
  document.getElementById('badge-top').textContent = `top ${data.length}`;
  if (chartTop) chartTop.destroy();
  chartTop = new Chart(document.getElementById('chart-top'), {
    type: 'bar',
    data: {
      labels: data.map(d => d.id ? `${d.id} - ${d.nome || 'S/N'}` : (d.nome || 'S/N')),
      datasets: [{
        label: 'Clientes',
        data: data.map(d => d.qtd),
        backgroundColor: data.map((_, i) => i === 0 ? '#3b7bff' : 'rgba(59,123,255,0.35)'),
        borderRadius: 4,
        borderSkipped: false
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { enabled: false },
        barLabels: {}
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: { stepSize: 1, color: '#555d75', font: { size: 11 } },
          grid: { color: 'rgba(255,255,255,0.04)' }
        },
        y: {
          ticks: { color: '#8b92a8', font: { size: 12 } },
          grid: { display: false }
        }
      }
    }
  });
}


function buildChartTipo() {
  // Conta quantos licenciados há por tipo (coluna L), sem somar qtd
  const map = {};
  allLicensees.forEach(l => {
    const t = l.tipo || 'Não informado';
    map[t] = (map[t] || 0) + 1;
  });
  const labels = Object.keys(map);
  const vals = Object.values(map);
  if (chartTipo) chartTipo.destroy();
  const legendEl = document.getElementById('legend-tipo');
  legendEl.innerHTML = labels.map((l, i) =>
    `<div class="legend-item"><span class="legend-dot" style="background:${COLORS[i % COLORS.length]}"></span>${l} <strong style="color:var(--text)">(${vals[i]})</strong></div>`
  ).join('');
  chartTipo = new Chart(document.getElementById('chart-tipo'), {
    type: 'doughnut',
    data: { labels, datasets: [{ data: vals, backgroundColor: COLORS, borderWidth: 0 }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '65%',
      plugins: { legend: { display: false } }
    }
  });
}

function buildChartTime() {
  const map = {};
  allClients.forEach(c => {
    if (!c.data) return;
    let key = '';
    const d = c.data;
    if (d instanceof Date && !isNaN(d)) {
      key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
    } else {
      const s = String(d);
      const m = s.match(/(\d{4})[\/\-](\d{1,2})/);
      if (m) key = m[1] + '-' + m[2].padStart(2, '0');
      else {
        const m2 = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (m2) key = m2[3] + '-' + m2[2].padStart(2, '0');
      }
    }
    if (key) map[key] = (map[key] || 0) + 1;
  });
  const sorted = Object.keys(map).sort();
  const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  document.getElementById('badge-time').textContent = sorted.length + ' meses';
  if (chartTime) chartTime.destroy();
  chartTime = new Chart(document.getElementById('chart-time'), {
    type: 'line',
    data: {
      labels: sorted.map(k => { const [y, m] = k.split('-'); return meses[+m - 1] + '/' + y.slice(2); }),
      datasets: [{
        label: 'Cadastros',
        data: sorted.map(k => map[k]),
        borderColor: '#3b7bff',
        backgroundColor: 'rgba(59,123,255,0.08)',
        fill: true,
        tension: 0.4,
        pointRadius: 4,
        pointBackgroundColor: '#3b7bff',
        pointBorderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { stepSize: 1, color: '#555d75', font: { size: 11 } },
          grid: { color: 'rgba(255,255,255,0.04)' }
        },
        x: {
          ticks: { color: '#555d75', font: { size: 11 }, autoSkip: true, maxTicksLimit: 14 },
          grid: { display: false }
        }
      }
    }
  });
}

function buildFilters() {
  // Filtro de tipo de licença
  const tipos = [...new Set(allLicensees.map(l => l.tipo).filter(Boolean))].sort();

  // Garante que os listeners não sejam duplicados ao recarregar planilha
  const searchEl = document.getElementById('search-input');
  const sel = document.getElementById('filter-licenciado');
  const newSearch = searchEl.cloneNode(true);
  const newSel = sel.cloneNode(false); // só o elemento, sem filhos

  newSel.innerHTML = '<option value="">Todos os tipos</option>' +
    tipos.map(t => `<option value="${t}">${t}</option>`).join('');

  searchEl.replaceWith(newSearch);
  sel.replaceWith(newSel);

  document.getElementById('search-input').addEventListener('input', () => { currentPage = 1; renderTable(); });
  document.getElementById('filter-licenciado').addEventListener('change', () => { currentPage = 1; renderTable(); });

  const thQtd = document.getElementById('th-qtd');
  if (thQtd) {
    thQtd.onclick = () => {
      sortQtd = sortQtd === 'desc' ? 'asc' : 'desc';
      document.getElementById('sort-arrow').textContent = sortQtd === 'desc' ? '↓' : '↑';
      currentPage = 1;
      renderTable();
    };
  }
}

function parseDate(d) {
  if (!d) return new Date(0);
  if (d instanceof Date) return isNaN(d) ? new Date(0) : d;
  const s = String(d).trim();
  // YYYY-MM-DD
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  // DD/MM/YYYY
  const m2 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m2) return new Date(Number(m2[3]), Number(m2[2]) - 1, Number(m2[1]));
  return new Date(d) || new Date(0);
}

function getFiltered() {
  const q = document.getElementById('search-input').value.toLowerCase();
  const tipo = document.getElementById('filter-licenciado').value;
  // Ordenação: mais recente para mais antigo por dataAtivo
  const filtered = allLicensees.filter(l => {
    if (q && !['id', 'nome', 'tipo'].some(k => (l[k] + '').toLowerCase().includes(q))) return false;
    if (tipo && l.tipo !== tipo) return false;
    return true;
  });
  if (sortQtd) {
    filtered.sort((a, b) => sortQtd === 'desc' ? b.qtd - a.qtd : a.qtd - b.qtd);
  } else {
    filtered.sort((a, b) => parseDate(b.dataAtivo) - parseDate(a.dataAtivo));
  }
  return filtered;
}

function fmtData(d) {
  if (!d && d !== 0) return '—';
  if (d instanceof Date && !isNaN(d)) return d.toLocaleDateString('pt-BR');
  const s = String(d).trim();
  // YYYY-MM-DD
  const m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m1) return `${m1[3]}/${m1[2]}/${m1[1]}`;
  // YYYY/MM/DD ou YYYY-M-D
  const m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m2) return `${m2[3].padStart(2,'0')}/${m2[2].padStart(2,'0')}/${m2[1]}`;
  // DD/MM/YYYY
  const m3 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m3) return s;
  return s || '—';
}

function badgeStatus(s) {
  const sl = (s + '').toLowerCase();
  const active = sl.includes('ativo') && !sl.includes('inativo');
  return `<span class="badge ${active ? 'badge-active' : 'badge-inactive'}">${s || '—'}</span>`;
}

function renderTable() {
  const filtered = getFiltered();
  const total = filtered.length;
  const start = (currentPage - 1) * PAGE_SIZE;
  const page = filtered.slice(start, start + PAGE_SIZE);

  document.getElementById('table-body').innerHTML = page.length
    ? page.map(l => `
      <tr>
        <td class="td-id">${l.id || '—'}</td>
        <td>${l.nome || '—'}</td>
        <td><span class="badge badge-tipo">${l.tipo || '—'}</span></td>
        <td class="td-muted" style="text-align:right;padding-right:2rem;">${l.qtd ? l.qtd.toLocaleString('pt-BR') : '0'}</td>
        <td class="td-muted">${fmtData(l.dataAtivo)}</td>
      </tr>
    `).join('')
    : '<tr class="empty-row"><td colspan="5">Nenhum registro encontrado para os filtros selecionados.</td></tr>';

  const pages = Math.ceil(total / PAGE_SIZE);
  document.getElementById('pagination-info').textContent = total
    ? `${start + 1}–${Math.min(start + PAGE_SIZE, total)} de ${total.toLocaleString('pt-BR')} registros`
    : '0 registros';
  document.getElementById('btn-prev').disabled = currentPage <= 1;
  document.getElementById('btn-next').disabled = currentPage >= pages;
}

function changePage(dir) {
  const pages = Math.ceil(getFiltered().length / PAGE_SIZE);
  currentPage = Math.max(1, Math.min(currentPage + dir, pages));
  renderTable();
}

function showToast(msg, type = 'success') {
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3500);
}