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

// Lê célula da sheet por coluna (letra) e linha (1-based)
function cellVal(sheet, col, row) {
  const addr = col.toUpperCase() + row;
  const cell = sheet[addr];
  if (!cell) return '';
  if (cell.t === 'd' || cell.v instanceof Date) return cell.v;
  return cell.v !== undefined ? cell.v : '';
}

function parseWorkbook(wb) {
  allClients = [];
  allLicensees = [];
  const names = wb.SheetNames.map(s => s.toLowerCase().trim());

  // Aba Relatório — Clientes por Licenciado
  const relIdx = names.findIndex(s =>
    s.includes('relatório') || s.includes('relatorio') || s.includes('clientes por licenciado')
  );
  // Aba Total de Clientes por Licenciados
  const totIdx = names.findIndex(s =>
    (s.includes('total') && s.includes('licenciado')) ||
    (s.includes('total') && s.includes('cliente'))
  );

  const relSheet = wb.Sheets[wb.SheetNames[relIdx >= 0 ? relIdx : 0]];
  const totSheet = totIdx >= 0 ? wb.Sheets[wb.SheetNames[totIdx]] : null;

  // --- RELATÓRIO: colunas fixas, header na linha 4, dados a partir da linha 5 ---
  // A=ID cliente, B=Nome cliente, C=E-mail, D=Data Cadastro,
  // E=ID licenciado, G=Tipo de licença, H=Status
  if (relSheet) {
    const range = XLSX.utils.decode_range(relSheet['!ref'] || 'A1');
    for (let r = 4; r <= range.e.r + 1; r++) { // linha 5 em diante (índice 4)
      const id    = cellVal(relSheet, 'A', r + 1);
      const nome  = cellVal(relSheet, 'B', r + 1);
      if (!id && !nome) continue; // linha vazia
      allClients.push({
        id:          id,
        nome:        nome,
        email:       cellVal(relSheet, 'C', r + 1),
        data:        cellVal(relSheet, 'D', r + 1),
        idLicenciado:cellVal(relSheet, 'E', r + 1),
        licenciado:  cellVal(relSheet, 'F', r + 1), // coluna F = Nome licenciado (se existir)
        tipo:        cellVal(relSheet, 'G', r + 1),
        status:      cellVal(relSheet, 'H', r + 1)
      });
    }
  }

  // --- TOTAL DE CLIENTES POR LICENCIADOS: header na linha 4, dados a partir da linha 5 ---
  // J=ID licenciado, K=Nome, L=Tipo de licença, M=Qtd. Clientes (ajuste se necessário)
  if (totSheet) {
    const range = XLSX.utils.decode_range(totSheet['!ref'] || 'A1');
    // Descobrir qual coluna tem "qtd" / "clientes" no header (linha 4)
    // Tentamos M, N — ou buscamos dinamicamente
    let qtdCol = 'M';
    for (let c = colIndex('J'); c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 3, c }); // linha 4 = índice 3
      const cell = totSheet[addr];
      if (cell && String(cell.v).toLowerCase().includes('qtd')) {
        qtdCol = XLSX.utils.encode_col(c);
        break;
      }
      if (cell && String(cell.v).toLowerCase().includes('cliente') && !String(cell.v).toLowerCase().includes('id')) {
        qtdCol = XLSX.utils.encode_col(c);
        break;
      }
    }

    for (let r = 4; r <= range.e.r + 1; r++) {
      const id   = cellVal(totSheet, 'J', r + 1);
      const nome = cellVal(totSheet, 'K', r + 1);
      if (!id && !nome) continue;
      const qtd  = Number(cellVal(totSheet, qtdCol, r + 1)) || 0;
      const tipo = cellVal(totSheet, 'L', r + 1);
      allLicensees.push({ id, nome, tipo, qtd });
    }
  }

  // Fallback: calcular licenciados a partir dos clientes se a aba total não existir
  if (allLicensees.length === 0 && allClients.length > 0) {
    const map = {};
    allClients.forEach(c => {
      const k = c.licenciado || c.idLicenciado || 'Sem licenciado';
      if (!map[k]) map[k] = { id: c.idLicenciado, nome: k, tipo: c.tipo, qtd: 0 };
      map[k].qtd++;
    });
    allLicensees = Object.values(map);
  }

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
  // Tipo de licença vem da aba Total (coluna L) — agrupa por tipo
  const map = {};
  if (allLicensees.length > 0) {
    allLicensees.forEach(l => {
      const t = l.tipo || 'Não informado';
      map[t] = (map[t] || 0) + (l.qtd || 1);
    });
  } else {
    allClients.forEach(c => {
      const t = c.tipo || 'Não informado';
      map[t] = (map[t] || 0) + 1;
    });
  }
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
  // Filtro de licenciado vem do ID licenciado (coluna E) da aba Relatório
  const ids = [...new Set(allClients.map(c => c.idLicenciado).filter(Boolean))].sort();
  const sel = document.getElementById('filter-licenciado');
  sel.innerHTML = '<option value="">Todos os licenciados</option>' +
    ids.map(id => {
      // Tenta encontrar o nome correspondente em allLicensees
      const lic = allLicensees.find(l => String(l.id) === String(id));
      const label = lic ? `${lic.nome} (${id})` : id;
      return `<option value="${id}">${label}</option>`;
    }).join('');

  // Garante que os listeners não sejam duplicados ao recarregar planilha
  const searchEl = document.getElementById('search-input');
  const statusEl = document.getElementById('filter-status');
  searchEl.replaceWith(searchEl.cloneNode(true));
  statusEl.replaceWith(statusEl.cloneNode(true));
  sel.replaceWith(sel.cloneNode(true));

  document.getElementById('search-input').addEventListener('input', () => { currentPage = 1; renderTable(); });
  document.getElementById('filter-status').addEventListener('change', () => { currentPage = 1; renderTable(); });
  document.getElementById('filter-licenciado').addEventListener('change', () => { currentPage = 1; renderTable(); });
}

function getFiltered() {
  const q = document.getElementById('search-input').value.toLowerCase();
  const st = document.getElementById('filter-status').value.toLowerCase();
  const lic = document.getElementById('filter-licenciado').value;
  return allClients.filter(c => {
    if (q && !['nome', 'email', 'licenciado', 'idLicenciado'].some(k => (c[k] + '').toLowerCase().includes(q))) return false;
    if (st) {
      const s = (c.status + '').toLowerCase();
      if (st === 'ativo' && (!s.includes('ativo') || s.includes('inativo'))) return false;
      if (st === 'inativo' && !s.includes('inativo')) return false;
    }
    if (lic && String(c.idLicenciado) !== String(lic)) return false;
    return true;
  });
}

function fmtData(d) {
  if (!d) return '—';
  if (d instanceof Date && !isNaN(d)) return d.toLocaleDateString('pt-BR');
  const s = String(d);
  const m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return `${m[3].padStart(2,'0')}/${m[2].padStart(2,'0')}/${m[1]}`;
  return s;
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
    ? page.map(c => `
      <tr>
        <td class="td-id">${c.id || '—'}</td>
        <td>${c.nome || '—'}</td>
        <td class="td-muted">${c.email || '—'}</td>
        <td class="td-muted">${fmtData(c.data)}</td>
        <td class="td-id">${c.idLicenciado || '—'}</td>
        <td><span class="badge badge-tipo">${c.tipo || '—'}</span></td>
        <td>${badgeStatus(c.status)}</td>
      </tr>
    `).join('')
    : '<tr class="empty-row"><td colspan="7">Nenhum registro encontrado para os filtros selecionados.</td></tr>';

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