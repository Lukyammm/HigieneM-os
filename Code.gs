const CONFIG = {
  DATA_SHEET_NAME: 'Respostas ao formulário 1',
  SETTINGS_SHEET_NAME: 'Configurações (Não Editar)',

  DATA_START_ROW: 2,
  SETTINGS_START_ROW: 2,

  // ABA DE DADOS
  DATA_COLUMN_DATE: 1,      // A
  DATA_COLUMN_UNIDADE: 2,   // B
  DATA_COLUMN_CATEGORIA: 3, // C
  DATA_COLUMN_MOMENTO: 4,   // D
  DATA_COLUMN_ACAO: 5,      // E

  // ABA DE CONFIGURAÇÕES
  SETTINGS_COLUMN_UNIDADE: 1,   // A
  SETTINGS_COLUMN_CATEGORIA: 2, // B
  SETTINGS_COLUMN_MOMENTO: 3,   // C
  SETTINGS_COLUMN_ACAO: 4       // D
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Análise de Higiene das Mãos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function openSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error(
      'Planilha não vinculada ao projeto. Defina Script Property SPREADSHEET_ID com o ID da planilha para o Web App.'
    );
  }

  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(
      'Não foi possível acessar a planilha (SPREADSHEET_ID=' + spreadsheetId + '). Verifique compartilhamento e permissões do Web App. Detalhe: ' +
      (error && error.message ? error.message : error)
    );
  }
}

function checkSpreadsheetAccess() {
  const ss = openSpreadsheet_();
  const sheets = ss.getSheets().map(s => s.getName());
  const dataSheet = resolveDataSheet_(ss);
  const settingsSheet = resolveSettingsSheet_(ss);

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    dataSheetExists: !!dataSheet,
    settingsSheetExists: !!settingsSheet,
    dataSheetName: dataSheet ? dataSheet.getName() : '',
    settingsSheetName: settingsSheet ? settingsSheet.getName() : '',
    availableSheets: sheets
  };
}

function getDashboardData(filters) {
  filters = filters || {};

  const baseRows = getBaseRows_();
  const normalized = baseRows.map(normalizeDataRow_).filter(Boolean);
  const filtered = applyFilters_(normalized, filters);
  const settings = getSettingsData_();

  return {
    meta: {
      generatedAt: new Date(),
      totalBase: normalized.length,
      totalFiltered: filtered.length,
      availableFilters: buildAvailableFilters_(normalized, settings)
    },
    kpis: buildKpis_(filtered),
    charts: {
      timeline: buildTimeline_(filtered),
      byUnidade: buildByField_(filtered, 'unidade'),
      byCategoria: buildByField_(filtered, 'categoria'),
      byMomento: buildByField_(filtered, 'momento'),
      byStatus: buildByField_(filtered, 'status'),
      byMetodo: buildByField_(filtered, 'metodo')
    },
    quality: buildQuality_(filtered),
    insights: buildInsights_(filtered),
    table: filtered.slice(0, 1000)
  };
}

function getFilteredTable(filters) {
  filters = filters || {};
  const baseRows = getBaseRows_();
  const normalized = baseRows.map(normalizeDataRow_).filter(Boolean);
  return applyFilters_(normalized, filters).slice(0, 2000);
}

function getBaseRows_() {
  const ss = openSpreadsheet_();
  const sheet = resolveDataSheet_(ss);

  if (!sheet) {
    const available = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(
      'A aba "' + CONFIG.DATA_SHEET_NAME + '" não foi encontrada. Abas disponíveis: ' + (available || 'nenhuma') + '.'
    );
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];

  const width = Math.max(CONFIG.DATA_COLUMN_ACAO, sheet.getLastColumn());

  return sheet.getRange(
    CONFIG.DATA_START_ROW,
    1,
    lastRow - CONFIG.DATA_START_ROW + 1,
    width
  ).getValues();
}

function getSettingsData_() {
  const ss = openSpreadsheet_();
  const sheet = resolveSettingsSheet_(ss);

  if (!sheet) {
    return {
      unidade: [],
      categoria: [],
      momento: [],
      acao: []
    };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.SETTINGS_START_ROW) {
    return {
      unidade: [],
      categoria: [],
      momento: [],
      acao: []
    };
  }

  const values = sheet.getRange(
    CONFIG.SETTINGS_START_ROW,
    1,
    lastRow - CONFIG.SETTINGS_START_ROW + 1,
    CONFIG.SETTINGS_COLUMN_ACAO
  ).getValues();

  const unidade = [];
  const categoria = [];
  const momento = [];
  const acao = [];

  values.forEach(row => {
    const u = cleanText_(row[CONFIG.SETTINGS_COLUMN_UNIDADE - 1]);
    const c = cleanText_(row[CONFIG.SETTINGS_COLUMN_CATEGORIA - 1]);
    const m = cleanText_(row[CONFIG.SETTINGS_COLUMN_MOMENTO - 1]);
    const a = cleanText_(row[CONFIG.SETTINGS_COLUMN_ACAO - 1]);

    if (u) unidade.push(u);
    if (c) categoria.push(c);
    if (m) momento.push(m);
    if (a) acao.push(a);
  });

  return {
    unidade: uniqueSorted_(unidade),
    categoria: uniqueSorted_(categoria),
    momento: uniqueSorted_(momento),
    acao: uniqueSorted_(acao)
  };
}

function normalizeDataRow_(row) {
  const rawDate = row[CONFIG.DATA_COLUMN_DATE - 1];
  const unidade = cleanText_(row[CONFIG.DATA_COLUMN_UNIDADE - 1]);
  const categoria = cleanText_(row[CONFIG.DATA_COLUMN_CATEGORIA - 1]);
  const momento = cleanText_(row[CONFIG.DATA_COLUMN_MOMENTO - 1]);
  const acao = cleanText_(row[CONFIG.DATA_COLUMN_ACAO - 1]);

  const parsedDate = parseDate_(rawDate);
  const status = classifyStatus_(acao);
  const metodo = classifyMethod_(acao);
  const completeness = classifyCompleteness_(unidade, categoria, momento, acao);

  return {
    timestamp: parsedDate ? parsedDate.toISOString() : '',
    data: parsedDate ? formatDateBr_(parsedDate) : '',
    ano: parsedDate ? parsedDate.getFullYear() : '',
    mesNumero: parsedDate ? parsedDate.getMonth() + 1 : '',
    mesLabel: parsedDate ? Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'MM/yyyy') : 'Sem data',
    diaLabel: parsedDate ? Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'Sem data',

    unidade: unidade || 'Não informado',
    categoria: categoria || 'Não informado',
    momento: momento || 'Não informado',
    acao: acao || 'Não informado',

    status: status,
    metodo: metodo,
    completeness: completeness,

    unidadeMissing: !unidade,
    categoriaMissing: !categoria,
    momentoMissing: !momento,
    acaoMissing: !acao
  };
}

function applyFilters_(rows, filters) {
  return rows.filter(r => {
    if (filters.unidade && filters.unidade !== 'TODOS' && r.unidade !== filters.unidade) return false;
    if (filters.categoria && filters.categoria !== 'TODOS' && r.categoria !== filters.categoria) return false;
    if (filters.momento && filters.momento !== 'TODOS' && r.momento !== filters.momento) return false;
    if (filters.status && filters.status !== 'TODOS' && r.status !== filters.status) return false;
    if (filters.metodo && filters.metodo !== 'TODOS' && r.metodo !== filters.metodo) return false;
    if (filters.completeness && filters.completeness !== 'TODOS' && r.completeness !== filters.completeness) return false;

    if (filters.startDate) {
      const start = new Date(filters.startDate + 'T00:00:00');
      if (!r.timestamp || new Date(r.timestamp) < start) return false;
    }

    if (filters.endDate) {
      const end = new Date(filters.endDate + 'T23:59:59');
      if (!r.timestamp || new Date(r.timestamp) > end) return false;
    }

    return true;
  });
}

function buildAvailableFilters_(rows, settings) {
  const dataFilters = {
    unidade: uniqueSorted_(rows.map(r => r.unidade).filter(v => v && v !== 'Não informado')),
    categoria: uniqueSorted_(rows.map(r => r.categoria).filter(v => v && v !== 'Não informado')),
    momento: uniqueSorted_(rows.map(r => r.momento).filter(v => v && v !== 'Não informado')),
    status: uniqueSorted_(rows.map(r => r.status)),
    metodo: uniqueSorted_(rows.map(r => r.metodo)),
    completeness: uniqueSorted_(rows.map(r => r.completeness))
  };

  return {
    unidade: mergeUniqueSorted_(settings.unidade, dataFilters.unidade),
    categoria: mergeUniqueSorted_(settings.categoria, dataFilters.categoria),
    momento: mergeUniqueSorted_(settings.momento, dataFilters.momento),
    status: dataFilters.status,
    metodo: dataFilters.metodo,
    completeness: dataFilters.completeness
  };
}

function buildKpis_(rows) {
  const total = rows.length;
  const realizados = rows.filter(r => r.status === 'Realizado').length;
  const naoRealizados = rows.filter(r => r.status === 'Não realizado').length;
  const incompletos = rows.filter(r => r.status === 'Incompleto').length;

  const preenchimentoCompleto = rows.filter(r => r.completeness === 'Completo').length;
  const preenchimentoIncompleto = rows.filter(r => r.completeness === 'Incompleto').length;

  return {
    totalObservacoes: total,
    realizados: realizados,
    naoRealizados: naoRealizados,
    incompletos: incompletos,
    adesaoGeral: percentage_(realizados, total),
    taxaNaoRealizacao: percentage_(naoRealizados, total),
    taxaIncompletos: percentage_(incompletos, total),
    taxaPreenchimentoCompleto: percentage_(preenchimentoCompleto, total),
    taxaPreenchimentoIncompleto: percentage_(preenchimentoIncompleto, total)
  };
}

function buildTimeline_(rows) {
  const map = {};

  rows.forEach(r => {
    const key = r.mesLabel || 'Sem data';

    if (!map[key]) {
      map[key] = {
        label: key,
        total: 0,
        realizado: 0,
        naoRealizado: 0,
        incompleto: 0
      };
    }

    map[key].total++;
    if (r.status === 'Realizado') map[key].realizado++;
    if (r.status === 'Não realizado') map[key].naoRealizado++;
    if (r.status === 'Incompleto') map[key].incompleto++;
  });

  return Object.values(map)
    .sort((a, b) => sortMonthLabel_(a.label, b.label))
    .map(item => ({
      label: item.label,
      total: item.total,
      realizado: item.realizado,
      naoRealizado: item.naoRealizado,
      incompleto: item.incompleto,
      adesao: percentageNumber_(item.realizado, item.total)
    }));
}

function buildByField_(rows, field) {
  const map = {};

  rows.forEach(r => {
    const key = r[field] || 'Não informado';

    if (!map[key]) {
      map[key] = {
        label: key,
        total: 0,
        realizado: 0,
        naoRealizado: 0,
        incompleto: 0
      };
    }

    map[key].total++;
    if (r.status === 'Realizado') map[key].realizado++;
    if (r.status === 'Não realizado') map[key].naoRealizado++;
    if (r.status === 'Incompleto') map[key].incompleto++;
  });

  return Object.values(map)
    .map(item => ({
      label: item.label,
      total: item.total,
      realizado: item.realizado,
      naoRealizado: item.naoRealizado,
      incompleto: item.incompleto,
      adesao: percentageNumber_(item.realizado, item.total),
      taxaNaoRealizacao: percentageNumber_(item.naoRealizado, item.total),
      taxaIncompletos: percentageNumber_(item.incompleto, item.total)
    }))
    .sort((a, b) => b.total - a.total);
}

function buildQuality_(rows) {
  const total = rows.length;

  const missingUnidade = rows.filter(r => r.unidadeMissing).length;
  const missingCategoria = rows.filter(r => r.categoriaMissing).length;
  const missingMomento = rows.filter(r => r.momentoMissing).length;
  const missingAcao = rows.filter(r => r.acaoMissing).length;

  return {
    total: total,
    missingUnidade: missingUnidade,
    missingCategoria: missingCategoria,
    missingMomento: missingMomento,
    missingAcao: missingAcao,
    pctMissingUnidade: percentage_(missingUnidade, total),
    pctMissingCategoria: percentage_(missingCategoria, total),
    pctMissingMomento: percentage_(missingMomento, total),
    pctMissingAcao: percentage_(missingAcao, total)
  };
}

function buildInsights_(rows) {
  if (!rows.length) {
    return ['Nenhum registro encontrado com os filtros aplicados.'];
  }

  const byUnidade = buildByField_(rows, 'unidade')
    .filter(x => x.label !== 'Não informado' && x.total >= 3);

  const byCategoria = buildByField_(rows, 'categoria')
    .filter(x => x.label !== 'Não informado' && x.total >= 3);

  const byMomento = buildByField_(rows, 'momento')
    .filter(x => x.label !== 'Não informado' && x.total >= 3);

  const worstUnidade = byUnidade.slice().sort((a, b) => b.taxaNaoRealizacao - a.taxaNaoRealizacao)[0];
  const bestUnidade = byUnidade.slice().sort((a, b) => b.adesao - a.adesao)[0];
  const worstCategoria = byCategoria.slice().sort((a, b) => b.taxaNaoRealizacao - a.taxaNaoRealizacao)[0];
  const worstMomento = byMomento.slice().sort((a, b) => b.taxaNaoRealizacao - a.taxaNaoRealizacao)[0];
  const mostIncompleteMomento = byMomento.slice().sort((a, b) => b.taxaIncompletos - a.taxaIncompletos)[0];

  const total = rows.length;
  const realizados = rows.filter(r => r.status === 'Realizado').length;
  const naoRealizados = rows.filter(r => r.status === 'Não realizado').length;
  const incompletos = rows.filter(r => r.status === 'Incompleto').length;

  const insights = [];

  insights.push(
    `A adesão geral no recorte atual está em ${percentage_(realizados, total)}, com ${naoRealizados} registro(s) de não realização e ${incompletos} registro(s) incompleto(s).`
  );

  if (worstUnidade) {
    insights.push(
      `A unidade com maior taxa de não realização é "${worstUnidade.label}", com ${formatPct_(worstUnidade.taxaNaoRealizacao)} em ${worstUnidade.total} observação(ões).`
    );
  }

  if (bestUnidade) {
    insights.push(
      `A melhor adesão entre as unidades analisadas aparece em "${bestUnidade.label}", com ${formatPct_(bestUnidade.adesao)} de realização.`
    );
  }

  if (worstCategoria) {
    insights.push(
      `A categoria profissional mais crítica no recorte atual é "${worstCategoria.label}", com ${formatPct_(worstCategoria.taxaNaoRealizacao)} de não realização.`
    );
  }

  if (worstMomento) {
    insights.push(
      `O momento com pior desempenho é "${worstMomento.label}", com ${formatPct_(worstMomento.taxaNaoRealizacao)} de não realização.`
    );
  }

  if (mostIncompleteMomento) {
    insights.push(
      `O maior percentual de registros incompletos aparece em "${mostIncompleteMomento.label}", com ${formatPct_(mostIncompleteMomento.taxaIncompletos)}.`
    );
  }

  return insights;
}

function classifyStatus_(acao) {
  const value = (acao || '').toLowerCase().trim();

  if (!value) return 'Incompleto';
  if (value.indexOf('não realizado') !== -1 || value.indexOf('nao realizado') !== -1) return 'Não realizado';
  if (value.indexOf('realizado') !== -1) return 'Realizado';

  return 'Incompleto';
}

function classifyMethod_(acao) {
  const value = (acao || '').toLowerCase().trim();

  if (!value) return 'Não informado';
  if (value.indexOf('água') !== -1 || value.indexOf('agua') !== -1 || value.indexOf('sabonete') !== -1) return 'Água e sabonete';
  if (value.indexOf('álcool') !== -1 || value.indexOf('alcool') !== -1 || value.indexOf('fricção') !== -1 || value.indexOf('friccao') !== -1) return 'Fricção com álcool';
  if (value.indexOf('não realizado') !== -1 || value.indexOf('nao realizado') !== -1) return 'Não realizado';
  if (value.indexOf('realizado') !== -1) return 'Realizado sem método detalhado';

  return 'Não informado';
}

function classifyCompleteness_(unidade, categoria, momento, acao) {
  return (unidade && categoria && momento && acao) ? 'Completo' : 'Incompleto';
}

function parseDate_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }

  const d = new Date(value);
  if (!isNaN(d)) return d;

  return null;
}

function resolveDataSheet_(ss) {
  const exact = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  if (exact) return exact;

  const normalizedExpected = normalizeSheetName_(CONFIG.DATA_SHEET_NAME);
  const byNormalized = ss.getSheets().find(s => normalizeSheetName_(s.getName()) === normalizedExpected);
  if (byNormalized) return byNormalized;

  const byPrefix = ss.getSheets().find(s => normalizeSheetName_(s.getName()).indexOf('respostas ao formulario') === 0);
  if (byPrefix) return byPrefix;

  return null;
}

function resolveSettingsSheet_(ss) {
  const exact = ss.getSheetByName(CONFIG.SETTINGS_SHEET_NAME);
  if (exact) return exact;

  const normalizedExpected = normalizeSheetName_(CONFIG.SETTINGS_SHEET_NAME);
  const byNormalized = ss.getSheets().find(s => normalizeSheetName_(s.getName()) === normalizedExpected);
  if (byNormalized) return byNormalized;

  const byPrefix = ss.getSheets().find(s => normalizeSheetName_(s.getName()).indexOf('configuracoes') === 0);
  if (byPrefix) return byPrefix;

  return null;
}

function normalizeSheetName_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/\u00A0/g, ' ')
    .trim();
}

function uniqueSorted_(arr) {
  return [...new Set(arr.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
}

function mergeUniqueSorted_(arr1, arr2) {
  return [...new Set([...(arr1 || []), ...(arr2 || [])])]
    .filter(Boolean)
    .sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
}

function percentage_(part, total) {
  if (!total) return '0,0%';
  return ((part / total) * 100).toFixed(1).replace('.', ',') + '%';
}

function percentageNumber_(part, total) {
  if (!total) return 0;
  return Number(((part / total) * 100).toFixed(1));
}

function formatPct_(n) {
  return String((n || 0).toFixed(1)).replace('.', ',') + '%';
}

function formatDateBr_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

function sortMonthLabel_(a, b) {
  if (a === 'Sem data') return 1;
  if (b === 'Sem data') return -1;

  const pa = a.split('/');
  const pb = b.split('/');

  if (pa.length !== 2 || pb.length !== 2) return String(a).localeCompare(String(b), 'pt-BR');

  const da = new Date(Number(pa[1]), Number(pa[0]) - 1, 1).getTime();
  const db = new Date(Number(pb[1]), Number(pb[0]) - 1, 1).getTime();

  return da - db;
}
