// ====================== CODE.gs (ATUALIZADO) ======================

const SPREADSHEET_ID = '1Z6X27MIUZ87czGEKyZmJvnqxsarmSEmSt4ifYpp2yMk';  // ← Planilha que você enviou
const SHEET_NAME = 'Respostas ao formulário 1';
const MIN_REQUIRED_COLUMNS = 5;
const LOWERCASE_CONNECTORS = new Set(['de', 'da', 'do', 'das', 'dos', 'e', 'em', 'na', 'no', 'nas', 'nos', 'a', 'o']);

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Análise de Higiene das Mãos - SCIH/HUC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ====================== FUNÇÃO PRINCIPAL ======================
function getAnalysis(filters = {}) {
  try {
    const records = loadNormalizedRecords();
    const filteredRecords = applyFilters(records, filters);
    
    const kpis = buildKPIs(filteredRecords);
    const chartData = buildChartData(filteredRecords);
    const insights = buildInsights(kpis, filteredRecords, records);
    const tableData = prepareTableData(filteredRecords);
    const filterOptions = getFilterOptions(records);

    return {
      success: true,
      kpis: kpis,
      chartData: chartData,
      insights: insights,
      tableData: tableData,
      filterOptions: filterOptions,
      totalProcessed: records.length
    };
  } catch (error) {
    console.error('Erro no getAnalysis:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

function loadNormalizedRecords() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(`Aba "${SHEET_NAME}" não encontrada na planilha informada.`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastColumn = Math.max(MIN_REQUIRED_COLUMNS, sheet.getLastColumn());
  const rawData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  return normalizeRecords(rawData);
}

// ====================== NORMALIZAR REGISTROS ======================
function normalizeRecords(rawData) {
  if (rawData.length < 2) return [];
  
  const records = [];
  
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    const timestamp = row[0] ? new Date(row[0]) : null;
    if (!timestamp) continue;
    
    let unidade = String(row[1] || '').trim();
    let categoria = String(row[2] || '').trim();
    let momento = String(row[3] || '').trim();
    let acaoRaw = String(row[4] || '').trim();
    
    unidade = toDisplayCase(unidade);
    categoria = toDisplayCase(categoria);
    momento = toDisplayCase(momento);
    
    const acao = acaoRaw || 'Não informado';
    
    const actionClassification = classifyAction(acaoRaw);
    const situacao = actionClassification.situacao;
    const metodo = actionClassification.metodo;
    
    const preenchimento = (unidade !== 'Não informado' && 
                           categoria !== 'Não informado' && 
                           momento !== 'Não informado' && 
                           acao !== 'Não informado') ? 'Completo' : 'Incompleto';
    
    records.push({
      data: timestamp,
      unidade: unidade,
      categoria: categoria,
      momento: momento,
      acao: acao,
      situacao: situacao,
      metodo: metodo,
      preenchimento: preenchimento
    });
  }
  return records;
}

function toDisplayCase(str) {
  if (!str) return 'Não informado';

  const normalized = String(str)
    .replace(/\s+/g, ' ')
    .trim()
    .toLocaleLowerCase('pt-BR');

  if (!normalized) return 'Não informado';

  return normalized
    .split(' ')
    .map((word, idx) => {
      if (!word) return word;
      if (LOWERCASE_CONNECTORS.has(word) && idx > 0) return word;
      return word.charAt(0).toLocaleUpperCase('pt-BR') + word.slice(1);
    })
    .join(' ');
}

// ====================== CLASSIFICAÇÃO ======================
function normalizeActionText(acao) {
  return String(acao || '')
    .trim()
    .toLocaleLowerCase('pt-BR')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function classifyAction(acao) {
  const normalized = normalizeActionText(acao);

  if (!normalized) {
    return { situacao: 'Incompleto', metodo: 'Não informado' };
  }

  const isNotDone = normalized.includes('nao realizado') ||
    normalized.includes('nao foi') ||
    normalized.includes('omit') ||
    normalized.includes('recus');

  if (isNotDone) {
    return { situacao: 'Não realizado', metodo: 'Não realizado' };
  }

  const usedSoapAndWater = normalized.includes('agua') ||
    normalized.includes('sabao') ||
    normalized.includes('sabonete') ||
    normalized.includes('lavagem') ||
    normalized.includes('lavou');

  if (usedSoapAndWater) {
    return { situacao: 'Realizado', metodo: 'Água e sabonete' };
  }

  const usedAlcohol = normalized.includes('alcool') ||
    normalized.includes('friccao') ||
    normalized.includes('gel');

  if (usedAlcohol) {
    return { situacao: 'Realizado', metodo: 'Fricção com álcool' };
  }

  if (normalized.includes('realizado')) {
    return { situacao: 'Realizado', metodo: 'Realizado sem método detalhado' };
  }

  return { situacao: 'Incompleto', metodo: 'Não informado' };
}

// ====================== APLICAR FILTROS ======================
function applyFilters(records, filters) {
  const startDate = filters.dataInicio ? new Date(filters.dataInicio) : null;
  const endDate = filters.dataFim ? new Date(filters.dataFim) : null;

  if (endDate) {
    endDate.setHours(23, 59, 59, 999);
  }

  const units = filters.unidades?.length ? new Set(filters.unidades) : null;
  const categories = filters.categorias?.length ? new Set(filters.categorias) : null;
  const moments = filters.momentos?.length ? new Set(filters.momentos) : null;
  const situations = filters.situacoes?.length ? new Set(filters.situacoes) : null;
  const methods = filters.metodos?.length ? new Set(filters.metodos) : null;
  const fillStatus = filters.preenchimentos?.length ? new Set(filters.preenchimentos) : null;

  return records.filter(record => {
    if (startDate && record.data < startDate) return false;
    if (endDate && record.data > endDate) return false;
    if (units && !units.has(record.unidade)) return false;
    if (categories && !categories.has(record.categoria)) return false;
    if (moments && !moments.has(record.momento)) return false;
    if (situations && !situations.has(record.situacao)) return false;
    if (methods && !methods.has(record.metodo)) return false;
    if (fillStatus && !fillStatus.has(record.preenchimento)) return false;
    
    return true;
  });
}

// ====================== CONSTRUIR KPIs ======================
function buildKPIs(filtered) {
  const total = filtered.length;
  if (total === 0) {
    return {
      totalObservacoes: 0,
      totalRealizado: 0,
      totalNaoRealizado: 0,
      totalIncompleto: 0,
      adesaoGeral: 0,
      taxaNaoRealizacao: 0,
      taxaPreenchimentoCompleto: 0,
      taxaPreenchimentoIncompleto: 0
    };
  }

  let realizado = 0;
  let naoRealizado = 0;
  let incompleto = 0;
  let completos = 0;

  for (const record of filtered) {
    if (record.situacao === 'Realizado') realizado++;
    else if (record.situacao === 'Não realizado') naoRealizado++;
    else incompleto++;

    if (record.preenchimento === 'Completo') completos++;
  }
  
  const adesaoGeral = Math.round((realizado / total) * 100) || 0;
  const taxaNaoRealizacao = Math.round((naoRealizado / total) * 100) || 0;
  const taxaPreenchimentoCompleto = Math.round((completos / total) * 100) || 0;
  
  return {
    totalObservacoes: total,
    totalRealizado: realizado,
    totalNaoRealizado: naoRealizado,
    totalIncompleto: incompleto,
    adesaoGeral: adesaoGeral,
    taxaNaoRealizacao: taxaNaoRealizacao,
    taxaPreenchimentoCompleto: taxaPreenchimentoCompleto,
    taxaPreenchimentoIncompleto: 100 - taxaPreenchimentoCompleto
  };
}

// ====================== CONSTRUIR GRÁFICOS ======================
function buildChartData(filtered) {
  return {
    temporal: buildTemporalData(filtered),
    porUnidade: buildGroupData(filtered, 'unidade'),
    porCategoria: buildGroupData(filtered, 'categoria'),
    porMomento: buildGroupData(filtered, 'momento'),
    distribuicaoSituacao: buildPieData(filtered, 'situacao'),
    distribuicaoMetodo: buildPieData(filtered, 'metodo')
  };
}

function buildTemporalData(filtered) {
  const groups = {};
  const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  filtered.forEach(r => {
    const key = `${r.data.getFullYear()}-${String(r.data.getMonth() + 1).padStart(2, '0')}`;
    if (!groups[key]) groups[key] = { total: 0, realizado: 0 };
    groups[key].total++;
    if (r.situacao === 'Realizado') groups[key].realizado++;
  });
  
  const labels = Object.keys(groups).sort();
  const formattedLabels = labels.map(key => {
    const [year, month] = key.split('-');
    const monthIndex = Number(month) - 1;
    const monthName = monthNames[monthIndex] || month;
    return `${monthName}/${year.slice(-2)}`;
  });

  return {
    labels: formattedLabels,
    adesao: labels.map(k => groups[k].total ? Math.round((groups[k].realizado / groups[k].total) * 100) : 0)
  };
}

function buildGroupData(filtered, key) {
  const groups = {};
  filtered.forEach(r => {
    const val = r[key] || 'Não informado';
    if (!groups[val]) groups[val] = { total: 0, realizado: 0 };
    groups[val].total++;
    if (r.situacao === 'Realizado') groups[val].realizado++;
  });
  
  const labels = Object.keys(groups).sort();
  return {
    labels: labels,
    adesaoPercent: labels.map(l => groups[l].total ? Math.round((groups[l].realizado / groups[l].total) * 100) : 0)
  };
}

function buildPieData(filtered, key) {
  const counts = {};
  filtered.forEach(r => {
    counts[r[key]] = (counts[r[key]] || 0) + 1;
  });

  let labels;
  if (key === 'situacao') {
    const situacaoOrder = ['Realizado', 'Não realizado', 'Incompleto'];
    labels = situacaoOrder.filter(label => counts[label] !== undefined);
  } else if (key === 'metodo') {
    const metodoOrder = ['Água e sabonete', 'Fricção com álcool', 'Não realizado', 'Realizado sem método detalhado', 'Não informado'];
    labels = metodoOrder.filter(label => counts[label] !== undefined);
  } else {
    labels = Object.keys(counts);
  }

  return {
    labels: labels,
    data: labels.map(l => counts[l])
  };
}

// ====================== INSIGHTS ======================
function buildInsights(kpis, filtered, allRecords) {
  const list = [];
  
  if (kpis.adesaoGeral >= 85) list.push('Excelente adesão institucional. Acima do padrão COSEP/OMS.');
  else if (kpis.adesaoGeral >= 70) list.push('Adesão boa. Dentro das metas institucionais.');
  else if (kpis.adesaoGeral >= 50) list.push('Adesão moderada. Recomenda-se ação corretiva.');
  else list.push('Adesão crítica. Intervenção urgente necessária.');
  
  list.push(`Foram processados <strong>${kpis.totalObservacoes}</strong> registros válidos.`);
  
  if (filtered.length > 0) {
    const unitData = buildGroupData(filtered, 'unidade');
    if (unitData.labels.length > 0) {
      const maxIdx = unitData.adesaoPercent.indexOf(Math.max(...unitData.adesaoPercent));
      list.push(`Melhor unidade: <strong>${unitData.labels[maxIdx]}</strong> (${unitData.adesaoPercent[maxIdx]}%)`);
    }
  }
  
  if (kpis.taxaPreenchimentoCompleto < 85) {
    list.push('Atenção: Alta taxa de preenchimento incompleto. Verificar qualidade do formulário.');
  }
  
  list.push(`Taxa de não realização: <strong>${kpis.taxaNaoRealizacao}%</strong>`);
  
  return list;
}

// ====================== TABELA ======================
function prepareTableData(filteredRecords) {
  return filteredRecords
    .sort((a, b) => b.data.getTime() - a.data.getTime())
    .map(r => ({
      data: r.data.toLocaleDateString('pt-BR') + ' ' + r.data.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' }),
      unidade: r.unidade,
      categoriaProfissional: r.categoria,
      momento: r.momento,
      acao: r.acao,
      situacao: r.situacao,
      metodo: r.metodo,
      preenchimento: r.preenchimento
    }));
}

// ====================== FILTROS ======================
function getFilterOptions(records) {
  const getUnique = (field) => [...new Set(records.map(r => r[field]))]
    .filter(v => v && v !== 'Não informado')
    .sort();
  
  return {
    unidades: getUnique('unidade'),
    categorias: getUnique('categoria'),
    momentos: getUnique('momento')
  };
}
