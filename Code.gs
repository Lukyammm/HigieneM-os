// ====================== CODE.gs (ATUALIZADO) ======================

const SPREADSHEET_ID = '1Z6X27MIUZ87czGEKyZmJvnqxsarmSEmSt4ifYpp2yMk';  // ← Planilha que você enviou
const SHEET_NAME = 'Respostas ao formulário 1';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Análise de Higiene das Mãos • COSEP')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ====================== FUNÇÃO PRINCIPAL ======================
function getAnalysis(filters = {}) {
  try {
    // Abre a planilha específica pelo ID (mais seguro e confiável)
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Aba "${SHEET_NAME}" não encontrada na planilha informada.`);
    }

    const rawData = sheet.getDataRange().getValues();
    const records = normalizeRecords(rawData);
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
    
    const toTitleCase = (str) => {
      if (!str) return 'Não informado';
      return str.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
    };
    
    unidade = toTitleCase(unidade);
    categoria = toTitleCase(categoria);
    momento = toTitleCase(momento);
    
    const acao = acaoRaw || 'Não informado';
    
    const situacao = determineSituacao(acaoRaw);
    const metodo = determineMetodo(acaoRaw);
    
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

// ====================== CLASSIFICAÇÃO ======================
function determineSituacao(acao) {
  if (!acao || acao.trim() === '') return 'Incompleto';
  
  const lower = acao.toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  
  if (lower.includes('nao realizado') || lower.includes('não realizado') || 
      lower.includes('nao foi') || lower.includes('não foi') || 
      lower.includes('omit') || lower.includes('recus')) {
    return 'Não realizado';
  }
  
  if (lower.includes('realizado')) {
    return 'Realizado';
  }
  
  return 'Incompleto';
}

function determineMetodo(acao) {
  if (!acao || acao.trim() === '') return 'Não informado';
  
  const lower = acao.toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  
  if (lower.includes('nao realizado') || lower.includes('não realizado')) {
    return 'Não realizado';
  }
  
  if (lower.includes('agua') || lower.includes('água') || 
      lower.includes('sabao') || lower.includes('sabonete') || 
      lower.includes('lavagem') || lower.includes('lavou')) {
    return 'Água e sabonete';
  }
  
  if (lower.includes('alcool') || lower.includes('álcool') || 
      lower.includes('friccao') || lower.includes('fricção') || 
      lower.includes('gel')) {
    return 'Fricção com álcool';
  }
  
  if (lower.includes('realizado')) {
    return 'Realizado sem método detalhado';
  }
  
  return 'Não informado';
}

// ====================== APLICAR FILTROS ======================
function applyFilters(records, filters) {
  return records.filter(record => {
    if (filters.dataInicio) {
      if (record.data < new Date(filters.dataInicio)) return false;
    }
    if (filters.dataFim) {
      const fim = new Date(filters.dataFim);
      fim.setHours(23, 59, 59, 999);
      if (record.data > fim) return false;
    }
    
    const check = (list, value) => !list || list.length === 0 || list.includes(value);
    
    if (!check(filters.unidades, record.unidade)) return false;
    if (!check(filters.categorias, record.categoria)) return false;
    if (!check(filters.momentos, record.momento)) return false;
    if (!check(filters.situacoes, record.situacao)) return false;
    if (!check(filters.metodos, record.metodo)) return false;
    if (!check(filters.preenchimentos, record.preenchimento)) return false;
    
    return true;
  });
}

// ====================== CONSTRUIR KPIs ======================
function buildKPIs(filtered) {
  const total = filtered.length;
  if (total === 0) return { totalObservacoes: 0, totalRealizado: 0, totalNaoRealizado: 0, totalIncompleto: 0, adesaoGeral: 0, taxaNaoRealizacao: 0, taxaPreenchimentoCompleto: 0, taxaPreenchimentoIncompleto: 0 };
  
  const realizado = filtered.filter(r => r.situacao === 'Realizado').length;
  const naoRealizado = filtered.filter(r => r.situacao === 'Não realizado').length;
  const incompleto = filtered.filter(r => r.situacao === 'Incompleto').length;
  
  const adesaoGeral = Math.round((realizado / total) * 100) || 0;
  const taxaNaoRealizacao = Math.round((naoRealizado / total) * 100) || 0;
  
  const completos = filtered.filter(r => r.preenchimento === 'Completo').length;
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
  filtered.forEach(r => {
    const key = `${r.data.getFullYear()}-${String(r.data.getMonth() + 1).padStart(2, '0')}`;
    if (!groups[key]) groups[key] = { total: 0, realizado: 0 };
    groups[key].total++;
    if (r.situacao === 'Realizado') groups[key].realizado++;
  });
  
  const labels = Object.keys(groups).sort();
  return {
    labels: labels,
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
  const labels = Object.keys(counts);
  return {
    labels: labels,
    data: labels.map(l => counts[l])
  };
}

// ====================== INSIGHTS ======================
function buildInsights(kpis, filtered, allRecords) {
  const list = [];
  
  if (kpis.adesaoGeral >= 85) list.push('🎯 Excelente adesão institucional! Acima do padrão COSEP/OMS.');
  else if (kpis.adesaoGeral >= 70) list.push('✅ Adesão boa. Dentro das metas institucionais.');
  else if (kpis.adesaoGeral >= 50) list.push('⚠️ Adesão moderada. Recomenda-se ação corretiva.');
  else list.push('🚨 Adesão crítica. Intervenção urgente necessária.');
  
  list.push(`Foram processados <strong>${kpis.totalObservacoes}</strong> registros válidos.`);
  
  if (filtered.length > 0) {
    const unitData = buildGroupData(filtered, 'unidade');
    if (unitData.labels.length > 0) {
      const maxIdx = unitData.adesaoPercent.indexOf(Math.max(...unitData.adesaoPercent));
      list.push(`🏆 Melhor unidade: <strong>${unitData.labels[maxIdx]}</strong> (${unitData.adesaoPercent[maxIdx]}%)`);
    }
  }
  
  if (kpis.taxaPreenchimentoCompleto < 85) {
    list.push('📝 Atenção: Alta taxa de preenchimento incompleto. Verificar qualidade do formulário.');
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
