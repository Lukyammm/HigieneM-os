# HigieneM-os - Monitoramento de Higiene das Mãos

Sistema web em **Google Apps Script + Google Sheets** para transformar respostas de formulário em um painel operacional de adesão à higiene das mãos, com KPIs, filtros, gráficos e tabela analítica para apoio à gestão de segurança do paciente.

## Fit de portfólio

Este projeto demonstra análise operacional em saúde, automação de leitura de planilhas, classificação de registros e construção de dashboard para tomada de decisão. Ele é relevante para vagas de BI, análise de dados, reporting e automação em ambientes com rotinas baseadas em formulários e Google Workspace.

**Palavras-chave:** healthcare analytics, patient safety, dashboard, KPI, Google Apps Script, Google Sheets, JavaScript, Chart.js, data quality, workflow automation.

## Problema que o sistema resolve

Em rotinas hospitalares, o monitoramento de higiene das mãos costuma depender de planilhas extensas, consolidação manual e análises lentas. Isso dificulta identificar queda de adesão, comparar unidades, acompanhar categorias profissionais e agir rapidamente sobre momentos críticos.

O sistema centraliza os registros em um painel único, classifica automaticamente as ações observadas e apresenta indicadores para apoiar educação permanente, reforço de protocolo e gestão de risco assistencial.

## Principais funcionalidades

- Leitura automática de respostas armazenadas em Google Sheets.
- Normalização de registros para padronizar texto, datas e categorias.
- Classificação automática da ação observada:
  - Realizado
  - Não realizado
  - Incompleto
- Identificação do método utilizado:
  - Água e sabonete
  - Fricção com álcool
  - Não informado
- KPIs operacionais:
  - total de observações
  - adesão geral
  - taxa de não realização
  - completude do preenchimento
- Filtros combináveis por período, unidade, categoria profissional, momento e método.
- Gráficos de tendência e distribuição para leitura gerencial.
- Tabela analítica para conferência e rastreabilidade.

## Tecnologias utilizadas

- Google Apps Script para backend e publicação do Web App.
- Google Sheets como base operacional.
- HTML5 e JavaScript para interface e lógica cliente.
- Tailwind CSS para layout e componentes visuais.
- Chart.js e `chartjs-plugin-datalabels` para visualização de dados.

## Estrutura do projeto

```text
HigieneM-os/
├── Code.gs         # Backend: leitura, tratamento, filtros e métricas
├── index.html      # Frontend: dashboard, filtros, gráficos e tabela
├── README.md       # Documentação técnica e de portfólio
└── PORTFOLIO.md    # Versão estratégica para apresentação
```

## Fluxo de funcionamento

1. Observações são coletadas via formulário institucional.
2. As respostas chegam na aba de dados do Google Sheets.
3. `Code.gs` lê os registros e aplica limpeza, padronização, classificação e cálculo de KPIs.
4. `index.html` renderiza o dashboard com filtros e visões analíticas.
5. O usuário aplica filtros e recebe resultados atualizados para análise operacional.

## Evidências técnicas

- Separação entre backend Apps Script e interface HTML/JavaScript.
- Transformação de dados brutos de formulário em métricas de gestão.
- Regras de classificação para reduzir análise manual.
- Filtros combináveis para investigação por unidade, período e categoria.
- Uso de visualizações para apoiar comparação e acompanhamento histórico.

## Como executar

### Pré-requisitos

- Conta Google com acesso ao Apps Script e Google Sheets.
- Planilha com estrutura compatível para timestamp, unidade, categoria, momento e ação observada.

### Passo a passo

1. Crie um projeto no Google Apps Script.
2. Copie o conteúdo de `Code.gs` para o arquivo `.gs` do projeto.
3. Copie o conteúdo de `index.html` para um arquivo HTML no projeto.
4. Ajuste as constantes no `Code.gs`:
   - `SPREADSHEET_ID`
   - `SHEET_NAME`
5. Clique em **Implantar > Nova implantação > Aplicativo da Web**.
6. Defina permissões de acesso conforme sua política institucional.
7. Abra a URL gerada e valide os dados no dashboard.

## Melhorias futuras

- Controle de acesso por perfil (gestão, coordenação, auditoria).
- Exportação de relatórios em PDF/Excel por filtro aplicado.
- Alertas automáticos para queda de adesão por unidade.
- Série histórica com metas mensais e comparação com benchmark.
- Módulo de plano de ação com acompanhamento de status.
- Testes automatizados para regras de classificação e cálculo de KPIs.

## Privacidade

Este repositório deve ser usado com dados fictícios, sintéticos ou anonimizados. Qualquer planilha real deve remover identificadores de pacientes, profissionais, instituições e registros sensíveis antes de uso público.

## Autor

Projeto de portfólio por **Lukyam Madeira**, com foco em análise operacional de saúde, indicadores de segurança do paciente e automação de dashboards em Google Workspace.
