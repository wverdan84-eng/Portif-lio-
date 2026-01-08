/**
 * SISTEMA DE GESTÃO DE INVESTIMENTOS PREMIUM
 * Versão: 2.0 Professional
 */

function criarPlanilhaInvestimentoPremium() {
  try {
    console.log('🚀 Iniciando criação da Planilha Premium...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Verificar se a planilha está vazia ou já existe
    if (ss.getSheets().length > 1) {
      const response = ui.alert(
        'Atenção',
        'Esta planilha já contém dados. Deseja criar uma nova estrutura?',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) {
        ui.alert('Operação cancelada pelo usuário.');
        return;
      }
    }
    
    // Criar todas as abas necessárias
    criarAbas(ss);
    
    // Obter referências das abas
    const dash = ss.getSheetByName("📈 DASHBOARD");
    const lanc = ss.getSheetByName("💰 LANÇAMENTOS");
    const metas = ss.getSheetByName("🎯 METAS");
    const analise = ss.getSheetByName("📊 ANÁLISE");
    const carteira = ss.getSheetByName("💼 CARTEIRA");
    const dividendos = ss.getSheetByName("📈 DIVIDENDOS");
    const alertas = ss.getSheetByName("🚨 ALERTAS");
    
    // Configurar cada aba
    configurarAbaLancamentos(lanc);
    configurarDashboard(dash, lanc, ss);
    configurarMetas(metas);
    configurarAnalise(analise, lanc);
    configurarCarteira(carteira, lanc);
    configurarDividendos(dividendos);
    configurarAlertas(alertas, lanc);
    
    // Aplicar formatação final
    aplicarFormatacaoFinal(ss);
    
    // Criar relatório de conclusão
    criarRelatorioConclusao(ss);
    
    console.log('✅ Sistema criado com sucesso!');
    mostrarMensagemSucesso(ui);
    
  } catch (error) {
    console.error('❌ Erro ao criar planilha:', error);
    SpreadsheetApp.getUi().alert(
      'Erro no Sistema',
      'Ocorreu um erro: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function criarAbas(ss) {
  const abasConfig = [
    { nome: "📈 DASHBOARD", cor: "#1c4587" },
    { nome: "💰 LANÇAMENTOS", cor: "#0b5394" },
    { nome: "🎯 METAS", cor: "#674ea7" },
    { nome: "📊 ANÁLISE", cor: "#45818e" },
    { nome: "💼 CARTEIRA", cor: "#3c78d8" },
    { nome: "📈 DIVIDENDOS", cor: "#6aa84f" },
    { nome: "🚨 ALERTAS", cor: "#cc0000" },
    { nome: "⚙️ CONFIG", cor: "#666666" }
  ];
  
  // Limpar abas existentes (exceto a primeira)
  const sheets = ss.getSheets();
  for (let i = sheets.length - 1; i > 0; i--) {
    ss.deleteSheet(sheets[i]);
  }
  
  // Criar novas abas
  abasConfig.forEach((aba, index) => {
    const sheet = ss.insertSheet(aba.nome, index + 1);
    sheet.setTabColor(aba.cor);
    if (aba.nome === "📈 DASHBOARD") {
      sheet.setHiddenGridlines(true);
    }
  });
}

function configurarAbaLancamentos(sheet) {
  const headers = [
    ["DATA", "TICKER", "ATIVO", "TIPO", "QUANTIDADE", 
     "PREÇO UNIT.", "TOTAL INVESTIDO", "CUSTOS", "TOTAL LÍQUIDO", 
     "CATEGORIA", "CORRETORA", "NOTA FISCAL", "OBSERVAÇÕES"]
  ];
  
  const headerRange = sheet.getRange("A1:M1");
  headerRange.setValues(headers)
    .setBackground("#1c4587")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(11);
  
  const colWidths = [100, 80, 180, 120, 90, 110, 120, 90, 120, 120, 130, 120, 200];
  colWidths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  
  aplicarValidacoesLancamentos(sheet);
  adicionarDadosExemplo(sheet);
  sheet.setFrozenRows(1);
  aplicarFormatacaoCondicionalLancamentos(sheet);
}

function aplicarValidacoesLancamentos(sheet) {
  const dataValidation = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setHelpText("Insira uma data válida")
    .build();
  sheet.getRange("A2:A1000").setDataValidation(dataValidation);
  
  const tipoValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["AÇÃO", "FII", "ETF", "RENDA FIXA", "CRYPTO", "FUNDO IMOBILIÁRIO", "EXTERIOR", "OUTROS"])
    .build();
  sheet.getRange("D2:D1000").setDataValidation(tipoValidation);
  
  const categoriaValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["VARIÁVEL", "FIXA", "EXTERIOR", "IMOBILIÁRIO", "CRIPTO", "RESERVA"])
    .build();
  sheet.getRange("J2:J1000").setDataValidation(categoriaValidation);
  
  sheet.getRange("E2:E1000").setNumberFormat("#,##0");
  sheet.getRange("F2:I1000").setNumberFormat("R$ #,##0.00");
  sheet.getRange("A2:A1000").setNumberFormat("dd/mm/yyyy");
}
function adicionarDadosExemplo(sheet) {
  const dadosExemplo = [
    [new Date(), "PETR4", "Petróleo Brasileiro S/A", "AÇÃO", 100, 36.50, "=E2*F2", 0.50, "=G2-H2", "VARIÁVEL", "XP Investimentos", "NF-001", "Compra inicial"],
    [new Date(), "ITUB4", "Itaú Unibanco", "AÇÃO", 50, 32.80, "=E3*F3", 0.45, "=G3-H3", "VARIÁVEL", "BTG Pactual", "NF-002", "Aumento posição"],
    [new Date(), "BOVA11", "ETF IBOVESPA", "ETF", 20, 112.30, "=E4*F4", 0.30, "=G4-H4", "VARIÁVEL", "Clear", "NF-003", "Diversificação"],
    [new Date(), "TS", "Tesouro Selic 2029", "RENDA FIXA", 1, 9850.00, "=E5*F5", 0, "=G5-H5", "FIXA", "Nubank", "NF-004", "Reserva segurança"],
    [new Date(), "IVVB11", "ETF S&P 500", "ETF", 15, 282.40, "=E6*F6", 0.60, "=G6-H6", "EXTERIOR", "Rico", "NF-005", "Exposição EUA"],
    [new Date(), "MXRF11", "Fundo Imobiliário", "FII", 30, 10.25, "=E7*F7", 0.20, "=G7-H7", "IMOBILIÁRIO", "Órama", "NF-006", "Renda passiva"]
  ];
  
  const range = sheet.getRange("A2:M7");
  range.setValues(dadosExemplo);
  range.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  
  for (let i = 2; i <= 7; i++) {
    const bgColor = i % 2 === 0 ? "#FFFFFF" : "#F8F9FA";
    sheet.getRange(`A${i}:M${i}`).setBackground(bgColor);
  }
}

function aplicarFormatacaoCondicionalLancamentos(sheet) {
  const regraValorAlto = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(10000)
    .setBackground("#FFF2CC")
    .setFontColor("#E69138")
    .setBold(true)
    .setRanges([sheet.getRange("I2:I1000")])
    .build();
  
  const regraCustos = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(10)
    .setBackground("#F4CCCC")
    .setFontColor("#CC0000")
    .setRanges([sheet.getRange("H2:H1000")])
    .build();
  
  sheet.setConditionalFormatRules([regraValorAlto, regraCustos]);
}

function configurarDashboard(sheet, lancSheet, ss) {
  sheet.clear();
  sheet.getRange("A1:Z100").setBackground("#F8F9FA");
  
  sheet.getRange("B1").setValue("DASHBOARD DE INVESTIMENTOS - SISTEMA PREMIUM")
    .setFontSize(20)
    .setFontWeight("bold")
    .setFontColor("#1c4587")
    .setHorizontalAlignment("center");
  sheet.getRange("B1:L1").merge();
  
  sheet.getRange("B2").setValue(`Gerado em: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}`)
    .setFontSize(10)
    .setFontColor("#666666");
  sheet.getRange("B2:L2").merge();
  
  criarCardKPI(sheet, 4, 2, "PATRIMÔNIO TOTAL", "=SUM('💰 LANÇAMENTOS'!I:I)", "#1c4587", "R$ #,##0.00");
  criarCardKPI(sheet, 4, 5, "INVESTIDO MÊS", "=SUMIF('💰 LANÇAMENTOS'!A:A,\">=\"&EOMONTH(TODAY(),-1)+1,'💰 LANÇAMENTOS'!I:I)", "#0b5394", "R$ #,##0.00");
  criarCardKPI(sheet, 4, 8, "RENTABILIDADE", "=IFERROR(F4/C4,0)", "#674ea7", "0.00%");
  criarCardKPI(sheet, 4, 11, "DIVERSIFICAÇÃO", "=COUNTA(UNIQUE('💰 LANÇAMENTOS'!B:B))-1", "#6aa84f", "#,##0 ativos");
  
  criarCardKPI(sheet, 8, 2, "MELHOR ATIVO", "=INDEX('💰 LANÇAMENTOS'!B:B,MATCH(MAX('💰 LANÇAMENTOS'!I:I),'💰 LANÇAMENTOS'!I:I,0))", "#3c78d8", "");
  criarCardKPI(sheet, 8, 5, "CUSTOS TOTAIS", "=SUM('💰 LANÇAMENTOS'!H:H)", "#cc0000", "R$ #,##0.00");
  criarCardKPI(sheet, 8, 8, "DIVIDENDOS MÊS", "='📈 DIVIDENDOS'!B2", "#45818e", "R$ #,##0.00");
  criarCardKPI(sheet, 8, 11, "METAS ATINGIDAS", "=COUNTIF('🎯 METAS'!F:F,\">=1\")", "#FF9900", "#,##0 / #,##0");
  
  criarGraficoAlocacao(sheet, lancSheet, 12, 2);
  criarTabelaTopAtivos(sheet, lancSheet, 28, 2);
  criarTabelaUltimosLancamentos(sheet, lancSheet, 28, 9);
  criarLegenda(sheet, 35, 2);
}

function criarCardKPI(sheet, linha, coluna, titulo, formula, cor, formato) {
  const rangeTitulo = sheet.getRange(linha, coluna, 1, 3);
  const rangeValor = sheet.getRange(linha + 1, coluna, 2, 3);
  
  rangeTitulo.setValue(titulo)
    .setBackground(cor)
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(11);
  rangeTitulo.merge();
  
  rangeValor.merge()
    .setFormula(formula)
    .setBackground("#FFFFFF")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "#DDDDDD")
    .setNumberFormat(formato);
}

function criarGraficoAlocacao(sheet, lancSheet, linha, coluna) {
  sheet.getRange(linha, coluna).setValue("DISTRIBUIÇÃO DA CARTEIRA")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#1c4587");
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(lancSheet.getRange("J2:J100"))
    .addRange(lancSheet.getRange("I2:I100"))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)
    .setNumHeaders(1)
    .setOption('legend.position', 'right')
    .setOption('colors', ['#1c4587', '#0b5394', '#674ea7', '#45818e', '#3c78d8'])
    .setOption('backgroundColor', '#F8F9FA')
    .setPosition(linha + 1, coluna, 0, 0)
    .setOption('width', 450)
    .setOption('height', 300)
    .build();
  
  sheet.insertChart(chart);
}

function criarTabelaTopAtivos(sheet, lancSheet, linha, coluna) {
  sheet.getRange(linha, coluna).setValue("TOP 5 ATIVOS MAIS VALIOSOS")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#1c4587");
  
  const headers = [["POSIÇÃO", "ATIVO", "VALOR INVESTIDO", "% CARTEIRA"]];
  sheet.getRange(linha + 1, coluna, 1, 4).setValues(headers)
    .setBackground("#1c4587")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  
  for (let i = 0; i < 5; i++) {
    const row = linha + 2 + i;
    sheet.getRange(row, coluna).setValue(i + 1);
    sheet.getRange(row, coluna + 1).setFormula(
      `=INDEX('💰 LANÇAMENTOS'!B:B, MATCH(LARGE('💰 LANÇAMENTOS'!I:I, ${i + 1}), '💰 LANÇAMENTOS'!I:I, 0))`
    );
    sheet.getRange(row, coluna + 2).setFormula(
      `=LARGE('💰 LANÇAMENTOS'!I:I, ${i + 1})`
    );
    sheet.getRange(row, coluna + 3).setFormula(
      `=IFERROR(${sheet.getRange(row, coluna + 2).getA1Notation()}/$C$5, 0)`
    );
  }
  
  const tabelaRange = sheet.getRange(linha + 1, coluna, 7, 4);
  tabelaRange.setBorder(true, true, true, true, true, true);
  
  sheet.getRange(linha + 2, coluna + 2, 5, 1).setNumberFormat("R$ #,##0.00");
  sheet.getRange(linha + 2, coluna + 3, 5, 1).setNumberFormat("0.00%");
  
  for (let i = 0; i < 5; i++) {
    const bgColor = i % 2 === 0 ? "#FFFFFF" : "#F8F9FA";
    sheet.getRange(linha + 2 + i, coluna, 1, 4).setBackground(bgColor);
  }
}