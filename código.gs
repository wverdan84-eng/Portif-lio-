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