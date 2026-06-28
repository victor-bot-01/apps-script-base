// ===== Script 1 (Ajustado) =====
function limparDadosMesAtual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Mês Atual");
  
  if (!sheet) {
    console.log("A aba 'Mês Atual' não foi encontrada.");
    return;
  }

  var lastRow = sheet.getLastRow();
  var rangeF = sheet.getRange(2, 6, lastRow - 1);
  var valuesF = rangeF.getValues();
  for (var i = 0; i < valuesF.length; i++) {
    if (valuesF[i][0]) {
      valuesF[i][0] = valuesF[i][0].replace(/\(Qtd:\s*\d+\)/g, "").trim();
    }
  }
  rangeF.setValues(valuesF);

  var rangeG = sheet.getRange(2, 7, lastRow - 1); 
  var valuesG = rangeG.getValues();
  for (var j = 0; j < valuesG.length; j++) {
    if (valuesG[j][0] === "Standard") {
      valuesG[j][0] = "Envio Próprio";
    } else if (valuesG[j][0] === "Logistica Amazon Dba") {
      valuesG[j][0] = "Coleta";
    }
  }
  rangeG.setValues(valuesG);

  // === NOVA FUNÇÃO ADICIONADA ===
  preencherResponsavelPorNumeroPedido(sheet);

  console.log("Limpeza e substituições concluídas com sucesso!");
}

// ===== Função Nova =====
function preencherResponsavelPorNumeroPedido(sheet) {
  var lastRow = sheet.getLastRow();
  var colB = sheet.getRange(2, 2, lastRow - 1).getValues(); // Número do pedido
  var colG = sheet.getRange(2, 7, lastRow - 1).getValues(); // Responsável existente

  var mapaResponsavel = {};

  // Criar um mapa com os valores já definidos na coluna G
  for (var i = 0; i < colB.length; i++) {
    var pedido = colB[i][0];
    var responsavel = colG[i][0];
    if (pedido && responsavel && responsavel.toString().trim() !== "") {
      mapaResponsavel[pedido] = responsavel;
    }
  }

  // Preencher as células vazias da coluna G com base no mapa
  for (var j = 0; j < colB.length; j++) {
    var pedidoAtual = colB[j][0];
    if (pedidoAtual && (!colG[j][0] || colG[j][0].toString().trim() === "")) {
      if (mapaResponsavel[pedidoAtual]) {
        colG[j][0] = mapaResponsavel[pedidoAtual];
      }
    }
  }

  // Atualizar a planilha
  sheet.getRange(2, 7, lastRow - 1).setValues(colG);
}

// ===== Script 2 =====
function copiarLinhasNaoVerdesParaVendas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getSheetByName("Mês Atual");
  const abaDestino = ss.getSheetByName("Vendas");

  if (!abaOrigem || !abaDestino) {
    Logger.log("❌ Aba 'Mês Atual' ou 'Vendas' não encontrada.");
    return;
  }

  const primeiraLinhaDados = 2;
  const ultimaLinha = abaOrigem.getLastRow();
  const totalLinhas = ultimaLinha - primeiraLinhaDados + 1;

  if (totalLinhas <= 0) {
    Logger.log("ℹ️ Nenhuma linha com dados para processar.");
    return;
  }

  const intervaloColA = abaOrigem.getRange(primeiraLinhaDados, 1, totalLinhas);
  const coresColA = intervaloColA.getBackgrounds();
  const valoresColB = abaOrigem.getRange(primeiraLinhaDados, 2, totalLinhas).getValues();
  const intervaloBateG = abaOrigem.getRange(primeiraLinhaDados, 2, totalLinhas, 6).getValues();

  const dadosFiltrados = [];
  const linhasParaPintar = [];

  for (let i = 0; i < totalLinhas; i++) {
    const cor = coresColA[i][0];
    const valorColB = valoresColB[i][0];

    if (cor.toLowerCase() !== "#00ff00" && valorColB !== "" && valorColB !== null) {
      dadosFiltrados.push(intervaloBateG[i]);
      linhasParaPintar.push(primeiraLinhaDados + i);
    }
  }

  if (dadosFiltrados.length > 0) {
    let ultimaLinhaDestino = abaDestino.getLastRow();
    let proximaLinhaLivre = 1;

    if (ultimaLinhaDestino > 0) {
      const valorUltimaLinha = abaDestino.getRange(ultimaLinhaDestino, 1).getValue();
      proximaLinhaLivre = valorUltimaLinha === "" ? ultimaLinhaDestino : ultimaLinhaDestino + 1;
    }

    abaDestino.getRange(proximaLinhaLivre, 1, dadosFiltrados.length, 6).setValues(dadosFiltrados);

    linhasParaPintar.forEach(linha => {
      if (linha > 0) {
        abaOrigem.getRange(linha, 1).setBackground("#0000ff");
      }
    });

    Logger.log(`✅ ${dadosFiltrados.length} linha(s) copiadas e marcadas como processadas.`);
  } else {
    Logger.log("ℹ️ Nenhuma linha foi encontrada para copiar.");
  }
}

// ===== Script 3 =====
function ListadeProdutos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Vendas');
  var dadosSheet = ss.getSheetByName("Dados");

  if (!sheet) {
    Logger.log('❌ Aba "Vendas" não encontrada.');
    return;
  }
  if (!dadosSheet) {
    Logger.log('❌ Aba "Dados" não encontrada.');
    return;
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var range = sheet.getRange(1, 1, sheet.getLastRow(), 9);
  var values = range.getValues();

  let linhasReorganizadas = 0;
  let linhasParaLimparColB = [];

  for (var i = 0; i < values.length; i++) {
    const colIval = (values[i][8] || "").toString().trim();
    if (!colIval) {
      if (values[i][1] && values[i][5]) {
        values[i][5] = values[i][1] + ' - ' + values[i][5];
      } else if (values[i][1]) {
        values[i][5] = values[i][1];
      }

      let valorColunaA = values[i][0];
      if (typeof valorColunaA === 'number') {
        valorColunaA = `'${valorColunaA.toFixed(0)}`;
      }

      var tempRow = [
        values[i][4], values[i][1], values[i][3], values[i][2],
        valorColunaA, values[i][5], values[i][6], values[i][7], today
      ];

      values[i] = tempRow;
      linhasParaLimparColB.push(i + 1);
      linhasReorganizadas++;
    }
  }

  range.setValues(values);
  linhasParaLimparColB.forEach(function(linha) {
    sheet.getRange(linha, 2).clearContent();
  });

  var rangeF = sheet.getRange(1, 6, sheet.getLastRow());
  var valuesF = rangeF.getValues();
  var rowsToDelete = [];

  for (var i = 0; i < valuesF.length; i++) {
    var cellValue = valuesF[i][0];
    var text = String(cellValue ?? ""); // garante string (mesmo se for número/data/null)

    if (
      text.includes('Mercado Envios Full') ||
      text.includes('Cancelado') ||
      text.includes('Pedido fulfillment Magalu')
    ) {
      rowsToDelete.push(i + 1);
    }
  }
  for (var i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  var vendasData = sheet.getRange("A:A").getValues().flat();
  var vendasColumnB = sheet.getRange("B1:B" + sheet.getLastRow());
  for (var i = 0; i < vendasData.length; i++) {
    var item = vendasData[i];
    if (item) {
      var dadosValues = dadosSheet.getRange(1, 1, dadosSheet.getLastRow(), dadosSheet.getLastColumn()).getValues();
      for (var j = 0; j < dadosValues.length; j++) {
        for (var k = 0; k < dadosValues[j].length; k++) {
          if (dadosValues[j][k] === item && dadosValues[j][k + 1]) {
            vendasColumnB.getCell(i + 1, 1).setValue(dadosValues[j][k + 1]);
            break;
          }
        }
      }
    }
  }
}

// ===== Script 4 =====
function corrigirNotacaoCientificaColunaE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("Vendas");
  if (!aba) { Logger.log("❌ Aba 'Vendas' não encontrada."); return; }

  const ultimaLinha = aba.getLastRow();
  const colEValores = aba.getRange(1, 5, ultimaLinha).getValues();
  for (let i = 0; i < colEValores.length; i++) {
    const valor = colEValores[i][0];
    if (typeof valor === "number") {
      aba.getRange(i + 1, 5).setValue("'" + valor.toFixed(0));
    } else if (typeof valor === "string" && /e\+?\d+/i.test(valor)) {
      const num = Number(valor);
      if (!isNaN(num)) {
        aba.getRange(i + 1, 5).setValue("'" + num.toFixed(0));
      }
    }
  }
  Logger.log("✅ Conversão de notação científica finalizada na coluna E.");
}

// ===== Script 5 (Etiquetas corrigido) =====
function EtiquetasdeSeparacao() {
  var planilhaVendas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendas");
  var planilhaEtiquetas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Etiquetas");

  var dadosVendas = planilhaVendas.getRange("B:J").getValues();
  dadosVendas = dadosVendas.filter(r => r[0] && r[1] && r[2] && r[3] && r[4]);
  dadosVendas.sort((a, b) => b[3] - a[3]);

  var dadosEtiquetas = [];
  dadosVendas.forEach(function(row) {
    var valorI = row[7] instanceof Date ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), "dd/MM/yyyy") : row[7];
    for (var i = 0; i < row[1]; i++) {
      var combinacao = valorI + '\n' + row[3] + '\n' + row[4] + '\n' + row[2] + '\n' + row[0];
      dadosEtiquetas.push([combinacao]);
    }
  });

  // --- Correção para evitar erro quando não há dados ---
  planilhaEtiquetas.clearContents();
  if (dadosEtiquetas.length === 0) {
    Logger.log("⚠️ Nenhuma etiqueta gerada.");
    return;
  }

  planilhaEtiquetas.getRange(1, 1, dadosEtiquetas.length, 1).setValues(dadosEtiquetas);

  organizarCelulasEtiquetas();
  concatenarCelulasDuplicadasComIdentificadorCompleto();
  joinFirstTwoLinesInColumnA();
  adicionarParenteses();
}

function concatenarCelulasDuplicadasComIdentificadorCompleto() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Etiquetas');
  const data = sheet.getRange("A1:A" + sheet.getLastRow()).getValues();
  const grupos = {};
  data.forEach((linha, i) => {
    const celula = linha[0];
    if (!celula) return;
    const linhas = celula.split('\n');
    const chave = linhas.slice(0, 4).join('||').trim();
    if (!grupos[chave]) {
      grupos[chave] = { linhasExtras: [], linhaPrincipal: i + 1, conteudoPrincipal: celula };
    } else {
      const extras = linhas.slice(4).join('\n');
      if (extras) grupos[chave].linhasExtras.push(extras);
      sheet.getRange(i + 1, 1).clearContent();
    }
  });
  for (const chave in grupos) {
    let novoConteudo = grupos[chave].conteudoPrincipal;
    if (grupos[chave].linhasExtras.length > 0) {
      novoConteudo += '\n' + grupos[chave].linhasExtras.join('\n');
    }
    sheet.getRange(grupos[chave].linhaPrincipal, 1).setValue(novoConteudo);
  }
  const novaData = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  for (let r = novaData.length - 1; r >= 0; r--) {
    if (novaData[r][0] === "") sheet.deleteRow(r + 1);
  }
}

function organizarCelulasEtiquetas() {
  var planilhaEtiquetas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Etiquetas");
  var dadosEtiquetas = planilhaEtiquetas.getRange("A:A").getValues();
  for (var i = 0; i < dadosEtiquetas.length; i++) {
    if (dadosEtiquetas[i][0]) {
      planilhaEtiquetas.getRange(i + 1, 1).setValue(dadosEtiquetas[i][0].split(";").join("\n"));
    }
  }
}

function joinFirstTwoLinesInColumnA() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Etiquetas");
  var values = sheet.getRange("A:A").getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) {
      var lines = values[i][0].split("\n");
      if (lines.length > 1) {
        values[i][0] = lines[0].trim() + " - " + lines[1].trim() + "\n" + lines.slice(2).join("\n");
      }
    }
  }
  sheet.getRange("A:A").setValues(values);
}

function adicionarParenteses() {
  var planilhaEtiquetas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Etiquetas");
  var dadosEtiquetas = planilhaEtiquetas.getRange(1, 1, planilhaEtiquetas.getLastRow(), 1).getValues();
  for (var i = 0; i < dadosEtiquetas.length; i++) {
    if (dadosEtiquetas[i][0]) {
      var linhas = dadosEtiquetas[i][0].split("\n");
      for (var j = 3; j < linhas.length; j++) {
        linhas[j] = linhas[j] + " (      )";
      }
      dadosEtiquetas[i][0] = linhas.join("\n");
    }
  }
  planilhaEtiquetas.getRange(1, 1, dadosEtiquetas.length, 1).setValues(dadosEtiquetas);
}

// ===== Script 6 =====
function manterLinhasComECompartilhado() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendas");
  const lastRow = sheet.getLastRow();
  const dados = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const manterIndices = new Set();
  const valoresEComBvazio = new Set();

  dados.forEach((linha, index) => {
    const valorB = linha[1];
    const valorE = linha[4];
    if ((!valorB || valorB.toString().trim() === "") && valorE) {
      valoresEComBvazio.add(valorE.toString().trim());
      manterIndices.add(index);
    }
  });

  dados.forEach((linha, index) => {
    const valorE = linha[4];
    if (valorE && valoresEComBvazio.has(valorE.toString().trim())) {
      manterIndices.add(index);
    }
  });

  for (let i = dados.length - 1; i >= 0; i--) {
    if (!manterIndices.has(i)) {
      sheet.deleteRow(i + 1);
    }
  }
  Logger.log("Linhas desnecessárias foram removidas com sucesso.");
}

// ===== Script 7 =====
function excluirEtiquetasQueContemValoresDaColunaE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaVendas = ss.getSheetByName("Vendas");
  var abaEtiquetas = ss.getSheetByName("Etiquetas");

  // --- Verificação de dados ---
  if (!abaVendas || !abaEtiquetas) {
    Logger.log("❌ Abas necessárias não encontradas.");
    return;
  }
  if (abaVendas.getLastRow() === 0 || abaEtiquetas.getLastRow() === 0) {
    Logger.log("⚠️ Nenhum dado para processar em uma das abas.");
    return;
  }

  var valoresColunaE = abaVendas.getRange("E1:E" + abaVendas.getLastRow()).getValues().flat();
  var valoresEFiltrados = valoresColunaE.map(v => v.toString().trim()).filter(v => v !== "");

  // --- Se não houver nenhum valor na coluna E, não há motivo para remover ---
  if (valoresEFiltrados.length === 0) {
    Logger.log("⚠️ Nenhum valor válido na coluna E para comparação.");
    return;
  }

  var valoresEtiquetas = abaEtiquetas.getRange("A1:A" + abaEtiquetas.getLastRow()).getValues().flat();

  for (var i = valoresEtiquetas.length - 1; i >= 0; i--) {
    var valorLinha = valoresEtiquetas[i];
    if (!valorLinha) continue;
    var contemValor = valoresEFiltrados.some(v => valorLinha.includes(v));
    if (contemValor) abaEtiquetas.deleteRow(i + 1);
  }

  Logger.log("✅ Etiquetas removidas com sucesso com base na coluna E.");
}

// ===== Pipeline Completo =====
function executarPipelineCompleto() {
  limparDadosMesAtual();
  copiarLinhasNaoVerdesParaVendas();
  ListadeProdutos();
  corrigirNotacaoCientificaColunaE();
  EtiquetasdeSeparacao();
  manterLinhasComECompartilhado();
  excluirEtiquetasQueContemValoresDaColunaE();
  Logger.log("Pipeline completo executado com sucesso!");
}
