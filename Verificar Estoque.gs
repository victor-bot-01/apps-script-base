function verificarEstoqueDescricoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const abaMesAtual = ss.getSheetByName("Mês Atual");
  if (!abaMesAtual) { Logger.log("❌ Aba 'Mês Atual' não encontrada."); return; }

  const abaPadrao = ss.getSheetByName("Padrão");
  if (!abaPadrao) { Logger.log("❌ Aba 'Padrão' não encontrada."); return; }

  const abaFormatar = ss.getSheetByName("Formatar");
  if (!abaFormatar) { Logger.log("❌ Aba 'Formatar' não encontrada."); return; }

  const ssDados = SpreadsheetApp.openById("15ueYlLK9JoLn_lXcP6JiFxs8Ds22KF3Km_FD2cCUMW4");
  const abaDados = ssDados.getSheetByName("Dados");
  if (!abaDados) { Logger.log("❌ Aba 'Dados' não encontrada na planilha externa."); return; }

  const ssInventario = SpreadsheetApp.openById("11yovl7B1uy6YxDjAAvoBfmeWv4t9hcNFibe5_UPA4D4");
  const abaInventario = ssInventario.getSheetByName("Inventário");
  if (!abaInventario) { Logger.log("❌ Aba 'Inventário' não encontrada na planilha externa."); return; }

  // Carrega todos os dados de uma vez
  const dadosValues = abaDados.getDataRange().getValues();
  const padraoValues = abaPadrao.getDataRange().getValues();

  const invLastRow = abaInventario.getLastRow();
  const invLastCol = abaInventario.getLastColumn();
  const inventarioValues = invLastRow > 0
    ? abaInventario.getRange(1, 1, invLastRow, invLastCol).getValues()
    : [];
  const inventarioBackgrounds = invLastRow > 0
    ? abaInventario.getRange(1, 1, 1, invLastCol).getBackgrounds()
    : [];

  // Palavras que devem ser ignoradas na verificação (singular e plural)
  const palavrasIgnorar = [
    "pingente", "pingentes",
    "colar", "colares",
    "anel", "anéis", "aneis",
    "pulseira", "pulseiras",
    "corrente", "correntes",
    "brinco", "brincos",
    "berloque", "berloques",
    "gema", "gemas",
    "larimar", "larimares"
  ];

  function deveIgnorar(descricao) {
    const desc = descricao.toLowerCase();
    return palavrasIgnorar.some(palavra => new RegExp("\\b" + palavra + "\\b").test(desc));
  }

  // Coluna F a partir da linha 2
  const lastRow = abaMesAtual.getLastRow();
  if (lastRow < 2) { Logger.log("⚠️ Nenhum dado na aba 'Mês Atual'."); return; }
  const colF = abaMesAtual.getRange(2, 6, lastRow - 1, 1).getValues();

  // Itens já existentes em "Formatar" para evitar duplicatas
  const ultimaLinhaFormatar = abaFormatar.getLastRow();
  const existingFormatar = new Set(
    ultimaLinhaFormatar > 0
      ? abaFormatar.getRange(1, 1, ultimaLinhaFormatar, 1).getValues().flat().map(v => v.toString().trim())
      : []
  );
  const novosFormatar = [];

  // Busca valor em array 2D (qualquer coluna); prioriza coluna com azul na linha 1 se pedido
  function findInValues(values, searchVal, backgrounds, prioritizeBlue) {
    const target = searchVal.toString().trim().toLowerCase();
    let normalResult = null;

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cellVal = values[r][c] ? values[r][c].toString().trim().toLowerCase() : "";
        if (cellVal === target) {
          if (prioritizeBlue && backgrounds && backgrounds[0] && backgrounds[0][c]) {
            if (backgrounds[0][c].toLowerCase() === "#4a86e8") {
              return { row: r, col: c };
            }
          }
          if (!normalResult) normalResult = { row: r, col: c };
        }
      }
    }
    return normalResult;
  }

  function findInPadrao(searchVal) {
    const target = searchVal.toString().trim().toLowerCase();
    for (let r = 0; r < padraoValues.length; r++) {
      for (let c = 0; c < padraoValues[r].length; c++) {
        if (padraoValues[r][c] && padraoValues[r][c].toString().trim().toLowerCase() === target) {
          return true;
        }
      }
    }
    return false;
  }

  for (let i = 0; i < colF.length; i++) {
    const descricao = colF[i][0];
    if (!descricao || descricao.toString().trim() === "") continue;

    const descricaoStr = descricao.toString().trim();

    if (deveIgnorar(descricaoStr)) continue;

    const rowInSheet = i + 2;
    const cellI = abaMesAtual.getRange(rowInSheet, 9);

    // Passo 1: busca na planilha "Dados"
    const dadosResult = findInValues(dadosValues, descricaoStr, null, false);

    if (!dadosResult) {
      cellI.setValue("Fora dos Dados");
      cellI.setBackground(null);
      cellI.setFontColor("#000000");
      continue;
    }

    // Pega o(s) produto(s) da célula subsequente
    const nextColDados = dadosResult.col + 1;
    let productCell = "";
    if (nextColDados < dadosValues[dadosResult.row].length) {
      productCell = dadosValues[dadosResult.row][nextColDados]
        ? dadosValues[dadosResult.row][nextColDados].toString().trim()
        : "";
    }

    if (!productCell) {
      cellI.setValue("Fora dos Dados");
      cellI.setBackground(null);
      cellI.setFontColor("#000000");
      continue;
    }

    const products = productCell.split(";").map(p => p.trim()).filter(p => p !== "");

    // Passo 2: processa cada produto
    // status: "com_estoque" | "sem_estoque_inv" | "sem_estoque_padrao" | "formatar"
    const results = [];

    for (const product of products) {
      const invResult = findInValues(inventarioValues, product, inventarioBackgrounds, true);

      if (!invResult) {
        if (findInPadrao(product)) {
          results.push({ status: "sem_estoque_padrao", productName: product });
        } else {
          results.push({ status: "formatar", productName: product });
          if (!existingFormatar.has(product) && !novosFormatar.includes(product)) {
            novosFormatar.push(product);
          }
        }
      } else {
        const qtyCol = invResult.col + 1;
        let qty = 0;
        if (qtyCol < inventarioValues[invResult.row].length) {
          qty = inventarioValues[invResult.row][qtyCol];
        }

        // Valor do cabeçalho = linha 1 da coluna onde o produto foi encontrado
        const headerValue = inventarioValues[0][invResult.col]
          ? inventarioValues[0][invResult.col].toString().trim()
          : "";

        const isZero = (qty === 0 || qty === "" || qty === null || qty === undefined);
        if (isZero) {
          results.push({ status: "sem_estoque_inv", productName: product, headerValue });
        } else {
          results.push({ status: "com_estoque", productName: product, headerValue });
        }
      }
    }

    // Passo 3: define o que escrever na coluna I
    const formatarResults      = results.filter(r => r.status === "formatar");
    const semEstoquePadrao     = results.filter(r => r.status === "sem_estoque_padrao");
    const semEstoqueInv        = results.filter(r => r.status === "sem_estoque_inv");
    const comEstoqueResults    = results.filter(r => r.status === "com_estoque");

    if (formatarResults.length > 0) {
      // Não encontrado em nenhuma base → Para Formatar
      cellI.setValue("Para Formatar");
      cellI.setBackground("#0000ff");
      cellI.setFontColor("#ffffff");

    } else if (semEstoquePadrao.length > 0) {
      // Não encontrado no Inventário, encontrado no Padrão → prevalece
      cellI.setValue("Sem Estoque");
      cellI.setBackground("#ffff00");
      cellI.setFontColor("#000000");

    } else if (semEstoqueInv.length > 0 && comEstoqueResults.length === 0) {
      // Todos no Inventário com quantidade 0
      cellI.setValue("Sem Estoque");
      cellI.setBackground("#ffff00");
      cellI.setFontColor("#000000");

    } else if (semEstoqueInv.length > 0 && comEstoqueResults.length > 0) {
      // Alguns com 0, alguns com estoque → Falta Parcial
      const faltando = semEstoqueInv.map(r => r.productName).join(", ");
      cellI.setValue("Falta Parcial + " + faltando);
      cellI.setBackground("#ffff00");
      cellI.setFontColor("#000000");

    } else {
      // Todos com estoque
      if (comEstoqueResults.length === 1) {
        cellI.setValue(comEstoqueResults[0].headerValue);
      } else {
        cellI.setValue(comEstoqueResults.map(r => r.headerValue).join("/"));
      }
      cellI.setBackground("#00ff00");
      cellI.setFontColor("#000000");
    }
  }

  // Grava novos itens na aba "Formatar"
  if (novosFormatar.length > 0) {
    const nextRow = abaFormatar.getLastRow() + 1;
    abaFormatar.getRange(nextRow, 1, novosFormatar.length, 1)
      .setValues(novosFormatar.map(p => [p]));
  }

  Logger.log("✅ Verificação de estoque concluída.");
}
