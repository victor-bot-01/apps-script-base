function copiarValoresFiltrados_paraDestino() {
  // ====== CONFIGURAÇÕES ======
  const NOME_ABA_ORIGEM = "Mês Atual";
  const HEX_AZUL  = "#0000ff"; // cor-alvo na coluna A
  const HEX_VERDE = "#00ff00"; // nova cor após copiar
  const DESTINO_SPREADSHEET_ID = "1RHTxClkkFamZUE4hM8romnZfWWTecNwvkvI4SO7rYmU";
  const DESTINO_ABA_NOME = "Junho";
  // ===========================

  const ssOrigem = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ssOrigem.getSheetByName(NOME_ABA_ORIGEM);
  if (!abaOrigem) throw new Error(`Aba de origem não encontrada: ${NOME_ABA_ORIGEM}`);

  const lastRowOrigem = abaOrigem.getLastRow();
  if (lastRowOrigem < 2) return;

  const numLinhas = lastRowOrigem - 1;
  const dadosBI   = abaOrigem.getRange(2, 2, numLinhas, 8).getValues();      // B2:I
  const fundosA   = abaOrigem.getRange(2, 1, numLinhas, 1).getBackgrounds(); // A2:A

  // Formatações extras omitidas aqui (igual já usamos)
  const fmBackgrounds  = abaOrigem.getRange(2, 2, numLinhas, 8).getBackgrounds();
  const fmNumFormats   = abaOrigem.getRange(2, 2, numLinhas, 8).getNumberFormats();
  const fmFontColors   = abaOrigem.getRange(2, 2, numLinhas, 8).getFontColors();
  const fmFontFamilies = abaOrigem.getRange(2, 2, numLinhas, 8).getFontFamilies();
  const fmFontSizes    = abaOrigem.getRange(2, 2, numLinhas, 8).getFontSizes();
  const fmFontWeights  = abaOrigem.getRange(2, 2, numLinhas, 8).getFontWeights();
  const fmFontStyles   = abaOrigem.getRange(2, 2, numLinhas, 8).getFontStyles();
  const fmWraps        = abaOrigem.getRange(2, 2, numLinhas, 8).getWraps();
  const fmHAligns      = abaOrigem.getRange(2, 2, numLinhas, 8).getHorizontalAlignments();
  const fmVAligns      = abaOrigem.getRange(2, 2, numLinhas, 8).getVerticalAlignments();

  const idxSelecionados = [];
  for (let i = 0; i < numLinhas; i++) {
    const cor = String(fundosA[i][0] || "").toLowerCase();
    if (cor === HEX_AZUL) idxSelecionados.push(i);
  }
  if (idxSelecionados.length === 0) return;

  const outValues  = [];
  const outBg      = [];
  const outNum     = [];
  const outFColor  = [];
  const outFFam    = [];
  const outFSize   = [];
  const outFWght   = [];
  const outFStyle  = [];
  const outWrap    = [];
  const outHAlign  = [];
  const outVAlign  = [];

  idxSelecionados.forEach((iSel) => {
    const row = dadosBI[iSel].slice();
    row[0] = (row[0] === null || row[0] === undefined) ? "" : String(row[0]); // força texto na col B
    outValues.push(row);
    outBg.push(fmBackgrounds[iSel]);
    outNum.push(fmNumFormats[iSel]);
    outFColor.push(fmFontColors[iSel]);
    outFFam.push(fmFontFamilies[iSel]);
    outFSize.push(fmFontSizes[iSel]);
    outFWght.push(fmFontWeights[iSel]);
    outFStyle.push(fmFontStyles[iSel]);
    outWrap.push(fmWraps[iSel]);
    outHAlign.push(fmHAligns[iSel]);
    outVAlign.push(fmVAligns[iSel]);
  });

  // Destino
  const ssDestino = SpreadsheetApp.openById(DESTINO_SPREADSHEET_ID);
  let abaDestino = ssDestino.getSheetByName(DESTINO_ABA_NOME);
  if (!abaDestino) abaDestino = ssDestino.insertSheet(DESTINO_ABA_NOME);

  let primeiraLinhaCola = getLastRowInColumn(abaDestino, 2) + 1;
  if (primeiraLinhaCola < 1) primeiraLinhaCola = 1;

  const linhasNecessarias = (primeiraLinhaCola - 1) + outValues.length;
  if (linhasNecessarias > abaDestino.getMaxRows()) {
    abaDestino.insertRowsAfter(abaDestino.getMaxRows(), linhasNecessarias - abaDestino.getMaxRows());
  }

  const destRange = abaDestino.getRange(primeiraLinhaCola, 2, outValues.length, 8);

  // Força coluna B como texto
  const fmt = [];
  for (let r = 0; r < outValues.length; r++) {
    fmt[r] = ["@", outNum[r][1], outNum[r][2], outNum[r][3], outNum[r][4], outNum[r][5], outNum[r][6], outNum[r][7]];
  }
  destRange.setNumberFormats(fmt);

  // Cola valores
  destRange.setValues(outValues);

  // Cola formatações
  destRange.setBackgrounds(outBg);
  destRange.setFontColors(outFColor);
  destRange.setFontFamilies(outFFam);
  destRange.setFontSizes(outFSize);
  destRange.setFontWeights(outFWght);
  destRange.setFontStyles(outFStyle);
  destRange.setWraps(outWrap);
  destRange.setHorizontalAlignments(outHAlign);
  destRange.setVerticalAlignments(outVAlign);

  // >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  // Pinta de VERDE a coluna A das linhas copiadas
  idxSelecionados.forEach((iSel) => {
    const srcRow = 2 + iSel;
    abaOrigem.getRange(srcRow, 1).setBackground(HEX_VERDE);
  });
  // <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  Logger.log(`Copiadas ${outValues.length} linhas. Linhas na origem marcadas em verde (#00ff00).`);
}

function getLastRowInColumn(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;
  const valores = sheet.getRange(1, col, lastRow, 1).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    if (String(valores[i][0]).trim() !== "") return i + 1;
  }
  return 0;
}
