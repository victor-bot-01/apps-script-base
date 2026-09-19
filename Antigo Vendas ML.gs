function importarUltimoExcelDaPasta(props, runId, TTL_MS) {
  const hb = () => {
    if (props && runId && TTL_MS) heartbeat_(props, runId, TTL_MS);
  };

  hb();

  const nomeDaPasta = "Vendas ML";
  const pasta = DriveApp.getFoldersByName(nomeDaPasta);

  if (!pasta.hasNext()) {
    Logger.log("Pasta 'Vendas ML' não encontrada.");
    return;
  }

  const pastaSelecionada = pasta.next();
  const arquivos = pastaSelecionada.getFilesByType(MimeType.MICROSOFT_EXCEL);

  let arquivoMaisRecente = null;
  let dataMaisRecente = 0;

  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const dataModificacao = arquivo.getLastUpdated().getTime();
    if (dataModificacao > dataMaisRecente) {
      dataMaisRecente = dataModificacao;
      arquivoMaisRecente = arquivo;
    }
  }

  if (!arquivoMaisRecente) {
    Logger.log("Nenhum arquivo .xlsx encontrado na pasta.");
    return;
  }

  hb();

  // Converte XLSX -> Google Sheets (Drive API avançada precisa estar habilitada)
  const blob = arquivoMaisRecente.getBlob();
  const arquivoConvertido = DriveApi.Files.insert(
    {
      title: "TEMP_CONVERTIDO_" + Date.now(),
      mimeType: MimeType.GOOGLE_SHEETS
    },
    blob
  );

  const tempId = arquivoConvertido.id;

  hb();

  const planilhaTemp = SpreadsheetApp.openById(tempId);
  const abaOrigem = planilhaTemp.getSheets()[0];
  const dados = abaOrigem.getDataRange().getValues();

  hb();

  const cabecalho = dados[5];
  const idxVenda = cabecalho.indexOf("N.º de venda");
  const idxEstado = cabecalho.indexOf("Estado");
  const idxUnidades = cabecalho.indexOf("Unidades");
  const idxTotal = cabecalho.indexOf("Total (BRL)");
  const idxTitulo = cabecalho.indexOf("Título do anúncio");
  const idxComprador = cabecalho.indexOf("Comprador");
  const idxEntrega = cabecalho.indexOf("Forma de entrega");

  if ([idxVenda, idxEstado, idxUnidades, idxTotal, idxTitulo, idxComprador, idxEntrega].some(i => i === -1)) {
    Logger.log("Uma ou mais colunas não foram encontradas no relatório do ML.");
    return;
  }

  // ⚠️ Se em algum momento isso falhar via gatilho, troque por openById('ID_DA_PLANILHA')
  const planilhaDestino = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = planilhaDestino.getSheetByName("Mês Atual");

  // Carregar pedidos existentes (coluna B)
  const pedidosExistentes = new Set();
  const lastRow = abaDestino.getLastRow();
  if (lastRow >= 2) {
    const linhasExistentes = abaDestino.getRange(2, 2, lastRow - 1, 1).getValues();
    linhasExistentes.forEach(linha => {
      const numero = linha[0];
      if (numero) pedidosExistentes.add(numero.toString().trim());
    });
  }

  hb();

  const dadosFiltrados = [];
  const linhasMultiploItem = new Set();

  let ultimoNumeroVenda = "";
  let ultimoComprador = "";
  let linhaPrincipalIndex = null;

  for (let i = 6; i < dados.length; i++) {
    if (i % 300 === 0) hb(); // ✅ heartbeat durante loop longo

    const linha = dados[i];
    const numeroVendaAtual = linha[idxVenda]?.toString().trim();
    const compradorAtual = linha[idxComprador]?.toString().trim();

    if (compradorAtual) {
      ultimoNumeroVenda = numeroVendaAtual;
      ultimoComprador = compradorAtual;

      if (pedidosExistentes.has(ultimoNumeroVenda)) {
        let j = i + 1;
        while (
          j < dados.length &&
          (!dados[j][idxComprador] || dados[j][idxComprador].toString().trim() === "")
        ) {
          j++;
        }
        i = j - 1;
        continue;
      }

      linhaPrincipalIndex = dadosFiltrados.length;
    } else {
      if (linhaPrincipalIndex !== null) {
        linhasMultiploItem.add(linhaPrincipalIndex);
      }
    }

    const numeroVendaFinal = compradorAtual ? numeroVendaAtual : ultimoNumeroVenda;
    const compradorFinal = compradorAtual || ultimoComprador;
    const valorEntrega = linha[idxEntrega]?.toString().trim();
    const valorEstado = linha[idxEstado]?.toString().trim();

    const itensG = [];

    if (valorEntrega === "Mercado Envios Flex") {
      itensG.push("Flex");
    } else if (valorEntrega === "Mercado Envios Full") {
      itensG.push(valorEntrega);
    }

    if (
      valorEstado.startsWith("Venda cancelada") ||
      valorEstado.startsWith("Para entregar na coleta")
    ) {
      itensG.push(valorEstado);
    }

    const valorColunaG = itensG.join(" | ");

    dadosFiltrados.push([
      "",                        // A
      numeroVendaFinal,          // B
      "ML",                      // C
      compradorFinal,            // D
      linha[idxUnidades],        // E
      linha[idxTitulo],          // F
      valorColunaG,              // G
      "",                        // H
      linha[idxTotal]            // I
    ]);

    if (!compradorAtual && linhaPrincipalIndex !== null) {
      linhasMultiploItem.add(dadosFiltrados.length - 1);
    }
  }

  hb();

  // Encontrar a primeira linha vazia (sem getRange em loop)
  const totalLinhas = abaDestino.getMaxRows();
  let inicioInsercao = 2;

  const bloco = abaDestino.getRange(2, 1, Math.max(1, totalLinhas - 1), 8).getValues();
  for (let r = 0; r < bloco.length; r++) {
    const vazia = bloco[r].every(cel => cel === "" || cel === null);
    if (vazia) {
      inicioInsercao = 2 + r;
      break;
    }
  }

  const totalLinhasAdicionadas = dadosFiltrados.length;

  if (totalLinhasAdicionadas > 0) {
    const rangeDestino = abaDestino.getRange(inicioInsercao, 1, totalLinhasAdicionadas, 9);
    rangeDestino.setValues(dadosFiltrados);
    rangeDestino.setBorder(true, true, true, true, true, true);

    hb();

    // Cinza para múltiplos itens
    linhasMultiploItem.forEach((idx) => {
      const linha = inicioInsercao + idx;
      abaDestino.getRange(linha, 1, 1, 9).setBackground("#cfcfcf");
    });

    hb();

    // Azul para "Mercado Envios Full" (coluna G)
    const colGRange = abaDestino.getRange(inicioInsercao, 7, totalLinhasAdicionadas, 1);
    const colGValues = colGRange.getValues();
    for (let i = 0; i < colGValues.length; i++) {
      if (i % 400 === 0) hb();
      const texto = String(colGValues[i][0] || "");
      if (texto.includes("Mercado Envios Full")) {
        abaDestino.getRange(inicioInsercao + i, 7).setBackground("#0000ff");
      }
    }

    hb();

    // Corrigir notação científica (coluna B)
    const colBRange = abaDestino.getRange(inicioInsercao, 2, totalLinhasAdicionadas, 1);
    const colBValues = colBRange.getValues();
    colBValues.forEach((row, i) => {
      if (i % 400 === 0) hb();
      const valor = row[0];
      if (typeof valor === "number" && valor.toExponential().includes("e")) {
        abaDestino.getRange(inicioInsercao + i, 2).setValue("'" + valor.toFixed(0));
      }
    });

    hb();

    // Remover linhas com E e F vazias (nas linhas recém-adicionadas)
    const valores = abaDestino.getRange(inicioInsercao, 1, totalLinhasAdicionadas, 9).getValues();
    let linhaOffset = 0;

    for (let i = 0; i < valores.length; i++) {
      if (i % 200 === 0) hb();

      const linha = valores[i];
      const linhaReal = inicioInsercao + i - linhaOffset;

      if (
        (!linha[4] || linha[4].toString().trim() === "") &&
        (!linha[5] || linha[5].toString().trim() === "")
      ) {
        const valorG = linha[6];
        const valorI = linha[8];

        if (valorG || valorI) {
          abaDestino.getRange(linhaReal + 1, 7).setValue(valorG);
          abaDestino.getRange(linhaReal + 1, 9).setValue(valorI);
        }

        abaDestino.deleteRow(linhaReal);
        linhaOffset++;
      }
    }
  } else {
    Logger.log("Nenhum novo pedido foi importado. Todos já existem na planilha.");
  }

  hb();

  // Excluir arquivos temporários
  DriveApp.getFileById(tempId).setTrashed(true);

  // ✅ Excluir o Excel original (intencional)
  DriveApp.getFileById(arquivoMaisRecente.getId()).setTrashed(true);

  Logger.log("Importação concluída.");
}
