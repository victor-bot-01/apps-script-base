function getAccessToken() {
  const refresh_token = PropertiesService.getScriptProperties().getProperty("refresh_token");
  const client_id = PropertiesService.getScriptProperties().getProperty("client_id");
  const client_secret = PropertiesService.getScriptProperties().getProperty("client_secret");

  const tokenUrl = "https://api.bling.com.br/Api/v3/oauth/token";
  const payload = {
    grant_type: "refresh_token",
    refresh_token: refresh_token
  };

  const headers = {
    Authorization: "Basic " + Utilities.base64Encode(client_id + ":" + client_secret),
    "Content-Type": "application/x-www-form-urlencoded",
    Accept: "application/json",
    "enable-jwt": "1"
  };

  const options = {
    method: "post",
    headers: headers,
    payload: Object.entries(payload).map(([k, v]) => `${k}=${encodeURIComponent(v)}`).join("&"),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(tokenUrl, options);
  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200 || !result.access_token) {
    Logger.log("Erro ao renovar token: " + JSON.stringify(result));
    throw new Error("Falha ao renovar access_token.");
  }

  PropertiesService.getScriptProperties().setProperty("access_token", result.access_token);
  if (result.refresh_token) {
    PropertiesService.getScriptProperties().setProperty("refresh_token", result.refresh_token);
  }

  Logger.log("✅ Novo token salvo com sucesso.");
  return result.access_token;
}

function importarPedidosUltimos3Dias() {
  const accessToken = getAccessToken();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mês Atual");
  if (!sheet) {
    Logger.log("❌ Aba 'Mês Atual' não encontrada.");
    return;
  }

  const lojasMapeadas = {
    "204386004": "Amazon 2", "204298767": "Americanas", "204313187": "Americanas 2",
    "204409641": "Americanas 3", "204699807": "ML 1", "204313280": "ML 4",
    "204324111": "Magalu", "204318944": "Magalu 2", "204524610": "Magalu 3",
    "203556912": "Olist", "204313211": "Olist 2", "205145842": "Raia Drogasil",
    "204769662": "Shein", "204313200": "Shopee", "203961843": "Shopee 2",
    "205423237": "Tiktok Essência do Brasil", "204313296": "ViaVarejo",
    "205330649": "Ponte Vecchio (Site)", "205438985": "Essência do Brasil - Site"
  };

  const headers = {
    method: "get",
    headers: { Authorization: `Bearer ${accessToken}`, "enable-jwt": "1" },
    muteHttpExceptions: true
  };

  const hoje = new Date();
  const tresDiasAtras = new Date(hoje.getTime() - 3 * 24 * 60 * 60 * 1000);
  const dataInicial = Utilities.formatDate(tresDiasAtras, "GMT-3", "yyyy-MM-dd'T00:00:00-03:00'");
  const dataFinal = Utilities.formatDate(hoje, "GMT-3", "yyyy-MM-dd'T23:59:59-03:00'");

  // ✅ URL corrigida:
  const url = `https://api.bling.com.br/Api/v3/pedidos/vendas?completo=true&limit=50&sort=-dataEmissao&filters=dataEmissao[${dataInicial} TO ${dataFinal}]`;
  Logger.log("🌐 URL usada: " + url);

  const response = UrlFetchApp.fetch(url, headers);
  const json = JSON.parse(response.getContentText());

  if (!json.data || json.data.length === 0) {
    Logger.log("⚠️ Nenhum pedido encontrado.");
    return;
  }

  const colB = sheet.getRange("B:B").getValues().map(r => String(r[0]).replace(/^'+|'+$/g, '').trim());
  const ultimaLinha = colB.filter(String).length;
  const pedidosExistentes = new Set(colB.slice(1, ultimaLinha));
  const linhasNovas = [];

  for (const resumo of json.data) {
    let numeroLoja = resumo.numeroLoja;
    numeroLoja = String(numeroLoja).replace(/^'+|'+$/g, '').trim();

    if (pedidosExistentes.has(numeroLoja)) {
      Logger.log(`⏩ Pedido ${numeroLoja} já registrado.`);
      continue;
    }

    const pedidoId = resumo.id;

    // ✅ URL corrigida:
    const detalheUrl = `https://api.bling.com.br/Api/v3/pedidos/vendas/${pedidoId}?completo=1`;
    const detalheRes = UrlFetchApp.fetch(detalheUrl, headers);
    const detalheJson = JSON.parse(detalheRes.getContentText());
    const pedido = detalheJson.data;

    if (!pedido || !pedido.itens || pedido.itens.length === 0) {
      Logger.log(`⚠️ Pedido ${numeroLoja} sem itens.`);
      continue;
    }

    const data = pedido.data || "";
    const lojaId = pedido.loja?.id || "";
    const lojaNome = lojasMapeadas[lojaId] || lojaId;
    const cliente = pedido.contato?.nome || "";
    const totalItens = pedido.totalProdutos || "";
    const produtos = pedido.itens || [];
    const volumes = pedido.transporte?.volumes || [];
    const observacoes = pedido.observacoes || "";

    let servico = (["204386004", "204324111"].includes(lojaId.toString()) && volumes.length > 0)
      ? volumes[0].servico || ""
      : "";

    if (!servico && lojaId.toString() === "204324111") {
      servico = observacoes || "";
    }

    const inicio = sheet.getRange("B:B").getValues().filter(r => r[0] !== "").length + 1;

    produtos.forEach((item, index) => {
      const quantidade = parseFloat(item.quantidade || 0);
      const descricao = item.descricao || "Sem descrição";
      const linha = [
        item.codigo,
        numeroLoja,
        lojaNome,
        cliente,
        quantidade,
        `${descricao} (Qtd: ${quantidade})`,
        index === 0 ? totalItens : "",
        index === 0 ? servico : ""
      ];
      sheet.getRange(inicio + index, 1, 1, 8).setValues([linha]);
      linhasNovas.push(inicio + index);
    });

    Logger.log(`✅ Pedido ${numeroLoja} importado com sucesso.`);
  }

  if (linhasNovas.length > 0) {
    const minLinha = Math.min(...linhasNovas);
    const maxLinha = Math.max(...linhasNovas);

    const colI = sheet.getRange(minLinha, 9, maxLinha - minLinha + 1).getValues();
    const colH = sheet.getRange(minLinha, 8, maxLinha - minLinha + 1).getValues();

    sheet.getRange(minLinha, 9, maxLinha - minLinha + 1).setValues(colI);
    sheet.getRange(minLinha, 7, maxLinha - minLinha + 1).clearContent();
    sheet.getRange(minLinha, 7, maxLinha - minLinha + 1).setValues(colH);
    sheet.getRange(minLinha, 8, maxLinha - minLinha + 1).clearContent();

    const colBValues = sheet.getRange(minLinha, 2, maxLinha - minLinha + 1).getValues();
    colBValues.forEach((row, i) => {
      const valor = row[0];
      if (typeof valor === "number" && valor.toExponential().includes("e")) {
        sheet.getRange(minLinha + i, 2).setValue("'" + valor.toFixed(0));
      }
    });

    sheet.getRange(minLinha, 1, maxLinha - minLinha + 1, sheet.getLastColumn())
         .sort({ column: 3, ascending: true });

    const novaColB = sheet.getRange(minLinha, 2, maxLinha - minLinha + 1).getValues().flat();
    let i = 0;
    let corAlternar = true;

    while (i < novaColB.length) {
      const valorAtual = novaColB[i];
      let j = i + 1;
      while (j < novaColB.length && novaColB[j] === valorAtual) j++;
      const tamanhoBloco = j - i;
      if (tamanhoBloco > 1) {
        const cor = corAlternar ? "#cfcfcf" : "#9499b3";
        corAlternar = !corAlternar;
        sheet.getRange(minLinha + i, 1, tamanhoBloco, 8).setBackground(cor);
      }
      i = j;
    }
  }
}
