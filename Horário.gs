/**
 * Mantém seu gatilho chamando esta função.
 * Ela só libera a execução no horário permitido.
 */
function horarioDeFuncionamento() {
  const agora = new Date();
  const dia = agora.getDay(); // 0 dom ... 6 sáb
  const minutosAgora = agora.getHours() * 60 + agora.getMinutes();

  let dentroDoHorario = false;

  if (dia >= 1 && dia <= 5) {
    const inicio = 6 * 60;
    const fim = 19 * 60 + 45;
    dentroDoHorario = minutosAgora >= inicio && minutosAgora <= fim;
  } else if (dia === 6) {
    const inicio = 6 * 60 + 30;
    const fim = 15 * 60;
    dentroDoHorario = minutosAgora >= inicio && minutosAgora <= fim;
  } else if (dia === 0) {
    const inicio = 17 * 60;
    const fim = 20 * 60;
    dentroDoHorario = minutosAgora >= inicio && minutosAgora <= fim;
  }

  if (dentroDoHorario) {
    rodarImportacoesComTrava_();
  } else {
    Logger.log("Fora do horário de funcionamento: " + agora);
  }
}

/**
 * Guardião: impede concorrência (manual + gatilho) no MESMO projeto.
 * Combina:
 * - ScriptLock (bloqueio do projeto)
 * - Lease em ScriptProperties com TTL renovável (mais robusto em execuções longas)
 */
function rodarImportacoesComTrava_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // espera até 30s

  const props = PropertiesService.getScriptProperties();
  const runId = Utilities.getUuid();

  // ✅ TTL maior que sua pior execução. Sugiro 120 min para ficar folgado.
  const TTL_MS = 120 * 60 * 1000;

  try {
    if (!acquireLease_(props, runId, TTL_MS)) {
      Logger.log("⛔ Abortado: já existe uma execução ativa (lease).");
      return;
    }

    Logger.log("🚀 Iniciando importações. runId=" + runId);
    heartbeat_(props, runId, TTL_MS);

    executarImportacoesReal_(props, runId, TTL_MS);

    Logger.log("🏁 Finalizado com sucesso. runId=" + runId);

  } catch (e) {
    Logger.log("❌ Erro: " + e);
    if (e && e.stack) Logger.log(e.stack);
    throw e;
  } finally {
    releaseLease_(props, runId);
    lock.releaseLock();
  }
}

/**
 * Sua lógica REAL.
 * ✅ Heartbeat agora RENOVA o TTL.
 */
function executarImportacoesReal_(props, runId, TTL_MS) {
  Logger.log("🚀 Iniciando importação de pedidos do Bling...");
  importarPedidosUltimos3Dias();
  heartbeat_(props, runId, TTL_MS);
  Logger.log("✅ Importação do Bling concluída.");

  Logger.log("📥 Iniciando importação do último Excel do ML...");
  importarUltimoExcelDaPasta(props, runId, TTL_MS);
  heartbeat_(props, runId, TTL_MS);
  Logger.log("✅ Importação do Excel do ML concluída.");

  Logger.log("🔄 Iniciando execução do pipeline completo...");
  executarPipelineCompleto();
  heartbeat_(props, runId, TTL_MS);
  Logger.log("✅ Pipeline completo concluído.");

  Logger.log("🖨️ Exportando etiquetas no formato 58mm...");
  exportarEtiquetas58mmParaPDF();
  heartbeat_(props, runId, TTL_MS);
  Logger.log("✅ PDF das etiquetas exportado com sucesso.");

  Logger.log("📤 Copiando valores filtrados para a planilha destino...");
  copiarValoresFiltrados_paraDestino();
  heartbeat_(props, runId, TTL_MS);
  Logger.log("✅ Valores copiados e linhas origem marcadas em verde.");
}

/* =========================
   LEASE (cadeado lógico)
   ========================= */

/**
 * ✅ Regra agora é simples e forte:
 * - Se expiresAt ainda está no futuro => outro processo está ativo => não entra.
 * (Sem janela de 10 minutos, que era o que deixava furar.)
 */
function acquireLease_(props, runId, ttlMs) {
  const now = Date.now();
  const raw = props.getProperty("IMPORT_LEASE");

  if (raw) {
    const lease = JSON.parse(raw);
    if (lease.expiresAt && lease.expiresAt > now) {
      return false;
    }
  }

  props.setProperty("IMPORT_LEASE", JSON.stringify({
    runId,
    startedAt: now,
    heartbeatAt: now,
    expiresAt: now + ttlMs,
  }));

  return true;
}

/**
 * ✅ Heartbeat RENOVA o TTL:
 * - atualiza heartbeatAt
 * - empurra expiresAt para now + ttlMs
 */
function heartbeat_(props, runId, ttlMs) {
  const raw = props.getProperty("IMPORT_LEASE");
  if (!raw) return;

  const lease = JSON.parse(raw);
  if (lease.runId !== runId) return;

  const now = Date.now();
  lease.heartbeatAt = now;
  lease.expiresAt = now + ttlMs;

  props.setProperty("IMPORT_LEASE", JSON.stringify(lease));
}

function releaseLease_(props, runId) {
  const raw = props.getProperty("IMPORT_LEASE");
  if (!raw) return;

  const lease = JSON.parse(raw);
  if (lease.runId === runId) props.deleteProperty("IMPORT_LEASE");
}

/**
 * Botão de emergência: use só se tiver certeza que não há execução rodando.
 */
function limparLeaseImportacao() {
  PropertiesService.getScriptProperties().deleteProperty("IMPORT_LEASE");
  Logger.log("🧹 Lease removido manualmente.");
}
