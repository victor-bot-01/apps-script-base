function executarImportacoes() {
  // Obtém um lock para o documento
  var lock = LockService.getDocumentLock();
  
  // Tenta adquirir o lock imediatamente (sem esperar)
  if (!lock.tryLock(0)) {
    Logger.log("Execução abortada: outro processo de importação já está rodando.");
    return; // aborta a execução
  }

  try {
    Logger.log("🚀 Iniciando importação de pedidos do Bling...");
    importarPedidosUltimos3Dias();
    Logger.log("✅ Importação do Bling concluída.");

    Logger.log("📥 Iniciando importação do último Excel do ML...");
    importarUltimoExcelDaPasta();
    Logger.log("✅ Importação do Excel do ML concluída.");

    Logger.log("🔄 Iniciando execução do pipeline completo...");
    executarPipelineCompleto();
    Logger.log("✅ Pipeline completo concluído.");

    Logger.log("🖨️ Exportando etiquetas no formato 58mm...");
    exportarEtiquetas58mmParaPDF();
    Logger.log("✅ PDF das etiquetas exportado com sucesso.");

    // 🔽 Aqui entra a sua nova função 🔽
    Logger.log("📤 Copiando valores filtrados para a planilha destino...");
    copiarValoresFiltrados_paraDestino();
    Logger.log("✅ Valores copiados e linhas origem marcadas em verde.");

    Logger.log("🏁 Todas as importações foram realizadas com sucesso.");
  } catch (erro) {
    Logger.log("❌ Ocorreu um erro durante a execução: " + erro.message);
  } finally {
    // Libera o lock (sempre)
    lock.releaseLock();
  }
}
