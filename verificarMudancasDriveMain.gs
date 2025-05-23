function verificarMudancasNoDrive() {
  var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista Coletas Noroeste");

  var pastaID = "1mcJKqLVGMZ6T55V1gEs-8uEw5migX4AZ"; // ID da pasta no Google Drive
  var pasta = DriveApp.getFolderById(pastaID);
  
  var arquivos = pasta.getFiles();

  // **********************
  var b1 = planilhaDestino.getRange("B1").getValue();
  var h1 = planilhaDestino.getRange("H1").getValue();
  // Converte os valores para datas sem horas
  if (b1 instanceof Date) b1 = Utilities.formatDate(b1, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (h1 instanceof Date) h1 = Utilities.formatDate(h1, Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (b1 !== h1) {
    Logger.log("🔄 B1 é diferente de H1. Chamando função importarDadosExcel...");
    importarDadosExcel();
    return;
  } // *********************


  var algumArquivoModificado = false; // Variável de controle

  var ultimaExecucao = PropertiesService.getScriptProperties().getProperty("ultimaExecucao");
  if (!ultimaExecucao) {
    ultimaExecucao = new Date(0).toISOString(); // Se for a primeira execução, considera uma data muito antiga
  }

  var novaUltimaExecucao = new Date().toISOString(); // Atualiza a última execução após processar os arquivos

  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    var dataModificacao = arquivo.getLastUpdated(); // Obtém a data completa como objeto Date

    // Formata a data e hora para exibição no log
    var dataFormatada = Utilities.formatDate(dataModificacao, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    if (dataModificacao.toISOString() > ultimaExecucao) { 
      Logger.log("📂 Arquivo alterado detectado: " + arquivo.getName());
      Logger.log("🕒 Última modificação: " + dataFormatada); // Exibe a hora formatada

      // Chama a função de importação passando o arquivo modificado
      importarArquivoEspecifico(arquivo); 
      algumArquivoModificado = true; // Indica que pelo menos um arquivo foi atualizado
    }    
  }

  // Se nenhum arquivo foi modificado, exibe a mensagem apenas uma vez
  if (!algumArquivoModificado) {
      Logger.log("Nenhum arquivo modificado");
      planilhaDestino.getRange("I1").setValue("⏳ Nenhum arquivo modificado...");
      Utilities.sleep(3000); 
      planilhaDestino.getRange("I1").setValue(""); 
  }

  // Atualiza o tempo da última verificação
  PropertiesService.getScriptProperties().setProperty("ultimaExecucao", novaUltimaExecucao);
}
