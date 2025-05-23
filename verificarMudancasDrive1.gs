function verificarMudancasNoDriveeeee() {
  var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista Coletas Noroeste");

  var pastaID = "1mcJKqLVGMZ6T55V1gEs-8uEw5migX4AZ"; // ID da pasta no Google Drive
  var pasta = DriveApp.getFolderById(pastaID);
  
  var arquivos = pasta.getFiles();

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
    }    
  }

  // Atualiza o tempo da última verificação
  PropertiesService.getScriptProperties().setProperty("ultimaExecucao", novaUltimaExecucao);
}
