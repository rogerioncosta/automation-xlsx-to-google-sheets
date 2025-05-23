var contadorArquivos = 0; // Variável global para contar os arquivos processados

function importarArquivoEspecifico(arquivo) {
  contadorArquivos++; // Incrementa o contador a cada execução

  var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista Coletas Noroeste");

  // var b1 = planilhaDestino.getRange("B1").getValue();
  // var h1 = planilhaDestino.getRange("H1").getValue();
  // // Converte os valores para datas sem horas
  // if (b1 instanceof Date) b1 = Utilities.formatDate(b1, Session.getScriptTimeZone(), "yyyy-MM-dd");
  // if (h1 instanceof Date) h1 = Utilities.formatDate(h1, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // if (b1 !== h1) {
  //   Logger.log("🔄 B1 é diferente de H1. Chamando função importarDadosExcel...");
  //   importarDadosExcel();
  //   return;
  // }

  Logger.log("📥 Processando arquivo: " + arquivo.getName());
  
  // Exibe a mensagem inicial de progresso
  planilhaDestino.getRange("I1").setValue("⏳ Importação em andamento...");
  planilhaDestino.getRange("I1").setValue("📂 Iniciando o processo...");
 
  var nomeArquivo = arquivo.getName();
  Logger.log("📥 Convertendo arquivo: " + nomeArquivo);

  var blob = arquivo.getBlob();
  var novoArquivo = DriveApp.createFile(blob);
  var novoArquivoID = novoArquivo.getId();
  var arquivoConvertido = Drive.Files.copy(
    { title: arquivo.getName().replace(".xlsx", ""), mimeType: MimeType.GOOGLE_SHEETS },
    novoArquivoID
  );

  var planilhaID = arquivoConvertido.id;
  Logger.log("✅ Planilha convertida: " + planilhaID);

  // Aguarda até que o arquivo esteja pronto para abrir
  Utilities.sleep(5000);
  var ss = SpreadsheetApp.openById(planilhaID);
  Logger.log("📖 Planilha aberta: " + ss.getName());

  var abas = ss.getSheets();
  Logger.log("🔍 Número de abas: " + abas.length);
  
  if (abas.length < 2) {
    Logger.log("❌ Arquivo sem a aba 2: " + arquivo.getName());
    DriveApp.getFileById(planilhaID).setTrashed(true);
    return;
  }

  var aba = abas[1]; // Segunda aba (índice 1)
  Logger.log("📄 Lendo aba: " + aba.getName());
  
  var idOrigem = aba.getRange("A1").getValue();
  Logger.log("🆔 ID da origem: " + idOrigem);
  
  var linhaDestino = encontrarLinhaPorID(planilhaDestino, idOrigem); 

  var i3 = aba.getRange("I3").getDisplayValue(); // Obtém o valor formatado como string
  Logger.log("Valor formatado da célula I3: " + i3);
  var a1 = aba.getRange("A1").getValue();
  var a11 = aba.getRange("A11").getValue();
  var i12 = aba.getRange("I12").getValue();
  var c12 = aba.getRange("C12").getValue();
  var g9 = aba.getRange("G9").getValue();
  var i7 = aba.getRange("I7").getValue();
  
  var linhaDados = [i3, a1, a11, i12, c12, g9, i7];

  // Atualizar apenas se houver ID correspondente
  if (linhaDestino > 0) {
    var linhaDestinoValores = planilhaDestino.getRange(linhaDestino, 2, 1, linhaDados.length).getValues()[0];

    for (var i = 0; i < linhaDados.length; i++) {
      if (linhaDados[i] !== linhaDestinoValores[i]) {
        planilhaDestino.getRange(linhaDestino, i + 2).setValue(linhaDados[i]);
      }
    }
    Logger.log("♻️ Atualizado na linha " + linhaDestino);
  } else {
    // Se ID não encontrado, inserir na próxima linha vazia
    linhaDestino = encontrarProximaLinhaVazia(planilhaDestino);
    planilhaDestino.getRange(linhaDestino, 2, 1, linhaDados.length).setValues([linhaDados]);
    Logger.log("✅ Novo registro inserido na linha " + linhaDestino);
  }

  // Exclui arquivos temporários
  DriveApp.getFileById(planilhaID).setTrashed(true);
  DriveApp.getFileById(novoArquivoID).setTrashed(true);

  Logger.log("🎯 Importação concluída!");

  // Atualiza a célula A1 com a mensagem final
  planilhaDestino.getRange("I1").setValue("✅ Importação concluída!");
  planilhaDestino.getRange("I1").setValue("✅ Processo concluído com sucesso.");

  SpreadsheetApp.flush(); // Força a atualização final

  // Aguarda 5 segundos
  Utilities.sleep(3000);

  // Limpa as células de progresso
  planilhaDestino.getRange("I1").setValue("");  // Limpa a célula de progresso
  SpreadsheetApp.flush(); // Força a atualização para limpar a célula

  // Verifica se já processou 5 arquivos e pausa
  if (contadorArquivos % 5 === 0) {
    Logger.log("⏳ Pausando por 5 segundos para evitar expiração...");
    Utilities.sleep(5000); // Pausa de 10 segundos
  }

}
