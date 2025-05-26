function importarDadosExcel_depreciado() {
  var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lista Coletas Noroeste");

  // Exibe a mensagem inicial de progresso
  planilhaDestino.getRange("I1").setValue("⏳ Importação em andamento...");
  planilhaDestino.getRange("I1").setValue("📂 Iniciando o processo...");
  SpreadsheetApp.flush(); // Força a atualização da planilha

  var pastaID = "ID-da-pasta-no-GoogleDrive"; // ID da pasta no Google Drive
  var pasta = DriveApp.getFolderById(pastaID);
  var arquivos = pasta.getFiles();
  var totalArquivos = 0;


  var b1 = planilhaDestino.getRange("B1").getValue();
  var h1 = planilhaDestino.getRange("H1").getValue();

  // Converte os valores para datas sem horas
  if (b1 instanceof Date) b1 = Utilities.formatDate(b1, Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (h1 instanceof Date) h1 = Utilities.formatDate(h1, Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (b1 !== h1) {
    Logger.log("🔄 B1 é diferente de H1. Limpando B3:K27...");
    planilhaDestino.getRange("B3:K27").clearContent(); // Limpa os valores sem remover a formatação

    // Atualiza B1 para ter o mesmo valor de H1
    // planilhaDestino.getRange("B1").setValue(new Date()); // Define para a data de hoje
    planilhaDestino.getRange("B1").setValue(h1); // Define o valor igual a H1
  } else {
    Logger.log("✅ B1 e H1 são iguais. Nenhuma ação necessária.");
  }

  var linhaDestino = 3; // Começa inserindo na linha 3 da planilha de destino

  // Conta o número total de arquivos para calcular o progresso
  while (arquivos.hasNext()) {
    arquivos.next();
    totalArquivos++;
  }

  arquivos = pasta.getFiles();  // Reinicia o loop para percorrer os arquivos

  var currentFile = 0;

  while (arquivos.hasNext()) {
    var arquivo = arquivos.next();
    var nomeArquivo = arquivo.getName();
    
    Logger.log("📂 Verificando arquivo: " + nomeArquivo);

    if (!nomeArquivo.endsWith(".xlsx")) {
      Logger.log("❌ Ignorado (não é .xlsx): " + nomeArquivo);
      continue;
    }

    try {
      currentFile++;
      var progress = Math.round((currentFile / totalArquivos) * 100);
      var message = "⏳ Processando arquivo " + currentFile + " de " + totalArquivos;

      // Atualiza a célula com o progresso
      planilhaDestino.getRange("I1").setValue(message + " (" + progress + "%)");
      SpreadsheetApp.flush(); // Força a atualização do progresso na planilha

      Logger.log("📥 Convertendo arquivo: " + nomeArquivo);
      
      // Converte o arquivo Excel para Google Sheets
      var blob = arquivo.getBlob();
      var novoArquivo = DriveApp.createFile(blob);
      var novoArquivoID = novoArquivo.getId();
      var arquivoConvertido = Drive.Files.copy(
        { title: nomeArquivo.replace(".xlsx", ""), mimeType: MimeType.GOOGLE_SHEETS },
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
      
      if (abas.length < 2) { // Agora verificamos se há pelo menos 2 abas
        Logger.log("❌ Arquivo sem a aba 2: " + nomeArquivo);
        DriveApp.getFileById(planilhaID).setTrashed(true);
        continue;
      }

      var aba = abas[1]; // Pega a segunda aba (índice 1)
      Logger.log("📄 Lendo aba: " + aba.getName());

      // Pega os valores específicos da segunda aba
      var i3 = aba.getRange("I3").getDisplayValue(); // Obtém o valor formatado como string
      Logger.log("Valor formatado da célula I3: " + i3);
      var a1 = aba.getRange("A1").getValue();
      var a11 = aba.getRange("A11").getValue();
      var i12 = aba.getRange("I12").getValue();
      var c12 = aba.getRange("C12").getValue();
      var g9 = aba.getRange("G9").getValue();
      var i7 = aba.getRange("I7").getValue();

      // *****
      // Monta a linha com os dados coletados
      var linhaDados = [i3, a1, a11, i12, c12, g9, i7];

      // Verifica a última linha preenchida na coluna B
      var ultimaLinhaPreenchida = planilhaDestino.getRange("B:B").getValues().filter(String).length;
      var linhaDestino = ultimaLinhaPreenchida + 1; // Começa na próxima linha disponível

      // Verifica se o ID de origem (A1) está vazio
      if (a1 === "" && planilhaDestino.getRange(linhaDestino, 3).getValue() === "") {
        // Se ambos os IDs estão em branco, preenche na próxima linha disponível
        planilhaDestino.getRange(linhaDestino, 2, 1, linhaDados.length).setValues([linhaDados]);
        Logger.log("✅ Dados inseridos na linha " + linhaDestino);
      } else {
        // Caso contrário, verifica os IDs
        var foundMatch = false;

        for (var i = 3; i <= 27; i++) {
          var idDestino = planilhaDestino.getRange(i, 3).getValue(); // ID na célula C da linha atual

          if (a1 === idDestino || idDestino === "") { // Se o ID de origem for igual ao ID de destino ou ID de destino estiver em branco
            foundMatch = true;

            // Preenche as células diferentes
            var linhaDestinoValores = planilhaDestino.getRange(i, 2, 1, linhaDados.length).getValues()[0];

            for (var j = 0; j < linhaDados.length; j++) {
              if (linhaDados[j] !== linhaDestinoValores[j]) {
                planilhaDestino.getRange(i, j + 2).setValue(linhaDados[j]);
              }
            }

            Logger.log("✅ Dados inseridos na linha " + i);      
            break; // Sai do loop se encontrar uma correspondência
          }
        }

        // Se não encontrar a correspondência, começa a preencher na próxima linha disponível
        if (!foundMatch) {
          planilhaDestino.getRange(linhaDestino, 2, 1, linhaDados.length).setValues([linhaDados]);
          Logger.log("✅ Dados inseridos na linha " + linhaDestino);
        }
      }

      // Logger.log("✅ Dados inseridos na linha " + linhaDestino);      
      
      // Log dos valores da linha de origem
      Logger.log("🌍 Linha de origem (dados da planilha convertida):");
      Logger.log(linhaDados);
      // Log dos valores da linha de destino
      Logger.log("🏁 Linha de destino (dados da planilha 'Lista Coletas Noroeste'):");
      Logger.log(linhaDestinoValores);

      // linhaDestino++; // Avança para a próxima linha

      // *****


      // Exclui arquivos temporários
      DriveApp.getFileById(planilhaID).setTrashed(true);
      DriveApp.getFileById(novoArquivoID).setTrashed(true);

    } catch (e) {
      Logger.log("❌ Erro ao processar " + nomeArquivo + ": " + e.toString());
    }
  }

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

  // Organiza de A a Z as linhas de B3:K27
  var rangeToSort = planilhaDestino.getRange("B3:K27");
  rangeToSort.sort({ column: 2, ascending: true }); // Ordena pela coluna B (segunda coluna)
}
