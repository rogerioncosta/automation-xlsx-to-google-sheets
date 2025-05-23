function onEdit(e) {
  var planilhaDestino = e.source.getSheetByName("Lista Coletas Noroeste");
  var intervaloEditado = e.range;  // A célula que foi editada
  var colunaEditada = intervaloEditado.getColumn();  // Coluna que foi editada
  var linhaEditada = intervaloEditado.getRow();  // Linha que foi editada

  // Verifica se a edição foi na coluna I (de I3 a I27)
  if (colunaEditada === 9 && linhaEditada >= 3 && linhaEditada <= 27) {
    var valorCelulaI = intervaloEditado.getValue();  // Valor editado na célula da coluna I

    // Se o valor da célula na coluna I for "OK"
    if (valorCelulaI === "OK" || valorCelulaI === "Ok" || valorCelulaI === "ok") {
      var horaAtual = new Date();  // Obtém a hora atual
      planilhaDestino.getRange(linhaEditada, 11).setValue(horaAtual);  // Preenche a hora na coluna K (coluna 11)
    } else {
      // Se o valor não for "OK", limpa a célula correspondente na coluna K
      planilhaDestino.getRange(linhaEditada, 11).clearContent();
    }
  }
}
