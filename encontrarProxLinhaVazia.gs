function encontrarProximaLinhaVazia(planilha) {
  var intervalo = planilha.getRange("B:B");
  var valores = intervalo.getValues();

  for (var i = 1; i < valores.length; i++) {
    if (!valores[i][0]) {
      return i + 1; // Retorna a linha vazia
    }
  }
  return valores.length + 1; // Caso todas as linhas estejam preenchidas
}
