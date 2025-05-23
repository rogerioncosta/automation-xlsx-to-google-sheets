function encontrarLinhaPorID(planilha, idOrigem) {
  if (!idOrigem) return -1; // Se ID estiver vazio, não faz nada

  var intervalo = planilha.getRange("C3:C"); // Coluna C a partir da linha 3
  var valores = intervalo.getValues();
  
  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] == idOrigem) {
      return i + 3; // Retorna a linha correspondente (começa em C3)
    }
  }
  return -1; // ID não encontrado
}
