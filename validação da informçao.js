function CalcularData() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var abaAtiva = planilha.getActiveSheet();
    var valoresColunaG = abaAtiva.getRange("G5:G" + abaAtiva.getLastRow()).getValues();
    var valoresColunaO = abaAtiva.getRange("O5:O" + abaAtiva.getLastRow()).getValues();
    var valoresColunaB = abaAtiva.getRange("B5:B" + abaAtiva.getLastRow()).getValues();
    var valoresColunaX = []; // Alterado para valoresColunaX
  
    for (var i = 0; i < valoresColunaG.length; i++) {
      var valorG = valoresColunaG[i][0];
      var valorO = valoresColunaO[i][0];
      var valorB = new Date(valoresColunaB[i][0]);
  
      if (valorG === "Atendimento (Externo Pós Vendas)") {
        valorB.setDate(valorB.getDate() + 7);
      } else if (valorG === "Atendimento (Externo Autorização )" && valorO === "Medicação") {
        valorB.setDate(valorB.getDate() + 10);
      } else if (valorG === "Atendimento (Externo Autorização )" && valorO === "Cirurgia") {
        valorB.setDate(valorB.getDate() + 21);
      } else if (valorG === "Atendimento (Externo Autorização )" && valorO === "Consulta") {
        valorB.setDate(valorB.getDate() + 5);
      } else if (valorG === "Atendimento (Externo Autorização )" && valorO === "Exame") {
        valorB.setDate(valorB.getDate() + 5);
      } else if (valorG === "Concluido" || valorG === "Cancelado" || valorG === "Atendimento"|| valorG === "Concluido (Externo Pós Vendas)" || valorG === "Concluido (Externo Agenda)"|| valorG === "Concluido (Externo Autorização)" ) {
        // Se for "Concluído", "Cancelado" ou "Atendimento", deixe a célula na coluna "B" vazia
       valorB = new Date(0); // Define como data vazia (01/01/1970)
      }
      var dataFormatada = Utilities.formatDate(valorB, 'GMT-3', 'dd/MM/yy');
      valoresColunaX.push([dataFormatada]); // Alterado para valoresColunaX
    }
  
    var rangeColunaX = abaAtiva.getRange("Y5:Y" + (valoresColunaX.length + 4)); // Alterado para coluna "X"
    rangeColunaX.setValues(valoresColunaX); // Alterado para coluna "X"
  }
  