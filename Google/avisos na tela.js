function verificarData() {
    // Obtenha a planilha ativa
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var abaAtiva = planilha.getActiveSheet();
    
    // Obtenha os dados da coluna A e B
    var dadosColunaA = abaAtiva.getRange("H5:H").getValues();
    var dadosColunaB = abaAtiva.getRange("I5:I").getValues();
    
    // Obtenha a data atual
    var dataAtual = new Date();
    
    // Itere pelas datas na coluna A
    for (var i = 0; i < dadosColunaA.length; i++) {
      var dataCelula = dadosColunaA[i][0];
      
      // Verifique se a data na célula é igual à data atual
      if (dataCelula instanceof Date && dataCelula.toDateString() === dataAtual.toDateString()) {
        // Obtenha o nome do paciente da coluna B
        var nomePaciente = dadosColunaB[i][0];
        
        // Exiba um aviso com o nome do paciente
        var ui = SpreadsheetApp.getUi();
        ui.alert('Paciente Correspondente à Data Atual', 'O paciente ' + nomePaciente + ' corresponde à data atual.', ui.ButtonSet.OK);
  
      }
    }
  }