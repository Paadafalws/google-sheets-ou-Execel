function ajustarDatas() {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var abaAtiva = planilha.getActiveSheet();
    var valoresColunaY = abaAtiva.getRange("Y5:Y" + abaAtiva.getLastRow()).getValues();
    var valoresColunaH = [];
  
    for (var i = 0; i < valoresColunaY.length; i++) {
      var dataY = valoresColunaY[i][0];
  
      if (dataY instanceof Date) {
        // Verifica se a data é anterior a 2020
        if (dataY < new Date('2020-01-01')) {
          valoresColunaH.push(['']); // Adiciona uma célula em branco
        } else {
          var diaDaSemana = dataY.getDay();
          if (diaDaSemana === 0 || diaDaSemana === 6) {
            dataY.setDate(dataY.getDate() + 2);
          }
          valoresColunaH.push([dataY]);
        }
      } else {
        // Se a célula em "Y" não contém uma data válida, mantém "H" em branco com a cor da fonte branca
        valoresColunaH.push(['']).setFontColor('white');
      }
    }
  
    var rangeColunaH = abaAtiva.getRange("H5:H" + (valoresColunaH.length + 4));
    rangeColunaH.setValues(valoresColunaH);
  }
  