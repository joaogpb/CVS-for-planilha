function importCSVsFromFolderToSeparateSheets() {
  // ID da pasta no Google Drive onde os arquivos CSV estão localizados
  var folderId = 'ID_pasta'; // ID da sua pasta
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.CSV);

  // ID da planilha onde os dados serão importados
  var spreadsheetId = 'ID_pasta';
  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Nome da aba de controle
  var controlSheetName = 'Control';
  var controlSheet = spreadsheet.getSheetByName(controlSheetName);

  // Se a aba de controle não existir, crie uma nova
  if (!controlSheet) {
    controlSheet = spreadsheet.insertSheet(controlSheetName);
    controlSheet.appendRow(['FileName']); // Cabeçalho para os nomes dos arquivos
  }else{
    controlSheet.clear();
    //controlSheet.appendRow(['FileName']);// Cabeçalho para os nomes dos arquivos
  }

  // Verificar se há dados na aba de controle
  var lastRow = controlSheet.getLastRow();
  var importedFiles = [];
  if (lastRow > 1) {
    importedFiles = controlSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  }

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    if (importedFiles.includes(fileName)) {
      Logger.log('Arquivo já importado: ' + fileName);
      continue;
    }

    try {
      var csvData = file.getBlob().getDataAsString();
      var rows = Utilities.parseCsv(csvData);

      // Nome da nova aba será o nome do arquivo (sem a extensão)
      var sheetName = fileName.replace('.csv', '');
      
      // Criar uma nova aba com o nome do arquivo
      var sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
      } else {
        // Limpar a aba existente antes de adicionar novos dados
        sheet.clear();
      }
      
      // Inserir os dados do CSV na nova aba
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      
      // Adicionar o nome do arquivo à aba de controle
      controlSheet.appendRow([fileName]);
    } catch (e) {
      Logger.log('Erro ao processar o arquivo: ' + fileName + '. Erro: ' + e.message);
    }
  }

  Logger.log('Todos os arquivos CSV foram importados com sucesso.');
}
