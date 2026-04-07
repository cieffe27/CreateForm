function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CreaForm')
    .addItem('Crea Google Modulo', 'createGoogleForm')
    .addToUi();
}

function createGoogleForm() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const spreadsheetName = spreadsheet.getName();
  const spreadsheetId = spreadsheet.getId();

  // Crea il modulo con lo stesso nome del FILE Google Fogli
  const form = FormApp.create(spreadsheetName);
  form.setIsQuiz(false);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const domanda = row[0];
    const risposte = row.slice(1, 5).filter(r => r !== '');

    if (!domanda || risposte.length < 2) continue;

    const item = form.addMultipleChoiceItem();
    item.setTitle(domanda);
    item.setChoices(
      risposte.map(r => item.createChoice(r))
    );
  }

  // --- SPOSTAMENTO DEL MODULO NELLA STESSA CARTELLA DEL FOGLIO ---

  const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  const parentFolders = spreadsheetFile.getParents();

  if (parentFolders.hasNext()) {
    const parentFolder = parentFolders.next();
    DriveApp.getFileById(form.getId()).moveTo(parentFolder);
  }

  SpreadsheetApp.getUi().alert(
    'Modulo creato con successo:\n' + form.getEditUrl()
  );
}
