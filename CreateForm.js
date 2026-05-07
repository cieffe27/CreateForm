function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CreaForm')
    .addItem('Crea Google Modulo', 'createGoogleForm') // ← Crea Form
    .addItem('Crea Doc', 'createGoogleDoc') // ← Doc Risposte
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

function createGoogleDoc() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const spreadsheetName = spreadsheet.getName();
  const spreadsheetId = spreadsheet.getId();
  const doc = DocumentApp.create(spreadsheetName);
  const body = doc.getBody();
  body.clear();
  const YELLOW = '#FFFF00'; // giallo fosforescente
  const colLetters = ['B', 'C', 'D', 'E']; // indici 1-4 → lettere colonna
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const domanda = row[0];
    const risposte = [row[1], row[2], row[3], row[4]];
    const evidenzia = String(row[5]).trim().toUpperCase(); // colonna F
    if (!domanda) continue;
    // Titolo H3
    const heading = body.appendParagraph(String(domanda));
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    // Lista con quadratino per ogni risposta non vuota
    for (let j = 0; j < risposte.length; j++) {
      const risposta = risposte[j];
      if (risposta === '' || risposta === null || risposta === undefined) continue;
      const listItem = body.appendListItem(String(risposta));
      listItem.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
      listItem.setNestingLevel(0);
      // Evidenzia se la lettera in col F corrisponde a questa colonna
      if (evidenzia === colLetters[j]) {
        listItem.editAsText().setBackgroundColor(YELLOW);
      }
    }
    // Riga vuota come spaziatore tra le domande
    body.appendParagraph('');
  }
  doc.saveAndClose();
  // Sposta nella stessa cartella del Google Sheet
  const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  const parentFolders = spreadsheetFile.getParents();
  if (parentFolders.hasNext()) {
    const parentFolder = parentFolders.next();
    DriveApp.getFileById(doc.getId()).moveTo(parentFolder);
  }
  SpreadsheetApp.getUi().alert('Documento creato con successo:\n' + doc.getUrl());
}

