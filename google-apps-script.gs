// Script Google Apps Script à coller dans votre Google Sheet
// Extensions > Apps Script > coller ce code > Déployer en tant qu'application Web

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      var headers = ['Date & heure', 'Prénom', 'Nom', 'Email', 'Téléphone', 'Dates du séjour', 'Voyageurs'];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#d9e1f2')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    }

    var p = e.parameter;
    var dateStr = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');

    sheet.appendRow([
      dateStr,
      p.prenom  || '',
      p.nom     || '',
      p.email   || '',
      p.tel     || '',
      p.dates   || '',
      p.guests  || ''
    ]);

    sheet.autoResizeColumns(1, 7);

    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.toString());
  }
}

// Fonction de test : ouvrez l'éditeur et cliquez sur Exécuter > testDoPost
function testDoPost() {
  var fakeEvent = {
    parameter: {
      prenom: 'Jean',
      nom: 'Dupont',
      email: 'jean@test.fr',
      tel: '06 00 00 00 00',
      dates: '15/01/2027 → 22/01/2027',
      guests: '2',
      ts: new Date().toISOString()
    }
  };
  var result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
