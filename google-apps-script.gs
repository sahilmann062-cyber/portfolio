function doPost(e) {
  try {
    var sheet = SpreadsheetApp
      .openById('1Hcby0TtcqyOWBj8fI3IwLvnB0BKzxoGMofUCrujKnKk')
      .getSheets()[0];

    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseError) {
        data = e.parameter || {};
      }
    } else {
      data = e.parameter || {};
    }

    var name = data.name || '';
    var email = data.email || '';
    var message = data.message || '';

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Message']);
    }

    sheet.appendRow([new Date(), name, email, message]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
