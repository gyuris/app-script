/**
 * Egy Google űrlapon történő válasz rögzítésére automatikus email küldése egy Google dokumentum sablonként való felhasználásával.
 * Az űrlap válaszait egy táblázatba kell legyűjteni, ezt a szkirptet ebben a táblázatban kell elindítani „űrlap bekülése” eseményre.
 * Gyuris Gellért
 */

const EMAIL_TEMPLATE_FILE_ID    = PropertiesService.getScriptProperties().getProperty('EMAIL_TEMPLATE_FILE_ID'); // Az email sablon azonosítója, érdemes zárolttá tenni az elkészülte után
const COMPLETE_TEXT             = PropertiesService.getScriptProperties().getProperty('COMPLETE_TEXT'); // A táblázat utolsó oszlopában megjelenő szöveg, ha az email ki lett küldve. Ezt az oszlopot fel kell venni, pl. "Visszajelzés"
const EMAIL_FROM_EMAIL          = PropertiesService.getScriptProperties().getProperty('EMAIL_FROM_EMAIL'); // Feladó e-mail címe
const EMAIL_FROM_NAME           = PropertiesService.getScriptProperties().getProperty('EMAIL_FROM_NAME'); // Feladó neve
const EMAIL_REPLY_TO            = PropertiesService.getScriptProperties().getProperty('EMAIL_REPLY_TO'); // Visszatérő (replay) email cím
const EMAIL_SUBJECT             = PropertiesService.getScriptProperties().getProperty('EMAIL_SUBJECT'); // Levél tárgya
const RECIPIENT_EMAIL_COLUMN_ID = PropertiesService.getScriptProperties().getProperty('RECIPIENT_EMAIL_COLUMN_ID'); // Az email címet tartalmazó oszlop fejléce (a fejlécben lévő szöveg)

function sendEmailRespond() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var numColumns = sheet.getLastColumn(); // Oszlopok száma
  var data = sheet.getRange(1, 1, sheet.getLastRow(), numColumns).getValues();

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
     // változónevek kigyűjtése
    if (i == 0) {
      var keys = row;
    } else if (row[numColumns-1] != COMPLETE_TEXT) {
      // szöveg lekérése a dokumentumból
      var message = DocumentApp.openById(EMAIL_TEMPLATE_FILE_ID).getBody().getText();
      // változók behelyettesítése  a szövegbe
      keys.forEach((key, index) => {
        message = message.replace("{{" + key + "}}", row[index].toString());
      });
      // e-mail küldése
      GmailApp.sendEmail(row[keys.indexOf(RECIPIENT_EMAIL_COLUMN_ID)], EMAIL_SUBJECT, message, {
        from: EMAIL_FROM_EMAIL,
        name: EMAIL_FROM_NAME,
        noReply: true,
        replyTo: EMAIL_REPLY_TO
      })
      // bejegyzés a táblázat utolsó oszlopába, hogy el lett neki küldve
      sheet.getRange(i + 1, numColumns).setValue(COMPLETE_TEXT);
    }
  }
}
