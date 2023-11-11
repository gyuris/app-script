/**
 * Egy Google űrlapon történő válasz rögzítésére automatikus email küldése egy Google dokumentum sablonként való felhasználásával.
 * Az űrlap válaszait egy táblázatba kell legyűjteni, ezt a szkirptet ebben a táblázatban kell elindítani „űrlap bekülése” eseményre.
 * Gyuris Gellért
 */

const EMAIL_TEMPLATE_FILE_ID    = "1BSJjUwcH7gFsNDCkYsv70ikWhDOddD0R3rVkRIOdkxs"; // Az email sablon azonosítója, érdemes zárolttá tenni az elkészülte után
const COMPLETE_TEXT             = "Visszajelzés elküldve"; // A táblázat utolsó oszlopában megjelenő szöveg, ha az email ki lett küldve. Ezt az utolsó oszlopot fel kell venni, pl. "Visszajelzés"
const EMAIL_FROM_EMAIL          = "gellert.gyuris@gmail.com"; // Feladó
const EMAIL_FROM_NAME           = "Gyuris Gellért automata"; // Feladó neve
const EMAIL_REPLY_TO            = "gyuris.gellert@kateketa.hu"; // Visszatérő email
const EMAIL_SUBJECT             = "Adventi flashmob regisztráció: demo és kotta" // Levél tárgya
const RECIPIENT_EMAIL_COLUMN_ID = "E-mail-cím"; // Az email címet tartalmazó oszlop fejléce

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
