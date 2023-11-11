/**
 * Egy Google űrlapon történő válasz rögzítésére autmatikus eseménybejegyzés létrehozása a naptárban
 * Az űrlap beküldéseit külön táblázatba kell leválasztani, majd ebben a táblázatban egy plusz oszlopot felvenni és az 
 * eseményindítót „űrlap beküldése” eseményre beállítani.
 *
 * Alap: http://wafflebytes.blogspot.com/2017/06/google-script-create-calendar-events.html
 * Hasonló: https://bionicteaching.com/google-calendar-events-via-google-form/
 * API: https://developers.google.com/apps-script/reference/calendar/calendar#createEvent(String,Date,Date,Object)
 * Gyuris Gellért
 */

function createCalendarEvent() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendar = CalendarApp.getCalendarById('adoratio.szeged@gmail.com');

  var startRow = 2;
  var numRows = sheet.getLastRow(); // Number of rows to process
  var numColumns = sheet.getLastColumn();

  var dataRange = sheet.getRange(startRow, 1, numRows-1, numColumns);
  var data = dataRange.getValues();

  var complete = "Esemény bejegyezve";

  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    if (row[8] != complete) {
      var currentCell = sheet.getRange(startRow + i, 9);
      var name = row[2];
      var rDate = new Date(row[3]);
      var eDate = new Date(new Date(row[3]).setHours( rDate.getHours() + 1 ));
      var contact = row[1];
      calendar.createEvent(name, rDate, eDate, {
        description: name + ' (' + row[5] + '): ' + row[4] + '\n' + contact + '\nBeküldve: ' + row[0].toLocaleDateString('hu-HU'),
        location: 'Szeged, Szegedi Szent József Templom, Dáni u. 3, 6722',
        guests: contact,
        sendInvites: true
      });
      currentCell.setValue(complete);
    }
  }
}
