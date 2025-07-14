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
  const SHEET = SpreadsheetApp.getActiveSheet();
  const CALENDARS_ID = {
    'csendes' : 'adoratio.szeged@gmail.com',
    'hangos, kötött ima (pl. zsolozsma, Rózsafüzér)' : 'd71ab7261481207181b04219a2aa964f68234b19aef0daed659b0ee34aa915bf@group.calendar.google.com',
    'Igehallgatás' : '518f8b15885e20f9801f3e8968a810328f76fba6eb1d7ccb0e2b8dbc69b217f2@group.calendar.google.com',
    'ének, hangszeres kísérettel (pl. dicsőítés)' : '359919ac3b0ae60f349cd7fa3eb4d54527c08259f6f80eb03b9ea3732e3ae684@group.calendar.google.com'
  }
  const STARTROW = 2;
  const COMPLETE = "Esemény bejegyezve";

  let numRows = SHEET.getLastRow(); // Feldolgozandó sorok száma
  let numColumns = SHEET.getLastColumn(); // Utolsó oszlop száma
  let data = SHEET.getRange(STARTROW, 1, numRows-1, numColumns).getValues(); // Adatok lekérdezése

  for (var i = 0; i < data.length; ++i) {
    let row = data[i];
    if (row[numColumns-1] != COMPLETE) {
      let contact = row[1];
      let name = row[2];
      let phone = row[3];
      let rDate = new Date(row[4]);
      let eDate = new Date(new Date(row[4]).setHours( rDate.getHours() + 1 ));
      let type = row[5];

      let calendar = CalendarApp.getCalendarById(CALENDARS_ID[type]);
      calendar.createEvent(name, rDate, eDate, {
        description: phone,
        location: 'Szeged, Szegedi Szent József Templom, Dáni u. 3, 6722',
        guests: contact,
        sendInvites: true
      });
      SHEET.getRange(STARTROW + i, numColumns).setValue(COMPLETE);
    }
  }
}
