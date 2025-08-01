/**
 * Adoratio-Szeged
 * Funkciók:
 * 1. sendEmail: Heti email küldése a listára a heti beosztásról (és a telefonszámokról)
 * 2. sendDocuments: Google naptárból naptárnézet és nyomtatható checklista generálása és küldése a felelősöknek
 * 3. sendPersonalNotification: Napi egyéni emlékeztető küldése N. órával az esemény előtt előtted és utánad következőkkel és telefonszámaikkal
 * Elvárt beállítás:
 * - A naptárakban az esemény neve a felelős neve
 * - A naptárakban az esemény leírása a felelős telefonszáma
 * - A naptárakban az esemény megghívottja a felelős e-mail címe
 * Időzítések: 1-2. heti egyszeri időpontra történő időzítéssel; 3-as óránkénti folyamatos időzítéssel. Időzítés a "Triggerek" menüpontban
 * API: https://developers.google.com/apps-script/reference/calendar/calendar
 * Gyuris Gellért
 */

// Globális beállítások:
const DESTINATION_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DESTINATION_FOLDER_ID'); // A mappa a beosztások számára
const SCHEDULE_TEMPLATE     = PropertiesService.getScriptProperties().getProperty('SCHEDULE_TEMPLATE');  // A beosztások sablona zárolt G dokumentum
const INTETIONS_TEMPLATE    = PropertiesService.getScriptProperties().getProperty('INTETIONS_TEMPLATE'); // A szándékok sablona zárolt G dokumentum
const NOTIFICATION_TEMPLATE = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_TEMPLATE'); // A napi értesítési levél sablona zárolt G dokumentum
const NOTIFICATION_DELTA    = PropertiesService.getScriptProperties().getProperty('NOTIFICATION_DELTA'); // A napi értesítés kb. ennyi órával előtte indítja az email-t
const CALENDAR_SILENT       = PropertiesService.getScriptProperties().getProperty('CALENDAR_SILENT');    // Alapértelmezett naptár: csendes ima
const CALENDAR_WORSHIP      = PropertiesService.getScriptProperties().getProperty('CALENDAR_WORSHIP');   // Naptár 2: énekes
const CALENDAR_LOUD         = PropertiesService.getScriptProperties().getProperty('CALENDAR_LOUD');      // Naptár 3: hangos, kötött
const CALENDAR_BIBLE        = PropertiesService.getScriptProperties().getProperty('CALENDAR_BIBLE');     // Naptár 4: Igeolvasás
const CALENDAR_CHURCH       = PropertiesService.getScriptProperties().getProperty('CALENDAR_CHURCH');    // Templomi propramok naptára
const TZ                    = PropertiesService.getScriptProperties().getProperty('TZ');                 // Időzőna
const RECIPIENT_LIST        = PropertiesService.getScriptProperties().getProperty('RECIPIENT_LIST');     // A heti beosztást megkapók e-mail-címe
const RECIPIENT_TEAM        = PropertiesService.getScriptProperties().getProperty('RECIPIENT_TEAM');     // A heti nyomtatási megkapók e-mail-címe
const INTETIONS_FILE        = PropertiesService.getScriptProperties().getProperty('INTETIONS_FILE');     // A szándékokat tartalmazó G táblázat
const START                 = getNextMonday(new Date()) // következő hét hétfője és rá egy hét
const END                   = getEndDate(START);


/**
 * Main: Email küldések
 * Egy funkcióban "Exceeded maximum execution time" hibát okoz, ezért ketté lett választva
 */
function sendEmail() {
  let html = createCalendarText();
  let messageHTML = '<h2>Kedves Adoráló testvérek!</h2><p>A következő hét beosztását alább találjátok.</p><p>Imaszándékaink:</p><ul><li>'
  + getIntentions(true).join('</li><li>')
  + '</li></ul><p> Áldott együttlétet kívánunk nektek az Úr előtt!<p/><p>Ha helyettesítés szükséges, kérjük, keressétek a koordinátort, Aradi Marit <a href="tel:+36204260219">+36204260219</a>.</p><p>Szeretettel: a vezetői csapat</p><hr>';
  GmailApp.sendEmail(
    RECIPIENT_LIST,
    "Következő hét beosztása: " + Utilities.formatDate(START, TZ, "yyyy, w") + ". hét\n",
    stripHtml(messageHTML + html),
    {
      name: 'Adoratio Szeged',
      htmlBody : messageHTML + html
    }
  );
  console.log('Heti beosztás elküldve a listára: ' + stripHtml(messageHTML + html) );
}

function sendDocuments() {
  let fileChecklist   = createChecklistDocument();
  let fileIntentions  = createIntentionsDocument();
  let fileCalendar    = createPrintableCalendarTable();
  messageHTML = "<h2>Helló Marika, Mária, Adorján, Gellért!</h2><p>Íme a heti ellenőrző lista, az áttekintő naptár és a heti imaszándék a szentségimádáshoz. Ezt a három mellékletet kell kinyomtatni és bevinni a hétfői nyitásig...</p><p>Fáradhatatlanul: a gép</p>";
  GmailApp.sendEmail(
    RECIPIENT_TEAM,
    "Nyomtasd ki és vidd el a Jezsikhez: " + Utilities.formatDate(START, TZ, "yyyy, w") + ". hét\n",
    stripHtml(messageHTML),
    {
      name: 'Adoratio Szeged',
      htmlBody : messageHTML,
      attachments: [fileChecklist.getAs(MimeType.PDF), fileCalendar.getAs(MimeType.PDF), fileIntentions.getAs(MimeType.PDF)]
    }
  );
  console.log('Heti beosztás nyomtatandó dokumentumai elküldve (' + RECIPIENT_TEAM + '):' + stripHtml(messageHTML) );
}

function sendPersonalNotification() {
  // több naptár a jelentől a delta órába eső eseményeinek összefűzése és sorbarendezése
  let start = new Date();
  let end = new Date(new Date().setHours(start.getHours() + Number(NOTIFICATION_DELTA)));
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE], start, end);

  events.forEach(function(item) {
    let tag = item.getTag('PersonalNotification');
    /*if (tag == 'processed' ) {
      item.deleteTag('PersonalNotification');
    }*/
    if (tag == null || tag != 'processed' ) {
      let email = getEmailFromDescription(item);
      if ( email != '' ) {
        let message = DocumentApp.openById(NOTIFICATION_TEMPLATE).getBody().getText();
        let difference = item.getStartTime().getTime() - new Date().getTime();
        let previousEvents = getEvents(
          [CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE],
          new Date(new Date(item.getStartTime()).setHours(item.getStartTime().getHours() - 1)),
          new Date(item.getStartTime())
        )
        let nextEvents = getEvents(
          [CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE],
          new Date(item.getEndTime()),
          new Date(new Date(item.getEndTime()).setHours(item.getEndTime().getHours() + 1))
        )
        let type = '';
        switch (item.getOriginalCalendarId()) {
          case CALENDAR_SILENT:
            type = 'Csendes ima';
            break;
          case CALENDAR_WORSHIP:
            type = 'ének hangszeres kísérettel (pl. dicsőítés)';
            break;
          case CALENDAR_LOUD:
            type = 'hangos, kötött ima (pl. zsolozsma, Rózsafüzér)';
            break;
          case CALENDAR_BIBLE:
            type = 'Igehallgatás, a Szentírás szavainak olvasása';
            break;
        }
        message = message.replace("{{TeljesNév}}", item.getTitle());
        message = message.replace("{{DeltaÓra}}", Math.floor((difference % 86400000) / 3600000));
        message = message.replace("{{DeltaPerc}}", Math.round(((difference % 86400000) % 3600000) / 60000));
        message = message.replace("{{DátumÉsIdőpont}}",  Utilities.formatDate(item.getStartTime(), TZ, "yyyy.MM.dd. HH:mm"));
        message = message.replace("{{Előzők}}", ( previousEvents.length > 0 ) ? concatenateEvents(previousEvents) : 'Nincs előtted senki.' )
        message = message.replace("{{Következők}}", ( nextEvents.length > 0 ) ? concatenateEvents(nextEvents) : 'Nincs utánad senki.' )
        message = message.replace("{{Típus}}", type);

        GmailApp.sendEmail(
          email,
          "Várunk a szentségimádásra! (" + Utilities.formatDate(item.getStartTime(), TZ, "yyyy.MM.dd. HH:mm") + ")",
          stripHtml(message),
          {
            name: 'Adoratio Szeged',
            htmlBody : toHtml(message)
          }
        );
        item.setTag('PersonalNotification', 'processed');
        console.log('Értesítés elküldve ' + email + ' számára: ' + message);
      }
    }
  })
}

/**
 * Nyomtathtó ellenőrő lista dokumentumának összeállítása
 */
function createChecklistDocument() {
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE], START, END);
  let templateFile = DriveApp.getFileById(SCHEDULE_TEMPLATE);
  let destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  let fileName =  Utilities.formatDate(START, TZ, "yyyy-MM-dd") + " Heti beosztás, ellenőrzőlista";
  let newFile = templateFile.makeCopy(fileName, destinationFolder);
  let fileToEdit = DocumentApp.openById(newFile.getId());
  let doc = fileToEdit.getBody();
  let previousStart, previousEnd, table, first;

  function setCellAttributes(cell) {
    cell.setPaddingTop(0);
    cell.setPaddingRight(0);
    cell.setPaddingBottom(0);
    cell.setPaddingLeft(0);
  }

  // hét beállítása
  doc.replaceText("{{HétSzáma}}", Utilities.formatDate(START, TZ, "yyyy, w"));
  // események végiglépdelése
  if (events.length > 0) {
    let day = Utilities.formatDate(new Date(), TZ, "d");
    for (i = 0; i < events.length; i++) {
      let event = events[i];
      // nap nevének kiiratása ha más, mint az előző
      if (Utilities.formatDate(event.getStartTime(), TZ, "d") != day) {
        let year = Utilities.formatDate(event.getStartTime(), TZ, "yyyy");
        let dateMonth = getHUNMonth(event.getStartTime().getMonth());
        let dateDayName = getHUNday(event.getStartTime().getDay());
        let dateDay = Utilities.formatDate(event.getStartTime(), TZ, "d");
        doc.appendParagraph(year + ". " + dateMonth + " " + dateDay + "., " + dateDayName).setHeading(DocumentApp.ParagraphHeading.HEADING1);
        // táblázat elindítása minden nap után
        table = doc.appendTable([]);
        first = true;
      }
      // az esemény adatainak beállítása
      let start = Utilities.formatDate(event.getStartTime(), TZ, "HH:mm");
      let end = Utilities.formatDate(event.getEndTime(), TZ, "HH:mm");
      let sTitle = "☐ " + event.getTitle();
      // ha nem különbözik az előzőtől, akkor üresen marad a cella
      let sInterval = (start == previousStart)  ? '' : start + "–" + end ;
      // // ha két időpont között megszakad a folytonosság (és nem közvetlenül a nap neve után vagyunk)
      if (first == false & start != previousStart & previousEnd != start) {
        sInterval = "\n" + sInterval;
        sTitle = "\n" + sTitle;
      }
      // adatok kiiratása a táblázatba
      let tableRow = table.appendTableRow();
      setCellAttributes(tableRow.appendTableCell(sInterval));
      setCellAttributes(tableRow.appendTableCell(sTitle));
      // oszlopok beállítása
      if (table.getNumChildren() == 1) {
        table.setColumnWidth(0, 70);
        table.setBorderWidth(0);
      }
      // következő nap vizsgálatához
      day = Utilities.formatDate(event.getStartTime(), TZ, "d");
      previousEnd = end;
      previousStart = start;
      first = false;
    }
  } else {
    Logger.log('Nincsenek következő események.');
  }
  fileToEdit.saveAndClose();
  return fileToEdit;
}

/**
 * E-mail-ben küldendő lista HTML kódjának összeállítása
 */
function createCalendarText(){
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE], START, END);
  // hanyadik hét
  let html = "<h3>" + Utilities.formatDate(START, TZ, "w") + ". hét</h3>";
  // események lekérdezése a naptárból a dátumok alapján
  if (events.length > 0) {
    let day = Utilities.formatDate(new Date(), TZ, "d");
    // végiglépdelés
    for (i = 0; i < events.length; i++) {
      let event = events[i];
      // nap nevének kiiratása ha más, mint az előző
      if (Utilities.formatDate(event.getStartTime(), TZ, "d") != day) {
        let year = Utilities.formatDate(event.getStartTime(), TZ, "yyyy");
        let dateMonth = getHUNMonth(event.getStartTime().getMonth());
        let dateDayName = getHUNday(event.getStartTime().getDay());
        let dateDay = Utilities.formatDate(event.getStartTime(), TZ, "d");
        html += '<h4>' + year + '. ' + dateMonth + ' ' + dateDay + '., ' + dateDayName + '</h4>';
      }
      // az esemény adatainak a kiiratása
      let start = Utilities.formatDate(event.getStartTime(), TZ, "HH:mm");
      let end = Utilities.formatDate(event.getEndTime(), TZ, "HH:mm");
      let title = event.getTitle();
      html += '<p><span>' + start + '–' + end + '</span> <strong>' + title + '</strong></p>';
           // következő nap vizsgálatához
      day = Utilities.formatDate(event.getStartTime(), TZ, "d");
    }
  } else {
    Logger.log('Nincsenek következő események.');
  }
  return html;
}

/**
 * Nyomtatható áttekintő naptár összeállítása (lásd még: CalendarTemplate.html)
 */
function createPrintableCalendarTable() {
  let t = HtmlService.createTemplateFromFile('CalendarTemplate');
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE, CALENDAR_CHURCH], START, END);
  // egész napos (vagy több, mint egy napig tartó események) eltávolítása
  let i = events.length;
  while (i--) {
    if (events[i].isAllDayEvent() || Math.abs(events[i].getStartTime() - events[i].getEndTime()) / 36e5 > 24) {
      events.splice(i, 1);
    }
  }
  // feldolgozás
  t.currentWeek = Utilities.formatDate(START, TZ, "yyyy, w");
  t.days = fillDays(); //events.end.setDate(events.end.getDate() - 1)
  t.gridTemplateAreas = fillGridTemplateAreas(events);
  t.eventsByAreas = fillAreas(t.gridTemplateAreas, events);
  let htmlBody = t.evaluate().getContent();
  let fileName =  Utilities.formatDate(START, TZ, "yyyy-MM-dd") + " Heti áttekintő naptár";
  let destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  let blob = Utilities.newBlob(htmlBody, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
  let file = destinationFolder.createFile(blob);
  return file;
}

// események dobozba rendezése
function fillAreas(areas, data) {
  function getEventsForArea(el) {
    let array = [];
    for (let i = 0; i < data.length; i++) {
      if (el.endsWith(Utilities.formatDate(data[i].getStartTime(), TZ, "yyyyMMdd-HH"))) {
        array.push(data[i]);
        data.splice(i, 1);
        i--;
      }
    }
    return array;
  }
  let names = {};
  for (let i = 1; i < areas.length; i++) {
    for (let j = 1; j < areas[i].length;j++) {
      if (typeof names[areas[i][j]] == 'undefined') {
        names[areas[i][j]] = getEventsForArea(areas[i][j]);
      }
    }
  }
  return names;
}

// grid-template-areas CSS feltöltési adatai
function fillGridTemplateAreas(data) {
  // alapértelmezet rácsozat kialakítása: minden lehetséges órának egy cella: day-2023-12-25-00 formáumkóddal
  let array = [ [".", "day1", "day2", "day3", "day4", "day5", "day6", "day7"]];
  for (let i = 1; i <= 24; i++) {
    array.push(["hour" + to2Number(i-1)]);
    let current = new Date(START);
    for (current; current < END; current.setDate(current.getDate() + 1)) {
      array[i].push("day-" + Utilities.formatDate(current, TZ, "yyyyMMdd") + "-" + to2Number(i-1));
    }
  }
  // felülírás ott, ahol két vagy több óra össze van vonva (kezdő és befejező érték nagyobb, mint 1 óra)
  for (i = 0; i < data.length; i++) {
    let event = data[i];
    let deltaHour = Math.abs(event.getStartTime() - event.getEndTime()) / 36e5;
    if (deltaHour > 1 ) {
      let name = "day-" + Utilities.formatDate(event.getStartTime(), TZ, "yyyyMMdd-HH");
      let index2D = indexOf2D(array, name);
      // a kezdő formátumkód felülírása
      for (let j = 1; j < deltaHour; j++) {
        array[index2D[0]+j][index2D[1]] = name;
      }
    }
  }
  return array;
}

// dátumok feltöltése
function fillDays() {
  let array = [];
  let current = new Date(START);
  for (current; current < END; current.setDate(current.getDate() + 1)) {
    array.push(Utilities.formatDate(current, TZ, "yyyy. MM. dd.") + ", " + getHUNday(current.getDay()).toUpperCase());
  }
  return array;
}

/**
 * Szándékok
 */
function getIntentions(addLink) {
  let intention = [];
  let ss = SpreadsheetApp.openById(INTETIONS_FILE);
  // fő szándék lekérése
  let sheet1 = ss.getSheets()[0];
  let lastRow = sheet1.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow()
  let i1 = sheet1.getRange(`B${lastRow}`).getCell(1,1).getValue();
  if (i1 != '') {
    intention.push( [sheet1.getName()] + ': ' + i1 );
  }
  // a pápai havi szándék lekérdezése
  let sheet2 = ss.getSheets()[1];
  let month = new Date(START);
  month.setDate(1);
  let index2 = sheet2.getDataRange().getValues().findIndex(row => new Date(row[0]).toDateString() == month.toDateString());
  let i2 = sheet2.getRange(`B${index2+1}`).getCell(1,1).getValue();
  let i2link = sheet2.getRange(`C${index2+1}`).getCell(1,1).getValue();
  if (i2 != '') {
    intention.push(
      [sheet2.getName()]
      + ' ('
      + month.getFullYear()
      + '. '
      + getHUNMonth(month.getMonth())
      +'): '
      + i2
      + (addLink && i2link != '' ? ` <a href="${i2link}">${i2link}</a>` : '')
      );
  }
  // Az Adoratio Szeged saját heti szándékának lekérése
  let sheet3 = ss.getSheets()[2];
  let index3 = sheet3.getDataRange().getValues().findIndex(row => new Date(row[0]).toDateString() == START.toDateString());
  let i3 = sheet3.getRange(`B${index3+1}`).getCell(1,1).getValue();
  if (i3 != '') {
    intention.push( [sheet3.getName()] + ': ' + i3 );
  }
  // visszatérés
  return intention;
}
function createIntentionsDocument() {
  let templateFile = DriveApp.getFileById(INTETIONS_TEMPLATE);
  let destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  let fileName =  Utilities.formatDate(START, TZ, "yyyy-MM-dd") + " Heti szándékok";
  let newFile = templateFile.makeCopy(fileName, destinationFolder);
  let fileToEdit = DocumentApp.openById(newFile.getId());
  let doc = fileToEdit.getBody();

  doc.replaceText("{{HétSzáma}}", Utilities.formatDate(START, TZ, "yyyy, w"));
  doc.replaceText("{{Szándékok}}", getIntentions().join('\n\n'));

  fileToEdit.saveAndClose();
  return fileToEdit;
}

/**
 * Napi értesítések
 */
function formatEvent(event) {
  let start = Utilities.formatDate(event.getStartTime(), TZ, "HH:mm");
  let end = Utilities.formatDate(event.getEndTime(), TZ, "HH:mm");
  let title = event.getTitle();
  let phone = getPhoneFromDescription(event);
  return '' + start + '–' + end + ' ' + title + ' (<a href="tel:' + phone + '">' + phone +'</a>)';
}

function concatenateEvents(events){
  let str = '<ul>';
  events.forEach(function(item) {
    str += '<li>' + formatEvent(item) + '</li>';
  });
  str += '</ul>'
  return str;
}

/**
 * Közös segédfunkciók
 */

function getPhoneFromDescription(event) {
  let results = stripHtml(event.getDescription()).match(/\+[\d-()\s]{8,15}/g);
  if ( results.length > 0 ) return results[0];
  return '';
}

function getEmailFromDescription(event) {
  let results = stripHtml(event.getDescription()).match(/[\w\-\.]+@([\w-]+\.)+[\w-]{2,4}/g);
  if (results.length > 0 ) return results[0];
  return '';
}

// események lekérdezése a naptárakból
function getEvents(aCalendar, start, end){
  // több naptár összefűzése és sorbarendezése
  let events = [];
  aCalendar.forEach(function(item) {
    events = events.concat(CalendarApp.getCalendarById(item).getEvents(start, end));
  })
  events.sort((a, b) => {return a.getStartTime().valueOf() - b.getStartTime().valueOf()});
  return events;
}

function getEndDate(date) {
  let week = new Date(date);
  week.setDate(week.getDate() + (((1 + 7 - week.getDay()) % 7) || 7));
  return week;
}

// a mai dátumhoz legközelebbi hétfő 00:00 meghatározása
function getNextMonday(date) {
  const dateCopy = new Date(date.getTime());
  const nextMonday = new Date(
    dateCopy.setDate(
      dateCopy.getDate() + ((7 - dateCopy.getDay() + 1) % 7 || 7),
    ),
  );
  nextMonday.setHours(0, 0, 0);
  return nextMonday;
}

// kétdimenziós tömbkeresés
function indexOf2D(array, item) {
  for (let i = 0; i < array.length; i++) {
    let position = array[i].indexOf(item);
    if (position > -1) return [i, position];
  }
  return -1;
}

// két karakterre formázás
function to2Number(n) {
    return (n < 10) ? '0' + n.toString() : n.toString();
}

// eltávolítja a HTML címkéket úgy, hogy ügyel a blokk szintű elemek új sorral való helyes helyettesítésére
function stripHtml(sHtml) {
  return sPlainText = sHtml.replace(/(<)(address|article|aside|blockquote|canvas|div|dl|dt|fieldset|figcaption|figure|footer|form|h1|h2|h3|h4|h5|h6|header|hr|main|nav|noscript|ol|pre|section|table|tfoot|video)/gi, "\n\n$1$2").replace(/(<)(ul|p|br\/?)/gi, "\n$1$2").replace(/(<)(li|dd)/gi, "\n - $1$2").replace(/<\/?[^>]+(>|$)/gi, "");
}

// a sorvégeket <br/>-re cseréli
function toHtml(str) {
   return str.replace(/\n/gi, "<br/>");
}

function getHUNMonth(n) {
  const months = new Array("január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december");
  return months[n];
}

function getHUNday(n) {
  const days = new Array("vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat", "vasárnap")
  return days[n];
}
