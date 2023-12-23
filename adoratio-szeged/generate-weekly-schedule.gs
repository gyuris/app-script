/**
 * Adoratio-Szeged
 * Google naptárból nyomtatható checklista generálása és * email küldése a listára a heti beosztásról és a telefonszámokról
 * Önálló szkript heti egyszeri időponton történő időzítéssel
 * Gyuris Gellért
 */

// Időzítés a "Triggerek" menüpontban

// Globális beállítások:
const DESTINATION_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DESTINATION_FOLDER_ID'); // A mappa a beosztások számára
const TEMPLATE_FILE_ID      = PropertiesService.getScriptProperties().getProperty('TEMPLATE_FILE_ID'); // A sablonként használt zárolt G dokumentum
const CALENDAR_SILENT       = PropertiesService.getScriptProperties().getProperty('CALENDAR_SILENT');  // Alapértelmezett naptár: csendes ima
const CALENDAR_WORSHIP      = PropertiesService.getScriptProperties().getProperty('CALENDAR_WORSHIP'); // Naptár 2: énekes
const CALENDAR_LOUD         = PropertiesService.getScriptProperties().getProperty('CALENDAR_LOUD');    // Naptár 3: hangos, kötött
const CALENDAR_BIBLE        = PropertiesService.getScriptProperties().getProperty('CALENDAR_BIBLE');   // Naptár 4: Igeolvasás
const CALENDAR_CHURCH       = PropertiesService.getScriptProperties().getProperty('CALENDAR_CHURCH');  // Templomi propramok naptára
const TZ                    = PropertiesService.getScriptProperties().getProperty('TZ');               // Időzőna
const RECIPIENT_LIST        = PropertiesService.getScriptProperties().getProperty('RECIPIENT_LIST');   // A heti beosztást megkapók e-mail-címe
const RECIPIENT_TEAM        = PropertiesService.getScriptProperties().getProperty('RECIPIENT_TEAM');   // A heti nyomtatási megkapók e-mail-címe
const START                 = getNextMonday(new Date()) // következő hét hétfője és rá egy hét
const END                   = getEndDate(START);

/**
 * Main: Email küldések
 */
function sendEmail() {
  let html = createCalendarText();
  let messageHTML = '<h2>Kedves Adoráló testvérek!</h2><p>A következő hét beosztását alább találjátok. Áldott együttlétet kívánunk nektek az Úr előtt!<p/><p>Ha helyettesítés szükséges, kérjük, keressétek a koordinátort, Aradi Marit <a href="tel:+36204260219">+36204260219</a>.</p><p>Szeretettel: a vezetői csapat</p><hr>' + html;
  GmailApp.sendEmail(
    RECIPIENT_LIST,
    "Következő hét beosztása: " + Utilities.formatDate(START, TZ, "yyyy, w") + ". hét\n",
    stripHtml(messageHTML),
    {
      name: 'Adoratio Szeged',
      htmlBody : messageHTML
    }
  );
  let fileChecklist = createChecklistDocument();
  let fileCalendar  = createPrintableCalendarTable();
  messageHTML = "<h2>Helló Marika, Adorján, Gellért!</h2><p>Íme a heti ellenőrző lista és az áttekintő naptár a szentségimádáshoz. Ezt a két mellékletet kell kinyomtatni és bevinni a hétfői nyitásig...</p><p>Fáradhatatlanul: a gép</p>";
  GmailApp.sendEmail(
    RECIPIENT_TEAM,
    "Nyomtasd ki és vidd el a Jezsikhez: " + Utilities.formatDate(START, TZ, "yyyy, w") + ". hét\n",
    stripHtml(messageHTML),
    {
      name: 'Adoratio Szeged',
      htmlBody : messageHTML,
      attachments: [fileChecklist.getAs(MimeType.PDF), fileCalendar.getAs(MimeType.PDF)]
    }
  );
}

/**
 * Nyomtathtó ellenőrő lista dokumentumának összeállítása
 */
function createChecklistDocument(sDate) {
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE]);
  let templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
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
  doc.replaceText("xx", Utilities.formatDate(START, TZ, "yyyy, w"));
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
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE]);
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
      let description = event.getDescription().replace(/<\/?[^>]+(>|$)/g, "");
      html += '<p><span>' + start + '–' + end + '</span> <strong>' + title + '</strong> (<a href="tel:' + description + '">' + description +'</a>)</p>';
           // következő nap vizsgálatához
      day = Utilities.formatDate(event.getStartTime(), TZ, "d");
    }
  } else {
    Logger.log('Nincsenek következő események.');
  }
  //Logger.log(html)
  return html;
}

/**
 * Nyomtatható áttekintő naptár összeállítása (lásd még: CalendarTemplate.html)
 */
function createPrintableCalendarTable() {
  let t = HtmlService.createTemplateFromFile('CalendarTemplate');
  let events = getEvents([CALENDAR_SILENT, CALENDAR_WORSHIP, CALENDAR_LOUD, CALENDAR_BIBLE, CALENDAR_CHURCH]);
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
    //console.log(array)
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
 * Közös segédfunkciók
 */
// események lekérdezése a naptárakból
function getEvents(aCalendar){
  // több naptár összefűzése és sorbarendezése
  let events = [];
  aCalendar.forEach(function(item) {
    events = events.concat(CalendarApp.getCalendarById(item).getEvents(START, END));
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
function stripHtml(sHTML) {
   return sPlainText = sHTML.replace(/(<(h1|h2|h3|h4|h5|h6|hr|div)>)/gi, "\n\n").replace(/(<(p|li)>)/gi, "\n").replace(/(<([^>]+)>)/gi, "");
}

function getHUNMonth(n) {
  const months = new Array("január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december");
  return months[n];
}

function getHUNday(n) {
  const days = new Array("vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat", "vasárnap")
  return days[n];
}
