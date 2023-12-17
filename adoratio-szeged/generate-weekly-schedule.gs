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
const CALENDAR_LAUD         = PropertiesService.getScriptProperties().getProperty('CALENDAR_LAUD');    // Naptár 3: hangos, kötött
const CALENDAR_BIBLE        = PropertiesService.getScriptProperties().getProperty('CALENDAR_BIBLE');   // Naptár 4: Igeolvasás
const TZ                    = PropertiesService.getScriptProperties().getProperty('TZ');               // Időzőna
const RECIPIENT_LIST        = PropertiesService.getScriptProperties().getProperty('RECIPIENT_LIST');   // A heti beosztást megkapók e-mail-címe
const RECIPIENT_TEAM        = PropertiesService.getScriptProperties().getProperty('RECIPIENT_TEAM');   // A heti nyomtatási megkapók e-mail-címe

// main
function sendEmail() {
  let events = getCalendarText();
  let messageHTML = '<h2>Kedves Adoráló testvérek!</h2><p>A következő hét beosztását alább találjátok. Áldott együttlétet kívánunk nektek az Úr előtt!<p/><p>Ha helyettesítés szükséges, kérjük, keressétek a koordinátort, Aradi Marit <a href="tel:+36204260219">+36204260219</a>.</p><p>Szeretettel: a vezetői csapat</p><hr>' + events.html;
  GmailApp.sendEmail(
    RECIPIENT_LIST,
    "Következő hét beosztása: " + Utilities.formatDate(events.start, TZ, "w") + ". hét\n",
    stripHtml(messageHTML),
    {
      name: 'Adoratio Szeged',
      htmlBody : messageHTML
    }
  );
  let file = createChecklistDocument();
  GmailApp.sendEmail(
    RECIPIENT_TEAM,
    "Nyomtasd ki és vidd el a Jezsikhez: " + Utilities.formatDate(events.start, TZ, "w") + ". hét\n",
    "Helló Marika, Gellért, Adorján!\nÍme a heti ellenőrző lista a szentségimádáshoz. Ezt a mellékletet kell kinyomtatni...\n\nFáradhatatlanul: a gép\n",
    {
      name: 'Adoratio Szeged',
      attachments: [file.getAs(MimeType.PDF)]
    }
  );
}

function createChecklistDocument(sDate) {
  let events = getEvents();
  let templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
  let destinationFolder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  let fileName =  Utilities.formatDate(events.start, TZ, "yyyy-MM-dd") + " Heti beosztás, ellenőrzőlista";
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
  doc.replaceText("xx", Utilities.formatDate(events.start, TZ, "w"));
  // események végiglépdelése
  if (events.data.length > 0) {
    let day = Utilities.formatDate(new Date(), TZ, "d");
    for (i = 0; i < events.data.length; i++) {
      let event = events.data[i];
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

function getCalendarText(){
  let events = getEvents();
  // hanyadik hét
  let html = "<h3>" + Utilities.formatDate(events.start, TZ, "w") + ". hét</h3>";
  // események lekérdezése a naptárból a dátumok alapján
  if (events.data.length > 0) {
    let day = Utilities.formatDate(new Date(), TZ, "d");
    // végiglépdelés
    for (i = 0; i < events.data.length; i++) {
      let event = events.data[i];
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
  return { html: html, start: events.start };
}

// események lekérdezése a naptárakból
function getEvents(){
  const calendar = CalendarApp.getDefaultCalendar();
  const calendarWorship = CalendarApp.getCalendarById(CALENDAR_WORSHIP);
  const calendarLaud = CalendarApp.getCalendarById(CALENDAR_LAUD);
  const calendarBible = CalendarApp.getCalendarById(CALENDAR_BIBLE);

  // következő hét hétfője és rá egy hét
  let monday = getNextMonday(new Date());
  let week = getNextMonday(new Date());
  week.setDate(week.getDate() + (((1 + 7 - week.getDay()) % 7) || 7));

  // több naptár összefűzése
  let arrayEvents = calendar.getEvents(monday, week);
  arrayEvents = arrayEvents.concat(calendarWorship.getEvents(monday, week));
  arrayEvents = arrayEvents.concat(calendarLaud.getEvents(monday, week));
  arrayEvents = arrayEvents.concat(calendarBible.getEvents(monday, week));
  arrayEvents.sort((a, b) => {return a.getStartTime().valueOf() - b.getStartTime().valueOf()});

  return { start: monday, end: week, data: arrayEvents };
}

// a mai dátumhoz legközelebbi hétfő 00:00 meghatározása
function getNextMonday(date = new Date()) {
  const dateCopy = new Date(date.getTime());
  const nextMonday = new Date(
    dateCopy.setDate(
      dateCopy.getDate() + ((7 - dateCopy.getDay() + 1) % 7 || 7),
    ),
  );
  nextMonday.setHours(0, 0, 0);
  return nextMonday;
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
