/**
 * Adoratio-Szeged
 * Google naptárból nyomtatható checklista generálása és * email küldése a listára a heti beosztásról és a telefonszámokról
 * Önálló szkript heti egyszeri időponton történő időzítéssel
 * Gyuris Gellért
 */

// Időzítés a "Triggerek" menüpontban

// Globális beállítások:
const DESTINATION_FOLDER_ID = '16eM0bclWDkKqwutW6KBfB5EPqz6if9Uc'; // "Heti beosztások" mappa
const TEMPLATE_FILE_ID      = '1BMJDG-KenFPPEC5WnNzFAWjFIQKO-lF2LTjp5MRu-mQ'; // "Heti beosztás, ellenőrzőlista SABLON" zárolt dokumentum
const CALENDAR_SILENT       = 'adoratio.szeged@gmail.com'; // alapértelmezett
const CALENDAR_WORSHIP      = '359919ac3b0ae60f349cd7fa3eb4d54527c08259f6f80eb03b9ea3732e3ae684@group.calendar.google.com';
const CALENDAR_LAUD         = 'd71ab7261481207181b04219a2aa964f68234b19aef0daed659b0ee34aa915bf@group.calendar.google.com';
const CALENDAR_BIBLE        = '518f8b15885e20f9801f3e8968a810328f76fba6eb1d7ccb0e2b8dbc69b217f2@group.calendar.google.com';
const RECIPIENT_LIST        = "adoratio-szeged@googlegroups.com";
const RECIPIENT_TEAM        = "miriamaradi@t-online.hu, jobel@ujevangelizacio.hu, csaladkozpont@gmail.com";
const TZ                    = "Europe/Budapest";

// main
function sendEmail() {
  let events = getCalendarText();
  GmailApp.sendEmail(
    RECIPIENT_LIST,
    "Következő hét beosztása: " + Utilities.formatDate(events.start, TZ, "w") + ". hét\n",
    "Kedves Adoráló testvérek!\n\nA következő hét beosztását alább találjátok. Áldott együttlétet kívánunk nektek az Úr előtt! Ha helyettesítés szükséges, kérjük, keressétek a koordinátort: Aradi Marit +36204260219\n\nSzeretettel: a vezetői csapat\n\n" + events.html,
    {
      name: 'Adoratio Szeged'
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
  let html = Utilities.formatDate(events.start, TZ, "w") + ". hét\n";
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
        html += "\n" + year + ". " + dateMonth + " " + dateDay + "., " + dateDayName + "\n";
      }
      // az esemény adatainak a kiiratása
      let start = Utilities.formatDate(event.getStartTime(), TZ, "HH:mm");
      let end = Utilities.formatDate(event.getEndTime(), TZ, "HH:mm");
      let title = event.getTitle();
      let description = event.getDescription().replace(/<\/?[^>]+(>|$)/g, "");
      html += start + "–" + end + " " + title + " (" + description +")\n";
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

function getHUNMonth(n) {
  const months = new Array("január", "február", "március", "április", "május", "június", "július", "augusztus", "szeptember", "október", "november", "december");
  return months[n];
}

function getHUNday(n) {
  const days = new Array("vasárnap", "hétfő", "kedd", "szerda", "csütörtök", "péntek", "szombat", "vasárnap")
  return days[n];
}
