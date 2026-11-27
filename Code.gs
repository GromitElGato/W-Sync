function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Acties')
    .addItem('Update Schoenen', 'updateShoesDistance')
    .addItem('Initialiseer Schoenen', 'initialiseerSchoenenTracking')
    .addItem('Installeer wijzigingstrigger', 'installeerOnChangeTrigger')
    .addToUi();

  SpreadsheetApp.getActive().getSheets().forEach(function(sheet) {
    updateShoesDistance(sheet);
  });
}


function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const column = e.range.getColumn();

  if (row < 2) return;

  // Kolom C = afstand, kolom E = schoen
  if (column === 3 || column === 5) {
    updateShoesDistance(sheet);
  }
}


/*
 * Wordt aangeroepen door de installeerbare onChange-trigger.
 * Zo worden schoenentotalen ook herberekend als een rij wordt verwijderd.
 */
function onChange(e) {
  if (!e) return;

  const t = e.changeType;

  if (t === 'REMOVE_ROW' || t === 'INSERT_ROW' || t === 'EDIT' || t === 'OTHER') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2026');
    if (sheet) {
      updateShoesDistance(sheet);
    }
  }
}


/*
 * Installeer deze functie één keer.
 * Bestaande onChange-triggers worden eerst verwijderd.
 */
function installeerOnChangeTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onChange') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onChange').forSpreadsheet(ss).onChange().create();

  Logger.log('onChange-trigger succesvol geïnstalleerd.');
}


function processRow(sheet, row) {
  const values = sheet.getRange(row, 1, 1, 6).getValues()[0];
  const activityName = values[1];

  if (!activityName || !activityName.toString().toLowerCase().includes('wandeling')) {
    sheet.deleteRow(row);
    return;
  }

  const distance = Number(values[2]) || 0;
  const link = values[5];

  if (link && typeof link === 'string' && link.startsWith('http')) {
    const richLink = SpreadsheetApp.newRichTextValue()
      .setText('Garmin')
      .setLinkUrl(link)
      .build();
    sheet.getRange(row, 6).setRichTextValue(richLink);
  }

  sheet.getRange(row, 3).setValue(Math.round(distance));
  sheet.getRange(row, 3).setNumberFormat('0');

  updateShoesDistance(sheet);
}


/*
 * Bereken de kilometerstand per schoen.
 *
 * Historische kilometerstanden vóór de initialisatie blijven onaangeroerd.
 * Alleen activiteiten vanaf de opgeslagen startregel worden meegerekend.
 * BLANCO wordt nooit als schoen meegerekend.
 */
function updateShoesDistance(sheet) {
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }

  // Alleen tabblad 2026 gebruiken voor schoenen.
  if (sheet.getName() !== '2026') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shoesSheet = ss.getSheetByName('Schoenen');
  if (!shoesSheet) return;

  const properties = PropertiesService.getScriptProperties();
  const baseData = properties.getProperty('SCHOENEN_BASISSTAND');

  // Schoenen zijn nog niet geïnitialiseerd.
  if (!baseData) return;

  const baseKm = JSON.parse(baseData);
  const startRow = Number(properties.getProperty('SCHOENEN_START_RIJ_2026')) || 2;
  const lastRow = sheet.getLastRow();
  const shoeLastRow = shoesSheet.getLastRow();

  if (shoeLastRow < 2) return;

  const shoes = shoesSheet.getRange(2, 1, shoeLastRow - 1, 2).getValues();

  // Nieuwe kilometers berekenen
  const addedKm = {};

  if (lastRow >= startRow) {
    const activities = sheet.getRange(startRow, 3, lastRow - startRow + 1, 3).getValues();

    activities.forEach(function(row) {
      const distanceMeters = Number(row[0]) || 0;
      const shoeName = String(row[2] || '').trim();

      // Geen schoen of BLANCO: activiteit telt niet mee.
      if (!shoeName || shoeName === 'BLANCO' || distanceMeters <= 0) return;

      if (!addedKm[shoeName]) addedKm[shoeName] = 0;
      addedKm[shoeName] += distanceMeters / 1000;
    });
  }

  // Basisstand + nieuwe kilometers
  const result = shoes.map(function(row) {
    const shoeName = String(row[0] || '').trim();

    if (!shoeName) return [''];

    // BLANCO blijft altijd 0.
    if (shoeName === 'BLANCO') return [0];

    const startKm = Number(baseKm[shoeName]) || 0;
    const newKm = Number(addedKm[shoeName]) || 0;

    return [Math.round((startKm + newKm) * 100) / 100];
  });

  shoesSheet.getRange(2, 2, result.length, 1).setValues(result);
}


/*
 * Geeft de schoenenlijst terug aan de webapp (Schoenen!A2:A).
 * Nieuwe schoenen verschijnen automatisch in de webapp.
 */
function getShoesList() {
  const shoesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schoenen');
  if (!shoesSheet) return [];

  const lastRow = shoesSheet.getLastRow();
  if (lastRow < 2) return [];

  const values = shoesSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const shoes = [];

  values.forEach(function(row) {
    const shoe = String(row[0] || '').trim();
    if (shoe) shoes.push(shoe);
  });

  return shoes;
}


/*
 * Geeft per schoen de totale kilometers terug aan de webapp
 * (Schoenen!A = naam, Schoenen!B = kilometers). BLANCO wordt overgeslagen.
 */
function getShoesTotals() {
  const shoesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Schoenen');
  if (!shoesSheet) return [];

  const lastRow = shoesSheet.getLastRow();
  if (lastRow < 2) return [];

  const out = [];

  shoesSheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach(function(row) {
    const name = String(row[0] || '').trim();
    if (!name || name === 'BLANCO') return;
    out.push({ n: name, km: Number(row[1]) || 0 });
  });

  return out;
}


/*
 * Datum (yyyy-MM-dd) van de laatste wandeling: de nieuwste datum in kolom A
 * van het meest recente jaartabblad (2026, 2027, ...).
 */
function getLastWalkDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const yearSheets = ss.getSheets()
    .filter(function(sh) { return /^\d{4}$/.test(sh.getName()); })
    .sort(function(a, b) { return Number(b.getName()) - Number(a.getName()); });

  for (let i = 0; i < yearSheets.length; i++) {
    const sh = yearSheets[i];
    const lastRow = sh.getLastRow();
    if (lastRow < 2) continue;

    let newest = null;
    sh.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function(row) {
      const v = row[0];
      if (v instanceof Date && !isNaN(v.getTime()) && (!newest || v > newest)) newest = v;
    });

    if (newest) return Utilities.formatDate(newest, tz, 'yyyy-MM-dd');
  }

  return '';
}


/*
 * Alle wandelingen uit de jaartabbladen (2026, 2027, ...), nieuwste eerst.
 * A = datum + tijd, B = activiteitsnaam, C = afstand in meter, F = link (Garmin of Strava).
 */
function getWalks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const out = [];

  ss.getSheets().forEach(function(sh) {
    if (!/^\d{4}$/.test(sh.getName())) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const n = lastRow - 1;
    const vals = sh.getRange(2, 1, n, 3).getValues();
    const rich = sh.getRange(2, 6, n, 1).getRichTextValues();

    for (let i = 0; i < n; i++) {
      const date = vals[i][0];
      if (!(date instanceof Date) || isNaN(date.getTime())) continue;

      out.push({
        t: date.getTime(),
        d: Utilities.formatDate(date, tz, "yyyy-MM-dd'T'HH:mm"),
        m: Math.round(Number(vals[i][2]) || 0),
        n: String(vals[i][1] || '').trim(),
        u: linkFromRichText_(rich[i][0])
      });
    }
  });

  out.sort(function(a, b) { return b.t - a.t; });

  return out.map(function(w) { return { d: w.d, m: w.m, n: w.n, u: w.u }; });
}


function linkFromRichText_(rt) {
  if (!rt) return '';

  let url = rt.getLinkUrl();

  if (!url) {
    rt.getRuns().forEach(function(run) {
      if (!url && run.getLinkUrl()) url = run.getLinkUrl();
    });
  }

  if (!url) {
    const text = String(rt.getText() || '').trim();
    if (/^https?:\/\//i.test(text)) url = text;
  }

  return url || '';
}


/*
 * Start Garmin-sync met de gekozen schoen.
 * De keuze wordt tijdelijk opgeslagen in Script Properties.
 * BLANCO wordt intern als lege keuze opgeslagen.
 */
function startGarminSync(selectedShoe) {
  const shoe = String(selectedShoe || '').trim();
  const properties = PropertiesService.getScriptProperties();

  if (shoe && shoe !== 'BLANCO') {
    properties.setProperty('GARMIN_SELECTED_SHOE', shoe);
  } else {
    properties.setProperty('GARMIN_SELECTED_SHOE', '');
  }

  const token = properties.getProperty('GITHUB_TOKEN');

  const url =
    'https://api.github.com/repos/' +
    'GromitElGato/garmin-walk-gsheets-sync/' +
    'actions/workflows/test-garmin.yml/dispatches';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
      Accept: 'application/vnd.github+json'
    },
    contentType: 'application/json',
    payload: JSON.stringify({ ref: 'main' }),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 204) {
    throw new Error(
      'GitHub fout: ' + response.getResponseCode() + ' - ' + response.getContentText()
    );
  }

  return 'Garmin Sync gestart.';
}


/*
 * API voor de app (GitHub Pages).
 *
 * Elke aanvraag moet de juiste app-sleutel meesturen.
 * De sleutel maak je één keer aan met maakAppSleutel().
 */
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (!p.action) {
    return ContentService.createTextOutput('Garmin Sync API');
  }

  try {
    const expected = PropertiesService.getScriptProperties().getProperty('APP_SLEUTEL');

    if (!expected) {
      return createJsonResponse({
        success: false,
        error: 'Er is nog geen app-sleutel. Voer maakAppSleutel() uit in Apps Script.'
      });
    }

    if (!p.key || p.key !== expected) {
      return createJsonResponse({ success: false, code: 'BAD_KEY', error: 'Ongeldige sleutel.' });
    }

    if (p.action === 'shoes') {
      return createJsonResponse({ success: true, shoes: getShoesList() });
    }

    if (p.action === 'totals') {
      return createJsonResponse({ success: true, shoes: getShoesTotals(), last: getLastWalkDate() });
    }

    if (p.action === 'walks') {
      return createJsonResponse({ success: true, walks: getWalks() });
    }

    if (p.action === 'sync') {
      const message = startGarminSync(p.shoe);
      return createJsonResponse({ success: true, message: message });
    }

    return createJsonResponse({ success: false, error: 'Onbekende actie.' });

  } catch (error) {
    return createJsonResponse({ success: false, error: String(error.message || error) });
  }
}


/*
 * Voer deze functie ÉÉN keer uit (Uitvoeren-knop in de editor).
 * De sleutel verschijnt in het Uitvoeringslogboek onderaan.
 */
function maakAppSleutel() {
  const key = Utilities.getUuid().replace(/-/g, '').substring(0, 20);
  PropertiesService.getScriptProperties().setProperty('APP_SLEUTEL', key);
  Logger.log('Jouw app-sleutel: ' + key);
}


/* Toont de bestaande sleutel opnieuw in het Uitvoeringslogboek. */
function toonAppSleutel() {
  Logger.log('Jouw app-sleutel: ' + PropertiesService.getScriptProperties().getProperty('APP_SLEUTEL'));
}


function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const properties = PropertiesService.getScriptProperties();
    const expectedToken = properties.getProperty('GARMIN_SYNC_TOKEN');

    if (!data.token || data.token !== expectedToken) {
      return createJsonResponse({ success: false, error: 'Ongeldige synchronisatietoken.' });
    }

    const sheetName = data.sheet || '2026';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (!sheet) {
      return createJsonResponse({
        success: false,
        error: 'Werkblad "' + sheetName + '" bestaat niet.'
      });
    }

    const activityId = String(data.activity_id || '');
    const activityName = String(data.activity_name || 'Wandelen');
    const startTime = data.start_time;
    const distance = Number(data.distance) || 0;
    const duration = Number(data.duration) || 0;
    const link = String(data.link || '');

    if (!activityId || !startTime || !link) {
      return createJsonResponse({
        success: false,
        error: 'Verplichte Garmin-gegevens ontbreken.'
      });
    }

    const date = new Date(startTime);

    if (isNaN(date.getTime())) {
      return createJsonResponse({ success: false, error: 'Ongeldige startdatum.' });
    }

    const cutoffDate = new Date(2026, 8, 15, 0, 0, 0);

    if (date < cutoffDate) {
      return createJsonResponse({
        success: true,
        added: false,
        message: 'Garmin-activiteit vóór 15/09/2026 wordt niet geïmporteerd.',
        activity_id: activityId
      });
    }

    // ---------------- GARMIN HISTORIEK ----------------

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName('GarminHistoriek');

    if (!historySheet) {
      historySheet = ss.insertSheet('GarminHistoriek');
      historySheet.getRange('A1').setValue('Garmin Activity ID');
    }

    const historyLastRow = historySheet.getLastRow();

    if (historyLastRow >= 2) {
      const historyIds = historySheet.getRange(2, 1, historyLastRow - 1, 1).getValues();
      const alreadyExists = historyIds.some(function(row) {
        return String(row[0]) === activityId;
      });

      if (alreadyExists) {
        return createJsonResponse({
          success: true,
          added: false,
          message: 'Garmin-activiteit staat in historiek en wordt niet opnieuw toegevoegd.',
          activity_id: activityId
        });
      }
    }

    // ---------------- GESELECTEERDE SCHOEN ----------------

    let selectedShoe = String(properties.getProperty('GARMIN_SELECTED_SHOE') || '').trim();

    // BLANCO mag nooit als echte schoen in de activiteit terechtkomen.
    if (selectedShoe === 'BLANCO') selectedShoe = '';

    // ---------------- DUUR ----------------

    const totalSeconds = Math.round(duration);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    const durationText =
      hours + ':' + String(minutes).padStart(2, '0') + ':' + String(seconds).padStart(2, '0');

    // ---------------- NIEUWE RIJ ----------------
    // A = datum, B = activiteit, C = afstand, D = duur, E = schoen, F = Garmin-link

    const newRow = sheet.getLastRow() + 1;

    sheet.getRange(newRow, 1, 1, 5).setValues([[
      date,
      activityName,
      Math.round(distance),
      durationText,
      selectedShoe
    ]]);

    // Opmaak van vorige rij (A t/m D) overnemen
    if (newRow > 2) {
      for (let col = 1; col <= 4; col++) {
        sheet.getRange(newRow - 1, col)
          .copyTo(sheet.getRange(newRow, col), { formatOnly: true });
      }
    }

    // ---------------- GARMIN-LINK IN KOLOM F ----------------

    const richLink = SpreadsheetApp.newRichTextValue()
      .setText('Garmin')
      .setLinkUrl(link)
      .build();

    sheet.getRange(newRow, 6).setRichTextValue(richLink);

    // Opmaak van F3 overnemen
    sheet.getRange(3, 6).copyTo(sheet.getRange(newRow, 6), { formatOnly: true });

    // Garmin-link opnieuw instellen
    sheet.getRange(newRow, 6).setRichTextValue(richLink);

    // ---------------- SCHOEN-DROPDOWN ----------------

    const shoeValidation = sheet.getRange('E2').getDataValidation();

    if (shoeValidation) {
      sheet.getRange(newRow, 5).setDataValidation(shoeValidation);
    }

    // ---------------- SCHOENEN BIJWERKEN ----------------

    updateShoesDistance(sheet);

    // ---------------- ACTIVITEIT REGISTREREN ----------------

    historySheet.appendRow([activityId]);

    // ---------------- WHATSAPP ----------------

    const km = (Math.round(distance) / 1000).toFixed(2);

    Logger.log('Garmin-activiteit toegevoegd. WhatsApp wordt nu verzonden...');

    stuurWhatsAppNaGarmin(km);

    return createJsonResponse({
      success: true,
      added: true,
      row: newRow,
      activity_id: activityId
    });

  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  }
}


function stuurWhatsAppNaGarmin(km) {
  const message = '🦔 ' + km + ' km gewandeld 🦔';

  Logger.log('WhatsApp bericht: ' + message);

  try {
    const resultaat = sendWhatsAppToGroup(message);
    Logger.log('WhatsApp verzonden: ' + JSON.stringify(resultaat));
    return resultaat;
  } catch (error) {
    Logger.log('WhatsApp verzenden mislukt: ' + error.toString());
    return null;
  }
}


function sendWhatsAppToGroup(message) {
  // WERKENDE GREEN API-GEGEVENS
  const instanceId = '7105245521';
  const token = '529c3b363f4a4ca48563ce003d47622d057998facdf44b8793';
  const chatId = '120363419720592353@g.us';

  const url =
    'https://api.green-api.com/waInstance' + instanceId + '/sendMessage/' + token;

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chatId: chatId, message: message }),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('Green API HTTP-status: ' + statusCode);
  Logger.log('Green API antwoord: ' + responseText);

  if (statusCode !== 200) {
    throw new Error('Green API fout ' + statusCode + ': ' + responseText);
  }

  return JSON.parse(responseText);
}


function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


function maakNieuwJaarTab(jaar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const naam = String(jaar);

  if (ss.getSheetByName(naam)) return;

  const vorigeSheet = ss.getSheetByName(String(Number(jaar) - 1));

  if (vorigeSheet) {
    const nieuweSheet = vorigeSheet.copyTo(ss);
    nieuweSheet.setName(naam);

    const lastRow = nieuweSheet.getLastRow();

    if (lastRow > 1) {
      nieuweSheet
        .getRange(2, 1, lastRow - 1, nieuweSheet.getLastColumn())
        .clearContent();
    }

    updateShoesDistance(nieuweSheet);
    return;
  }

  ss.insertSheet(naam);
}


function autoUpdateAllSheets() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(sheet) {
    if (/^\d{4}$/.test(sheet.getName())) {
      updateShoesDistance(sheet);
    }
  });
}


function verwijderDubbeleGarminActiviteiten() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GarminHistoriek');
  if (!historySheet) return;

  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) return;

  const values = historySheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const unique = [];
  const seen = {};

  values.forEach(function(row) {
    const id = String(row[0]);

    if (id && !seen[id]) {
      seen[id] = true;
      unique.push([id]);
    }
  });

  historySheet.getRange(2, 1, lastRow - 1, 1).clearContent();

  if (unique.length > 0) {
    historySheet.getRange(2, 1, unique.length, 1).setValues(unique);
  }
}


function verwijderOudeGarminActiviteiten() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GarminHistoriek');
  if (!historySheet) return;

  const lastRow = historySheet.getLastRow();
  if (lastRow < 2) return;

  const values = historySheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const rowsToKeep = [];

  values.forEach(function(row) {
    if (row[0]) rowsToKeep.push(row);
  });

  historySheet.getRange(2, 1, lastRow - 1, 1).clearContent();

  if (rowsToKeep.length > 0) {
    historySheet.getRange(2, 1, rowsToKeep.length, 1).setValues(rowsToKeep);
  }
}


function initialiseerGarminHistoriek() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GarminHistoriek');

  if (!sheet) {
    sheet = ss.insertSheet('GarminHistoriek');
  }

  sheet.getRange('A1').setValue('Garmin Activity ID');
}


function initialiseerSchoenenTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('2026');
  const shoesSheet = ss.getSheetByName('Schoenen');

  if (!sheet || !shoesSheet) {
    throw new Error('Tabblad 2026 of Schoenen ontbreekt.');
  }

  const properties = PropertiesService.getScriptProperties();
  const shoeLastRow = shoesSheet.getLastRow();

  if (shoeLastRow < 2) {
    throw new Error('Er staan geen schoenen in het tabblad Schoenen.');
  }

  const shoes = shoesSheet.getRange(2, 1, shoeLastRow - 1, 2).getValues();
  const baseKm = {};

  shoes.forEach(function(row) {
    const shoeName = String(row[0] || '').trim();
    if (shoeName) {
      baseKm[shoeName] = Number(row[1]) || 0;
    }
  });

  // Huidige kilometerstanden worden eenmalig als basisstand opgeslagen.
  properties.setProperty('SCHOENEN_BASISSTAND', JSON.stringify(baseKm));

  // Alles wat nu al in 2026 staat, beschouwen we als bestaande data.
  properties.setProperty('SCHOENEN_START_RIJ_2026', String(sheet.getLastRow() + 1));

  Logger.log('Schoenen-tracking geïnitialiseerd.');
  Logger.log('Basisstanden: ' + JSON.stringify(baseKm));
  Logger.log('Nieuwe activiteiten vanaf rij ' + (sheet.getLastRow() + 1));
}


function testWhatsApp() {
  sendWhatsAppToGroup('🦔 TESTBERICHT vanuit Google Apps Script 🦔');
}
