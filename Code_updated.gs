/**
 * KASSEN-BACKEND FÜR GOOGLE SHEETS
 * ---------------------------------
 * Diese Datei kommt komplett in den Apps-Script-Editor deines Google Sheets
 * (Erweiterungen → Apps Script) — ersetzt den kompletten bisherigen Inhalt.
 *
 * Erwartet VIER Tabellenblätter im Sheet:
 *   "Sales"    Kopfzeile: sale_id | timestamp | zahlungsart | huette | artikel | menge | einzelpreis | zeilensumme | storniert | gebucht_von | storniert_von
 *              -> jede Buchung wird zeilenweise pro Artikel eingetragen, gut lesbar, keine JSON-Klumpen
 *              -> Spalte 9 "storniert" (TRUE/leer) markiert stornierte Buchungen
 *              -> NEU: Spalte 10 "gebucht_von" = Gerätename/Person, die gebucht hat
 *              -> NEU: Spalte 11 "storniert_von" = Gerätename/Person, die storniert hat
 *   "Summary"  Kopfzeile: artikel | menge | umsatz
 *   "Payments" Kopfzeile: zahlungsart | umsatz
 *   "Products" Kopfzeile: id | category | name | price | emoji
 *
 * NEU in dieser Version:
 *   - action=recent (GET): letzte 30 Buchungen fürs Storno-Fenster in der App
 *   - type=cancel_sale (POST): storniert eine Buchung, zieht Summary/Payments zurück,
 *     protokolliert in Spalte 11, WER storniert hat (vom Gerätenamen aus der App)
 *   - jede Buchung trägt jetzt mit, WELCHES GERÄT/WER gebucht hat (Spalte 10)
 *   - resetAllBookings(): über Menü "Kasse > Alle Buchungen zurücksetzen" ausführbar
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Kasse')
    .addItem('Alle Buchungen zurücksetzen (neues Event)', 'resetAllBookings')
    .addToUi();
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const action = e && e.parameter ? e.parameter.action : null;

  if (!action) {
    return jsonOut({ info: 'Backend läuft. Nutze ?action=summary, ?action=products oder ?action=recent' });
  }

  if (action === 'summary') {
    const sh = ss.getSheetByName('Summary');
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1)
      .filter(r => r[0])
      .map(r => ({ name: r[0], qty: r[1], revenue: r[2] }));
    return jsonOut(rows);
  }

  if (action === 'payments') {
    const sh = ss.getSheetByName('Payments');
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1)
      .filter(r => r[0])
      .map(r => ({ method: r[0], revenue: r[1] }));
    return jsonOut(rows);
  }

  if (action === 'products') {
    const sh = ss.getSheetByName('Products');
    const data = sh.getDataRange().getValues();
    const rows = data.slice(1)
      .filter(r => r[0])
      .map(r => ({ id: r[0], category: r[1], name: r[2], price: r[3], emoji: r[4] }));
    return jsonOut(rows);
  }

  if (action === 'recent') {
    return jsonOut(getRecentSales_(ss, 30));
  }

  return jsonOut({ error: 'unbekannte action' });
}

// Gruppiert die zeilenweisen Sales-Einträge nach sale_id zurück zu einer
// Buchung (für die Storno-Liste in der App). Neueste zuerst.
function getRecentSales_(ss, limit) {
  const sh = ss.getSheetByName('Sales');
  const data = sh.getDataRange().getValues();
  const bySale = {};
  const order = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const saleId = row[0];
    if (!saleId) continue;
    if (!bySale[saleId]) {
      bySale[saleId] = {
        clientId: saleId,
        ts: row[1] instanceof Date ? row[1].getTime() : row[1],
        payment: row[2],
        hut: row[3],
        total: 0,
        cancelled: !!row[8],
        bookedBy: row[9] || '',
        cancelledBy: row[10] || ''
      };
      order.push(saleId);
    }
    bySale[saleId].total += Number(row[7]) || 0;
    if (row[8]) bySale[saleId].cancelled = true;
    if (row[10]) bySale[saleId].cancelledBy = row[10];
  }
  const sales = order.map(id => bySale[id]).reverse();
  return sales.slice(0, limit || 30);
}

function doPost(e) {
  // LockService verhindert, dass zwei gleichzeitige Buchungen sich gegenseitig überschreiben
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!e || !e.postData) {
      return jsonOut({ error: 'kein postData – diese Funktion muss über einen echten POST-Request aufgerufen werden, nicht manuell im Editor' });
    }
    const body = JSON.parse(e.postData.contents);

    if (body.type === 'sale') {
      const salesSh = ss.getSheetByName('Sales');
      const summarySh = ss.getSheetByName('Summary');
      const paymentsSh = ss.getSheetByName('Payments');
      const saleId = body.clientId;
      const payment = body.payment === 'ec' ? 'ec' : 'bar';
      const bookedBy = body.device || '';

      // Idempotenz: Falls diese Buchung (gleiche clientId) bei einem Retry schon
      // erfolgreich gespeichert wurde, NICHT nochmal eintragen/zählen.
      const existing = salesSh.getDataRange().getValues();
      for (let i = 1; i < existing.length; i++) {
        if (existing[i][0] === saleId) {
          return jsonOut({ ok: true, id: saleId, duplicate: true });
        }
      }

      const ts = new Date(body.ts || Date.now());
      body.items.forEach(item => {
        const lineTotal = item.qty * item.price;
        salesSh.appendRow([saleId, ts, payment, body.hut || '', item.name, item.qty, item.price, lineTotal, '', bookedBy, '']);
        updateSummary(summarySh, item.name, item.qty, lineTotal);
        updatePayments(paymentsSh, payment, lineTotal);
      });

      return jsonOut({ ok: true, id: saleId });
    }

    if (body.type === 'cancel_sale') {
      const salesSh = ss.getSheetByName('Sales');
      const summarySh = ss.getSheetByName('Summary');
      const paymentsSh = ss.getSheetByName('Payments');
      const saleId = body.clientId;
      const cancelledBy = body.device || '';
      const data = salesSh.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[0] !== saleId) continue;
        if (row[8]) continue; // schon storniert, nichts doppelt zurückziehen
        found = true;
        const payment = row[2];
        const productName = row[4];
        const qty = Number(row[5]) || 0;
        const lineTotal = Number(row[7]) || 0;
        salesSh.getRange(i + 1, 9).setValue(true);     // Spalte 9 = storniert
        salesSh.getRange(i + 1, 11).setValue(cancelledBy); // Spalte 11 = storniert_von
        updateSummary(summarySh, productName, -qty, -lineTotal);
        updatePayments(paymentsSh, payment, -lineTotal);
      }
      if (!found) return jsonOut({ error: 'Buchung nicht gefunden oder bereits storniert.' });
      return jsonOut({ ok: true, id: saleId, cancelled: true });
    }

    if (body.type === 'product') {
      const sh = ss.getSheetByName('Products');
      const id = Utilities.getUuid();
      sh.appendRow([id, body.category, body.name, body.price, body.emoji || '']);
      return jsonOut({ ok: true, id: id });
    }

    if (body.type === 'delete_product') {
      const sh = ss.getSheetByName('Products');
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === body.id) {
          sh.deleteRow(i + 1);
          break;
        }
      }
      return jsonOut({ ok: true });
    }

    return jsonOut({ error: 'unbekannter type' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function updateSummary(sh, productName, qty, revenue) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === productName) {
      sh.getRange(i + 1, 2).setValue(data[i][1] + qty);
      sh.getRange(i + 1, 3).setValue(data[i][2] + revenue);
      return;
    }
  }
  sh.appendRow([productName, qty, revenue]);
}

function updatePayments(sh, method, revenue) {
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === method) {
      sh.getRange(i + 1, 2).setValue(data[i][1] + revenue);
      return;
    }
  }
  sh.appendRow([method, revenue]);
}

// Einmalig manuell im Editor ausführen (Run-Dropdown) ODER über das Menü
// "Kasse > Alle Buchungen zurücksetzen (neues Event)" im Sheet selbst.
// Löscht ALLE Zeilen in Sales, Summary und Payments (Kopfzeilen bleiben).
// Products (die Getränkekarte inkl. selbst hinzugefügter Getränke) bleibt
// unangetastet — nur die Verkaufsdaten/Strichliste werden zurückgesetzt.
function resetAllBookings() {
  ['Sales', 'Summary', 'Payments'].forEach(name => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
    }
  });
  SpreadsheetApp.getUi().alert('Alle Buchungen wurden zurückgesetzt. Die Getränkekarte (Products) bleibt erhalten.');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
