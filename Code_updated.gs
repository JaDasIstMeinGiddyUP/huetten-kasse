/**
 * KASSEN-BACKEND FÜR GOOGLE SHEETS
 * ---------------------------------
 * Diese Datei kommt komplett in den Apps-Script-Editor deines Google Sheets
 * (Erweiterungen → Apps Script) — ersetzt den kompletten bisherigen Inhalt.
 *
 * Erwartet Tabellenblätter im Sheet:
 *   "Sales"    Kopfzeile: sale_id | timestamp | zahlungsart | huette | artikel | menge | einzelpreis | zeilensumme | storniert | gebucht_von | storniert_von | storno_grund
 *              -> jede Buchung wird zeilenweise pro Artikel eingetragen, gut lesbar, keine JSON-Klumpen
 *              -> Spalte 9 "storniert" (TRUE/leer) markiert stornierte Buchungen
 *              -> Spalte 10 "gebucht_von" = Gerätename/Person, die gebucht hat
 *              -> Spalte 11 "storniert_von" = Gerätename/Person, die storniert hat
 *              -> Spalte 12 "storno_grund" = vom Personal eingegebene Begründung der Stornierung
 *   "Summary"  (nicht mehr live beschrieben, siehe PERFORMANCE-FIX unten) Kopfzeile: artikel | menge | umsatz
 *   "Payments" (nicht mehr live beschrieben, siehe PERFORMANCE-FIX unten) Kopfzeile: zahlungsart | umsatz
 *   "Products" Kopfzeile: id | category | name | price | emoji
 *
 * PERFORMANCE-FIX (2026-10-03): Bisher wurde bei JEDER Buchung/Stornierung
 * zusätzlich das komplette "Summary"- und "Payments"-Blatt eingelesen und
 * zeilenweise durchsucht, um die Summen live nachzuführen. Das machte jede
 * Buchung umso langsamer, je mehr Buchungen an dem Abend schon im Sheet
 * standen (merklich spürbar bei stark besuchten Events mit vielen Geräten
 * gleichzeitig, die sich zusätzlich per LockService hintereinanderreihen).
 * Jetzt wird bei einer Buchung NUR NOCH die Sales-Zeile angehängt (schnell,
 * O(1)) — Summary/Payments werden stattdessen bei Bedarf (action=summary /
 * action=payments) direkt aus "Sales" live berechnet. Für die Kassierenden
 * ändert sich nichts, nur das ständige "Synchronisiere..." wird spürbar
 * kürzer bzw. bleibt über den ganzen Abend hinweg gleich schnell statt
 * immer länger zu dauern.
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
    return jsonOut(computeAggregates_(ss).summary);
  }

  if (action === 'payments') {
    return jsonOut(computeAggregates_(ss).payments);
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

// Ein einziger Durchlauf über "Sales" liefert sowohl die Artikel-Summary
// als auch die Zahlungsart-Summe (storno-bereinigt). Wird nur bei Bedarf
// (Übersicht öffnen) aufgerufen, nicht bei jeder Buchung.
function computeAggregates_(ss) {
  const sh = ss.getSheetByName('Sales');
  const data = sh.getDataRange().getValues();
  const byProduct = {}; // name -> {qty, revenue}
  const byPayment = {}; // method -> revenue
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[8]) continue; // storniert -> zählt nicht mit
    const payment = row[2];
    const productName = row[4];
    const qty = Number(row[5]) || 0;
    const lineTotal = Number(row[7]) || 0;

    if (!byProduct[productName]) byProduct[productName] = { qty: 0, revenue: 0 };
    byProduct[productName].qty += qty;
    byProduct[productName].revenue += lineTotal;

    byPayment[payment] = (byPayment[payment] || 0) + lineTotal;
  }
  return {
    summary: Object.keys(byProduct).map(name => ({ name, qty: byProduct[name].qty, revenue: byProduct[name].revenue })),
    payments: Object.keys(byPayment).map(method => ({ method, revenue: byPayment[method] }))
  };
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
        cancelledBy: row[10] || '',
        cancelReason: row[11] || ''
      };
      order.push(saleId);
    }
    bySale[saleId].total += Number(row[7]) || 0;
    if (row[8]) bySale[saleId].cancelled = true;
    if (row[10]) bySale[saleId].cancelledBy = row[10];
    if (row[11]) bySale[saleId].cancelReason = row[11];
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
        salesSh.appendRow([saleId, ts, payment, body.hut || '', item.name, item.qty, item.price, lineTotal, '', bookedBy, '', '']);
      });

      return jsonOut({ ok: true, id: saleId });
    }

    if (body.type === 'cancel_sale') {
      const salesSh = ss.getSheetByName('Sales');
      const saleId = body.clientId;
      const cancelledBy = body.device || '';
      const reason = (body.reason || '').toString().trim();
      if (!reason) {
        return jsonOut({ error: 'Begründung fehlt – Stornierung ohne Grund ist nicht erlaubt.' });
      }
      const data = salesSh.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[0] !== saleId) continue;
        if (row[8]) continue; // schon storniert, nichts doppelt zurückziehen
        found = true;
        salesSh.getRange(i + 1, 9).setValue(true);        // Spalte 9 = storniert
        salesSh.getRange(i + 1, 11).setValue(cancelledBy); // Spalte 11 = storniert_von
        salesSh.getRange(i + 1, 12).setValue(reason);      // Spalte 12 = storno_grund
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

// Einmalig manuell im Editor ausführen (Run-Dropdown) ODER über das Menü
// "Kasse > Alle Buchungen zurücksetzen (neues Event)" im Sheet selbst.
// Löscht ALLE Zeilen in Sales (Summary/Payments werden nicht mehr befüllt,
// bleiben zur Sicherheit aber im Reset enthalten falls noch alte Daten drin
// stehen). Products (die Getränkekarte) bleibt unangetastet.
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

// Einmalig manuell im Editor ausführen (Run-Dropdown -> diese Funktion
// auswählen -> Run), wenn man die Tabellenblätter "Summary" und "Payments"
// im Sheet selbst mal auf den aktuellen Stand bringen will (z.B. zum
// Draufschauen/Exportieren). Für den laufenden Kassenbetrieb NICHT nötig,
// die App liest ihre Zahlen live direkt aus "Sales" (siehe doGet oben).
// Komplett getrennt von doPost/doGet, beeinflusst laufenden Betrieb nicht.
function refreshSummaryPaymentsSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const agg = computeAggregates_(ss);

  const summarySh = ss.getSheetByName('Summary');
  if (summarySh) {
    const lastRow = summarySh.getLastRow();
    if (lastRow > 1) summarySh.getRange(2, 1, lastRow - 1, summarySh.getLastColumn()).clearContent();
    const rows = agg.summary.map(r => [r.name, r.qty, r.revenue]);
    if (rows.length) summarySh.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  const paymentsSh = ss.getSheetByName('Payments');
  if (paymentsSh) {
    const lastRow = paymentsSh.getLastRow();
    if (lastRow > 1) paymentsSh.getRange(2, 1, lastRow - 1, paymentsSh.getLastColumn()).clearContent();
    const rows = agg.payments.map(r => [r.method, r.revenue]);
    if (rows.length) paymentsSh.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  Logger.log('Summary & Payments wurden auf den aktuellen Stand gebracht.');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
