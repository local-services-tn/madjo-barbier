/**
 * MADJO — Barbier · Réservations vers Google Sheets
 * =====================================================
 * Ce script reçoit les réservations du site (index.html) et
 * les enregistre dans votre Google Sheet. Il répond aussi aux
 * demandes de disponibilité : le site affiche en vert les
 * créneaux libres et en grisé ceux qui sont déjà pris.
 *
 * Au premier appel, il crée tout seul la feuille « Réservations »
 * et la ligne d'en-têtes : rien à préparer dans la feuille.
 *
 * -----------------------------------------------------
 * MISE EN PLACE
 * -----------------------------------------------------
 * 1. Ouvrez votre Google Sheet : Extensions → Apps Script →
 *    collez ce fichier → Enregistrer (Ctrl+S).
 * 2. Déployer → Nouveau déploiement → Application Web :
 *      - Exécuter en tant que « Moi »
 *      - Qui peut accéder « Toute personne »
 *    → Déployer → copiez l'URL « Web App » (…/exec).
 * 3. Collez cette URL dans index.html :
 *      const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/.../exec';
 *
 * Tableau de bord : ?action=bookings&key=madjo2026 renvoie toutes les
 * réservations (protégé par DASH_KEY). dashboard.html l'utilise.
 *
 * Clients : ?action=lookup&phone=20123456 renvoie les réservations d'un
 * numéro ; POST {"action":"cancel","ref":"MJO-XXX","phone":"20123456"}
 * annule une réservation (colonne « Statut » → Annulé, créneau libéré).
 * =====================================================
 */

const SHEET_ID = '1ZffXXIHJn1jYP3x6niG5ow9rz0lRMLXdeyqfSdCJ1lw';
const SHEET_NAME = 'Réservations';
const HEADERS = ['Référence', 'Date création', 'Service', 'Durée (min)', 'Prix',
  'Barbier', 'Rôle', 'Date RDV', 'Heure', 'Nom', 'Téléphone', 'Email', 'Notes', 'Statut'];

// Mot de passe du tableau de bord (dashboard.html). Protège les données
// clients. Changez-le à votre goût, puis mettez à jour dashboard.html
// avec la même valeur.
const DASH_KEY = 'madjo2026';

/** Récupère (ou crée) la feuille + en-têtes, force Date RDV/Heure en texte. */
function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
  } else {
    // Met à jour la ligne d'en-têtes si une colonne a été ajoutée (« Statut »).
    var firstRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (firstRow.length < HEADERS.length) sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  // Colonnes H (Date RDV) et I (Heure) en TEXTE : la date/heure
  // reste identique (« 2026-08-20 », « 08:00 ») sans conversion.
  sh.getRange('H2:I').setNumberFormat('@');
  return sh;
}

/**
 * Recevoir une réservation (POST).
 * Le site envoie un JSON : {ref, barber, service, durationMin,
 * price, date, time, name, phone, email, notes}
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'cancel') return cancelBooking_(data);

    var sh = getSheet_();

    var roles = {
      'Madjo':   'Maître barbier',
      'Yassine': 'Expert coupes',
      'Mehdi':   'Spécialiste barbe'
    };

    // Refus d'un double créneau : même barbier, même jour, même heure déjà pris.
    var existing = sh.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (!existing[i][0]) continue;
      if (String(existing[i][5]).trim() === String(data.barber).trim()
          && String(existing[i][7]) === String(data.date)
          && String(existing[i][8]) === String(data.time)) {
        return json_({ ok: false, error: 'slot_taken', ref: data.ref });
      }
    }

    sh.appendRow([
      data.ref,
      new Date(),
      data.service,
      data.durationMin,
      data.price,
      data.barber,
      roles[data.barber] || '',
      String(data.date),   // « YYYY-MM-DD », gardé en texte
      String(data.time),   // « HH:MM », gardé en texte
      data.name,
      data.phone,
      data.email || '',
      data.notes || '',
      'Confirmé'
    ]);

    return json_({ ok: true, ref: data.ref });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * Disponibilité (GET).
 * Appelé par le site : ?action=avail&barber=Madjo
 * Renvoie les créneaux déjà pris pour chaque jour :
 *   { ok:true, barber:"Madjo",
 *     bookings: { "2026-08-16": ["08:00","09:30"], ... } }
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'ping';

    if (action === 'avail') {
      var barber = (e.parameter.barber || '').trim();
      var sh = getSheet_();
      var values = sh.getDataRange().getValues();

      var bookings = {};
      for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (!row[0]) continue;
        if (String(row[13]).trim() === 'Annulé') continue; // créneau libéré
        var rowBarber = String(row[5]).trim();
        var dateKey = toDateKey_(row[7]);
        var timeStr = toTimeStr_(row[8]);
        if (!dateKey || !timeStr) continue;
        if (barber === '' || rowBarber === barber) {
          if (!bookings[dateKey]) bookings[dateKey] = [];
          if (bookings[dateKey].indexOf(timeStr) === -1) bookings[dateKey].push(timeStr);
        }
      }
      return json_({ ok: true, barber: barber, bookings: bookings });
    }

    if (action === 'lookup') {
      var phone = phoneKey_(e.parameter.phone || '');
      if (!phone) return json_({ ok: false, error: 'phone_required' });
      var shL = getSheet_();
      var valsL = shL.getDataRange().getValues();
      var found = [];
      for (var l = 1; l < valsL.length; l++) {
        var rL = valsL[l];
        if (!rL[0]) continue;
        if (phoneKey_(rL[10]) !== phone) continue;
        found.push({
          ref: String(rL[0] || ''),
          service: String(rL[2] || ''),
          price: rL[4],
          barber: String(rL[5] || ''),
          date: toDateKey_(rL[7]),
          time: toTimeStr_(rL[8]),
          name: String(rL[9] || ''),
          status: String(rL[13]).trim() === 'Annulé' ? 'cancelled' : 'confirmed'
        });
      }
      found.sort(function (a, b) {
        var ka = a.date + ' ' + a.time, kb = b.date + ' ' + b.time;
        return ka < kb ? -1 : ka > kb ? 1 : 0;
      });
      return json_({ ok: true, bookings: found });
    }

    if (action === 'bookings') {
      if ((e.parameter.key || '') !== DASH_KEY) {
        return json_({ ok: false, error: 'unauthorized' });
      }
      var sh2 = getSheet_();
      var values2 = sh2.getDataRange().getValues();
      var rows = [];
      for (var j = 1; j < values2.length; j++) {
        var r = values2[j];
        if (!r[0]) continue; // ignore une ligne vide
        rows.push({
          ref: String(r[0] || ''),
          createdAt: r[1] instanceof Date ? Utilities.formatDate(r[1], tz_(), 'yyyy-MM-dd HH:mm') : String(r[1] || ''),
          service: String(r[2] || ''),
          durationMin: String(r[3] || ''),
          price: r[4],
          barber: String(r[5] || ''),
          role: String(r[6] || ''),
          date: toDateKey_(r[7]),
          time: toTimeStr_(r[8]),
          name: String(r[9] || ''),
          phone: String(r[10] || ''),
          email: String(r[11] || ''),
          notes: String(r[12] || ''),
          status: String(r[13]).trim() === 'Annulé' ? 'cancelled' : 'confirmed'
        });
      }
      return json_({ ok: true, headers: HEADERS, rows: rows });
    }

    return json_({ ok: true, message: 'Madjo — endpoint actif.' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ---------- aide ---------- */

/** Annule une réservation : marque « Annulé » (colonne N), garde la trace. */
function cancelBooking_(data) {
  var ref = String(data.ref || '').trim();
  var phone = phoneKey_(data.phone || '');
  if (!ref || !phone) return json_({ ok: false, error: 'bad_request' });

  var sh = getSheet_();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    if (String(row[0]).trim() === ref && phoneKey_(row[10]) === phone) {
      if (String(row[13]).trim() === 'Annulé') {
        return json_({ ok: false, error: 'already_cancelled' });
      }
      sh.getRange(i + 1, 14).setValue('Annulé');
      return json_({ ok: true, cancelled: ref });
    }
  }
  return json_({ ok: false, error: 'not_found' });
}

/** Normalise un téléphone : chiffres seuls, sans l'indicatif 216. */
function phoneKey_(v) {
  var s = String(v || '').replace(/[^\d]/g, '');
  if (s.length >= 10 && s.slice(0, 3) === '216') s = s.slice(3);
  return s;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** « 2026-08-16 » ou « 16/08/2026 » → « 2026-08-16 ». Renvoie '' si invalide. */
function toDateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var dd = m[1].length === 1 ? '0' + m[1] : m[1];
    var mm = m[2].length === 1 ? '0' + m[2] : m[2];
    return m[3] + '-' + mm + '-' + dd;
  }
  return '';
}

/** « 08:00 » → « 08:00 ». Accepte aussi un objet Date (saisie manuelle). */
function toTimeStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  var s = String(v).trim();
  var m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return (m[1].length === 1 ? '0' + m[1] : m[1]) + ':' + m[2];
  return s;
}

function tz_() {
  try { return Session.getScriptTimeZone(); }
  catch (e) { return 'Africa/Tunis'; }
}
