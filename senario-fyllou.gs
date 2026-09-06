/**
 * Σενάριο για την εφαρμογή «Πρόσκληση & Οργάνωση Εκδήλωσης».
 * Μπαίνει σε φύλλο Google, μέσα από Επεκτάσεις → Apps Script.
 * Δημιουργεί μόνο του δύο καρτέλες: «ekdilosi» και «apantiseis».
 */

var FYLLO_EKD = 'ekdilosi';
var FYLLO_APA = 'apantiseis';

var STILES = ['id','onoma','apantisi','enilikes','paidia','vrefi','onomata','pota','fagita',
              'elef_pota','elef_fagita','diatrofi','alergies','ferno','doro','posoDorou',
              'metakinisi','theseis','simeiosi','pote'];

function vresFyllo_(onoma) {
  var vivlio = SpreadsheetApp.getActiveSpreadsheet();
  var f = vivlio.getSheetByName(onoma);
  if (!f) {
    f = vivlio.insertSheet(onoma);
    if (onoma === FYLLO_APA) f.appendRow(STILES);
  }
  return f;
}

function apantisi_(antikeimeno) {
  return ContentService
    .createTextOutput(JSON.stringify(antikeimeno))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Ανάγνωση ---------- */
function doGet(e) {
  try {
    var fEkd = vresFyllo_(FYLLO_EKD);
    var kelí = fEkd.getRange('A1').getValue();
    var vasi = kelí ? JSON.parse(kelí) : {};

    var fApa = vresFyllo_(FYLLO_APA);
    var dedomena = fApa.getDataRange().getValues();
    var apantiseis = [];

    for (var i = 1; i < dedomena.length; i++) {
      var grammi = dedomena[i];
      if (!grammi[0]) continue;
      var a = {};
      for (var j = 0; j < STILES.length; j++) a[STILES[j]] = grammi[j];
      a.enilikes = Number(a.enilikes) || 0;
      a.paidia   = Number(a.paidia)   || 0;
      a.vrefi    = Number(a.vrefi)    || 0;
      a.theseis  = Number(a.theseis)  || 0;
      a.pote     = Number(a.pote)     || 0;
      ['onomata','pota','fagita','diatrofi','ferno'].forEach(function (k) {
        try { a[k] = a[k] ? JSON.parse(a[k]) : []; } catch (err) { a[k] = []; }
      });
      apantiseis.push(a);
    }

    return apantisi_({
      ekdilosi:  vasi.ekdilosi  || null,
      pota:      vasi.pota      || [],
      fagita:    vasi.fagita    || [],
      istoriko:  vasi.istoriko  || [],
      ekdId:     vasi.ekdId     || '',
      apantiseis: apantiseis
    });
  } catch (err) {
    return apantisi_({ sfalma: String(err) });
  }
}

/* ---------- Εγγραφή ---------- */
function doPost(e) {
  var kleidoma = LockService.getScriptLock();
  kleidoma.waitLock(20000);
  try {
    var aitima = JSON.parse(e.postData.contents);

    if (aitima.energeia === 'ekdilosi') {
      vresFyllo_(FYLLO_EKD).getRange('A1').setValue(JSON.stringify(aitima.dedomena));
      return apantisi_({ ok: true });
    }

    if (aitima.energeia === 'apantisi') {
      var a = aitima.dedomena;
      var f = vresFyllo_(FYLLO_APA);
      var grammi = STILES.map(function (k) {
        var t = a[k];
        if (t === undefined || t === null) return '';
        return (typeof t === 'object') ? JSON.stringify(t) : t;
      });

      var ids = f.getRange(1, 1, Math.max(f.getLastRow(), 1), 1).getValues();
      var thesi = -1;
      for (var i = 1; i < ids.length; i++) if (ids[i][0] === a.id) { thesi = i + 1; break; }

      if (thesi > 0) f.getRange(thesi, 1, 1, STILES.length).setValues([grammi]);
      else f.appendRow(grammi);

      return apantisi_({ ok: true });
    }

    if (aitima.energeia === 'katharismos') {
      var fa = vresFyllo_(FYLLO_APA);
      if (fa.getLastRow() > 1) fa.deleteRows(2, fa.getLastRow() - 1);
      return apantisi_({ ok: true });
    }

    return apantisi_({ sfalma: 'άγνωστη ενέργεια' });
  } catch (err) {
    return apantisi_({ sfalma: String(err) });
  } finally {
    kleidoma.releaseLock();
  }
}
