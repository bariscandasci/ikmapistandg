/**
 * Tepe Güvenlik 4. Direktörlük — paylaşımlı proje veritabanı + aday yorumları + AI.
 * Ankara ve İstanbul Anadolu sayfalarından AYRI bir Google Sheet'e yazar.
 *
 * Kurulum:
 * 1. Yeni bir Google Sheet aç (ör. "Tepe Güvenlik 4. Direktörlük").
 * 2. Uzantılar > Apps Script, bu dosyanın tamamını yapıştır, kaydet.
 * 3. Dağıt > Yeni dağıtım > Web uygulaması.
 *    Yürütme: Ben, Erişim: Herkes. Dağıt.
 * 4. Çıkan /exec adresini app.js içindeki SHEET_API_URL'e yapıştır.
 * İlk sayfa açılışı boş "Projeler" sekmesine 10 güvenlik projesini yazar
 * (urgent=false, active=true). Sonraki aç/kapa ve acil kayıtları aynı satırlara
 * yazılır. AI için Script Özellikleri'ne GEMINI_API_KEY eklenebilir.
 * Kod güncellemesinde Dağıtımları yönet > mevcut dağıtımın yeni sürümünü yayınla
 * (adres değişmez).
 */

const SHEET_NAME = "Projeler";
const FIELDS = [
  "id", "name", "sector", "address", "lat", "lng", "accessStopId",
  "shift", "salary", "meal", "transport", "referral", "gender",
  "urgent", "active", "capacity", "position", "ageMin", "ageMax",
];

/**
 * "capacity", "position", "ageMin" ve "ageMax" sonradan eklenen alanlar — canlı Sheet'in başlık
 * satırında henüz yoksa (ör. bu kod ilk kez deploy edildiğinde) burada
 * otomatik olarak sütun başlıkları eklenir. Böylece kullanıcı Sheet'i elle
 * düzenlemek zorunda kalmaz. Mevcut satırlardaki hücreler boş kalır — bu
 * satırlar "Tümü" dışındaki bir Pozisyon Filtresi'nde görünmez, İK Sheet'ten
 * elle doldurabilir.
 */
function ensureLaterAddedColumns_() {
  const sheet = getSheet_();
  let lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  ["capacity", "position", "ageMin", "ageMax"].forEach((col) => {
    if (!header.includes(col)) {
      lastCol += 1;
      sheet.getRange(1, lastCol).setValue(col);
      header.push(col);
    }
  });
}

// Aday yorumları — "Comments" sekmesi. Bu sekmenin kendi id sütunu yok; bir
// satır candidateId+candidateName+author+text+createdAt ile tanımlanır.
const COMMENTS_SHEET_NAME = "Comments";
const COMMENT_FIELDS = ["candidateId", "candidateName", "author", "text", "createdAt"];

/** "Projeler" sekmesi yoksa (ör. ilk kurulum) başlık satırıyla birlikte oluşturur. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(FIELDS);
  }
  return sheet;
}

/** "Comments" sekmesi yoksa (ör. ilk kurulum) başlık satırıyla birlikte oluşturur. */
function getCommentsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(COMMENTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(COMMENTS_SHEET_NAME);
    sheet.appendRow(COMMENT_FIELDS);
  }
  return sheet;
}

function readAllComments_() {
  const sheet = getCommentsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0].map((h) => String(h).trim());
  return values.slice(1)
    .filter((row) => row[header.indexOf("candidateId")])
    .map((row) => {
      const obj = {};
      header.forEach((key, i) => {
        obj[key] = row[i] == null ? "" : String(row[i]);
      });
      return obj;
    });
}

function readAllRows_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const header = values[0].map((h) => String(h).trim());
  return values.slice(1)
    .filter((row) => row[header.indexOf("id")]) // boş satırları atla
    .map((row) => {
      const obj = {};
      header.forEach((key, i) => {
        let v = row[i];
        if (key === "urgent" || key === "active") {
          v = v === true || v === "TRUE" || v === "true";
        } else if (key === "lat" || key === "lng") {
          v = Number(v);
        } else if (key === "capacity" || key === "ageMin" || key === "ageMax") {
          // Boş hücre = tanımsız (kontenjan için sınırsız); dolu hücre sayıya çevrilir.
          v = v === "" || v == null ? null : Number(v);
        } else {
          v = v == null ? "" : String(v);
        }
        obj[key] = v;
      });
      return obj;
    });
}

function findRowIndexById_(id) {
  const sheet = getSheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // 1-index + header satırı
  }
  return -1;
}

/**
 * "Projeler" sekmesi boşsa Tepe Güvenlik 4. Direktörlük'teki 10 projeyi yazar.
 * Hepsi açık (active=true) ve acil değil (urgent=false) başlar. Sekmede veri
 * varsa hiçbir satıra dokunmaz — aç/kapa ve acil işaretleri korunur.
 * doGet ilk okumada bunu kendisi çağırır; editörden de çalıştırılabilir.
 */
function seedIfEmpty() {
  const sheet = getSheet_();
  if (sheet.getLastRow() > 1) return;
  const rows = [
    ["proj_kentplus_3352", "Kentplus 3352 Ada Sitesi", "Güvenlik", "Barbaros Mahallesi, Ataşehir", 40.997671, 29.104585, "stop_yenisahra", "12 saat · 2 gündüz 2 gece 2 izin", "47.000₺ · net 51.000₺", "Projede", "3.628₺", "Projeye yönlendirilecek — Operasyon: Ferdi Hakan / İK: Feray Çorlu — Elit site; özel güvenlik kimlik kartı zorunlu", "Erkek", false, true, "", "Özel Güvenlik", 22, 43],
    ["proj_pendorya_avm", "Pendorya AVM", "Güvenlik", "Çamçeşme, Pendik", 40.866101, 29.275137, "stop_tersane", "12 saat · 2 gündüz 2 gece 2 izin", "44.500₺ · net 47.500₺", "Projede", "3.628₺", "Projeye yönlendirilecek — Operasyon: Selim Öztürk / İK: Feray Çorlu — Mesai imkânı mevcut", "Erkek", false, true, "", "Özel Güvenlik", 22, 44],
    ["proj_tepe_nautilus", "Tepe Nautilus AVM", "Güvenlik", "Acıbadem, Fatih Sk. No:1, Kadıköy", 40.999617, 29.031495, "stop_ayrilik_cesmesi", "07:00-15:00 / 15:00-23:00 / 23:00-07:00 · 6 çalış 1 izin", "41.000₺ · net 44.628₺", "Projede", "3.628₺", "Projeye yönlendirilecek — Operasyon: Berkok Çelikcan / İK: Ezgi Nur Taşkın — Vergi kesintisi yok", "Kadın/Erkek", false, true, "", "Özel Güvenlik", 22, 45],
    ["proj_anadolu_grubu", "Anadolu Grubu", "Güvenlik", "Buyaka, Balkan Cd. No:58, Ümraniye", 41.026465, 29.127646, "stop_ihlamurkuyu", "8 saat · 6 çalış 1 izin", "54.000₺ · net 62.618₺", "4.990₺", "3.628₺", "Projeye yönlendirilecek — Operasyon: Nazım Güler / İK: Feray Çorlu", "Erkek", false, true, "", "Özel Güvenlik", 22, 32],
    ["proj_acarer_metal", "Acarer Metal (BOSB)", "Güvenlik", "Aydınlı, 1 Nolu Cd. No:14, BOSB, Tuzla", 40.877745, 29.397198, "stop_sabiha_gokcen_havalimani", "12 saat · 3 çalış 2 off", "41.750₺ · net 47.550₺", "Projede", "5.250₺", "Projeye yönlendirilecek — Operasyon: Berkok Çelikcan / İK: Ezgi Nur Taşkın — Metal fabrikası", "Erkek", false, true, "", "Özel Güvenlik", 22, 46],
    ["proj_sukru_saracoglu", "Şükrü Saracoğlu", "Güvenlik", "Zühtüpaşa, Recep Peker Cd. No:3, Kadıköy", 40.987581, 29.036977, "stop_sogutlucesme", "12 saat · 2 gündüz 2 gece 2 izin", "39.600₺ · net 53.628₺", "3.628₺", "10.400₺", "Projeye yönlendirilecek — Operasyon: Berkok Çelikcan / İK: Ezgi Nur Taşkın — Stadyum projesi", "Erkek", false, true, "", "Özel Güvenlik", 22, 44],
    ["proj_kuhne_nagel", "Kühne Nagel Depo", "Güvenlik", "Orta, Öztaş Sk. No:12, Logipark, Tuzla", 40.891272, 29.396216, "stop_sabiha_gokcen_havalimani", "12 saat · 4 çalış 3 off", "42.300₺", "Projede", "Servis", "Projeye yönlendirilecek — Operasyon: Berkok Çelikcan / İK: Ezgi Nur Taşkın", "Erkek", false, true, "", "Özel Güvenlik", 22, 47],
    ["proj_sutas", "Sütaş", "Güvenlik", "Dudullu OSB, Ümraniye", 40.996189, 29.177048, "stop_modoko", "12 saat · 2 gündüz 2 gece 2 izin", "32.720₺ · net 43.632₺", "Projede", "Servis", "Operasyon görüşmesi, ardından projeye yönlendirilecek — Operasyon: Murat Doğan / İK: S. Nur Güleç — Tecilli ve emekli alımı yok", "Erkek", false, true, "", "Özel Güvenlik", 22, 40],
    ["proj_medipol_kavacik", "Medipol Üniversitesi Kampüsü", "Güvenlik", "Kavacık, Beykoz", 41.091038, 29.091431, "stop_carsi_2", "12 saat · 2 gündüz 2 gece 2 izin", "41.563₺", "Projede", "3.628₺", "Projeye yönlendirilecek — Operasyon: Murat Doğan / İK: S. Nur Güleç — Tecilli alımı yok; emekli alımı var", "Erkek", false, true, "", "Özel Güvenlik", 22, 45],
    ["proj_bomontiada", "Bomontiada", "Güvenlik", "Bomonti, Şişli", 41.058118, 28.980848, "stop_osmanbey", "07:00-15:00 / 15:00-23:00 / 23:00-07:00 · 6 çalış 1 izin", "38.432₺ · net 41.730₺", "7.800₺", "3.628₺", "Projeye yönlendirilecek — Operasyon: Ersan Artan / İK: S. Nur Güleç — X-ray bilgisi gerekli", "Kadın", false, true, "", "Özel Güvenlik", 22, 39],
  ];
  sheet.getRange(2, 1, rows.length, FIELDS.length).setValues(rows);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureLaterAddedColumns_();
    seedIfEmpty();
    const projects = readAllRows_();
    const comments = readAllComments_();
    return ContentService.createTextOutput(JSON.stringify({ ok: true, projects, comments }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonError_("Geçersiz istek gövdesi");
  }

  const action = body.action;

  // AI istekleri Sheet'e dokunmaz — sheet kilidinin dışında, hemen işlenir
  // (aksi halde bir Gemini çağrısı sürerken diğer kullanıcıların proje
  // kaydetme/güncelleme istekleri gereksiz yere bekler).
  if (action === "aiComplete") {
    return callGemini_(body.system, body.prompt, body.maxTokens);
  }

  const sheet = getSheet_();

  // Aynı anda birden fazla yazma isteği gelirse (ör. birkaç kişi aynı anda
  // kaydet'e basarsa) satır numaraları kayabilir ve yanlış satırlar
  // güncellenebilir — bu kilit, bir istek bitmeden diğerinin başlamasını
  // engelleyerek bunu önler.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (action === "add") {
      ensureLaterAddedColumns_();
      const p = body.project || {};
      if (!p.name) return jsonError_("name zorunlu");
      const id = p.id || ("proj_" + p.name.toLowerCase()
        .replace(/[ığüşöçİĞÜŞÖÇ]/g, (c) => ({ "ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c", "İ": "i", "Ğ": "g", "Ü": "u", "Ş": "s", "Ö": "o", "Ç": "c" }[c] || c))
        .replace(/[^a-z0-9]+/g, "_")
        .replace(/^_+|_+$/g, "") + "_" + Date.now().toString(36));
      const row = FIELDS.map((f) => {
        if (f === "id") return id;
        if (f === "urgent" || f === "active") return p[f] ? true : (f === "active" ? true : false);
        return p[f] != null ? p[f] : "";
      });
      sheet.appendRow(row);
      return jsonOk_({ id });
    }

    if (action === "update") {
      const id = body.id;
      const patch = body.patch || {};
      if (!id) return jsonError_("id zorunlu");
      const rowIndex = findRowIndexById_(id);
      if (rowIndex === -1) return jsonError_("proje bulunamadı: " + id);
      const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      Object.keys(patch).forEach((key) => {
        const col = header.indexOf(key);
        if (col !== -1) sheet.getRange(rowIndex, col + 1).setValue(patch[key]);
      });
      return jsonOk_({ id });
    }

    // Birden fazla projeyi TEK bir istekte günceller — ama SADECE gerçekten
    // değişen hücrelere yazar. (ÖNEMLİ: önceki sürüm bütün satırı okuyup
    // range.setValues() ile TÜMÜNÜ geri yazıyordu; bu, dokunulmayan lat/lng
    // hücrelerinin bile Sheets tarafından yeniden yorumlanıp ondalık
    // noktalarının silinmesine — ör. 39.87 -> 3987 — yol açtı ve TÜM
    // projelerin koordinatını bozdu. Şimdi sadece id sütunu okunuyor, her
    // patch yalnızca kendi hücresine yazılıyor; başka hiçbir hücreye
    // dokunulmuyor.)
    if (action === "updateBatch") {
      ensureLaterAddedColumns_();
      const patches = body.patches || []; // [{ id, patch }, ...]
      const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      const idCol = header.indexOf("id");
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return jsonOk_({ updated: 0, notFound: patches.map((p) => p.id) });
      const idValues = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
      const idToRow = {};
      idValues.forEach((r, i) => { idToRow[r[0]] = i + 2; }); // 1-index + başlık satırı

      const notFound = [];
      let updated = 0;
      patches.forEach(({ id, patch }) => {
        const row = idToRow[id];
        if (row === undefined) { notFound.push(id); return; }
        Object.keys(patch || {}).forEach((key) => {
          const col = header.indexOf(key);
          if (col !== -1) sheet.getRange(row, col + 1).setValue(patch[key]);
        });
        updated++;
      });
      return jsonOk_({ updated, notFound });
    }

    if (action === "remove") {
      const id = body.id;
      if (!id) return jsonError_("id zorunlu");
      const rowIndex = findRowIndexById_(id);
      if (rowIndex === -1) return jsonError_("proje bulunamadı: " + id);
      sheet.deleteRow(rowIndex);
      return jsonOk_({ id });
    }

    if (action === "addComment") {
      const c = body.comment || {};
      if (!c.candidateId || !c.author || !c.text) {
        return jsonError_("candidateId, author ve text zorunlu");
      }
      const record = {
        candidateId: c.candidateId,
        candidateName: c.candidateName || "",
        author: c.author,
        text: c.text,
        createdAt: new Date().toISOString(),
      };
      getCommentsSheet_().appendRow(COMMENT_FIELDS.map((f) => record[f]));
      return jsonOk_({ comment: record });
    }

    // Bir yorumu candidateId+createdAt eşleşmesiyle bulup satırı siler (ör.
    // hatalı/test amaçlı girilmiş bir yorumu temizlemek için). createdAt her
    // yorumda benzersiz olduğundan (appendRow anında üretilir) bu ikili tek
    // bir satırı işaret etmeye yeter.
    if (action === "removeComment") {
      const target = body.comment || {};
      if (!target.candidateId || !target.createdAt) return jsonError_("candidateId ve createdAt zorunlu");
      const commentsSheet = getCommentsSheet_();
      const values = commentsSheet.getDataRange().getValues();
      const header = values[0].map(String);
      const candCol = header.indexOf("candidateId");
      const createdCol = header.indexOf("createdAt");
      for (let r = values.length - 1; r >= 1; r--) {
        if (String(values[r][candCol]) === target.candidateId && String(values[r][createdCol]) === target.createdAt) {
          commentsSheet.deleteRow(r + 1);
          return jsonOk_({ removed: true });
        }
      }
      return jsonError_("yorum bulunamadı");
    }

    return jsonError_("bilinmeyen action: " + action);
  } catch (err) {
    return jsonError_(String(err));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Google Gemini API'ye tek seferlik bir tamamlama isteği gönderir — İK
 * ekibine sunulan AI özellikleri (eşleşme gerekçesi, doğal dil filtre,
 * sütun eşleştirme önerisi) hepsi bu tek genel amaçlı uç noktayı kullanır.
 * API anahtarı bilerek istemci tarafında (app.js) DEĞİL, burada, Apps
 * Script "Script Özellikleri"nde tutulur — statik bir sitede anahtarı
 * tarayıcıya koymak herkese açık hale getirir.
 *
 * Kurulum: https://aistudio.google.com/apikey üzerinden ücretsiz bir
 * Gemini API anahtarı alın, sonra Apps Script editöründe Proje Ayarları >
 * Script Özellikleri'ne GEMINI_API_KEY adında bir özellik ekleyip bu
 * anahtarı değer olarak girin.
 */
function callGemini_(system, prompt, maxTokens) {
  if (!prompt) return jsonError_("prompt zorunlu");
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    return jsonError_("GEMINI_API_KEY tanımlı değil — Apps Script > Proje Ayarları > Script Özellikleri'nden ekleyin (aistudio.google.com/apikey'den ücretsiz alınır).");
  }
  // Model, görünür metinden ÖNCE değişken miktarda "düşünme" tokenı da
  // harcıyor (thinkingConfig bu modelde desteklenmiyor — "invalid argument"
  // hatası verdi; sabit bir alt sınır tahmini de yetmedi, görev karmaşıklığına
  // göre değişiyor — canlı testte gözlemlendi). Bu yüzden çağıranın istediği
  // değer YOK SAYILIYOR, her istekte doğrudan izin verilen tavan (2048)
  // kullanılıyor — bu kısa, tek seferlik İK istekleri için maliyet önemsiz.
  const payload = {
    contents: [{ role: "user", parts: [{ text: String(prompt) }] }],
    generationConfig: { maxOutputTokens: 2048 },
  };
  if (system) payload.systemInstruction = { parts: [{ text: String(system) }] };

  const model = "gemini-flash-lite-latest";
  const res = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + encodeURIComponent(apiKey),
    {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );

  let data;
  try {
    data = JSON.parse(res.getContentText());
  } catch (err) {
    return jsonError_("Gemini yanıtı ayrıştırılamadı: " + res.getContentText());
  }
  if (res.getResponseCode() !== 200) {
    return jsonError_("Gemini API hatası: " + ((data.error && data.error.message) || res.getContentText()));
  }
  const parts = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) || [];
  const text = parts.map((p) => p.text || "").join("");
  return jsonOk_({ text });
}

function jsonOk_(extra) {
  return ContentService.createTextOutput(JSON.stringify(Object.assign({ ok: true }, extra)))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonError_(message) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, error: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
