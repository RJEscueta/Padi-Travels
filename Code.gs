/**
 * Travel agency backend. Bind this script to your Google Sheet
 * (Sheet → Extensions → Apps Script), then run setup() once.
 * Sheets: Leads, Users, Images, Prices, Packages, Offers, Testimonials, Quotes, Settings.
 * Share this Sheet only with the owner's Google account. Staff use the admin
 * dashboard (admin.html), never the Sheet directly.
 */
var ADMIN_EMAIL  = 'you@example.com';   // first admin; temp password is emailed here
var NOTIFY_EMAIL = 'sales@example.com'; // gets an alert for every new lead
var AGENCY       = 'Your Agency Name';
var ADMIN_URL    = 'https://yoursite.example.com/admin.html'; // your deployed admin.html — used in reset-password emails
var SESSION_SECS = 21600;               // 6 hours (Apps Script cache maximum)
var RESET_SECS   = 1800;                // password-reset links expire after 30 minutes

var LEAD_COLS = ['lead_id','created','status','source','first','last','phone','email','country','contact',
  'destination','travel_date','return_date','travelers','travel_type','budget','message','heard',
  'selected_package','utm_source','utm_medium','utm_campaign','assigned_agent','next_followup','notes'];
var STATUSES = ['NEW','CONTACTED','QUALIFIED','QUOTE SENT','NEGOTIATING','BOOKED','LOST','FOLLOW-UP'];
var USER_COLS = ['email','name','role','salt','hash','active','must_change','created'];
var IMG_KEYS = ['hero','bangkok','phuket','pattaya','krabi','chiang-mai','ayutthaya'];
var DEST_KEYS = IMG_KEYS.filter(function (k) { return k !== 'hero'; });
var EDITABLE = ['status','assigned_agent','next_followup','notes'];
var SETTINGS_FIELDS = ['agency_name','phone','whatsapp','line','email','office','hours','about_text'];
var PKG_COLS   = ['id','title','tag','duration','inclusions','price','was_price','active','hot',
  'tour_type','duration_type','vehicle','style','price_number'];
var TOUR_TYPES = ['Sea Tour','Land Tour'];
var DURATION_TYPES = ['Half-Day','One-Day','Multi-Day'];
var VEHICLES = ['Speedboat','Long-tail Boat','Land Vehicle','Van'];
var TOUR_STYLES = ['Join Tour','Private Tour'];
var OFFER_COLS = ['id','title','validity','price','active'];
var TESTI_COLS = ['id','quote','name','active'];
var QUOTE_STATUSES = ['DRAFT','SENT','NEGOTIATING','ACCEPTED','DECLINED','EXPIRED'];
var QUOTE_COLS = ['quote_id','lead_id','created','valid_until','customer','phone','destination',
  'travel_date','travelers','items','currency','total','status','agent','notes'];

/* ---------- one-time setup ---------- */
function setup() {
  tab_('Leads', LEAD_COLS); tab_('Users', USER_COLS); tab_('Images', ['key', 'url']);
  tab_('Prices', ['key', 'price']); tab_('Packages', PKG_COLS); tab_('Offers', OFFER_COLS);
  tab_('Testimonials', TESTI_COLS); tab_('Quotes', QUOTE_COLS); tab_('Settings', ['key', 'value']);
  if (!findUser_(ADMIN_EMAIL)) {
    var pw = tempPw_();
    addUser_(ADMIN_EMAIL, 'Owner', 'admin', pw);
    MailApp.sendEmail(ADMIN_EMAIL, AGENCY + ' admin login', 'Email: ' + ADMIN_EMAIL + '\nTemporary password: ' + pw + '\nYou must change it at first sign-in.');
  }
}
/** Run from the editor if the owner is locked out. Emails a new temporary password. */
function resetOwnerPassword() {
  var pw = tempPw_(), u = findUser_(ADMIN_EMAIL);
  if (!u) return setup();
  setPw_(u.row, pw, 1);
  MailApp.sendEmail(ADMIN_EMAIL, AGENCY + ' admin password reset', 'Temporary password: ' + pw);
}
function tab_(name, cols) {
  var s = SpreadsheetApp.getActive(), t = s.getSheetByName(name) || s.insertSheet(name);
  if (t.getLastRow() === 0) { t.appendRow(cols); t.setFrozenRows(1); }
}

/* ---------- HTTP ---------- */
function doGet(e) {
  var a = e.parameter.action;
  if (a === 'images') return out_({ ok: true, images: images_() });   // kept for backward compatibility
  if (a === 'content') return out_({ ok: true, content: publicContent_() });
  return out_({ ok: true });
}
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents), a = d.action;
    if (a === 'lead') return out_(newLead_(d));
    if (a === 'login') return out_(login_(d));
    if (a === 'forgotPassword') return out_(forgotPassword_(d));
    if (a === 'resetWithToken') return out_(resetWithToken_(d));
    var s = auth_(d);
    if (!s) return out_({ ok: false, error: 'Please sign in again.', auth: false });
    if (a === 'logout') { CacheService.getScriptCache().remove('t_' + d.token); return out_({ ok: true }); }
    if (a === 'changePassword') return out_(changePw_(s, d));
    if (s.must) return out_({ ok: false, error: 'Change your temporary password first.', must: true });

    if (a === 'leads') return out_({ ok: true, leads: leads_(), statuses: STATUSES });
    if (a === 'updateLead') return out_(updateLead_(d));
    if (a === 'quotes') return out_({ ok: true, quotes: quotes_(), statuses: QUOTE_STATUSES });
    if (a === 'saveQuote') return out_(saveQuote_(d, s));

    var ADMIN_ONLY = ['users','addUser','resetUser','setActive','setImage','removeImage',
      'adminLists','saveSettings','savePackages','saveOffers','saveTestimonials','setPrice','removePrice'];
    if (ADMIN_ONLY.indexOf(a) > -1) {
      if (s.role !== 'admin') return out_({ ok: false, error: 'Admins only.' });
      if (a === 'users') return out_({ ok: true, users: users_() });
      if (a === 'addUser') return out_(addUserApi_(d));
      if (a === 'resetUser') return out_(resetUser_(d));
      if (a === 'setActive') return out_(setActive_(d, s));
      if (a === 'setImage') return out_(setImage_(d));
      if (a === 'removeImage') return out_(removeImage_(d));
      if (a === 'adminLists') return out_({ ok: true, content: adminContent_() });
      if (a === 'saveSettings') return out_(saveSettings_(d));
      if (a === 'savePackages') return out_(saveRows_('Packages', PKG_COLS, d.rows));
      if (a === 'saveOffers') return out_(saveRows_('Offers', OFFER_COLS, d.rows));
      if (a === 'saveTestimonials') return out_(saveRows_('Testimonials', TESTI_COLS, d.rows));
      if (a === 'setPrice') return out_(setPrice_(d));
      return out_(removePrice_(d));
    }
    return out_({ ok: false, error: 'Unknown action.' });
  } catch (err) { return out_({ ok: false, error: 'Server error.' }); }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- leads ---------- */
function clean_(v, n) {
  v = (v == null ? '' : String(v)).trim().slice(0, n || 500);
  return /^[=+\-@]/.test(v) ? "'" + v : v;   // blocks spreadsheet formula injection
}
function newLead_(d) {
  if (d.website) return { ok: true, lead_id: 'LEAD-0' };            // honeypot: bots fill this
  var c = CacheService.getScriptCache(), ip = clean_(d.phone, 30) || 'x', k = 'r_' + ip;
  if (c.get(k)) return { ok: false, error: 'Please wait a minute before sending again.' };
  c.put(k, '1', 60);
  if (!clean_(d.first) || !clean_(d.phone, 30)) return { ok: false, error: 'Name and phone are required.' };
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  var p = PropertiesService.getScriptProperties(), n = (parseInt(p.getProperty('lead_n') || '0', 10)) + 1;
  p.setProperty('lead_n', String(n)); lock.releaseLock();
  var id = 'LEAD-' + new Date().getFullYear() + '-' + ('000000' + n).slice(-6);
  var row = LEAD_COLS.map(function (col) {
    if (col === 'lead_id') return id;
    if (col === 'created') return new Date();
    if (col === 'status') return 'NEW';
    if (col === 'message') return clean_(d[col], 2000);
    return clean_(d[col], 200);
  });
  SpreadsheetApp.getActive().getSheetByName('Leads').appendRow(row);
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, 'New lead ' + id, 'From: ' + clean_(d.first) + ' ' + clean_(d.last) + '\nPhone: ' + clean_(d.phone, 30) + '\nDestination: ' + clean_(d.destination) + '\nOpen the admin dashboard.');
    var em = clean_(d.email, 120);
    if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em))
      MailApp.sendEmail(em, 'Thank you for contacting ' + AGENCY, 'Thank you for contacting ' + AGENCY + '. We’ve received your travel request (' + id + ') and our travel specialist will contact you shortly.');
  } catch (err) {}
  return { ok: true, lead_id: id };
}
function leads_() {
  var v = SpreadsheetApp.getActive().getSheetByName('Leads').getDataRange().getValues(), h = v.shift();
  return v.reverse().slice(0, 300).map(function (r) {
    var o = {}; h.forEach(function (k, i) { o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i]; }); return o;
  });
}
function updateLead_(d) {
  var t = SpreadsheetApp.getActive().getSheetByName('Leads'), ids = t.getRange(1, 1, t.getLastRow(), 1).getValues();
  for (var i = 1; i < ids.length; i++) if (ids[i][0] === d.lead_id) {
    EDITABLE.forEach(function (f) {
      if (d.fields && d.fields[f] != null) {
        var val = clean_(d.fields[f], 2000);
        if (f === 'status' && STATUSES.indexOf(val) < 0) return;
        t.getRange(i + 1, LEAD_COLS.indexOf(f) + 1).setValue(val);
      }
    });
    return { ok: true };
  }
  return { ok: false, error: 'Lead not found.' };
}

/* ---------- auth & users ---------- */
function b64_(b) { return Utilities.base64Encode(b); }
function hash_(pw, salt) {
  var h = pw + salt;
  for (var i = 0; i < 300; i++) h = b64_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt + i));
  return h;
}
function tempPw_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12) + 'Aa1'; }
function usersSheet_() { return SpreadsheetApp.getActive().getSheetByName('Users'); }
function findUser_(email) {
  email = String(email || '').trim().toLowerCase();
  var v = usersSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).toLowerCase() === email)
    return { row: i + 1, email: v[i][0], name: v[i][1], role: v[i][2], salt: v[i][3], hash: v[i][4], active: v[i][5], must: v[i][6] };
  return null;
}
function addUser_(email, name, role, pw) {
  var salt = Utilities.getUuid();
  usersSheet_().appendRow([email.toLowerCase(), name, role, salt, hash_(pw, salt), 1, 1, new Date()]);
}
function setPw_(row, pw, must) {
  var salt = Utilities.getUuid(), t = usersSheet_();
  t.getRange(row, 4, 1, 2).setValues([[salt, hash_(pw, salt)]]);
  t.getRange(row, 7).setValue(must);
}
function login_(d) {
  var c = CacheService.getScriptCache(), em = String(d.email || '').trim().toLowerCase(), k = 'f_' + em;
  var fails = parseInt(c.get(k) || '0', 10);
  if (fails >= 5) return { ok: false, error: 'Too many attempts. Try again in 15 minutes.' };
  var u = findUser_(em);
  if (!u || !u.active || hash_(String(d.password || ''), u.salt) !== u.hash) {
    c.put(k, String(fails + 1), 900); return { ok: false, error: 'Wrong email or password.' };
  }
  c.remove(k);
  var tok = Utilities.getUuid() + Utilities.getUuid();
  c.put('t_' + tok, JSON.stringify({ email: u.email, name: u.name, role: u.role, must: u.must == 1 }), SESSION_SECS);
  return { ok: true, token: tok, name: u.name, role: u.role, must: u.must == 1 };
}
function auth_(d) {
  var j = d.token && CacheService.getScriptCache().get('t_' + d.token);
  return j ? JSON.parse(j) : null;
}

/* ---------- self-service "forgot password" (email link), for owner and admins alike ---------- */
function forgotPassword_(d) {
  var c = CacheService.getScriptCache(), em = String(d.email || '').trim().toLowerCase();
  var generic = { ok: true, message: 'If that email has an account, a reset link has been sent.' };
  if (!em) return generic;
  var rk = 'fp_' + em;
  if (c.get(rk)) return generic;              // already sent one recently; don't spam or leak existence
  c.put(rk, '1', 300);
  var u = findUser_(em);
  if (u && u.active) {
    var token = Utilities.getUuid() + Utilities.getUuid();
    c.put('rt_' + token, em, RESET_SECS);
    var link = ADMIN_URL + (ADMIN_URL.indexOf('?') > -1 ? '&' : '?') + 'reset=' + token;
    try {
      MailApp.sendEmail(em, AGENCY + ' password reset',
        'We received a request to reset your admin password.\n\n' +
        'Set a new password here (link expires in 30 minutes):\n' + link +
        '\n\nIf you did not request this, you can ignore this email — your password will not change.');
    } catch (err) {}
  }
  return generic;
}
function resetWithToken_(d) {
  var c = CacheService.getScriptCache(), em = c.get('rt_' + String(d.token || ''));
  if (!em) return { ok: false, error: 'This reset link is invalid or has expired. Request a new one.' };
  if (!policy_(d.newPassword)) return { ok: false, error: 'New password needs 10+ characters with letters and numbers.' };
  var u = findUser_(em);
  if (!u || !u.active) return { ok: false, error: 'Account not found or inactive.' };
  setPw_(u.row, d.newPassword, 0);
  c.remove('rt_' + String(d.token));
  return { ok: true };
}
function policy_(pw) { return pw && pw.length >= 10 && /[a-z]/i.test(pw) && /\d/.test(pw); }
function changePw_(s, d) {
  var u = findUser_(s.email);
  if (!u || hash_(String(d.oldPassword || ''), u.salt) !== u.hash) return { ok: false, error: 'Current password is wrong.' };
  if (!policy_(d.newPassword)) return { ok: false, error: 'New password needs 10+ characters with letters and numbers.' };
  if (d.newPassword === d.oldPassword) return { ok: false, error: 'Choose a different password.' };
  setPw_(u.row, d.newPassword, 0);
  var c = CacheService.getScriptCache(); s.must = false; c.put('t_' + d.token, JSON.stringify(s), SESSION_SECS);
  return { ok: true };
}
function users_() {
  return usersSheet_().getDataRange().getValues().slice(1).map(function (r) {
    return { email: r[0], name: r[1], role: r[2], active: r[5] == 1 };
  });
}
function addUserApi_(d) {
  var em = String(d.email || '').trim().toLowerCase(), role = d.role === 'admin' ? 'admin' : 'agent';
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) return { ok: false, error: 'Enter a valid email.' };
  if (findUser_(em)) return { ok: false, error: 'User already exists.' };
  var pw = tempPw_(); addUser_(em, clean_(d.name, 60) || em, role, pw);
  return { ok: true, tempPassword: pw };
}
function resetUser_(d) {
  var u = findUser_(d.email); if (!u) return { ok: false, error: 'User not found.' };
  var pw = tempPw_(); setPw_(u.row, pw, 1); return { ok: true, tempPassword: pw };
}
function setActive_(d, s) {
  var u = findUser_(d.email); if (!u) return { ok: false, error: 'User not found.' };
  if (u.email.toLowerCase() === s.email.toLowerCase()) return { ok: false, error: 'You cannot deactivate yourself.' };
  usersSheet_().getRange(u.row, 6).setValue(d.active ? 1 : 0); return { ok: true };
}

/* ---------- images (files in Drive, links in the Images sheet) ---------- */
function images_() {
  var c = CacheService.getScriptCache(), j = c.get('imgs');
  if (j) return JSON.parse(j);
  var m = {}; SpreadsheetApp.getActive().getSheetByName('Images').getDataRange().getValues().slice(1)
    .forEach(function (r) { if (r[0] && r[1]) m[r[0]] = r[1]; });
  c.put('imgs', JSON.stringify(m), 300); return m;
}
function keyRow_(sheet, key) {
  var v = SpreadsheetApp.getActive().getSheetByName(sheet).getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === key) return i + 1; return 0;
}
function setImage_(d) {
  if (IMG_KEYS.indexOf(d.key) < 0) return { ok: false, error: 'Unknown image slot.' };
  var t = SpreadsheetApp.getActive().getSheetByName('Images'), url = '';
  var m = /^data:(image\/(?:webp|jpeg|png));base64,([A-Za-z0-9+\/=]+)$/.exec(d.dataUrl || '');
  if (m) {
    if (m[2].length > 2000000) return { ok: false, error: 'Image too large (max ~1.5 MB).' };
    var p = PropertiesService.getScriptProperties(), fid = p.getProperty('folder'), folder;
    try { folder = DriveApp.getFolderById(fid); } catch (e) { folder = DriveApp.createFolder(AGENCY + ' Site Images'); p.setProperty('folder', folder.getId()); }
    var f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], d.key + '-' + Date.now()));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    url = 'https://lh3.googleusercontent.com/d/' + f.getId();
  } else if (/^https:\/\/[^\s"']+$/.test(d.url || '')) url = d.url;
  else return { ok: false, error: 'Send a photo or an https link.' };
  var r = keyRow_('Images', d.key); if (r) t.getRange(r, 2).setValue(url); else t.appendRow([d.key, url]);
  CacheService.getScriptCache().remove('imgs'); return { ok: true, url: url };
}
function removeImage_(d) {
  var r = keyRow_('Images', d.key); if (r) SpreadsheetApp.getActive().getSheetByName('Images').deleteRow(r);
  CacheService.getScriptCache().remove('imgs'); return { ok: true };
}

/* ---------- prices (per destination "starting price" text) ---------- */
function prices_() {
  var m = {}; SpreadsheetApp.getActive().getSheetByName('Prices').getDataRange().getValues().slice(1)
    .forEach(function (r) { if (r[0] && r[1]) m[r[0]] = r[1]; });
  return m;
}
function setPrice_(d) {
  if (DEST_KEYS.indexOf(d.key) < 0) return { ok: false, error: 'Unknown destination.' };
  var t = SpreadsheetApp.getActive().getSheetByName('Prices'), val = clean_(d.price, 40);
  var r = keyRow_('Prices', d.key); if (r) t.getRange(r, 2).setValue(val); else t.appendRow([d.key, val]);
  return { ok: true };
}
function removePrice_(d) {
  var r = keyRow_('Prices', d.key); if (r) SpreadsheetApp.getActive().getSheetByName('Prices').deleteRow(r);
  return { ok: true };
}

/* ---------- settings (agency name & contact details) ---------- */
function settings_() {
  var m = {}; SpreadsheetApp.getActive().getSheetByName('Settings').getDataRange().getValues().slice(1)
    .forEach(function (r) { if (r[0]) m[r[0]] = r[1]; });
  return m;
}
function saveSettings_(d) {
  var t = SpreadsheetApp.getActive().getSheetByName('Settings');
  SETTINGS_FIELDS.forEach(function (f) {
    if (d.fields && d.fields[f] != null) {
      var val = clean_(d.fields[f], f === 'about_text' ? 1200 : 300), r = keyRow_('Settings', f);
      if (r) t.getRange(r, 2).setValue(val); else t.appendRow([f, val]);
    }
  });
  return { ok: true };
}

/* ---------- packages / offers / testimonials (admin edits full list at once) ---------- */
function readRows_(sheet, cols) {
  var v = SpreadsheetApp.getActive().getSheetByName(sheet).getDataRange().getValues().slice(1);
  return v.filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {}; cols.forEach(function (c, i) { o[c] = r[i]; }); return o;
  });
}
function saveRows_(sheet, cols, rows) {
  if (!Array.isArray(rows)) return { ok: false, error: 'No data received.' };
  if (rows.length > 200) return { ok: false, error: 'Too many rows.' };
  var t = SpreadsheetApp.getActive().getSheetByName(sheet), last = t.getLastRow();
  if (last > 1) t.getRange(2, 1, last - 1, cols.length).clearContent();
  if (rows.length) {
    var data = rows.map(function (r, idx) {
      return cols.map(function (c) {
        if (c === 'id') return clean_(r.id || ('id' + Date.now() + idx), 40);
        if (c === 'active' || c === 'hot') return r[c] ? 1 : 0;
        if (c === 'price_number') return Number(r[c]) || 0;
        return clean_(r[c], c === 'inclusions' ? 3000 : 500);
      });
    });
    t.getRange(2, 1, data.length, cols.length).setValues(data);
  }
  return { ok: true };
}

/* ---------- content for the public website ---------- */
function publicContent_() {
  function act(rows) { return rows.filter(function (r) { return r.active == 1 || r.active === true; }); }
  return {
    settings: settings_(), images: images_(), prices: prices_(),
    packages: act(readRows_('Packages', PKG_COLS)),
    offers: act(readRows_('Offers', OFFER_COLS)),
    testimonials: act(readRows_('Testimonials', TESTI_COLS))
  };
}
/** Unfiltered lists (including inactive) for the admin dashboard editors. */
function adminContent_() {
  return {
    settings: settings_(), images: images_(), prices: prices_(),
    packages: readRows_('Packages', PKG_COLS),
    offers: readRows_('Offers', OFFER_COLS),
    testimonials: readRows_('Testimonials', TESTI_COLS)
  };
}

/* ---------- quotes ---------- */
function quotesSheet_() { return SpreadsheetApp.getActive().getSheetByName('Quotes'); }
function quotes_() {
  var v = quotesSheet_().getDataRange().getValues(), h = v.shift();
  return v.reverse().slice(0, 300).map(function (r) {
    var o = {}; h.forEach(function (k, i) {
      o[k] = r[i] instanceof Date ? r[i].toISOString() : r[i];
    });
    try { o.items = JSON.parse(o.items || '[]'); } catch (e) { o.items = []; }
    return o;
  });
}
function quoteRow_(id) {
  var v = quotesSheet_().getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === id) return i + 1; return 0;
}
function saveQuote_(d, s) {
  if (QUOTE_STATUSES.indexOf(d.status) < 0) d.status = 'DRAFT';
  var items = Array.isArray(d.items) ? d.items.slice(0, 40).map(function (it) {
    return { desc: clean_(it.desc, 200), qty: Math.max(1, parseInt(it.qty, 10) || 1), price: Number(it.price) || 0 };
  }) : [];
  var total = items.reduce(function (sum, it) { return sum + it.qty * it.price; }, 0);
  var t = quotesSheet_(), id = clean_(d.quote_id, 40), row = quoteRow_(id);
  if (!row) {
    var p = PropertiesService.getScriptProperties(), n = (parseInt(p.getProperty('quote_n') || '0', 10)) + 1;
    p.setProperty('quote_n', String(n));
    id = 'QT-' + new Date().getFullYear() + '-' + ('000000' + n).slice(-6);
  }
  var vals = {
    quote_id: id, created: row ? undefined : new Date(), valid_until: clean_(d.valid_until, 20),
    customer: clean_(d.customer, 120), phone: clean_(d.phone, 30), destination: clean_(d.destination, 120),
    travel_date: clean_(d.travel_date, 20), travelers: clean_(d.travelers, 10),
    items: JSON.stringify(items), currency: clean_(d.currency, 10) || 'THB', total: total,
    status: d.status, agent: clean_(d.agent || s.name, 80), notes: clean_(d.notes, 2000)
  };
  if (row) {
    QUOTE_COLS.forEach(function (c, i) { if (vals[c] !== undefined) t.getRange(row, i + 1).setValue(vals[c]); });
  } else {
    t.appendRow(QUOTE_COLS.map(function (c) { return vals[c] !== undefined ? vals[c] : ''; }));
  }
  if (d.markSent && d.lead_id) {
    updateLead_({ lead_id: d.lead_id, fields: { status: 'QUOTE SENT', next_followup: d.followup || '' } });
  }
  return { ok: true, quote_id: id, total: total };
}
