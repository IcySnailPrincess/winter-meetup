// Backend for the winter meetup page: availability + Secret Santa.
// Deploy: Deploy > Manage deployments > Edit > Version: New version (keeps the same /exec URL).
const SHEET_ID = '1xjXAdBd90J1rFODyk3pdB8M5BWEF9a3peQ6510Ljn3g';
const PROPS = PropertiesService.getScriptProperties();

// ---------- helpers ----------
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function sheet_() { return ss_().getSheets()[0]; } // availability
function tab_(name, header) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); }
  return sh;
}
function people_() { return tab_('santa_people', ['name', 'pin_hash', 'salt', 'wishlist', 'joined']); }
function msgs_() { return tab_('santa_msgs', ['giver', 'from_role', 'text', 'time']); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function norm_(s) { return String(s || '').trim().toLowerCase(); }
function cleanName_(s) { return String(s || '').trim().replace(/\s+/g, ' ').slice(0, 40); }
function hash_(pin, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function validPin_(pin) { return /^\d{4,8}$/.test(String(pin || '')); }
function getJSON_(key, dflt) { const v = PROPS.getProperty(key); return v ? JSON.parse(v) : dflt; }
function setJSON_(key, val) { PROPS.setProperty(key, JSON.stringify(val)); }

// Brute-force guard: 5 wrong PINs per key -> locked 15 min.
function checkLock_(key) {
  const n = Number(CacheService.getScriptCache().get('fail:' + key) || 0);
  if (n >= 5) throw new Error('Too many wrong PINs. Try again in 15 minutes.');
}
function fail_(key) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get('fail:' + key) || 0) + 1;
  c.put('fail:' + key, String(n), 900);
  throw new Error('Wrong name or PIN.');
}

// ---------- availability ----------
function rows_() {
  const values = sheet_().getDataRange().getValues().slice(1);
  return values.filter(r => r[0]).map(r => {
    let slots = {};
    try { slots = JSON.parse(r[1] || '{}'); } catch (e) {}
    return { name: String(r[0]), slots: slots, updated: r[2] ? new Date(r[2]).toISOString() : null };
  });
}
function saveAvailability_(body) {
  const name = cleanName_(body.name);
  if (!name) throw new Error('name required');
  const sh = sheet_();
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) if (norm_(data[i][0]) === norm_(name)) { rowIdx = i + 1; break; }
  if (body.delete) { if (rowIdx > 0) sh.deleteRow(rowIdx); return rows_(); }
  const slots = {};
  const allowed = { free: 1, maybe: 1, busy: 1 };
  Object.keys(body.slots || {}).forEach(k => {
    if (/^\d{4}-\d{2}-\d{2}-(AM|PM)$/.test(k) && allowed[body.slots[k]]) slots[k] = body.slots[k];
  });
  const row = [name, JSON.stringify(slots), new Date()];
  if (rowIdx > 0) sh.getRange(rowIdx, 1, 1, 3).setValues([row]); else sh.appendRow(row);
  return rows_();
}

// ---------- secret santa ----------
function santaPeople_() {
  const sh = people_();
  return sh.getDataRange().getValues().slice(1).map((r, i) => ({
    row: i + 2, name: String(r[0]), hash: String(r[1]), salt: String(r[2]), wishlist: String(r[3] || '')
  })).filter(p => p.name);
}
function findPerson_(name) { return santaPeople_().find(p => norm_(p.name) === norm_(name)); }
function auth_(name, pin) {
  const key = 'p:' + norm_(name);
  checkLock_(key);
  const p = findPerson_(name);
  if (!p || !validPin_(pin) || hash_(pin, p.salt) !== p.hash) fail_(key);
  return p;
}
function assignments_() { return getJSON_('santa_assign', null); } // {giverNorm: receiverName}
function santaPublic_() {
  return {
    people: santaPeople_().map(p => p.name),
    budget: PROPS.getProperty('santa_budget') || '',
    date: PROPS.getProperty('santa_date') || '',
    drawn: !!assignments_(),
    organizerSet: !!PROPS.getProperty('santa_org_hash')
  };
}
function thread_(giver) {
  return msgs_().getDataRange().getValues().slice(1)
    .filter(r => norm_(r[0]) === norm_(giver))
    .map(r => ({ from: String(r[1]), text: String(r[2]), time: r[3] ? new Date(r[3]).toISOString() : null }));
}
function santaFor_(name) { // who drew `name`
  const a = assignments_() || {};
  const g = Object.keys(a).find(k => norm_(a[k]) === norm_(name));
  if (!g) return null;
  const p = santaPeople_().find(x => norm_(x.name) === g);
  return p ? p.name : null;
}

function santaJoin_(body) {
  const name = cleanName_(body.name);
  const pin = String(body.pin || '');
  const wishlist = String(body.wishlist || '').slice(0, 1000);
  if (!name) throw new Error('Enter your name.');
  if (!validPin_(pin)) throw new Error('PIN must be 4–8 digits.');
  const existing = findPerson_(name);
  if (existing) {
    auth_(name, pin);
    people_().getRange(existing.row, 4).setValue(wishlist);
  } else {
    if (assignments_()) throw new Error('Names were already drawn. Ask the organizer to redraw to add you.');
    const salt = Utilities.getUuid();
    people_().appendRow([name, hash_(pin, salt), salt, wishlist, new Date()]);
  }
  return santaPublic_();
}

function santaMe_(body) {
  const me = auth_(body.name, body.pin);
  const a = assignments_();
  const out = { name: me.name, wishlist: me.wishlist, drawn: !!a, giftee: null, threads: { giftee: [], santa: [] } };
  if (a && a[norm_(me.name)]) {
    const g = findPerson_(a[norm_(me.name)]);
    out.giftee = { name: g ? g.name : a[norm_(me.name)], wishlist: g ? g.wishlist : '' };
    out.threads.giftee = thread_(me.name); // I'm the santa in this thread
  }
  const mySanta = santaFor_(me.name);
  if (mySanta) out.threads.santa = thread_(mySanta); // I'm the giftee in my santa's thread
  return out;
}

function santaMsg_(body) {
  const me = auth_(body.name, body.pin);
  const text = String(body.text || '').trim().slice(0, 500);
  if (!text) throw new Error('Message is empty.');
  if (!assignments_()) throw new Error('Names have not been drawn yet.');
  if (body.to === 'giftee') msgs_().appendRow([me.name, 'santa', text, new Date()]);
  else if (body.to === 'santa') {
    const s = santaFor_(me.name);
    if (!s) throw new Error('No Secret Santa found for you.');
    msgs_().appendRow([s, 'giftee', text, new Date()]);
  } else throw new Error('bad recipient');
  return santaMe_(body);
}

function draw_(names, exclusions) {
  const ex = {};
  exclusions.forEach(([a, b]) => { ex[norm_(a) + '|' + norm_(b)] = 1; ex[norm_(b) + '|' + norm_(a)] = 1; });
  for (let t = 0; t < 5000; t++) {
    const rec = names.slice();
    for (let i = rec.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [rec[i], rec[j]] = [rec[j], rec[i]]; }
    let ok = true;
    for (let i = 0; i < names.length && ok; i++) {
      if (norm_(names[i]) === norm_(rec[i]) || ex[norm_(names[i]) + '|' + norm_(rec[i])]) ok = false;
    }
    if (ok) { const out = {}; names.forEach((n, i) => out[norm_(n)] = rec[i]); return out; }
  }
  throw new Error('No valid draw is possible with these exclusions. Remove some and try again.');
}

function santaAdmin_(body) {
  const pin = String(body.pin || '');
  const orgHash = PROPS.getProperty('santa_org_hash');
  if (body.op === 'claim') {
    if (orgHash) throw new Error('Organizer PIN is already set.');
    if (!validPin_(pin)) throw new Error('PIN must be 4–8 digits.');
    const salt = Utilities.getUuid();
    PROPS.setProperty('santa_org_salt', salt);
    PROPS.setProperty('santa_org_hash', hash_(pin, salt));
  } else {
    checkLock_('org');
    if (!orgHash || !validPin_(pin) || hash_(pin, PROPS.getProperty('santa_org_salt')) !== orgHash) fail_('org');
  }
  const op = body.op;
  if (op === 'settings') {
    PROPS.setProperty('santa_budget', String(body.budget || '').slice(0, 40));
    PROPS.setProperty('santa_date', String(body.date || '').slice(0, 40));
  } else if (op === 'exclusions') {
    const pairs = (body.pairs || []).filter(p => Array.isArray(p) && p.length === 2 && p[0] && p[1] && norm_(p[0]) !== norm_(p[1]))
      .map(p => [cleanName_(p[0]), cleanName_(p[1])]).slice(0, 100);
    setJSON_('santa_excl', pairs);
  } else if (op === 'draw') {
    const names = santaPeople_().map(p => p.name);
    if (names.length < 3) throw new Error('Need at least 3 people to draw.');
    setJSON_('santa_assign', draw_(names, getJSON_('santa_excl', [])));
    const m = msgs_(); if (m.getLastRow() > 1) m.deleteRows(2, m.getLastRow() - 1); // fresh threads after a (re)draw
  } else if (op === 'reset') {
    PROPS.deleteProperty('santa_assign');
    const m = msgs_(); if (m.getLastRow() > 1) m.deleteRows(2, m.getLastRow() - 1);
  } else if (op === 'remove') {
    const p = findPerson_(body.name);
    if (p) people_().deleteRow(p.row);
    if (assignments_()) PROPS.deleteProperty('santa_assign'); // removing someone invalidates the draw
  } else if (op === 'release') {
    ['santa_org_hash', 'santa_org_salt'].forEach(k => PROPS.deleteProperty(k));
  } else if (op !== 'claim' && op !== 'login') {
    throw new Error('unknown op');
  }
  return Object.assign(santaPublic_(), { exclusions: getJSON_('santa_excl', []) });
}

// ---------- entry points ----------
function doGet() {
  return json_({ availability: rows_(), santa: santaPublic_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case undefined: case null: case '': return json_(saveAvailability_(body)); // legacy availability save
      case 'santaJoin': return json_(santaJoin_(body));
      case 'santaMe': return json_(santaMe_(body));
      case 'santaMsg': return json_(santaMsg_(body));
      case 'santaAdmin': return json_(santaAdmin_(body));
      default: return json_({ error: 'unknown action' });
    }
  } catch (err) {
    return json_({ error: err.message });
  } finally {
    lock.releaseLock();
  }
}
