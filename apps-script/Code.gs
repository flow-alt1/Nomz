/**
 * NOMZ KITCHEN — menu backend (Google Apps Script + Google Sheet)
 *
 * Setup (once):
 *  1. Create a Google Sheet, then Extensions > Apps Script. Paste this file.
 *  2. Run setup() once and approve permissions. Open View > Logs (or the
 *     Execution log) to read your admin password. Change it any time in
 *     Project Settings > Script properties > ADMIN_PASSWORD.
 *  3. Deploy > New deployment > Web app.
 *       Execute as: Me    Who has access: Anyone
 *     Copy the /exec URL into config.js as API_URL.
 *  4. Square checkout: Project Settings > Script properties > add
 *       SQUARE_TOKEN        your Square Production Access Token (keep private)
 *       SQUARE_LOCATION_ID  your Square Location ID
 *     Then run testSquare() and read the Execution log.
 *  After editing this code later, use Deploy > Manage deployments > Edit >
 *  New version so the same URL keeps working.
 */

const SHEET_NAME = 'Menu';
const COLS = ['id','sort','featured','available','nameEn','nameFa','price','category',
              'description','ingredients','allergens','image','url','updated'];
const STORE_HOST = 'nomzkitchen.square.site';
const CACHE_KEY = 'menu_public_v1';

/* ------------------------------------------------------------------ */
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() < 2) {
    const now = new Date();
    const rows = SEED.map(it => COLS.map(c => c === 'updated' ? now : (it[c] === undefined ? '' : it[c])));
    sh.getRange(2, 1, rows.length, COLS.length).setValues(rows);
  }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASSWORD')) {
    const pw = 'nomz-' + Utilities.getUuid().slice(0, 8);
    props.setProperty('ADMIN_PASSWORD', pw);
    Logger.log('Admin password created: ' + pw);
  } else {
    Logger.log('Admin password already set (see Project Settings > Script properties).');
  }
  photoFolder_(); // asks for Google Drive permission (photo uploads)
  clearCache_();
}

/* --------------------------- Public API --------------------------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'list') {
    const cache = CacheService.getScriptCache();
    const hit = cache.get(CACHE_KEY);
    if (hit) return raw_(hit);
    const items = readItems_().filter(it => it.available);
    const body = JSON.stringify({ ok: true, items: items });
    try { cache.put(CACHE_KEY, body, 300); } catch (err) {}
    return raw_(body);
  }
  return json_({ ok: false, error: 'Unknown action' });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }

  // Public: customers starting a checkout (no password needed)
  if (req.action === 'checkout') {
    try { return json_(Object.assign({ ok: true }, checkout_(req))); }
    catch (err) { return json_({ ok: false, error: String(err.message || err) }); }
  }

  const auth = checkAuth_(req.password);
  if (auth !== true) return json_({ ok: false, error: auth });

  try {
    switch (req.action) {
      case 'login':     return json_({ ok: true });
      case 'adminList': return json_({ ok: true, items: readItems_() });
      case 'save':      return json_({ ok: true, item: saveItem_(req.item || {}) });
      case 'delete':    deleteItem_(String(req.id || '')); return json_({ ok: true });
      case 'scrape':    return json_({ ok: true, item: scrape_(String(req.url || '')) });
      case 'upload':    return json_(Object.assign({ ok: true }, upload_(req)));
      default:          return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ----------------------------- Auth ------------------------------- */
function checkAuth_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= 10) return 'Too many wrong passwords. Try again in 15 minutes.';
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (real && pw && String(pw) === real) return true;
  cache.put('fails', String(fails + 1), 900);
  return 'Wrong password.';
}

/* ----------------------------- Data ------------------------------- */
function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Run setup() first.');
  return sh;
}

function readItems_() {
  const values = sheet_().getDataRange().getValues();
  const head = values.shift() || [];
  return values.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((h, i) => { o[h] = r[i]; });
    o.id = String(o.id);
    o.sort = Number(o.sort) || 999;
    o.price = o.price === '' ? '' : Number(o.price);
    o.featured = o.featured === true || String(o.featured).toUpperCase() === 'TRUE';
    o.available = !(o.available === false || String(o.available).toUpperCase() === 'FALSE');
    o.updated = o.updated instanceof Date ? o.updated.toISOString() : String(o.updated || '');
    return o;
  }).sort((a, b) => a.sort - b.sort);
}

function productId_(url) {
  const m = String(url).match(/\/product\/[^\/]+\/([A-Za-z0-9]+)/);
  return m ? m[1].toUpperCase() : '';
}

function cleanUrl_(url) {
  const u = String(url || '').trim();
  const m = u.match(/^https:\/\/([^\/?#]+)(\/product\/[^?#]+)/i);
  if (!m || m[1].toLowerCase() !== STORE_HOST) throw new Error('Paste a product link from ' + STORE_HOST);
  return 'https://' + STORE_HOST + m[2];
}

function saveItem_(it) {
  const rawUrl = String(it.url || '').trim();
  const url = rawUrl ? cleanUrl_(rawUrl) : '';
  if (url && !productId_(url)) throw new Error('That link does not look like a Square product page.');
  const nameEn = String(it.nameEn || '').trim();
  if (!nameEn) throw new Error('English name is required.');
  const nameFa = String(it.nameFa || '').trim();
  const description = String(it.description || '').trim();
  const price = String(it.price).trim() === '' ? '' : Number(it.price);
  if (price !== '' && (isNaN(price) || price < 0)) throw new Error('Price must be a number, like 16.95');
  if (!url && price === '') throw new Error('Add a price. It is required for dishes without a Square store link.');
  const image = String(it.image || '').trim();
  if (image && !/^https:\/\//i.test(image)) throw new Error('Photo URL must start with https://');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const items = readItems_();
    const prevId = String(it.originalId || '');
    let id = url ? productId_(url) : prevId;
    // Dishes without a Square store link get (or update) their own item in the Square catalog
    if (!url) id = syncSquareItem_(id, nameEn, nameFa, description, price) || id || ('LOCAL-' + Utilities.getUuid().slice(0, 8).toUpperCase());
    const existing = items.find(x => x.id === id);
    const target = existing || (prevId ? items.find(x => x.id === prevId) : null);
    const row = {
      id: id,
      sort: Number(it.sort) || (target ? target.sort : (items.reduce((m, x) => Math.max(m, x.sort === 999 ? 0 : x.sort), 0) + 1)),
      featured: !!it.featured,
      available: it.available !== false,
      nameEn: nameEn, nameFa: nameFa, price: price,
      category: String(it.category || '').trim() || 'Menu',
      description: description,
      ingredients: String(it.ingredients || '').trim(),
      allergens: String(it.allergens || '').trim(),
      image: image, url: url, updated: new Date()
    };
    const values = COLS.map(c => row[c]);
    if (target) {
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
      sh.getRange(ids.indexOf(target.id) + 2, 1, 1, COLS.length).setValues([values]);
    } else {
      sh.appendRow(values);
    }
    clearCache_();
    row.updated = row.updated.toISOString();
    return row;
  } finally {
    lock.releaseLock();
  }
}

// Creates or updates a Square catalog item for a dish that isn't on the Square store yet.
// Returns the Square item id, or null when Square isn't connected.
function syncSquareItem_(id, nameEn, nameFa, description, price) {
  try { sqProps_(); } catch (e) { return null; }
  const name = nameFa ? nameEn + ' | ' + nameFa : nameEn;
  const money = { amount: Math.round(Number(price) * 100), currency: 'USD' };
  let obj = null;
  if (id && id.indexOf('LOCAL-') !== 0) {
    try { obj = sq_('get', '/catalog/object/' + encodeURIComponent(id)).object; } catch (e) {}
  }
  if (obj && obj.type === 'ITEM') {
    obj.item_data.name = name;
    obj.item_data.description = description;
    delete obj.item_data.description_html;
    delete obj.item_data.description_plaintext;
    const v = obj.item_data.variations && obj.item_data.variations[0];
    if (v) { v.item_variation_data.pricing_type = 'FIXED_PRICING'; v.item_variation_data.price_money = money; }
  } else {
    obj = { type: 'ITEM', id: '#dish', item_data: { name: name, description: description,
      variations: [{ type: 'ITEM_VARIATION', id: '#dishvar',
        item_variation_data: { name: 'Regular', pricing_type: 'FIXED_PRICING', price_money: money } }] } };
  }
  const res = sq_('post', '/catalog/object', { idempotency_key: Utilities.getUuid(), object: obj });
  const saved = res.catalog_object;
  const vid = pickVariation_(saved);
  if (vid) CacheService.getScriptCache().put('var_' + saved.id, vid, 21600);
  return saved.id;
}

/* ------------------------- Photo upload (Drive) ------------------------- */
function photoFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Nomz Website Photos');
  p.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function upload_(req) {
  const m = String(req.data || '').match(/^data:(image\/(jpeg|png|webp));base64,(.+)$/);
  if (!m) throw new Error('Upload a JPG, PNG or WEBP photo.');
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 8 * 1024 * 1024) throw new Error('That photo is too large (max 8 MB).');
  const ext = m[2] === 'jpeg' ? 'jpg' : m[2];
  const file = photoFolder_().createFile(Utilities.newBlob(bytes, m[1], 'dish-' + Date.now() + '.' + ext));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: 'https://lh3.googleusercontent.com/d/' + file.getId() };
}

function deleteItem_(id) {
  if (!id) throw new Error('Missing id');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]));
    const idx = ids.indexOf(id);
    if (idx < 0) throw new Error('Dish not found');
    sh.deleteRow(idx + 2);
    clearCache_();
  } finally {
    lock.releaseLock();
  }
}

/* ------------------- Read details from a Square page ------------------- */
function scrape_(url) {
  const clean = cleanUrl_(url);
  const res = UrlFetchApp.fetch(clean, { muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) throw new Error('Square returned ' + res.getResponseCode() + ' for that link.');
  const html = res.getContentText();
  const meta = name => {
    const tags = html.match(/<meta\b[^>]*>/gi) || [];
    for (const tag of tags) {
      const key = tag.match(/(?:property|name)\s*=\s*("([^"]*)"|'([^']*)')/i);
      if (!key || (key[2] || key[3] || '').toLowerCase() !== name) continue;
      const val = tag.match(/content\s*=\s*("([^"]*)"|'([^']*)')/i);
      if (val) return decode_(val[2] !== undefined ? val[2] : val[3]);
    }
    return '';
  };
  let title = meta('og:title') || (html.match(/<title>([^<]*)<\/title>/i) || [])[1] || '';
  title = decode_(title).replace(/\s*\|\s*Nomz\s*$/i, '');
  const parts = title.split('|');
  const desc = meta('og:description') || meta('description');
  const ing = desc.split(/Ingredients:/i);
  const allerg = (ing[1] || '').split(/Allergens?:/i);
  return {
    id: productId_(clean),
    url: clean,
    nameEn: (parts[0] || '').trim(),
    nameFa: parts.slice(1).join('|').trim(),
    description: ing[0].trim(),
    ingredients: (allerg[0] || '').trim(),
    allergens: (allerg[1] || '').trim(),
    image: meta('og:image')
  };
}

function decode_(s) {
  return String(s)
    .replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCharCode(parseInt(n, 16)));
}

/* ========================= SQUARE CHECKOUT ========================= */
const SQUARE_API = 'https://connect.squareup.com/v2';

function sqProps_() {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty('SQUARE_TOKEN'), location = p.getProperty('SQUARE_LOCATION_ID');
  if (!token || !location) throw new Error('Online checkout is not connected yet.');
  return { token: token, location: location };
}

function sq_(method, path, body) {
  const opts = {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + sqProps_().token },
    muteHttpExceptions: true
  };
  if (body) opts.payload = JSON.stringify(body);
  const res = UrlFetchApp.fetch(SQUARE_API + path, opts);
  const code = res.getResponseCode();
  let j = {};
  try { j = JSON.parse(res.getContentText() || '{}'); } catch (e) {}
  if (code >= 300) {
    const e0 = j.errors && j.errors[0];
    throw new Error(e0 ? (e0.detail || e0.code) : 'Square error ' + code);
  }
  return j;
}

function pickVariation_(obj) {
  if (!obj) return null;
  if (obj.type === 'ITEM_VARIATION') return obj.id;
  const v = obj.item_data && obj.item_data.variations;
  return v && v.length ? v[0].id : null;
}

// Finds the Square catalog variation for a dish: first by the ID in its
// square.site link, then by name. Cached for 6 hours.
function variationFor_(item) {
  const cache = CacheService.getScriptCache(), key = 'var_' + item.id;
  const hit = cache.get(key);
  if (hit) return hit === 'none' ? null : hit;
  let vid = null;
  try { vid = pickVariation_(sq_('get', '/catalog/object/' + encodeURIComponent(item.id)).object); } catch (e) {}
  if (!vid) {
    try {
      const words = String(item.nameEn).split(/\s+/).filter(w => w.length > 2).slice(0, 3);
      const j = sq_('post', '/catalog/search', {
        object_types: ['ITEM'], limit: 100,
        query: { text_query: { keywords: words } }
      });
      const want = String(item.nameEn).trim().toLowerCase();
      const m = (j.objects || []).find(o => o.item_data &&
        String(o.item_data.name || '').trim().toLowerCase().indexOf(want) === 0);
      vid = pickVariation_(m);
    } catch (e) {}
  }
  cache.put(key, vid || 'none', vid ? 21600 : 600);
  return vid;
}

function checkout_(req) {
  const location = sqProps_().location;
  const menu = readItems_().filter(x => x.available);
  const lines = (Array.isArray(req.items) ? req.items : []).slice(0, 40).map(l => ({
    item: menu.find(x => x.id === String(l.id)),
    qty: Math.max(1, Math.min(50, parseInt(l.qty, 10) || 0))
  })).filter(l => l.item);
  if (!lines.length) throw new Error('Your order is empty.');

  const line_items = lines.map(({ item, qty }) => {
    const vid = variationFor_(item);
    if (vid) return { catalog_object_id: vid, quantity: String(qty) };
    // Not found in the Square catalog: use the name and price from the sheet
    if (item.price === '' || isNaN(item.price)) throw new Error(item.nameEn + ' has no price set.');
    return {
      name: item.nameFa ? item.nameEn + ' | ' + item.nameFa : item.nameEn,
      quantity: String(qty),
      base_price_money: { amount: Math.round(Number(item.price) * 100), currency: 'USD' }
    };
  });

  const method = req.method === 'Delivery' ? 'Delivery' : 'Pickup';
  const note = String(req.note || '').replace(/\s+/g, ' ').trim().slice(0, 300);
  const body = {
    idempotency_key: Utilities.getUuid(),
    order: { location_id: location, line_items: line_items, pricing_options: { auto_apply_taxes: true } },
    checkout_options: { ask_for_shipping_address: method === 'Delivery' },
    payment_note: (method.toUpperCase() + ' - nomzkitchen.com' + (note ? ' - Note: ' + note : '')).slice(0, 500)
  };
  const ret = String(req.returnUrl || '');
  if (/^https:\/\/\S{1,250}$/.test(ret)) body.checkout_options.redirect_url = ret;

  const j = sq_('post', '/online-checkout/payment-links', body);
  return { url: j.payment_link.url };
}

// Run this from the editor to check the Square connection.
function testSquare() {
  const loc = sq_('get', '/locations/' + sqProps_().location).location;
  Logger.log('Connected to Square location: ' + loc.name + ' (' + loc.status + ')');
  readItems_().forEach(it => {
    CacheService.getScriptCache().remove('var_' + it.id);
    const v = variationFor_(it);
    Logger.log((v ? 'OK      ' : 'NO MATCH (uses sheet price, no auto tax) ') + it.nameEn + (v ? '  -> ' + v : ''));
  });
}

/* ----------------------------- Helpers ---------------------------- */
function clearCache_() { CacheService.getScriptCache().remove(CACHE_KEY); }
function json_(o) { return raw_(JSON.stringify(o)); }
function raw_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON); }

/* Initial menu (used by setup() only) */
const SEED = [
  {
    "id": "NS7PE7TU5FFYZAEXSGXOM7TO",
    "sort": 1,
    "featured": true,
    "available": true,
    "nameEn": "Zereshk Polo with Chicken",
    "nameFa": "زرشک پلو با مرغ",
    "price": 16.95,
    "category": "Rice dishes",
    "description": "Traditional Persian saffron rice topped with sweet and tangy barberries, served with tender chicken cooked in a flavorful tomato and saffron sauce.",
    "ingredients": "Basmati rice, chicken, barberries, saffron, onion, tomato paste, turmeric, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/7TKS5KLNBRV6ZNVOECYHAVRB.jpeg",
    "url": "https://nomzkitchen.square.site/product/zereshk-polo-with-chicken-/NS7PE7TU5FFYZAEXSGXOM7TO"
  },
  {
    "id": "ZUTM5VYXZ2R4KVZQJCGIUUYX",
    "sort": 2,
    "featured": true,
    "available": true,
    "nameEn": "Baghali Polo with Lamb Shank",
    "nameFa": "باقالی پلو با ماهیچه",
    "price": 19.95,
    "category": "Rice dishes",
    "description": "A classic Persian specialty: a tender, slow-braised lamb shank served alongside fragrant basmati rice mixed with fresh dill, fava beans, and saffron.",
    "ingredients": "Lamb shank, basmati rice, fava beans, dill, saffron, onion, turmeric, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/RAGT3246OYZJDSCAQLQ33WUJ.jpeg",
    "url": "https://nomzkitchen.square.site/product/baghali-polo-with-lamb-shank-/ZUTM5VYXZ2R4KVZQJCGIUUYX"
  },
  {
    "id": "AT22IAPJ355ZEKMDRFDQBQS7",
    "sort": 3,
    "featured": false,
    "available": true,
    "nameEn": "Baghali Polo with Chicken",
    "nameFa": "باقالی پلو با مرغ",
    "price": 17.95,
    "category": "Rice dishes",
    "description": "Basmati rice mixed with fragrant dill, tender fava beans, and saffron, served with juicy, slow-cooked chicken in a saffron sauce.",
    "ingredients": "Basmati rice, chicken, fava beans, dill, saffron, onion, turmeric, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/OAKORTVFOMSNQSAD4PR62RLR.jpeg",
    "url": "https://nomzkitchen.square.site/product/baghali-polo-with-chicken-/AT22IAPJ355ZEKMDRFDQBQS7"
  },
  {
    "id": "HD5YIFAMRRL7DJP2O7547KGQ",
    "sort": 4,
    "featured": false,
    "available": true,
    "nameEn": "Sabzi Polo with Salmon",
    "nameFa": "سبزی پلو با ماهی سالمون",
    "price": 21.95,
    "category": "Rice dishes",
    "description": "Fragrant Persian basmati rice mixed with fresh herbs, served with a tender, perfectly cooked salmon fillet.",
    "ingredients": "Salmon, basmati rice, dill, parsley, cilantro, green onions, saffron, vegetable oil, lemon, salt, and spices.",
    "allergens": "Fish (salmon)",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/EOFD2LVRHYJABTEZIZSLAYVP.jpeg",
    "url": "https://nomzkitchen.square.site/product/sabzi-polo-with-salmon-/HD5YIFAMRRL7DJP2O7547KGQ"
  },
  {
    "id": "FRZLVLUYCOIZHM3LUPVC6UHL",
    "sort": 5,
    "featured": false,
    "available": true,
    "nameEn": "Adas Polo",
    "nameFa": "عدس پلو",
    "price": 16.95,
    "category": "Rice dishes",
    "description": "Persian rice layered with tender lentils, seasoned ground beef, sweet raisins, and caramelized onions, finished with fragrant saffron.",
    "ingredients": "Basmati rice, lentils, ground beef, raisins, onions, saffron, turmeric, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/HJPKKFQD6HGXFGZYH4ID7AEC.jpeg",
    "url": "https://nomzkitchen.square.site/product/adas-polo-/FRZLVLUYCOIZHM3LUPVC6UHL"
  },
  {
    "id": "RMEUAF2QGK6RWU4MSIOQU2G7",
    "sort": 6,
    "featured": false,
    "available": true,
    "nameEn": "Loobia Polo",
    "nameFa": "لوبیا پلو",
    "price": 16.95,
    "category": "Rice dishes",
    "description": "Tender green beans, seasoned ground beef, tomatoes, and warm spices layered with fluffy basmati rice and gently steamed.",
    "ingredients": "Basmati rice, green beans, ground beef, tomatoes, tomato paste, onion, turmeric, cinnamon, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/IE4FOVLTLOLWPMVAGB2OSBTL.jpeg",
    "url": "https://nomzkitchen.square.site/product/loobia-polo-/RMEUAF2QGK6RWU4MSIOQU2G7"
  },
  {
    "id": "RKJ6HENZ7JIT3DIVQDYSXYTB",
    "sort": 7,
    "featured": true,
    "available": true,
    "nameEn": "Ghormeh Sabzi Stew",
    "nameFa": "خورشت قورمه‌سبزی",
    "price": 16.95,
    "category": "Stews",
    "description": "The classic Persian herb stew, slow-cooked with tender beef, kidney beans, dried limes, and a blend of fresh herbs. Served with basmati rice.",
    "ingredients": "Beef, parsley, cilantro, fenugreek, green onions or leeks, kidney beans, dried Persian limes, onion, turmeric, basmati rice, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/Y6MWLQOMYPEGPJVOKZ2FUVWR.jpeg",
    "url": "https://nomzkitchen.square.site/product/ghormeh-sabzi-stew-/RKJ6HENZ7JIT3DIVQDYSXYTB"
  },
  {
    "id": "FFKZRT2HXUXRH5J7JIUDLGGC",
    "sort": 8,
    "featured": true,
    "available": true,
    "nameEn": "Fesenjan Stew",
    "nameFa": "خورشت فسنجان",
    "price": 18.95,
    "category": "Stews",
    "description": "A rich stew of finely ground walnuts and pomegranate molasses, slow-cooked into a sweet and tangy sauce. Served with basmati rice.",
    "ingredients": "Chicken, ground walnuts, pomegranate molasses, onion, turmeric, basmati rice, salt, and spices.",
    "allergens": "Tree nuts (walnuts)",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/V3F3HK37EW5PGLIJYEBMNQPW.jpeg",
    "url": "https://nomzkitchen.square.site/product/fesenjan-stew-/FFKZRT2HXUXRH5J7JIUDLGGC"
  },
  {
    "id": "S52DNVRL5MLJ3W2POOZMQF3O",
    "sort": 9,
    "featured": false,
    "available": true,
    "nameEn": "Gheymeh Bademjan Stew",
    "nameFa": "خورشت قیمه بادمجان",
    "price": 16.95,
    "category": "Stews",
    "description": "Tender beef, yellow split peas, and golden fried eggplant simmered in a tomato sauce with dried Persian limes. Served with saffron basmati rice.",
    "ingredients": "Beef, eggplant, yellow split peas, basmati rice, onion, tomato paste, dried Persian limes, saffron, turmeric, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/UQ4D7DKIEXXAGWLW3Q6W5VCQ.jpeg",
    "url": "https://nomzkitchen.square.site/product/gheymeh-bademjan-stew-/S52DNVRL5MLJ3W2POOZMQF3O"
  },
  {
    "id": "3QDSVOM5YVBB6XCVVYO6NXB7",
    "sort": 10,
    "featured": false,
    "available": true,
    "nameEn": "Gheymeh Stew with Potatoes",
    "nameFa": "خورشت قیمه با سیب‌زمینی",
    "price": 16.95,
    "category": "Stews",
    "description": "Tender beef, yellow split peas, and dried Persian limes in a rich tomato sauce, topped with crispy golden fried potatoes. Served with saffron basmati rice.",
    "ingredients": "Beef, yellow split peas, potatoes, basmati rice, onion, tomato paste, dried Persian limes, saffron, turmeric, vegetable oil, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/2XXIDPEAVB7G7BMYCB2ZD5X2.jpeg",
    "url": "https://nomzkitchen.square.site/product/gheymeh-stew-with-potatoes-/3QDSVOM5YVBB6XCVVYO6NXB7"
  },
  {
    "id": "TMISBGCUDFQ57I5M45BJZ7QB",
    "sort": 11,
    "featured": false,
    "available": true,
    "nameEn": "Abgoosht",
    "nameFa": "آبگوشت",
    "price": 19.95,
    "category": "Traditional",
    "description": "Slow-cooked lamb stew with chickpeas, white beans, potatoes, and tomatoes, served with bread.",
    "ingredients": "Lamb, chickpeas, white beans, potatoes, tomatoes, onion, turmeric, dried lime, salt, and spices.",
    "allergens": "",
    "image": "https://156522947.cdn6.editmysite.com/uploads/1/5/6/5/156522947/RLLX3KARAI4ECIEODFIIOZQK.jpeg",
    "url": "https://nomzkitchen.square.site/product/abgoosht-/TMISBGCUDFQ57I5M45BJZ7QB"
  }
];
