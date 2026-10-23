/**
 * Jesus Life Church - Giving records (Google Apps Script)
 *
 * SETUP (once):
 *  1. Create a new Google Sheet (name it e.g. "JLC Giving").
 *  2. In the sheet: Extensions > Apps Script. Delete any code, paste this whole file, Save.
 *  3. Choose the function "setup" in the toolbar and click Run. Approve the permissions.
 *  4. Deploy > New deployment > type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy and copy the Web app URL.
 *  5. Paste that URL into CONFIG.sheetUrl in give.html and receipt.html.
 *
 * If you ever change this script: Deploy > Manage deployments > edit (pencil) >
 * Version: New version > Deploy. The URL stays the same.
 *
 * DAILY USE:
 *  Each gift appears as a new row with Status = Pending.
 *  Check your EcoCash, then set Status to "Confirmed" (dropdown). The giver can then
 *  open "Check my receipt" and see their confirmed receipt.
 *  Keep this sheet private: it contains givers' phone numbers.
 */

const SHEET_NAME = 'Giving';
const TZ = 'Africa/Harare';

const HEADERS = [
  'Receipt No', 'Date', 'Name', 'Private?', 'Phone', 'Email', 'Paid From',
  'Giving Toward', 'Pledge / Message', 'Amount', 'Currency', 'Frequency',
  'Method', 'Payment', 'Reference', 'Status', 'Confirmed At'
];

// Column numbers (1-based)
const COL = {
  REF: 1, DATE: 2, NAME: 3, ANON: 4, PHONE: 5, EMAIL: 6, PAYER: 7, TYPE: 8,
  PLEDGE: 9, AMOUNT: 10, CURRENCY: 11, FREQ: 12, METHOD: 13, PAYMENT: 14,
  REFERENCE: 15, STATUS: 16, CONFIRMED_AT: 17
};

/** Run this once to prepare the sheet. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const rows = sh.getMaxRows() - 1;

  sh.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#5b21b6')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);

  // Everything is plain text except the amount. This keeps phone numbers intact
  // and stops anything typed by a giver from being treated as a formula.
  for (let c = 1; c <= HEADERS.length; c++) {
    if (c === COL.AMOUNT) continue;
    sh.getRange(2, c, rows, 1).setNumberFormat('@');
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Confirmed', 'Rejected'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, COL.STATUS, rows, 1).setDataValidation(rule);

  const statusRange = sh.getRange(2, COL.STATUS, rows, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Confirmed')
      .setBackground('#bbf7d0').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending')
      .setBackground('#fef08a').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Rejected')
      .setBackground('#fecaca').setRanges([statusRange]).build()
  ]);

  sh.setColumnWidths(1, HEADERS.length, 130);
  sh.setColumnWidth(COL.PLEDGE, 260);
}

/** Stamps the confirmation time when you change Status in the sheet. */
function onEdit(e) {
  if (!e || !e.range) return;
  const r = e.range;
  const sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME) return;
  if (r.getColumn() > COL.STATUS || r.getLastColumn() < COL.STATUS) return;

  for (let row = r.getRow(); row <= r.getLastRow(); row++) {
    if (row < 2) continue;
    const status = sh.getRange(row, COL.STATUS).getValue();
    const stamp = sh.getRange(row, COL.CONFIRMED_AT);
    if (status === 'Confirmed') {
      if (!stamp.getValue()) stamp.setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
    } else {
      stamp.setValue('');
    }
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'content') {
    return ContentService.createTextOutput(contentJson_())
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Jesus Life Church giving service is running.');
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'submit') out = submit_(req);
    else if (req.action === 'lookup') out = lookup_(req);
    else if (['visitor', 'prayer', 'testimony'].indexOf(req.action) >= 0) out = form_(req);
    else if (req.action === 'order') out = order_(req);
    else if (req.action === 'booking') out = booking_(req);
    else if (req.action === 'anon') out = anon_(req);
    else out = { ok: false, error: 'Unknown request.' };
  } catch (err) {
    out = { ok: false, error: 'Bad request.' };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

function digits_(v) {
  return String(v == null ? '' : v).replace(/\D/g, '');
}

function newRef_(sh) {
  const day = Utilities.formatDate(new Date(), TZ, 'yyMMdd');
  for (let i = 0; i < 5; i++) {
    const rand = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
    const ref = 'JLC-' + day + '-' + rand;
    const last = sh.getLastRow();
    if (last < 2) return ref;
    const hit = sh.getRange(2, COL.REF, last - 1, 1)
      .createTextFinder(ref).matchEntireCell(true).findNext();
    if (!hit) return ref;
  }
  throw new Error('Could not create a unique receipt number');
}

/* ---------- submit a gift ---------- */

function submit_(req) {
  if (req.website) return { ok: false, error: 'Rejected.' }; // honeypot for bots

  const r = req.record || {};
  const amount = Number(r.amount);
  if (!(amount > 0) || amount > 1000000) return { ok: false, error: 'Invalid amount.' };
  if (digits_(r.phone).length < 9) return { ok: false, error: 'Invalid phone number.' };

  const currency = ['USD', 'ZiG'].indexOf(r.currency) >= 0 ? r.currency : 'USD';
  const freq = r.freq === 'Monthly' ? 'Monthly' : 'One-time';
  const paid = r.paid === true;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    if (!sh) return { ok: false, error: 'Sheet not set up. Run setup().' };
    const ref = newRef_(sh);
    const row = [];
    row[COL.REF - 1] = ref;
    row[COL.DATE - 1] = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    row[COL.NAME - 1] = clean_(r.name, 80);
    row[COL.ANON - 1] = r.anon === true ? 'Yes' : 'No';
    row[COL.PHONE - 1] = clean_(r.phone, 20);
    row[COL.EMAIL - 1] = clean_(r.email, 80);
    row[COL.PAYER - 1] = clean_(r.payer, 20);
    row[COL.TYPE - 1] = clean_(r.type, 40);
    row[COL.PLEDGE - 1] = clean_(r.pledge, 300);
    row[COL.AMOUNT - 1] = amount;
    row[COL.CURRENCY - 1] = currency;
    row[COL.FREQ - 1] = freq;
    row[COL.METHOD - 1] = clean_(r.method, 30);
    row[COL.PAYMENT - 1] = paid ? 'Reported paid' : 'Pledge (pay later)';
    row[COL.REFERENCE - 1] = clean_(r.reference, 40);
    row[COL.STATUS - 1] = 'Pending';       // always Pending: only you can confirm
    row[COL.CONFIRMED_AT - 1] = '';
    sh.appendRow(row);
    return { ok: true, ref: ref };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- look up a receipt ---------- */

function lookup_(req) {
  const ref = clean_(req.ref, 40).toUpperCase();
  const d = digits_(req.phone).slice(-9);
  if (!ref || d.length < 9) {
    return { ok: false, error: 'Please enter your receipt number and phone number.' };
  }

  // Slow down guessing: 10 wrong tries per phone number per 10 minutes
  const cache = CacheService.getScriptCache();
  const key = 'fail_' + d;
  const fails = Number(cache.get(key) || 0);
  if (fails >= 10) {
    return { ok: false, error: 'Too many attempts. Please wait 10 minutes or contact the church.' };
  }

  const sh = sheet_();
  const last = sh ? sh.getLastRow() : 0;
  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[COL.REF - 1]).toUpperCase() !== ref) continue;
      const matches = [row[COL.PHONE - 1], row[COL.PAYER - 1]]
        .some(function (v) { return digits_(v).slice(-9) === d; });
      if (!matches) break;
      return {
        ok: true,
        found: true,
        receipt: {
          ref: String(row[COL.REF - 1]),
          date: String(row[COL.DATE - 1]),
          name: String(row[COL.NAME - 1]),
          anon: row[COL.ANON - 1] === 'Yes',
          phone: String(row[COL.PHONE - 1]),
          payer: String(row[COL.PAYER - 1]),
          type: String(row[COL.TYPE - 1]),
          pledge: String(row[COL.PLEDGE - 1]),
          amount: Number(row[COL.AMOUNT - 1]),
          currency: String(row[COL.CURRENCY - 1]),
          freq: String(row[COL.FREQ - 1]),
          method: String(row[COL.METHOD - 1]),
          payment: String(row[COL.PAYMENT - 1]),
          reference: String(row[COL.REFERENCE - 1]),
          status: String(row[COL.STATUS - 1]) || 'Pending',
          confirmedAt: String(row[COL.CONFIRMED_AT - 1])
        }
      };
    }
  }

  cache.put(key, String(fails + 1), 600);
  // Same answer whether the receipt number or the phone was wrong
  return { ok: true, found: false };
}

/* =====================================================================
 * WEBSITE CONTENT, FORMS AND SETTINGS
 *
 * Run setupContent() once (and again whenever you replace this script).
 * It is safe to run many times: it never deletes your rows.
 *
 * Tabs it creates, and what they do:
 *   Events        what's coming (dates, flyers)
 *   Leaders       the sliding leadership photos
 *   Sermons       media archive links
 *   Devotionals   one row per day
 *   Ministries    youth, women, men, children, etc.
 *   Gallery       photo gallery
 *   Testimonies   people send theirs here; set Approved = Yes to show it
 *   Visitors      the "I'm new" form lands here
 *   Prayer Requests   the prayer form lands here
 *   Settings      announcement bar and "Live now" switch
 *
 * After pasting this new code: Deploy > Manage deployments > pencil >
 * Version: New version > Deploy.  The web app URL stays the same.
 * ===================================================================== */

const TAB_DEFS = {
  Events:      { key: 'events',      main: 'title',       headers: ['Date', 'End Date', 'Time', 'Title', 'Description', 'Location', 'Show', 'Image'] },
  Leaders:     { key: 'leaders',     main: 'name',        headers: ['Group', 'Name', 'Role', 'Photo', 'Bio', 'Order', 'Show'] },
  Sermons:     { key: 'sermons',     main: 'title',       headers: ['Date', 'Title', 'Speaker', 'Type', 'Link', 'Series', 'Show'] },
  Devotionals: { key: 'devotionals', main: 'title',       headers: ['Date', 'Title', 'Verse', 'Message', 'Prayer'] },
  Ministries:  { key: 'ministries',  main: 'name',        headers: ['Name', 'Description', 'Leader', 'Meets', 'Order', 'Show'] },
  Gallery:     { key: 'gallery',     main: 'photo',       headers: ['Photo', 'Caption', 'Date', 'Show'] },
  Testimonies: { key: 'testimonies', main: 'testimony',   headers: ['Date', 'Name', 'Category', 'Testimony', 'Phone', 'Approved', 'Real Name'], approval: true, hide: ['phone', 'approved', 'realname'] },
  Visitors:    { key: null, headers: ['Date', 'Name', 'Phone', 'Email', 'Plan', 'Message', 'Status'], statusCol: 7, statuses: ['New', 'Contacted', 'Done'] },
  'Prayer Requests': { key: null, headers: ['Date', 'Name', 'Phone', 'Request', 'Private?', 'Status'], statusCol: 6, statuses: ['New', 'Praying', 'Done'] },
  Products:    { key: 'products', main: 'name', headers: ['Name', 'Description', 'Price', 'Currency', 'Options', 'Image', 'Category', 'Available', 'Show'] },
  Orders:      { key: null, headers: ['Date', 'Order No', 'Name', 'Phone', 'Items', 'Total', 'Delivery', 'Notes', 'Status'], statusCol: 9, statuses: ['New', 'Confirmed', 'Ready', 'Collected', 'Cancelled'] },
  Appointments: { key: null, headers: ['Date Requested', 'Ref', 'Name', 'Phone', 'Email', 'Topic', 'Preferred Date', 'Preferred Time', 'Meeting Type', 'Note', 'Status'], statusCol: 11, statuses: ['Requested', 'Confirmed', 'Declined', 'Done'] },
  'Anonymous Messages': { key: null, headers: ['Date', 'Category', 'Message', 'Contact (optional)', 'Status'], statusCol: 5, statuses: ['New', 'Read', 'Actioned'] }
};

const NOTES = {
  Events: 'Date format: 2026-11-08 (year-month-day). End Date only for events lasting several days. Image = flyer file name in the "photos" folder (e.g. anniversary.jpg) or a full https link. Show = No hides a row.',
  Leaders: 'Group examples: Senior Leadership, Pastors, Admin, Praise & Worship, Ushers. Photo = file name inside the "photos" folder, or a full https link. Order = 1, 2, 3 (smaller shows first).',
  Sermons: 'Type: Video, Audio or Notes. Link = YouTube, Facebook, Google Drive or any https link. Newest date shows first.',
  Devotionals: 'One row per day. Date format: 2026-10-23. The website shows today\u2019s row, or the latest one before today.',
  Ministries: 'One row per ministry (Youth, Women, Men, Children...). Meets = e.g. "Saturdays 2 PM". Order = 1, 2, 3.',
  Gallery: 'Photo = file name inside the "photos" folder, or a full https link. Caption and Date are optional.',
  Testimonies: 'People send testimonies from the website. They start as Pending. Change Approved to Yes to show one on the site. "Name" is the name to display; "Real Name" is private.',
  Visitors: 'New-visitor forms. Change Status to Contacted / Done as you follow up. Keep this sheet private.',
  'Prayer Requests': 'Prayer requests. Change Status as you pray and follow up. Keep this sheet private.',
  Products: 'One row per item. Price = a number only (e.g. 12.50), Currency = USD or ZiG. Options = choices separated by commas (e.g. S,M,L,XL or Black,White). Image = file name in the "photos" folder or a full https link. Available = No shows "Sold out". Show = Yes to put it in the store.',
  Orders: 'Orders from the website store. Totals are worked out by the server from the Products tab. Change Status as you confirm, prepare and hand over. Keep this sheet private.',
  Appointments: 'One-on-one requests. Nothing is confirmed until you contact the person. Change Status to Confirmed / Declined / Done. Keep this sheet private.',
  'Anonymous Messages': 'Messages sent without a name. Do not try to find out who sent them. Contact is filled only if the person chose to leave it. Keep this sheet private.'
};

/** Run this to create (or refresh) all website tabs. Safe to run again. */
function setupContent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(TAB_DEFS).forEach(function (name) {
    const def = TAB_DEFS[name];
    const headers = def.headers;
    let sh = ss.getSheetByName(name);
    const isNew = !sh;
    if (!sh) sh = ss.insertSheet(name);

    // clear the old "how to use" note area (right of the table) so it can be rewritten
    sh.getRange(1, headers.length + 1, 1, 6).clearContent();
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#5b21b6').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, headers.length).setNumberFormat('@').setWrap(true).setVerticalAlignment('top');
    sh.setColumnWidths(1, headers.length, 140);

    const showIdx = headers.indexOf('Show');
    if (showIdx >= 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(true).build();
      sh.getRange(2, showIdx + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
    }
    const availIdx = headers.indexOf('Available');
    if (availIdx >= 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(true).build();
      sh.getRange(2, availIdx + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
    }
    if (def.statuses) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(def.statuses, true).setAllowInvalid(true).build();
      sh.getRange(2, def.statusCol, sh.getMaxRows() - 1, 1).setDataValidation(rule);
    }
    if (def.approval) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Pending', 'Yes', 'No'], true).setAllowInvalid(true).build();
      sh.getRange(2, headers.indexOf('Approved') + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
      sh.setColumnWidth(headers.indexOf('Testimony') + 1, 360);
    }

    if (isNew) {
      if (name === 'Events') {
        sh.getRange(2, 1, 2, headers.length).setValues([
          ['2026-10-23', '', 'All night', 'Night of Visitation 2', 'An all-night service. Come expecting.', '', 'Yes', ''],
          ['2026-11-08', '', '', '5th Church Anniversary & Prophet\u2019s Birthday Celebration', 'Join us as we celebrate five years of God\u2019s faithfulness.', '', 'Yes', '']
        ]);
      }
      if (name === 'Products') {
        sh.getRange(2, 1, 5, headers.length).setValues([
          ['Anointing Oil', 'Made to order. Ask about sizes.', '', 'USD', '', '', 'Oils', 'Yes', 'No'],
          ['Church Calendar', 'Made to order.', '', 'USD', '', '', 'Print', 'Yes', 'No'],
          ['JLC Stickers', 'Made to order.', '', 'USD', '', '', 'Print', 'Yes', 'No'],
          ['JLC T-Shirt', 'With the JLC logo. Made to order.', '', 'USD', 'S,M,L,XL', '', 'Clothing', 'Yes', 'No'],
          ['JLC Cap', 'With the JLC logo. Made to order.', '', 'USD', '', '', 'Clothing', 'Yes', 'No']
        ]);
      }
      if (name === 'Leaders') {
        sh.getRange(2, 1, 1, headers.length).setValues([
          ['Senior Leadership', 'Prophets M & B Chiyangwa', 'Senior Leaders', '', '', '1', 'Yes']
        ]);
      }
    }
    if (NOTES[name]) {
      sh.getRange(1, headers.length + 2).setValue('HOW TO USE: ' + NOTES[name]).setFontStyle('italic').setWrap(false);
    }
  });

  // Settings tab (announcement bar + live switch)
  let st = ss.getSheetByName('Settings');
  if (!st) {
    st = ss.insertSheet('Settings');
    st.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']])
      .setFontWeight('bold').setBackground('#5b21b6').setFontColor('#ffffff');
    st.getRange(2, 1, 4, 2).setValues([
      ['Announcement text', ''],
      ['Announcement link', ''],
      ['Live now (Yes/No)', 'No'],
      ['Live link', '']
    ]);
    st.setFrozenRows(1);
    st.getRange(1, 1, 10, 2).setNumberFormat('@');
    st.setColumnWidth(1, 200);
    st.setColumnWidth(2, 460);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build();
    st.getRange(4, 2).setDataValidation(rule);
    st.getRange(1, 4).setValue('HOW TO USE: Type a message in "Announcement text" to show a bar at the top of every page (add a link if you want). Set "Live now" to Yes and paste your YouTube/Facebook live link to show a red LIVE NOW bar. Leave the text empty to hide the bar.')
      .setFontStyle('italic').setWrap(false);
  }
  ensureSettings_();
}

function ensureSettings_() {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!st) return;
  const wanted = [
    ['Store note', 'Everything is made to order. Place your request and we will confirm availability and payment details with you on WhatsApp.'],
    ['Booking note', 'Please tell us when you would like to meet. We will contact you to confirm the day and time.']
  ];
  const have = st.getLastRow() >= 2 ? st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
  wanted.forEach(function (w) {
    if (have.indexOf(w[0].toLowerCase()) < 0) {
      const row = st.getLastRow() + 1;
      st.getRange(row, 1, 1, 2).setNumberFormat('@').setValues([w]);
    }
  });
}

function contentJson_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('content_v3');
  if (hit) return hit;
  const json = JSON.stringify(content_());
  try { cache.put('content_v3', json, 60); } catch (err) { /* too large to cache, fine */ }
  return json;
}

function content_() {
  const out = { ok: true, settings: {} };
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(TAB_DEFS).forEach(function (name) {
    const def = TAB_DEFS[name];
    if (!def.key) return; // private tabs are never sent to the website
    const headers = def.headers;
    const sh = ss.getSheetByName(name);
    const rows = [];
    if (sh && sh.getLastRow() >= 2) {
      const values = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
      values.forEach(function (row) {
        const obj = {};
        headers.forEach(function (h, i) {
          let v = row[i];
          if (v instanceof Date) v = Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
          obj[h.replace(/[^A-Za-z]/g, '').toLowerCase()] = String(v == null ? '' : v).trim();
        });
        if (obj.show && obj.show.toLowerCase() === 'no') return;
        if (def.approval && String(obj.approved).toLowerCase() !== 'yes') return;
        if (!obj[def.main]) return;
        (def.hide || []).forEach(function (k) { delete obj[k]; });
        rows.push(obj);
      });
    }
    if (def.key === 'leaders' || def.key === 'ministries') {
      rows.sort(function (a, b) { return (Number(a.order) || 999) - (Number(b.order) || 999); });
    }
    if (def.key === 'testimonies') {
      rows.sort(function (a, b) { return String(b.date).localeCompare(String(a.date)); });
    }
    out[def.key] = rows;
  });

  const st = ss.getSheetByName('Settings');
  if (st && st.getLastRow() >= 2) {
    st.getRange(2, 1, st.getLastRow() - 1, 2).getValues().forEach(function (r) {
      const k = String(r[0]).trim().toLowerCase();
      const v = String(r[1] == null ? '' : r[1]).trim();
      if (k === 'announcement text') out.settings.announcement = v;
      if (k === 'announcement link') out.settings.announcementlink = v;
      if (k.indexOf('live now') === 0) out.settings.live = v.toLowerCase() === 'yes' ? 'yes' : 'no';
      if (k === 'live link') out.settings.livelink = v;
      if (k === 'store note') out.settings.storenote = v;
      if (k === 'booking note') out.settings.bookingnote = v;
    });
  }
  return out;
}

/* ---------- forms from the website (visitors, prayer, testimonies) ---------- */

function form_(req) {
  if (req.website) return { ok: false, error: 'Rejected.' }; // honeypot for bots
  const r = req.record || {};
  const names = { visitor: 'Visitors', prayer: 'Prayer Requests', testimony: 'Testimonies' };
  const sheetName = names[req.action];

  // slow down floods: at most 40 of each kind per hour
  const cache = CacheService.getScriptCache();
  const rk = 'rate_' + req.action;
  const n = Number(cache.get(rk) || 0);
  if (n >= 40) return { ok: false, error: 'We are receiving many messages right now. Please try again later or use WhatsApp.' };
  cache.put(rk, String(n + 1), 3600);

  const when = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  let row;
  if (req.action === 'visitor') {
    if (!clean_(r.name, 80) || digits_(r.phone).length < 9) return { ok: false, error: 'Please enter your name and phone number.' };
    row = [when, clean_(r.name, 80), clean_(r.phone, 20), clean_(r.email, 80), clean_(r.plan, 60), clean_(r.message, 600), 'New'];
  } else if (req.action === 'prayer') {
    if (!clean_(r.request, 800)) return { ok: false, error: 'Please write your prayer request.' };
    row = [when, clean_(r.name, 80), clean_(r.phone, 20), clean_(r.request, 800), r.isPrivate === true ? 'Yes' : 'No', 'New'];
  } else {
    if (clean_(r.testimony, 1500).length < 20) return { ok: false, error: 'Please write a little more of your testimony.' };
    if (r.consent !== true) return { ok: false, error: 'Please tick the box to allow us to share your testimony.' };
    row = [when, clean_(r.displayName, 80) || 'Anonymous', clean_(r.category, 40), clean_(r.testimony, 1500), clean_(r.phone, 20), 'Pending', clean_(r.realName, 80)];
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh) return { ok: false, error: 'Not set up yet. Run setupContent().' };
    sh.appendRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}


/* ---------- store orders, one-on-one bookings, anonymous messages ---------- */

function throttle_(kind, max) {
  const cache = CacheService.getScriptCache();
  const k = 'rate_' + kind;
  const n = Number(cache.get(k) || 0);
  if (n >= max) return false;
  cache.put(k, String(n + 1), 3600);
  return true;
}

function appendTo_(sheetName, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sh) return false;
    sh.appendRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function ref_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
}

function order_(req) {
  if (req.website) return { ok: false, error: 'Rejected.' };
  if (!throttle_('order', 60)) return { ok: false, error: 'We are receiving many orders right now. Please try again later or use WhatsApp.' };
  const r = req.record || {};
  if (!clean_(r.name, 80) || digits_(r.phone).length < 9) return { ok: false, error: 'Please enter your name and phone number.' };
  const items = Array.isArray(r.items) ? r.items.slice(0, 30) : [];
  if (!items.length) return { ok: false, error: 'Your basket is empty.' };

  // Prices always come from the Products tab, never from the visitor's browser
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Products');
  const catalog = {};
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, TAB_DEFS.Products.headers.length).getValues().forEach(function (row) {
      const name = String(row[0]).trim();
      if (!name) return;
      catalog[name.toLowerCase()] = {
        name: name,
        price: Number(String(row[2]).replace(/[^0-9.]/g, '')) || 0,
        currency: String(row[3]).trim() || 'USD',
        available: String(row[7]).trim().toLowerCase() !== 'no',
        shown: String(row[8]).trim().toLowerCase() !== 'no'
      };
    });
  }

  const lines = [];
  const totals = {};
  let unpriced = false;
  for (let i = 0; i < items.length; i++) {
    const it = items[i] || {};
    const p = catalog[String(it.name || '').trim().toLowerCase()];
    if (!p || !p.shown) return { ok: false, error: 'Sorry, one of the items is no longer in the store. Please refresh the page.' };
    if (!p.available) return { ok: false, error: 'Sorry, ' + p.name + ' is sold out.' };
    const qty = Math.max(1, Math.min(50, parseInt(it.qty, 10) || 1));
    const opt = clean_(it.option, 40);
    lines.push(qty + ' x ' + p.name + (opt ? ' (' + opt + ')' : ''));
    if (p.price > 0) totals[p.currency] = (totals[p.currency] || 0) + p.price * qty; else unpriced = true;
  }
  let total = Object.keys(totals).map(function (c) { return c + ' ' + totals[c].toFixed(2); }).join(' + ');
  if (!total) total = 'To be confirmed';
  else if (unpriced) total += ' (+ items priced on request)';

  const ref = ref_('ORD');
  const ok = appendTo_('Orders', [
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), ref, clean_(r.name, 80), clean_(r.phone, 20),
    lines.join('; '), total, clean_(r.delivery, 40), clean_(r.notes, 400), 'New'
  ]);
  if (!ok) return { ok: false, error: 'Not set up yet. Run setupContent().' };
  return { ok: true, ref: ref, total: total, items: lines.join('; ') };
}

function booking_(req) {
  if (req.website) return { ok: false, error: 'Rejected.' };
  if (!throttle_('booking', 40)) return { ok: false, error: 'We are receiving many requests right now. Please try again later or use WhatsApp.' };
  const r = req.record || {};
  if (!clean_(r.name, 80) || digits_(r.phone).length < 9) return { ok: false, error: 'Please enter your name and phone number.' };
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(r.date || '')) ? r.date : '';
  const ref = ref_('APT');
  const ok = appendTo_('Appointments', [
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), ref, clean_(r.name, 80), clean_(r.phone, 20), clean_(r.email, 80),
    clean_(r.topic, 60), date, clean_(r.time, 40), clean_(r.mode, 40), clean_(r.note, 500), 'Requested'
  ]);
  if (!ok) return { ok: false, error: 'Not set up yet. Run setupContent().' };
  return { ok: true, ref: ref };
}

function anon_(req) {
  if (req.website) return { ok: false, error: 'Rejected.' };
  if (!throttle_('anon', 60)) return { ok: false, error: 'Please try again a little later.' };
  const r = req.record || {};
  const msg = clean_(r.message, 2000);
  if (msg.length < 10) return { ok: false, error: 'Please write a little more.' };
  // Only the date, category, message and (if the person chose) a contact are stored. Nothing else.
  const contact = r.wantReply === true ? clean_(r.contact, 60) : '';
  const ok = appendTo_('Anonymous Messages', [
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), clean_(r.category, 40), msg, contact, 'New'
  ]);
  if (!ok) return { ok: false, error: 'Not set up yet. Run setupContent().' };
  return { ok: true };
}
