/**
 * L'CHAIM — BACKEND (Google Apps Script + Google Sheets as DB)
 * Covers all phases: Fresh Food ordering (Phase 1), Daily Reflections /
 * Pillars / Stories (Phase 1-2 content), Coaching bookings (Phase 2-3).
 *
 * SETUP
 * 1. Create a Google Sheet with these tabs & header rows (row 1 exactly):
 *
 *    Menu:        ID | Tag | Name | Description | Price | Emoji | Phase | Active
 *    Reflections: ID | Quote | Active
 *    Pillars:     ID | Icon | Name | Description
 *    Stories:     ID | Quote | Who
 *    Orders:      OrderID | Timestamp | Name | Email | Phone | ItemsJSON | Total | Status | StripeRef
 *    Coaching:    BookingID | Timestamp | Name | Email | Pillar | Message | PreferredTime | Status
 *    ConciergePlans: RequestID | Timestamp | Name | Email | Phone | Dimensions | PlanJSON | Status
 *
 *    Note: Coaching's Pillar column now also carries the index.html requests —
 *    "Mind — Clarity & Stress Coaching", "Social — Table Circle RSVP",
 *    "Spirit — Purpose & Values Mentorship", "Spirit — Restoration Retreat Waitlist" —
 *    so you can still filter/report by pillar, just with more specific labels.
 *
 * 2. Extensions > Apps Script, paste this file in as code.gs.
 * 3. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the deployment URL into WEBAPP_URL in the frontend fetch calls.
 * 5. (Payments) Store your Stripe secret key in
 *      Project Settings > Script Properties as STRIPE_SECRET_KEY,
 *      then use createStripeCheckoutSession() from placeOrder().
 */

const SHEET_NAMES = {
  MENU: 'Menu',
  REFLECTIONS: 'Reflections',
  PILLARS: 'Pillars',
  STORIES: 'Stories',
  ORDERS: 'Orders',
  COACHING: 'Coaching',
  CONCIERGE: 'ConciergePlans'
};

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Reads a sheet into an array of objects using row 1 as keys.
function readSheet_(sheetName) {
  const sheet = ss_().getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(row => row.join('') !== '') // skip blank rows
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function appendRow_(sheetName, rowArray) {
  const sheet = ss_().getSheetByName(sheetName);
  sheet.appendRow(rowArray);
}

// ---------- GET: read-only data for the site ----------
function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  try {
    switch (action) {
      case 'getmenu': {
        const items = readSheet_(SHEET_NAMES.MENU).filter(r => r.Active === true || r.Active === 'TRUE');
        return json_({ ok: true, items });
      }
      case 'getreflections': {
        const items = readSheet_(SHEET_NAMES.REFLECTIONS).filter(r => r.Active === true || r.Active === 'TRUE');
        return json_({ ok: true, items });
      }
      case 'getpillars': {
        return json_({ ok: true, items: readSheet_(SHEET_NAMES.PILLARS) });
      }
      case 'getstories': {
        return json_({ ok: true, items: readSheet_(SHEET_NAMES.STORIES) });
      }
      default:
        return json_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// ---------- POST: writes (orders, coaching bookings) ----------
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Invalid JSON body' });
  }

  const action = (body.action || '').toLowerCase();

  try {
    switch (action) {
      case 'placeorder':
        return json_(placeOrder_(body));
      case 'bookcoaching':
        return json_(bookCoaching_(body));
      case 'requestconcierge':
        return json_(requestConciergePlan_(body));
      default:
        return json_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

// Phase 1 — food ordering
function placeOrder_(data) {
  const required = ['name', 'phone', 'items', 'total']; // email is optional
  for (const field of required) {
    if (!data[field]) return { ok: false, error: 'Missing field: ' + field };
  }

  const orderId = 'ORD-' + new Date().getTime();

  appendRow_(SHEET_NAMES.ORDERS, [
    orderId,
    new Date(),
    data.name,
    data.email || '',
    data.phone,
    JSON.stringify(data.items),
    data.total,
    'pending_payment',
    ''
  ]);

  const itemNames = data.items.map(i => i.name).join(', ');
  sendWhatsAppNotification_(
    `🥗 New L'chaim order ${orderId}\n` +
    `Name: ${data.name}\n` +
    `Phone: ${data.phone}\n` +
    `Items: ${itemNames}\n` +
    `Total: $${data.total}`
  );

  // Payment: create the Stripe Checkout session here and return its URL
  // for the frontend to redirect to. Order status flips to "paid" when
  // Stripe's webhook calls this same web app with action=confirmpayment
  // (add that case to doPost once Stripe is wired in).
  // const checkoutUrl = createStripeCheckoutSession_(orderId, data);

  return { ok: true, orderId /*, checkoutUrl */ };
}

// Phase 2/3 — mind / social / spiritual coaching booking requests
function bookCoaching_(data) {
  const required = ['name', 'email', 'pillar'];
  for (const field of required) {
    if (!data[field]) return { ok: false, error: 'Missing field: ' + field };
  }

  const bookingId = 'BK-' + new Date().getTime();

  appendRow_(SHEET_NAMES.COACHING, [
    bookingId,
    new Date(),
    data.name,
    data.email,
    data.pillar,
    data.message || '',
    data.preferredTime || '',
    'new'
  ]);

  sendWhatsAppNotification_(
    `✨ New ${data.pillar} session request ${bookingId}\n` +
    `Name: ${data.name}\n` +
    `Email: ${data.email}`
  );

  return { ok: true, bookingId };
}

// Phase 4 — "generate my plan" full assessment (all dimensions in one request)
function requestConciergePlan_(data) {
  const required = ['name', 'phone', 'dimensions'];
  for (const field of required) {
    if (!data[field]) return { ok: false, error: 'Missing field: ' + field };
  }
  if (!Array.isArray(data.dimensions) || data.dimensions.length === 0) {
    return { ok: false, error: 'dimensions must be a non-empty array' };
  }

  const requestId = 'PLAN-' + new Date().getTime();
  const plan = generatePlanContent_(data.dimensions);

  appendRow_(SHEET_NAMES.CONCIERGE, [
    requestId,
    new Date(),
    data.name,
    data.email || '',
    data.phone,
    data.dimensions.join(', '),
    JSON.stringify(plan),
    'new'
  ]);

  // Auto-enroll: log an actionable follow-up in Coaching for each non-physical
  // dimension so staff has a concrete task, not just a raw checkbox list.
  ['mental', 'social', 'spiritual'].forEach((dim) => {
    if (!plan[dim]) return;
    appendRow_(SHEET_NAMES.COACHING, [
      'BK-' + new Date().getTime() + '-' + dim,
      new Date(),
      data.name,
      data.email || '',
      plan[dim].title + ' (via Concierge Plan)',
      'Auto-enrolled from concierge request ' + requestId,
      '',
      'new'
    ]);
  });

  // Email the plan straight to the customer, if we have an address.
  if (data.email) {
    try {
      MailApp.sendEmail(data.email, "Your L'chaim Plan", buildPlanEmailBody_(data.name, plan));
    } catch (err) {
      console.error('Plan email failed: ' + err.message);
    }
  }

  sendWhatsAppNotification_(
    `🌿 New concierge plan request ${requestId}\n` +
    `Name: ${data.name}\n` +
    `Phone: ${data.phone}\n` +
    `Dimensions: ${data.dimensions.join(', ')}`
  );

  return { ok: true, requestId, plan };
}

// Builds the actual plan content per selected dimension, pulling live data
// from the Menu and Reflections sheets rather than static copy.
function generatePlanContent_(dimensions) {
  const dims = dimensions.map(d => d.toLowerCase());
  const plan = {};

  if (dims.includes('physical')) {
    const menu = readSheet_(SHEET_NAMES.MENU).filter(r => r.Active === true || r.Active === 'TRUE');
    const picks = menu.slice(0, 3).map(r => r.Name);
    plan.physical = {
      title: 'Physical — Your starter meal plan',
      detail: picks.length
        ? 'Your first three meals: ' + picks.join(', ') + '.'
        : "Fresh picks from today's menu — see the Fresh Foods section to choose."
    };
  }

  if (dims.includes('mental')) {
    plan.mental = {
      title: 'Mental — Clarity & Stress Coaching',
      detail: "You're enrolled. We'll reach out within 1-2 days to schedule your first session."
    };
  }

  if (dims.includes('social')) {
    plan.social = {
      title: "Social — L'chaim Table Circles",
      detail: "You're on the list for the next Table Circle — watch your email for the date and topic."
    };
  }

  if (dims.includes('spiritual')) {
    const reflections = readSheet_(SHEET_NAMES.REFLECTIONS).filter(r => r.Active === true || r.Active === 'TRUE');
    const pick = reflections.length
      ? reflections[Math.floor(Math.random() * reflections.length)].Quote
      : 'Take a quiet moment today to reflect on what matters most.';
    plan.spiritual = {
      title: "Spiritual — Today's devotional",
      detail: pick
    };
  }

  return plan;
}

function buildPlanEmailBody_(name, plan) {
  const sections = Object.values(plan).map(p => `${p.title}\n${p.detail}`).join('\n\n');
  return `Hi ${name},\n\nHere is your personalized L'chaim plan:\n\n${sections}\n\nTo life,\nThe L'chaim Team`;
}

/**
 * Sends a WhatsApp message to the business owner via CallMeBot
 * (https://www.callmebot.com/blog/free-api-whatsapp-messages/) —
 * a free service for sending yourself WhatsApp notifications from code.
 * Not meant for messaging customers or high volume — for that, use the
 * official Twilio or Meta WhatsApp Business API instead.
 *
 * ONE-TIME SETUP:
 * 1. Save +34 644 59 71 67 as a contact on the phone you want notified.
 * 2. From that WhatsApp, send: "I allow callmebot to send me messages"
 *    to that number.
 * 3. You'll receive an API key by reply within a minute or two.
 * 4. In Apps Script: Project Settings > Script Properties, add:
 *      WHATSAPP_PHONE      = your number with country code, no + or spaces
 *                             (e.g. 15551234567)
 *      CALLMEBOT_API_KEY   = the key you received
 * 5. Re-deploy (Deploy > Manage deployments > Edit > New version).
 *
 * If either property is missing, this silently skips (logged only) —
 * it never blocks an order or booking from being recorded.
 */
function sendWhatsAppNotification_(message) {
  const props = PropertiesService.getScriptProperties();
  const phone = props.getProperty('WHATSAPP_PHONE');
  const apiKey = props.getProperty('CALLMEBOT_API_KEY');

  if (!phone || !apiKey) {
    console.warn('WhatsApp notification skipped — set WHATSAPP_PHONE and CALLMEBOT_API_KEY in Script Properties.');
    return;
  }

  const url = 'https://api.callmebot.com/whatsapp.php'
    + '?phone=' + encodeURIComponent(phone)
    + '&text=' + encodeURIComponent(message)
    + '&apikey=' + encodeURIComponent(apiKey);

  try {
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (err) {
    console.error('WhatsApp notification failed: ' + err.message);
  }
}

