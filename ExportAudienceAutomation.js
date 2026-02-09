/******************************************************
 * GLOBAL CONFIG
 ******************************************************/
const CONFIG = {
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  MIN_PHONE_LENGTH: 8,
  BATCH_SIZE: 75,
  COUNTRY_CODE: 'ID',
  LANGUAGE_CODE: 'id'
};

/******************************************************
 * MENU SYSTEM
 ******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🧹 Utilities')
    .addItem('📞 Normalize Phone to E.164 Indonesia', 'normalizePhonesToE164')
    .addToUi();

  ui.createMenu('⬆️ Export To CRM/MA/BSP')
    .addSubMenu(ui.createMenu('Braze')
      .addItem('🔌 Check Upload API Call', 'checkBrazeConnection')
      .addSeparator()
      .addItem('📤 Upload to Braze', 'uploadToBraze')
      .addSeparator()
      .addItem('📊 View Stats', 'viewBrazeStats')
      .addItem('🔄 Reset Status', 'resetBrazeUploadStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚠️MoEngage⚠️')
      .addItem('🔌 Check Upload API Call ⚠️', 'checkMoEngageConnection')
      .addSeparator()
      .addItem('📤 Upload to MoEngage ⚠️', 'uploadToMoEngage')
      .addSeparator()
      .addItem('📊 View Stats', 'viewMoEngageStats')
      .addItem('🔄 Reset Status', 'resetMoEngageUploadStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Infobip')
      .addItem('🔌 Check Upload API Call', 'checkInfobipConnection')
      .addSeparator()
      .addItem('📤 Upload to Infobip', 'sendToInfobip')
      .addSeparator()
      .addItem('📊 View Stats', 'viewInfobipStats')
      .addItem('👥 List Recent 3', 'listRecentInfobipPeople')
      .addItem('🔄 Reset Status', 'resetInfobipUploadStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚠️Bitbybit⚠️')
      .addItem('🔌 Check Upload API Call ⚠️', 'checkBitbybitConnection')    
      .addSeparator()
      .addItem('📤 Upload to Bitbybit ⚠️', 'uploadToBitbybit')
      .addSeparator()
      .addItem('📊 View Stats', 'viewBitbybitStats')
      .addItem('🔄 Reset Status', 'resetBitbybitUploadStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚠️Gupshup⚠️')
      .addItem('🔌 Check Upload API Call ⚠️', 'checkGupshupConnection')      
      .addSeparator()
      .addItem('📤 Upload to Gupshup ⚠️', 'uploadToGupshup')
      .addSeparator()
      .addItem('📊 View Stats', 'viewGupshupStats')
      .addItem('🔄 Reset Status', 'resetGupshupUploadStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚠️Mekari Qontak⚠️')
      .addItem('🔌 Check Upload API Call ⚠️', 'checkQontakConnection')      
      .addSeparator()
      .addItem('📤 Upload to Qontak ⚠️', 'uploadToQontak')
      .addSeparator()
      .addItem('📊 View Stats', 'viewQontakStats')
      .addItem('🔄 Reset Status', 'resetQontakUploadStatus'))
    .addToUi();
    
  ui.createMenu('⬆️ Export To CDP/Analytics')
    .addSubMenu(ui.createMenu('Mixpanel')
      .addItem('🔌 Check Upload API Call (send User Profile)', 'checkMixpanelConnection')
      .addSeparator()
      .addItem('📤 Upload to Mixpanel', 'uploadToMixpanel')
      .addSeparator()
      .addItem('📊 View Stats', 'viewMixpanelStats')
      .addItem('👥 List Recent 3', 'listRecentMixpanelUsers')
      .addItem('🔄 Reset Status', 'resetMixpanelUploadStatus'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Segment')
      .addItem('🔌 Check Upload API Call', 'checkSegmentConnection')
      .addSeparator()
      .addItem('📤 Upload to Segment', 'uploadToSegment')
      .addSeparator()
      .addItem('📊 View Stats', 'viewSegmentStats')
      .addItem('🔄 Reset Status', 'resetSegmentUploadStatus'))
    .addToUi();
}

/******************************************************
 * BRAZE UPLOAD
 ******************************************************/
function uploadToBraze() {
  const sheet = SpreadsheetApp.getActiveSheet();
  setupGenericColumns(sheet, 'Braze Upload Status', 'Braze Upload Date');
  const ctx = getSheetContext(sheet, 'Braze Upload Status', 'Braze Upload Date');
  const items = [];
  const rows = [];

  ctx.data.forEach((r, i) => {
    if (r[ctx.statusCol] === 'Uploaded') return;
    const email = cleanEmail(r[ctx.emailCol]), phone = normalizeIndonesianPhone(r[ctx.phoneCol]), name = formatName(r[ctx.nameCol]);
    if (!email || !phone) return;
    items.push({ external_id: email, email, phone, first_name: name, country: CONFIG.COUNTRY_CODE, language: CONFIG.LANGUAGE_CODE });
    rows.push(i + 2);
  });

  if (handleEmptyOrConfirm(items.length, 'Braze')) return;
  for (let i = 0; i < items.length; i += CONFIG.BATCH_SIZE) {
    const p = PropertiesService.getScriptProperties();
    const res = UrlFetchApp.fetch(`${p.getProperty('BRAZE_REST_ENDPOINT')}/users/track`, {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: `Bearer ${p.getProperty('BRAZE_API_KEY')}` },
      payload: JSON.stringify({ attributes: items.slice(i, i + CONFIG.BATCH_SIZE) }), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 201) throw new Error(res.getContentText());
  }
  updateStatusOnSheet(sheet, rows, ctx.statusCol, ctx.dateCol, 'Braze');
}

/******************************************************
 * MOENGAGE UPLOAD
 ******************************************************/
function uploadToMoEngage() {
  const sheet = SpreadsheetApp.getActiveSheet();
  setupGenericColumns(sheet, 'MoEngage Upload Status', 'MoEngage Upload Date');
  const ctx = getSheetContext(sheet, 'MoEngage Upload Status', 'MoEngage Upload Date');
  const elements = [];
  const rows = [];

  ctx.data.forEach((r, i) => {
    if (r[ctx.statusCol] === 'Uploaded') return;
    const email = cleanEmail(r[ctx.emailCol]), phone = normalizeIndonesianPhone(r[ctx.phoneCol]), name = formatName(r[ctx.nameCol]);
    if (!email) return;
    elements.push({ type: "customer", attributes: { customer_id: email, full_name: name, email, mobile: phone, country: "ID" } });
    rows.push(i + 2);
  });

  if (handleEmptyOrConfirm(elements.length, 'MoEngage')) return;
  const p = PropertiesService.getScriptProperties();
  const res = UrlFetchApp.fetch(`https://api-${p.getProperty('MOENGAGE_REGION')}.moengage.com/v1/customer/add`, {
    method: 'post', contentType: 'application/json',
    headers: { "MOE-APP-ID": p.getProperty('MOE_APP_ID'), "X-Api-Key": p.getProperty('MOE_DATA_API_KEY') },
    payload: JSON.stringify({ elements }), muteHttpExceptions: true
  });
  updateStatusOnSheet(sheet, rows, ctx.statusCol, ctx.dateCol, 'MoEngage');
}

/******************************************************
 * INFOBIP UPLOAD
 ******************************************************/
function sendToInfobip() {
  console.log("Starting sendToInfobip with duplicate detection...");
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    setupGenericColumns(sheet, 'Infobip Upload Status', 'Infobip Upload Date');

    const p = PropertiesService.getScriptProperties();
    const API_KEY = p.getProperty('INFOBIP_API_KEY');
    const BASE_URL = p.getProperty('INFOBIP_BASE_URL');

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) throw new Error('No data rows.');

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const nameCol = headers.indexOf('Name');
    const emailCol = headers.indexOf('Email');
    const phoneCol = headers.indexOf('Phone Number');
    const statusCol = headers.indexOf('Infobip Upload Status');
    const dateCol = headers.indexOf('Infobip Upload Date');

    const queue = [];
    data.forEach((r, i) => {
      // Skip if already uploaded OR already identified as a duplicate
      if (r[statusCol] === 'Uploaded' || r[statusCol] === 'Already in Database') return;

      const name = formatName(r[nameCol]);
      const email = cleanEmail(r[emailCol]);
      const phone = normalizeIndonesianPhone(r[phoneCol]);

      if (!name || (!email && !phone)) return;

      const parts = name.split(' ');
      queue.push({
        rowIndex: i + 2,
        payload: {
          firstName: parts[0],
          lastName: parts.slice(1).join(' ') || null,
          country: 'Indonesia',
          contactInformation: {
            email: email ? [{ address: email }] : [],
            phone: phone ? [{ number: phone }] : []
          }
        }
      });
    });

    if (handleEmptyOrConfirm(queue.length, 'Infobip')) return;

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    queue.forEach(item => {
      const res = UrlFetchApp.fetch(`https://${BASE_URL}/people/2/persons`, {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'App ' + API_KEY },
        payload: JSON.stringify(item.payload),
        muteHttpExceptions: true
      });

      const responseCode = res.getResponseCode();
      const responseBody = res.getContentText();

      if (responseCode >= 200 && responseCode < 300) {
        // Success: New person created
        sheet.getRange(item.rowIndex, statusCol + 1).setValue('Uploaded');
        sheet.getRange(item.rowIndex, dateCol + 1).setValue(now);
      } 
      else if (responseCode === 409 || (responseCode === 400 && responseBody.includes("already exists"))) {
        // Duplicate Found: Mark accordingly
        sheet.getRange(item.rowIndex, statusCol + 1).setValue('Already in Database');
        sheet.getRange(item.rowIndex, dateCol + 1).setValue(now);
        console.log("Duplicate detected at row " + item.rowIndex);
      } 
      else {
        console.error("Infobip API error on row " + item.rowIndex + ": " + responseBody);
      }
    });

    SpreadsheetApp.getUi().alert('✅ Infobip Sync Complete', 'Check the status column for new uploads and existing duplicates.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    console.error("Infobip Error: " + e.message);
    SpreadsheetApp.getUi().alert('❌ Infobip Error: ' + e.message);
  }
}

/******************************************************
 * BITBYBIT UPLOAD
 ******************************************************/
function uploadToBitbybit() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    setupGenericColumns(sheet, 'Bitbybit Upload Status', 'Bitbybit Upload Date');
    const ctx = getSheetContext(sheet, 'Bitbybit Upload Status', 'Bitbybit Upload Date');
    const p = PropertiesService.getScriptProperties();

    const items = [];
    ctx.data.forEach((r, i) => {
      if (r[ctx.statusCol] === 'Uploaded') return;
      const email = cleanEmail(r[ctx.emailCol]), phone = normalizeIndonesianPhone(r[ctx.phoneCol]), name = formatName(r[ctx.nameCol]);
      if (!email && !phone) return;
      items.push({ row: i + 2, payload: { name, email, phone_number: phone } });
    });

    if (handleEmptyOrConfirm(items.length, 'Bitbybit')) return;

    items.forEach(item => {
      const res = UrlFetchApp.fetch("https://api.bitbybit.id/v1/contacts", {
        method: 'post',
        contentType: 'application/json',
        headers: { "X-API-KEY": p.getProperty('BITBYBIT_API_KEY') },
        payload: JSON.stringify(item.payload),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() < 300) updateStatusOnSheet(sheet, [item.row], ctx.statusCol, ctx.dateCol, null);
    });
    SpreadsheetApp.getUi().alert('✅ Bitbybit Upload Complete');
  } catch (e) { SpreadsheetApp.getUi().alert('❌ Bitbybit Error: ' + e.message); }
}

/******************************************************
 * GUPSHUP UPLOAD
 ******************************************************/
function uploadToGupshup() {
  console.log("Starting uploadToGupshup...");
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    setupGenericColumns(sheet, 'Gupshup Upload Status', 'Gupshup Upload Date');
    const ctx = getSheetContext(sheet, 'Gupshup Upload Status', 'Gupshup Upload Date');
    const p = PropertiesService.getScriptProperties();
    const apiKey = p.getProperty('GUPSHUP_API_KEY');
    const appName = p.getProperty('GUPSHUP_APP_NAME');

    if (!apiKey || !appName) throw new Error("Missing GUPSHUP_API_KEY or GUPSHUP_APP_NAME in Script Properties.");

    const queue = [];
    ctx.data.forEach((r, i) => {
      if (r[ctx.statusCol] === 'Uploaded') return;
      const phone = normalizeIndonesianPhone(r[ctx.phoneCol]);
      if (!phone) return;

      queue.push({
        rowIndex: i + 2,
        payload: {
          name: formatName(r[ctx.nameCol]),
          phone: phone.replace('+', ''), // Gupshup often expects digits only
          email: cleanEmail(r[ctx.emailCol])
        }
      });
    });

    if (handleEmptyOrConfirm(queue.length, 'Gupshup')) return;

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    let successCount = 0;

    queue.forEach(item => {
      // Gupshup Upsert Contact Endpoint
      const url = `https://api.gupshup.io/sm/api/v1/app/${appName}/contact`;
      const options = {
        method: 'post',
        headers: { 'apikey': apiKey },
        payload: item.payload, // Sent as form-data by default in UrlFetchApp
        muteHttpExceptions: true
      };

      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
        sheet.getRange(item.rowIndex, ctx.statusCol + 1).setValue('Uploaded');
        sheet.getRange(item.rowIndex, ctx.dateCol + 1).setValue(now);
        successCount++;
      } else {
        console.error(`Gupshup error on row ${item.rowIndex}: ${res.getContentText()}`);
      }
    });

    SpreadsheetApp.getUi().alert('✅ Gupshup Sync Complete', `Successfully synced ${successCount} contacts.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Gupshup Error: ' + e.message);
  }
}

/******************************************************
 * MEKARI QONTAK UPLOAD
 ******************************************************/
function uploadToQontak() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    setupGenericColumns(sheet, 'Qontak Upload Status', 'Qontak Upload Date');
    const ctx = getSheetContext(sheet, 'Qontak Upload Status', 'Qontak Upload Date');
    const p = PropertiesService.getScriptProperties();
    const listId = p.getProperty('QONTAK_LIST_ID');

    const contacts = [];
    const rows = [];

    ctx.data.forEach((r, i) => {
      if (r[ctx.statusCol] === 'Uploaded') return;
      const email = cleanEmail(r[ctx.emailCol]), phone = normalizeIndonesianPhone(r[ctx.phoneCol]), name = formatName(r[ctx.nameCol]);
      if (!phone) return;

      contacts.push({
        name: name,
        phone_number: phone,
        email: email
      });
      rows.push(i + 2);
    });

    if (handleEmptyOrConfirm(contacts.length, 'Mekari Qontak')) return;

    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: "Bearer " + p.getProperty('QONTAK_API_TOKEN') },
      payload: JSON.stringify({
        contact_list_id: listId,
        contacts: contacts
      }),
      muteHttpExceptions: true
    };
    
    const res = UrlFetchApp.fetch("https://service-chat.qontak.com/api/open/v1/contacts/contact_lists/add_contacts", options);
    if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
      updateStatusOnSheet(sheet, rows, ctx.statusCol, ctx.dateCol, 'Mekari Qontak');
    } else {
      throw new Error(res.getContentText());
    }
  } catch (e) { SpreadsheetApp.getUi().alert('❌ Qontak Error: ' + e.message); }
}

/******************************************************
 * MIXPANEL UPLOAD
 ******************************************************/
function uploadToMixpanel() {
  const sheet = SpreadsheetApp.getActiveSheet();
  setupGenericColumns(sheet, 'Mixpanel Upload Status', 'Mixpanel Upload Date');
  const ctx = getSheetContext(sheet, 'Mixpanel Upload Status', 'Mixpanel Upload Date');
  const profiles = [];
  const rows = [];

  ctx.data.forEach((r, i) => {
    if (r[ctx.statusCol] === 'Uploaded') return;
    const email = cleanEmail(r[ctx.emailCol]), phone = normalizeIndonesianPhone(r[ctx.phoneCol]), name = formatName(r[ctx.nameCol]);
    if (!email) return;
    profiles.push({ "$token": PropertiesService.getScriptProperties().getProperty('MIXPANEL_PROJECT_TOKEN'), "$distinct_id": email, "$set": { "$first_name": name.split(' ')[0], "$last_name": name.split(' ').slice(1).join(' ') || "", "$email": email, "$phone": phone, "Country": "ID" } });
    rows.push(i + 2);
  });

  if (handleEmptyOrConfirm(profiles.length, 'Mixpanel')) return;
  for (let i = 0; i < profiles.length; i += 50) {
    const res = UrlFetchApp.fetch("https://api.mixpanel.com/engage#profile-set", {
      method: 'post', contentType: 'application/json', headers: { 'Accept': 'text/plain' },
      payload: JSON.stringify(profiles.slice(i, i + 50)), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) throw new Error(res.getContentText());
  }
  updateStatusOnSheet(sheet, rows, ctx.statusCol, ctx.dateCol, 'Mixpanel');
}

/******************************************************
 * SEGMENT UPLOAD
 ******************************************************/
function uploadToSegment() {
  const sheet = SpreadsheetApp.getActiveSheet();
  setupGenericColumns(sheet, 'Segment Upload Status', 'Segment Upload Date');
  const ctx = getSheetContext(sheet, 'Segment Upload Status', 'Segment Upload Date');
  const batch = [];
  const rows = [];

  ctx.data.forEach((r, i) => {
    if (r[ctx.statusCol] === 'Uploaded') return;
    const email = cleanEmail(r[ctx.emailCol]), phone = normalizeIndonesianPhone(r[ctx.phoneCol]), name = formatName(r[ctx.nameCol]);
    if (!email) return;
    batch.push({ type: "identify", userId: email, traits: { name, email, phone, address: { country: "ID" } } });
    rows.push(i + 2);
  });

  if (handleEmptyOrConfirm(batch.length, 'Segment')) return;
  const auth = Utilities.base64Encode(PropertiesService.getScriptProperties().getProperty('SEGMENT_WRITE_KEY') + ":");
  for (let i = 0; i < batch.length; i += CONFIG.BATCH_SIZE) {
    const res = UrlFetchApp.fetch("https://api.segment.io/v1/batch", {
      method: 'post', contentType: 'application/json', headers: { Authorization: "Basic " + auth },
      payload: JSON.stringify({ batch: batch.slice(i, i + CONFIG.BATCH_SIZE) }), muteHttpExceptions: true
    });
  }
  updateStatusOnSheet(sheet, rows, ctx.statusCol, ctx.dateCol, 'Segment');
}

/******************************************************
 * CONNECTION CHECKS
 ******************************************************/
function checkBrazeConnection() {
  console.log("Checking Braze Connection...");
  const p = PropertiesService.getScriptProperties();
  const endpoint = p.getProperty('BRAZE_REST_ENDPOINT');
  const targetUrl = `${endpoint}/users/track`;
  
  try {
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: `Bearer ${p.getProperty('BRAZE_API_KEY')}` },
      payload: JSON.stringify({ external_ids: ['test'] }),
      muteHttpExceptions: true
    });
    
    const code = res.getResponseCode();
    const isOk = (code === 200 || code === 400);
    console.log("Braze Connection response code: " + code);
    
    SpreadsheetApp.getUi().alert(
      isOk ? '✅ Braze Connection OK' : '❌ Braze Connection Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) { 
    console.error("Connection check error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ Braze error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    ); 
  }
}

function checkMoEngageConnection() {
  console.log("Checking MoEngage Connection...");
  const p = PropertiesService.getScriptProperties();
  const region = p.getProperty('MOENGAGE_REGION');
  const targetUrl = `https://api-${region}.moengage.com/v1/customer/add`;
  
  try {
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { 
        "MOE-APP-ID": p.getProperty('MOE_APP_ID'), 
        "X-Api-Key": p.getProperty('MOE_DATA_API_KEY') 
      },
      payload: JSON.stringify({ elements: [] }),
      muteHttpExceptions: true
    });
    
    const code = res.getResponseCode();
    console.log("MoEngage Response Code: " + code);

    SpreadsheetApp.getUi().alert(
      code === 200 ? '✅ MoEngage Connection OK' : '❌ MoEngage Connection Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) { 
    console.error("MoEngage Connection error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ MoEngage error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    ); 
  }
}

function checkInfobipConnection() {
  console.log("Checking Infobip Connection...");
  const p = PropertiesService.getScriptProperties();
  const baseUrl = p.getProperty('INFOBIP_BASE_URL');
  const targetUrl = `https://${baseUrl}/people/2/persons?limit=1`;
  
  try {
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'get',
      headers: { Authorization: 'App ' + p.getProperty('INFOBIP_API_KEY') },
      muteHttpExceptions: true
    });
    
    const code = res.getResponseCode();
    console.log("Infobip Connection response code: " + code);
    
    SpreadsheetApp.getUi().alert(
      code === 200 ? '✅ Infobip Connection OK' : '❌ Infobip Connection Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) { 
    console.error("Connection check error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ Infobip error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    ); 
  }
}

function checkBitbybitConnection() {
  console.log("Checking Bitbybit Connection...");
  const p = PropertiesService.getScriptProperties();
  const targetUrl = "https://api.bitbybit.id/v1/groups";
  
  try {
    const apiKey = p.getProperty('BITBYBIT_API_KEY');
    
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'get',
      headers: { "X-API-KEY": apiKey },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    console.log("Bitbybit Response Code: " + code);

    SpreadsheetApp.getUi().alert(
      code === 200 ? '✅ Bitbybit Connection OK' : '❌ Bitbybit Connection Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    console.error("Bitbybit Connection error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ Bitbybit error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function checkGupshupConnection() {
  console.log("Checking Gupshup Connection...");
  const p = PropertiesService.getScriptProperties();
  const appName = p.getProperty('GUPSHUP_APP_NAME');
  const targetUrl = `https://api.gupshup.io/sm/api/v1/app/${appName}/msg/history`;
  
  try {
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'get',
      headers: { 'apikey': p.getProperty('GUPSHUP_API_KEY') },
      muteHttpExceptions: true
    });
    
    const code = res.getResponseCode();
    const isOk = (code !== 401);
    
    SpreadsheetApp.getUi().alert(
      isOk ? '✅ Gupshup Connected' : '❌ Gupshup Auth Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) { 
    console.error("Gupshup Connection error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ Gupshup error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    ); 
  }
}

function checkQontakConnection() {
  console.log("Checking Mekari Qontak Connection...");
  const p = PropertiesService.getScriptProperties();
  const targetUrl = "https://service-chat.qontak.com/api/open/v1/contacts/contact_lists";
  
  try {
    const token = p.getProperty('QONTAK_API_TOKEN');
    
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'get',
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    console.log("Qontak Response Code: " + code);

    SpreadsheetApp.getUi().alert(
      code === 200 ? '✅ Mekari Qontak Connection OK' : '❌ Qontak Connection Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}\nResponse: ${res.getContentText().substring(0, 100)}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    console.error("Qontak Connection error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ Qontak error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function checkMixpanelConnection() {
  console.log("Checking Mixpanel Connection...");
  try {
    const p = PropertiesService.getScriptProperties();
    const token = p.getProperty('MIXPANEL_PROJECT_TOKEN');
    if(!token) throw new Error("MIXPANEL_PROJECT_TOKEN is missing in Script Properties.");
    
    // Test with a dummy update for a test user
    const testPayload = [{
      "$token": token,
      "$distinct_id": "connection_test",
      "$set": { "Last Connection Test": new Date().toISOString() }
    }];
    
    sendToMixpanel(testPayload);
    SpreadsheetApp.getUi().alert('✅ Mixpanel Connection OK');
  } catch (e) { 
    console.error("Connection check error: " + e.message);
    SpreadsheetApp.getUi().alert('❌ Mixpanel error: ' + e.message); 
  }
}
function sendToMixpanel(batch) {
  const url = "https://api.mixpanel.com/engage#profile-set";
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Accept': 'text/plain' },
    payload: JSON.stringify(batch),
    muteHttpExceptions: true
  };
  
  const res = UrlFetchApp.fetch(url, options);
  // Mixpanel returns '1' on success for this endpoint
  if (res.getResponseCode() !== 200 || res.getContentText() !== '1') {
    console.error("Mixpanel API Error: " + res.getContentText());
    throw new Error("Mixpanel refused data: " + res.getContentText());
  }
}

function checkSegmentConnection() {
  console.log("Checking Segment Connection...");
  const p = PropertiesService.getScriptProperties();
  const targetUrl = "https://api.segment.io/v1/batch";
  
  try {
    const writeKey = p.getProperty('SEGMENT_WRITE_KEY');
    const auth = Utilities.base64Encode(writeKey + ":");
    
    const res = UrlFetchApp.fetch(targetUrl, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: "Basic " + auth },
      payload: JSON.stringify({ batch: [], sentAt: new Date().toISOString() }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    console.log("Segment Response Code: " + code);

    SpreadsheetApp.getUi().alert(
      code === 200 ? '✅ Segment Connection OK' : '❌ Segment Connection Failed',
      `Endpoint: ${targetUrl}\nStatus Code: ${code}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    console.error("Segment Connection error: " + e.message);
    SpreadsheetApp.getUi().alert(
      '❌ Segment error: ' + e.message,
      `Attempted URL: ${targetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function checkConn(url, name) { 
  try { const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); SpreadsheetApp.getUi().alert(`${name} check returned ${res.getResponseCode()}. Connection reachable.`); } 
  catch(e) { SpreadsheetApp.getUi().alert(`❌ ${name} Unreachable: ` + e.message); }
}

/******************************************************
 * STATS
 ******************************************************/
function viewBrazeStats() { viewStats('Braze Upload Status', 'Braze'); }
function viewMoEngageStats() { viewStats('MoEngage Upload Status', 'MoEngage'); }
function viewInfobipStats() { viewStats('Infobip Upload Status', 'Infobip'); }
function viewBitbybitStats() { viewStats('Bitbybit Upload Status', 'Bitbybit'); }
function viewGupshupStats() { viewStats('Gupshup Upload Status', 'Gupshup'); }
function viewQontakStats() { viewStats('Qontak Upload Status', 'Mekari Qontak'); }
function viewMixpanelStats() { viewStats('Mixpanel Upload Status', 'Mixpanel'); }
function viewSegmentStats() { viewStats('Segment Upload Status', 'Segment'); }

function viewStats(statusHeader, title) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ctx = getSheetContext(sheet, statusHeader, '');
  
  if (ctx.statusCol === -1) {
    SpreadsheetApp.getUi().alert('❌ Column Missing', 'The column "' + statusHeader + '" was not found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const totalRecords = ctx.data.length;
  if (totalRecords === 0) return;

  // Calculate specific statuses
  const uploadedCount = ctx.data.filter(r => r[ctx.statusCol] === 'Uploaded').length;
  const duplicateCount = ctx.data.filter(r => r[ctx.statusCol] === 'Already in Database').length;
  
  // Both "Uploaded" and "Already in Database" count as "Processed/Synced"
  const totalSynced = uploadedCount + duplicateCount;
  const pendingCount = totalRecords - totalSynced;
  const completionRate = ((totalSynced / totalRecords) * 100).toFixed(1);

  // Visual Progress Bar
  const progressBarLength = 10;
  const completedBlocks = Math.round((totalSynced / totalRecords) * progressBarLength);
  const progressBar = '✅' + '🟩'.repeat(completedBlocks) + '⬜'.repeat(progressBarLength - completedBlocks);

  // Build Report Message
  let reportMessage = 
    `Platform: ${title}\n` +
    `${progressBar} ${completionRate}%\n\n` +
    `📈 Total Records: ${totalRecords}\n` +
    `✨ New Uploads: ${uploadedCount}\n` +
    `👥 Already in DB: ${duplicateCount}\n` +
    `⏳ Pending: ${pendingCount}\n\n`;

  if (pendingCount === 0) {
    reportMessage += "Status: Fully Optimized! 🎉";
  } else {
    reportMessage += "Status: Sync Required.";
  }

  SpreadsheetApp.getUi().alert('📊 ' + title + ' Statistics', reportMessage, SpreadsheetApp.getUi().ButtonSet.OK);
}

/******************************************************
 * RETURN LISTS
 ******************************************************/
function listRecentInfobipPeople() {
  console.log("Fetching recent people from Infobip...");
  try {
    const p = PropertiesService.getScriptProperties();
    const res = UrlFetchApp.fetch(`https://${p.getProperty('INFOBIP_BASE_URL')}/people/2/persons?limit=3&sort=createdAt,desc`, {
      headers: { Authorization: 'App ' + p.getProperty('INFOBIP_API_KEY') },
      muteHttpExceptions: true
    });

    const responseText = res.getContentText();
    const data = JSON.parse(responseText);

    if (res.getResponseCode() === 200 && data.persons && data.persons.length > 0) {
      console.log("Infobip list retrieved successfully.");
      
      // Map the raw data into a readable string
      const report = data.persons.map(person => {
        const phone = (person.contactInformation.phone && person.contactInformation.phone.length > 0) 
                      ? person.contactInformation.phone[0].number 
                      : 'N/A';
        const email = (person.contactInformation.email && person.contactInformation.email.length > 0) 
                      ? person.contactInformation.email[0].address 
                      : 'N/A';
        
        return `👤 Name: ${person.firstName} ${person.lastName || ''}\n` +
               `   ID: ${person.id}\n` +
               `   📞 Phone: ${phone}\n` +
               `   📧 Email: ${email}\n` +
               `   📅 Created: ${person.createdAt.replace('T', ' ')}`;
      }).join('\n\n' + "─".repeat(30) + '\n\n');

      SpreadsheetApp.getUi().alert('👥 Recent Infobip People', report, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      console.error("Failed to parse Infobip list: " + responseText);
      SpreadsheetApp.getUi().alert('❌ Error', 'Could not retrieve recent users. Response: ' + responseText.substring(0, 200));
    }
  } catch (e) {
    console.error("listRecentInfobipPeople Error: " + e.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

function listRecentMixpanelUsers() {
  console.log("Fetching recent profiles from Mixpanel Engage Query API...");
  try {
    const p = PropertiesService.getScriptProperties();
    const projectId = p.getProperty('MIXPANEL_PROJECT_ID');
    const serviceAccountUser = p.getProperty('MIXPANEL_SERVICE_ACCOUNT_USER');
    const serviceAccountSecret = p.getProperty('MIXPANEL_SERVICE_ACCOUNT_SECRET');

    if (!projectId || !serviceAccountUser || !serviceAccountSecret) {
      throw new Error("Missing Mixpanel Query Credentials in Script Properties.");
    }

    const authHeader = "Basic " + Utilities.base64Encode(serviceAccountUser + ":" + serviceAccountSecret);

    // Endpoint for Engage Query
    const url = `https://mixpanel.com/api/2.0/engage?project_id=${projectId}`;

    const options = {
      method: 'post',
      headers: {
        "Authorization": authHeader,
        "Accept": "application/json",
        "Content-Type": "application/x-www-form-urlencoded"
      },
      payload: "limit=3", 
      muteHttpExceptions: true
    };

    const res = UrlFetchApp.fetch(url, options);
    const responseText = res.getContentText();
    const result = JSON.parse(responseText);

    if (res.getResponseCode() === 200 && result.results) {
      if (result.results.length === 0) {
        SpreadsheetApp.getUi().alert('ℹ️ No profiles found in Mixpanel.');
        return;
      }

      console.log("Mixpanel list retrieved successfully.");
      
      // Map the raw data into a readable string
      const report = result.results.map(u => {
        const props = u.$properties || {};
        const firstName = props.$first_name || '';
        const lastName = props.$last_name || '';
        const fullName = (firstName || lastName) ? `${firstName} ${lastName}`.trim() : 'Anonymous';
        
        return `👤 Name: ${fullName}\n` +
               `   ID: ${u.$distinct_id}\n` +
               `   📧 Email: ${props.$email || 'N/A'}\n` +
               `   📞 Phone: ${props.$phone || 'N/A'}\n` +
               `   📍 Country: ${props.Country || 'N/A'}`;
      }).join('\n\n' + "─".repeat(30) + '\n\n');

      SpreadsheetApp.getUi().alert('👥 Recent Mixpanel Profiles', report, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      console.error("Mixpanel Query Error: " + responseText);
      SpreadsheetApp.getUi().alert('❌ Query Failed', (result.error || "Check execution logs"));
    }
  } catch (e) {
    console.error("listRecentMixpanelUsers Error: " + e.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e.message);
  }
}

/******************************************************
 * RESETS
 ******************************************************/
function resetBrazeUploadStatus() { resetStatusColumns('Braze Upload Status', 'Braze Upload Date'); }
function resetMoEngageUploadStatus() { resetStatusColumns('MoEngage Upload Status', 'MoEngage Upload Date'); }
function resetInfobipUploadStatus() { resetStatusColumns('Infobip Upload Status', 'Infobip Upload Date'); }
function resetBitbybitUploadStatus() { resetStatusColumns('Bitbybit Upload Status', 'Bitbybit Upload Date'); }
function resetGupshupUploadStatus() { resetStatusColumns('Gupshup Upload Status', 'Gupshup Upload Date'); }
function resetQontakUploadStatus() { resetStatusColumns('Qontak Upload Status', 'Qontak Upload Date'); }
function resetMixpanelUploadStatus() { resetStatusColumns('Mixpanel Upload Status', 'Mixpanel Upload Date'); }
function resetSegmentUploadStatus() { resetStatusColumns('Segment Upload Status', 'Segment Upload Date'); }

function resetStatusColumns(statusHeader, dateHeader) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const sIdx = headers.indexOf(statusHeader);
  const dIdx = headers.indexOf(dateHeader);

  // 1. Safety Check: Ensure the column actually exists
  if (sIdx === -1) {
    console.warn("Reset aborted: Column " + statusHeader + " not found.");
    SpreadsheetApp.getUi().alert('❌ Error', 'Column "' + statusHeader + '" was not found in this sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // 2. Confirmation Prompt
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '🔄 Confirm Status Reset',
    'Are you sure you want to clear the upload status and date for ' + statusHeader.replace(' Upload Status', '') + '?\n\nThis will allow you to re-upload these users.',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    console.log("Reset cancelled by user for: " + statusHeader);
    return;
  }

  // 3. Execution
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // Nothing to clear

  console.log("Resetting columns: " + statusHeader + " and " + dateHeader);
  
  // Clear the Status Column
  sheet.getRange(2, sIdx + 1, lastRow - 1).clearContent();
  
  // Clear the Date Column if it exists
  if (dIdx !== -1) {
    sheet.getRange(2, dIdx + 1, lastRow - 1).clearContent();
  }

  // 4. Result Notification
  ui.alert(
    '✅ Reset Successful', 
    'The upload status for ' + statusHeader.replace(' Upload Status', '') + ' has been cleared for ' + (lastRow - 1) + ' rows.', 
    ui.ButtonSet.OK
  );
}

/******************************************************
 * SHARED HELPERS & UTILS
 ******************************************************/
function getSheetContext(sheet, statusName, dateName) {
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { data, statusCol: headers.indexOf(statusName), dateCol: headers.indexOf(dateName), nameCol: headers.indexOf('Name'), emailCol: headers.indexOf('Email'), phoneCol: headers.indexOf('Phone Number') };
}
function handleEmptyOrConfirm(count, platform) {
  if (count === 0) { SpreadsheetApp.getUi().alert(`No new users for ${platform}.`); return true; }
  return SpreadsheetApp.getUi().alert('Confirm', `Upload ${count} to ${platform}?`, SpreadsheetApp.getUi().ButtonSet.YES_NO) !== SpreadsheetApp.getUi().Button.YES;
}
function updateStatusOnSheet(sheet, rows, sIdx, dIdx, platform) {
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  rows.forEach(r => { 
    sheet.getRange(r, sIdx + 1).setValue('Uploaded'); 
    sheet.getRange(r, dIdx + 1).setValue(now); 
  });
  if (platform) SpreadsheetApp.getUi().alert(`✅ ${platform} Upload Complete`);
}
function setupGenericColumns(sheet, statusLabel, dateLabel) {
  const lastCol = sheet.getLastColumn();
  // Handle empty sheets
  if (lastCol === 0) {
    sheet.getRange(1, 1).setValue(statusLabel);
    sheet.getRange(1, 2).setValue(dateLabel);
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Create Status Column if missing
  if (headers.indexOf(statusLabel) === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(statusLabel);
  }
  
  // Re-fetch headers to check for Date column
  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (updatedHeaders.indexOf(dateLabel) === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue(dateLabel);
  }
}

/******************************************************
 * DATA CLEANING
 ******************************************************/
function normalizePhonesToE164() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const idx = data[0].map(h => h.toLowerCase()).findIndex(h => h.includes('phone'));
  if (idx === -1) return;
  for (let i = 1; i < data.length; i++) { if (data[i][idx]) data[i][idx] = normalizeIndonesianPhone(data[i][idx]); }
  sheet.getDataRange().setValues(data);
  SpreadsheetApp.getUi().alert('✅ Phone formatting complete.');
}
function normalizeIndonesianPhone(v) {
  if (!v) return '';
  let d = v.toString().replace(/\D/g, '');
  if (d.length < 8) return '';
  if (d.startsWith('62')) return '+' + d;
  if (d.startsWith('0')) return '+62' + d.slice(1);
  return '+62' + d;
}
function cleanEmail(v) { 
  return v ? v.toString().trim().toLowerCase() : ''; 
  }



function formatName(v) { 
  return v ? v.toString().toLowerCase().split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ') : ''; 
  }