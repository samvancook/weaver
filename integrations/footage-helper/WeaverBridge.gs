// Inventory is read-only. Only receipt pointers are stored in Script Properties.
// Configure WEAVER_BRIDGE_URL to the reviewed Weaver HTTPS origin, and
// WEAVER_BRIDGE_ENABLED=true after the receiver and OAuth trust are configured.
function weaverManifest_(eventName) {
  const report = getIntakeReport(eventName);
  if (!report.event) throw new Error('No Add Event record was found.');
  if (report.closed) throw new Error('This event is already closed.');
  const folder = String(report.event.omFolderLink || '').match(/folders\/([-\w]+)/);
  if (!folder) throw new Error('The event needs an OM folder link.');
  return {
    sourceSpreadsheetId: RETITLE_CONFIG.spreadsheetId,
    eventName: String(report.eventName).replace(/\s+/g, ' ').trim(),
    omFolderId: folder[1], closed: false,
    titles: report.titles.map(function(t) {
      return {sourceRow:t.rowNumber, performer:t.performer, title:t.title,
        rights:t.rights, contentLink:t.contentLink};
    })
  };
}
function weaverRequest_(path, method, payload) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('WEAVER_BRIDGE_ENABLED') !== 'true') throw new Error('Weaver bridge is not enabled yet.');
  const origin = String(props.getProperty('WEAVER_BRIDGE_URL') || '').replace(/\/$/, '');
  if (!/^https:\/\/[a-zA-Z0-9.-]+$/.test(origin)) throw new Error('Configure the reviewed Weaver HTTPS origin.');
  const token = ScriptApp.getIdentityToken();
  if (!token) throw new Error('Owner authorization is required for the Weaver connection.');
  const options = {method:method, muteHttpExceptions:true,
    headers:{Authorization:'Bearer ' + token}};
  if (payload) { options.contentType='application/json'; options.payload=JSON.stringify(payload); }
  const response = UrlFetchApp.fetch(origin + path, options);
  let result;
  try { result = JSON.parse(response.getContentText()); }
  catch (e) { throw new Error('Weaver did not return a readable receipt.'); }
  if (response.getResponseCode() !== 200 || !result.ok || !result.receipt) {
    throw new Error(result.error || 'Weaver did not confirm this handoff.');
  }
  return result.receipt;
}
function weaverPointerKey_(manifest) {
  return 'WEAVER_EVENT_' + Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify([manifest.omFolderId, intakeKey_(manifest.eventName)]))
    .map(function(b) {return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
}
function sendEventToWeaver(eventName) {
  // Always reconstruct from current source rows; never trust browser title data.
  const manifest = weaverManifest_(eventName);
  const receipt = weaverRequest_('/api/footage/events', 'post', manifest);
  PropertiesService.getScriptProperties().setProperty(weaverPointerKey_(manifest), receipt.eventId);
  return weaverReceiptForUi_(receipt);
}
function checkWeaverEvent(eventName) {
  const manifest = weaverManifest_(eventName);
  const eventId = PropertiesService.getScriptProperties().getProperty(weaverPointerKey_(manifest));
  if (!eventId) throw new Error('No Weaver receipt is recorded for this event. Send it first.');
  const receipt = weaverRequest_('/api/footage/events/' + encodeURIComponent(eventId), 'get');
  // A prior successful receipt does not confirm newer edits to the inventory.
  const clean = function(x) {return String(x || '').trim();};
  const current = manifest.titles.map(function(t) {
    return {sourceRow:t.sourceRow, performer:clean(t.performer), title:clean(t.title), rights:clean(t.rights), contentLink:clean(t.contentLink)};
  }).sort(function(a,b){return a.sourceRow-b.sourceRow;});
  if (JSON.stringify(current) !== JSON.stringify(receipt.manifest.titles)) {
    receipt.readyForClosureReview=false; receipt.status='inventory_changed';
  }
  return weaverReceiptForUi_(receipt);
}
function weaverReceiptForUi_(r) {
  return {eventId:r.eventId, status:r.status, submittedCount:r.submittedCount,
    acceptedCount:r.acceptedCount, blockedCount:r.blockedCount,
    readyForClosureReview:r.readyForClosureReview, checkedAt:r.checkedAt,
    currentUnavailableFileIds:r.currentUnavailableFileIds || [],
    blocked:(r.blocked || []).map(function(t){return {sourceRow:t.sourceRow, performer:t.performer, title:t.title, reason:t.reason};})};
}

// Owner-only editor diagnostic. Never returns or logs the bearer token.
// The decoded audience identifies the configuration; Weaver still verifies the
// signed token independently on every request.
function weaverConnectionInfo_() {
  const token = ScriptApp.getIdentityToken();
  if (!token) throw new Error('Authorize openid and userinfo.email first.');
  const payload = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(token.split('.')[1])).getDataAsString());
  return {audience:payload.aud, effectiveEmail:payload.email, emailVerified:payload.email_verified};
}
