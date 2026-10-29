// Reviewed live-test endpoint. This function accepts no browser-supplied URL.
const WEAVER_REVIEWED_TEST_URL = 'https://footage-bridge---weaver-iasomqy6ja-uc.a.run.app';
function weaverRequireOwner_() {
  const active = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const owner = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!active || !owner || active !== owner) throw new Error('The deploying owner must authorize and connect Weaver first.');
}
function configureWeaverLiveTest() {
  weaverRequireOwner_();
  const info = weaverConnectionInfo_();
  if (!info.audience || !info.emailVerified || !info.effectiveEmail) throw new Error('Google did not confirm the connection identity.');
  PropertiesService.getScriptProperties().setProperties({
    WEAVER_BRIDGE_URL: WEAVER_REVIEWED_TEST_URL,
    WEAVER_BRIDGE_ENABLED: 'true'
  });
  return info;
}
function getWeaverConnectionInfo() {
  weaverRequireOwner_();
  return weaverConnectionInfo_();
}
