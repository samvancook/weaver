import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';

const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const source = app.slice(app.indexOf('function setAdminAuthUi('), app.indexOf('async function fetchSheetValues('));

function setup() {
  const clients = [], timers = new Map(), requests = [];
  let time = 1000, timerId = 0;
  const elements = { adminSignIn: {}, adminAuthStatus: {} };
  const context = {
    elements, Headers, Number, Promise, Error,
    Date: { now: () => time },
    runtimeConfig: { googleOAuthClientId: 'test-client' }, GOOGLE_ADMIN_SCOPES: 'unchanged-scope',
    setTimeout: (fn, ms) => { const id = ++timerId; timers.set(id, { fn, ms }); return id; },
    clearTimeout: id => timers.delete(id),
    window: { navigator: { userActivation: { isActive: false } }, google: { accounts: { oauth2: {
      initTokenClient(config) {
        const client = { config, requestAccessToken(options) { client.options = options; } };
        clients.push(client); return client;
      }
    } } } },
    fetch: async (path, options) => { requests.push({ path, options }); return { status: 200 }; }
  };
  const api = runInNewContext(`let googleAdminAccessToken='', googleAdminTokenRequest=null,
    googleAdminTokenExpiresAt=0, googleAdminExpiryTimer=null;
    ${source}
    ({getGoogleAdminAccessToken, adminFetch, clearGoogleAdminSession,
      token:()=>googleAdminAccessToken})`, context);
  return { api, context, clients, timers, requests, elements,
    advance: ms => { time += ms; },
    fire: ms => { const found = [...timers.values()].find(t => t.ms === ms); assert.ok(found); found.fn(); } };
}

async function signIn(s, token = 'test-token', expires = 3600) {
  const pending = s.api.getGoogleAdminAccessToken({ interactive: true });
  s.clients.at(-1).config.callback({ access_token: token, expires_in: expires });
  return pending;
}

test('background loads never open Google or send a request without a session', async () => {
  const s = setup();
  await assert.rejects(s.api.adminFetch('/private'), /Sign in button/);
  assert.equal(s.clients.length, 0); assert.equal(s.requests.length, 0);
  assert.equal(s.elements.adminSignIn.disabled, false);
});

test('concurrent user actions share one popup and reuse only an unexpired token', async () => {
  const s = setup(); s.context.window.navigator.userActivation.isActive = true;
  const a = s.api.getGoogleAdminAccessToken(), b = s.api.getGoogleAdminAccessToken();
  assert.equal(s.clients.length, 1);
  assert.equal(s.clients[0].config.scope, 'unchanged-scope');
  s.clients[0].config.callback({ access_token: 'test-token', expires_in: 3600 });
  assert.deepEqual(await Promise.all([a, b]), ['test-token', 'test-token']);
  s.context.window.navigator.userActivation.isActive = false;
  await s.api.adminFetch('/private');
  assert.equal(s.requests[0].options.headers.get('Authorization'), 'Bearer test-token');
  assert.equal(s.clients.length, 1);
  s.advance(3600000);
  await assert.rejects(s.api.adminFetch('/private'), /Sign in button/);
  assert.equal(s.api.token(), ''); assert.equal(s.requests.length, 1);
});

test('blocked or closed popups fail promptly and permit retry without reloading', async () => {
  for (const type of ['popup_failed_to_open', 'popup_closed', 'unknown']) {
    const s = setup();
    const p = s.api.getGoogleAdminAccessToken({ interactive: true });
    s.clients[0].config.error_callback({ type });
    await assert.rejects(p, /popup|closed|could not open/);
    assert.equal(s.timers.size, 0); assert.equal(s.elements.adminSignIn.disabled, false);
    const retry = s.api.getGoogleAdminAccessToken({ interactive: true });
    assert.equal(s.clients.length, 2);
    s.clients[0].config.callback({ access_token: 'late-old-token', expires_in: 3600 });
    assert.equal(s.api.token(), '');
    s.clients[1].config.callback({ access_token: 'new-token', expires_in: 3600 });
    assert.equal(await retry, 'new-token');
  }
});

test('timed out or denied attempts cannot revive a session through a late callback', async () => {
  const s = setup();
  const p = s.api.getGoogleAdminAccessToken({ interactive: true });
  s.fire(120000); await assert.rejects(p, /no page refresh/);
  s.clients[0].config.callback({ access_token: 'late-token', expires_in: 3600 });
  assert.equal(s.api.token(), '');
  const denied = s.api.getGoogleAdminAccessToken({ interactive: true });
  s.clients[1].config.callback({ error: 'access_denied' });
  await assert.rejects(denied, /not completed/);
  assert.equal(s.elements.adminSignIn.disabled, false);
});

test('expiry updates the visible session state without touching form data', async () => {
  const s = setup(); s.elements.unsavedDecision = { value: 'keep this note' };
  await signIn(s, 'test-token', 100);
  assert.equal(s.elements.adminSignIn.textContent, 'Switch account');
  s.fire(90000);
  assert.equal(s.api.token(), ''); assert.match(s.elements.adminAuthStatus.textContent, /expired/);
  assert.equal(s.elements.adminSignIn.textContent, 'Sign in');
  assert.equal(s.elements.unsavedDecision.value, 'keep this note');
});

test('a 401 never replays the original mutation or opens a second popup', async () => {
  const s = setup(); await signIn(s);
  s.context.fetch = async () => { s.requests.push('one mutation'); return { status: 401 }; };
  await assert.rejects(s.api.adminFetch('/save', { method: 'POST', body: 'decision' }), /retry this action/);
  assert.equal(s.requests.length, 1); assert.equal(s.clients.length, 1);
  assert.equal(s.api.token(), '');
});

test('an older in-flight 401 cannot clear a newer sign-in', async () => {
  const s = setup(); await signIn(s, 'old-token');
  let finish;
  s.context.fetch = () => new Promise(resolve => { finish = resolve; });
  const oldRequest = s.api.adminFetch('/read'); await Promise.resolve();
  s.api.clearGoogleAdminSession(); await signIn(s, 'new-token');
  finish({ status: 401 }); await assert.rejects(oldRequest, /retry this action/);
  assert.equal(s.api.token(), 'new-token');
});

test('missing configuration, unloaded Google and malformed tokens remain recoverable', async () => {
  const s = setup(); s.context.runtimeConfig.googleOAuthClientId = '';
  await assert.rejects(s.api.getGoogleAdminAccessToken({ interactive: true }), /not configured/);
  assert.equal(s.elements.adminSignIn.disabled, false);
  s.context.runtimeConfig.googleOAuthClientId = 'test-client';
  s.context.window.google = null;
  await assert.rejects(s.api.getGoogleAdminAccessToken({ interactive: true }), /still loading/);
  const t = setup(); const p = t.api.getGoogleAdminAccessToken({ interactive: true });
  t.clients[0].config.callback({ access_token: 'invalid-without-expiry' });
  await assert.rejects(p, /valid session/);
});
