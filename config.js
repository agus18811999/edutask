/* EduTask browser configuration. Never put private credentials in this file. */
(() => {
  'use strict';
  const API_URL = '/api/proxy';
  const TOKEN_KEY = 'edutask_session_v1';
  const READ_ACTIONS = new Set([
    'health','auth.session','dashboard.intern','dashboard.mentor','dashboard.refresh',
    'tasks.list','tasks.get','tasks.activity','records.list','users.list','users.get',
    'projects.list','batches.list','notifications.list','evaluations.list','reports.tasks',
    'reports.intern','reports.performance','tasks.comments'
  ]);
  const inFlight = new Map();
  const state = { token: sessionStorage.getItem(TOKEN_KEY) || '', user: null };
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const normalizeError = (error, action) => ({
    code: error?.code || (error?.name === 'AbortError' ? 'TIMEOUT' : 'NETWORK_ERROR'),
    message: error?.message || `Permintaan ${action} gagal. Periksa koneksi internet.`,
    action
  });
  const api = async (action, payload = {}, options = {}) => {
    const reading = READ_ACTIONS.has(action) && options.dedupe !== false;
    const key = reading ? JSON.stringify([action, payload, state.token]) : '';
    if (reading && inFlight.has(key)) return inFlight.get(key);
    const execute = async () => {
      const retries = reading ? 1 : 0;
      for (let attempt = 0; attempt <= retries; attempt++) {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), options.timeout || 25000);
        try {
          const response = await fetch(API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
            body: JSON.stringify({ action, payload, token: state.token, requestId: options.requestId || undefined }),
            credentials: 'same-origin', signal: controller.signal, cache: 'no-store'
          });
          const raw = await response.text();
          let result;
          try { result = JSON.parse(raw); } catch (_) { throw { code: 'INVALID_RESPONSE', message: 'Server mengembalikan respons yang tidak valid.' }; }
          if (!response.ok || !result.success) {
            const error = { code: result.code || 'API_ERROR', message: result.message || `HTTP ${response.status}` };
            if (error.code === 'UNAUTHORIZED' || error.code === 'SESSION_EXPIRED') {
              setSession('');
              window.dispatchEvent(new CustomEvent('edutask:unauthorized'));
            }
            throw error;
          }
          return result.data;
        } catch (error) {
          if (attempt < retries && ['AbortError','TypeError'].includes(error?.name)) { await sleep(400 * (attempt + 1)); continue; }
          throw normalizeError(error, action);
        } finally { clearTimeout(timeout); }
      }
    };
    const promise = execute();
    if (reading) inFlight.set(key, promise);
    try { return await promise; } finally { if (reading) inFlight.delete(key); }
  };
  function setSession(token, user = null) {
    state.token = token || '';
    state.user = user;
    if (token) sessionStorage.setItem(TOKEN_KEY, token);
    else sessionStorage.removeItem(TOKEN_KEY);
    inFlight.clear();
  }
  function idempotencyKey() {
    if (crypto?.randomUUID) return crypto.randomUUID();
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    return [...bytes].map(b => b.toString(16).padStart(2, '0')).join('');
  }
  window.Edutask = Object.freeze({ api, state, setSession, idempotencyKey, API_URL, TOKEN_KEY });
})();
