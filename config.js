/* EduTask browser configuration. Never put private credentials in this file. */
(() => {
  'use strict';
  const API_URL = '/api/proxy';
  const TOKEN_KEY = 'edutask_session_v1';
  // Same-tab session; do not persist bearer tokens in localStorage by default.
  const CACHE_TTL = {'directory.snapshot':12000,'records.view':12000,'users.list':20000,'projects.list':20000,'tasks.list':12000,'notifications.list':12000,'evaluations.list':12000,'records.list':15000,'reports.intern':15000,'reports.performance':15000};
  const readCache = new Map();
  const USER_HINT_KEY = 'edutask_user_hint_v2';
  const perfSamples = [];
  const markApi = (action, elapsed, success, backendMs, extra = {}) => {
    const row = {action, roundtripMs: Math.round(elapsed), backendMs: Number.isFinite(backendMs)?backendMs:null, proxyMs: Number.isFinite(extra.proxyMs)?extra.proxyMs:null, hashLookupMs: Number.isFinite(extra.hashLookupMs)?extra.hashLookupMs:null, kdfMs: Number.isFinite(extra.kdfMs)?extra.kdfMs:null, ok: success, time:new Date().toLocaleTimeString('id-ID')};
    perfSamples.push(row); if(perfSamples.length>80) perfSamples.shift();
    console.info(`[EduTask • API] ${action}: ${row.roundtripMs}ms${row.backendMs===null?'':` • GAS ${row.backendMs}ms`}${row.kdfMs===null?'':` • Native KDF ${row.kdfMs}ms`}${row.hashLookupMs===null?'':` • Salt Lookup ${row.hashLookupMs}ms`} • ${success?'OK':'ERROR'}`);
  };
  const perfReport=()=>{
    console.group('⚡ EduTask Performance Report');
    console.table(perfSamples);
    if (performance.getEntriesByType) {
      const nav=performance.getEntriesByType('navigation')[0];
      if(nav) console.table({domReadyMs:Math.round(nav.domContentLoadedEventEnd-nav.startTime), loadMs:Math.round(nav.loadEventEnd-nav.startTime), ttfbMs:Math.round(nav.responseStart-nav.requestStart)});
    }
    console.groupEnd(); return perfSamples.slice();
  };
  window.EdutaskPerf = Object.freeze({report:perfReport,samples:()=>perfSamples.slice()});
  const READ_ACTIONS = new Set([
    'health','auth.session','dashboard.intern','dashboard.mentor','dashboard.refresh',
    'tasks.list','tasks.get','tasks.activity','records.list','records.view','directory.snapshot','users.list','users.get',
    'projects.list','batches.list','notifications.list','evaluations.list','reports.tasks',
    'reports.intern','reports.performance','tasks.comments'
  ]);
  const inFlight = new Map();
  const cachePrefix = (action,payload) => JSON.stringify([action,payload,state.token]);
  function invalidateReads(actions=null){
    if(!actions){readCache.clear();return;}
    for(const [key,row] of readCache)if(actions.includes(row.action))readCache.delete(key);
  }
  const safeUserHint = () => { try { return JSON.parse(sessionStorage.getItem(USER_HINT_KEY)||'null'); } catch (_) { return null; } };
  const state = { token: sessionStorage.getItem(TOKEN_KEY) || '', user: safeUserHint() };
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const normalizeError = (error, action) => ({
    code: error?.code || (error?.name === 'AbortError' ? 'TIMEOUT' : 'NETWORK_ERROR'),
    message: error?.message || `Permintaan ${action} gagal. Periksa koneksi internet.`,
    action
  });
  const api = async (action, payload = {}, options = {}) => {
    const reading = READ_ACTIONS.has(action) && !(action==='tasks.comments' && payload?.mode==='add');
    const deduping = reading && options.dedupe !== false;
    const key = reading ? cachePrefix(action,payload) : '';
    const ttl = CACHE_TTL[action] || 0;
    if(reading && ttl && !options.fresh){
      const entry=readCache.get(key);
      if(entry && Date.now()-entry.time<ttl){
        console.debug(`[EduTask • Cache] ${action} instant`);
        return entry.data;
      }
    }
    if (deduping && inFlight.has(key)) return inFlight.get(key);
    const execute = async () => {
      const retries = reading ? 1 : 0;
      for (let attempt = 0; attempt <= retries; attempt++) {
        const controller = new AbortController();
        const apiStart = performance.now();
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
          markApi(action,performance.now()-apiStart,response.ok && !!result.success, Number(result.serverMs), result);
          if (!response.ok || !result.success) {
            const error = { code: result.code || 'API_ERROR', message: result.message || `HTTP ${response.status}` };
            if (error.code === 'UNAUTHORIZED' || error.code === 'SESSION_EXPIRED') {
              setSession('');
              window.dispatchEvent(new CustomEvent('edutask:unauthorized'));
            }
            throw error;
          }
          if(reading && ttl) readCache.set(key,{action,data:result.data,time:Date.now()});
          if(!reading && !['auth.login','auth.hashParams','auth.session','auth.forgotPassword'].includes(action)) invalidateReads();
          return result.data;
        } catch (error) {
          if(error.code==='NETWORK_ERROR'||error.name==='TypeError'||error.name==='AbortError') markApi(action,performance.now()-apiStart,false,null);
          if (attempt < retries && ['AbortError','TypeError'].includes(error?.name)) { await sleep(400 * (attempt + 1)); continue; }
          throw normalizeError(error, action);
        } finally { clearTimeout(timeout); }
      }
    };
    const promise = execute();
    if (deduping) inFlight.set(key, promise);
    try { return await promise; } finally { if (deduping) inFlight.delete(key); }
  };
  function setSession(token, user = null) {
    state.token = token || '';
    state.user = user;
    if (token) { sessionStorage.setItem(TOKEN_KEY, token); if(user) sessionStorage.setItem(USER_HINT_KEY, JSON.stringify({id:user.id,fullName:user.fullName,role:user.role})); }
    else { sessionStorage.removeItem(TOKEN_KEY); sessionStorage.removeItem(USER_HINT_KEY); }
    inFlight.clear();
    readCache.clear();
  }
  function idempotencyKey() {
    if (crypto?.randomUUID) return crypto.randomUUID();
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    return [...bytes].map(b => b.toString(16).padStart(2, '0')).join('');
  }
  window.Edutask = Object.freeze({ api, state, setSession, idempotencyKey, invalidateReads, API_URL, TOKEN_KEY });
})();
