/* EduTask V6 | instant shell, account-scoped safe summaries, targeted cache invalidation. */
(() => {
  'use strict';
  const API_URL='/api/proxy';
  const TOKEN_KEY='edutask_session_v1', USER_HINT_KEY='edutask_user_hint_v2';
  const SCOPE_PREFIX='edutask:v6:summary:';
  // Every record or user detail stays only in RAM; never persist token or sensitive rows.
  const TTL={
    'directory.snapshot':90000,'references.snapshot':180000,'records.view':75000,
    'users.list':90000,'projects.list':90000,'batches.list':90000,
    'tasks.list':30000,'tasks.get':20000,'notifications.list':30000,
    'evaluations.list':45000,'records.list':75000,'reports.intern':45000,
    'reports.performance':45000,'dashboard.mentor':30000,'dashboard.intern':30000,
    'dashboard.refresh':30000,'dashboard.metrics':30000
  };
  const READ_ACTIONS=new Set(['health','auth.session','dashboard.intern','dashboard.mentor','dashboard.refresh','dashboard.metrics',
    'tasks.list','tasks.get','tasks.activity','records.list','records.view','references.snapshot',
    'directory.snapshot','users.list','users.get','projects.list','batches.list','notifications.list',
    'evaluations.list','reports.tasks','reports.intern','reports.performance','tasks.comments']);
  const readCache=new Map(), inFlight=new Map(), perfSamples=[];
  const readHint=()=>{try{return JSON.parse(sessionStorage.getItem(USER_HINT_KEY)||'null')}catch(_){return null}};
  const state={token:sessionStorage.getItem(TOKEN_KEY)||'',user:readHint()};
  const keyFor=(action,payload)=>JSON.stringify([state.user?.id||'unknown',state.user?.role||'unknown',action,payload]);
  const sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
  function perf(action,elapsed,success,res={}){
    const item={action,roundtripMs:Math.round(elapsed),backendMs:Number.isFinite(Number(res.serverMs))?Number(res.serverMs):null,
      proxyMs:Number.isFinite(Number(res.proxyMs))?Number(res.proxyMs):null,
      kdfMs:Number.isFinite(Number(res.kdfMs))?Number(res.kdfMs):null,ok:success,time:new Date().toLocaleTimeString('id-ID')};
    perfSamples.push(item);if(perfSamples.length>120)perfSamples.shift();
    console.info(`⚡ [EduTask API] ${action}: ${item.roundtripMs}ms${item.backendMs===null?'':` • GAS ${item.backendMs}ms`} • ${success?'OK':'FAIL'}`);
  }
  function report(){console.group('⚡ EduTask Instant V6 Performance');console.table(perfSamples);
    const n=performance.getEntriesByType?.('navigation')?.[0];if(n)console.table({domReadyMs:Math.round(n.domContentLoadedEventEnd-n.startTime),loadMs:Math.round(n.loadEventEnd-n.startTime),ttfbMs:Math.round(n.responseStart-n.requestStart)});
    console.groupEnd();return perfSamples.slice();}
  window.EdutaskPerf=Object.freeze({report,samples:()=>perfSamples.slice()});
  // Only numeric aggregates are eligible for disk persistence; no names, emails, tasks or IDs in payload.
  function summaryKey(user=state.user){return user?.id&&user?.role?SCOPE_PREFIX+encodeURIComponent(user.id)+':'+user.role:null;}
  function toSafeSummary(d){
    const valid=(x)=>Number.isFinite(Number(x))?Math.max(0,Number(x)):0;
    const fields=['total','todo','inProgress','inReview','revision','completed','blocked','overdue','completionRate'];
    const stats={};for(const field of fields)stats[field]=valid(d?.stats?.[field]);
    return {stats,projects:valid(d?.projects),interns:valid(d?.interns),
      weekly:(Array.isArray(d?.weekly)?d.weekly:[]).slice(-7).map(x=>({date:/^\d{4}-\d{2}-\d{2}$/.test(x.date)?x.date:'',count:valid(x.count)})),
      earliest:[],notifications:[],activities:[],upcoming:[],latestSubmissions:[]};
  }
  const scoped={
    get(){try{const k=summaryKey();if(!k)return null;const raw=localStorage.getItem(k);if(!raw)return null;
      const value=JSON.parse(raw);if(value.version!==6||Date.now()-value.timestamp>600000||value.timestamp>Date.now()+30000){localStorage.removeItem(k);return null;}
      return {data:toSafeSummary(value.data),timestamp:value.timestamp};
    }catch(_){return null}},
    put(data){try{const k=summaryKey();if(!k)return;localStorage.setItem(k,JSON.stringify({version:6,timestamp:Date.now(),data:toSafeSummary(data)}));}catch(_){}},
    clear(){try{const k=summaryKey();if(k)localStorage.removeItem(k)}catch(_){}}
  };
  function peek(action,payload={},options={}){const row=readCache.get(keyFor(action,payload));if(!row)return null;
    if(!options.allowStale&&Date.now()-row.time>Number(TTL[action]||0))return null;
    return {data:row.data,ageMs:Date.now()-row.time};}
  function invalidateReads(actions=null,scopePredicate=null){
    if(!actions){readCache.clear();return;}
    const filter=new Set(actions);
    for(const [key,row] of readCache)if(filter.has(row.action)&&(!scopePredicate||scopePredicate(row.payload)))readCache.delete(key);
  }
  function invalidateRelated(action,payload={}){
    if(action.startsWith('auth.')&&action!=='auth.changePassword')return;
    const affected=new Set();
    const add=(...x)=>x.forEach(y=>affected.add(y));
    if(action.startsWith('tasks.')||payload.table==='Subtasks')add('tasks.list','tasks.get','dashboard.refresh','dashboard.intern','dashboard.mentor','dashboard.metrics','reports.intern','reports.performance','reports.tasks','notifications.list');
    if(action.startsWith('evaluations.')||payload.table==='Evaluations')add('evaluations.list','reports.performance','dashboard.refresh');
    if(action.startsWith('notifications.'))add('notifications.list','dashboard.refresh','dashboard.intern','dashboard.mentor','dashboard.metrics');
    if(action.startsWith('users.')||payload.table==='MentorAssignments')add('users.list','users.get','directory.snapshot','references.snapshot','dashboard.refresh','dashboard.intern','dashboard.mentor','dashboard.metrics');
    if(payload.table){add('records.view','records.list');if(['Projects','InternshipPrograms','Batches','Teams','ProjectMembers'].includes(payload.table))add('references.snapshot','directory.snapshot');
      if(['Projects','ProjectMembers'].includes(payload.table))add('projects.list','dashboard.refresh');
      if(payload.table==='Batches')add('batches.list');}
    if(action.startsWith('projects.')||action.startsWith('batches.'))add('records.view','records.list','references.snapshot','directory.snapshot','dashboard.refresh','projects.list','batches.list');
    if(affected.size)invalidateReads([...affected],p=> !payload.table || !p?.table || p.table===payload.table);
  }
  function normalizeError(error,action){return {code:error?.code||(error?.name==='AbortError'?'TIMEOUT':'NETWORK_ERROR'),message:error?.message||`Permintaan ${action} gagal.`,action};}
  async function api(action,payload={},options={}){
    const reading=READ_ACTIONS.has(action)&&!(action==='tasks.comments'&&payload?.mode==='add');
    const key=keyFor(action,payload),ttl=TTL[action]||0,requestToken=state.token;
    if(reading&&!options.fresh){const saved=readCache.get(key);if(saved&&Date.now()-saved.time<ttl){console.debug(`⚡ [EduTask Cache] ${action} RAM ${Date.now()-saved.time}ms`);return saved.data;}}
    if(reading&&options.dedupe!==false&&inFlight.has(key))return inFlight.get(key);
    const execute=async()=>{
      const retries=reading?1:0;
      for(let attempt=0;attempt<=retries;attempt++){
        const controller=new AbortController(),start=performance.now();const timeout=setTimeout(()=>controller.abort(),options.timeout||25000);
        try{
          const response=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'application/json','Accept':'application/json'},
            credentials:'same-origin',cache:'no-store',signal:controller.signal,
            body:JSON.stringify({action,payload,token:requestToken,requestId:options.requestId||undefined})});
          let result;try{result=JSON.parse(await response.text())}catch(_){throw {code:'INVALID_RESPONSE',message:'Respons server bukan JSON yang valid.'}}
          perf(action,performance.now()-start,response.ok&&!!result.success,result);
          if(!response.ok||!result.success){const error={code:result.code||'API_ERROR',message:result.message||`HTTP ${response.status}`};
            if((error.code==='UNAUTHORIZED'||error.code==='SESSION_EXPIRED')&&requestToken===state.token){setSession('');window.dispatchEvent(new CustomEvent('edutask:unauthorized'));}
            throw error;}
          // Reject late results from a previous account instead of hydrating another account's cache.
          if(requestToken!==state.token)throw {code:'SESSION_CHANGED',message:'Sesi pengguna sudah berubah.'};
          if(reading&&ttl)readCache.set(key,{action,payload,data:result.data,time:Date.now()});
          if(!reading)invalidateRelated(action,payload);
          return result.data;
        }catch(error){
          if(attempt<retries&&['AbortError','TypeError'].includes(error?.name)){await sleep(300*(attempt+1));continue;}
          throw normalizeError(error,action);
        }finally{clearTimeout(timeout)}
      }
    };
    const promise=execute();if(reading&&options.dedupe!==false)inFlight.set(key,promise);
    try{return await promise}finally{if(reading&&inFlight.get(key)===promise)inFlight.delete(key)}
  }
  function setSession(token,user=null){
    if(!token)scoped.clear();
    state.token=token||'';state.user=user;
    if(token){sessionStorage.setItem(TOKEN_KEY,token);if(user)sessionStorage.setItem(USER_HINT_KEY,JSON.stringify({id:user.id,fullName:user.fullName,role:user.role}));}
    else{sessionStorage.removeItem(TOKEN_KEY);sessionStorage.removeItem(USER_HINT_KEY);}
    inFlight.clear();readCache.clear();
  }
  function idempotencyKey(){if(crypto?.randomUUID)return crypto.randomUUID();return [...crypto.getRandomValues(new Uint8Array(16))].map(x=>x.toString(16).padStart(2,'0')).join('')}
  window.Edutask=Object.freeze({api,state,setSession,idempotencyKey,invalidateReads,invalidateRelated,peek,scoped,API_URL,TOKEN_KEY});
})();
