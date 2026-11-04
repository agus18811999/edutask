/**
 * EduTask V3 Vercel proxy — accelerated PBKDF2 with signed requests to Apps Script.
 * Secrets stay server-side. Browser sends the original action/payload as before.
 * Requires EDUTASK_GAS_URL, EDUTASK_KDF_PEPPER and EDUTASK_PROXY_SIGNING_KEY.
 */
import {createHmac, randomBytes, randomUUID, pbkdf2 as pbkdf2Callback} from 'node:crypto';
import {promisify} from 'node:util';
const pbkdf2=promisify(pbkdf2Callback);
const UPGRADE_ITERATIONS=600000;

function error(code,message,status=400){return {status,json:{success:false,data:null,code,message}};}
function validatePassword(p,confirm){
  if(typeof p!=='string'||p.length<12||p.length>128) throw Object.assign(new Error('Password harus 12–128 karakter.'),{code:'WEAK_PASSWORD'});
  if(confirm!==undefined&&p!==confirm) throw Object.assign(new Error('Konfirmasi password tidak cocok.'),{code:'VALIDATION_ERROR'});
}
function signedBody(input,key){
  const ts=String(Date.now()),nonce=randomUUID(),body=JSON.stringify(input);
  const sig=createHmac('sha256',key).update(ts+'.'+nonce+'.'+body).digest('hex');
  return JSON.stringify({...input,_edge:{ts,nonce,sig}});
}
async function upstream(input,gasUrl,signingKey,sign=false){
  const raw=sign?signedBody(input,signingKey):JSON.stringify(input);
  const signal=AbortSignal.timeout(28000);
  const response=await fetch(gasUrl,{method:'POST',headers:{'Content-Type':'text/plain; charset=utf-8'},body:raw,redirect:'follow',signal,cache:'no-store'});
  const responseText=await response.text();
  let parsed;
  try{parsed=JSON.parse(responseText);}catch(_){throw Object.assign(new Error('Google Apps Script tidak mengembalikan JSON. Periksa URL /exec.'),{code:'GAS_RESPONSE_ERROR'});}
  return parsed;
}
async function derive(password,salt,iterations,pepper){
  if(!/^[a-f0-9]{64}$/i.test(salt)||!Number.isSafeInteger(Number(iterations))||Number(iterations)<12000||Number(iterations)>1000000)
    throw Object.assign(new Error('Parameter derivasi password tidak valid.'),{code:'HASH_PARAMS_ERROR'});
  const hash=await pbkdf2((password+pepper).trim(),Buffer.from(salt,'hex'),Number(iterations),32,'sha256');
  return {hash:hash.toString('hex'),salt,iterations:Number(iterations)};
}
const secureInputs=new Set(['auth.login','auth.register','auth.resetPassword','auth.changePassword']);
export default async function handler(req,res){
  const started=Date.now();
  res.setHeader('Cache-Control','no-store, max-age=0');
  res.setHeader('X-Content-Type-Options','nosniff');
  res.setHeader('Content-Type','application/json; charset=utf-8');
  if(req.method!=='POST')return res.status(405).json(error('METHOD_NOT_ALLOWED','POST only.',405).json);
  const gasUrl=process.env.EDUTASK_GAS_URL||'';
  if(!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(gasUrl))
    return res.status(503).json(error('CONFIG_ERROR','EDUTASK_GAS_URL harus berisi URL Google Apps Script /exec.',503).json);
  try{
    const origin=req.headers.origin;
    if(origin&&new URL(origin).host!==req.headers.host)return res.status(403).json(error('FORBIDDEN','Cross-origin API access denied.',403).json);
    const input=typeof req.body==='string'?JSON.parse(req.body):req.body;
    if(!input||typeof input!=='object'||Array.isArray(input))return res.status(400).json(error('VALIDATION_ERROR','Payload tidak valid.').json);
    const bodyLength=Buffer.byteLength(JSON.stringify(input),'utf8');
    if(typeof input.action!=='string'||!/^[a-z][a-zA-Z0-9.]{0,99}$/.test(input.action))
      return res.status(400).json(error('VALIDATION_ERROR','API action tidak valid.').json);
    if(bodyLength>30000)return res.status(413).json(error('TOO_LARGE','Payload terlalu besar.',413).json);
    const action=String(input.action||'');
    const token=typeof input.token==='string'?input.token:'';
    const p=input.payload&&typeof input.payload==='object'&&!Array.isArray(input.payload)?input.payload:{};
    const signingKey=process.env.EDUTASK_PROXY_SIGNING_KEY||'';
    const pepper=process.env.EDUTASK_KDF_PEPPER||'';
    const fastReady=signingKey.length>=32&&pepper.length>=32;
    if(secureInputs.has(action)&&(signingKey||pepper)&&!fastReady)
      return res.status(503).json(error('CONFIG_ERROR','Konfigurasi fast auth belum lengkap di Vercel.',503).json);
    let updated={...input};
    let fast=false, lookupMs=0, kdfMs=0;
    if(fastReady&&secureInputs.has(action)){
      fast=true;
      if(action==='auth.login'||action==='auth.changePassword'){
        const t=Date.now();
        const hashParamRequest={action:'auth.hashParams',payload:action==='auth.login'?{email:p.email}:{},token};
        const meta=await upstream(hashParamRequest,gasUrl,signingKey,true);
        lookupMs=Date.now()-t;
        if(!meta.success)return res.status(200).json(meta);
        const salt=meta.data?.salt,iterations=Number(meta.data?.iterations);
        if(action==='auth.login'){
          if(typeof p.password!=='string'||p.password.length<1||p.password.length>128)
            return res.status(200).json(error('INVALID_CREDENTIALS','Email atau password salah.').json);
          const start=Date.now(),proof=await derive(p.password,salt,iterations,pepper);kdfMs=Date.now()-start;
          updated.payload={email:p.email,fastProof:proof};
        }else{
          validatePassword(p.newPassword);
          const t2=Date.now(),current=await derive(String(p.currentPassword||''),salt,iterations,pepper);
          const fresh=await derive(p.newPassword,randomBytes(32).toString('hex'),UPGRADE_ITERATIONS,pepper);
          kdfMs=Date.now()-t2;
          updated.payload={currentProof:current,fastProof:fresh};
        }
      }else{
        validatePassword(p.password,p.confirmPassword);
        const t2=Date.now();
        const proof=await derive(p.password,randomBytes(32).toString('hex'),UPGRADE_ITERATIONS,pepper);
        kdfMs=Date.now()-t2;
        const {password,confirmPassword,...rest}=p;
        updated.payload={...rest,fastProof:proof};
      }
    }
    const gasStart=Date.now();
    const result=await upstream(updated,gasUrl,signingKey,fast);
    const gasRoundTripMs=Date.now()-gasStart, totalMs=Date.now()-started;
    res.setHeader('Server-Timing',`edutask_gas_request;dur=${gasRoundTripMs}, edutask_kdf;dur=${kdfMs}, edutask_lookup;dur=${lookupMs}`);
    console.info('[EduTask Proxy]',action,{fastAuth:fast,hashLookupMs:lookupMs,kdfMs,gasRoundTripMs,totalMs});
    // Response is still the same shape that config.js already understands.
    return res.status(200).json({...result,proxyMs:totalMs,hashLookupMs:lookupMs,kdfMs,gasRoundTripMs});
  }catch(e){
    console.error('[EduTask Proxy]',e?.code||e?.name||'ERROR',String(e?.message||e));
    const known=['WEAK_PASSWORD','VALIDATION_ERROR','HASH_PARAMS_ERROR','GAS_RESPONSE_ERROR'].includes(e?.code);
    return res.status(known?400:502).json(error(known?e.code:'UPSTREAM_UNAVAILABLE',known?e.message:'Tidak dapat memproses permintaan. Periksa deployment atau koneksi backend.',known?400:502).json);
  }
}
