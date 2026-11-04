/**
 * EduTask | Google Apps Script V8 backend.
 * Deploy as: Execute as Me / Anyone. All application authorization is server-side.
 * Custom credentials on Apps Script are a tradeoff; read SECURITY in README.
 */
const EDUTASK = (() => {
  'use strict';
  const COLS = Object.freeze({
    Users: 'email,fullName,passwordHash,passwordSalt,iterations,role,emailVerified,programId,batchId,teamId,mentorId,avatarUrl,phone,joinDate,lastLogin,preferences',
    Invitations: 'email,role,codeHash,programId,batchId,teamId,mentorId,expiresAt,usedAt,createdBy',
    EmailVerifications: 'userId,codeHash,expiresAt,attempts,resendAt,usedAt',
    PasswordResets: 'userId,codeHash,expiresAt,attempts,usedAt',
    Sessions: 'userId,tokenHash,expiresAt,revokedAt,lastSeenAt',
    InternshipPrograms: 'name,description,startDate,endDate,ownerId',
    Batches: 'name,programId,startDate,endDate,description,mentorId,ownerId',
    Teams: 'name,description,batchId,mentorId,ownerId',
    MentorAssignments: 'mentorId,internId,projectId,ownerId',
    InternProfiles: 'userId,objectives,learningGoals,skills,achievements,bio,linkedin,github,ownerId',
    Projects: 'name,description,mentorId,batchId,programId,startDate,dueDate,priority,progress,ownerId',
    ProjectMembers: 'projectId,userId,ownerId',
    Milestones: 'projectId,name,description,dueDate,progress,ownerId',
    Tasks: 'title,description,instructions,projectId,category,priority,startDate,dueDate,estimateHours,checklist,links,notes,mentorId,createdBy,requestId',
    TaskAssignments: 'taskId,internId,mentorId,progress,startedAt,submittedAt,completedAt,revisions,ownerId',
    Subtasks: 'assignmentId,taskId,title,done,ownerId',
    TaskTemplates: 'title,description,category,priority,estimateHours,checklist,instructions,ownerId',
    TaskComments: 'assignmentId,taskId,userId,message,ownerId',
    TaskSubmissions: 'assignmentId,taskId,internId,link,notes,requestId,version,ownerId',
    TaskReviews: 'submissionId,assignmentId,mentorId,decision,feedback,score,ownerId',
    TaskActivity: 'assignmentId,taskId,userId,action,details,ownerId',
    DailyReports: 'userId,date,type,goal,priorities,completed,inProgress,blockers,learnings,tomorrow,durationMinutes,ownerId',
    Evaluations: 'internId,mentorId,period,quality,timeliness,communication,initiative,collaboration,problemSolving,ownership,learning,notes,published,ownerId',
    EvaluationCriteria: 'name,weight,description,ownerId',
    Notifications: 'userId,title,message,link,type,isRead,ownerId',
    Announcements: 'title,content,audience,programId,batchId,publishAt,ownerId',
    AuditLogs: 'actorId,entity,entityId,action,detail,ownerId',
    AppSettings: 'key,value,ownerId'
  });
  const TABLES = Object.keys(COLS);
  const PUBLIC = new Set(['health','auth.register','auth.login','auth.verifyEmail','auth.resendVerification','auth.forgotPassword','auth.resetPassword']);
  const ADMIN_EDIT = new Set(['InternshipPrograms','Batches','Teams','MentorAssignments','ProjectMembers','EvaluationCriteria','AppSettings']);
  const STAFF_EDIT = new Set(['Projects','Milestones','TaskTemplates','Announcements']);
  const INTERN_EDIT = new Set(['DailyReports','InternProfiles']);
  const DEFAULT_STATUS = 'Active';
  const iso = () => new Date().toISOString();
  const str = v => String(v === null || v === undefined ? '' : v).trim();
  const lower = v => str(v).toLowerCase();
  const uid = prefix => prefix + '_' + Utilities.getUuid().replace(/-/g, '');
  const props = () => PropertiesService.getScriptProperties();
  const requireField = (val, label) => { if (!str(val)) fail('VALIDATION_ERROR', label + ' wajib diisi.'); return str(val); };
  const fail = (code, message) => { const e = new Error(message); e.code = code; throw e; };
  const limit = (val, max = 5000) => { const s = str(val); if (s.length > max) fail('VALIDATION_ERROR','Teks terlalu panjang.'); return s; };
  const validEmail = email => { const v = lower(email); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v) || v.length > 254) fail('VALIDATION_ERROR','Alamat email tidak valid.'); return v; };
  const dateOK = x => !x || /^\d{4}-\d{2}-\d{2}$/.test(str(x));
  const safeURL = u => { if (!u) return ''; if (!/^https:\/\/[^\s]+$/i.test(str(u)) || str(u).length > 2000) fail('VALIDATION_ERROR','Gunakan URL HTTPS yang valid.'); return str(u); };
  const safeSheet = value => {
    let s = typeof value === 'object' && value !== null ? JSON.stringify(value) : String(value === undefined || value === null ? '' : value);
    if (s.length > 20000) fail('VALIDATION_ERROR','Nilai field terlalu panjang.');
    // Sheets and CSV formula injection: apostrophe escape for user-controlled text.
    if (/^[\t\r\n\s]*[=+\-@]/.test(s)) s = "'" + s;
    return s;
  };
  const schema = name => {
    if (!Object.prototype.hasOwnProperty.call(COLS, name)) fail('BAD_TABLE','Tabel tidak dikenal.');
    return ['id', ...COLS[name].split(','), 'status', 'createdAt', 'updatedAt'];
  };
  let spreadsheet;
  const ss = () => {
    if (spreadsheet) return spreadsheet;
    const id = props().getProperty('EDUTASK_SPREADSHEET_ID');
    if (!id) fail('CONFIG_ERROR','EDUTASK_SPREADSHEET_ID belum diatur pada Script Properties.');
    spreadsheet = SpreadsheetApp.openById(id);
    return spreadsheet;
  };
  const sheet = name => { schema(name); const s = ss().getSheetByName(name); if (!s) fail('SETUP_REQUIRED','Jalankan setupEdutask() dahulu.'); return s; };
  const all = (name, ctx) => {
    if (ctx && ctx.cache && ctx.cache[name]) return ctx.cache[name];
    const s = sheet(name), n = s.getLastRow();
    if (n < 2) { if (ctx) ctx.cache[name] = []; return []; }
    const heads = schema(name), raw = s.getRange(2, 1, n - 1, heads.length).getDisplayValues();
    const data = raw.map((r, i) => { const ob = { _row: i + 2 }; heads.forEach((h,j) => ob[h] = r[j] || ''); return ob; }).filter(x => x.id);
    if (ctx) ctx.cache[name] = data;
    return data;
  };
  const byId = (name, id, ctx) => all(name, ctx).find(r => r.id === str(id)) || null;
  const active = r => r && r.status !== 'Archived' && r.status !== 'Deleted' && r.status !== 'Inactive';
  const dbInsert = (name, data, ctx) => {
    const heads = schema(name), now = iso();
    const row = { ...data, id: data.id || uid(name.substring(0,3).toLowerCase()), status: data.status || DEFAULT_STATUS, createdAt: now, updatedAt: now };
    const target = sheet(name);
    const nextRow = target.getLastRow() + 1;
    const range = target.getRange(nextRow, 1, 1, heads.length);
    range.setNumberFormat('@'); // Preserve ISO dates, IDs and tokens as literal text.
    range.setValues([heads.map(h => safeSheet(row[h]))]);
    if (ctx && ctx.cache) delete ctx.cache[name];
    return row;
  };
  const dbUpdate = (name, id, changes, ctx) => {
    const old = byId(name, id, ctx);
    if (!old) fail('NOT_FOUND','Record tidak ditemukan.');
    const row = { ...old, ...changes, id: old.id, createdAt: old.createdAt, updatedAt: iso() };
    const heads = schema(name);
    const range = sheet(name).getRange(old._row, 1, 1, heads.length);
    range.setNumberFormat('@');
    range.setValues([heads.map(h => safeSheet(row[h]))]);
    if (ctx && ctx.cache) delete ctx.cache[name];
    return row;
  };
  const locked = fn => {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) fail('BUSY','Sistem sedang sibuk. Coba lagi sebentar.');
    try { return fn(); } finally { lock.releaseLock(); }
  };
  const hex = values => values.map(x => (x & 255).toString(16).padStart(2, '0')).join('');
  const sha = value => hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value));
  const pepper = () => { const p = props().getProperty('EDUTASK_SECRET'); if (!p) fail('SETUP_REQUIRED','Server secret belum dibuat.'); return p; };
  const tokenHash = val => sha(str(val) + ':' + pepper());
  const constantEquals = (a, b) => {
    const x = str(a), y = str(b); let diff = x.length ^ y.length;
    const n = Math.max(x.length, y.length);
    for (let i=0;i<n;i++) diff |= (x.charCodeAt(i) || 0) ^ (y.charCodeAt(i) || 0);
    return diff === 0;
  };
  function pbkdf2(password, saltHex, iterations) {
    // PBKDF2-HMAC-SHA256, one 32-byte block. GAS Utilities HMAC is significantly slower than native crypto.
    const p = str(password), it = Number(iterations);
    const salt = (saltHex.match(/.{2}/g) || []).map(b => parseInt(b,16)).concat([0,0,0,1]);
    const key = Utilities.newBlob(p).getBytes();
    let u = Utilities.computeHmacSha256Signature(salt, key);
    const out = u.map(n => n & 255);
    for (let i = 1; i < it; i++) {
      u = Utilities.computeHmacSha256Signature(u, key);
      for (let j=0;j<32;j++) out[j] ^= (u[j] & 255);
    }
    return hex(out);
  }
  function passwordParts(password) {
    if (str(password).length < 12 || str(password).length > 128) fail('WEAK_PASSWORD','Password harus 12–128 karakter.');
    const iterations = Number(props().getProperty('EDUTASK_PBKDF2_ITERATIONS') || '12000');
    const salt = sha(Utilities.getUuid() + Utilities.getUuid());
    return { passwordSalt: salt, iterations, passwordHash: pbkdf2(password + pepper(), salt, iterations) };
  }
  const checkPassword = (password, user) => Boolean(user && user.passwordSalt && constantEquals(pbkdf2(password + pepper(), user.passwordSalt, user.iterations), user.passwordHash));
  const publicUser = user => ({ id:user.id, fullName:user.fullName, email:user.email, role:user.role, avatarUrl:user.avatarUrl || '', phone:user.phone || '', programId:user.programId || '', batchId:user.batchId || '', teamId:user.teamId || '', mentorId:user.mentorId || '' });
  const audit = (ctx, entity, entityId, action, detail = '') => dbInsert('AuditLogs', {
    actorId: ctx?.user?.id || 'system', entity, entityId, action, detail: limit(typeof detail === 'string' ? detail : JSON.stringify(detail), 1500), ownerId: ctx?.user?.id || 'system'
  }, ctx);
  function rate(name, value, maximum, seconds) {
    const cache = CacheService.getScriptCache(), key = 'rate:' + sha(name + ':' + lower(value)).slice(0,55);
    const count = Number(cache.get(key) || 0);
    if (count >= maximum) fail('RATE_LIMIT','Terlalu banyak percobaan. Coba beberapa saat lagi.');
    cache.put(key, String(count + 1), seconds);
  }
  function mail(to, subject, text) { MailApp.sendEmail({ to, subject, body: text, name: 'EduTask' }); }
  const code = () => String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000).padStart(6, '0');
  function createCode(name, idField, idValue, ctx) {
    const now = Date.now(), value = code();
    const pending = all(name, ctx).filter(x => x[idField] === idValue && !x.usedAt);
    pending.forEach(x => dbUpdate(name,x.id,{ usedAt:iso(),status:'Used' },ctx));
    dbInsert(name, { [idField]:idValue, codeHash:tokenHash(value), expiresAt:new Date(now + 15*60000).toISOString(), attempts:'0', resendAt:new Date(now + 60000).toISOString() },ctx);
    return value;
  }
  function confirmCode(name, idField, idValue, value, ctx) {
    const r = all(name,ctx).filter(x => x[idField] === idValue && !x.usedAt).sort((a,b) => b.createdAt.localeCompare(a.createdAt))[0];
    if (!r || Date.parse(r.expiresAt) < Date.now()) fail('OTP_EXPIRED','Kode kedaluwarsa. Minta kode baru.');
    if (Number(r.attempts) >= 5) fail('OTP_LOCKED','Terlalu banyak percobaan kode.');
    dbUpdate(name,r.id,{ attempts:String(Number(r.attempts)+1) },ctx);
    if (!constantEquals(r.codeHash,tokenHash(value))) fail('OTP_INVALID','Kode verifikasi salah.');
    dbUpdate(name,r.id,{ usedAt:iso(),status:'Used' },ctx);
    return true;
  }
  function authenticate(token, ctx) {
    if (!token || str(token).length < 40) fail('UNAUTHORIZED','Silakan login terlebih dahulu.');
    const session = all('Sessions',ctx).find(s => s.tokenHash === tokenHash(token) && !s.revokedAt);
    if (!session || Date.parse(session.expiresAt) <= Date.now()) fail('SESSION_EXPIRED','Sesi berakhir. Silakan login lagi.');
    const user = byId('Users',session.userId,ctx);
    if (!active(user) || user.emailVerified !== 'true') fail('UNAUTHORIZED','Akun belum aktif atau belum diverifikasi.');
    ctx.user = user;
    return user;
  }
  const isStaff = ctx => ['admin','mentor'].includes(ctx.user.role);
  const isAdmin = ctx => ctx.user.role === 'admin';
  const requireStaff = ctx => { if (!isStaff(ctx)) fail('FORBIDDEN','Hanya admin dan mentor yang dapat mengakses fitur ini.'); };
  const requireAdmin = ctx => { if (!isAdmin(ctx)) fail('FORBIDDEN','Hanya admin yang dapat mengakses fitur ini.'); };
  function allowedIntern(ctx, internId) {
    if (isAdmin(ctx) || ctx.user.id === internId) return true;
    if (ctx.user.role !== 'mentor') return false;
    const intern = byId('Users',internId,ctx);
    return Boolean(intern && (intern.mentorId === ctx.user.id || all('MentorAssignments',ctx).some(m => active(m) && m.mentorId === ctx.user.id && m.internId === internId)));
  }
  const staffIntern = (ctx, id) => { requireStaff(ctx); if (!allowedIntern(ctx,id) || id === ctx.user.id) fail('FORBIDDEN','Peserta berada di luar kewenangan Anda.'); };
  function projectAllowed(ctx, id) {
    const project = byId('Projects',id,ctx);
    if (!project) return false;
    if (isAdmin(ctx) || project.mentorId === ctx.user.id) return true;
    return all('ProjectMembers',ctx).some(x => active(x) && x.projectId === id && x.userId === ctx.user.id);
  }
  function assignmentAllowed(ctx, id) {
    const a = byId('TaskAssignments',id,ctx);
    if (!a || !active(a)) return false;
    return allowedIntern(ctx,a.internId);
  }
  function canSee(ctx,name,r) {
    if (!active(r) && name !== 'TaskAssignments' && name !== 'Sessions') return false;
    if (isAdmin(ctx)) return true;
    const me = ctx.user.id;
    if (name === 'Users') return r.id === me || (ctx.user.role === 'mentor' && r.role === 'intern' && allowedIntern(ctx,r.id));
    if (name === 'InternProfiles' || name === 'DailyReports' || name === 'Evaluations') return allowedIntern(ctx,r.userId || r.internId) && (name !== 'Evaluations' || isStaff(ctx) || r.published === 'true');
    if (name === 'Notifications') return r.userId === me;
    if (name === 'Announcements') return (isStaff(ctx) && r.ownerId === me) || ((!r.publishAt || Date.parse(r.publishAt) <= Date.now()) && (r.audience === 'all' || r.audience === ctx.user.role || !r.audience) && (!r.batchId || r.batchId === ctx.user.batchId));
    if (name === 'Tasks') return isStaff(ctx) ? (r.mentorId === me || all('TaskAssignments',ctx).some(a => a.taskId === r.id && allowedIntern(ctx,a.internId))) : all('TaskAssignments',ctx).some(a => a.taskId === r.id && a.internId === me && active(a));
    if (name === 'TaskAssignments') return allowedIntern(ctx,r.internId);
    if (name === 'TaskSubmissions') return assignmentAllowed(ctx,r.assignmentId);
    if (name === 'TaskReviews' || name === 'TaskComments' || name === 'TaskActivity' || name === 'Subtasks') return assignmentAllowed(ctx,r.assignmentId);
    if (name === 'Projects') return projectAllowed(ctx,r.id);
    if (name === 'ProjectMembers' || name === 'Milestones') return projectAllowed(ctx,r.projectId);
    if (name === 'MentorAssignments') return r.mentorId === me || r.internId === me;
    if (name === 'Batches') return isStaff(ctx) ? r.mentorId === me || r.id === ctx.user.batchId : r.id === ctx.user.batchId;
    if (name === 'Teams') return r.mentorId === me || r.id === ctx.user.teamId;
    if (name === 'InternshipPrograms') return r.id === ctx.user.programId;
    if (name === 'TaskTemplates') return isStaff(ctx) && (r.ownerId === me || !r.ownerId);
    if (name === 'EvaluationCriteria') return true;
    if (name === 'AuditLogs') return ctx.user.role === 'mentor' && r.actorId === me;
    return false;
  }
  const visibleList = (name,ctx) => all(name,ctx).filter(r => canSee(ctx,name,r));
  function notify(ctx, userId, title, message, link, type='info') {
    if (!userId) return;
    dbInsert('Notifications',{userId,title,message,link,type,isRead:'false',ownerId:userId},ctx);
  }
  function taskEvent(ctx, a, action, details) {
    dbInsert('TaskActivity',{assignmentId:a.id,taskId:a.taskId,userId:ctx.user.id,action,details:limit(details || '',1200),ownerId:a.internId},ctx);
  }
  const taskByAssignment = (a, ctx) => {
    const t = byId('Tasks',a.taskId,ctx);
    if (!t || !active(t)) return null;
    return { ...t, assignmentId:a.id, internId:a.internId, mentorId:a.mentorId || t.mentorId,
      taskStatus:a.status, progress:Number(a.progress || 0), startedAt:a.startedAt, submittedAt:a.submittedAt,
      completedAt:a.completedAt, revisions:Number(a.revisions || 0), isOverdue:a.status !== 'Completed' && Boolean(t.dueDate) && t.dueDate < new Date().toISOString().slice(0,10) };
  };
  const taskList = ctx => visibleList('TaskAssignments',ctx).map(a => taskByAssignment(a,ctx)).filter(Boolean);
  function dashboard(ctx) {
    const tasks = taskList(ctx);
    const count = status => tasks.filter(t => t.taskStatus === status).length;
    const finished = count('Completed'), overdue = tasks.filter(x=>x.isOverdue).length;
    const stats = { total:tasks.length, todo:count('To Do'), inProgress:count('In Progress'), inReview:count('In Review'), revision:count('Revision Required'), completed:finished, blocked:count('Blocked'), overdue,
      completionRate: tasks.length ? Math.round(finished / tasks.length * 100):0 };
    const earliest = tasks.filter(t=>t.taskStatus !== 'Completed' && t.dueDate).sort((a,b) => a.dueDate.localeCompare(b.dueDate)).slice(0,7);
    const notifications = visibleList('Notifications',ctx).sort((a,b)=>b.createdAt.localeCompare(a.createdAt)).slice(0,8);
    const activities = visibleList('TaskActivity',ctx).sort((a,b)=>b.createdAt.localeCompare(a.createdAt)).slice(0,8);
    const dates = Array.from({length:7},(_,i)=>{ const d=new Date(); d.setUTCDate(d.getUTCDate()-6+i); return d.toISOString().slice(0,10); });
    const weekly = dates.map(d=>({date:d,count:tasks.filter(t=>t.completedAt && t.completedAt.slice(0,10) === d).length}));
    const interns = isStaff(ctx) ? visibleList('Users',ctx).filter(u=>u.role==='intern') : [];
    const projects = visibleList('Projects',ctx);
    return { user: publicUser(ctx.user), stats, earliest, notifications, activities, weekly,
      projects:projects.length, interns:interns.length,
      upcoming: earliest.filter(t=>t.dueDate >= iso().slice(0,10)).slice(0,5),
      latestSubmissions:isStaff(ctx)?visibleList('TaskSubmissions',ctx).sort((a,b)=>b.createdAt.localeCompare(a.createdAt)).slice(0,5):[] };
  }
  function createSession(ctx, user) {
    const token = Array.from({length:4},()=>Utilities.getUuid().replace(/-/g,'')).join('');
    const expiresAt = new Date(Date.now() + 8*3600000).toISOString();
    dbInsert('Sessions', {userId:user.id,tokenHash:tokenHash(token),expiresAt,lastSeenAt:iso()},ctx);
    return { token, expiresAt, user:publicUser(user) };
  }
  function auth(ctx, action, p) {
    if (action === 'auth.register') return locked(()=>{
      rate('register',p.email,5,3600);
      const email=validEmail(p.email), name=limit(requireField(p.fullName,'Nama lengkap'),120);
      if (p.password !== p.confirmPassword) fail('VALIDATION_ERROR','Konfirmasi password tidak cocok.');
      const invitation=all('Invitations',ctx).find(i=>i.email===email && !i.usedAt && i.codeHash===tokenHash(str(p.invitationCode)) && Date.parse(i.expiresAt)>Date.now());
      if (!invitation) fail('INVITE_REQUIRED','Undangan tidak valid atau sudah kedaluwarsa.');
      if (all('Users',ctx).some(u=>u.email===email)) fail('ACCOUNT_EXISTS','Email sudah terdaftar.');
      const parts=passwordParts(p.password);
      const user=dbInsert('Users',{email,fullName:name,...parts,role:invitation.role,emailVerified:'false',programId:invitation.programId,
        batchId:invitation.batchId,teamId:invitation.teamId,mentorId:invitation.mentorId},ctx);
      dbUpdate('Invitations',invitation.id,{usedAt:iso(),status:'Used'},ctx);
      const otp=createCode('EmailVerifications','userId',user.id,ctx);
      mail(email,'Kode verifikasi akun EduTask',`Halo ${name},\n\nKode verifikasi akun EduTask: ${otp}\nBerlaku 15 menit. Jangan bagikan kode kepada siapa pun.\n`);
      audit(ctx,'Users',user.id,'register',email);
      return {email,verifyRequired:true};
    });
    if (action === 'auth.verifyEmail') return locked(()=>{
      const email=validEmail(p.email), user=all('Users',ctx).find(u=>u.email===email);
      if (!user) fail('OTP_INVALID','Kode tidak valid.');
      if (user.emailVerified==='true') return {verified:true};
      rate('verify',email,12,1800);
      confirmCode('EmailVerifications','userId',user.id,str(p.code),ctx);
      dbUpdate('Users',user.id,{emailVerified:'true'},ctx);
      return {verified:true};
    });
    if (action === 'auth.resendVerification') return locked(()=>{
      const email=validEmail(p.email); rate('resend',email,3,3600);
      const user=all('Users',ctx).find(x=>x.email===email);
      if (!user || user.emailVerified==='true') return {message:'Jika akun belum terverifikasi, kode akan dikirim.'};
      const last=all('EmailVerifications',ctx).filter(x=>x.userId===user.id && !x.usedAt).sort((a,b)=>b.createdAt.localeCompare(a.createdAt))[0];
      if (last && Date.parse(last.resendAt)>Date.now()) fail('COOLDOWN','Tunggu 60 detik sebelum mengirim ulang.');
      const otp=createCode('EmailVerifications','userId',user.id,ctx);
      mail(email,'Kode verifikasi EduTask',`Kode terbaru: ${otp}. Berlaku 15 menit.`);
      return {message:'Jika akun belum terverifikasi, kode akan dikirim.'};
    });
    if (action === 'auth.login') return locked(()=>{
      const email=validEmail(p.email); rate('login',email,8,900);
      const user=all('Users',ctx).find(u=>u.email===email);
      // Keep hash computation cost consistent even on an unknown account.
      const valid = user ? checkPassword(str(p.password),user) : false;
      if (!valid || !active(user)) fail('INVALID_CREDENTIALS','Email atau password salah.');
      if (user.emailVerified !== 'true') fail('EMAIL_UNVERIFIED','Silakan verifikasi email terlebih dahulu.');
      dbUpdate('Users',user.id,{lastLogin:iso()},ctx);
      const login=createSession(ctx,user);
      ctx.user=user; audit(ctx,'Users',user.id,'login');
      return login;
    });
    if (action === 'auth.forgotPassword') return locked(()=>{
      const email=validEmail(p.email); rate('forgot',email,3,3600);
      const user=all('Users',ctx).find(u=>u.email===email && active(u));
      if (user) {
        const reset=createCode('PasswordResets','userId',user.id,ctx);
        mail(email,'Reset password EduTask',`Kode reset password Anda: ${reset}. Berlaku 15 menit. Jika Anda tidak meminta reset, abaikan email ini.`);
      }
      return {message:'Jika email terdaftar, kode reset akan dikirim.'};
    });
    if (action === 'auth.resetPassword') return locked(()=>{
      const email=validEmail(p.email); rate('reset',email,10,3600);
      const user=all('Users',ctx).find(u=>u.email===email);
      if (!user) fail('OTP_INVALID','Kode atau email salah.');
      if (p.password !== p.confirmPassword) fail('VALIDATION_ERROR','Konfirmasi password tidak cocok.');
      // Validate password before consuming the single-use token.
      if (str(p.password).length < 12 || str(p.password).length > 128) fail('WEAK_PASSWORD','Password harus 12–128 karakter.');
      confirmCode('PasswordResets','userId',user.id,str(p.code),ctx);
      dbUpdate('Users',user.id,passwordParts(p.password),ctx);
      all('Sessions',ctx).filter(s=>s.userId===user.id && !s.revokedAt).forEach(s=>dbUpdate('Sessions',s.id,{revokedAt:iso()},ctx));
      audit(ctx,'Users',user.id,'resetPassword');
      return {reset:true};
    });
    if (action === 'auth.logout') return locked(()=>{
      const s=all('Sessions',ctx).find(x=>x.tokenHash===tokenHash(ctx.token) && !x.revokedAt);
      if (s) dbUpdate('Sessions',s.id,{revokedAt:iso()},ctx);
      return {loggedOut:true};
    });
    if (action === 'auth.session') return {user:publicUser(ctx.user)};
    if (action === 'auth.changePassword') return locked(()=>{
      if (!checkPassword(str(p.currentPassword),ctx.user)) fail('INVALID_CREDENTIALS','Password lama tidak cocok.');
      const parts=passwordParts(p.newPassword);
      dbUpdate('Users',ctx.user.id,parts,ctx);
      all('Sessions',ctx).filter(s=>s.userId===ctx.user.id && s.tokenHash!==tokenHash(ctx.token) && !s.revokedAt)
        .forEach(s=>dbUpdate('Sessions',s.id,{revokedAt:iso()},ctx));
      audit(ctx,'Users',ctx.user.id,'changePassword');
      return {changed:true};
    });
    fail('UNKNOWN_ACTION','Aksi autentikasi tidak dikenal.');
  }
  function usersApi(ctx,action,p) {
    if (action==='users.list') return visibleList('Users',ctx).map(publicUser);
    if (action==='users.get') { const r=byId('Users',p.id,ctx); if (!r||!canSee(ctx,'Users',r)) fail('NOT_FOUND','Pengguna tidak ditemukan.'); return publicUser(r); }
    if (action==='users.invite') return locked(()=>{
      requireAdmin(ctx);
      const email=validEmail(p.email), role=str(p.role);
      if (!['intern','mentor'].includes(role)) fail('VALIDATION_ERROR','Role tidak valid.');
      if (all('Users',ctx).some(x=>x.email===email)) fail('ACCOUNT_EXISTS','Email sudah terdaftar.');
      rate('invite',email,5,3600);
      const value=Array.from({length:2},()=>Utilities.getUuid().replace(/-/g,'')).join('').slice(0,24).toUpperCase();
      const record=dbInsert('Invitations',{ email, role, codeHash:tokenHash(value), programId:str(p.programId),batchId:str(p.batchId),
        teamId:str(p.teamId),mentorId:str(p.mentorId),createdBy:ctx.user.id,expiresAt:new Date(Date.now()+7*86400000).toISOString() },ctx);
      const url=props().getProperty('EDUTASK_APP_URL') || '';
      mail(email,'Undangan EduTask',`Anda diundang sebagai ${role} di EduTask.\nKode undangan: ${value}\nBerlaku 7 hari.\nBuka: ${url ? url.replace(/\/$/,'') + '/index.html?tab=register&email='+encodeURIComponent(email) : 'website EduTask internal'}\nJangan berikan kode kepada orang lain.`);
      audit(ctx,'Invitations',record.id,'invite',email);
      return {sent:true,email};
    });
    if (action==='users.update') return locked(()=>{
      const target=byId('Users',p.id||ctx.user.id,ctx);
      if (!target || (!isAdmin(ctx) && target.id!==ctx.user.id)) fail('FORBIDDEN','Anda tidak dapat mengubah akun ini.');
      const changes={};
      const self=['fullName','avatarUrl','phone','preferences'];
      const extra=['role','programId','batchId','teamId','mentorId','status'];
      [...self,...(isAdmin(ctx)?extra:[])].forEach(k=>{ if (p[k]!==undefined) changes[k]=limit(p[k],k==='preferences'?1500:250); });
      if (changes.role && !['admin','mentor','intern'].includes(changes.role)) fail('VALIDATION_ERROR','Role tidak valid.');
      if (changes.avatarUrl) changes.avatarUrl=safeURL(changes.avatarUrl);
      if (target.id===ctx.user.id && (changes.role || changes.status==='Inactive')) fail('FORBIDDEN','Tidak dapat mengubah role/status akun sendiri.');
      const updated=dbUpdate('Users',target.id,changes,ctx);
      audit(ctx,'Users',target.id,'update',Object.keys(changes).join(','));
      return publicUser(updated);
    });
    if (action==='users.deactivate') return locked(()=>{
      requireAdmin(ctx);
      if (p.id===ctx.user.id) fail('FORBIDDEN','Tidak dapat menonaktifkan diri sendiri.');
      const result=dbUpdate('Users',p.id,{status:'Inactive'},ctx);
      all('Sessions',ctx).filter(s=>s.userId===p.id && !s.revokedAt).forEach(s=>dbUpdate('Sessions',s.id,{revokedAt:iso()},ctx));
      audit(ctx,'Users',p.id,'deactivate'); return publicUser(result);
    });
    fail('UNKNOWN_ACTION','Aksi pengguna tidak dikenal.');
  }
  function validateTask(ctx,p,editing) {
    const obj={};
    ['title','description','instructions','projectId','category','priority','startDate','dueDate','estimateHours','checklist','links','notes'].forEach(k=>{
      if (p[k]!==undefined) obj[k]=limit(p[k],k==='description'||k==='instructions'?12000:4000);
    });
    if (!editing) obj.title=requireField(obj.title,'Judul tugas');
    if (obj.priority && !['Low','Medium','High','Urgent'].includes(obj.priority)) fail('VALIDATION_ERROR','Prioritas tidak valid.');
    if (obj.startDate && !dateOK(obj.startDate) || obj.dueDate && !dateOK(obj.dueDate)) fail('VALIDATION_ERROR','Format tanggal tidak valid.');
    if (obj.startDate && obj.dueDate && obj.dueDate < obj.startDate) fail('VALIDATION_ERROR','Deadline lebih awal dari tanggal mulai.');
    if (obj.projectId && !projectAllowed(ctx,obj.projectId)) fail('FORBIDDEN','Proyek di luar akses.');
    if (obj.links) obj.links=obj.links.split('\n').filter(Boolean).map(safeURL).join('\n');
    return obj;
  }
  function makeAssignments(ctx, task, ids) {
    const unique=[...new Set((Array.isArray(ids)?ids:[]).map(str))];
    if (!unique.length || unique.length>50) fail('VALIDATION_ERROR','Pilih 1–50 peserta.');
    unique.forEach(id=>{
      staffIntern(ctx,id);
      const intern=byId('Users',id,ctx);
      if (!intern || intern.role!=='intern' || !active(intern)) fail('VALIDATION_ERROR','Peserta tidak aktif.');
      if (all('TaskAssignments',ctx).some(a=>active(a)&&a.taskId===task.id&&a.internId===id)) return;
      const a=dbInsert('TaskAssignments',{taskId:task.id,internId:id,mentorId:task.mentorId,progress:'0',status:'To Do',revisions:'0',ownerId:id},ctx);
      taskEvent(ctx,a,'assigned','Tugas ditugaskan oleh '+ctx.user.fullName);
      notify(ctx,id,'Tugas baru: '+task.title,'Tenggat: '+(task.dueDate || 'belum ditentukan'),'tasks','task');
    });
  }
  function tasksApi(ctx,action,p,requestId) {
    if (action==='tasks.list') return taskList(ctx);
    if (action==='tasks.get') {
      const a=byId('TaskAssignments',p.assignmentId,ctx);
      if (!a || !assignmentAllowed(ctx,a.id)) fail('NOT_FOUND','Penugasan tidak ditemukan.');
      const t=taskByAssignment(a,ctx);
      if (!t) fail('NOT_FOUND','Tugas tidak ditemukan.');
      const byAssign=name=>visibleList(name,ctx).filter(x=>x.assignmentId===a.id).sort((x,y)=>x.createdAt.localeCompare(y.createdAt));
      return {task:t,comments:byAssign('TaskComments'),submissions:byAssign('TaskSubmissions'),reviews:byAssign('TaskReviews'),activity:byAssign('TaskActivity'),subtasks:byAssign('Subtasks')};
    }
    if (action==='tasks.create') return locked(()=>{
      requireStaff(ctx);
      const rid=limit(requestId || p.requestId || '',130);
      if (rid) { const prev=all('Tasks',ctx).find(t=>t.requestId===rid && t.createdBy===ctx.user.id); if (prev) return {id:prev.id,duplicate:true}; }
      const data=validateTask(ctx,p,false);
      const ids=Array.isArray(p.assigneeIds)?p.assigneeIds:[];
      if (!ids.length) fail('VALIDATION_ERROR','Pilih setidaknya satu peserta.');
      ids.forEach(id=>staffIntern(ctx,id));
      const task=dbInsert('Tasks',{...data,mentorId:ctx.user.id,createdBy:ctx.user.id,requestId:rid,status:'Active'},ctx);
      makeAssignments(ctx,task,ids);
      audit(ctx,'Tasks',task.id,'create'); return {id:task.id};
    });
    if (action==='tasks.update') return locked(()=>{
      requireStaff(ctx);
      const task=byId('Tasks',p.id,ctx);
      if (!task || !canSee(ctx,'Tasks',task)) fail('FORBIDDEN','Tidak dapat mengubah tugas ini.');
      if (!isAdmin(ctx) && task.mentorId!==ctx.user.id) fail('FORBIDDEN','Hanya pembuat tugas/admin dapat mengubah informasi tugas.');
      const edited=dbUpdate('Tasks',task.id,validateTask(ctx,p,true),ctx);
      audit(ctx,'Tasks',task.id,'update'); return {id:edited.id};
    });
    if (action==='tasks.archive') return locked(()=>{
      requireStaff(ctx);
      const t=byId('Tasks',p.id,ctx);
      if (!t || !canSee(ctx,'Tasks',t) || (!isAdmin(ctx)&&t.mentorId!==ctx.user.id)) fail('FORBIDDEN','Akses ditolak.');
      dbUpdate('Tasks',t.id,{status:'Archived'},ctx);
      audit(ctx,'Tasks',t.id,'archive'); return {archived:true};
    });
    if (action==='tasks.assign') return locked(()=>{
      requireStaff(ctx); const t=byId('Tasks',p.taskId,ctx);
      if (!t || !active(t) || (!isAdmin(ctx)&&t.mentorId!==ctx.user.id)) fail('FORBIDDEN','Tidak dapat menugaskan tugas.');
      makeAssignments(ctx,t,p.assigneeIds); audit(ctx,'Tasks',t.id,'assign'); return {assigned:true};
    });
    if (action==='tasks.updateProgress') return locked(()=>{
      const a=byId('TaskAssignments',p.assignmentId,ctx);
      if (!a || !active(a) || (a.internId!==ctx.user.id && !isStaff(ctx))) fail('FORBIDDEN','Penugasan tidak dapat diubah.');
      if (!assignmentAllowed(ctx,a.id)) fail('FORBIDDEN','Penugasan di luar akses.');
      if (['Completed','In Review'].includes(a.status)) fail('INVALID_TRANSITION','Tugas sedang ditinjau atau sudah selesai.');
      const progress=Math.max(0,Math.min(100,Number(p.progress)));
      if (!Number.isFinite(progress)) fail('VALIDATION_ERROR','Progres tidak valid.');
      const status=p.status || 'In Progress';
      if (!['To Do','In Progress','Blocked','Revision Required'].includes(status) || (status==='Revision Required' && a.status!=='Revision Required')) fail('INVALID_TRANSITION','Status tidak valid.');
      const changes={progress:String(Math.round(progress)),status};
      if (!a.startedAt && status==='In Progress') changes.startedAt=iso();
      dbUpdate('TaskAssignments',a.id,changes,ctx);
      taskEvent(ctx,a,'progress',status+' · '+Math.round(progress)+'%');
      return {updated:true};
    });
    if (action==='tasks.submit') return locked(()=>{
      const a=byId('TaskAssignments',p.assignmentId,ctx);
      if (!a || a.internId!==ctx.user.id || !active(a)) fail('FORBIDDEN','Anda tidak memiliki penugasan ini.');
      if (!['To Do','In Progress','Revision Required','Blocked'].includes(a.status)) fail('INVALID_TRANSITION','Status tidak mengizinkan pengumpulan.');
      const rid=limit(requestId || p.requestId || '',130);
      if (rid) { const prev=all('TaskSubmissions',ctx).find(x=>x.requestId===rid && x.internId===ctx.user.id); if (prev) return {id:prev.id,duplicate:true}; }
      const link=p.link?safeURL(p.link):'';
      const notes=limit(p.notes,5000);
      if (!link && !notes) fail('VALIDATION_ERROR','Lampirkan tautan pekerjaan atau catatan hasil.');
      const version=all('TaskSubmissions',ctx).filter(s=>s.assignmentId===a.id).length+1;
      const submission=dbInsert('TaskSubmissions',{assignmentId:a.id,taskId:a.taskId,internId:a.internId,link,notes,requestId:rid,version:String(version),ownerId:a.internId},ctx);
      dbUpdate('TaskAssignments',a.id,{status:'In Review',progress:'100',submittedAt:iso()},ctx);
      taskEvent(ctx,a,'submitted','Pengumpulan versi '+version);
      notify(ctx,a.mentorId,'Submission membutuhkan review',ctx.user.fullName+' mengumpulkan pekerjaan.','reviews','review');
      audit(ctx,'TaskSubmissions',submission.id,'submit'); return {id:submission.id,version};
    });
    if (action==='tasks.requestRevision' || action==='tasks.approve') return locked(()=>{
      requireStaff(ctx);
      const a=byId('TaskAssignments',p.assignmentId,ctx);
      if (!a || !assignmentAllowed(ctx,a.id) || a.status!=='In Review') fail('INVALID_TRANSITION','Hanya submission In Review yang dapat dinilai.');
      const submissions=all('TaskSubmissions',ctx).filter(s=>s.assignmentId===a.id).sort((x,y)=>y.createdAt.localeCompare(x.createdAt));
      const sub=submissions[0]; if (!sub) fail('NOT_FOUND','Submission tidak ditemukan.');
      const revision=action==='tasks.requestRevision', feedback=limit(p.feedback,4000);
      if (revision && !feedback) fail('VALIDATION_ERROR','Feedback revisi wajib diisi.');
      const score=p.score===''||p.score==null?'':String(Math.max(0,Math.min(100,Number(p.score))));
      dbInsert('TaskReviews',{submissionId:sub.id,assignmentId:a.id,mentorId:ctx.user.id,decision:revision?'Revision Required':'Approved',feedback,score,ownerId:a.internId},ctx);
      dbUpdate('TaskAssignments',a.id, revision ? {status:'Revision Required',revisions:String(Number(a.revisions||0)+1),progress:'75'} : {status:'Completed',completedAt:iso(),progress:'100'},ctx);
      taskEvent(ctx,a,revision?'revision':'approved',feedback);
      notify(ctx,a.internId,revision?'Revisi tugas diperlukan':'Tugas disetujui',feedback || 'Mentor telah menyetujui tugas Anda.','tasks',revision?'revision':'approval');
      audit(ctx,'TaskReviews',sub.id,revision?'revision':'approve');
      return {reviewed:true,decision:revision?'Revision Required':'Completed'};
    });
    if (action==='tasks.comments') {
      const a=byId('TaskAssignments',p.assignmentId,ctx);
      if (!a || !assignmentAllowed(ctx,a.id)) fail('FORBIDDEN','Komentar berada di luar akses Anda.');
      if (p.mode!=='add') return visibleList('TaskComments',ctx).filter(x=>x.assignmentId===a.id);
      return locked(()=>{
        const message=limit(requireField(p.message,'Komentar'),2500);
        const item=dbInsert('TaskComments',{assignmentId:a.id,taskId:a.taskId,userId:ctx.user.id,message,ownerId:a.internId},ctx);
        taskEvent(ctx,a,'comment',message.slice(0,100));
        notify(ctx,ctx.user.id===a.internId?a.mentorId:a.internId,'Komentar baru pada tugas',message.slice(0,160),'tasks','comment');
        return {id:item.id};
      });
    }
    if (action==='tasks.activity') return visibleList('TaskActivity',ctx).filter(x=>!p.assignmentId||x.assignmentId===p.assignmentId);
    fail('UNKNOWN_ACTION','Aksi tugas tidak dikenal.');
  }
  function recordsApi(ctx,action,p) {
    const name = str(p.table);
    if (!TABLES.includes(name)) fail('BAD_TABLE','Tabel tidak tersedia.');
    if (['Sessions','EmailVerifications','PasswordResets','Invitations'].includes(name)) fail('FORBIDDEN','Tabel internal tidak tersedia melalui API.');
    if (action==='records.list') {
      const data=visibleList(name,ctx);
      return (name==='Users'?data.map(publicUser):data.map(({_row,...rest})=>rest));
    }
    if (action==='records.save') return locked(()=>{
      if (!ADMIN_EDIT.has(name) && !STAFF_EDIT.has(name) && !INTERN_EDIT.has(name) && name!=='Subtasks') fail('FORBIDDEN','Gunakan endpoint khusus untuk mengubah data ini.');
      const data=p.data && typeof p.data==='object' ? p.data : {};
      const existing=data.id?byId(name,data.id,ctx):null;
      if (data.id && !existing) fail('NOT_FOUND','Data tidak ditemukan.');
      if (existing && !canSee(ctx,name,existing)) fail('FORBIDDEN','Tidak memiliki akses untuk mengubah data.');
      if (ADMIN_EDIT.has(name)) requireAdmin(ctx);
      if (STAFF_EDIT.has(name)) requireStaff(ctx);
      if (name==='DailyReports' && existing && existing.userId!==ctx.user.id) fail('FORBIDDEN','Hanya pembuat laporan dapat mengubah laporan.');
      if (name==='InternProfiles') {
        const id=existing?.userId || data.userId || ctx.user.id;
        if (!allowedIntern(ctx,id)) fail('FORBIDDEN','Profil berada di luar kewenangan.');
        if (existing && existing.userId!==id) fail('FORBIDDEN','Pemilik profil tidak bisa diubah.');
      }
      if (name==='Subtasks') {
        const aid=existing?.assignmentId || data.assignmentId;
        if (!assignmentAllowed(ctx,aid)) fail('FORBIDDEN','Subtask di luar akses.');
        if (!isStaff(ctx) && existing && data.title!==undefined && data.title!==existing.title) fail('FORBIDDEN','Judul subtask hanya dapat diubah mentor.');
      }
      if (name==='Milestones' && !projectAllowed(ctx,existing?.projectId || data.projectId)) fail('FORBIDDEN','Proyek di luar akses.');
      if (name==='Projects' && existing && !isAdmin(ctx) && existing.mentorId!==ctx.user.id) fail('FORBIDDEN','Hanya PIC dapat mengubah proyek.');
      if (name==='Announcements' && existing && !isAdmin(ctx) && existing.ownerId!==ctx.user.id) fail('FORBIDDEN','Anda bukan pembuat pengumuman.');
      if (name==='TaskTemplates' && existing && !isAdmin(ctx) && existing.ownerId!==ctx.user.id) fail('FORBIDDEN','Anda bukan pembuat template.');
      const cols=schema(name).filter(k=>!['id','createdAt','updatedAt','ownerId','status','passwordHash','passwordSalt','iterations','role'].includes(k));
      const clean={};
      cols.forEach(k=>{if (data[k]!==undefined) clean[k]=limit(data[k],['description','content','notes','checklist','learningGoals'].includes(k)?12000:4000);});
      for (const [k,v] of Object.entries(clean)) {
        if (/Date$/i.test(k) && !dateOK(v)) fail('VALIDATION_ERROR','Tanggal tidak valid: '+k);
        if (/^(link|avatarUrl|linkedin|github)$/i.test(k) && v) clean[k]=safeURL(v);
      }
      if (name==='DailyReports') { clean.userId=ctx.user.id; if (!clean.date && !existing) clean.date=iso().slice(0,10); if (!clean.type && !existing) clean.type='Check-Out'; }
      if (name==='InternProfiles') clean.userId=existing?.userId || data.userId || ctx.user.id;
      if (name==='Projects' && !existing) clean.mentorId=isAdmin(ctx) && data.mentorId ? str(data.mentorId) : ctx.user.id;
      if (name==='Projects' && clean.mentorId) {
        const mentor=byId('Users',clean.mentorId,ctx);
        if (!mentor || !['admin','mentor'].includes(mentor.role)) fail('VALIDATION_ERROR','PIC proyek harus admin/mentor.');
      }
      if (name==='Subtasks') {
        clean.assignmentId=existing?.assignmentId || str(data.assignmentId);
        const a=byId('TaskAssignments',clean.assignmentId,ctx);
        clean.taskId=a.taskId;
        if (clean.done!==undefined) clean.done=clean.done===true || clean.done==='true' ? 'true':'false';
      }
      if (name==='ProjectMembers') {
        const target=byId('Users',clean.userId || existing?.userId,ctx);
        if (!target || !active(target)) fail('VALIDATION_ERROR','Pengguna tidak ditemukan.');
        if (!byId('Projects',clean.projectId || existing?.projectId,ctx)) fail('VALIDATION_ERROR','Proyek tidak ditemukan.');
      }
      if (name==='MentorAssignments') {
        const mentor=byId('Users',clean.mentorId || existing?.mentorId,ctx);
        const intern=byId('Users',clean.internId || existing?.internId,ctx);
        if (!mentor || mentor.role!=='mentor' || !intern || intern.role!=='intern') fail('VALIDATION_ERROR','Pasangan mentor/peserta tidak valid.');
      }
      if (name==='Announcements' && !existing) clean.audience=clean.audience || 'all';
      if (name==='EvaluationCriteria' && clean.weight && (!Number.isFinite(Number(clean.weight)) || Number(clean.weight)<0)) fail('VALIDATION_ERROR','Bobot tidak valid.');
      const required={InternshipPrograms:['name'],Batches:['name'],Teams:['name'],Projects:['name'],Milestones:['name','projectId'],
        TaskTemplates:['title'],Announcements:['title','content'],EvaluationCriteria:['name'],AppSettings:['key','value'],
        ProjectMembers:['projectId','userId'],MentorAssignments:['mentorId','internId'],Subtasks:['assignmentId','title']};
      (required[name]||[]).forEach(k=>{if (!str(clean[k] || existing?.[k])) fail('VALIDATION_ERROR',k+' wajib diisi.');});
      const updated=existing ? dbUpdate(name,existing.id,clean,ctx) : dbInsert(name,{...clean,ownerId:ctx.user.id},ctx);
      audit(ctx,name,updated.id,existing?'update':'create');
      if (name==='Announcements' && !existing) {
        const users=all('Users',ctx).filter(u=>u.emailVerified==='true'&&active(u)&& (clean.audience==='all'||!clean.audience||u.role===clean.audience) && (!clean.batchId||u.batchId===clean.batchId));
        users.slice(0,200).forEach(u=>notify(ctx,u.id,'Pengumuman baru',clean.title,'announcements','announcement'));
      }
      return {id:updated.id};
    });
    if (action==='records.archive') return locked(()=>{
      const row=byId(name,p.id,ctx);
      if (!row || !canSee(ctx,name,row)) fail('FORBIDDEN','Record di luar akses.');
      if (ADMIN_EDIT.has(name)) requireAdmin(ctx);
      else if (STAFF_EDIT.has(name)) { requireStaff(ctx); if (!isAdmin(ctx) && row.ownerId!==ctx.user.id && row.mentorId!==ctx.user.id) fail('FORBIDDEN','Bukan pemilik record.'); }
      else if (name==='DailyReports' && row.userId===ctx.user.id) {} 
      else if (name==='Subtasks' && assignmentAllowed(ctx,row.assignmentId)) {}
      else fail('FORBIDDEN','Record tidak dapat diarsipkan.');
      dbUpdate(name,row.id,{status:'Archived'},ctx);
      audit(ctx,name,row.id,'archive'); return {archived:true};
    });
    fail('UNKNOWN_ACTION','Aksi records tidak dikenal.');
  }
  function evaluationsApi(ctx,action,p) {
    if (action==='evaluations.list') return visibleList('Evaluations',ctx).map(({_row,...rest})=>rest);
    if (action==='evaluations.create' || action==='evaluations.update') return locked(()=>{
      requireStaff(ctx);
      const old=p.id?byId('Evaluations',p.id,ctx):null;
      if (p.id && (!old || !canSee(ctx,'Evaluations',old))) fail('FORBIDDEN','Evaluasi tidak dapat diubah.');
      const internId=old?.internId || str(p.internId); staffIntern(ctx,internId);
      const changes={internId,mentorId:old?.mentorId || ctx.user.id,period:limit(p.period || old?.period,40),notes:limit(p.notes,5000)};
      ['quality','timeliness','communication','initiative','collaboration','problemSolving','ownership','learning'].forEach(k=>{
        const v=p[k]!==undefined?p[k]:old?.[k];
        if (v!==undefined && v!=='') { const n=Number(v); if (!Number.isFinite(n)||n<0||n>100) fail('VALIDATION_ERROR','Nilai '+k+' harus 0–100.'); changes[k]=String(n); }
      });
      if (!changes.period) fail('VALIDATION_ERROR','Periode evaluasi wajib diisi.');
      if (!old) changes.published='false';
      const row=old?dbUpdate('Evaluations',old.id,changes,ctx):dbInsert('Evaluations',{...changes,ownerId:internId},ctx);
      audit(ctx,'Evaluations',row.id,old?'update':'create'); return {id:row.id};
    });
    if (action==='evaluations.publish') return locked(()=>{
      requireStaff(ctx);
      const e=byId('Evaluations',p.id,ctx);
      if (!e || !canSee(ctx,'Evaluations',e)) fail('FORBIDDEN','Evaluasi tidak dapat dipublikasikan.');
      dbUpdate('Evaluations',e.id,{published:'true'},ctx);
      notify(ctx,e.internId,'Evaluasi telah tersedia','Hasil evaluasi performa Anda sudah dipublikasikan.','performance','evaluation');
      audit(ctx,'Evaluations',e.id,'publish'); return {published:true};
    });
    fail('UNKNOWN_ACTION','Aksi evaluasi tidak dikenal.');
  }
  function reports(ctx,action,p) {
    if (action==='reports.intern') {
      const id=p.internId || ctx.user.id;
      if (!allowedIntern(ctx,id)) fail('FORBIDDEN','Peserta berada di luar akses.');
      const tasks=taskList(ctx).filter(t=>t.internId===id);
      const completed=tasks.filter(t=>t.taskStatus==='Completed');
      const onTime=completed.filter(t=>!t.dueDate || (t.submittedAt && t.submittedAt.slice(0,10)<=t.dueDate));
      const reports=visibleList('DailyReports',ctx).filter(r=>r.userId===id);
      const reviews=visibleList('TaskReviews',ctx).filter(r=>tasks.some(t=>t.assignmentId===r.assignmentId));
      return {internId:id,total:tasks.length,completed:completed.length,completionRate:tasks.length?Math.round(completed.length/tasks.length*100):0,
        onTime:onTime.length,late:completed.length-onTime.length,revisions:tasks.reduce((sum,t)=>sum+t.revisions,0),reports:reports.length,
        averageScore:reviews.filter(r=>r.score!=='').length?Math.round(reviews.filter(r=>r.score!=='').reduce((sum,r)=>sum+Number(r.score),0)/reviews.filter(r=>r.score!=='').length):null};
    }
    if (action==='reports.tasks') return taskList(ctx);
    if (action==='reports.performance') {
      requireStaff(ctx);
      return visibleList('Users',ctx).filter(u=>u.role==='intern').map(u=>({user:publicUser(u),...reports(ctx,'reports.intern',{internId:u.id})}));
    }
    if (action==='reports.export') {
      const type=p.type || 'tasks';
      const map={tasks:()=>reports(ctx,'reports.tasks',{}),performance:()=>reports(ctx,'reports.performance',{}),
        daily:()=>visibleList('DailyReports',ctx),evaluations:()=>visibleList('Evaluations',ctx),
        projects:()=>visibleList('Projects',ctx),audits:()=>{requireAdmin(ctx);return visibleList('AuditLogs',ctx);}};
      if (!map[type]) fail('VALIDATION_ERROR','Jenis ekspor tidak dikenal.');
      const rows=map[type]();
      const keys=rows.length ? [...new Set(rows.flatMap(r=>Object.keys(r).filter(k=>!k.startsWith('_'))))] : ['message'];
      const escapeCSV = v => {
        let s=String(v===undefined||v===null?'':typeof v==='object'?JSON.stringify(v):v);
        if (/^[\s]*[=+\-@]/.test(s)) s="'"+s;
        return '"'+s.replace(/"/g,'""')+'"';
      };
      const csv=[keys.map(escapeCSV).join(','),...rows.map(r=>keys.map(k=>escapeCSV(r[k])).join(','))].join('\r\n');
      if (csv.length>1400000) fail('TOO_LARGE','Ekspor terlalu besar. Persempit data.');
      return {fileName:`edutask_${type}_${iso().slice(0,10)}.csv`,csv:'\uFEFF'+csv,count:rows.length};
    }
    fail('UNKNOWN_ACTION','Aksi laporan tidak dikenal.');
  }
  function notificationsApi(ctx,action,p) {
    if (action==='notifications.list') return visibleList('Notifications',ctx).sort((a,b)=>b.createdAt.localeCompare(a.createdAt)).map(({_row,...rest})=>rest);
    if (action==='notifications.markRead') return locked(()=>{
      const note=byId('Notifications',p.id,ctx);
      if (!note || note.userId!==ctx.user.id) fail('FORBIDDEN','Notifikasi tidak ditemukan.');
      dbUpdate('Notifications',note.id,{isRead:'true'},ctx); return {read:true};
    });
    fail('UNKNOWN_ACTION','Aksi notifikasi tidak dikenal.');
  }
  function handle(e) {
    let action='unknown',ctx={cache:{},token:''};
    try {
      const req=JSON.parse(e?.postData?.contents || '{}');
      action=str(req.action); if (!action || action.length>100) fail('VALIDATION_ERROR','API action diperlukan.');
      const p=req.payload && typeof req.payload==='object' && !Array.isArray(req.payload)?req.payload:{};
      const rid=str(req.requestId);
      if (rid.length>140) fail('VALIDATION_ERROR','Request id tidak valid.');
      ctx.token=str(req.token);
      if (action==='health') return response({status:'ok',version:'1.0.0'});
      if (!PUBLIC.has(action)) authenticate(ctx.token,ctx);
      let data;
      if (action.startsWith('auth.')) data=auth(ctx,action,p);
      else if (action==='dashboard.intern' || action==='dashboard.mentor' || action==='dashboard.refresh') {
        if (action==='dashboard.mentor') requireStaff(ctx);
        data=dashboard(ctx);
      }
      else if (action.startsWith('tasks.')) data=tasksApi(ctx,action,p,rid);
      else if (action.startsWith('users.')) data=usersApi(ctx,action,p);
      else if (action.startsWith('records.')) data=recordsApi(ctx,action,p);
      else if (action.startsWith('evaluations.')) data=evaluationsApi(ctx,action,p);
      else if (action.startsWith('reports.')) data=reports(ctx,action,p);
      else if (action.startsWith('notifications.')) data=notificationsApi(ctx,action,p);
      else if (action==='projects.list' || action==='batches.list') data=visibleList(action.startsWith('projects')?'Projects':'Batches',ctx).map(({_row,...rest})=>rest);
      else if (action==='projects.create' || action==='projects.update') data=recordsApi(ctx,'records.save',{table:'Projects',data:p});
      else if (action==='batches.create' || action==='batches.update') data=recordsApi(ctx,'records.save',{table:'Batches',data:p});
      else if (action==='projects.members') data=visibleList('ProjectMembers',ctx).filter(x=>x.projectId===p.projectId);
      else fail('UNKNOWN_ACTION','Endpoint tidak dikenal.');
      return response(data);
    } catch (error) {
      const permitted=new Set(['VALIDATION_ERROR','WEAK_PASSWORD','INVALID_CREDENTIALS','EMAIL_UNVERIFIED','ACCOUNT_EXISTS',
        'INVITE_REQUIRED','OTP_EXPIRED','OTP_LOCKED','OTP_INVALID','RATE_LIMIT','COOLDOWN','UNAUTHORIZED','SESSION_EXPIRED',
        'FORBIDDEN','NOT_FOUND','INVALID_TRANSITION','BAD_TABLE','BUSY','SETUP_REQUIRED','CONFIG_ERROR','TOO_LARGE','UNKNOWN_ACTION']);
      const code=permitted.has(error.code)?error.code:'INTERNAL_ERROR';
      console.error('EduTask API',action,code,String(error));
      return response(null,false,code,code==='INTERNAL_ERROR'?'Terjadi kesalahan server. Periksa Apps Script executions.':error.message);
    }
  }
  function response(data,success=true,code='',message='Operation successful') {
    return ContentService.createTextOutput(JSON.stringify({success,data,code,message,timestamp:iso()})).setMimeType(ContentService.MimeType.JSON);
  }
  function setup() {
    const id=props().getProperty('EDUTASK_SPREADSHEET_ID');
    if (!id) throw new Error('Isi Script Property EDUTASK_SPREADSHEET_ID dengan ID Google Sheets lebih dulu.');
    const s=ss();
    TABLES.forEach(name=>{
      const headers=schema(name), tab=s.getSheetByName(name) || s.insertSheet(name);
      if (tab.getLastRow()===0) { tab.getRange(1,1,1,headers.length).setValues([headers]); }
      else {
        const existing=tab.getRange(1,1,1,Math.max(tab.getLastColumn(),1)).getDisplayValues()[0].filter(Boolean);
        existing.forEach((h,i)=>{if (headers[i]!==h) throw new Error(name+' memiliki header berbeda pada kolom '+(i+1)+'. Jangan menimpa data: koreksi/migrasikan manual.');});
        if (existing.length<headers.length) tab.getRange(1,existing.length+1,1,headers.length-existing.length).setValues([headers.slice(existing.length)]);
      }
      tab.setFrozenRows(1);
      tab.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#087F8C').setFontColor('#FFFFFF');
    });
    if (!props().getProperty('EDUTASK_SECRET')) props().setProperty('EDUTASK_SECRET',Array.from({length:5},()=>Utilities.getUuid()).join('-'));
    if (!props().getProperty('EDUTASK_PBKDF2_ITERATIONS')) props().setProperty('EDUTASK_PBKDF2_ITERATIONS','12000');
    return 'EduTask database ready: '+TABLES.length+' sheets. Configure EDUTASK_BOOTSTRAP_EMAIL and EDUTASK_BOOTSTRAP_PASSWORD, then run bootstrapInitialAdmin().';
  }
  function bootstrap() {
    const email=validEmail(props().getProperty('EDUTASK_BOOTSTRAP_EMAIL'));
    const password=props().getProperty('EDUTASK_BOOTSTRAP_PASSWORD');
    if (!password || password.length<16) throw new Error('Set EDUTASK_BOOTSTRAP_PASSWORD minimal 16 karakter pada Script Properties.');
    return locked(()=>{
      if (all('Users').some(x=>x.role==='admin')) throw new Error('Admin sudah ada. Bootstrap terkunci agar tidak menambah admin lain.');
      if (all('Users').some(x=>x.email===email)) throw new Error('Email sudah digunakan.');
      const p=passwordParts(password);
      const user=dbInsert('Users',{email,fullName:'EduTask Administrator',...p,role:'admin',emailVerified:'true',status:'Active'});
      props().deleteProperty('EDUTASK_BOOTSTRAP_PASSWORD');
      audit(null,'Users',user.id,'bootstrapAdmin');
      return 'Admin pertama berhasil dibuat: '+email+'. Password bootstrap dihapus dari Script Properties. Ganti password setelah login.';
    });
  }
  return Object.freeze({handle,setup,bootstrap});
})();
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({success:true,data:{service:'EduTask API',status:'running'}})).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) { return EDUTASK.handle(e); }
function setupEdutask() { return EDUTASK.setup(); }
function bootstrapInitialAdmin() { return EDUTASK.bootstrap(); }
