/**
 * 스케줄러 v5 — Apps Script 백엔드 (수석비서 연동)
 * 시트: <스케쥴러 개인용>  탭: Tasks / Log / Config / AI가이드
 * 본부 시트: Config의 HQ_SHEET_ID (<수석비서 - 모든 업무>)  탭: Orders / Reports / Status / Teams / 프로젝트 마스터
 *
 * v5에서 더해진 것
 *  - 읽을 때 본부 시트도 함께 읽는다: 팀 목록, 프로젝트 이름, 대표 결정 필요(Status), 보고함(Reports 신규), 맡긴 일의 Orders 상태.
 *  - delegate: 할 일을 팀에 맡긴다 → Orders에 한 줄 추가, 할 일은 상태=대기, 대기상대=팀, 연결=o-XXXX.
 *  - 팀이 Orders를 완료/실패/보류로 바꾸면, 다음 읽기 때 그 할 일을 오늘 목록으로 되돌리고 결과를 메모에 붙인다.
 *  - ack: Reports 행의 상태를 <확인>으로 바꾼다 (대표가 봤다는 표시).
 *
 * v3에서 바뀐 점
 *  - id는 쓰는 쪽(앱·AI)이 직접 만든다. 같은 번호가 두 번 생기지 않는다.
 *  - 읽을 때마다 빈 id·중복 id·빈 순서를 자동으로 고친다 (AI가 시트에 직접 써도 안전).
 *  - 삭제는 상태=삭제 로 표시만 한다. 실제 정리는 cleanup()이 보관 탭으로 옮긴다.
 *  - 순서(Q열) 저장, 여러 개 한꺼번에 순서 바꾸기(reorder).
 *  - 같은 요청이 두 번 와도 결과가 같다 (재시도 안전).
 *  - 서버가 바쁘거나 오류가 나도 항상 JSON으로 답한다 (앱의 <저장 실패: Unexpected token> 방지).
 *
 * 배포: 이 파일 전체를 Code.gs에 붙여넣고 저장
 *       → 배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포   (주소는 그대로 유지됨)
 * 선택: 편집기 위쪽 함수 목록에서 installTriggers 를 한 번 실행하면
 *       매일 새벽 반복 할 일 되살리기 + 오래된 완료·삭제 정리가 자동으로 돈다.
 */

const SHEET_ID = '1qGFfbeqUBn6tViJKn0vm2nckAb5eb5vYp6ht8Fp-VA4';
const VERSION = 'v5';
const FEATURES = ['hq'];
const TZ = 'Asia/Seoul';

const COLS = ['id','title','bucket','today','status','waitingOn','category','project',
              'due','minutes','repeat','memo','createdAt','updatedAt','doneAt','source','order','link'];
const HEAD = ['id','제목','구간','오늘','상태','대기상대','분류','프로젝트',
              '마감','예상분','반복','메모','생성일','수정일','완료일','출처','순서','연결'];
const NCOL = COLS.length;
const C = {}; COLS.forEach((k, i) => C[k] = i);

const BUCKETS = ['지금','나중','아이디어'];
const STATUSES = ['할일','진행중','대기','완료','삭제'];
const DATE_KEYS = ['due','createdAt','updatedAt','doneAt'];
const EDITABLE = ['title','bucket','today','status','waitingOn','category','project',
                  'due','minutes','repeat','memo','doneAt','order','link'];

// ---------------- 공통 ----------------
function ss_(){ return SpreadsheetApp.openById(SHEET_ID); }
function sheet_(n){ return ss_().getSheetByName(n); }
function todayStr_(){ return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowStr_(){ return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function newId_(prefix){
  const t = Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss');
  const r = Math.random().toString(36).slice(2, 5);
  return (prefix || 's') + t + r;
}

function config_(){
  const sh = sheet_('Config'); const c = {};
  if (!sh) return c;
  sh.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) c[String(r[0]).trim()] = String(r[1]); });
  return c;
}

/** 시트에 보이는 날짜(2026-10-15, 2026. 10. 15, 10/15/2026 등)를 2026-10-15 로 통일 */
function normDate_(v){
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return s;
}
function bool_(v){ return v === true || /^(true|1|y|yes|o|✔|✓)$/i.test(String(v).trim()); }
function num_(v){ if (v === '' || v === null || v === undefined) return ''; const n = Number(v); return isNaN(n) ? '' : n; }

function rowToTask_(r){
  const t = {};
  COLS.forEach((k, i) => t[k] = (r[i] === undefined || r[i] === null) ? '' : String(r[i]).trim());
  t.today = bool_(t.today);
  t.minutes = num_(t.minutes);
  t.order = num_(t.order);
  DATE_KEYS.forEach(k => t[k] = normDate_(t[k]));
  if (BUCKETS.indexOf(t.bucket) < 0) t.bucket = '지금';
  if (STATUSES.indexOf(t.status) < 0) t.status = '할일';
  return t;
}
function taskToRow_(t){
  return COLS.map(k => {
    if (k === 'today') return t.today ? 'TRUE' : 'FALSE';
    const v = t[k];
    return (v === undefined || v === null) ? '' : v;
  });
}

// ---------------- 시트 읽기 + 자동 복구 ----------------
/** Tasks 탭 전체를 읽는다. rows[i].row 는 실제 시트 행 번호 */
function readAll_(){
  const sh = sheet_('Tasks');
  ensureHeader_(sh);
  const last = sh.getLastRow();
  if (last < 2) return { sh, rows: [] };
  const vals = sh.getRange(2, 1, last - 1, NCOL).getDisplayValues();
  const rows = [];
  vals.forEach((r, i) => {
    if (!String(r[C.id]).trim() && !String(r[C.title]).trim()) return; // 완전 빈 줄
    const t = rowToTask_(r); t._row = i + 2; rows.push(t);
  });
  return { sh, rows };
}

function ensureHeader_(sh){
  if (sh.getMaxColumns() < NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), NCOL - sh.getMaxColumns());
  const h = sh.getRange(1, 1, 1, NCOL).getDisplayValues()[0];
  if (String(h[C.order]).trim() !== '순서') sh.getRange(1, C.order + 1).setValue('순서');
  if (String(h[C.link]).trim() !== '연결') sh.getRange(1, C.link + 1).setValue('연결');
}

/** 빈 id, 중복 id, 빈 순서를 고친다. 고친 게 있으면 true */
function repair_(sh, rows){
  const seen = {}; let maxOrder = 0; const fixes = [];
  rows.forEach(t => { if (t.order !== '' && t.order > maxOrder) maxOrder = t.order; });
  rows.forEach(t => {
    let changed = false;
    if (!t.id){ t.id = newId_('r'); changed = true; fixes.push('빈 id → ' + t.id + ' (' + t.title + ')'); }
    else if (seen[t.id]){ const old = t.id; t.id = newId_('r'); changed = true; fixes.push('중복 id ' + old + ' → ' + t.id + ' (' + t.title + ')'); }
    if (t.order === ''){ maxOrder += 10; t.order = maxOrder; changed = true; }
    if (t.status === '완료' && t.today){ t.today = false; changed = true; }
    if (t.status === '완료' && !t.doneAt){ t.doneAt = t.updatedAt || todayStr_(); changed = true; }
    seen[t.id] = true;
    if (changed) sh.getRange(t._row, 1, 1, NCOL).setValues([taskToRow_(t)]);
  });
  if (fixes.length) log_('system', 'repair', '', '', fixes.join(' / '));
  return fixes.length > 0;
}

function findIndex_(rows, id, expectTitle){
  const hits = rows.filter(t => t.id === String(id));
  if (!hits.length) return null;
  if (hits.length > 1 && expectTitle){
    const exact = hits.filter(t => t.title === expectTitle);
    if (exact.length) return exact[0];
  }
  return hits[0];
}

function writeRow_(sh, rowNum, t){
  const rng = sh.getRange(rowNum, 1, 1, NCOL);
  rng.setValues([taskToRow_(t)]);
}

function appendTask_(sh, rows, t){
  let lastRow = 1;
  rows.forEach(r => { if (r._row > lastRow) lastRow = r._row; });
  const rowNum = lastRow + 1;
  if (rowNum > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 20);
  writeRow_(sh, rowNum, t);
  t._row = rowNum; rows.push(t);
}

function log_(who, action, id, title, detail){
  try { sheet_('Log').appendRow([nowStr_(), who || '', action || '', id || '', title || '', detail || '']); } catch (e) {}
}

function publicTask_(t){ const o = {}; COLS.forEach(k => o[k] = t[k]); return o; }

// ---------------- 읽기 ----------------
function doGet(e){
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'ping') return json_({ ok: true, version: VERSION, features: FEATURES, today: todayStr_(), hq: !!hqBook_() });
    if (p.action === 'config'){ const c = config_(); delete c.ADMIN_TOKEN; return json_({ ok: true, version: VERSION, config: c }); }

    let data = readAll_();
    const needs = data.rows.some((t, i) => !t.id || t.order === '' || data.rows.findIndex(x => x.id === t.id) !== i
                                         || (t.status === '완료' && (t.today || !t.doneAt)));
    if (needs){
      const lock = LockService.getScriptLock();
      if (lock.tryLock(5000)){
        try { data = readAll_(); repair_(data.sh, data.rows); } finally { lock.releaseLock(); }
      }
    }
    let hq = null;
    if (p.hq !== '0') hq = readHq_(data);          // 본부 시트 (없거나 못 열면 null)
    const tasks = data.rows.filter(t => t.status !== '삭제').map(publicTask_);
    return json_({ ok: true, version: VERSION, features: FEATURES, today: todayStr_(), tasks, hq });
  } catch (err) {
    return json_({ ok: false, version: VERSION, error: String(err && err.message || err) });
  }
}

// ---------------- 본부 시트 (수석비서) ----------------
function hqBook_(){
  const id = String(config_().HQ_SHEET_ID || '').trim();
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}
function hqVals_(book, name, rng){
  const sh = book.getSheetByName(name); if (!sh) return [];
  const last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(rng.replace('{N}', last)).getDisplayValues();
}
function isBlank_(v){ const s = String(v || '').trim(); return !s || s === '없음' || s === '-'; }

/** 본부 시트에서 앱이 보여줄 것만 뽑는다. 맡긴 일(Orders) 상태도 Tasks에 되비춘다. */
function readHq_(data){
  const book = hqBook_();
  if (!book) return { ok: false, error: 'HQ_SHEET_ID 없음 또는 열 수 없음' };
  try {
    const teams = hqVals_(book, 'Teams', 'A2:H{N}').filter(r => r[0])
      .map(r => ({ name: r[0], label: r[1] || r[0], area: r[2], cycle: r[3], active: String(r[6]).trim() === '가동중' }));
    const projects = hqVals_(book, '프로젝트 마스터', 'A2:F{N}').filter(r => r[0])
      .map(r => ({ name: r[0], team: r[2], status: r[3] }));
    const decisions = hqVals_(book, 'Status', 'A2:F{N}').filter(r => r[0] && !isBlank_(r[5]))
      .map(r => ({ team: r[0], updated: r[1], text: r[5] }));
    const reports = hqVals_(book, 'Reports', 'A2:J{N}');
    const inbox = [];
    reports.forEach((r, i) => {
      if (!r[0]) return;
      const status = String(r[8] || '').trim();
      const type = String(r[2] || '').trim();
      if (status !== '신규') return;
      if (isBlank_(r[6]) && type !== '경보' && type !== '제안') return;
      inbox.push({ row: i + 2, date: r[0], team: r[1], type, priority: r[3], title: r[4],
                   content: String(r[5] || '').slice(0, 400), action: isBlank_(r[6]) ? '' : r[6], due: r[7], source: r[9] });
    });
    // 맡긴 일의 Orders 상태
    const linked = data.rows.filter(t => /^o-\d+$/.test(t.link));
    const orders = {};
    if (linked.length){
      const want = {}; linked.forEach(t => want[t.link] = true);
      hqVals_(book, 'Orders', 'A2:K{N}').forEach(r => {
        if (want[r[0]]) orders[r[0]] = { team: r[3], status: r[7], at: r[8], result: r[9], note: r[10] };
      });
      syncOrders_(data, orders);
    }
    return { ok: true, teams, projects, decisions, inbox, orders };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/** 팀이 끝낸 일은 오늘 목록으로 되돌린다 (한 번만). */
function syncOrders_(data, orders){
  const done = data.rows.filter(t => t.status === '대기' && orders[t.link] && ['완료','실패','보류'].indexOf(orders[t.link].status) >= 0);
  if (!done.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    done.forEach(t => {
      const o = orders[t.link];
      const tag = '[팀 ' + o.status + '] ';
      const memo = (t.memo && t.memo.indexOf(tag) === 0) ? t.memo : (tag + (o.result || o.note || '') + (t.memo ? '\n' + t.memo : '')).trim();
      updateTask_(data.sh, data.rows, t.id, { status: '할일', today: true, memo }, 'hq', 'team-' + o.status);
    });
  } finally { lock.releaseLock(); }
}

function nextOrderId_(sh){
  let max = 0;
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getDisplayValues().forEach(r => { const m = String(r[0]).match(/^o-(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  return 'o-' + String(max + 1).padStart(4, '0');
}

/** 할 일을 팀에 맡긴다: Orders 한 줄 + 할 일은 대기 상태 */
function delegate_(sh, rows, op, who){
  const cur = findIndex_(rows, op.id, op.expectTitle);
  if (!cur) return { ok: false, error: 'not found', fatal: true };
  if (/^o-\d+$/.test(cur.link) && cur.status === '대기') return { ok: true, task: publicTask_(cur), orderId: cur.link, duplicate: true };
  const book = hqBook_();
  if (!book) return { ok: false, error: '본부 시트를 열 수 없습니다 (Config HQ_SHEET_ID 확인)', fatal: true };
  const team = String(op.team || '').trim();
  const teams = hqVals_(book, 'Teams', 'A2:H{N}');
  const tr = teams.filter(r => r[0] === team)[0];
  if (!tr) return { ok: false, error: '팀을 찾을 수 없습니다: ' + team, fatal: true };
  if (String(tr[6]).trim() !== '가동중') return { ok: false, error: team + '은(는) 아직 준비중입니다', fatal: true };
  const osh = book.getSheetByName('Orders');
  const oid = nextOrderId_(osh);
  const text = String(op.text || cur.title).trim();
  const pri = ['상','중','하'].indexOf(op.priority) >= 0 ? op.priority : '중';
  const due = normDate_(op.due || cur.due || '');
  const requester = String(config_().REQUESTER || '박보민(스케줄러)');
  osh.appendRow([oid, nowStr_(), requester, team, text, pri, due, '신규', '', '', '스케줄러 ' + cur.id]);
  const r = updateTask_(sh, rows, cur.id, { status: '대기', waitingOn: team, link: oid, today: false, due: due || cur.due }, who, 'delegate');
  r.orderId = oid; return r;
}

/** 보고를 봤다고 표시한다: Reports I열 신규 → 확인 */
function ackReport_(op){
  const book = hqBook_();
  if (!book) return { ok: false, error: '본부 시트를 열 수 없습니다', fatal: true };
  const sh = book.getSheetByName('Reports');
  const row = +op.row;
  if (!sh || !(row >= 2) || row > sh.getLastRow()) return { ok: false, error: 'bad row', fatal: true };
  const cell = sh.getRange(row, 9);
  const cur = String(cell.getDisplayValue()).trim();
  if (cur === '신규') cell.setValue(op.value === '완료' ? '완료' : '확인');
  log_('app', 'ack', '', 'Reports ' + row, cur + ' → ' + (op.value === '완료' ? '완료' : '확인'));
  return { ok: true, row, status: cur === '신규' ? (op.value === '완료' ? '완료' : '확인') : cur };
}

// ---------------- 쓰기 ----------------
function doPost(e){
  let body = {};
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, version: VERSION, error: 'bad json' }); }

  try {
    const c = config_();
    if (c.ADMIN_TOKEN && body.token !== c.ADMIN_TOKEN) return json_({ ok: false, version: VERSION, error: 'unauthorized', fatal: true });

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return json_({ ok: false, version: VERSION, error: 'busy', retry: true });
    try {
      const who = body.who || 'app';
      const data = readAll_();
      const ops = body.action === 'batch' ? (body.ops || []) : [body];
      const results = ops.map(op => runOp_(data.sh, data.rows, op, who));
      SpreadsheetApp.flush();
      if (body.action === 'batch') return json_({ ok: results.every(r => r.ok), version: VERSION, results });
      const r = results[0]; r.version = VERSION; return json_(r);
    } finally { lock.releaseLock(); }
  } catch (err) {
    return json_({ ok: false, version: VERSION, error: String(err && err.message || err), retry: true });
  }
}

function runOp_(sh, rows, op, who){
  switch (op.action){
    case 'add':     return addTask_(sh, rows, op.task || {}, who);
    case 'update':  return updateTask_(sh, rows, op.id, op.patch || {}, who, 'update', op.expectTitle);
    case 'done':    return updateTask_(sh, rows, op.id, { status: '완료' }, who, 'done', op.expectTitle);
    case 'undone':  return updateTask_(sh, rows, op.id, { status: '할일' }, who, 'undone', op.expectTitle);
    case 'delete':  return updateTask_(sh, rows, op.id, { status: '삭제', today: false }, who, 'delete', op.expectTitle);
    case 'restore': return updateTask_(sh, rows, op.id, { status: '할일' }, who, 'restore', op.expectTitle);
    case 'reorder': return reorder_(sh, rows, op.items || [], who);
    case 'delegate': return delegate_(sh, rows, op, who);
    case 'ack':      return ackReport_(op);
    default:        return { ok: false, error: 'unknown action', fatal: true };
  }
}

function cleanPatch_(p){
  const out = {};
  EDITABLE.forEach(k => { if (Object.prototype.hasOwnProperty.call(p, k)) out[k] = p[k]; });
  if ('title' in out) out.title = String(out.title || '').trim();
  if ('today' in out) out.today = bool_(out.today);
  if ('minutes' in out) out.minutes = num_(out.minutes);
  if ('order' in out) out.order = num_(out.order);
  if ('bucket' in out && BUCKETS.indexOf(out.bucket) < 0) delete out.bucket;
  if ('status' in out && STATUSES.indexOf(out.status) < 0) delete out.status;
  DATE_KEYS.forEach(k => { if (k in out) out[k] = normDate_(out[k]); });
  return out;
}

function addTask_(sh, rows, t, who){
  const id = String(t.id || '').trim();
  if (id){
    const exists = rows.filter(r => r.id === id)[0];
    if (exists) return { ok: true, task: publicTask_(exists), duplicate: true }; // 재시도로 같은 요청이 또 온 경우
  }
  const p = cleanPatch_(t);
  if (!p.title) return { ok: false, error: 'title required', fatal: true };
  let maxOrder = 0; rows.forEach(r => { if (r.order !== '' && r.order > maxOrder) maxOrder = r.order; });
  const today = todayStr_();
  const task = {
    id: id || newId_(who === 'claude' ? 'c' : 's'),
    title: p.title, bucket: p.bucket || '지금', today: !!p.today, status: p.status || '할일',
    waitingOn: p.waitingOn || '', category: p.category || '', project: p.project || '',
    due: p.due || '', minutes: p.minutes === undefined ? '' : p.minutes, repeat: p.repeat || '', memo: p.memo || '',
    createdAt: today, updatedAt: today, doneAt: '', source: who, link: p.link || '',
    order: (p.order === '' || p.order === undefined) ? maxOrder + 10 : p.order
  };
  if (task.status === '완료'){ task.doneAt = today; task.today = false; }
  appendTask_(sh, rows, task);
  log_(who, 'add', task.id, task.title, '');
  return { ok: true, task: publicTask_(task) };
}

function updateTask_(sh, rows, id, patch, who, actionName, expectTitle){
  const cur = findIndex_(rows, id, expectTitle);
  if (!cur) return { ok: false, error: 'not found', fatal: true };
  const p = cleanPatch_(patch);
  const today = todayStr_();
  if ('status' in p){
    if (p.status === '완료' && cur.status !== '완료'){ if (!p.doneAt) p.doneAt = today; p.today = false; }
    if (p.status !== '완료' && p.status !== '삭제') p.doneAt = '';
  }
  const next = Object.assign({}, cur, p, { updatedAt: today });
  writeRow_(sh, cur._row, next);
  Object.assign(cur, next);
  log_(who, actionName || 'update', cur.id, cur.title, JSON.stringify(p));
  return { ok: true, task: publicTask_(cur) };
}

function reorder_(sh, rows, items, who){
  let n = 0;
  items.forEach(it => {
    const cur = findIndex_(rows, it.id);
    const o = num_(it.order);
    if (!cur || o === '') return;
    cur.order = o;
    sh.getRange(cur._row, C.order + 1).setValue(o);
    n++;
  });
  if (n) log_(who, 'reorder', '', '', n + '건');
  return { ok: true, count: n };
}

// ---------------- 정기 작업 (installTriggers 로 켬) ----------------
/** 반복 할 일: 완료된 반복 항목을 다음 날짜가 되면 다시 할일로 되살린다 */
function rolloverRepeats(){
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const { sh, rows } = readAll_();
    const today = todayStr_();
    rows.filter(t => t.repeat && t.status === '완료').forEach(t => {
      const d = new Date((t.doneAt || today) + 'T00:00:00+09:00');
      let next = new Date(d);
      if (t.repeat === '매일') next.setDate(d.getDate() + 1);
      else if (t.repeat === '매주') next.setDate(d.getDate() + 7);
      else if (t.repeat === '매월') next.setMonth(d.getMonth() + 1);
      else if (t.repeat === '월말') next = new Date(d.getFullYear(), d.getMonth() + 2, 0);
      else return;
      const nextStr = Utilities.formatDate(next, TZ, 'yyyy-MM-dd');
      if (nextStr <= today) updateTask_(sh, rows, t.id, { status: '할일', due: nextStr, today: true }, 'system', 'repeat');
    });
  } finally { lock.releaseLock(); }
}

/** 삭제 표시 7일 지난 것, 완료 30일 지난 것을 <보관> 탭으로 옮긴다 */
function cleanup(){
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const { sh, rows } = readAll_();
    const ss = ss_();
    let arc = ss.getSheetByName('보관');
    if (!arc){ arc = ss.insertSheet('보관'); arc.getRange(1, 1, 1, NCOL).setValues([HEAD]); arc.setFrozenRows(1); }
    const today = new Date(todayStr_() + 'T00:00:00+09:00');
    const age = s => s ? Math.floor((today - new Date(s + 'T00:00:00+09:00')) / 86400000) : 0;
    const move = rows.filter(t => (t.status === '삭제' && age(t.updatedAt) >= 7) || (t.status === '완료' && !t.repeat && age(t.doneAt) >= 30));
    if (!move.length) return;
    arc.getRange(arc.getLastRow() + 1, 1, move.length, NCOL).setValues(move.map(taskToRow_));
    move.map(t => t._row).sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
    log_('system', 'cleanup', '', '', move.length + '건 보관 탭으로 이동');
  } finally { lock.releaseLock(); }
}

function daily(){ rolloverRepeats(); cleanup(); }

/** 편집기에서 한 번만 실행: 매일 새벽 5시에 daily() 실행 */
function installTriggers(){
  ScriptApp.getProjectTriggers().forEach(tr => { if (tr.getHandlerFunction() === 'daily') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('daily').timeBased().everyDays(1).atHour(5).inTimezone(TZ).create();
}
