/**
 * LAST CALL 슬롯 — 구글 시트 연결 코드 (Apps Script 웹 앱)
 *
 * 붙이는 곳: 'LAST CALL 슬롯 관리' 시트 → 확장 프로그램 → Apps Script → Code.gs 내용을 이 파일로 바꾸기
 * 배포: 배포 → 새 배포 → 유형 '웹 앱' → 다음 사용자 인증 정보로 실행 '나' → 액세스 권한 '모든 사용자'
 *
 * 탭 구성
 *   테이블: 테이블 | 상태 | 닉네임 | 세션 | 시작 | 마지막 활동 | 돌린 횟수
 *   쿠폰:   쿠폰번호 | 테이블 | 닉네임 | 경품 | 당첨 시각 | 지급 | 지급 시각 | 세션
 *   기록:   시각 | 테이블 | 닉네임 | 내용 | 세션
 *   경품:   코드 | 그림 | 경품
 *   설정:   관리키 / 알림 채널 / 자동 리셋(시간)
 */
const SHEET = { tables: '테이블', coupons: '쿠폰', log: '기록', prizes: '경품', settings: '설정' };
const ST = { free: '게임 가능', playing: '게임 중', prize: '경품 대기', done: '참여 완료' };
const PRIZE_IDS = ['seven', 'bar', 'bell', 'cilantro'];
const TZ = 'Asia/Seoul';

// ---------- entry points ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'table') return out_(tableInfo_(p.table, p.session));
    if (p.action === 'state') return out_(ownerState_(p.key));
    return out_({ ok: true, service: 'LAST CALL' });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ ok: false, error: '요청 형식이 잘못됐어요' }); }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return out_({ ok: false, error: 'busy' }); }
  try {
    switch (b.action) {
      case 'join':   return out_(join_(b));
      case 'spin':   return out_(spin_(b));
      case 'prize':  return out_(prize_(b));
      case 'end':    return out_(end_(b));
      case 'serve':  auth_(b.key); return out_(serve_(b));
      case 'reset':  auth_(b.key); return out_(reset_(b));
      case 'prizes': auth_(b.key); return out_(setPrizes_(b));
      case 'test':
        auth_(b.key);
        push_(settings_(), 'LAST CALL 알림 테스트', '이 알림이 보이면 설정 완료예요. 손님 입장과 당첨 알림이 여기로 와요.', ['white_check_mark'], 3);
        return out_({ ok: true });
      default: return out_({ ok: false, error: 'unknown action' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// ---------- helpers ----------
function sh_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ms_(d) { return d instanceof Date ? d.getTime() : (d ? (new Date(d).getTime() || 0) : 0); }
// game ids start with a letter so the sheet can never read them as a number or a date
function sid_() { return 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 11); }
// Write text exactly as typed: Sheets otherwise turns '007' into 7 and '3/4' into a date.
function asText_(sheet, row, cols) { cols.forEach(function (c) { sheet.getRange(row, c).setNumberFormat('@'); }); }
function appendText_(sheet, values, textCols) {
  const row = sheet.getLastRow() + 1;
  asText_(sheet, row, textCols);
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
  return row;
}
function todayStart_() { return new Date(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + 'T00:00:00+09:00').getTime(); }

function settings_() {
  const v = sh_(SHEET.settings).getRange(2, 1, 10, 2).getValues();
  const m = {};
  v.forEach(function (r) { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  return {
    key: String(m['관리키'] || '').trim(),
    topic: String(m['알림 채널'] || '').trim(),
    resetHours: Number(m['자동 리셋(시간)']) || 3,
  };
}
function auth_(key) {
  const k = settings_().key;
  if (!k || String(key || '') !== k) throw new Error('관리키가 맞지 않아요');
}
function prizes_() {
  const v = sh_(SHEET.prizes).getRange(2, 1, PRIZE_IDS.length, 3).getValues();
  const p = {};
  v.forEach(function (r) { const id = String(r[0]).trim(); if (PRIZE_IDS.indexOf(id) >= 0) p[id] = String(r[2]).trim(); });
  return p;
}
function readTables_() {
  const s = sh_(SHEET.tables), n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, 7).getValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r, i) {
      return { row: i + 2, no: String(r[0]).trim(), status: String(r[1] || ST.free), name: String(r[2] || ''),
               session: String(r[3] || ''), started: ms_(r[4]), last: ms_(r[5]), spins: Number(r[6]) || 0 };
    });
}
function writeTable_(t) {
  asText_(sh_(SHEET.tables), t.row, [3, 4]);
  sh_(SHEET.tables).getRange(t.row, 2, 1, 6).setValues([[
    t.status, t.name, t.session,
    t.started ? new Date(t.started) : '', t.last ? new Date(t.last) : '', t.spins || '',
  ]]);
}
function findTable_(no) {
  const t = readTables_().filter(function (x) { return x.no === String(no); })[0];
  if (!t) throw new Error('없는 테이블이에요');
  return t;
}
function readCoupons_() {
  const s = sh_(SHEET.coupons), n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, 8).getValues().map(function (r, i) {
    return { row: i + 2, code: String(r[0]), table: String(r[1]), name: String(r[2]), prize: String(r[3]),
             time: ms_(r[4]), served: r[5] === true, servedAt: ms_(r[6]), session: String(r[7]) };
  });
}
// a finished table shows '경품 대기' until every coupon of its session is served (also when ticked in the sheet)
function settle_(t, coupons) {
  if (t.status === ST.prize || t.status === ST.done) {
    const left = coupons.filter(function (c) { return c.session === t.session && !c.served; }).length;
    t.status = left ? ST.prize : ST.done;
  }
  return t;
}
// idle tables open up again after N hours
function effective_(t, cfg) {
  if (t.status !== ST.free && t.last && Date.now() - t.last > cfg.resetHours * 3600e3) {
    return Object.assign({}, t, { status: ST.free, name: '', session: '', started: 0, last: 0, spins: 0, expired: true });
  }
  return t;
}
function log_(table, name, text, session) { appendText_(sh_(SHEET.log), [new Date(), table, name, text, session || ''], [3, 4, 5]); }
function push_(cfg, title, message, tags, priority) {
  if (!cfg.topic) return;
  try {
    UrlFetchApp.fetch('https://ntfy.sh/', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ topic: cfg.topic, title: title, message: message, tags: tags, priority: priority }),
    });
  } catch (err) { /* the game keeps going even if the push fails */ }
}
function own_(b) {
  const t = findTable_(b.table);
  if (!b.session || t.session !== String(b.session)) throw new Error('이 테이블의 지금 게임이 아니에요');
  return t;
}

// ---------- guest ----------
function tableInfo_(no, session) {
  const cfg = settings_(), coupons = readCoupons_();
  const t = settle_(effective_(findTable_(no), cfg), coupons);
  const mine = !!session && t.session === String(session);
  const res = { ok: true, table: t.no, status: t.status, name: t.status === ST.free ? '' : t.name, mine: mine, prizes: prizes_() };
  if (mine) res.coupons = coupons.filter(function (c) { return c.session === t.session; })
    .map(function (c) { return { code: c.code, prize: c.prize, served: c.served }; });
  return res;
}
function join_(b) {
  const cfg = settings_();
  const t = findTable_(b.table);
  const eff = settle_(effective_(t, cfg), readCoupons_());
  if (eff.status !== ST.free) return { ok: false, reason: eff.status, name: eff.name };
  const name = String(b.name || '').trim().slice(0, 10) || '손님';
  t.status = ST.playing; t.name = name; t.session = sid_(); t.started = Date.now(); t.last = Date.now(); t.spins = 0;
  writeTable_(t);
  log_(t.no, name, '입장', t.session);
  push_(cfg, '손님 입장', '테이블 ' + t.no + ' · ' + name + ' 님이 게임을 시작했어요', ['slot_machine'], 3);
  return { ok: true, session: t.session, table: t.no, name: name, prizes: prizes_() };
}
function spin_(b) {
  const t = own_(b);
  t.last = Date.now(); t.spins = Math.max(t.spins, Number(b.spins) || 0);
  writeTable_(t);
  return { ok: true };
}
function prize_(b) {
  const cfg = settings_(), t = own_(b);
  const sym = String(b.symbol);
  if (PRIZE_IDS.indexOf(sym) < 0) throw new Error('잘못된 경품이에요');
  const prize = prizes_()[sym] || sym;
  const used = {};
  readCoupons_().forEach(function (c) { if (!c.served) used[c.code] = true; });
  let code;
  do { code = String(1000 + Math.floor(Math.random() * 9000)); } while (used[code]);
  const s = sh_(SHEET.coupons);
  const row = appendText_(s, [code, t.no, t.name, prize, new Date(), false, '', t.session], [3, 4, 8]);
  s.getRange(row, 6).insertCheckboxes();
  t.last = Date.now(); t.spins = Math.max(t.spins, Number(b.spins) || 0);
  writeTable_(t);
  log_(t.no, t.name, prize + ' 당첨 (쿠폰 ' + code + ')', t.session);
  push_(cfg, '경품 당첨', '테이블 ' + t.no + ' · ' + t.name + ' — ' + prize + ' 당첨 (쿠폰 ' + code + ')', ['tada'], 4);
  return { ok: true, code: code, prize: prize };
}
function end_(b) {
  const cfg = settings_(), t = own_(b);
  if (t.status !== ST.playing) return { ok: true, status: t.status };
  const mine = readCoupons_().filter(function (c) { return c.session === t.session; });
  t.status = mine.some(function (c) { return !c.served; }) ? ST.prize : ST.done;
  t.last = Date.now(); t.spins = Math.max(t.spins, Number(b.spins) || 0);
  writeTable_(t);
  const sum = mine.length ? mine.map(function (c) { return c.prize + ' (쿠폰 ' + c.code + ')'; }).join(', ') : '당첨 없음';
  log_(t.no, t.name, '게임 종료 · ' + t.spins + '번 돌림 · ' + sum, t.session);
  push_(cfg, '게임 종료', '테이블 ' + t.no + ' · ' + t.name + ' — ' + t.spins + '번 돌림 · ' + sum, ['checkered_flag'], 3);
  return { ok: true, status: t.status };
}

// ---------- owner ----------
function serve_(b) {
  const undo = !!b.undo;
  const c = readCoupons_().filter(function (x) { return x.code === String(b.code) && x.served === undo; })[0];
  if (!c) throw new Error('쿠폰을 찾지 못했어요');
  sh_(SHEET.coupons).getRange(c.row, 6, 1, 2).setValues([[!undo, undo ? '' : new Date()]]);
  log_(c.table, c.name, c.prize + (undo ? ' 지급 취소' : ' 지급') + ' (쿠폰 ' + c.code + ')', c.session);
  const t = readTables_().filter(function (x) { return x.session === c.session; })[0];
  if (t) { const before = t.status; settle_(t, readCoupons_()); if (t.status !== before) writeTable_(t); }
  return { ok: true };
}
function reset_(b) {
  const all = b.table === 'all';
  readTables_().forEach(function (t) {
    if (!all && t.no !== String(b.table)) return;
    if (t.status !== ST.free) log_(t.no, t.name, '새 손님 받기 (리셋)', t.session);
    t.status = ST.free; t.name = ''; t.session = ''; t.started = 0; t.last = 0; t.spins = 0;
    writeTable_(t);
  });
  return { ok: true };
}
function setPrizes_(b) {
  const s = sh_(SHEET.prizes);
  s.getRange(2, 1, PRIZE_IDS.length, 1).getValues().forEach(function (r, i) {
    const id = String(r[0]).trim();
    if (b.prizes && b.prizes[id]) s.getRange(i + 2, 3).setValue(String(b.prizes[id]).trim().slice(0, 24));
  });
  return { ok: true, prizes: prizes_() };
}
function ownerState_(key) {
  auth_(key);
  const cfg = settings_(), coupons = readCoupons_(), today = todayStart_();
  const tables = readTables_().map(function (t) {
    const e = settle_(effective_(t, cfg), coupons);
    return { no: e.no, status: e.status, name: e.name, session: e.session, started: e.started, last: e.last, spins: e.spins };
  });
  const list = coupons
    .filter(function (c) { return !c.served || c.time >= today || c.servedAt >= today; })
    .map(function (c) { return { code: c.code, table: c.table, name: c.name, prize: c.prize, time: c.time, served: c.served, servedAt: c.servedAt, session: c.session }; });
  let entries = 0;
  const ls = sh_(SHEET.log), last = ls.getLastRow();
  if (last > 1) {
    const from = Math.max(2, last - 2000);
    ls.getRange(from, 1, last - from + 1, 4).getValues().forEach(function (r) {
      if (ms_(r[0]) >= today && r[3] === '입장') entries++;
    });
  }
  return { ok: true, tables: tables, coupons: list, prizes: prizes_(), entriesToday: entries,
           topic: cfg.topic, resetHours: cfg.resetHours, now: Date.now() };
}
