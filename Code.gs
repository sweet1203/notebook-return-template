/**
 * 노트북 반납 확인 웹앱 - 템플릿 (Google Apps Script)
 * https://github.com/sweet1203/notebook-return-template
 *
 * 시트 구성
 *  - 사용방법 : 설치·사용 안내
 *  - 학생입력 : 선생님이 학번·이름(·시리얼번호·비고)을 붙여넣는 곳
 *  - 반납현황 : 학생 1명당 1행. 담임이 체크한 반납 물품이 기록됨 (최종 수합 파일)
 *  - 담임설정 : 반 | 담임선생님 | 링크키 | 반별 링크   (메뉴로 자동 생성)
 *  - 설정     : 제목 / 관리자비밀번호 / 웹앱 URL / 관리자 링크
 *
 * 시트 메뉴 [📦 노트북 반납]
 *  1. 학생명단 반영하기 : 학생입력 → 반납현황·담임설정 만들기 (이미 체크한 기록은 유지)
 *  2. 반별 링크 만들기   : 설정의 웹앱 URL로 관리자 링크·반별 링크 만들기
 */

// 화면 주소 (모든 학교가 같이 씀. 각 학교의 웹앱 주소는 링크의 api 값으로 전달됨)
const FRONT_URL = 'https://sweet1203.github.io/notebook-return-template/';

const SHEET_INPUT = '학생입력';
const SHEET_STATUS = '반납현황';
const SHEET_TEACHER = '담임설정';
const SHEET_CONFIG = '설정';

const ST_NONE = '미제출';
const ST_DONE = '제출완료';

const COLS = ['학번', '반', '번호', '이름', '시리얼번호', '비고',
  '노트북', '충전기', '충전케이블', '상태', '제출시각'];
const ITEM_COLS = { notebook: '노트북', charger: '충전기', cable: '충전케이블' };
const CONFIRM_COLS = ['관리자확인', '확인시각'];   // 관리자 수령 확인 (열이 없으면 자동으로 추가)
const TEACHER_HEADER = ['반', '담임선생님', '링크키', '반별 링크'];

// 화면에서 호출할 수 있는 함수 목록
const API = {
  info: info,
  classLoad: classLoad,
  classSave: classSave,
  adminLoad: adminLoad,
  adminSave: adminSave,
  adminConfirm: adminConfirm,
  adminReset: adminReset
};

/* ---------- 웹앱 (JSON API) ---------- */

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 요청 본문: {"fn": "함수이름", "args": [...]}
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.fn];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    return json_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

// 웹앱 주소를 브라우저로 직접 열면 관리자 화면으로 보냄
function doGet() {
  const url = ScriptApp.getService().getUrl();
  const admin = FRONT_URL + '?api=' + apiId_(url) + '&mode=admin';
  return HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">노트북 반납 API 동작 중입니다. ' +
    '<a href="' + admin + '" target="_top">관리자 화면 열기</a></p>');
}

/* ---------- 공통 ---------- */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 시트가 없습니다.');
  return sh;
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function clean_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '');
}

// 설정 시트에서 A열 이름으로 B열 값 읽기
function config_(name) {
  const v = sheet_(SHEET_CONFIG).getDataRange().getDisplayValues();
  for (let r = 0; r < v.length; r++) {
    if (clean_(v[r][0]) === clean_(name)) return String(v[r][1]).trim();
  }
  return '';
}

function setConfig_(name, value) {
  const sh = sheet_(SHEET_CONFIG);
  const v = sh.getDataRange().getDisplayValues();
  for (let r = 0; r < v.length; r++) {
    if (clean_(v[r][0]) === clean_(name)) { sh.getRange(r + 1, 2).setValue(value); return; }
  }
  sh.appendRow([name, value]);
}

function title_() {
  return config_('제목') || '노트북 반납 확인';
}

// 웹앱 URL(https://script.google.com/macros/s/아이디/exec) → 아이디
function apiId_(url) {
  const m = String(url || '').match(/\/macros\/s\/([A-Za-z0-9_-]+)\/exec/);
  return m ? m[1] : '';
}

// 반납현황 시트 전체를 읽어 {sheet, idx(열 위치), rows(객체 배열)} 로 돌려줌
function loadStatus_() {
  const sh = sheet_(SHEET_STATUS);
  const values = sh.getDataRange().getDisplayValues();
  const header = values[0].map(String);
  const idx = {};
  COLS.forEach(function (c) {
    idx[c] = header.indexOf(c);
    if (idx[c] < 0) throw new Error('"' + SHEET_STATUS + '" 시트에 "' + c + '" 열이 없습니다. 메뉴에서 [학생명단 반영하기]를 먼저 실행해 주세요.');
  });
  CONFIRM_COLS.forEach(function (c) {
    idx[c] = header.indexOf(c);
    if (idx[c] < 0) {
      idx[c] = header.length;
      header.push(c);
      sh.getRange(1, idx[c] + 1).setValue(c).setFontWeight('bold');
    }
  });
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    if (!clean_(values[r][idx['학번']])) continue;
    const o = { _r: r };
    COLS.concat(CONFIRM_COLS).forEach(function (c) {
      const v = values[r][idx[c]];
      o[c] = String(v == null ? '' : v).trim();
    });
    rows.push(o);
  }
  return { sheet: sh, idx: idx, rows: rows };
}

function toClient_(o) {
  const nb = o['노트북'] === 'O', ch = o['충전기'] === 'O', ca = o['충전케이블'] === 'O';
  return {
    sid: o['학번'], grade: o['학번'].charAt(0), cls: o['반'], no: o['번호'], name: o['이름'],
    serial: o['시리얼번호'], memo: o['비고'],
    notebook: nb, charger: ch, cable: ca,
    status: (nb || ch || ca) ? ST_DONE : ST_NONE,
    submittedAt: o['제출시각'],
    adminOk: o['관리자확인'] === 'O', adminOkAt: o['확인시각']
  };
}

function sortRows_(rows) {
  return rows.sort(function (a, b) { return a['학번'].localeCompare(b['학번']); });
}

// 반 번호 + 링크키가 맞으면 {cls, name}
function checkClass_(cls, key) {
  cls = clean_(cls);
  const v = sheet_(SHEET_TEACHER).getDataRange().getDisplayValues();
  for (let r = 1; r < v.length; r++) {
    if (clean_(v[r][0]) === cls && clean_(v[r][2]) && clean_(v[r][2]) === clean_(key)) {
      return { cls: cls, name: String(v[r][1]).trim() };
    }
  }
  throw new Error('반 링크가 올바르지 않습니다. 관리자에게 받은 링크로 다시 접속해 주세요.');
}

function checkAdmin_(pw) {
  const real = clean_(config_('관리자비밀번호'));
  if (!real) throw new Error('설정 시트에 관리자 비밀번호가 비어 있습니다. 먼저 정해 주세요.');
  if (real !== clean_(pw)) throw new Error('관리자 비밀번호가 올바르지 않습니다.');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// 학생별 물품 체크를 시트에 기록. changes = [{sid, notebook, charger, cable}]
// 물품이 하나라도 체크되면 제출완료 + 제출시각(지금), 모두 해제되면 미제출
// 물품이 바뀐 학생은 관리자 수령 확인을 지움 (다시 확인해야 함)
function writeItems_(data, changes, onlyCls) {
  const range = data.sheet.getDataRange();
  const values = range.getValues();
  const bySid = {};
  data.rows.forEach(function (o) { bySid[clean_(o['학번'])] = o; });
  const time = now_();
  let count = 0;
  changes.forEach(function (c) {
    const o = bySid[clean_(c.sid)];
    if (!o || (onlyCls && clean_(o['반']) !== onlyCls)) return;
    const row = values[o._r];
    const any = !!(c.notebook || c.charger || c.cable);
    const changed = Object.keys(ITEM_COLS).some(function (k) {
      return (row[data.idx[ITEM_COLS[k]]] === 'O') !== !!c[k];
    });
    if (changed) { row[data.idx['관리자확인']] = ''; row[data.idx['확인시각']] = ''; }
    Object.keys(ITEM_COLS).forEach(function (k) { row[data.idx[ITEM_COLS[k]]] = c[k] ? 'O' : ''; });
    row[data.idx['상태']] = any ? ST_DONE : ST_NONE;
    row[data.idx['제출시각']] = any ? time : '';
    count++;
  });
  range.setValues(values);
  return count;
}

/* ---------- API: 반별 체크 (담임) ---------- */

function info() {
  return { title: title_() };
}

function classLoad(cls, key) {
  const t = checkClass_(cls, key);
  const data = loadStatus_();
  const rows = sortRows_(data.rows.filter(function (o) { return clean_(o['반']) === t.cls; }));
  return {
    title: title_(), cls: t.cls, grade: rows.length ? rows[0]['학번'].charAt(0) : '',
    teacher: t.name, rows: rows.map(toClient_), loadedAt: now_()
  };
}

function classSave(cls, key, changes) {
  const t = checkClass_(cls, key);
  const count = withLock_(function () {
    return writeItems_(loadStatus_(), changes || [], t.cls);
  });
  const res = classLoad(cls, key);
  res.message = count + '명 저장했습니다';
  return res;
}

/* ---------- API: 관리자 ---------- */

function adminLoad(pw) {
  checkAdmin_(pw);
  const data = loadStatus_();
  const classes = [];
  sheet_(SHEET_TEACHER).getDataRange().getDisplayValues().slice(1).forEach(function (r) {
    if (clean_(r[0])) classes.push({ cls: clean_(r[0]), teacher: String(r[1]).trim(), key: clean_(r[2]) });
  });
  const ss = ss_();
  return {
    title: title_(),
    rows: sortRows_(data.rows).map(toClient_),
    classes: classes,
    sheetUrl: ss.getUrl(),
    xlsxUrl: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    loadedAt: now_()
  };
}

// 관리자가 학생별 물품 체크 저장 (모든 반)
function adminSave(pw, changes) {
  checkAdmin_(pw);
  const count = withLock_(function () {
    return writeItems_(loadStatus_(), changes || []);
  });
  const res = adminLoad(pw);
  res.message = count + '명 저장했습니다';
  return res;
}

// 관리자 수령 확인. sids 학생(제출완료인 학생만)을 on=true 면 확인, false 면 확인 취소
function adminConfirm(pw, sids, on) {
  checkAdmin_(pw);
  const want = {};
  (sids || []).forEach(function (s) { want[clean_(s)] = true; });
  const count = withLock_(function () {
    const data = loadStatus_();
    const range = data.sheet.getDataRange();
    const values = range.getValues();
    const time = now_();
    let n = 0;
    data.rows.forEach(function (o) {
      if (!want[clean_(o['학번'])]) return;
      if (!(o['노트북'] === 'O' || o['충전기'] === 'O' || o['충전케이블'] === 'O')) return;
      values[o._r][data.idx['관리자확인']] = on ? 'O' : '';
      values[o._r][data.idx['확인시각']] = on ? time : '';
      n++;
    });
    range.setValues(values);
    return n;
  });
  const res = adminLoad(pw);
  res.message = count + '명 ' + (on ? '수령 확인했습니다' : '수령 확인을 취소했습니다');
  return res;
}

// 제출 삭제 → 물품 체크를 모두 지우고 미제출로
function adminReset(pw, sid) {
  checkAdmin_(pw);
  withLock_(function () {
    writeItems_(loadStatus_(), [{ sid: sid }]);
  });
  const res = adminLoad(pw);
  res.message = sid + ' 학생의 제출을 삭제했습니다';
  return res;
}

/* ---------- 시트 메뉴 ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 노트북 반납')
    .addItem('1. 학생명단 반영하기', 'syncRoster')
    .addItem('2. 반별 링크 만들기', 'makeLinks')
    .addToUi();
}

function randomKey_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function ensureSheet_(name, header) {
  let sh = ss_().getSheetByName(name);
  if (!sh) sh = ss_().insertSheet(name);
  if (header && !clean_(sh.getRange(1, 1).getDisplayValue())) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 학생입력 → 반납현황 (이미 체크한 물품·제출시각은 유지), 담임설정에 반 추가
function syncRoster() {
  const ui = SpreadsheetApp.getUi();
  const input = sheet_(SHEET_INPUT).getDataRange().getDisplayValues();
  const errors = [];
  const seen = {};
  const students = [];
  for (let r = 1; r < input.length; r++) {
    const sid = clean_(input[r][0]);
    const name = String(input[r][1] || '').trim();
    if (!sid && !name) continue;
    if (!/^\d{5}$/.test(sid)) { errors.push((r + 1) + '행: 학번 "' + input[r][0] + '"은(는) 5자리 숫자가 아닙니다'); continue; }
    if (!name) { errors.push((r + 1) + '행: 이름이 비어 있습니다'); continue; }
    if (seen[sid]) { errors.push((r + 1) + '행: 학번 ' + sid + ' 이(가) ' + seen[sid] + '행과 중복됩니다'); continue; }
    seen[sid] = r + 1;
    students.push({
      sid: sid, name: name,
      serial: String(input[r][2] || '').trim(), memo: String(input[r][3] || '').trim()
    });
  }
  if (errors.length) {
    ui.alert('학생입력 시트를 확인해 주세요',
      errors.slice(0, 15).join('\n') + (errors.length > 15 ? '\n… 외 ' + (errors.length - 15) + '건' : ''),
      ui.ButtonSet.OK);
    return;
  }
  if (!students.length) {
    ui.alert('학생입력 시트에 학생이 없습니다. 2행부터 학번과 이름을 붙여넣어 주세요.');
    return;
  }

  const dropped = withLock_(function () {
    // 기존 체크 기록
    const prev = {};
    const st = ensureSheet_(SHEET_STATUS, COLS);
    try {
      loadStatus_().rows.forEach(function (o) { prev[clean_(o['학번'])] = o; });
    } catch (e) { /* 처음 실행이면 기록 없음 */ }

    students.sort(function (a, b) { return a.sid.localeCompare(b.sid); });
    const rows = students.map(function (s) {
      const p = prev[s.sid] || {};
      return [s.sid, String(parseInt(s.sid.substr(1, 2), 10)), String(parseInt(s.sid.substr(3, 2), 10)), s.name,
        s.serial, s.memo, p['노트북'] || '', p['충전기'] || '', p['충전케이블'] || '',
        p['상태'] || ST_NONE, p['제출시각'] || '', p['관리자확인'] || '', p['확인시각'] || ''];
    });
    st.clearContents();
    const head = COLS.concat(CONFIRM_COLS);
    st.getRange(1, 1, rows.length + 1, head.length).setNumberFormat('@');
    st.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    st.getRange(2, 1, rows.length, head.length).setValues(rows);
    st.setFrozenRows(1);

    // 담임설정: 새 반만 추가 (기존 담임 이름·링크키는 유지)
    const th = ensureSheet_(SHEET_TEACHER, TEACHER_HEADER);
    const tv = th.getDataRange().getDisplayValues();
    const have = {};
    for (let r = 1; r < tv.length; r++) if (clean_(tv[r][0])) have[clean_(tv[r][0])] = true;
    const classes = {};
    rows.forEach(function (row) { classes[row[1]] = true; });
    Object.keys(classes).sort(function (a, b) { return a - b; }).forEach(function (c) {
      if (!have[c]) th.appendRow([c, '', randomKey_(), '']);
    });
    th.getRange(2, 1, Math.max(th.getLastRow() - 1, 1), 3).setNumberFormat('@');

    return Object.keys(prev).filter(function (sid) { return !seen[sid]; });
  });

  ui.alert('반영 완료',
    students.length + '명을 반납현황에 반영했습니다.\n' +
    (dropped.length ? '학생입력에 없는 ' + dropped.length + '명(' + dropped.slice(0, 5).join(', ') +
      (dropped.length > 5 ? ' 등' : '') + ')은 반납현황에서 뺐습니다.\n' : '') + '\n' +
    '다음 단계\n· 담임설정 시트에 담임선생님 성함 입력(선택)\n· 설정 시트에 관리자 비밀번호 정하기\n· 웹앱 배포 후 [2. 반별 링크 만들기]',
    ui.ButtonSet.OK);
}

// 설정의 웹앱 URL → 관리자 링크(설정 시트)와 반별 링크(담임설정 D열)
function makeLinks() {
  const ui = SpreadsheetApp.getUi();
  const id = apiId_(config_('웹앱 URL'));
  if (!id) {
    ui.alert('설정 시트의 [웹앱 URL] 칸에 배포한 웹앱 주소를 붙여넣어 주세요.\n' +
      '(https://script.google.com/macros/s/…/exec 형식, 사용방법 시트 참고)');
    return;
  }
  const base = FRONT_URL + '?api=' + id;
  setConfig_('관리자 링크', base + '&mode=admin');
  const th = sheet_(SHEET_TEACHER);
  const tv = th.getDataRange().getDisplayValues();
  const lines = [];
  for (let r = 1; r < tv.length; r++) {
    const cls = clean_(tv[r][0]);
    if (!cls) continue;
    let key = clean_(tv[r][2]);
    if (!key) { key = randomKey_(); th.getRange(r + 1, 3).setValue(key); }
    const link = base + '&class=' + cls + '&key=' + key;
    th.getRange(r + 1, 4).setValue(link);
    lines.push(cls + '반: ' + link);
  }
  ui.alert('링크를 만들었습니다',
    '관리자 링크는 설정 시트, 반별 링크는 담임설정 시트 D열에 있습니다.\n' +
    '관리자 화면에서도 반별 링크를 복사할 수 있습니다.\n\n관리자 링크:\n' + base + '&mode=admin',
    ui.ButtonSet.OK);
}
