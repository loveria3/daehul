// ================================================================
// 🍊 정성가득 대흘귤집 황금향 — Google Apps Script 백엔드
// ================================================================
// 사용 방법:
//   1. Google 스프레드시트를 새로 만들고 URL에서 ID를 복사
//      예) https://docs.google.com/spreadsheets/d/【여기가 ID】/edit
//   2. 아래 SPREADSHEET_ID 에 붙여넣기
//   3. 스크립트 편집기에서 [배포] → [새 배포] → 웹 앱 선택
//      - 실행 계정: 나(내 계정)
//      - 액세스 권한: 모든 사용자
//   4. 배포 후 나오는 URL을 HTML 파일의 APPS_SCRIPT_URL 에 붙여넣기
// ================================================================

var CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',  // ← 스프레드시트 ID 입력
  ADMIN_PASSWORD: 'daheul2025',            // ← 관리자 비밀번호 (변경 가능)
  ORDERS_SHEET:   '주문내역',
  CONTACTS_SHEET: '문의내역',
  SEASON_YEAR:    2026
};

// ================================================================
// POST 핸들러 — 주문 접수 / 문의 접수
// ================================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'order';

    if (action === 'order')   return saveOrder(data);
    if (action === 'contact') return saveContact(data);

    return jsonOk({ success: false, message: '알 수 없는 요청입니다' });
  } catch (err) {
    return jsonOk({ success: false, message: '서버 오류: ' + err.toString() });
  }
}

// ================================================================
// GET 핸들러 — 관리자 조회 / 상태 업데이트
// ================================================================
function doGet(e) {
  var p      = e.parameter;
  var action = p.action || '';

  // 관리자 전용 액션은 비밀번호 확인
  var adminActions = ['getOrders', 'getContacts', 'updateOrder'];
  if (adminActions.indexOf(action) >= 0 && p.password !== CONFIG.ADMIN_PASSWORD) {
    return jsonOk({ success: false, message: '비밀번호가 올바르지 않습니다' });
  }

  if (action === 'getOrders')   return getOrders(p);
  if (action === 'getContacts') return getContacts(p);
  if (action === 'updateOrder') return updateOrderField(p);
  if (action === 'ping')        return jsonOk({ success: true, message: 'pong', year: CONFIG.SEASON_YEAR });

  return jsonOk({ success: false, message: '알 수 없는 액션입니다' });
}

// ================================================================
// 주문 저장
// ================================================================
function saveOrder(data) {
  var ss    = getSpreadsheet();
  var sheet = getOrCreateSheet(ss, CONFIG.ORDERS_SHEET);

  // 최초 실행 시 헤더 생성
  if (sheet.getLastRow() === 0) {
    var headers = [
      'NO', '주문일자', '접수시각', '주문채널',
      '보내는분성명', '보내는분전화번호', '보내는분주소',
      '받는분성명', '받는분전화번호', '받는분기타연락처', '받는분주소',
      '품목명', '박스수량', '단가', '합계',
      '배송메시지', '상태', '입금확인액', '비고'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#1E3A2F')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 100);  // 보내는분성명
    sheet.setColumnWidth(11, 250); // 받는분주소
  }

  var no  = Math.max(sheet.getLastRow(), 1); // 헤더 포함 행 수
  var now = new Date();

  sheet.appendRow([
    no,
    data['주문일자'] || Utilities.formatDate(now, 'Asia/Seoul', 'M월d일'),
    Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    data['주문채널']     || '웹주문',
    data['보내는분성명']  || '',
    data['보내는분전화번호'] || '',
    data['보내는분주소']  || '',
    data['받는분성명']    || '',
    data['받는분전화번호']  || '',
    data['받는분기타연락처'] || '',
    data['받는분주소']    || '',
    data['품목명']       || '',
    Number(data['박스수량']) || 1,
    Number(data['단가'])    || 0,
    Number(data['합계'])    || 0,
    data['배송메시지']    || '',
    '접수대기',  // 초기 상태
    0,           // 입금확인액
    ''           // 비고
  ]);

  return jsonOk({ success: true, message: '주문이 접수되었습니다' });
}

// ================================================================
// 주문 목록 조회 (관리자)
// ================================================================
function getOrders(params) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);

  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonOk({ success: true, orders: [], total: 0 });
  }

  var values  = sheet.getDataRange().getValues();
  var headers = values[0];
  var orders  = [];

  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var val = values[i][j];
      // Date 객체는 문자열로 변환
      if (val instanceof Date) val = Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      row[headers[j]] = val;
    }
    row['_row'] = i + 1; // 실제 시트 행 번호 (업데이트 시 사용)
    orders.push(row);
  }

  // 검색 필터
  if (params.search) {
    var q = String(params.search).toLowerCase();
    orders = orders.filter(function(o) {
      var combined = [o['보내는분성명'], o['받는분성명'], o['받는분전화번호'], o['받는분주소']].join(' ').toLowerCase();
      return combined.indexOf(q) >= 0;
    });
  }

  // 상태 필터
  if (params.status && params.status !== 'all') {
    orders = orders.filter(function(o) { return o['상태'] === params.status; });
  }

  // 최신순 정렬
  orders.sort(function(a, b) {
    return String(b['접수시각']).localeCompare(String(a['접수시각']));
  });

  // 통계
  var stats = {
    total:      orders.length,
    totalBoxes: orders.reduce(function(s, o) { return s + (Number(o['박스수량']) || 0); }, 0),
    totalAmt:   orders.reduce(function(s, o) { return s + (Number(o['합계']) || 0); }, 0),
    totalPaid:  orders.reduce(function(s, o) { return s + (Number(o['입금확인액']) || 0); }, 0),
    unpaid:     orders.filter(function(o) { return o['상태'] !== '취소' && (Number(o['입금확인액']) || 0) < (Number(o['합계']) || 0); }).length
  };

  return jsonOk({ success: true, orders: orders, stats: stats });
}

// ================================================================
// 주문 필드 업데이트 (관리자)  — GET 파라미터로 처리
// ================================================================
function updateOrderField(params) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET);
  if (!sheet) return jsonOk({ success: false, message: '시트가 없습니다' });

  var rowNum = parseInt(params.row, 10);
  if (!rowNum || rowNum < 2) return jsonOk({ success: false, message: '잘못된 행 번호입니다' });

  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colIndex = headers.indexOf(params.field) + 1; // 1-based
  if (colIndex <= 0) return jsonOk({ success: false, message: '필드를 찾을 수 없습니다: ' + params.field });

  var value = params.value;
  // 숫자 필드
  if (params.field === '입금확인액' || params.field === '박스수량') {
    value = Number(value) || 0;
  }

  sheet.getRange(rowNum, colIndex).setValue(value);
  return jsonOk({ success: true, message: '업데이트 완료' });
}

// ================================================================
// 문의 저장
// ================================================================
function saveContact(data) {
  var ss    = getSpreadsheet();
  var sheet = getOrCreateSheet(ss, CONFIG.CONTACTS_SHEET);

  if (sheet.getLastRow() === 0) {
    var headers = ['접수시각', '성함', '연락처', '문의유형', '문의내용', '처리상태'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground('#1E3A2F')
         .setFontColor('#FFFFFF')
         .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    data['성함']   || '',
    data['연락처'] || '',
    data['문의유형'] || '',
    data['문의내용'] || '',
    '미처리'
  ]);

  return jsonOk({ success: true, message: '문의가 접수되었습니다' });
}

// ================================================================
// 문의 목록 조회 (관리자)
// ================================================================
function getContacts(params) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);

  if (!sheet || sheet.getLastRow() <= 1) {
    return jsonOk({ success: true, contacts: [] });
  }

  var values  = sheet.getDataRange().getValues();
  var headers = values[0];
  var contacts = [];

  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var val = values[i][j];
      if (val instanceof Date) val = Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      row[headers[j]] = val;
    }
    row['_row'] = i + 1;
    contacts.push(row);
  }

  contacts.sort(function(a, b) {
    return String(b['접수시각']).localeCompare(String(a['접수시각']));
  });

  return jsonOk({ success: true, contacts: contacts });
}

// ================================================================
// 유틸리티
// ================================================================
function getSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function jsonOk(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
