// ════════════════════════════════════════════════════════════
// 透析コンソール管理システム - GAS コード
// ════════════════════════════════════════════════════════════
// ★ 設定：スプレッドシートIDを貼り付けてください
const SS_ID = 'YOUR_SPREADSHEET_ID';

// ★ シート名（実際のシート名に合わせてください）
const LEASE_SHEET   = 'リース管理';
const INSPECT_SHEET = '点検記録';
const REPAIR_SHEET  = '部品交換・修理記録';

// ════════════════════════════════════════════════════════════
// 列番号の定義（0始まり）
// ════════════════════════════════════════════════════════════

// リース管理シート
// A=0: 機器ID, B=1: 拠点, C=2: カテゴリ, D=3: 機械番号,
// E=4: 型番, F=5: メーカー, G=6: リース会社,
// H=7: リース開始日, I=8: リース終了日, J=9: 買替推奨年,
// K=10: 稼働状況（"稼働中"/"要注意"/"要対応"）

// 点検記録シート
// A=0: 機器ID, B=1: 拠点, C=2: 型番,
// D=3: 点検日, E=4: 点検種別, F=5: 点検内容,
// G=6: 担当者, H=7: 次回点検予定日

// 部品交換・修理記録シート
// A=0: 機器ID, B=1: 拠点, C=2: 機械番号, D=3: 型番,
// E=4: 対応日, F=5: 記録種別, G=6: 交換部品・対応内容,
// H=7: 理由・詳細, I=8: 対応業者, J=9: 担当者,
// K=10: 次回予定日  ← 今回追加（スプレッドシートのK列）

// ════════════════════════════════════════════════════════════

function doGet(e) {
  const p = e.parameter;
  try {
    let result;
    switch (p.action) {
      case 'getList':        result = getList();           break;
      case 'getDevice':      result = getDevice(p.id);     break;
      case 'getHistory':     result = getHistory(p.id);    break;
      case 'getSchedule':    result = getSchedule();       break;
      case 'getAllHistory':   result = getAllHistory();      break;
      case 'getReplaceYears':result = getReplaceYears();   break;
      default: result = { error: 'Unknown action: ' + p.action };
    }
    return jsonRes(result);
  } catch (err) {
    return jsonRes({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    let result;
    switch (body.action) {
      case 'addRecord':    result = addRecord(body);    break;
      case 'updateStatus': result = updateStatus(body); break;
      default: result = { error: 'Unknown action: ' + body.action };
    }
    return jsonRes(result);
  } catch (err) {
    return jsonRes({ error: err.toString() });
  }
}

function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 日付フォーマット ──
function fmtDate(val) {
  if (!val) return '';
  try {
    return Utilities.formatDate(new Date(val), 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) {
    return String(val);
  }
}

// ── リース管理シートの行 → 機器オブジェクト ──
function rowToDevice(row) {
  const statusMap = { '稼働中': 'ok', '要注意': 'warn', '要対応': 'alert' };
  const rawStatus = row[10] || '稼働中';
  return {
    id:           String(row[0] || ''),
    site:         String(row[1] || ''),
    category:     String(row[2] || ''),
    machineNo:    row[3] !== '' && row[3] !== null && row[3] !== undefined ? row[3] : '-',
    model:        String(row[4] || ''),
    maker:        String(row[5] || ''),
    leaseCompany: String(row[6] || ''),
    startDate:    fmtDate(row[7]),
    endDate:      fmtDate(row[8]),
    replaceYear:  row[9] ? Number(row[9]) : '',
    status:       statusMap[rawStatus] || 'ok',
    statusLabel:  rawStatus,
  };
}

// ── 機器一覧 ──
function getList() {
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEASE_SHEET);
  const data  = sheet.getDataRange().getValues();
  const devices = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    devices.push(rowToDevice(data[i]));
  }
  return { devices };
}

// ── 機器詳細 ──
function getDevice(id) {
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEASE_SHEET);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return rowToDevice(data[i]);
  }
  return {};
}

// ── 機器履歴（点検 + 修理を結合・日付降順） ──
function getHistory(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const records = [];

  // 点検記録
  const iData = ss.getSheetByName(INSPECT_SHEET).getDataRange().getValues();
  for (let i = 1; i < iData.length; i++) {
    const r = iData[i];
    if (String(r[0]) !== String(id)) continue;
    records.push({
      type:     String(r[4] || '定期点検'),
      date:     fmtDate(r[3]),
      content:  String(r[5] || ''),
      staff:    String(r[6] || ''),
      nextDate: fmtDate(r[7]),
    });
  }

  // 部品交換・修理記録
  const rData = ss.getSheetByName(REPAIR_SHEET).getDataRange().getValues();
  for (let i = 1; i < rData.length; i++) {
    const r = rData[i];
    if (String(r[0]) !== String(id)) continue;
    records.push({
      type:     String(r[5] || '部品交換'),
      date:     fmtDate(r[4]),
      parts:    String(r[6] || ''),
      content:  String(r[7] || ''),
      vendor:   String(r[8] || ''),
      staff:    String(r[9] || ''),
      nextDate: fmtDate(r[10]),
    });
  }

  records.sort((a, b) => new Date(b.date) - new Date(a.date));
  return { records };
}

// ── 今月の予定 ──
// 各シートの「次回予定日」が当月のもので、かつ完了していないものを返す
function getSchedule() {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const now  = new Date();
  const thisYear  = now.getFullYear();
  const thisMonth = now.getMonth(); // 0始まり

  // 機器情報マップ
  const devMap = {};
  ss.getSheetByName(LEASE_SHEET).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    devMap[String(r[0])] = { machineNo: r[3] || '-', model: String(r[4] || ''), site: String(r[1] || '') };
  });

  const schedules = [];

  // ─ 点検記録の次回予定 ─
  const iRows = ss.getSheetByName(INSPECT_SHEET).getDataRange().getValues().slice(1);

  // 完了判定用：機器IDごとの最終点検日
  const lastInspect = {};
  iRows.forEach(r => {
    const devId = String(r[0]);
    const d = r[3] ? new Date(r[3]) : null;
    if (d && (!lastInspect[devId] || d > lastInspect[devId])) lastInspect[devId] = d;
  });

  iRows.forEach(r => {
    if (!r[7]) return;
    const next = new Date(r[7]);
    if (next.getFullYear() !== thisYear || next.getMonth() !== thisMonth) return;
    const devId = String(r[0]);
    // 予定日以降に新しい記録があれば完了済みとみなす
    if (lastInspect[devId] && lastInspect[devId] >= next) return;
    const dev = devMap[devId] || {};
    schedules.push({
      deviceId:  devId,
      machineNo: dev.machineNo,
      model:     dev.model,
      site:      dev.site,
      type:      String(r[4] || '定期点検'),
      nextDate:  fmtDate(next),
      content:   String(r[5] || ''),
    });
  });

  // ─ 部品交換・修理記録の次回予定 ─
  const rRows = ss.getSheetByName(REPAIR_SHEET).getDataRange().getValues().slice(1);

  const lastRepair = {};
  rRows.forEach(r => {
    const devId = String(r[0]);
    const d = r[4] ? new Date(r[4]) : null;
    if (d && (!lastRepair[devId] || d > lastRepair[devId])) lastRepair[devId] = d;
  });

  rRows.forEach(r => {
    if (!r[10]) return;
    const next = new Date(r[10]);
    if (next.getFullYear() !== thisYear || next.getMonth() !== thisMonth) return;
    const devId = String(r[0]);
    if (lastRepair[devId] && lastRepair[devId] >= next) return;
    const dev = devMap[devId] || {};
    schedules.push({
      deviceId:  devId,
      machineNo: dev.machineNo,
      model:     dev.model,
      site:      dev.site,
      type:      String(r[5] || '部品交換'),
      nextDate:  fmtDate(next),
      parts:     String(r[6] || ''),
      content:   String(r[7] || ''),
    });
  });

  schedules.sort((a, b) => new Date(a.nextDate) - new Date(b.nextDate));
  return { schedules };
}

// ── 全機器履歴（履歴一覧ページ用） ──
function getAllHistory() {
  const ss = SpreadsheetApp.openById(SS_ID);

  // 機械番号マップ
  const machineNoMap = {};
  ss.getSheetByName(LEASE_SHEET).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    machineNoMap[String(r[0])] = r[3] || '-';
  });

  const records = [];

  ss.getSheetByName(INSPECT_SHEET).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[3]) return;
    records.push({
      deviceId:  String(r[0]),
      machineNo: machineNoMap[String(r[0])] || '-',
      type:      String(r[4] || '定期点検'),
      date:      fmtDate(r[3]),
      content:   String(r[5] || ''),
      staff:     String(r[6] || ''),
      nextDate:  fmtDate(r[7]),
    });
  });

  ss.getSheetByName(REPAIR_SHEET).getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0] || !r[4]) return;
    records.push({
      deviceId:  String(r[0]),
      machineNo: machineNoMap[String(r[0])] || '-',
      type:      String(r[5] || '部品交換'),
      date:      fmtDate(r[4]),
      parts:     String(r[6] || ''),
      content:   String(r[7] || ''),
      vendor:    String(r[8] || ''),
      staff:     String(r[9] || ''),
      nextDate:  fmtDate(r[10]),
    });
  });

  records.sort((a, b) => new Date(b.date) - new Date(a.date));
  return { records };
}

// ── 買替推奨年一覧 ──
function getReplaceYears() {
  const data = SpreadsheetApp.openById(SS_ID)
    .getSheetByName(LEASE_SHEET).getDataRange().getValues().slice(1);
  const devices = [];
  data.forEach(r => {
    if (!r[0] || !r[9]) return;
    devices.push({
      id:           String(r[0]),
      site:         String(r[1] || ''),
      machineNo:    r[3] !== '' ? r[3] : '-',
      model:        String(r[4] || ''),
      leaseCompany: String(r[6] || ''),
      startDate:    fmtDate(r[7]),
      replaceYear:  Number(r[9]),
    });
  });
  devices.sort((a, b) => (a.replaceYear || 9999) - (b.replaceYear || 9999));
  return { devices };
}

// ── 記録追加 ──
function addRecord(body) {
  const ss = SpreadsheetApp.openById(SS_ID);

  if (body.recordType === '点検' || body.inspectType) {
    // 点検記録シートへ追記
    ss.getSheetByName(INSPECT_SHEET).appendRow([
      body.deviceId  || '',
      body.site      || '',
      body.model     || '',
      body.date      || '',
      body.inspectType || body.recordType || '定期点検',
      body.content   || '',
      body.staff     || '',
      body.nextDate  || '',
    ]);
  } else {
    // 部品交換・修理記録シートへ追記（K列=次回予定日）
    ss.getSheetByName(REPAIR_SHEET).appendRow([
      body.deviceId  || '',
      body.site      || '',
      body.machineNo || '',
      body.model     || '',
      body.date      || '',
      body.recordType|| '部品交換',
      body.parts     || '',
      body.content   || '',
      body.vendor    || '',
      body.staff     || '',
      body.nextDate  || '',  // K列
    ]);
  }
  return { success: true };
}

// ── 稼働状況更新（リース管理シートK列） ──
function updateStatus(body) {
  const sheet = SpreadsheetApp.openById(SS_ID).getSheetByName(LEASE_SHEET);
  const data  = sheet.getDataRange().getValues();
  const labelMap = { ok: '稼働中', warn: '要注意', alert: '要対応' };
  const newLabel = labelMap[body.status] || '稼働中';

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.deviceId)) {
      sheet.getRange(i + 1, 11).setValue(newLabel); // 11列目 = K列
      return { success: true };
    }
  }
  return { error: '機器が見つかりません: ' + body.deviceId };
}
