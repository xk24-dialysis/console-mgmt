// ════════════════════════════════════════════════════════════
// 透析コンソール管理システム - GAS コード
// スプレッドシートから開いたGASはIDなしで自動認識します。
// ════════════════════════════════════════════════════════════

// ★ 機器台帳シート（複数拠点を配列で指定）
const DEVICE_SHEETS = ['機器台帳_川奈', '機器台帳_伊豆高原'];
// ★ 後方互換用（updateStatus等で使用）
const LEASE_SHEET   = '機器台帳_川奈';
const INSPECT_SHEET = '定期点検記録';
const REPAIR_SHEET  = '部品交換・修理記録';

// ════════════════════════════════════════════════════════════
// ヘッダー名の候補（実際のスプレッドシートの表記に合わせて追加可）
// ════════════════════════════════════════════════════════════
const HEADER_ALIASES = {
  id:           ['機器ID', 'ID', 'デバイスID', '機器id'],
  site:         ['拠点', '施設', 'サイト'],
  category:     ['カテゴリ', '区分', '種別'],
  machineNo:    ['機械番号', '機番', 'No', 'NO', '号機'],
  model:        ['型式', '型番', 'モデル', 'MODEL', '機種'],
  maker:        ['メーカー', '製造会社', 'メーカ'],
  leaseCompany: ['リース会社', 'リース先', '会社'],
  startDate:    ['リース開始日', '導入日', '開始日', 'リース開始'],
  endDate:      ['リース終了日', '終了日', 'リース満了'],
  replaceYear:  ['買替推奨年', '買替年', '推奨買替年', '更新推奨年'],
  status:       ['稼働状況', 'ステータス', '状態', '状況'],
  // 点検記録シート
  inspDate:     ['点検日', '実施日', '日付'],
  inspType:     ['点検種別', '種別', 'タイプ'],
  inspContent:  ['点検内容', '内容', '所見'],
  staff:        ['担当者', '担当', 'スタッフ'],
  nextDate:     ['次回予定日', '次回点検予定日', '次回日', '予定日'],
  // 部品交換・修理記録シート
  repairDate:   ['対応日', '実施日', '日付'],
  repairType:   ['記録種別', '種別', 'タイプ'],
  parts:        ['交換部品', '部品', '対応内容', '交換部品・対応内容'],
  repairContent:['理由・詳細', '詳細', '内容', '理由'],
  vendor:       ['対応業者', '業者', 'ベンダー'],
  repairNextDate:['次回予定日', '次回日', '予定日'],
};

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

// ── ヘッダー行からカラムマップを作成 ──
function buildColMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    if (h !== null && h !== undefined && h !== '') {
      map[String(h).trim()] = i;
    }
  });
  return map;
}

// ── 先頭10行のうちヘッダーらしき行を自動検出 ──
// 既知のエイリアスに最も多くマッチする行をヘッダーと判定する
function findHeaderRowIndex(data) {
  const allAliases = new Set(Object.values(HEADER_ALIASES).flat());
  let bestIdx = 0;
  let bestScore = 0;
  const limit = Math.min(10, data.length);
  for (let i = 0; i < limit; i++) {
    let score = 0;
    for (const cell of data[i]) {
      if (allAliases.has(String(cell).trim())) score++;
    }
    if (score > bestScore) { bestScore = score; bestIdx = i; }
  }
  return bestIdx;
}

// ── カラムマップからフィールド値を取得（候補名を順番に試す）──
function getField(row, colMap, fieldKey) {
  const aliases = HEADER_ALIASES[fieldKey] || [];
  for (const alias of aliases) {
    if (colMap[alias] !== undefined) {
      const val = row[colMap[alias]];
      if (val !== null && val !== undefined && val !== '') return val;
    }
  }
  return '';
}

// ── リース管理シートの行 → 機器オブジェクト ──
function rowToDevice(row, colMap) {
  const statusMap = { '稼働中': 'ok', '要注意': 'warn', '要対応': 'alert' };
  const rawStatus = String(getField(row, colMap, 'status') || '稼働中');
  const machineNo = getField(row, colMap, 'machineNo');

  return {
    id:           String(getField(row, colMap, 'id')           || row[0] || ''),
    site:         String(getField(row, colMap, 'site')         || ''),
    category:     String(getField(row, colMap, 'category')     || ''),
    machineNo:    (machineNo !== '' && machineNo !== null) ? machineNo : '-',
    model:        String(getField(row, colMap, 'model')        || ''),
    maker:        String(getField(row, colMap, 'maker')        || ''),
    leaseCompany: String(getField(row, colMap, 'leaseCompany') || ''),
    startDate:    fmtDate(getField(row, colMap, 'startDate')),
    endDate:      fmtDate(getField(row, colMap, 'endDate')),
    replaceYear:  getField(row, colMap, 'replaceYear') ? Number(getField(row, colMap, 'replaceYear')) : '',
    status:       statusMap[rawStatus] || 'ok',
    statusLabel:  rawStatus,
  };
}

// ── 全機器台帳シートからデバイス行を取得するヘルパー ──
function getAllDeviceRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];
  for (const sheetName of DEVICE_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    // メモ行などを飛ばしてヘッダー行を自動検出
    const headerIdx = findHeaderRowIndex(data);
    const colMap    = buildColMap(data[headerIdx]);
    const idCol     = findIdCol(colMap);

    for (let i = headerIdx + 1; i < data.length; i++) {
      const id = String(data[i][idCol] || '').trim();
      // 空行・ヘッダーの繰り返し行はスキップ
      if (!id || id === '機器ID' || id === 'ID') continue;
      result.push({ row: data[i], colMap, sheetName, rowIndex: i + 1 });
    }
  }
  return result;
}

// ── 機器一覧 ──
function getList() {
  const devices = getAllDeviceRows().map(({ row, colMap }) => rowToDevice(row, colMap));
  return { devices };
}

// ── 機器詳細 ──
function getDevice(id) {
  for (const { row, colMap } of getAllDeviceRows()) {
    if (String(getField(row, colMap, 'id') || row[0]) === String(id)) {
      return rowToDevice(row, colMap);
    }
  }
  return {};
}

// ── ID列のインデックスを取得 ──
function findIdCol(colMap) {
  for (const alias of HEADER_ALIASES.id) {
    if (colMap[alias] !== undefined) return colMap[alias];
  }
  return 0; // デフォルトはA列
}

// ── 機器履歴（点検 + 修理を結合・日付降順） ──
function getHistory(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const records = [];

  // 点検記録
  const iSheet = ss.getSheetByName(INSPECT_SHEET);
  const iData  = iSheet.getDataRange().getValues();
  if (iData.length > 1) {
    const iMap   = buildColMap(iData[0]);
    const iIdCol = findIdCol(iMap);
    for (let i = 1; i < iData.length; i++) {
      if (String(iData[i][iIdCol]) !== String(id)) continue;
      records.push({
        type:     String(getField(iData[i], iMap, 'inspType')    || '定期点検'),
        date:     fmtDate(getField(iData[i], iMap, 'inspDate')),
        content:  String(getField(iData[i], iMap, 'inspContent') || ''),
        staff:    String(getField(iData[i], iMap, 'staff')       || ''),
        nextDate: fmtDate(getField(iData[i], iMap, 'nextDate')),
      });
    }
  }

  // 部品交換・修理記録
  const rSheet = ss.getSheetByName(REPAIR_SHEET);
  const rData  = rSheet.getDataRange().getValues();
  if (rData.length > 1) {
    const rMap   = buildColMap(rData[0]);
    const rIdCol = findIdCol(rMap);
    for (let i = 1; i < rData.length; i++) {
      if (String(rData[i][rIdCol]) !== String(id)) continue;
      records.push({
        type:     String(getField(rData[i], rMap, 'repairType')    || '部品交換'),
        date:     fmtDate(getField(rData[i], rMap, 'repairDate')),
        parts:    String(getField(rData[i], rMap, 'parts')         || ''),
        content:  String(getField(rData[i], rMap, 'repairContent') || ''),
        vendor:   String(getField(rData[i], rMap, 'vendor')        || ''),
        staff:    String(getField(rData[i], rMap, 'staff')         || ''),
        nextDate: fmtDate(getField(rData[i], rMap, 'repairNextDate')),
      });
    }
  }

  records.sort((a, b) => new Date(b.date) - new Date(a.date));
  return { records };
}

// ── 今月の予定 ──
function getSchedule() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const thisYear  = now.getFullYear();
  const thisMonth = now.getMonth();

  // 機器情報マップ（全機器台帳シートから構築）
  const devMap = buildDevMap();

  const schedules = [];

  // 点検記録の次回予定
  const iData  = ss.getSheetByName(INSPECT_SHEET).getDataRange().getValues();
  const iMap   = buildColMap(iData[0]);
  const iIdCol = findIdCol(iMap);
  const lastInspect = {};
  for (let i = 1; i < iData.length; i++) {
    const devId = String(iData[i][iIdCol] || '');
    const d = getField(iData[i], iMap, 'inspDate') ? new Date(getField(iData[i], iMap, 'inspDate')) : null;
    if (d && (!lastInspect[devId] || d > lastInspect[devId])) lastInspect[devId] = d;
  }
  for (let i = 1; i < iData.length; i++) {
    const nextRaw = getField(iData[i], iMap, 'nextDate');
    if (!nextRaw) continue;
    const next = new Date(nextRaw);
    if (next.getFullYear() !== thisYear || next.getMonth() !== thisMonth) continue;
    const devId = String(iData[i][iIdCol] || '');
    if (lastInspect[devId] && lastInspect[devId] >= next) continue;
    const dev = devMap[devId] || {};
    schedules.push({
      deviceId:  devId,
      machineNo: dev.machineNo,
      model:     dev.model,
      site:      dev.site,
      type:      String(getField(iData[i], iMap, 'inspType') || '定期点検'),
      nextDate:  fmtDate(next),
      content:   String(getField(iData[i], iMap, 'inspContent') || ''),
    });
  }

  // 部品交換・修理記録の次回予定
  const rData  = ss.getSheetByName(REPAIR_SHEET).getDataRange().getValues();
  const rMap   = buildColMap(rData[0]);
  const rIdCol = findIdCol(rMap);
  const lastRepair = {};
  for (let i = 1; i < rData.length; i++) {
    const devId = String(rData[i][rIdCol] || '');
    const d = getField(rData[i], rMap, 'repairDate') ? new Date(getField(rData[i], rMap, 'repairDate')) : null;
    if (d && (!lastRepair[devId] || d > lastRepair[devId])) lastRepair[devId] = d;
  }
  for (let i = 1; i < rData.length; i++) {
    const nextRaw = getField(rData[i], rMap, 'repairNextDate');
    if (!nextRaw) continue;
    const next = new Date(nextRaw);
    if (next.getFullYear() !== thisYear || next.getMonth() !== thisMonth) continue;
    const devId = String(rData[i][rIdCol] || '');
    if (lastRepair[devId] && lastRepair[devId] >= next) continue;
    const dev = devMap[devId] || {};
    schedules.push({
      deviceId:  devId,
      machineNo: dev.machineNo,
      model:     dev.model,
      site:      dev.site,
      type:      String(getField(rData[i], rMap, 'repairType') || '部品交換'),
      nextDate:  fmtDate(next),
      parts:     String(getField(rData[i], rMap, 'parts')      || ''),
      content:   String(getField(rData[i], rMap, 'repairContent') || ''),
    });
  }

  schedules.sort((a, b) => new Date(a.nextDate) - new Date(b.nextDate));
  return { schedules };
}

// ── 機械番号マップを全機器台帳から構築 ──
function buildMachineNoMap() {
  const map = {};
  for (const { row, colMap } of getAllDeviceRows()) {
    const id = String(getField(row, colMap, 'id') || row[0] || '');
    if (id) map[id] = getField(row, colMap, 'machineNo') || '-';
  }
  return map;
}

// ── devMapを全機器台帳から構築（getSchedule用）──
function buildDevMap() {
  const map = {};
  for (const { row, colMap } of getAllDeviceRows()) {
    const id = String(getField(row, colMap, 'id') || row[0] || '');
    if (id) map[id] = {
      machineNo: getField(row, colMap, 'machineNo') || '-',
      model:     String(getField(row, colMap, 'model') || ''),
      site:      String(getField(row, colMap, 'site')  || ''),
    };
  }
  return map;
}

// ── 全機器履歴 ──
function getAllHistory() {
  const machineNoMap = buildMachineNoMap();

  const records = [];

  const iData  = ss.getSheetByName(INSPECT_SHEET).getDataRange().getValues();
  const iMap   = buildColMap(iData[0]);
  const iIdCol = findIdCol(iMap);
  for (let i = 1; i < iData.length; i++) {
    const devId = String(iData[i][iIdCol] || '');
    if (!devId || !getField(iData[i], iMap, 'inspDate')) continue;
    records.push({
      deviceId:  devId,
      machineNo: machineNoMap[devId] || '-',
      type:      String(getField(iData[i], iMap, 'inspType')    || '定期点検'),
      date:      fmtDate(getField(iData[i], iMap, 'inspDate')),
      content:   String(getField(iData[i], iMap, 'inspContent') || ''),
      staff:     String(getField(iData[i], iMap, 'staff')       || ''),
      nextDate:  fmtDate(getField(iData[i], iMap, 'nextDate')),
    });
  }

  const rData  = ss.getSheetByName(REPAIR_SHEET).getDataRange().getValues();
  const rMap   = buildColMap(rData[0]);
  const rIdCol = findIdCol(rMap);
  for (let i = 1; i < rData.length; i++) {
    const devId = String(rData[i][rIdCol] || '');
    if (!devId || !getField(rData[i], rMap, 'repairDate')) continue;
    records.push({
      deviceId:  devId,
      machineNo: machineNoMap[devId] || '-',
      type:      String(getField(rData[i], rMap, 'repairType')    || '部品交換'),
      date:      fmtDate(getField(rData[i], rMap, 'repairDate')),
      parts:     String(getField(rData[i], rMap, 'parts')         || ''),
      content:   String(getField(rData[i], rMap, 'repairContent') || ''),
      vendor:    String(getField(rData[i], rMap, 'vendor')        || ''),
      staff:     String(getField(rData[i], rMap, 'staff')         || ''),
      nextDate:  fmtDate(getField(rData[i], rMap, 'repairNextDate')),
    });
  }

  records.sort((a, b) => new Date(b.date) - new Date(a.date));
  return { records };
}

// ── 買替推奨年一覧 ──
function getReplaceYears() {
  const devices = [];
  for (const { row, colMap } of getAllDeviceRows()) {
    const id          = String(getField(row, colMap, 'id') || row[0] || '');
    const replaceYear = getField(row, colMap, 'replaceYear');
    if (!id || !replaceYear) continue;
    devices.push({
      id:           id,
      site:         String(getField(row, colMap, 'site')         || ''),
      machineNo:    getField(row, colMap, 'machineNo')            || '-',
      model:        String(getField(row, colMap, 'model')        || ''),
      leaseCompany: String(getField(row, colMap, 'leaseCompany') || ''),
      startDate:    fmtDate(getField(row, colMap, 'startDate')),
      replaceYear:  Number(replaceYear),
    });
  }
  devices.sort((a, b) => (a.replaceYear || 9999) - (b.replaceYear || 9999));
  return { devices };
}

// ── 記録追加 ──
function addRecord(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (body.recordType === '点検' || body.inspectType) {
    const sheet  = ss.getSheetByName(INSPECT_SHEET);
    const data   = sheet.getDataRange().getValues();
    const colMap = buildColMap(data[0]);

    // ヘッダーに合わせて追記行を組み立て
    const newRow = new Array(data[0].length).fill('');
    const set = (fieldKey, val) => {
      for (const alias of (HEADER_ALIASES[fieldKey] || [])) {
        if (colMap[alias] !== undefined) { newRow[colMap[alias]] = val; return; }
      }
    };
    set('id',          body.deviceId   || '');
    set('site',        body.site        || '');
    set('model',       body.model       || '');
    set('inspDate',    body.date        || '');
    set('inspType',    body.inspectType || body.recordType || '定期点検');
    set('inspContent', body.content     || '');
    set('staff',       body.staff       || '');
    set('nextDate',    body.nextDate    || '');
    sheet.appendRow(newRow);

  } else {
    const sheet  = ss.getSheetByName(REPAIR_SHEET);
    const data   = sheet.getDataRange().getValues();
    const colMap = buildColMap(data[0]);

    const newRow = new Array(data[0].length).fill('');
    const set = (fieldKey, val) => {
      for (const alias of (HEADER_ALIASES[fieldKey] || [])) {
        if (colMap[alias] !== undefined) { newRow[colMap[alias]] = val; return; }
      }
    };
    set('id',              body.deviceId  || '');
    set('site',            body.site       || '');
    set('machineNo',       body.machineNo  || '');
    set('model',           body.model      || '');
    set('repairDate',      body.date       || '');
    set('repairType',      body.recordType || '部品交換');
    set('parts',           body.parts      || '');
    set('repairContent',   body.content    || '');
    set('vendor',          body.vendor     || '');
    set('staff',           body.staff      || '');
    set('repairNextDate',  body.nextDate   || '');
    sheet.appendRow(newRow);
  }

  return { success: true };
}

// ── 稼働状況更新（全機器台帳シートを検索して更新）──
function updateStatus(body) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const labelMap = { ok: '稼働中', warn: '要注意', alert: '要対応' };
  const newLabel = labelMap[body.status] || '稼働中';

  for (const sheetName of DEVICE_SHEETS) {
    const sheet  = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    const data   = sheet.getDataRange().getValues();
    const colMap = buildColMap(data[0]);
    const idCol  = findIdCol(colMap);

    let statusCol = -1;
    for (const alias of HEADER_ALIASES.status) {
      if (colMap[alias] !== undefined) { statusCol = colMap[alias]; break; }
    }
    if (statusCol === -1) continue;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(body.deviceId)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newLabel);
        return { success: true };
      }
    }
  }
  return { error: '機器が見つかりません: ' + body.deviceId };
}
