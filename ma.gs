function doGet(e) {
  var action   = e && e.parameter && e.parameter.action;
  var callback = e && e.parameter && e.parameter.callback;

  function jsonOut(data) {
    var json = JSON.stringify(data);
    var body = callback ? callback + '(' + json + ')' : json;
    var mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
    return ContentService.createTextOutput(body).setMimeType(mime);
  }

  if (action === 'getDashboardData')    return jsonOut(getDashboardData());
  if (action === 'getDashboardComments') return jsonOut(getDashboardComments());
  if (action === 'getCEOProjects')      return jsonOut(getCEOProjectsData());
  if (action === 'getTOActions')        return jsonOut(getTOActionsData());
  if (action === 'getMeetingMinutes')   return jsonOut(getMeetingMinutesData());
  if (action === 'getRiskSOS')          return jsonOut(getRiskSOSData());
  if (action === 'getCEODecisions')     return jsonOut(getCEODecisionsData());
  if (action === 'getDataQualityFull')  return jsonOut(getDataQualityFullData());

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('GSP NEXT 30 - CEO Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var payload = body.payload || {};
    var result;

    if (action === 'saveDashboardComment') {
      result = saveDashboardComment(payload);
    } else if (action === 'likeDashboardComment') {
      result = likeDashboardComment(payload);
    } else if (action === 'deleteDashboardComment') {
      result = deleteDashboardComment(payload);
    } else {
      result = { success: false, error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* =====================================================
   CONFIG
===================================================== */

const DASHBOARD_SHEETS = [
  { name: '01_Brand', label: '01_Brand', display: '01 Thương hiệu' },
  { name: '02_Sales', label: '02_Sales', display: '02 Kinh doanh' },
  { name: '03_Operation', label: '03_Operation', display: '03 Vận hành' },
  { name: '04_Digital_AI', label: '04_Digital_AI', display: '04 Digital / AI' },
  { name: '05_People_Culture', label: '05_People_Culture', display: '05 Con người / Văn hóa' },
  { name: '06_System_KPI', label: '06_System_KPI', display: '06 Hệ thống / KPI' }
];

const COMMENT_SHEET_NAME = 'CEO_Comments';
const COMMENT_FOLDER_NAME = 'GSP_NEXT30_Dashboard_Comments';

/* =====================================================
   DASHBOARD DATA
===================================================== */

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rows = [];

  DASHBOARD_SHEETS.forEach(function(config) {
    const sheet = ss.getSheetByName(config.name);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 4) return;

    const headers = values[2].map(function(h) {
      return String(h || '').trim();
    });

    const sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();

    values.slice(3).forEach(function(row, index) {
      const obj = {};

      headers.forEach(function(h, i) {
        if (h) obj[h] = row[i];
      });

      const rowNumber = index + 4;
      const rowId = getVal(obj, ['Mã dòng (ID)', 'Mã dòng', 'ID']);

      if (!hasText(rowId)) return;

      const rowUrl = ss.getUrl() + '#gid=' + sheet.getSheetId() + '&range=A' + rowNumber + ':Z' + rowNumber;

      const rowObj = {
        RowId: String(rowId || ''),
        Workstream: config.label,
        WorkstreamDisplay: config.display,
        SourceSheet: config.name,
        SourceUrl: sheetUrl,
        RowUrl: rowUrl,
        Unit: String(getVal(obj, ['Đơn vị / Khối', 'Đơn vị', 'Khối', 'Bộ phận', 'Phòng ban']) || ''),
        PIC: String(getVal(obj, ['PIC Lead', 'PIC', 'Đầu mối']) || ''),
        Action: String(getVal(obj, ['Kế hoạch hành động', 'Action Plan', 'Action', 'Việc cần làm', 'Cam kết hành động', 'Hành động', 'Nhiệm vụ']) || ''),
        Owner: String(getVal(obj, ['Owner', 'Người phụ trách', 'Người chịu trách nhiệm']) || ''),
        Deadline: formatDate(getVal(obj, ['Deadline', 'Hạn hoàn thành', 'Thời hạn'])),
        Status: normalizeStatus(getVal(obj, ['Status', 'Trạng thái', 'Tình trạng'])),
        RAG: normalizeRag(getVal(obj, ['RAG', 'Đèn RAG', 'Đèn báo', 'Cảnh báo'])),
        Risk: String(getVal(obj, ['Risk / SOS', 'Risk', 'SOS', 'Rủi ro', 'Blocker', 'Vướng mắc']) || ''),
        CEODecision: String(getVal(obj, ['CEO Decision Needed', 'CEO cần chốt', 'CEO quyết định', 'Cần CEO chốt', 'Cần anh Vinh chốt']) || ''),
        PMOComment: String(getVal(obj, ['PMO Comment', 'Ghi chú PMO', 'Nhận xét PMO']) || '')
      };

      rowObj.IsOverdue = isPastDeadline(rowObj.Deadline, rowObj.Status);
      rowObj.DataQuality = calcRowDataQuality_(rowObj);
      rowObj.PriorityScore = calcPriorityScore_(rowObj);

      rows.push(rowObj);
    });
  });

  const doneRows = rows.filter(r => r.Status === 'Done');
  const inProgressRows = rows.filter(r => r.Status === 'In Progress');
  const overdueRows = rows.filter(r => r.Status === 'Overdue' || r.IsOverdue);
  const blockedRows = rows.filter(r => r.Status === 'Blocked');

  const greenRows = rows.filter(r => r.RAG === 'Green');
  const amberRows = rows.filter(r => r.RAG === 'Amber');
  const redRows = rows.filter(r => r.RAG === 'Red');

  const riskRows = rows.filter(r => hasText(r.Risk));

  const sosRows = rows.filter(function(r) {
    const riskText = String(r.Risk || '').toLowerCase();
    return r.RAG === 'Red' || riskText.includes('sos') || riskText.includes('khẩn');
  });

  const ceoDecisionRows = rows.filter(r => isDecision(r.CEODecision));
  const workstreamSummary = buildWorkstreamSummary(rows, ss);

  const dashboardData = {
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),

    summary: {
      totalActions: rows.length,
      doneActions: doneRows.length,
      inProgressActions: inProgressRows.length,
      overdueActions: overdueRows.length,
      blockedActions: blockedRows.length,
      completionRate: rows.length > 0 ? Math.round(doneRows.length / rows.length * 100) : 0,
      green: greenRows.length,
      amber: amberRows.length,
      red: redRows.length,
      riskCount: riskRows.length,
      sosCount: sosRows.length,
      ceoDecisionCount: ceoDecisionRows.length
    },

    metricSources: {
      totalActions: buildMetricSources(rows),
      doneActions: buildMetricSources(doneRows),
      inProgressActions: buildMetricSources(inProgressRows),
      overdueActions: buildMetricSources(overdueRows),
      blockedActions: buildMetricSources(blockedRows),
      green: buildMetricSources(greenRows),
      amber: buildMetricSources(amberRows),
      red: buildMetricSources(redRows),
      riskCount: buildMetricSources(riskRows),
      sosCount: buildMetricSources(sosRows),
      ceoDecisionCount: buildMetricSources(ceoDecisionRows)
    },

    workstreamSummary: workstreamSummary,
    workstreamSources: buildWorkstreamSources(rows, workstreamSummary),

    topOverdue: overdueRows.slice(0, 30),
    topRisks: uniqueRows(riskRows.concat(redRows)).slice(0, 30),
    decisionList: ceoDecisionRows.slice(0, 40),

    actions: rows.map(r => buildItem(r)),
    dataQuality: buildDataQuality_(rows),
    topCritical: buildTopCritical_(rows).slice(0, 10)
  };

  return JSON.parse(JSON.stringify(dashboardData));
}

function buildMetricSources(rows) {
  const groups = {};

  rows.forEach(function(r) {
    const key = r.Workstream || 'Chưa phân loại';

    if (!groups[key]) {
      groups[key] = {
        Workstream: key,
        WorkstreamDisplay: r.WorkstreamDisplay || key,
        Count: 0,
        SourceUrl: r.SourceUrl || '',
        Items: []
      };
    }

    groups[key].Count += 1;
    groups[key].Items.push(buildItem(r));
  });

  return Object.keys(groups).map(k => groups[k]);
}

function buildWorkstreamSources(rows, workstreamSummary) {
  const output = {};

  workstreamSummary.forEach(function(ws) {
    output[ws.Workstream] = {
      Workstream: ws.Workstream,
      WorkstreamDisplay: ws.WorkstreamDisplay,
      SourceUrl: ws.SourceUrl,
      Summary: ws,
      Items: rows.filter(r => r.Workstream === ws.Workstream).map(r => buildItem(r))
    };
  });

  return output;
}

function buildItem(r) {
  return {
    RowId: r.RowId || '',
    Workstream: r.Workstream || '',
    WorkstreamDisplay: r.WorkstreamDisplay || '',
    Unit: r.Unit || '',
    PIC: r.PIC || '',
    Action: r.Action || '',
    Owner: r.Owner || '',
    Deadline: r.Deadline || '',
    Status: r.Status || '',
    RAG: r.RAG || '',
    Risk: r.Risk || '',
    CEODecision: r.CEODecision || '',
    PMOComment: r.PMOComment || '',
    SourceUrl: r.SourceUrl || '',
    RowUrl: r.RowUrl || r.SourceUrl || '',
    PriorityScore: r.PriorityScore || 0,
    DataQuality: r.DataQuality || 0
  };
}

function buildWorkstreamSummary(rows, ss) {
  return DASHBOARD_SHEETS.map(function(config) {
    const sheet = ss.getSheetByName(config.name);
    const sourceUrl = sheet ? ss.getUrl() + '#gid=' + sheet.getSheetId() : '';

    const data = rows.filter(r => r.Workstream === config.label);
    const total = data.length;
    const done = data.filter(r => r.Status === 'Done').length;
    const overdue = data.filter(r => r.Status === 'Overdue' || r.IsOverdue).length;
    const blocked = data.filter(r => r.Status === 'Blocked').length;
    const risks = data.filter(r => hasText(r.Risk)).length;
    const decisions = data.filter(r => isDecision(r.CEODecision)).length;

    let rag = 'Green';
    if (data.some(r => r.RAG === 'Red') || overdue > 0 || blocked > 0) rag = 'Red';
    else if (data.some(r => r.RAG === 'Amber') || risks > 0 || decisions > 0) rag = 'Amber';
    else if (total === 0) rag = 'Amber';

    return {
      Workstream: config.label,
      WorkstreamDisplay: config.display,
      SourceUrl: sourceUrl,
      Total: total,
      Done: done,
      Completion: total > 0 ? Math.round(done / total * 100) : 0,
      Overdue: overdue,
      Blocked: blocked,
      Risks: risks,
      Decisions: decisions,
      RAG: rag
    };
  });
}

/* =====================================================
   PMO ADD-ONS
===================================================== */

function buildDataQuality_(rows) {
  const total = rows.length;

  const missingOwner = rows.filter(r => !hasText(r.Owner)).length;
  const missingPIC = rows.filter(r => !hasText(r.PIC)).length;
  const missingDeadline = rows.filter(r => !hasText(r.Deadline)).length;
  const missingAction = rows.filter(r => !hasText(r.Action)).length;
  const missingStatus = rows.filter(r => !hasText(r.Status)).length;
  const missingRAG = rows.filter(r => !hasText(r.RAG)).length;

  const score = total > 0
    ? Math.round(rows.reduce((sum, r) => sum + calcRowDataQuality_(r), 0) / total)
    : 0;

  return {
    score: score,
    total: total,
    missingOwner: missingOwner,
    missingPIC: missingPIC,
    missingDeadline: missingDeadline,
    missingAction: missingAction,
    missingStatus: missingStatus,
    missingRAG: missingRAG
  };
}

function buildTopCritical_(rows) {
  return rows.map(function(r) {
    const item = buildItem(r);
    item.PriorityScore = calcPriorityScore_(r);
    return item;
  }).sort(function(a, b) {
    return Number(b.PriorityScore || 0) - Number(a.PriorityScore || 0);
  });
}

function calcRowDataQuality_(r) {
  const fields = [
    r.Action,
    r.Owner,
    r.PIC,
    r.Deadline,
    r.Status,
    r.RAG
  ];

  const filled = fields.filter(x => hasText(x)).length;
  return Math.round(filled / fields.length * 100);
}

function calcPriorityScore_(r) {
  let score = 0;

  if (r.RAG === 'Red') score += 40;
  if (r.RAG === 'Amber') score += 15;
  if (r.Status === 'Blocked') score += 35;
  if (r.Status === 'Overdue') score += 35;
  if (r.IsOverdue || isPastDeadline(r.Deadline, r.Status)) score += 30;
  if (hasText(r.Risk)) score += 25;
  if (isDecision(r.CEODecision)) score += 25;
  if (!hasText(r.Owner)) score += 10;
  if (!hasText(r.Deadline)) score += 10;
  if (!hasText(r.PIC)) score += 8;
  if (!hasText(r.Action)) score += 8;

  return score;
}

/* =====================================================
   COMMENT / CHAT
===================================================== */

function saveDashboardComment(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateCommentSheet_(ss);

  payload = payload || {};

  const commentId = Utilities.getUuid();
  const parentId = String(payload.parentId || '').trim();
  const type = parentId ? 'Reply' : 'Comment';

  const nowDate = new Date();
  const nowText = Utilities.formatDate(nowDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

  const name = String(payload.name || 'Người dùng').trim();
  const role = String(payload.role || 'Khác').trim();
  const message = String(payload.message || '').trim();

  const imageName = String(payload.imageName || '').trim();
  const imageBase64 = String(payload.imageBase64 || '').trim();
  const imageMimeType = String(payload.imageMimeType || '').trim();
  const imagePreview = String(payload.imagePreview || '').trim();

  let imageUrl = '';

  if (imageBase64 && imageName && imageMimeType) {
    imageUrl = saveCommentImage_(imageBase64, imageName, imageMimeType);
  }

  sheet.appendRow([
    commentId,
    parentId,
    type,
    nowDate,
    name,
    role,
    message,
    imageUrl,
    imagePreview,
    0,
    '',
    'Mới',
    ''
  ]);

  SpreadsheetApp.flush();

  return {
    success: true,
    savedAt: nowText
  };
}

function getDashboardComments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateCommentSheet_(ss);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  return values
    .filter(function(row) {
      return row.join('').trim() !== '';
    })
    .map(function(row) {
      const timeValue = row[3];
      let timeText = '';

      if (Object.prototype.toString.call(timeValue) === '[object Date]' && !isNaN(timeValue)) {
        timeText = Utilities.formatDate(timeValue, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      } else {
        timeText = String(timeValue || '');
      }

      return {
        Id: String(row[0] || ''),
        ParentId: String(row[1] || ''),
        Type: String(row[2] || ''),
        Time: timeText,
        Name: String(row[4] || ''),
        Role: String(row[5] || ''),
        Message: String(row[6] || ''),
        ImageUrl: String(row[7] || ''),
        ImagePreview: String(row[8] || ''),
        Likes: Number(row[9] || 0),
        LikedBy: String(row[10] || ''),
        Status: String(row[11] || ''),
        PMONote: String(row[12] || '')
      };
    });
}

function likeDashboardComment(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateCommentSheet_(ss);

  payload = payload || {};

  const commentId = String(payload.commentId || '').trim();
  const likerName = String(payload.name || 'Người dùng').trim();

  if (!commentId) {
    return {
      success: false,
      comments: getDashboardComments()
    };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {
      success: false,
      comments: []
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === commentId) {
      const rowIndex = i + 2;
      const currentLikes = Number(sheet.getRange(rowIndex, 10).getValue() || 0);
      const likedByText = String(sheet.getRange(rowIndex, 11).getValue() || '');

      let likedBy = likedByText
        ? likedByText.split(',').map(x => x.trim()).filter(Boolean)
        : [];

      if (likedBy.indexOf(likerName) === -1) {
        likedBy.push(likerName);
        sheet.getRange(rowIndex, 10).setValue(currentLikes + 1);
        sheet.getRange(rowIndex, 11).setValue(likedBy.join(', '));
      }

      SpreadsheetApp.flush();

      return {
        success: true,
        comments: getDashboardComments()
      };
    }
  }

  return {
    success: false,
    comments: getDashboardComments()
  };
}

function deleteDashboardComment(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateCommentSheet_(ss);

  payload = payload || {};

  const commentId = String(payload.commentId || '').trim();

  if (!commentId) {
    return {
      success: false,
      comments: getDashboardComments()
    };
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return {
      success: false,
      comments: []
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 13).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const rowCommentId = String(values[i][0] || '');
    const rowParentId = String(values[i][1] || '');

    if (rowCommentId === commentId || rowParentId === commentId) {
      sheet.deleteRow(i + 2);
    }
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    comments: getDashboardComments()
  };
}

function getOrCreateCommentSheet_(ss) {
  let sheet = ss.getSheetByName(COMMENT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(COMMENT_SHEET_NAME);
  }

  const headers = [
    'Comment ID',
    'Parent ID',
    'Loại',
    'Thời gian',
    'Người gửi',
    'Vai trò',
    'Lời nhắn / Comment',
    'Link ảnh đính kèm',
    'Preview ảnh',
    'Like',
    'Người đã like',
    'Trạng thái',
    'Ghi chú PMO'
  ];

  sheet.getRange(1, 1, 1, 13).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 13)
    .setFontWeight('bold')
    .setBackground('#064e3b')
    .setFontColor('#ffffff');

  return sheet;
}

function saveCommentImage_(base64Data, fileName, mimeType) {
  const folders = DriveApp.getFoldersByName(COMMENT_FOLDER_NAME);

  let folder;

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(COMMENT_FOLDER_NAME);
  }

  const cleanBase64 = String(base64Data).split(',').pop();
  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

/* =====================================================
   HELPERS
===================================================== */

function uniqueRows(rows) {
  const seen = {};
  const output = [];

  rows.forEach(function(r) {
    const key = r.RowId || r.Workstream + '_' + r.Action;
    if (seen[key]) return;

    seen[key] = true;
    output.push(r);
  });

  return output;
}

function getVal(obj, keys) {
  const headers = Object.keys(obj || {});

  for (let i = 0; i < keys.length; i++) {
    const key = String(keys[i] || '').toLowerCase();

    const found = headers.find(function(h) {
      return String(h || '').toLowerCase().includes(key);
    });

    if (
      found &&
      obj[found] !== '' &&
      obj[found] !== null &&
      obj[found] !== undefined
    ) {
      return obj[found];
    }
  }

  return '';
}

function normalizeStatus(value) {
  const s = String(value || '').toLowerCase();

  if (s.includes('hoàn thành') || s.includes('done') || s.includes('xong')) return 'Done';
  if (s.includes('quá hạn') || s.includes('overdue') || s.includes('trễ')) return 'Overdue';
  if (s.includes('vướng') || s.includes('blocked') || s.includes('chờ') || s.includes('tắc')) return 'Blocked';
  if (s.includes('đang') || s.includes('progress') || s.includes('triển khai')) return 'In Progress';

  return 'Not Started';
}

function normalizeRag(value) {
  const s = String(value || '').toLowerCase();

  if (s.includes('xanh') || s.includes('green') || s.includes('🟢')) return 'Green';
  if (s.includes('đỏ') || s.includes('red') || s.includes('🔴')) return 'Red';
  if (s.includes('vàng') || s.includes('amber') || s.includes('yellow') || s.includes('🟡')) return 'Amber';

  return 'Amber';
}

function isPastDeadline(deadlineText, status) {
  if (!deadlineText || status === 'Done') return false;

  const d = parseDate(deadlineText);

  if (!d) return false;

  const today = new Date();

  today.setHours(0, 0, 0, 0);
  d.setHours(0, 0, 0, 0);

  return d < today;
}

function parseDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }

  const s = String(value || '').trim();
  const ddmmyyyy = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);

  if (ddmmyyyy) {
    return new Date(Number(ddmmyyyy[3]), Number(ddmmyyyy[2]) - 1, Number(ddmmyyyy[1]));
  }

  const d = new Date(s);

  if (!isNaN(d)) return d;

  return null;
}

function formatDate(value) {
  if (!value) return '';

  const d = parseDate(value);

  if (d) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  return String(value);
}

function hasText(value) {
  return value !== null && value !== undefined && String(value).trim() !== '';
}

function isDecision(value) {
  const s = String(value || '').trim().toLowerCase();

  if (!s) return false;

  const noValues = [
    'không',
    'khong',
    'không có',
    'khong co',
    'no',
    'none',
    'n/a',
    'na',
    'chưa có',
    'chua co'
  ];

  return noValues.indexOf(s) === -1;
}

/* =====================================================
   TEST
===================================================== */

function testGetDashboardData() {
  const data = getDashboardData();
  Logger.log(JSON.stringify(data.summary, null, 2));
}

function testReadDashboardComments() {
  const comments = getDashboardComments();
  Logger.log('Số comment đọc được: ' + comments.length);
  Logger.log(JSON.stringify(comments, null, 2));
}

function testSaveCommentManual() {
  const result = saveDashboardComment({
    parentId: '',
    name: 'Mr Sơn',
    role: 'PMO',
    message: 'Test lưu comment thủ công từ Apps Script',
    imageName: '',
    imageBase64: '',
    imageMimeType: '',
    imagePreview: ''
  });

  Logger.log(JSON.stringify(result, null, 2));
}

/* =====================================================
   SAFE SHEET READER
   Đọc sheet với header ở hàng 1, data từ hàng 2.
   Trả về { exists, headers, rows, rowCount }.
   Không throw nếu sheet chưa tạo.
===================================================== */

function safeReadSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { exists: false, headers: [], rows: [], rowCount: 0 };
  }
  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { exists: true, headers: values && values[0] ? values[0] : [], rows: [], rowCount: 0 };
  }
  var headers = values[0].map(function(h) { return String(h || '').trim(); });
  var rows = values.slice(1)
    .filter(function(row) {
      return row.some(function(cell) { return String(cell || '').trim() !== ''; });
    })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = (row[i] !== undefined && row[i] !== null) ? row[i] : ''; });
      return obj;
    });
  return { exists: true, headers: headers, rows: rows, rowCount: rows.length };
}

/* =====================================================
   CEO PROJECTS
   Sheet: CEO_PROJECTS (header hàng 1)
   → Khi Vướng mắc ≠ "" hoặc RAG = Red → tự vào Risk/SOS
   → Khi CEO cần chốt ≠ "" → tự vào CEO Decision Board
===================================================== */

function getCEOProjectsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = safeReadSheet(ss, 'CEO_PROJECTS');

  var records = data.rows.map(function(r, i) {
    var rag = normalizeRag(r['RAG'] || '');
    var status = normalizeStatus(r['Trạng thái'] || r['Status'] || '');
    return {
      RowId:       String(r['Project ID'] || ('PRJ-' + String(i + 1).padStart(3, '0'))),
      ProjectName: String(r['Tên dự án']   || r['Project Name'] || ''),
      PICLead:     String(r['PIC Lead']    || r['PIC'] || ''),
      Owner:       String(r['Owner']       || r['CEO phụ trách'] || ''),
      Workstream:  String(r['Workstream']  || r['Lĩnh vực'] || ''),
      StartDate:   formatDate(r['Ngày bắt đầu']),
      Deadline:    formatDate(r['Deadline']),
      Status:      status,
      RAG:         rag,
      Progress:    String(r['% Hoàn thành'] || r['Progress'] || ''),
      Milestone:   String(r['Milestone gần nhất'] || r['Milestone'] || ''),
      NextMile:    String(r['Milestone tiếp theo'] || ''),
      Issues:      String(r['Vướng mắc']   || r['Risk/SOS'] || ''),
      CEODecision: String(r['CEO cần chốt'] || r['CEO Decision'] || ''),
      Note:        String(r['Ghi chú'] || ''),
      RowUrl:      String(r['Link'] || r['URL'] || '')
    };
  });

  return { exists: data.exists, records: records, total: records.length };
}

/* =====================================================
   TO ACTIONS
   Sheet: TO_ACTIONS (header hàng 1)
   → Khi CEO cần biết ≠ "" → tự vào CEO Decision Board
   → Khi RAG = Red hoặc Mức độ ưu tiên = Cao → tự vào Risk/SOS
===================================================== */

function getTOActionsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = safeReadSheet(ss, 'TO_ACTIONS');

  var records = data.rows.map(function(r, i) {
    return {
      RowId:       String(r['Action ID']          || ('TOA-' + String(i + 1).padStart(4, '0'))),
      ActionID:    String(r['Action ID']          || ('TOA-' + String(i + 1).padStart(4, '0'))),
      Date:        formatDate(r['Ngày phát sinh'] || r['Ngày'] || ''),
      MeetingDate: formatDate(r['Ngày phát sinh'] || ''),
      Meeting:     String(r['Cuộc họp']          || r['Bối cảnh'] || ''),
      MeetingID:   String(r['Meeting ID']         || ''),
      Action:      String(r['Nội dung action']    || r['Action'] || ''),
      Content:     String(r['Nội dung action']    || r['Action'] || ''),
      PIC:         String(r['PIC thực hiện']      || r['PIC'] || ''),
      Owner:       String(r['Owner theo dõi']     || r['Owner'] || ''),
      Deadline:    formatDate(r['Deadline']        || ''),
      Status:      normalizeStatus(r['Trạng thái'] || r['Status'] || ''),
      Priority:    String(r['Mức độ ưu tiên']     || r['Priority'] || ''),
      RAG:         normalizeRag(r['RAG']           || ''),
      Result:      String(r['Kết quả']            || r['Output'] || ''),
      CEONote:     String(r['CEO cần biết']        || ''),
      Note:        String(r['Ghi chú']            || ''),
      RowUrl:      String(r['Link biên bản']       || r['Link'] || '')
    };
  });

  return { exists: data.exists, records: records, total: records.length };
}

/* =====================================================
   MEETING MINUTES
   Sheet: MEETING_MINUTES (header hàng 1)
   → Khi Risk/SOS phát sinh ≠ "" → tự vào Risk/SOS Center
   → Khi Chỉ đạo CEO hoặc CEO cần quyết tiếp ≠ "" → CEO Decision
===================================================== */

function getMeetingMinutesData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = safeReadSheet(ss, 'MEETING_MINUTES');

  var records = data.rows.map(function(r, i) {
    return {
      MeetingID:    String(r['Meeting ID']            || ('MTG-' + String(i + 1).padStart(3, '0'))),
      Date:         formatDate(r['Ngày họp']          || ''),
      Topic:        String(r['Chủ đề họp']            || r['Chủ đề'] || ''),
      Chair:        String(r['Người chủ trì']         || ''),
      Attendees:    String(r['Thành phần tham dự']    || r['Thành phần'] || ''),
      Scope:        String(r['Phạm vi']               || ''),
      Summary:      String(r['Tóm tắt chính']         || r['Tóm tắt'] || ''),
      CEODirective: String(r['Chỉ đạo CEO']           || ''),
      Decisions:    String(r['Quyết định đã chốt']    || ''),
      Actions:      String(r['Action phát sinh']       || ''),
      RiskSOS:      String(r['Risk/SOS phát sinh']     || r['Risk SOS phát sinh'] || ''),
      CEODecision:  String(r['CEO cần quyết tiếp']     || ''),
      FileMinutes:  String(r['File biên bản']          || ''),
      FileAudio:    String(r['File ghi âm']            || r['File transcript'] || ''),
      RecordedBy:   String(r['Người lập biên bản']     || ''),
      Status:       String(r['Trạng thái']             || 'Nháp'),
      SecretNote:   String(r['Ghi chú bảo mật']        || '')
    };
  });

  return {
    exists:  data.exists,
    records: records,
    total:   records.length,
    pending: records.filter(function(r) { return String(r.Status || '').includes('Chờ'); }).length
  };
}

/* =====================================================
   RISK / SOS CENTER — Tổng hợp từ TẤT CẢ nguồn
   Nguồn 1: 6 sheets PMO — cột Risk/SOS ≠ "" hoặc RAG = Red
   Nguồn 2: CEO_PROJECTS — Vướng mắc ≠ "" hoặc RAG = Red
   Nguồn 3: MEETING_MINUTES — Risk/SOS phát sinh ≠ ""
   Nguồn 4: TO_ACTIONS — RAG = Red hoặc Mức độ ưu tiên = Cao
===================================================== */

function getRiskSOSData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allRisks = [];
  var seen = {};

  function addRisk(item) {
    var key = (item.Source || '') + '||' + (item.RowId || '') + '||' + String(item.Risk || '').slice(0, 40);
    if (seen[key]) return;
    seen[key] = true;
    allRisks.push(item);
  }

  // --- Nguồn 1: 6 sheets PMO (header hàng 3, data từ hàng 4) ---
  DASHBOARD_SHEETS.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 4) return;
    var headers = values[2].map(function(h) { return String(h || '').trim(); });
    var sheetId = sheet.getSheetId();
    values.slice(3).forEach(function(row, idx) {
      var obj = {};
      headers.forEach(function(h, i) { if (h) obj[h] = row[i]; });
      var rowId = getVal(obj, ['Mã dòng (ID)', 'Mã dòng', 'ID']);
      if (!hasText(rowId)) return;
      var risk = String(getVal(obj, ['Risk / SOS', 'Risk', 'SOS', 'Rủi ro', 'Blocker', 'Vướng mắc']) || '');
      var rag  = normalizeRag(getVal(obj, ['RAG', 'Đèn RAG', 'Đèn báo']));
      if (!hasText(risk) && rag !== 'Red') return;
      addRisk({
        Source:           config.display,
        Workstream:       config.label,
        WorkstreamDisplay:config.display,
        RowId:            String(rowId),
        PIC:              String(getVal(obj, ['PIC Lead', 'PIC', 'Đầu mối']) || ''),
        Owner:            String(getVal(obj, ['Owner', 'Người phụ trách']) || ''),
        Risk:             risk || 'RAG Red',
        Action:           String(getVal(obj, ['Kế hoạch hành động', 'Action Plan', 'Action', 'Hành động']) || ''),
        Deadline:         formatDate(getVal(obj, ['Deadline', 'Hạn hoàn thành'])),
        RAG:              rag,
        Status:           normalizeStatus(getVal(obj, ['Status', 'Trạng thái', 'Tình trạng'])),
        CEODecision:      String(getVal(obj, ['CEO Decision Needed', 'CEO cần chốt', 'Cần CEO chốt']) || ''),
        RowUrl:           ss.getUrl() + '#gid=' + sheetId + '&range=A' + (idx + 4) + ':Z' + (idx + 4)
      });
    });
  });

  // --- Nguồn 2: CEO_PROJECTS ---
  safeReadSheet(ss, 'CEO_PROJECTS').rows.forEach(function(r) {
    var issues = String(r['Vướng mắc'] || r['Risk/SOS'] || '');
    var rag    = normalizeRag(r['RAG'] || '');
    if (!hasText(issues) && rag !== 'Red') return;
    addRisk({
      Source:           'CEO Projects',
      Workstream:       'CEO Projects',
      WorkstreamDisplay:'CEO Projects',
      RowId:            String(r['Project ID'] || ''),
      PIC:              String(r['PIC Lead'] || r['PIC'] || ''),
      Owner:            String(r['Owner'] || r['CEO phụ trách'] || ''),
      Risk:             issues || 'RAG Red',
      Action:           String(r['Tên dự án'] || r['Project Name'] || ''),
      Deadline:         formatDate(r['Deadline'] || ''),
      RAG:              rag,
      Status:           normalizeStatus(r['Trạng thái'] || r['Status'] || ''),
      CEODecision:      String(r['CEO cần chốt'] || ''),
      RowUrl:           String(r['Link'] || '')
    });
  });

  // --- Nguồn 3: MEETING_MINUTES ---
  safeReadSheet(ss, 'MEETING_MINUTES').rows.forEach(function(r) {
    var riskText = String(r['Risk/SOS phát sinh'] || r['Risk SOS phát sinh'] || '');
    if (!hasText(riskText)) return;
    addRisk({
      Source:           'Biên bản họp',
      Workstream:       'Meeting',
      WorkstreamDisplay:'BB: ' + String(r['Chủ đề họp'] || '').slice(0, 30),
      RowId:            String(r['Meeting ID'] || ''),
      PIC:              String(r['Người chủ trì'] || ''),
      Owner:            String(r['Người lập biên bản'] || ''),
      Risk:             riskText,
      Action:           String(r['Chủ đề họp'] || ''),
      Deadline:         formatDate(r['Ngày họp'] || ''),
      RAG:              'Amber',
      Status:           String(r['Trạng thái'] || 'Nháp'),
      CEODecision:      String(r['CEO cần quyết tiếp'] || ''),
      RowUrl:           String(r['File biên bản'] || '')
    });
  });

  // --- Nguồn 4: TO_ACTIONS (RAG=Red hoặc Ưu tiên Cao) ---
  safeReadSheet(ss, 'TO_ACTIONS').rows.forEach(function(r) {
    var rag      = normalizeRag(r['RAG'] || '');
    var priority = String(r['Mức độ ưu tiên'] || r['Priority'] || '').toLowerCase();
    if (rag !== 'Red' && !priority.includes('cao')) return;
    addRisk({
      Source:           'TO Actions',
      Workstream:       'TO',
      WorkstreamDisplay:'TO: ' + String(r['Cuộc họp'] || '').slice(0, 25),
      RowId:            String(r['Action ID'] || ''),
      PIC:              String(r['PIC thực hiện'] || r['PIC'] || ''),
      Owner:            String(r['Owner theo dõi'] || r['Owner'] || ''),
      Risk:             String(r['Nội dung action'] || ''),
      Action:           String(r['Nội dung action'] || ''),
      Deadline:         formatDate(r['Deadline'] || ''),
      RAG:              rag,
      Status:           normalizeStatus(r['Trạng thái'] || r['Status'] || ''),
      CEODecision:      String(r['CEO cần biết'] || ''),
      RowUrl:           String(r['Link biên bản'] || r['Link'] || '')
    });
  });

  // Sắp xếp: Red → Amber → Green
  var ragOrder = { Red: 0, Amber: 1, Green: 2 };
  allRisks.sort(function(a, b) {
    return (ragOrder[a.RAG] || 1) - (ragOrder[b.RAG] || 1);
  });

  return {
    exists:     true,
    topRisks:   allRisks,
    total:      allRisks.length,
    redCount:   allRisks.filter(function(r) { return r.RAG === 'Red'; }).length,
    amberCount: allRisks.filter(function(r) { return r.RAG === 'Amber'; }).length
  };
}

/* =====================================================
   CEO DECISION BOARD — Tổng hợp từ TẤT CẢ nguồn
   Nguồn 1: 6 sheets PMO — cột "CEO cần chốt" ≠ ""
   Nguồn 2: CEO_PROJECTS — cột "CEO cần chốt" ≠ ""
   Nguồn 3: MEETING_MINUTES — "Chỉ đạo CEO" hoặc "CEO cần quyết tiếp" ≠ ""
   Nguồn 4: TO_ACTIONS — cột "CEO cần biết" ≠ ""
===================================================== */

function getCEODecisionsData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allDecisions = [];
  var seen = {};

  function addDecision(item) {
    var key = (item.Source || '') + '||' + (item.RowId || '') + '||' + String(item.CEODecision || '').slice(0, 40);
    if (seen[key]) return;
    seen[key] = true;
    allDecisions.push(item);
  }

  // --- Nguồn 1: 6 sheets PMO ---
  DASHBOARD_SHEETS.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 4) return;
    var headers = values[2].map(function(h) { return String(h || '').trim(); });
    var sheetId = sheet.getSheetId();
    values.slice(3).forEach(function(row, idx) {
      var obj = {};
      headers.forEach(function(h, i) { if (h) obj[h] = row[i]; });
      var rowId  = getVal(obj, ['Mã dòng (ID)', 'Mã dòng', 'ID']);
      if (!hasText(rowId)) return;
      var ceoCol = String(getVal(obj, ['CEO Decision Needed', 'CEO cần chốt', 'CEO quyết định', 'Cần CEO chốt']) || '');
      if (!isDecision(ceoCol)) return;
      addDecision({
        Source:           config.display,
        Workstream:       config.label,
        WorkstreamDisplay:config.display,
        RowId:            String(rowId),
        PIC:              String(getVal(obj, ['PIC Lead', 'PIC', 'Đầu mối']) || ''),
        Owner:            String(getVal(obj, ['Owner', 'Người phụ trách']) || ''),
        CEODecision:      ceoCol,
        Action:           String(getVal(obj, ['Kế hoạch hành động', 'Action Plan', 'Action', 'Hành động']) || ''),
        Deadline:         formatDate(getVal(obj, ['Deadline', 'Hạn hoàn thành'])),
        RAG:              normalizeRag(getVal(obj, ['RAG', 'Đèn RAG'])),
        Risk:             String(getVal(obj, ['Risk / SOS', 'Risk', 'SOS', 'Vướng mắc']) || ''),
        RowUrl:           ss.getUrl() + '#gid=' + sheetId + '&range=A' + (idx + 4) + ':Z' + (idx + 4)
      });
    });
  });

  // --- Nguồn 2: CEO_PROJECTS ---
  safeReadSheet(ss, 'CEO_PROJECTS').rows.forEach(function(r) {
    var ceoCol = String(r['CEO cần chốt'] || r['CEO Decision'] || '');
    if (!isDecision(ceoCol)) return;
    addDecision({
      Source:           'CEO Projects',
      Workstream:       'CEO Projects',
      WorkstreamDisplay:'CEO Projects',
      RowId:            String(r['Project ID'] || ''),
      PIC:              String(r['PIC Lead'] || r['PIC'] || ''),
      Owner:            String(r['Owner'] || ''),
      CEODecision:      ceoCol,
      Action:           String(r['Tên dự án'] || r['Project Name'] || ''),
      Deadline:         formatDate(r['Deadline'] || ''),
      RAG:              normalizeRag(r['RAG'] || ''),
      Risk:             String(r['Vướng mắc'] || ''),
      RowUrl:           String(r['Link'] || '')
    });
  });

  // --- Nguồn 3: MEETING_MINUTES ---
  safeReadSheet(ss, 'MEETING_MINUTES').rows.forEach(function(r) {
    var directive = String(r['Chỉ đạo CEO'] || '');
    var nextDecision = String(r['CEO cần quyết tiếp'] || '');
    if (!hasText(directive) && !hasText(nextDecision)) return;
    addDecision({
      Source:           'Biên bản họp',
      Workstream:       'Meeting',
      WorkstreamDisplay:'BB: ' + String(r['Chủ đề họp'] || '').slice(0, 30),
      RowId:            String(r['Meeting ID'] || ''),
      PIC:              String(r['Người chủ trì'] || ''),
      Owner:            String(r['Người lập biên bản'] || ''),
      CEODecision:      nextDecision || directive,
      Action:           directive || String(r['Chủ đề họp'] || ''),
      Deadline:         formatDate(r['Ngày họp'] || ''),
      RAG:              'Amber',
      Risk:             String(r['Risk/SOS phát sinh'] || ''),
      RowUrl:           String(r['File biên bản'] || '')
    });
  });

  // --- Nguồn 4: TO_ACTIONS ---
  safeReadSheet(ss, 'TO_ACTIONS').rows.forEach(function(r) {
    var ceoCol = String(r['CEO cần biết'] || '');
    if (!hasText(ceoCol)) return;
    addDecision({
      Source:           'TO Actions',
      Workstream:       'TO',
      WorkstreamDisplay:'TO: ' + String(r['Cuộc họp'] || '').slice(0, 25),
      RowId:            String(r['Action ID'] || ''),
      PIC:              String(r['PIC thực hiện'] || r['PIC'] || ''),
      Owner:            String(r['Owner theo dõi'] || r['Owner'] || ''),
      CEODecision:      ceoCol,
      Action:           String(r['Nội dung action'] || ''),
      Deadline:         formatDate(r['Deadline'] || ''),
      RAG:              normalizeRag(r['RAG'] || ''),
      Risk:             '',
      RowUrl:           String(r['Link biên bản'] || r['Link'] || '')
    });
  });

  // Sắp xếp: Red → Amber → Green
  var ragOrder = { Red: 0, Amber: 1, Green: 2 };
  allDecisions.sort(function(a, b) {
    return (ragOrder[a.RAG] || 1) - (ragOrder[b.RAG] || 1);
  });

  return {
    exists:       true,
    decisionList: allDecisions,
    total:        allDecisions.length,
    redCount:     allDecisions.filter(function(r) { return r.RAG === 'Red'; }).length
  };
}

/* =====================================================
   DATA QUALITY — Kiểm tra từng sheet
   Quy tắc DQ:
   PMO: bắt buộc RowId, Action, PIC, Deadline, Status, RAG
   CEO_PROJECTS: bắt buộc Project ID, Tên dự án, PIC Lead, Deadline, Trạng thái, RAG
   TO_ACTIONS: bắt buộc Action ID, Nội dung action, PIC thực hiện, Deadline, Trạng thái
   MEETING_MINUTES: bắt buộc Meeting ID, Ngày họp, Chủ đề họp, Người chủ trì, Trạng thái
===================================================== */

function getDataQualityFullData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var modules = [];

  // Tính DQ score cho 1 mảng rows với danh sách field bắt buộc
  function calcDQ(rows, requiredFields, label, sheetName) {
    if (!rows.length) return { module: label, sheetName: sheetName, total: 0, score: 100, missing: {} };
    var missingCount = {};
    requiredFields.forEach(function(f) { missingCount[f] = 0; });
    rows.forEach(function(r) {
      requiredFields.forEach(function(f) {
        if (!hasText(r[f])) missingCount[f]++;
      });
    });
    var totalChecks  = rows.length * requiredFields.length;
    var totalMissing = requiredFields.reduce(function(s, f) { return s + missingCount[f]; }, 0);
    var score = totalChecks > 0 ? Math.round((1 - totalMissing / totalChecks) * 100) : 100;
    return { module: label, sheetName: sheetName, total: rows.length, score: score, missing: missingCount };
  }

  // ---- PMO 6 sheets (đọc theo cấu trúc đặc biệt: header hàng 3, data từ hàng 4) ----
  DASHBOARD_SHEETS.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      modules.push({ module: config.display, sheetName: config.name, sheetExists: false, total: 0, score: 0, missing: {} });
      return;
    }
    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 4) {
      modules.push({ module: config.display, sheetName: config.name, sheetExists: true, total: 0, score: 100, missing: {} });
      return;
    }
    var headers = values[2].map(function(h) { return String(h || '').trim(); });
    var rows = values.slice(3)
      .filter(function(row) {
        return headers.some(function(h, i) {
          return (h.includes('Mã dòng') || h === 'ID') && hasText(row[i]);
        });
      })
      .map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { if (h) obj[h] = row[i]; });
        // Chuẩn hóa về key cố định để kiểm tra
        return {
          'Action':   getVal(obj, ['Kế hoạch hành động', 'Action Plan', 'Action', 'Hành động', 'Nhiệm vụ']),
          'PIC':      getVal(obj, ['PIC Lead', 'PIC', 'Đầu mối']),
          'Owner':    getVal(obj, ['Owner', 'Người phụ trách']),
          'Deadline': getVal(obj, ['Deadline', 'Hạn hoàn thành']),
          'Status':   getVal(obj, ['Status', 'Trạng thái', 'Tình trạng']),
          'RAG':      getVal(obj, ['RAG', 'Đèn RAG', 'Đèn báo'])
        };
      });
    var dq = calcDQ(rows, ['Action', 'PIC', 'Owner', 'Deadline', 'Status', 'RAG'], config.display, config.name);
    dq.sheetExists = true;
    modules.push(dq);
  });

  // ---- CEO_PROJECTS ----
  var cptData = safeReadSheet(ss, 'CEO_PROJECTS');
  var cptDQ = calcDQ(cptData.rows, ['Tên dự án', 'PIC Lead', 'Deadline', 'Trạng thái', 'RAG'], 'CEO Projects', 'CEO_PROJECTS');
  cptDQ.sheetExists = cptData.exists;
  modules.push(cptDQ);

  // ---- TO_ACTIONS ----
  var toaData = safeReadSheet(ss, 'TO_ACTIONS');
  var toaDQ = calcDQ(toaData.rows, ['Nội dung action', 'PIC thực hiện', 'Deadline', 'Trạng thái'], 'TO Actions', 'TO_ACTIONS');
  toaDQ.sheetExists = toaData.exists;
  modules.push(toaDQ);

  // ---- MEETING_MINUTES ----
  var mmData = safeReadSheet(ss, 'MEETING_MINUTES');
  var mmDQ = calcDQ(mmData.rows, ['Ngày họp', 'Chủ đề họp', 'Người chủ trì', 'Trạng thái'], 'Biên bản họp', 'MEETING_MINUTES');
  mmDQ.sheetExists = mmData.exists;
  modules.push(mmDQ);

  // ---- Tổng hợp ----
  var existingModules = modules.filter(function(m) { return m.sheetExists && m.total > 0; });
  var totalRows  = existingModules.reduce(function(s, m) { return s + m.total; }, 0);
  var avgScore   = existingModules.length > 0
    ? Math.round(existingModules.reduce(function(s, m) { return s + m.score; }, 0) / existingModules.length)
    : 0;

  return {
    exists:   true,
    overall:  { score: avgScore, total: totalRows, moduleCount: existingModules.length },
    modules:  modules
  };
}

/* =====================================================
   TEST — New endpoints
===================================================== */

function testGetCEOProjects() {
  Logger.log(JSON.stringify(getCEOProjectsData(), null, 2));
}

function testGetTOActions() {
  Logger.log(JSON.stringify(getTOActionsData(), null, 2));
}

function testGetMeetingMinutes() {
  Logger.log(JSON.stringify(getMeetingMinutesData(), null, 2));
}

function testGetRiskSOS() {
  var data = getRiskSOSData();
  Logger.log('Tổng Risk/SOS: ' + data.total + ' (Red: ' + data.redCount + ', Amber: ' + data.amberCount + ')');
}

function testGetCEODecisions() {
  var data = getCEODecisionsData();
  Logger.log('Tổng CEO Decisions: ' + data.total + ' (Red: ' + data.redCount + ')');
}

function testGetDataQualityFull() {
  var data = getDataQualityFullData();
  Logger.log('Điểm DQ tổng thể: ' + data.overall.score + '% (' + data.overall.total + ' dòng)');
  data.modules.forEach(function(m) {
    Logger.log(m.module + ': ' + m.score + '% (' + m.total + ' dòng) - Sheet exists: ' + m.sheetExists);
  });
}