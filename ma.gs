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
  if (action === 'getTOActions') {
    try {
      return jsonOut(getTOActionsData());
    } catch(eTOA) {
      Logger.log('[getTOActions] ERROR: ' + eTOA.message);
      return jsonOut({ exists: false, records: [], total: 0, _error: eTOA.message });
    }
  }
  if (action === 'getMeetingMinutes')   return jsonOut(getMeetingMinutesData());
  if (action === 'getRiskSOS')          return jsonOut(getRiskSOSData());
  if (action === 'getCEODecisions')     return jsonOut(getCEODecisionsData());
  if (action === 'getDataQualityFull')  return jsonOut(getDataQualityFullData());
  if (action === 'getCommandCenterData') return jsonOut(getCommandCenterData());
  if (action === 'getPicPmoData')        return jsonOut(getPicPmoData());

  if (action === 'ping') return jsonOut({ ok: true, ts: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss') });

  if (action === 'getToActionsRaw') {
    // DEBUG: đọc thẳng TO_ACTIONS, trả raw data
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var sh  = ss2.getSheetByName('TO_ACTIONS');
    if (!sh) return jsonOut({ ok: false, reason: 'Sheet TO_ACTIONS không tìm thấy', allSheets: ss2.getSheets().map(function(s){ return s.getName(); }) });
    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) return jsonOut({ ok: false, reason: 'Sheet trống (lastRow=' + lastRow + ')' });
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var sample  = sh.getRange(2, 1, Math.min(3, lastRow - 1), lastCol).getValues();
    return jsonOut({ ok: true, rowCount: lastRow - 1, headers: headers, sample: sample });
  }

  if (action === 'getToActionCenterData') {
    try {
      return jsonOut(getToActionCenterData());
    } catch(e) {
      Logger.log('[getToActionCenterData] ERROR: ' + e.message);
      // Fallback: đọc đơn giản qua getTOActionsData
      try {
        var toa = getTOActionsData();
        var recs = toa.records || [];
        return jsonOut({
          actions: recs, issues: [], factories: [],
          evidenceFiles: [], byArea: [],
          urgent24h: [], urgent72h: [], inPlan: recs,
          noOwnerList: [], ceoList: [],
          _error: e.message,
          issuesSummary: { total: 0, byRAG: { Green: 0, Amber: 0, Red: 0 }, open: 0 },
          actionsSummary: { total: recs.length, done: 0, inProgress: 0, overdue: 0,
            noOwner: 0, ceoNeeded: 0, urgent24hCount: 0, urgent72hCount: 0, inPlanCount: recs.length }
        });
      } catch(e2) {
        return jsonOut({ actions: [], issues: [], factories: [], _error: e.message + ' | ' + e2.message });
      }
    }
  }

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
   PMO NEW MODEL — 16 PIC SHEETS + PMO_ALL_ACTIONS
===================================================== */

const PIC_SHEETS_CONFIG = [
  { name: 'PIC01_KinhDoanh',      code: 'PIC01', label: '01 Kinh Doanh' },
  { name: 'PIC02_SalesTech',      code: 'PIC02', label: '02 Sales Tech' },
  { name: 'PIC03_SCM',            code: 'PIC03', label: '03 SCM' },
  { name: 'PIC04_MuaHang',        code: 'PIC04', label: '04 Mua Hàng' },
  { name: 'PIC05_CongNghe',       code: 'PIC05', label: '05 Công Nghệ' },
  { name: 'PIC06_CoDien_QLTB',    code: 'PIC06', label: '06 Cơ Điện/QLTB' },
  { name: 'PIC07_IT_ERP',         code: 'PIC07', label: '07 IT/ERP' },
  { name: 'PIC08_QA',             code: 'PIC08', label: '08 QA' },
  { name: 'PIC09_PC',             code: 'PIC09', label: '09 P&C' },
  { name: 'PIC10_GiamSatTuanThu', code: 'PIC10', label: '10 Giám Sát Tuân Thủ' },
  { name: 'PIC11_KiemSoatNoiBo',  code: 'PIC11', label: '11 Kiểm Soát Nội Bộ' },
  { name: 'PIC12_HCNS',           code: 'PIC12', label: '12 HCNS' },
  { name: 'PIC13_TaiChinhKeToan', code: 'PIC13', label: '13 Tài Chính/Kế Toán' },
  { name: 'PIC14_GS1',            code: 'PIC14', label: '14 GS1' },
  { name: 'PIC15_GS5',            code: 'PIC15', label: '15 GS5' },
  { name: 'PIC16_GS6',            code: 'PIC16', label: '16 GS6' }
];

const PMO_STANDARD_HEADERS = [
  'Action ID', 'PIC Code', 'PIC Name', 'Bộ phận/Đơn vị', 'Nhà máy/Khối',
  'Workstream', 'Nhóm việc', 'Action', 'Output kỳ vọng', 'Owner', 'Phối hợp',
  'Start Date', 'Deadline', 'Status', 'RAG', '% Tiến độ', 'Update mới nhất',
  'Risk/SOS', 'CEO Decision', 'PMO Comment', 'Evidence Link',
  'Source File', 'Last Update', 'Update By'
];

const PMO_COLUMN_ALIASES = {
  'Action ID':        ['Action ID', 'Mã dòng (ID)', 'Mã dòng', 'ID', 'STT', 'No.', 'No'],
  'PIC Code':         ['PIC Code', 'Mã PIC'],
  'PIC Name':         ['PIC Name', 'Tên PIC'],
  'Bộ phận/Đơn vị':  ['Bộ phận/Đơn vị', 'Bộ phận', 'Đơn vị', 'Đơn vị / Khối', 'Khối/Bộ phận'],
  'Nhà máy/Khối':    ['Nhà máy/Khối', 'Nhà máy', 'Factory', 'Site', 'Khối'],
  'Workstream':       ['Workstream', 'Trục', 'Lĩnh vực', 'Trục chiến lược'],
  'Nhóm việc':        ['Nhóm việc', 'Nhóm', 'Phân nhóm', 'Group', 'Category', 'Mã'],
  'Action':           ['Action', 'Kế hoạch hành động', 'Action Plan', 'Việc cần làm', 'Nhiệm vụ', 'Nội dung việc', 'Task',
                       'Hành động', 'Cam kết hành động', 'Hành động / Công việc', 'Hành động/Công việc', 'Công việc'],
  'Output kỳ vọng':   ['Output kỳ vọng', 'Expected Output', 'Output', 'Kết quả kỳ vọng', 'Deliverable', 'Kết quả đầu ra',
                       'Kết quả / Output nếu có', 'Kết quả / Output', 'Kết quả/Output', 'Kết quả'],
  'Owner':            ['Owner', 'PIC', 'PIC Lead', 'Đầu mối', 'Người phụ trách', 'Người chịu trách nhiệm', 'Người thực hiện',
                       'Owner (Người phụ trách)'],
  'Phối hợp':         ['Phối hợp', 'Coordinator', 'Người phối hợp', 'Support', 'Hỗ trợ'],
  'Start Date':       ['Start Date', 'Ngày bắt đầu', 'Start', 'Bắt đầu', 'Ngày KH bắt đầu'],
  'Deadline':         ['Deadline', 'Hạn hoàn thành', 'Thời hạn', 'Due Date', 'Due', 'Hạn KH',
                       'Deadline ngày/tuần', 'Deadline ngày', 'Deadline tuần',
                       'Mốc thời gian', 'Mốc hoàn thành', 'Ngày hoàn thành',
                       'Tuần hoàn thành', 'Timeline', 'Thời điểm hoàn thành',
                       'Thời hạn hoàn thành', 'Hạn chót', 'Target date'],
  'Status':           ['Status', 'Trạng thái', 'Tình trạng'],
  'RAG':              ['RAG', 'Đèn RAG', 'Đèn báo', 'Cảnh báo', 'RAG Status'],
  '% Tiến độ':        ['% Tiến độ', '% Done', '% Hoàn thành', 'Progress', 'Tiến độ', 'Completion %'],
  'Update mới nhất':  ['Update mới nhất', 'Update', 'Cập nhật', 'Latest Update', 'Cập nhật mới nhất', 'Ghi chú update',
                       'Chi tiết công việc theo ngày/theo tuần', 'Chi tiết công việc', 'Chi tiết'],
  'Risk/SOS':         ['Risk/SOS', 'Risk / SOS', 'Risk', 'SOS', 'Rủi ro', 'Blocker', 'Vướng mắc'],
  'CEO Decision':     ['CEO Decision', 'CEO Decision Needed', 'CEO cần chốt', 'Cần CEO chốt', 'Cần anh Vinh chốt', 'CEO quyết định'],
  'PMO Comment':      ['PMO Comment', 'Ghi chú PMO', 'Nhận xét PMO', 'PMO Note'],
  'Evidence Link':    ['Evidence Link', 'Link bằng chứng', 'Evidence', 'File minh chứng', 'Link minh chứng'],
  'Source File':      ['Source File', 'File nguồn', 'Source'],
  'Last Update':      ['Last Update', 'Cập nhật lần cuối', 'Updated At'],
  'Update By':        ['Update By', 'Người cập nhật', 'Updated By']
};

// Đọc sheet có standard headers (row 1 = header, row 2+ = data) → internal row format
function readStandardPMORows_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return null;
  var vals = sheet.getDataRange().getValues();
  if (vals.length < 2) return null;

  var hdrs = vals[0].map(function(h) { return String(h || '').trim(); });
  var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
  var rows = [];

  vals.slice(1).forEach(function(row, i) {
    var obj = {};
    hdrs.forEach(function(h, j) { if (h) obj[h] = row[j]; });

    var rowId = getVal(obj, ['Action ID', 'Mã dòng (ID)', 'Mã dòng', 'ID', 'STT']);
    if (!hasText(rowId)) return;

    var rowUrl = ss.getUrl() + '#gid=' + sheet.getSheetId() + '&range=A' + (i + 2) + ':Z' + (i + 2);
    var deadline = formatDate(getVal(obj, ['Deadline', 'Hạn hoàn thành', 'Thời hạn', 'Due Date']));
    var status   = normalizeStatus(getVal(obj, ['Status', 'Trạng thái', 'Tình trạng']));

    // Workstream: từ cột Workstream, fallback sang PIC Name
    var ws     = String(getVal(obj, ['Workstream', 'Trục']) || '');
    var wsDisp = ws || String(getVal(obj, ['PIC Name', 'Bộ phận/Đơn vị']) || '');

    var rowObj = {
      RowId:             String(rowId),
      Workstream:        ws || wsDisp,
      WorkstreamDisplay: wsDisp,
      SourceSheet:       sheetName,
      SourceUrl:         sheetUrl,
      RowUrl:            rowUrl,
      Unit:              String(getVal(obj, ['Bộ phận/Đơn vị', 'Đơn vị / Khối', 'Đơn vị', 'Nhà máy/Khối']) || ''),
      PIC:               String(getVal(obj, ['Owner', 'PIC Name', 'PIC', 'PIC Lead', 'Đầu mối']) || ''),
      Action:            String(getVal(obj, ['Action', 'Kế hoạch hành động', 'Việc cần làm', 'Nhiệm vụ']) || ''),
      Owner:             String(getVal(obj, ['Owner', 'Người phụ trách', 'Người chịu trách nhiệm']) || ''),
      Deadline:          deadline,
      Status:            status,
      RAG:               normalizeRag(getVal(obj, ['RAG', 'Đèn RAG', 'Đèn báo', 'Cảnh báo'])),
      Risk:              String(getVal(obj, ['Risk/SOS', 'Risk / SOS', 'Risk', 'Rủi ro', 'Vướng mắc']) || ''),
      CEODecision:       String(getVal(obj, ['CEO Decision', 'CEO cần chốt', 'CEO Decision Needed', 'Cần CEO chốt']) || ''),
      PMOComment:        String(getVal(obj, ['PMO Comment', 'Ghi chú PMO', 'Nhận xét PMO']) || ''),
      // Extra fields (new model)
      PICCode:           String(getVal(obj, ['PIC Code']) || ''),
      PICName:           String(getVal(obj, ['PIC Name']) || ''),
      Progress:          String(getVal(obj, ['% Tiến độ', '% Done', '% Hoàn thành']) || ''),
      Output:            String(getVal(obj, ['Output kỳ vọng', 'Output', 'Expected Output']) || ''),
      UpdateNote:        String(getVal(obj, ['Update mới nhất', 'Cập nhật mới nhất', 'Update']) || '')
    };
    rowObj.IsOverdue    = isPastDeadline(deadline, status);
    rowObj.DataQuality  = calcRowDataQuality_(rowObj);
    rowObj.PriorityScore= calcPriorityScore_(rowObj);
    rows.push(rowObj);
  });
  return rows.length > 0 ? rows : null;
}

/* =====================================================
   DASHBOARD DATA
===================================================== */

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rows = [];
  let dataSource = 'legacy';

  // ── Priority 1: PMO_ALL_ACTIONS (new unified model) ──
  var allActRows = readStandardPMORows_(ss, 'PMO_ALL_ACTIONS');
  if (allActRows && allActRows.length > 0) {
    rows = allActRows;
    dataSource = 'PMO_ALL_ACTIONS';
  } else {
    // ── Priority 2: 16 PIC sheets ──
    PIC_SHEETS_CONFIG.forEach(function(pic) {
      var picRows = readStandardPMORows_(ss, pic.name);
      if (picRows) rows = rows.concat(picRows);
    });
    if (rows.length > 0) {
      dataSource = 'PIC_SHEETS';
    } else {
      // ── Priority 3: Fallback — 6 workstream sheets cũ ──
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
    }
  }

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
    dataSource: dataSource,
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

  // ── TO Action Center data (đọc TO_ISSUES + TO_ACTIONS) ──
  try {
    var toCenter = getToActionCenterData();
    dashboardData.toActions   = toCenter.actions   || [];
    dashboardData.toIssues    = toCenter.issues     || [];
    dashboardData.toFactories = toCenter.factories  || [];
    dashboardData.toSummary   = {
      issuesSummary:  toCenter.issuesSummary  || {},
      actionsSummary: toCenter.actionsSummary || {}
    };
  } catch (e) {
    Logger.log('getDashboardData [TO]: ' + e.message);
    dashboardData.toActions   = [];
    dashboardData.toIssues    = [];
    dashboardData.toFactories = [];
    dashboardData.toSummary   = {};
  }

  // ── Command Center Summary (các module sheet riêng) ──
  try {
    var rawProjects  = readSheetSafe('CEO_PROJECTS');
    var rawRisks     = readSheetSafe('RISK_SOS');
    var rawDecisions = readSheetSafe('CEO_DECISIONS');
    var rawMinutes   = readSheetSafe('MEETING_MINUTES');
    var ccData = {
      pmoActions:     rows,
      ceoProjects:    rawProjects.map(normalizeCEOProject_),
      toActions:      dashboardData.toActions || [],
      riskSos:        rawRisks.map(normalizeRiskSOS_),
      ceoDecisions:   rawDecisions.map(normalizeCEODecision_),
      meetingMinutes: rawMinutes.map(normalizeMeetingMinutes_)
    };
    dashboardData.commandCenterSummary = buildCommandCenterSummary(ccData);
    dashboardData.ceoProjects = ccData.ceoProjects;
  } catch(e) {
    Logger.log('getDashboardData [commandCenterSummary]: ' + e.message);
    dashboardData.commandCenterSummary = {};
    dashboardData.ceoProjects = [];
  }

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
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var data  = safeReadSheet(ss, 'TO_ACTIONS');

  var records = data.rows.map(function(r, i) {
    var deadline = formatDate(r['Deadline'] || '');
    var status   = normalizeStatus(r['Trạng thái'] || r['Status'] || '');
    return {
      actionId:    String(r['Action ID']             || ('TOA-' + String(i + 1).padStart(4, '0'))),
      factory:     String(r['Nhà máy']              || '').trim(),
      reporter:    String(r['Người ghi nhận']        || r['Nguồn báo cáo'] || '').trim(),
      reportDate:  formatDate(r['Ngày ghi nhận']     || ''),
      issueDate:   formatDate(r['Ngày phát sinh']    || r['Ngày'] || ''),
      meetingId:   String(r['Meeting ID']            || '').trim(),
      context:     String(r['Cuộc họp / Bối cảnh']  || r['Cuộc họp'] || r['Bối cảnh'] || '').trim(),
      issueId:     String(r['Issue ID liên quan']    || '').trim(),
      area:        String(r['Khu vực']              || '').trim(),
      action:      String(r['Nội dung action']       || r['Action'] || '').trim(),
      pic:         String(r['PIC thực hiện']         || r['PIC'] || '').trim(),
      owner:       String(r['Owner theo dõi']        || r['Owner'] || '').trim(),
      deadline:    deadline,
      planType:    String(r['Loại kế hoạch']        || r['Loại'] || '').trim(),
      status:      status,
      priority:    String(r['Mức độ ưu tiên']       || r['Priority'] || '').trim(),
      rag:         normalizeRag(r['RAG'] || ''),
      output:      String(r['Kết quả / Output']      || r['Kết quả'] || '').trim(),
      ceoNote:     String(r['CEO cần biết']          || '').trim(),
      evidenceLink:String(r['File nguồn']            || r['Evidence Link'] || '').trim(),
      rowUrl:      String(r['Link biên bản']         || r['Link'] || '').trim(),
      note:        String(r['Ghi chú']              || '').trim(),
      // PascalCase aliases (for rendering compatibility)
      RowId:        String(r['Action ID']            || ('TOA-' + String(i + 1).padStart(4, '0'))),
      ActionID:     String(r['Action ID']            || ('TOA-' + String(i + 1).padStart(4, '0'))),
      Date:         formatDate(r['Ngày phát sinh']   || r['Ngày'] || ''),
      MeetingID:    String(r['Meeting ID']           || '').trim(),
      IssueIDRef:   String(r['Issue ID liên quan']   || '').trim(),
      Area:         String(r['Khu vực']             || '').trim(),
      Source:       String(r['Cuộc họp / Bối cảnh'] || r['Cuộc họp'] || r['Bối cảnh'] || '').trim(),
      Action:       String(r['Nội dung action']      || r['Action'] || '').trim(),
      PIC:          String(r['PIC thực hiện']        || r['PIC'] || '').trim(),
      Owner:        String(r['Owner theo dõi']       || r['Owner'] || '').trim(),
      Deadline:     deadline,
      PlanType:     String(r['Loại kế hoạch']       || r['Loại'] || '').trim(),
      Status:       status,
      Priority:     String(r['Mức độ ưu tiên']      || r['Priority'] || '').trim(),
      RAG:          normalizeRag(r['RAG'] || ''),
      Output:       String(r['Kết quả / Output']     || r['Kết quả'] || '').trim(),
      CEONote:      String(r['CEO cần biết']         || '').trim(),
      EvidenceLink: String(r['File nguồn']           || r['Evidence Link'] || '').trim(),
      RowUrl:       String(r['Link biên bản']        || r['Link'] || '').trim(),
      Note:         String(r['Ghi chú']             || '').trim(),
      Factory:      String(r['Nhà máy']             || '').trim(),
      Reporter:     String(r['Người ghi nhận']       || r['Nguồn báo cáo'] || '').trim(),
      ReportDate:   formatDate(r['Ngày ghi nhận']    || ''),
      IsOverdue:    isPastDeadline(deadline, status)
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

/* =====================================================
   READ SHEET SAFE — Helper chuẩn cho Command Center
   Input : sheetName (string)
   Output: [] nếu sheet chưa tồn tại hoặc chưa có data
           [{col1: val, col2: val, ...}, ...] nếu có data
   Header ở hàng 1, data từ hàng 2.
   Bỏ qua dòng trống hoàn toàn. Không throw exception.
===================================================== */

function readSheetSafe(sheetName) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return [];
    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];
    var headers = values[0].map(function(h) { return String(h || '').trim(); });
    return values.slice(1)
      .filter(function(row) {
        return row.some(function(c) { return String(c || '').trim() !== ''; });
      })
      .map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) {
          obj[h] = (row[i] !== null && row[i] !== undefined) ? row[i] : '';
        });
        return obj;
      });
  } catch (e) {
    Logger.log('readSheetSafe error [' + sheetName + ']: ' + e.message);
    return [];
  }
}

/* =====================================================
   GET COMMAND CENTER DATA
   Action: getCommandCenterData
   Trả JSON đầy đủ cho toàn bộ CEO Command Center.
   Không ảnh hưởng getDashboardData (PMO cũ).
===================================================== */

function getCommandCenterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Đọc từng sheet an toàn ----
  var rawPMO        = readSheetSafe('PMO_90D_ACTIONS');
  var rawProjects   = readSheetSafe('CEO_PROJECTS');
  var rawUpdates    = readSheetSafe('CEO_PROJECT_UPDATES');
  var rawTOActions  = readSheetSafe('TO_ACTIONS');
  var rawMinutes    = readSheetSafe('MEETING_MINUTES');
  var rawDecisions  = readSheetSafe('CEO_DECISIONS');
  var rawRisks      = readSheetSafe('RISK_SOS');
  var rawMasterPIC  = readSheetSafe('MASTER_PIC');
  var rawMasterSts  = readSheetSafe('MASTER_STATUS');
  var rawConfig     = readSheetSafe('CONFIG');

  // ---- Normalize từng module ----
  var pmoActions        = rawPMO.map(normalizePMOAction_);
  var ceoProjects       = rawProjects.map(normalizeCEOProject_);
  var ceoProjectUpdates = rawUpdates.map(normalizeCEOProjectUpdate_);
  var toActions         = rawTOActions.map(normalizeTOAction_);
  var meetingMinutes    = rawMinutes.map(normalizeMeetingMinutes_);
  var ceoDecisions      = rawDecisions.map(normalizeCEODecision_);
  var riskSos           = rawRisks.map(normalizeRiskSOS_);

  // MASTER & CONFIG trả raw (web dùng để build dropdown, hiển thị nhãn)
  var masterPic    = rawMasterPIC.map(function(r) {
    return {
      Name:       String(r['Tên PIC'] || r['Name'] || r['PIC'] || ''),
      Workstream: String(r['Workstream'] || r['Trục'] || ''),
      Email:      String(r['Email'] || ''),
      Phone:      String(r['Số điện thoại'] || r['Phone'] || ''),
      Role:       String(r['Vai trò'] || r['Role'] || '')
    };
  });
  var masterStatus = rawMasterSts.map(function(r) {
    return {
      Value:    String(r['Giá trị'] || r['Value'] || r['Trạng thái'] || ''),
      Category: String(r['Nhóm'] || r['Category'] || ''),
      Color:    String(r['Màu'] || r['Color'] || ''),
      Order:    Number(r['Thứ tự'] || r['Order'] || 0)
    };
  });

  var data = {
    pmoActions:        pmoActions,
    ceoProjects:       ceoProjects,
    ceoProjectUpdates: ceoProjectUpdates,
    toActions:         toActions,
    meetingMinutes:    meetingMinutes,
    ceoDecisions:      ceoDecisions,
    riskSos:           riskSos,
    masterPic:         masterPic,
    masterStatus:      masterStatus,
    spreadsheetName:   ss.getName(),
    updatedAt:         Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
  };

  data.commandCenterSummary = buildCommandCenterSummary(data);

  return JSON.parse(JSON.stringify(data));
}

/* =====================================================
   NORMALIZE FUNCTIONS — Mỗi sheet có normalize riêng.
   Input : row object từ readSheetSafe (key = tên cột header)
   Output: object chuẩn hóa với key cố định cho web dùng.
   Dùng getVal() để hỗ trợ alias tên cột đa dạng.
===================================================== */

function normalizePMOAction_(r, i) {
  var rag      = normalizeRag(getVal(r, ['RAG', 'Đèn RAG', 'Đèn báo', 'Cảnh báo']));
  var status   = normalizeStatus(getVal(r, ['Status', 'Trạng thái', 'Tình trạng']));
  var deadline = formatDate(getVal(r, ['Deadline', 'Hạn hoàn thành', 'Thời hạn']));
  var action   = String(getVal(r, ['Hành động', 'Action', 'Kế hoạch hành động', 'Action Plan', 'Nhiệm vụ', 'Cam kết']) || '');
  var pic      = String(getVal(r, ['PIC', 'PIC Lead', 'Đầu mối', 'Người thực hiện']) || '');
  var owner    = String(getVal(r, ['Owner', 'Người phụ trách', 'Người chịu trách nhiệm']) || '');
  return {
    RowId:       String(getVal(r, ['Mã dòng', 'Mã dòng (ID)', 'ID', 'Row ID']) || ('PMO-' + String(i + 1).padStart(4, '0'))),
    Workstream:  String(getVal(r, ['Workstream', 'Trục', 'Lĩnh vực', 'Trục công việc']) || ''),
    Unit:        String(getVal(r, ['Đơn vị', 'Khối', 'Unit', 'Bộ phận']) || ''),
    Action:      action,
    PIC:         pic,
    Owner:       owner,
    Deadline:    deadline,
    Status:      status,
    RAG:         rag,
    Risk:        String(getVal(r, ['Risk/SOS', 'Risk', 'SOS', 'Rủi ro', 'Vướng mắc', 'Blocker']) || ''),
    CEODecision: String(getVal(r, ['CEO cần chốt', 'CEO Decision', 'CEO Decision Needed', 'Cần CEO chốt']) || ''),
    PMOComment:  String(getVal(r, ['PMO Comment', 'Ghi chú PMO', 'Nhận xét PMO', 'Ghi chú']) || ''),
    IsOverdue:   isPastDeadline(deadline, status),
    DQ:          calcRowDataQuality_({ Action: action, Owner: owner, PIC: pic, Deadline: deadline, Status: status, RAG: rag })
  };
}

function normalizeCEOProject_(r, i) {
  var rag      = normalizeRag(r['RAG'] || '');
  var status   = normalizeStatus(r['Trạng thái'] || r['Status'] || '');
  var deadline = formatDate(r['Deadline'] || '');
  return {
    RowId:         String(r['Project ID'] || ('PRJ-' + String(i + 1).padStart(3, '0'))),
    ProjectName:   String(r['Tên dự án']      || r['Project Name'] || ''),
    Workstream:    String(r['Workstream']      || r['Lĩnh vực'] || ''),
    PICLead:       String(r['PIC Lead']        || r['PIC'] || ''),
    Owner:         String(r['Owner']           || r['CEO phụ trách'] || ''),
    StartDate:     formatDate(r['Ngày bắt đầu'] || ''),
    Deadline:      deadline,
    Status:        status,
    RAG:           rag,
    Progress:      String(r['% Hoàn thành']    || r['Progress'] || ''),
    Milestone:     String(r['Milestone gần nhất'] || r['Milestone'] || ''),
    NextMilestone: String(r['Milestone tiếp theo'] || ''),
    Issues:        String(r['Vướng mắc']       || r['Risk/SOS'] || ''),
    CEODecision:   String(r['CEO cần chốt']    || r['CEO Decision'] || ''),
    Note:          String(r['Ghi chú']         || ''),
    RowUrl:        String(r['Link']            || r['URL'] || ''),
    IsOverdue:     isPastDeadline(deadline, status),
    DQ:            calcRowDataQuality_({ Action: r['Tên dự án'] || '', Owner: r['Owner'] || '', PIC: r['PIC Lead'] || '', Deadline: deadline, Status: status, RAG: rag })
  };
}

function normalizeCEOProjectUpdate_(r, i) {
  return {
    UpdateID:      String(r['Update ID']         || ('UPD-' + String(i + 1).padStart(4, '0'))),
    ProjectID:     String(r['Project ID']         || ''),
    Date:          formatDate(r['Ngày cập nhật']  || r['Date'] || ''),
    UpdatedBy:     String(r['PIC cập nhật']       || r['Người cập nhật'] || ''),
    Progress:      String(r['% Hoàn thành']       || r['Progress'] || ''),
    MilestoneDone: String(r['Milestone đã đạt']   || r['Milestone Done'] || ''),
    NextMilestone: String(r['Milestone tiếp theo'] || r['Next Milestone'] || ''),
    Issues:        String(r['Vướng mắc']          || r['Issues'] || ''),
    CEONote:       String(r['CEO cần biết']        || r['CEO Note'] || ''),
    RAG:           normalizeRag(r['RAG']           || ''),
    Note:          String(r['Ghi chú']            || '')
  };
}

function normalizeTOAction_(r, i) {
  var rag      = normalizeRag(r['RAG'] || '');
  var status   = normalizeStatus(r['Trạng thái'] || r['Status'] || '');
  var deadline = formatDate(r['Deadline'] || '');
  return {
    RowId:     String(r['Action ID']          || ('TOA-' + String(i + 1).padStart(4, '0'))),
    ActionID:  String(r['Action ID']          || ('TOA-' + String(i + 1).padStart(4, '0'))),
    Date:      formatDate(r['Ngày phát sinh'] || r['Ngày'] || ''),
    MeetingID: String(r['Meeting ID']         || ''),
    Meeting:   String(r['Cuộc họp']          || r['Bối cảnh'] || ''),
    Action:    String(r['Nội dung action']    || r['Action'] || ''),
    PIC:       String(r['PIC thực hiện']      || r['PIC'] || ''),
    Owner:     String(r['Owner theo dõi']     || r['Owner'] || ''),
    Deadline:  deadline,
    Status:    status,
    Priority:  String(r['Mức độ ưu tiên']    || r['Priority'] || ''),
    RAG:       rag,
    Result:    String(r['Kết quả']           || r['Output'] || ''),
    CEONote:   String(r['CEO cần biết']       || ''),
    Note:      String(r['Ghi chú']           || ''),
    RowUrl:    String(r['Link biên bản']      || r['Link'] || ''),
    IsOverdue: isPastDeadline(deadline, status),
    DQ:        calcRowDataQuality_({ Action: r['Nội dung action'] || '', Owner: r['Owner theo dõi'] || '', PIC: r['PIC thực hiện'] || '', Deadline: deadline, Status: status, RAG: rag })
  };
}

function normalizeMeetingMinutes_(r, i) {
  return {
    MeetingID:    String(r['Meeting ID']           || ('MTG-' + String(i + 1).padStart(3, '0'))),
    Date:         formatDate(r['Ngày họp']         || ''),
    Topic:        String(r['Chủ đề họp']           || r['Chủ đề'] || ''),
    Chair:        String(r['Người chủ trì']        || ''),
    Attendees:    String(r['Thành phần tham dự']   || r['Thành phần'] || ''),
    Scope:        String(r['Phạm vi']              || ''),
    Summary:      String(r['Tóm tắt chính']        || r['Tóm tắt'] || ''),
    CEODirective: String(r['Chỉ đạo CEO']          || ''),
    Decisions:    String(r['Quyết định đã chốt']   || ''),
    Actions:      String(r['Action phát sinh']      || ''),
    RiskSOS:      String(r['Risk/SOS phát sinh']    || r['Risk SOS phát sinh'] || ''),
    CEODecision:  String(r['CEO cần quyết tiếp']    || ''),
    FileMinutes:  String(r['File biên bản']         || ''),
    FileAudio:    String(r['File ghi âm']           || r['File transcript'] || ''),
    RecordedBy:   String(r['Người lập biên bản']    || ''),
    Status:       String(r['Trạng thái']            || 'Nháp'),
    SecretNote:   String(r['Ghi chú bảo mật']       || ''),
    DQ:           calcRowDataQuality_({ Action: r['Chủ đề họp'] || '', Owner: r['Người lập biên bản'] || '', PIC: r['Người chủ trì'] || '', Deadline: r['Ngày họp'] || '', Status: r['Trạng thái'] || '', RAG: '' })
  };
}

function normalizeCEODecision_(r, i) {
  var deadline = formatDate(r['Deadline'] || r['Hạn chốt'] || '');
  var status   = String(r['Trạng thái'] || r['Status'] || '');
  return {
    RowId:        String(r['Decision ID']         || ('DEC-' + String(i + 1).padStart(4, '0'))),
    DecisionID:   String(r['Decision ID']         || ('DEC-' + String(i + 1).padStart(4, '0'))),
    Source:       String(r['Nguồn']              || r['Source'] || ''),
    SourceID:     String(r['Source ID']          || r['Mã nguồn'] || ''),
    Workstream:   String(r['Workstream']         || r['Trục'] || ''),
    Content:      String(r['Nội dung cần chốt']  || r['Content'] || ''),
    Proposer:     String(r['PIC đề xuất']        || r['PIC'] || ''),
    Deadline:     deadline,
    Status:       status,
    Decision:     String(r['Quyết định CEO']     || r['CEO Decision'] || ''),
    DecisionDate: formatDate(r['Ngày chốt']      || r['Decision Date'] || ''),
    Note:         String(r['Ghi chú']            || ''),
    IsOverdue:    isPastDeadline(deadline, status),
    IsPending:    /chờ|pending|chưa/i.test(status)
  };
}

function normalizeRiskSOS_(r, i) {
  var rag      = normalizeRag(r['RAG'] || r['Mức độ'] || '');
  var deadline = formatDate(r['Deadline xử lý'] || r['Deadline'] || '');
  var status   = String(r['Trạng thái'] || r['Status'] || '');
  return {
    RowId:       String(r['Risk ID']             || ('RSK-' + String(i + 1).padStart(4, '0'))),
    RiskID:      String(r['Risk ID']             || ('RSK-' + String(i + 1).padStart(4, '0'))),
    Source:      String(r['Nguồn']              || r['Source'] || ''),
    SourceID:    String(r['Source ID']          || ''),
    Workstream:  String(r['Workstream']         || r['Trục'] || ''),
    PIC:         String(r['PIC xử lý']          || r['PIC'] || ''),
    Description: String(r['Mô tả rủi ro']       || r['Description'] || r['Risk'] || ''),
    Level:       String(r['Mức độ']             || r['Level'] || ''),
    RAG:         rag,
    Deadline:    deadline,
    Status:      status,
    Action:      String(r['Biện pháp xử lý']    || r['Action'] || ''),
    CEONote:     String(r['CEO cần biết']        || ''),
    IsOverdue:   isPastDeadline(deadline, status),
    IsOpen:      !/done|hoàn thành|đã xử lý|resolved/i.test(status)
  };
}

/* =====================================================
   BUILD COMMAND CENTER SUMMARY
   Tổng hợp số liệu từ tất cả module cho Overview.
   Input : data object từ getCommandCenterData()
   Output: summary stats object
===================================================== */

function buildCommandCenterSummary(data) {
  var pmo      = data.pmoActions        || [];
  var projects = data.ceoProjects       || [];
  var toActs   = data.toActions         || [];
  var mm       = data.meetingMinutes    || [];
  var decs     = data.ceoDecisions      || [];
  var risks    = data.riskSos           || [];

  // ---- PMO ----
  var pmoOverdue = pmo.filter(function(r) { return r.IsOverdue; }).length;
  var pmoRed     = pmo.filter(function(r) { return r.RAG === 'Red'; }).length;
  var pmoAmber   = pmo.filter(function(r) { return r.RAG === 'Amber'; }).length;
  var pmoGreen   = pmo.filter(function(r) { return r.RAG === 'Green'; }).length;
  var pmoDone    = pmo.filter(function(r) { return r.Status === 'Done'; }).length;

  // ---- CEO Projects ----
  var projActive  = projects.filter(function(r) { return r.Status === 'In Progress'; }).length;
  var projDone    = projects.filter(function(r) { return r.Status === 'Done'; }).length;
  var projOverdue = projects.filter(function(r) { return r.IsOverdue; }).length;
  var projRed     = projects.filter(function(r) { return r.RAG === 'Red'; }).length;

  // ---- TO Actions ----
  var toOverdue = toActs.filter(function(r) { return r.IsOverdue; }).length;
  var toDone    = toActs.filter(function(r) { return r.Status === 'Done'; }).length;
  var toRed     = toActs.filter(function(r) { return r.RAG === 'Red'; }).length;

  // ---- Meeting Minutes ----
  var mmPending    = mm.filter(function(r) { return /chờ|pending/i.test(r.Status || ''); }).length;
  var mmHasCEO     = mm.filter(function(r) { return hasText(r.CEODirective); }).length;
  var mmHasRisk    = mm.filter(function(r) { return hasText(r.RiskSOS); }).length;
  var mmHasDecision= mm.filter(function(r) { return hasText(r.CEODecision); }).length;

  // ---- CEO Decisions ----
  var decPending  = decs.filter(function(r) { return r.IsPending; }).length;
  var decOverdue  = decs.filter(function(r) { return r.IsOverdue; }).length;
  var decDone     = decs.filter(function(r) { return !r.IsPending; }).length;

  // ---- Risk/SOS ----
  var riskOpen    = risks.filter(function(r) { return r.IsOpen; }).length;
  var riskRed     = risks.filter(function(r) { return r.RAG === 'Red'; }).length;
  var riskAmber   = risks.filter(function(r) { return r.RAG === 'Amber'; }).length;
  var riskOverdue = risks.filter(function(r) { return r.IsOverdue; }).length;

  // ---- Data Quality across all ----
  var allRows = [].concat(pmo, projects, toActs);
  var dqTotal   = allRows.length;
  var dqAvgScore= dqTotal > 0
    ? Math.round(allRows.reduce(function(s, r) { return s + (r.DQ || 0); }, 0) / dqTotal)
    : 0;

  // ---- Overall health ----
  var totalRed = pmoRed + projRed + toRed + riskRed;
  var health = 'Green';
  if (totalRed > 0 || riskOpen > 2) health = 'Red';
  else if (pmoAmber > 0 || projOverdue > 0 || decOverdue > 0) health = 'Amber';

  return {
    // PMO
    pmoTotal:       pmo.length,
    pmoDone:        pmoDone,
    pmoOverdue:     pmoOverdue,
    pmoRed:         pmoRed,
    pmoAmber:       pmoAmber,
    pmoGreen:       pmoGreen,
    pmoCompletionRate: pmo.length > 0 ? Math.round(pmoDone / pmo.length * 100) : 0,

    // CEO Projects
    projectsTotal:   projects.length,
    projectsActive:  projActive,
    projectsDone:    projDone,
    projectsOverdue: projOverdue,
    projectsRed:     projRed,

    // TO Actions
    toActionsTotal:   toActs.length,
    toActionsDone:    toDone,
    toActionsOverdue: toOverdue,
    toActionsRed:     toRed,

    // Meetings
    meetingsTotal:      mm.length,
    meetingsPending:    mmPending,
    meetingsWithCEO:    mmHasCEO,
    meetingsWithRisk:   mmHasRisk,
    meetingsWithDecision: mmHasDecision,

    // CEO Decisions
    decisionsTotal:   decs.length,
    decisionsPending: decPending,
    decisionsOverdue: decOverdue,
    decisionsDone:    decDone,

    // Risks
    risksTotal:   risks.length,
    risksOpen:    riskOpen,
    risksRed:     riskRed,
    risksAmber:   riskAmber,
    risksOverdue: riskOverdue,

    // DQ & Health
    dataQualityScore: dqAvgScore,
    overallHealth:    health,

    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
  };
}

/* =====================================================
   TEST — Command Center
===================================================== */

function testGetCommandCenterData() {
  var data = getCommandCenterData();
  var s    = data.commandCenterSummary;
  Logger.log('=== COMMAND CENTER SUMMARY ===');
  Logger.log('PMO: ' + data.pmoActions.length + ' actions | Overdue: ' + s.pmoOverdue + ' | Red: ' + s.pmoRed);
  Logger.log('Projects: ' + data.ceoProjects.length + ' | Active: ' + s.projectsActive + ' | Overdue: ' + s.projectsOverdue);
  Logger.log('TO Actions: ' + data.toActions.length + ' | Overdue: ' + s.toActionsOverdue);
  Logger.log('Meetings: ' + data.meetingMinutes.length + ' | Pending: ' + s.meetingsPending);
  Logger.log('Decisions: ' + data.ceoDecisions.length + ' | Pending: ' + s.decisionsPending);
  Logger.log('Risks: ' + data.riskSos.length + ' | Open: ' + s.risksOpen + ' | Red: ' + s.risksRed);
  Logger.log('Overall Health: ' + s.overallHealth + ' | DQ Score: ' + s.dataQualityScore + '%');
  Logger.log('Master PIC: ' + data.masterPic.length + ' | Status: ' + data.masterStatus.length);
}

function testReadSheetSafe() {
  var sheets = [
    'PMO_90D_ACTIONS', 'CEO_PROJECTS', 'CEO_PROJECT_UPDATES',
    'TO_ACTIONS', 'MEETING_MINUTES', 'CEO_DECISIONS',
    'RISK_SOS', 'MASTER_PIC', 'MASTER_STATUS', 'CONFIG'
  ];
  sheets.forEach(function(name) {
    var rows = readSheetSafe(name);
    Logger.log(name + ': ' + rows.length + ' dòng' + (rows.length === 0 ? ' (chưa tạo hoặc chưa có data)' : ''));
  });
}


/* ================================================================
   PMO — SETUP, IMPORT, SYNC
================================================================ */

// Tạo 16 PIC sheets, PMO_ALL_ACTIONS, PMO_IMPORT_LOG, PMO_SOURCE_FILES
function setupPicSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];

  PIC_SHEETS_CONFIG.forEach(function(pic) {
    if (!ss.getSheetByName(pic.name)) {
      var sh = ss.insertSheet(pic.name);
      sh.getRange(1, 1, 1, PMO_STANDARD_HEADERS.length).setValues([PMO_STANDARD_HEADERS]);
      sh.setFrozenRows(1);
      created.push(pic.name);
    }
  });

  if (!ss.getSheetByName('PMO_ALL_ACTIONS')) {
    var sh = ss.insertSheet('PMO_ALL_ACTIONS');
    sh.getRange(1, 1, 1, PMO_STANDARD_HEADERS.length).setValues([PMO_STANDARD_HEADERS]);
    sh.setFrozenRows(1);
    created.push('PMO_ALL_ACTIONS');
  }

  if (!ss.getSheetByName('PMO_IMPORT_LOG')) {
    var logHdr = ['Timestamp', 'PIC Code', 'PIC Name', 'Source URL', 'Sheet Name', 'Rows Imported', 'Status', 'Error'];
    var sh = ss.insertSheet('PMO_IMPORT_LOG');
    sh.getRange(1, 1, 1, logHdr.length).setValues([logHdr]);
    sh.setFrozenRows(1);
    created.push('PMO_IMPORT_LOG');
  }

  if (!ss.getSheetByName('PMO_SOURCE_FILES')) {
    var srcHdr = ['PIC Code', 'PIC Name', 'Source URL', 'Sheet Name', 'Last Imported', 'Row Count', 'Active'];
    var sh = ss.insertSheet('PMO_SOURCE_FILES');
    sh.getRange(1, 1, 1, srcHdr.length).setValues([srcHdr]);
    sh.setFrozenRows(1);
    created.push('PMO_SOURCE_FILES');
  }

  Logger.log('setupPicSheets: ' + (created.length > 0 ? 'Đã tạo: ' + created.join(', ') : 'Tất cả sheets đã tồn tại'));
  return { created: created };
}

/* ----------------------------------------------------------------
   importPicPlanToDatabase — đọc file Google Sheets nguồn của PIC,
   map về header chuẩn, ghi vào PIC sheet và log.

   Params:
     sourceUrl    — URL của Google Sheets file PIC
     picSheetName — tên sheet trong file nguồn (null = tự tìm)
     picCode      — VD: 'PIC01'
     picName      — tên hiển thị (null = lấy từ config)
     forceReimport— true = ghi đè ngay cả khi đã import trước đó

   Returns: { imported, picCode, targetSheet } hoặc { skipped, reason }
---------------------------------------------------------------- */
// Normalize header text: bỏ newline, dấu tiếng Việt, trim khoảng trắng,
// bỏ ký tự đặc biệt → dùng để so sánh fuzzy giữa alias và header file nguồn.
function normHeader_(s) {
  return String(s || '')
    .replace(/\r?\n|\r/g, ' ')           // newline → space
    .replace(/\s+/g, ' ').trim()
    .replace(/đ/gi, 'd')                 // đ/Đ không decompose được bằng NFD
    .normalize('NFD')                    // tách dấu ra khỏi ký tự gốc
    .replace(/[̀-ͯ]/g, '')     // bỏ dấu combining
    .toLowerCase()
    .replace(/[^a-z0-9 \/%+\-]/g, ' ')  // bỏ ký tự lạ, giữ / % + -
    .replace(/\s+/g, ' ').trim();
}

// Tìm dòng header của bảng action bằng từ khóa, không chỉ đếm số ô.
function findActionHeaderRow_(values) {
  var keywords = ['stt', 'mã', 'hành động', 'công việc', 'deadline', 'owner',
                  'người phụ trách', 'phối hợp', 'kết quả', 'output'];
  var best = { idx: 0, score: 0 };
  for (var i = 0; i < Math.min(25, values.length); i++) {
    var rowStr = values[i].map(function(c) { return normHeader_(c); }).join('|');
    var score = 0;
    keywords.forEach(function(kw) { if (rowStr.indexOf(kw) >= 0) score++; });
    if (score > best.score) best = { idx: i, score: score };
  }
  if (best.score >= 2) return best.idx;
  // Fallback: dòng nhiều ô nhất trong 10 dòng đầu
  var maxCells = 0, fallback = 0;
  for (var i = 0; i < Math.min(10, values.length); i++) {
    var cnt = values[i].filter(function(c) { return String(c || '').trim() !== ''; }).length;
    if (cnt > maxCells) { maxCells = cnt; fallback = i; }
  }
  return fallback;
}

// Trích xuất metadata từ phần mô tả đầu file (trước dòng header bảng action).
function extractPicMetadata_(values, headerRowIdx) {
  var meta = { department: '', workstream: '', goal90: '' };

  // Tìm value kế tiếp label trong cùng 1 row
  function findValueInRow(row, labelKeywords) {
    for (var j = 0; j < row.length; j++) {
      var cell = String(row[j] || '').trim();
      var cellLow = cell.toLowerCase();
      for (var ki = 0; ki < labelKeywords.length; ki++) {
        if (cellLow.indexOf(labelKeywords[ki]) >= 0) {
          // Thử lấy phần sau dấu ':'
          var colonIdx = cell.indexOf(':');
          if (colonIdx >= 0 && cell.substring(colonIdx + 1).trim()) {
            return cell.substring(colonIdx + 1).trim();
          }
          // Tìm ô không rỗng tiếp theo trong cùng row
          for (var k = j + 1; k < row.length; k++) {
            var v = String(row[k] || '').trim();
            if (v) return v;
          }
        }
      }
    }
    return '';
  }

  for (var i = 0; i < Math.min(headerRowIdx, 25); i++) {
    var row = values[i];
    if (!meta.department) meta.department = findValueInRow(row, ['bộ phận', 'department']);
    if (!meta.workstream) meta.workstream = findValueInRow(row, ['workstream', 'trục chiến lược', 'trục']);
    if (!meta.goal90)     meta.goal90     = findValueInRow(row, ['mục tiêu 90', 'mục tiêu']);
  }
  return meta;
}

function importPicPlanToDatabase(sourceUrl, picSheetName, picCode, picName, forceReimport) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ts = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');

  if (!sourceUrl) throw new Error('sourceUrl là bắt buộc');
  if (!picCode)   throw new Error('picCode là bắt buộc');

  var picCfg = null;
  for (var pi = 0; pi < PIC_SHEETS_CONFIG.length; pi++) {
    if (PIC_SHEETS_CONFIG[pi].code === picCode) { picCfg = PIC_SHEETS_CONFIG[pi]; break; }
  }
  if (!picCfg) throw new Error('Không tìm thấy PIC config cho code: ' + picCode);

  var targetSheetName = picCfg.name;
  if (!picName) picName = picCfg.label;

  // Kiểm tra đã import chưa (trừ khi forceReimport)
  if (!forceReimport) {
    var srcFileSh = ss.getSheetByName('PMO_SOURCE_FILES');
    if (srcFileSh && srcFileSh.getLastRow() > 1) {
      var sfVals = srcFileSh.getDataRange().getValues();
      for (var si = 1; si < sfVals.length; si++) {
        if (String(sfVals[si][0]) === picCode && String(sfVals[si][2]) === sourceUrl) {
          Logger.log('File đã import. Dùng forceReimport=true để import lại.');
          return { skipped: true, reason: 'Already imported on ' + sfVals[si][4] };
        }
      }
    }
  }

  // Mở file nguồn
  var srcSS;
  try { srcSS = SpreadsheetApp.openByUrl(sourceUrl); }
  catch(e) {
    pmoLogImport_(ss, ts, picCode, picName, sourceUrl, picSheetName || '', 0, 'ERROR', 'Cannot open: ' + e.message);
    throw new Error('Không mở được file nguồn: ' + e.message);
  }

  // Tìm sheet nguồn
  var srcSheet;
  if (picSheetName) {
    srcSheet = srcSS.getSheetByName(picSheetName);
    if (!srcSheet) {
      pmoLogImport_(ss, ts, picCode, picName, sourceUrl, picSheetName, 0, 'ERROR', 'Sheet not found: ' + picSheetName);
      throw new Error('Sheet "' + picSheetName + '" không tồn tại trong file nguồn');
    }
  } else {
    // Lấy sheet có nhiều dữ liệu nhất
    var sheets = srcSS.getSheets();
    var maxRows = 0;
    sheets.forEach(function(sh) {
      var lr = sh.getLastRow();
      if (lr > maxRows) { maxRows = lr; srcSheet = sh; }
    });
    if (!srcSheet) throw new Error('File không có sheet nào');
    picSheetName = srcSheet.getName();
  }

  // Đọc dữ liệu nguồn
  var values = srcSheet.getDataRange().getValues();
  if (values.length < 2) {
    pmoLogImport_(ss, ts, picCode, picName, sourceUrl, picSheetName, 0, 'WARN', 'Sheet empty');
    return { imported: 0, warning: 'Sheet nguồn không có dữ liệu' };
  }

  // Tìm header row bằng từ khóa (thay vì chỉ đếm ô)
  var headerRowIdx = findActionHeaderRow_(values);
  var srcHeaders = values[headerRowIdx].map(function(h) { return String(h || '').trim(); });

  // Trích xuất metadata từ phần đầu file (Bộ phận, Workstream, Mục tiêu 90 ngày)
  var meta = extractPicMetadata_(values, headerRowIdx);
  Logger.log('[importPic] headerRowIdx=' + headerRowIdx + ' | dept=' + meta.department + ' | ws=' + meta.workstream);

  // Normalize source headers (bỏ newline, dấu tiếng Việt, ký tự lạ)
  var srcHeadersNorm = srcHeaders.map(function(h) { return normHeader_(h); });
  Logger.log('[importPic] srcHeadersNorm=' + JSON.stringify(srcHeadersNorm));

  // Build column index map: stdCol → source column index (so sánh qua normHeader_)
  var colIdx = {};
  Object.keys(PMO_COLUMN_ALIASES).forEach(function(stdCol) {
    var aliases = PMO_COLUMN_ALIASES[stdCol];
    for (var ai = 0; ai < aliases.length; ai++) {
      var aliasNorm = normHeader_(aliases[ai]);
      for (var ci = 0; ci < srcHeadersNorm.length; ci++) {
        if (srcHeadersNorm[ci] === aliasNorm && srcHeadersNorm[ci] !== '') {
          if (!(stdCol in colIdx)) colIdx[stdCol] = ci;
          break;
        }
      }
    }
  });

  // Log chi tiết deadline
  var dlIdx = colIdx['Deadline'];
  var dlHeader = dlIdx !== undefined ? srcHeaders[dlIdx] : '(không tìm thấy)';
  var dlSamples = [];
  if (dlIdx !== undefined) {
    for (var si = 1; si <= 5 && (headerRowIdx + si) < values.length; si++) {
      dlSamples.push(String(values[headerRowIdx + si][dlIdx] || '').trim());
    }
  }
  Logger.log('[importPic] deadlineColumnIndex=' + (dlIdx !== undefined ? dlIdx : -1));
  Logger.log('[importPic] deadlineHeaderMatched="' + dlHeader + '"');
  Logger.log('[importPic] deadline5Samples=' + JSON.stringify(dlSamples));
  Logger.log('[importPic] colIdx=' + JSON.stringify(colIdx));

  // Map từng row dữ liệu sang standard headers
  var outputRows = [];
  var rowNum = 0;
  for (var ri = headerRowIdx + 1; ri < values.length; ri++) {
    var srcRow = values[ri];
    var nonEmpty = srcRow.filter(function(c) { return String(c || '').trim() !== ''; }).length;
    if (nonEmpty === 0) continue;
    rowNum++;

    var newRow = PMO_STANDARD_HEADERS.map(function(col) {
      if (col === 'PIC Code')        return picCode;
      if (col === 'PIC Name')        return picName;
      if (col === 'Source File')     return srcSS.getName() + ' | ' + picSheetName;
      if (col === 'Last Update')     return ts;
      if (col === 'Update By')       return 'PMO Import';
      // Metadata từ đầu file
      if (col === 'Bộ phận/Đơn vị') return meta.department || '';
      if (col === 'Workstream')      return meta.workstream || '';

      if (col === 'Action ID') {
        var rawId = colIdx[col] !== undefined ? String(srcRow[colIdx[col]] || '').trim() : '';
        return rawId || (picCode + '-' + String(rowNum).padStart(4, '0'));
      }
      if (col === 'Status') {
        var sv = colIdx[col] !== undefined ? String(srcRow[colIdx[col]] || '').trim() : '';
        return sv || 'Chưa bắt đầu';
      }
      if (col === 'RAG') {
        // Không tự bịa — chỉ lấy từ file nếu có
        return colIdx[col] !== undefined ? String(srcRow[colIdx[col]] || '').trim() : '';
      }
      if (colIdx[col] !== undefined) {
        var val = srcRow[colIdx[col]];
        if (val instanceof Date) {
          // Deadline là ngày thật → format chuẩn
          return Utilities.formatDate(val, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
        }
        var strVal = String(val === null || val === undefined ? '' : val).trim();
        return strVal;
      }
      return '';
    });
    outputRows.push(newRow);
  }

  if (outputRows.length === 0) {
    pmoLogImport_(ss, ts, picCode, picName, sourceUrl, picSheetName, 0, 'WARN', 'No data rows after mapping');
    return { imported: 0, warning: 'Không tìm thấy dòng dữ liệu sau khi map cột' };
  }

  // Ghi vào PIC sheet
  var targetSheet = ss.getSheetByName(targetSheetName) || ss.insertSheet(targetSheetName);
  targetSheet.clearContents();
  var allData = [PMO_STANDARD_HEADERS].concat(outputRows);
  targetSheet.getRange(1, 1, allData.length, PMO_STANDARD_HEADERS.length).setValues(allData);
  targetSheet.setFrozenRows(1);

  // Log
  pmoLogImport_(ss, ts, picCode, picName, sourceUrl, picSheetName, outputRows.length, 'OK', '');
  pmoUpdateSourceFiles_(ss, picCode, picName, sourceUrl, picSheetName, ts, outputRows.length);

  Logger.log('importPicPlanToDatabase: ' + outputRows.length + ' rows → ' + targetSheetName);
  return { imported: outputRows.length, picCode: picCode, targetSheet: targetSheetName };
}

function pmoLogImport_(ss, ts, picCode, picName, url, sheetName, rowCount, status, error) {
  var sh = ss.getSheetByName('PMO_IMPORT_LOG');
  if (!sh) return;
  sh.appendRow([ts, picCode, picName, url, sheetName, rowCount, status, error]);
}

function pmoUpdateSourceFiles_(ss, picCode, picName, url, sheetName, ts, rowCount) {
  var sh = ss.getSheetByName('PMO_SOURCE_FILES');
  if (!sh) return;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === picCode && String(data[i][2]) === url) {
      sh.getRange(i + 1, 1, 1, 7).setValues([[picCode, picName, url, sheetName, ts, rowCount, 'TRUE']]);
      return;
    }
  }
  sh.appendRow([picCode, picName, url, sheetName, ts, rowCount, 'TRUE']);
}

/* ----------------------------------------------------------------
   syncPicSheetsToPmoAllActions — gom dữ liệu từ 16 PIC sheets
   vào PMO_ALL_ACTIONS. Xóa dữ liệu cũ, ghi lại toàn bộ.
---------------------------------------------------------------- */
function syncPicSheetsToPmoAllActions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ts = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
  var allRows = [];
  var picReport = [];

  PIC_SHEETS_CONFIG.forEach(function(pic) {
    var sheet = ss.getSheetByName(pic.name);
    if (!sheet || sheet.getLastRow() < 2) {
      picReport.push({ pic: pic.code, rows: 0 });
      return;
    }
    var vals = sheet.getDataRange().getValues();
    var hdrs = vals[0].map(function(h) { return String(h || '').trim(); });
    // Chỉ đọc sheet có standard header
    if (hdrs[0] !== 'Action ID') {
      picReport.push({ pic: pic.code, rows: 0, skip: 'No standard header' });
      return;
    }
    var count = 0;
    vals.slice(1).forEach(function(row) {
      var nonEmpty = row.filter(function(c) { return String(c || '').trim() !== ''; }).length;
      if (nonEmpty === 0) return;
      // Map theo position (PIC sheets dùng standard headers)
      var mapped = PMO_STANDARD_HEADERS.map(function(h) {
        var idx = hdrs.indexOf(h);
        return idx >= 0 ? row[idx] : '';
      });
      allRows.push(mapped);
      count++;
    });
    picReport.push({ pic: pic.code, rows: count });
  });

  var allSheet = ss.getSheetByName('PMO_ALL_ACTIONS') || ss.insertSheet('PMO_ALL_ACTIONS');
  allSheet.clearContents();
  var output = [PMO_STANDARD_HEADERS].concat(allRows);
  allSheet.getRange(1, 1, output.length, PMO_STANDARD_HEADERS.length).setValues(output);
  allSheet.setFrozenRows(1);

  var total = allRows.length;
  Logger.log('syncPicSheetsToPmoAllActions: ' + total + ' rows → PMO_ALL_ACTIONS (' + ts + ')');
  Logger.log('Per-PIC: ' + JSON.stringify(picReport));
  return { total: total, updatedAt: ts, perPic: picReport };
}

/* ----------------------------------------------------------------
   getPicPmoSummary — trả JSON tóm tắt theo từng PIC (dùng cho web)
---------------------------------------------------------------- */
function getPicPmoSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = [];

  PIC_SHEETS_CONFIG.forEach(function(pic) {
    var rows = readStandardPMORows_(ss, pic.name);
    if (!rows) { result.push({ code: pic.code, label: pic.label, total: 0 }); return; }
    var done    = rows.filter(function(r) { return r.Status === 'Done'; }).length;
    var overdue = rows.filter(function(r) { return r.IsOverdue; }).length;
    var red     = rows.filter(function(r) { return r.RAG === 'Red'; }).length;
    result.push({ code: pic.code, label: pic.label, total: rows.length, done: done, overdue: overdue, red: red });
  });

  return { pics: result, updatedAt: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss') };
}

/* ----------------------------------------------------------------
   TEST helpers
---------------------------------------------------------------- */
function testSetupPicSheets() {
  var r = setupPicSheets();
  Logger.log('Created: ' + JSON.stringify(r.created));
}

function testSyncPicSheets() {
  var r = syncPicSheetsToPmoAllActions();
  Logger.log('Sync result: ' + JSON.stringify(r));
}

// ── Chạy từng function trong Apps Script Editor để import file PIC ──
// Dán link Google Sheets thật vào mỗi function rồi Run.

function testImportPIC01() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1wcw0aNb8fpK-g4Ed1vSv32Rbs54VRYR9NZ6EaR4UJU8/edit?gid=674565336#gid=674565336', null, 'PIC01', null, true);
  Logger.log('PIC01: ' + JSON.stringify(r));
}

function testImportPIC02() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1ovnXqPb_z9xjFnJveyNosRZmECqDzVrzQWQGyyeDv_M/edit?gid=1025666321#gid=1025666321', null, 'PIC02', null, false);
  Logger.log('PIC02: ' + JSON.stringify(r));
}

function testImportPIC03() {
  var r = importPicPlanToDatabase('DAN_LINK_PIC03_VAO_DAY', null, 'PIC03', null, false);
  Logger.log('PIC03: ' + JSON.stringify(r));
}

function testImportPIC04() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1iE7uZEW2K8brwYEowtoGPg-hCxY7WpqyA9ICuQMXygY/edit?gid=2025716379#gid=2025716379', null, 'PIC04', null, false);
  Logger.log('PIC04: ' + JSON.stringify(r));
}

function testImportPIC05() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1qzxOFNJJri6d9R8FRuMgvOwWcPzFkTTYq-CflA49jRk/edit?gid=1330008661#gid=1330008661', null, 'PIC05', null, false);
  Logger.log('PIC05: ' + JSON.stringify(r));
}

function testImportPIC06() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1xWGxzAqeGW2W7ZEwopYhHHDpNzpyxmktXp-qbxSX3XU/edit?gid=1106585649#gid=1106585649', null, 'PIC06', null, false);
  Logger.log('PIC06: ' + JSON.stringify(r));
}

function testImportPIC07() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1jR4v5KHwFC9q3wfdRE8navbKtOnJ5MQJQP2pThNNP0c/edit?gid=980701653#gid=980701653', null, 'PIC07', null, false);
  Logger.log('PIC07: ' + JSON.stringify(r));
}

function testImportPIC08() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1oaDzR7nwfzaOog-knI-GSvym58B0NinOb-tgW5gWp58/edit?gid=504626587#gid=504626587', null, 'PIC08', null, false);
  Logger.log('PIC08: ' + JSON.stringify(r));
}

function testImportPIC09() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1kY1Y2doayQgXg54j9BH5xsH3NHsZyr2tqZmtwGLQh0g/edit?gid=1970385017#gid=1970385017', null, 'PIC09', null, false);
  Logger.log('PIC09: ' + JSON.stringify(r));
}

function testImportPIC10() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1RcaiX-MR-taWUD5hxbWGj_aH4_Ra0TDgTWewWEem-fM/edit?gid=2057573109#gid=2057573109', null, 'PIC10', null, false);
  Logger.log('PIC10: ' + JSON.stringify(r));
}

function testImportPIC11() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1jWggECv-GhG_ybk_7NishuW2Lb0cXeXlXayOzFv5KgU/edit?gid=1400779779#gid=1400779779', null, 'PIC11', null, false);
  Logger.log('PIC11: ' + JSON.stringify(r));
}

function testImportPIC12() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1qyEtg8SD__spubxLgW8Vr5O2JI-nrGTgAJHYboKlwfQ/edit?gid=1262195897#gid=1262195897', null, 'PIC12', null, false);
  Logger.log('PIC12: ' + JSON.stringify(r));
}

function testImportPIC13() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/181muxohTIWaitQF7xsaJGMRG2BuY517zbduoEG4pIvw/edit?gid=1120467208#gid=1120467208', null, 'PIC13', null, false);
  Logger.log('PIC13: ' + JSON.stringify(r));
}

function testImportPIC14() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1ux9-bBoG_Z5WLlWwVSS2ensivMNOeCPEsCleH8eIIzk/edit?gid=867676644#gid=867676644', null, 'PIC14', null, false);
  Logger.log('PIC14: ' + JSON.stringify(r));
}

function testImportPIC15() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1hhrMgvXlv9JXSblPizyCk34CFP889k7KcEcTIT2HkqE/edit?gid=681657984#gid=681657984', null, 'PIC15', null, false);
  Logger.log('PIC15: ' + JSON.stringify(r));
}

function testImportPIC16() {
  var r = importPicPlanToDatabase('https://docs.google.com/spreadsheets/d/1GhMrKtUnoRegr8Hf7ThfivAggppafJ-65oTQI5qJAdg/edit?gid=46456607#gid=46456607', null, 'PIC16', null, false);
  Logger.log('PIC16: ' + JSON.stringify(r));
}

function testGetDashboardDataSource() {
  var data = getDashboardData();
  Logger.log('dataSource: ' + data.dataSource + ' | rows: ' + (data.summary ? data.summary.totalActions : '?'));
}

/* ================================================================
   SETUP DATABASE — setupGspNext30Database()
   Tạo và cấu hình toàn bộ sheets cho GSP NEXT 30 Command Center
   - Tạo sheet nếu chưa tồn tại
   - Ghi header chuẩn chỉ cho sheet mới tạo
   - Freeze row 1, định dạng header xanh đậm
   - Tạo dropdown validation cho các cột chuẩn
   - KHÔNG xóa dữ liệu cũ nếu sheet đã tồn tại
================================================================ */

function setupGspNext30Database() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];
  var created = 0;
  var updated = 0;

  var SHEETS = getDbSheetConfigs_();

  SHEETS.forEach(function(cfg) {
    var isNew = !ss.getSheetByName(cfg.name);
    var sheet;

    if (isNew) {
      sheet = ss.insertSheet(cfg.name);
      // Ghi headers chỉ khi tạo mới
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      // Seed data cho master sheets
      if (cfg.seedData && cfg.seedData.length > 0) {
        sheet.getRange(2, 1, cfg.seedData.length, cfg.seedData[0].length).setValues(cfg.seedData);
      }
      created++;
      log.push('[TẠO MỚI] ' + cfg.name + ' — ' + cfg.headers.length + ' cột');
    } else {
      sheet = ss.getSheetByName(cfg.name);
      updated++;
      log.push('[ĐÃ TỒN TẠI] ' + cfg.name + ' — giữ nguyên dữ liệu, cập nhật định dạng');
    }

    // Áp dụng định dạng header (an toàn cho cả sheet mới và cũ)
    applyDbHeaderFormat_(sheet, cfg.headers.length);

    // Áp dụng dropdown validations
    applyDbValidations_(sheet, cfg.headers, cfg.validations || {});

    // Auto resize columns
    try { sheet.autoResizeColumns(1, cfg.headers.length); } catch (e) {}
  });

  var summary = '=== GSP NEXT 30 Database Setup ===\n' +
    'Tạo mới: ' + created + ' sheet | Đã tồn tại: ' + updated + ' sheet\n\n' +
    log.join('\n');
  Logger.log(summary);

  try {
    ss.toast('Tạo mới: ' + created + ' | Cập nhật định dạng: ' + updated + ' | Xem log để biết chi tiết.',
             'GSP NEXT 30 Database Setup ✓', 8);
  } catch (e) {}

  return { created: created, updated: updated, log: log };
}

/* ----------------------------------------------------------------
   Helper: Định dạng header row (freeze + màu nền + chữ trắng bold)
---------------------------------------------------------------- */
function applyDbHeaderFormat_(sheet, numCols) {
  if (numCols < 1) return;
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#064e3b');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setVerticalAlignment('middle');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(false);
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(1);
}

/* ----------------------------------------------------------------
   Helper: Tạo dropdown data validation theo tên cột
---------------------------------------------------------------- */
function applyDbValidations_(sheet, headers, validations) {
  var MAX_DATA_ROWS = 1000;
  Object.keys(validations).forEach(function(colName) {
    var options = validations[colName];
    if (!options || options.length === 0) return;
    var colIdx = headers.indexOf(colName) + 1;
    if (colIdx <= 0) return; // cột không tìm thấy trong headers

    var range = sheet.getRange(2, colIdx, MAX_DATA_ROWS);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options, true)   // true = hiển thị dropdown arrow
      .setAllowInvalid(true)               // cho phép nhập ngoài list, không block
      .setHelpText('Chọn từ danh sách')
      .build();
    range.setDataValidation(rule);
  });
}

/* ----------------------------------------------------------------
   Định nghĩa toàn bộ sheets và cấu hình
---------------------------------------------------------------- */
function getDbSheetConfigs_() {

  // ── Dropdown constants ──
  var RAG          = ['🟢 Green', '🟡 Amber', '🔴 Red'];
  var WORKSTREAM   = ['01 Thương hiệu', '02 Kinh doanh', '03 Vận hành',
                      '04 Digital / AI', '05 Con người / Văn hóa', '06 Hệ thống / KPI'];
  var FACTORY      = ['Nhà máy Hà Nội', 'Nhà máy HCM', 'Nhà máy Đà Nẵng',
                      'Văn phòng HQ', 'Khác'];
  var REGION       = ['Miền Bắc', 'Miền Trung', 'Miền Nam', 'Toàn quốc'];
  var LEVEL        = ['Cao', 'Trung bình', 'Thấp'];
  var ISSUE_GROUP  = ['Vận hành', 'Kinh doanh', 'Con người',
                      'Hệ thống', 'Chất lượng', 'An toàn', 'Khác'];
  var SECURITY     = ['Công khai', 'Nội bộ', 'Bảo mật', 'Tối mật'];
  var SOURCE       = ['PMO', 'CEO Projects', 'TO Actions', 'Biên bản họp', 'Khác'];
  var ROLE         = ['PMO', 'CEO', 'GĐ Workstream', 'PIC', 'Hỗ trợ', 'Khác'];

  var PMO_STATUS   = ['Chưa bắt đầu', 'Đang làm', 'Hoàn thành', 'Vướng mắc', 'Quá hạn'];
  var PROJ_STATUS  = ['Chưa bắt đầu', 'Đang triển khai', 'Hoàn thành', 'Tạm dừng', 'Quá hạn'];
  var TOA_STATUS   = ['Chờ phân công', 'Chưa xử lý', 'Đang xử lý', 'Hoàn thành', 'Quá hạn'];
  var MM_STATUS    = ['Nháp', 'Chờ duyệt', 'Đã duyệt', 'Lưu trữ'];
  var DEC_STATUS   = ['Chờ CEO chốt', 'Đã chốt', 'Quá hạn', 'Hủy'];
  var RISK_STATUS  = ['Mới phát sinh', 'Đang theo dõi', 'Đang xử lý', 'Đã xử lý', 'Đóng lại'];
  var MM_SCOPE     = ['Toàn ban lãnh đạo', 'TO', 'Theo dự án', 'GSP NEXT 30', 'Khác'];
  var PRIORITY     = ['Cao', 'Trung bình', 'Thấp'];

  return [

    /* ── 1. PMO_90D_ACTIONS ── */
    {
      name: 'PMO_90D_ACTIONS',
      headers: [
        'Mã dòng', 'Workstream', 'Nhà máy', 'Khu vực', 'Đơn vị / Khối',
        'Hành động / Action', 'PIC', 'Owner', 'Ngày bắt đầu', 'Deadline',
        'Trạng thái', 'RAG', 'Nhóm vấn đề', 'Risk/SOS', 'CEO cần chốt',
        'PMO Comment', 'Ghi chú', 'Link sheet'
      ],
      validations: {
        'Workstream'  : WORKSTREAM,
        'Nhà máy'    : FACTORY,
        'Khu vực'    : REGION,
        'Trạng thái' : PMO_STATUS,
        'RAG'        : RAG,
        'Nhóm vấn đề': ISSUE_GROUP
      }
    },

    /* ── 2. CEO_PROJECTS ── */
    {
      name: 'CEO_PROJECTS',
      headers: [
        'Project ID', 'Tên dự án', 'Workstream', 'PIC Lead', 'Owner',
        'Ngày bắt đầu', 'Deadline', 'Trạng thái', 'RAG', '% Hoàn thành',
        'Milestone gần nhất', 'Milestone tiếp theo', 'Vướng mắc',
        'CEO cần chốt', 'Ghi chú', 'Link'
      ],
      validations: {
        'Workstream'  : WORKSTREAM,
        'Trạng thái' : PROJ_STATUS,
        'RAG'        : RAG
      }
    },

    /* ── 3. CEO_PROJECT_UPDATES ── */
    {
      name: 'CEO_PROJECT_UPDATES',
      headers: [
        'Update ID', 'Project ID', 'Ngày cập nhật', 'PIC cập nhật',
        '% Hoàn thành', 'Milestone đã đạt', 'Milestone tiếp theo',
        'Vướng mắc', 'CEO cần biết', 'RAG', 'Ghi chú'
      ],
      validations: {
        'RAG': RAG
      }
    },

    /* ── 4. TO_ACTIONS ── */
    {
      name: 'TO_ACTIONS',
      headers: [
        'Action ID', 'Nhà máy', 'Người ghi nhận', 'Ngày ghi nhận',
        'Ngày phát sinh', 'Meeting ID', 'Cuộc họp / Bối cảnh',
        'Issue ID liên quan', 'Khu vực',
        'Nội dung action', 'PIC thực hiện', 'Owner theo dõi', 'Deadline',
        'Loại kế hoạch', 'Trạng thái', 'Mức độ ưu tiên', 'RAG',
        'Kết quả / Output', 'CEO cần biết',
        'File nguồn', 'Link biên bản', 'Ghi chú'
      ],
      validations: {
        'Nhà máy'       : FACTORY,
        'Khu vực'       : REGION,
        'Loại kế hoạch' : ['Xử lý ngay (24h)', 'Xử lý khẩn (72h)', 'Kế hoạch dài hạn'],
        'Trạng thái'    : TOA_STATUS,
        'Mức độ ưu tiên': PRIORITY,
        'RAG'           : RAG
      }
    },

    /* ── 4b. TO_ISSUES ── */
    {
      name: 'TO_ISSUES',
      headers: [
        'Issue ID', 'Nhà máy', 'Người ghi nhận', 'Ngày ghi nhận', 'File nguồn',
        'Khu vực', 'Workstream', 'Loại vấn đề', 'Mô tả vấn đề',
        'Nguyên nhân', 'Ảnh hưởng', 'Mức độ', 'RAG',
        'Trạng thái', 'Action ID liên quan', 'Ghi chú'
      ],
      validations: {
        'Nhà máy'    : FACTORY,
        'Workstream' : WORKSTREAM,
        'Khu vực'   : REGION,
        'Loại vấn đề': ISSUE_GROUP,
        'Mức độ'    : LEVEL,
        'RAG'       : RAG,
        'Trạng thái': ['Mới', 'Đang xử lý', 'Đã xử lý', 'Đóng']
      }
    },

    /* ── 5. MEETING_MINUTES ── */
    {
      name: 'MEETING_MINUTES',
      headers: [
        'Meeting ID', 'Ngày họp', 'Chủ đề họp', 'Người chủ trì',
        'Thành phần tham dự', 'Phạm vi', 'Tóm tắt chính', 'Chỉ đạo CEO',
        'Quyết định đã chốt', 'Action phát sinh', 'Risk/SOS phát sinh',
        'CEO cần quyết tiếp', 'File biên bản', 'File ghi âm',
        'Người lập biên bản', 'Trạng thái', 'Ghi chú bảo mật'
      ],
      validations: {
        'Phạm vi'        : MM_SCOPE,
        'Trạng thái'     : MM_STATUS,
        'Ghi chú bảo mật': SECURITY
      }
    },

    /* ── 6. CEO_DECISIONS ── */
    {
      name: 'CEO_DECISIONS',
      headers: [
        'Decision ID', 'Nguồn', 'Source ID', 'Workstream',
        'Nội dung cần chốt', 'PIC đề xuất', 'Deadline',
        'Trạng thái', 'Quyết định CEO', 'Ngày chốt', 'Ghi chú'
      ],
      validations: {
        'Nguồn'      : SOURCE,
        'Workstream' : WORKSTREAM,
        'Trạng thái' : DEC_STATUS
      }
    },

    /* ── 7. RISK_SOS ── */
    {
      name: 'RISK_SOS',
      headers: [
        'Risk ID', 'Nguồn', 'Source ID', 'Workstream', 'Nhà máy', 'Khu vực',
        'PIC xử lý', 'Mô tả rủi ro', 'Nhóm vấn đề', 'Mức độ', 'RAG',
        'Deadline xử lý', 'Trạng thái', 'Biện pháp xử lý', 'CEO cần biết'
      ],
      validations: {
        'Nguồn'      : SOURCE,
        'Workstream' : WORKSTREAM,
        'Nhà máy'   : FACTORY,
        'Khu vực'   : REGION,
        'Nhóm vấn đề': ISSUE_GROUP,
        'Mức độ'    : LEVEL,
        'RAG'       : RAG,
        'Trạng thái': RISK_STATUS
      }
    },

    /* ── 8. MASTER_PIC ── */
    {
      name: 'MASTER_PIC',
      headers: [
        'Tên PIC', 'Workstream', 'Nhà máy', 'Email', 'Số điện thoại', 'Vai trò', 'Ghi chú'
      ],
      validations: {
        'Workstream': WORKSTREAM,
        'Nhà máy'  : FACTORY,
        'Vai trò'  : ROLE
      },
      seedData: []   // Người dùng tự nhập danh sách PIC
    },

    /* ── 9. MASTER_STATUS ── */
    {
      name: 'MASTER_STATUS',
      headers: ['Giá trị', 'Nhóm', 'Màu HEX', 'Thứ tự', 'Ghi chú'],
      validations: {
        'Nhóm': ['RAG', 'Trạng thái PMO', 'Trạng thái Dự án',
                 'Trạng thái TO Actions', 'Trạng thái Biên bản',
                 'Trạng thái Quyết định', 'Trạng thái Rủi ro']
      },
      seedData: [
        ['🟢 Green',        'RAG',                    '#16a34a', 1, ''],
        ['🟡 Amber',        'RAG',                    '#d97706', 2, ''],
        ['🔴 Red',          'RAG',                    '#dc2626', 3, ''],
        ['Chưa bắt đầu',   'Trạng thái PMO',         '#6b7280', 1, ''],
        ['Đang làm',       'Trạng thái PMO',         '#2563eb', 2, ''],
        ['Hoàn thành',     'Trạng thái PMO',         '#16a34a', 3, ''],
        ['Vướng mắc',      'Trạng thái PMO',         '#d97706', 4, ''],
        ['Quá hạn',        'Trạng thái PMO',         '#dc2626', 5, ''],
        ['Chưa bắt đầu',   'Trạng thái Dự án',       '#6b7280', 1, ''],
        ['Đang triển khai','Trạng thái Dự án',       '#2563eb', 2, ''],
        ['Hoàn thành',     'Trạng thái Dự án',       '#16a34a', 3, ''],
        ['Tạm dừng',       'Trạng thái Dự án',       '#d97706', 4, ''],
        ['Quá hạn',        'Trạng thái Dự án',       '#dc2626', 5, ''],
        ['Chờ CEO chốt',   'Trạng thái Quyết định',  '#d97706', 1, ''],
        ['Đã chốt',        'Trạng thái Quyết định',  '#16a34a', 2, ''],
        ['Quá hạn',        'Trạng thái Quyết định',  '#dc2626', 3, ''],
        ['Hủy',            'Trạng thái Quyết định',  '#6b7280', 4, ''],
        ['Mới phát sinh',  'Trạng thái Rủi ro',      '#dc2626', 1, ''],
        ['Đang theo dõi',  'Trạng thái Rủi ro',      '#d97706', 2, ''],
        ['Đang xử lý',     'Trạng thái Rủi ro',      '#2563eb', 3, ''],
        ['Đã xử lý',       'Trạng thái Rủi ro',      '#16a34a', 4, ''],
        ['Đóng lại',       'Trạng thái Rủi ro',      '#6b7280', 5, '']
      ]
    },

    /* ── 10. CONFIG ── */
    {
      name: 'CONFIG',
      headers: ['Key', 'Value', 'Mô tả'],
      validations: {},
      seedData: [
        ['APP_VERSION',        '1.0.0',               'Phiên bản ứng dụng'],
        ['APP_NAME',           'GSP NEXT 30',          'Tên ứng dụng'],
        ['SPREADSHEET_NAME',   'GSP NEXT 30 Database', 'Tên file Google Sheet'],
        ['TIMEZONE',           'Asia/Ho_Chi_Minh',     'Múi giờ hệ thống'],
        ['PMO_REVIEW_DAY',     'Monday',               'Ngày họp PMO hàng tuần'],
        ['CEO_REVIEW_DAY',     'Friday',               'Ngày họp CEO review'],
        ['RAG_RED_THRESHOLD',  '30',                   '% action quá hạn để tô đỏ workstream'],
        ['SETUP_DATE',         '',                     'Ngày chạy setupGspNext30Database()']
      ]
    }

  ]; // end return
}

/* ----------------------------------------------------------------
   testSetupDatabase — chạy từ GAS Editor để kiểm tra setup
---------------------------------------------------------------- */
function testSetupDatabase() {
  var result = setupGspNext30Database();
  Logger.log('Setup hoàn thành: ' + JSON.stringify(result.log));
}


/* ================================================================
   SETUP PIC INPUT SHEETS — setupPicInputSheets()
   Tạo 16 sheet nhập liệu cho từng phòng ban / PIC
   - Tạo sheet mới nếu chưa tồn tại
   - Thêm header chuẩn nếu sheet mới hoặc trống
   - Giữ nguyên dữ liệu nếu sheet đã có data
   - Freeze, định dạng header xanh đậm, dropdown, auto-resize
================================================================ */

function setupPicInputSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var PIC_SHEETS = [
    'PIC01_KinhDoanh',
    'PIC02_SalesTech',
    'PIC03_SCM',
    'PIC04_MuaHang',
    'PIC05_CongNghe',
    'PIC06_CoDien_QLTB',
    'PIC07_IT_ERP',
    'PIC08_QA',
    'PIC09_HR',
    'PIC10_TaiChinh',
    'PIC11_GS1',
    'PIC12_GS5',
    'PIC13_GS6',
    'PIC14_PMO',
    'PIC15_AI_Data',
    'PIC16_TO'
  ];

  var HEADERS = [
    'Mã dòng',
    'Workstream',
    'Module',
    'Nội dung hành động',
    'PIC chính',
    'PIC hỗ trợ',
    'Ngày bắt đầu',
    'Deadline',
    'Trạng thái',
    'RAG',
    '% Hoàn thành',
    'Kết quả / Output',
    'Vướng mắc / Rủi ro',
    'CEO cần chốt',
    'Ghi chú'
  ];

  var VALIDATIONS = {
    'Workstream': [
      '01 Thương hiệu',
      '02 Kinh doanh',
      '03 Vận hành',
      '04 Digital / AI',
      '05 Con người / Văn hóa',
      '06 Hệ thống / KPI'
    ],
    'Module': [
      'Kế hoạch & Ngân sách',
      'Vận hành thường kỳ',
      'Dự án / Cải tiến',
      'Báo cáo & KPI',
      'Nhân sự & Đào tạo',
      'Mua sắm & Đấu thầu',
      'Kiểm soát chất lượng',
      'Hệ thống & Công nghệ',
      'Khách hàng & Kinh doanh',
      'Khác'
    ],
    'Trạng thái': [
      'Chưa bắt đầu',
      'Đang làm',
      'Hoàn thành',
      'Vướng mắc',
      'Quá hạn'
    ],
    'RAG': [
      '🟢 Green',
      '🟡 Amber',
      '🔴 Red'
    ],
    'CEO cần chốt': [
      'Không cần',
      'Cần thông báo CEO',
      'Cần CEO chốt',
      'CEO đã chốt'
    ]
  };

  var log = [];
  var created = 0;
  var headersAdded = 0;
  var kept = 0;

  PIC_SHEETS.forEach(function(sheetName) {
    var existing = ss.getSheetByName(sheetName);
    var isNew    = !existing;
    var isEmpty  = !isNew && (
      existing.getLastRow() === 0 ||
      String(existing.getRange(1, 1).getValue()).trim() === ''
    );
    var sheet;

    if (isNew) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      created++;
      log.push('[TẠO MỚI] ' + sheetName);
    } else if (isEmpty) {
      sheet = existing;
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      headersAdded++;
      log.push('[THÊM HEADER] ' + sheetName + ' — sheet trống');
    } else {
      sheet = existing;
      kept++;
      log.push('[GIỮ NGUYÊN DATA] ' + sheetName + ' — cập nhật định dạng');
    }

    // Định dạng + validation (an toàn với sheet đã có data)
    applyDbHeaderFormat_(sheet, HEADERS.length);
    applyDbValidations_(sheet, HEADERS, VALIDATIONS);
    try { sheet.autoResizeColumns(1, HEADERS.length); } catch (e) {}
  });

  var summary = '=== PIC Input Sheets Setup ===\n' +
    'Tạo mới: ' + created + ' | Thêm header: ' + headersAdded + ' | Giữ nguyên: ' + kept + '\n\n' +
    log.join('\n');
  Logger.log(summary);

  try {
    ss.toast(
      'Tạo mới: ' + created + '  |  Thêm header: ' + headersAdded + '  |  Giữ nguyên data: ' + kept,
      'PIC Input Sheets Setup ✓', 8
    );
  } catch (e) {}

  return { created: created, headersAdded: headersAdded, kept: kept, log: log };
}

/* ----------------------------------------------------------------
   testSetupPicSheets — chạy từ GAS Editor để kiểm tra
---------------------------------------------------------------- */
function testSetupPicSheets() {
  var result = setupPicInputSheets();
  Logger.log('PIC Setup hoàn thành: ' + JSON.stringify(result));
}


/* ================================================================
   PIC PMO DATA — getPicPmoData()
   Đọc toàn bộ sheet có tiền tố "PIC", gom thành pmoAllActions,
   tổng hợp theo PIC / Workstream / Module / Status / RAG,
   tạo danh sách quá hạn, Risk/SOS, CEO cần chốt, Data Quality.
   Fallback 6 sheet cũ vẫn hoạt động qua getDashboardData().
================================================================ */

// Workstream mặc định theo tên sheet nếu cột Workstream bị bỏ trống
var PIC_SHEET_WORKSTREAM_ = {
  'PIC01_KinhDoanh':   '02 Kinh doanh',
  'PIC02_SalesTech':   '02 Kinh doanh',
  'PIC03_SCM':         '03 Vận hành',
  'PIC04_MuaHang':     '03 Vận hành',
  'PIC05_CongNghe':    '03 Vận hành',
  'PIC06_CoDien_QLTB': '03 Vận hành',
  'PIC07_IT_ERP':      '04 Digital / AI',
  'PIC08_QA':          '03 Vận hành',
  'PIC09_HR':          '05 Con người / Văn hóa',
  'PIC10_TaiChinh':    '06 Hệ thống / KPI',
  'PIC11_GS1':         '03 Vận hành',
  'PIC12_GS5':         '03 Vận hành',
  'PIC13_GS6':         '03 Vận hành',
  'PIC14_PMO':         '06 Hệ thống / KPI',
  'PIC15_AI_Data':     '04 Digital / AI',
  'PIC16_TO':          '06 Hệ thống / KPI'
};

/* ----------------------------------------------------------------
   Main endpoint
---------------------------------------------------------------- */
function getPicPmoData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Tìm tất cả sheet có tiền tố "PIC"
  var allSheets = ss.getSheets();
  var picSheets = allSheets.filter(function(s) {
    return s.getName().toUpperCase().indexOf('PIC') === 0;
  });

  var sheetsFound  = [];
  var sheetsEmpty  = [];
  var allActions   = [];

  // 2. Đọc và normalize từng sheet
  picSheets.forEach(function(sheet) {
    var sheetName = sheet.getName();
    var sheetUrl  = ss.getUrl() + '#gid=' + sheet.getSheetId();
    var rows      = readSheetSafe(sheetName);   // trả về [] nếu trống — không báo lỗi

    if (rows.length === 0) {
      sheetsEmpty.push(sheetName);
      return;
    }

    sheetsFound.push(sheetName);
    rows.forEach(function(row) {
      var normalized = normalizePicRow_(row, sheetName, sheetUrl, ss);
      if (normalized) allActions.push(normalized);
    });
  });

  // 3. Lọc các nhóm cơ bản
  var doneRows       = allActions.filter(function(r) { return r.Status === 'Done'; });
  var ipRows         = allActions.filter(function(r) { return r.Status === 'In Progress'; });
  var overdueRows    = allActions.filter(function(r) { return r.Status === 'Overdue' || r.IsOverdue; });
  var blockedRows    = allActions.filter(function(r) { return r.Status === 'Blocked'; });
  var notStartedRows = allActions.filter(function(r) { return r.Status === 'Not Started'; });

  var greenRows = allActions.filter(function(r) { return r.RAG === 'Green'; });
  var amberRows = allActions.filter(function(r) { return r.RAG === 'Amber'; });
  var redRows   = allActions.filter(function(r) { return r.RAG === 'Red'; });

  // Risk/SOS: cột Vướng mắc / Rủi ro có nội dung
  var riskRows = allActions.filter(function(r) { return hasText(r.Risk); });

  // CEO cần chốt: cột CEO cần chốt có nội dung và không phải "Không cần"
  var ceoRows = allActions.filter(function(r) { return isDecision(r.CEODecision); });

  // 4. Tổng hợp
  return JSON.parse(JSON.stringify({
    source:       'PIC_SHEETS',
    updatedAt:    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    sheetsFound:  sheetsFound,
    sheetsEmpty:  sheetsEmpty,

    summary: {
      total:             allActions.length,
      done:              doneRows.length,
      inProgress:        ipRows.length,
      overdue:           overdueRows.length,
      blocked:           blockedRows.length,
      notStarted:        notStartedRows.length,
      green:             greenRows.length,
      amber:             amberRows.length,
      red:               redRows.length,
      completionRate:    allActions.length > 0
                           ? Math.round(doneRows.length / allActions.length * 100) : 0,
      riskCount:         riskRows.length,
      ceoDecisionCount:  ceoRows.length
    },

    dataQuality:    buildPicDataQuality_(allActions),
    bySheet:        buildPicBySheet_(allActions, picSheets, ss),
    byWorkstream:   buildPicByGroup_(allActions, 'Workstream'),
    byModule:       buildPicByGroup_(allActions, 'Module'),
    byPIC:          buildPicByPIC_(allActions),
    byStatus: {
      'Done':        doneRows.length,
      'In Progress': ipRows.length,
      'Overdue':     overdueRows.length,
      'Blocked':     blockedRows.length,
      'Not Started': notStartedRows.length
    },
    byRAG: {
      'Green': greenRows.length,
      'Amber': amberRows.length,
      'Red':   redRows.length
    },

    overdueList:      overdueRows,
    riskList:         riskRows,
    ceoDecisionList:  ceoRows,
    allActions:       allActions
  }));
}

/* ----------------------------------------------------------------
   Normalize một row từ PIC sheet thành cấu trúc chuẩn
---------------------------------------------------------------- */
function normalizePicRow_(row, sheetName, sheetUrl, ss) {
  // Bỏ qua dòng hoàn toàn trống
  var rowId = String(row['Mã dòng'] || '').trim();
  if (!rowId) return null;

  var workstreamRaw = String(row['Workstream'] || '').trim();
  var workstream    = workstreamRaw || (PIC_SHEET_WORKSTREAM_[sheetName] || 'Chưa phân loại');
  var module        = String(row['Module'] || '').trim();
  var action        = String(row['Nội dung hành động'] || '').trim();
  var pic           = String(row['PIC chính'] || '').trim();
  var owner         = String(row['PIC hỗ trợ'] || '').trim();
  var startDate     = formatDate(row['Ngày bắt đầu']);
  var deadline      = formatDate(row['Deadline']);
  var status        = normalizeStatus(row['Trạng thái']);
  var rag           = normalizeRag(row['RAG']);
  var progress      = String(row['% Hoàn thành'] || '').trim();
  var output        = String(row['Kết quả / Output'] || '').trim();
  var risk          = String(row['Vướng mắc / Rủi ro'] || '').trim();
  var ceoDecision   = String(row['CEO cần chốt'] || '').trim();
  var note          = String(row['Ghi chú'] || '').trim();

  var isOverdue = isPastDeadline(deadline, status);

  var dqScore = (function() {
    var fields = [action, pic, deadline, status, rag, module];
    var filled = fields.filter(function(x) { return hasText(x); }).length;
    return Math.round(filled / fields.length * 100);
  })();

  var priorityScore = (function() {
    var s = 0;
    if (rag === 'Red')              s += 40;
    if (rag === 'Amber')            s += 15;
    if (status === 'Blocked')       s += 35;
    if (status === 'Overdue')       s += 35;
    if (isOverdue)                  s += 30;
    if (hasText(risk))              s += 25;
    if (isDecision(ceoDecision))    s += 25;
    if (!hasText(pic))              s += 8;
    if (!hasText(action))           s += 8;
    if (!hasText(deadline))         s += 10;
    return s;
  })();

  return {
    RowId:         rowId,
    SourceSheet:   sheetName,
    SourceUrl:     sheetUrl,
    Workstream:    workstream,
    Module:        module,
    Action:        action,
    PIC:           pic,
    Owner:         owner,
    StartDate:     startDate,
    Deadline:      deadline,
    Status:        status,
    RAG:           rag,
    Progress:      progress,
    Output:        output,
    Risk:          risk,
    CEODecision:   ceoDecision,
    Note:          note,
    IsOverdue:     isOverdue,
    DataQuality:   dqScore,
    PriorityScore: priorityScore
  };
}

/* ----------------------------------------------------------------
   Data Quality cho PIC sheets
---------------------------------------------------------------- */
function buildPicDataQuality_(rows) {
  var total = rows.length;
  if (total === 0) return { score: 0, total: 0, missingPIC: 0, missingAction: 0,
                             missingDeadline: 0, missingStatus: 0, missingRAG: 0, missingModule: 0 };

  return {
    score:          Math.round(rows.reduce(function(s, r) { return s + r.DataQuality; }, 0) / total),
    total:          total,
    missingPIC:     rows.filter(function(r) { return !hasText(r.PIC); }).length,
    missingAction:  rows.filter(function(r) { return !hasText(r.Action); }).length,
    missingDeadline:rows.filter(function(r) { return !hasText(r.Deadline); }).length,
    missingStatus:  rows.filter(function(r) { return r.Status === 'Not Started' && !hasText(r.Status); }).length,
    missingRAG:     rows.filter(function(r) { return !hasText(r.RAG) || r.RAG === 'Amber'; }).length,
    missingModule:  rows.filter(function(r) { return !hasText(r.Module); }).length
  };
}

/* ----------------------------------------------------------------
   Tổng hợp theo sheet (bySheet)
---------------------------------------------------------------- */
function buildPicBySheet_(rows, picSheets, ss) {
  var result = {};

  picSheets.forEach(function(sheet) {
    var name  = sheet.getName();
    var url   = ss.getUrl() + '#gid=' + sheet.getSheetId();
    var label = name.replace(/_/g, ' ');
    var data  = rows.filter(function(r) { return r.SourceSheet === name; });
    var total = data.length;
    var done  = data.filter(function(r) { return r.Status === 'Done'; }).length;
    var ov    = data.filter(function(r) { return r.Status === 'Overdue' || r.IsOverdue; }).length;
    var red   = data.filter(function(r) { return r.RAG === 'Red'; }).length;

    var rag = 'Green';
    if (red > 0 || ov > 0) rag = 'Red';
    else if (data.some(function(r) { return r.RAG === 'Amber'; })) rag = 'Amber';
    else if (total === 0) rag = 'Amber';

    result[name] = {
      sheetName:      name,
      label:          label,
      url:            url,
      total:          total,
      done:           done,
      overdue:        ov,
      blocked:        data.filter(function(r) { return r.Status === 'Blocked'; }).length,
      completionRate: total > 0 ? Math.round(done / total * 100) : 0,
      riskCount:      data.filter(function(r) { return hasText(r.Risk); }).length,
      ceoCount:       data.filter(function(r) { return isDecision(r.CEODecision); }).length,
      rag:            rag
    };
  });

  return result;
}

/* ----------------------------------------------------------------
   Tổng hợp theo một field (Workstream hoặc Module)
---------------------------------------------------------------- */
function buildPicByGroup_(rows, field) {
  var groups = {};

  rows.forEach(function(r) {
    var key = hasText(r[field]) ? r[field] : '(Chưa điền)';
    if (!groups[key]) {
      groups[key] = { label: key, total: 0, done: 0, overdue: 0, blocked: 0,
                      green: 0, amber: 0, red: 0, riskCount: 0, ceoCount: 0 };
    }
    var g = groups[key];
    g.total++;
    if (r.Status === 'Done')                     g.done++;
    if (r.Status === 'Overdue' || r.IsOverdue)   g.overdue++;
    if (r.Status === 'Blocked')                  g.blocked++;
    if (r.RAG === 'Green')                       g.green++;
    if (r.RAG === 'Amber')                       g.amber++;
    if (r.RAG === 'Red')                         g.red++;
    if (hasText(r.Risk))                         g.riskCount++;
    if (isDecision(r.CEODecision))               g.ceoCount++;
  });

  // Gắn completionRate và rag tổng thể cho mỗi nhóm
  Object.keys(groups).forEach(function(k) {
    var g = groups[k];
    g.completionRate = g.total > 0 ? Math.round(g.done / g.total * 100) : 0;
    g.rag = (g.red > 0 || g.overdue > 0) ? 'Red'
          : (g.amber > 0 || g.riskCount > 0) ? 'Amber'
          : 'Green';
  });

  return groups;
}

/* ----------------------------------------------------------------
   Tổng hợp theo PIC (trả về mảng để dễ sort)
---------------------------------------------------------------- */
function buildPicByPIC_(rows) {
  var map = {};

  rows.forEach(function(r) {
    var key = hasText(r.PIC) ? r.PIC : '(Chưa phân công)';
    if (!map[key]) {
      map[key] = { pic: key, total: 0, done: 0, overdue: 0, blocked: 0,
                   green: 0, amber: 0, red: 0, riskCount: 0, ceoCount: 0,
                   sheets: {} };
    }
    var p = map[key];
    p.total++;
    if (r.Status === 'Done')                    p.done++;
    if (r.Status === 'Overdue' || r.IsOverdue)  p.overdue++;
    if (r.Status === 'Blocked')                 p.blocked++;
    if (r.RAG === 'Green')                      p.green++;
    if (r.RAG === 'Amber')                      p.amber++;
    if (r.RAG === 'Red')                        p.red++;
    if (hasText(r.Risk))                        p.riskCount++;
    if (isDecision(r.CEODecision))              p.ceoCount++;
    p.sheets[r.SourceSheet] = (p.sheets[r.SourceSheet] || 0) + 1;
  });

  return Object.keys(map).map(function(k) {
    var p = map[k];
    p.completionRate = p.total > 0 ? Math.round(p.done / p.total * 100) : 0;
    p.rag = (p.red > 0 || p.overdue > 0) ? 'Red'
           : (p.amber > 0) ? 'Amber'
           : 'Green';
    return p;
  }).sort(function(a, b) {
    // Sort: nhiều action nhất lên đầu
    return b.total - a.total;
  });
}

/* ----------------------------------------------------------------
   testGetPicPmoData — chạy từ GAS Editor để kiểm tra
---------------------------------------------------------------- */
function testGetPicPmoData() {
  var result = getPicPmoData();
  Logger.log('Sheets found: ' + result.sheetsFound.join(', '));
  Logger.log('Sheets empty: ' + result.sheetsEmpty.join(', '));
  Logger.log('Total actions: ' + result.summary.total);
  Logger.log('Overdue: ' + result.summary.overdue);
  Logger.log('Risk/SOS: ' + result.summary.riskCount);
  Logger.log('CEO cần chốt: ' + result.summary.ceoDecisionCount);
  Logger.log('DQ Score: ' + result.dataQuality.score + '%');
}


/* ================================================================
   TO ACTION CENTER — getToActionCenterData()
   Đọc TO_ISSUES + TO_ACTIONS, tổng hợp theo RAG/Khu vực/Urgency
================================================================ */

function getToActionCenterData() {
  var now   = new Date();
  var in24h = new Date(now.getTime() + 24 * 3600 * 1000);
  var in72h = new Date(now.getTime() + 72 * 3600 * 1000);

  var issues  = readSheetSafe('TO_ISSUES').map(normalizeToIssue_);
  var actions = readSheetSafe('TO_ACTIONS').map(normalizeTOActionFull_);

  // ── Urgency buckets ──
  var urgent24h = [];
  var urgent72h = [];
  var inPlan    = [];
  var noOwner   = [];
  var ceoList   = [];

  actions.forEach(function(a) {
    var planType = (a.PlanType || '').toLowerCase();
    var dl       = a.Deadline ? parseDate(a.Deadline) : null;
    var active   = (a.Status !== 'Done');

    if (active) {
      var is24  = planType.includes('24') || planType.includes('ngay')  || (dl && dl <= in24h);
      var is72  = !is24 && (planType.includes('72') || planType.includes('khẩn') || planType.includes('khan') || (dl && dl > in24h && dl <= in72h));

      if (is24)       urgent24h.push(a);
      else if (is72)  urgent72h.push(a);
      else            inPlan.push(a);
    }

    if ((isNeedAssign_(a.PIC) || isNeedAssign_(a.Owner)) && active) noOwner.push(a);
    if (hasText(a.CEONote) && !isNoValue_(a.CEONote))    ceoList.push(a);
  });

  // ── Issues aggregation ──
  var issRAG    = { Green: 0, Amber: 0, Red: 0 };
  var areaMap   = {};
  var evidenceMap = {};

  issues.forEach(function(iss) {
    var r = iss.RAG || 'Amber';
    if (issRAG[r] !== undefined) issRAG[r]++; else issRAG['Amber']++;

    var area = iss.Area || 'Chưa phân loại';
    if (!areaMap[area]) areaMap[area] = { area: area, total: 0, green: 0, amber: 0, red: 0 };
    areaMap[area].total++;
    if (r === 'Green')      areaMap[area].green++;
    else if (r === 'Red')   areaMap[area].red++;
    else                    areaMap[area].amber++;

    if (hasText(iss.EvidenceLink)) {
      evidenceMap[iss.EvidenceLink] = { url: iss.EvidenceLink, name: iss.Reporter + (iss.ReportDate ? '  —  ' + iss.ReportDate : '') };
    }
  });

  actions.forEach(function(a) {
    if (hasText(a.EvidenceLink)) {
      evidenceMap[a.EvidenceLink] = { url: a.EvidenceLink, name: a.Source || a.Meeting || 'File báo cáo' };
    }
  });

  var done = actions.filter(function(a) { return a.Status === 'Done'; });
  var overdue = actions.filter(function(a) { return a.Status === 'Overdue' || a.IsOverdue; });

  // ── Unique factories (for frontend factory tab generation) ──
  var factorySet = {};
  issues.forEach(function(x)  { if (x.Factory) factorySet[x.Factory] = true; });
  actions.forEach(function(x) { if (x.Factory) factorySet[x.Factory] = true; });
  var factories = Object.keys(factorySet).sort();

  return JSON.parse(JSON.stringify({
    updatedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),

    factories: factories,

    issuesSummary: {
      total:   issues.length,
      byRAG:   issRAG,
      open:    issues.filter(function(i) { return i.Status !== 'Đã xử lý' && i.Status !== 'Đóng'; }).length,
      byStatus: buildStatusCounts_(issues, 'Status')
    },

    actionsSummary: {
      total:          actions.length,
      done:           done.length,
      inProgress:     actions.filter(function(a) { return a.Status === 'In Progress'; }).length,
      overdue:        overdue.length,
      noOwner:        noOwner.length,
      ceoNeeded:      ceoList.length,
      urgent24hCount: urgent24h.length,
      urgent72hCount: urgent72h.length,
      inPlanCount:    inPlan.length
    },

    byArea:        Object.keys(areaMap).map(function(k) { return areaMap[k]; }),
    evidenceFiles: Object.keys(evidenceMap).map(function(k) { return evidenceMap[k]; }),

    urgent24h:   urgent24h,
    urgent72h:   urgent72h,
    inPlan:      inPlan,
    noOwnerList: noOwner,
    ceoList:     ceoList,
    issues:      issues,
    actions:     actions
  }));
}

/* ── Count helper (reusable) ── */
function buildStatusCounts_(rows, field) {
  var counts = {};
  rows.forEach(function(r) {
    var k = String(r[field] || 'Không rõ');
    counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

function isNoValue_(v) {
  var s = String(v || '').toLowerCase().trim();
  return !s || s === 'không' || s === 'khong' || s === 'không cần' || s === 'no' || s === 'n/a';
}

function isNeedAssign_(v) {
  var s = String(v || '').trim().toLowerCase();
  return !s || s === 'cần phân công' || s === 'can phan cong'
    || s === 'chưa phân công' || s === 'chua phan cong'
    || s === 'tbd' || s === 'chưa có' || s === 'chua co';
}

/* ── Normalize TO_ISSUES row ── */
function normalizeToIssue_(r, i) {
  var rag    = normalizeRag(r['RAG'] || '');
  var status = String(r['Trạng thái'] || r['Status'] || 'Mới').trim();
  return {
    IssueID:      String(r['Issue ID']            || ('ISS-' + String(i + 1).padStart(4, '0'))),
    Factory:      String(r['Nhà máy']             || '').trim(),
    Reporter:     String(r['Người ghi nhận']      || r['Nguồn báo cáo'] || r['Người báo cáo'] || '').trim(),
    ReportDate:   formatDate(r['Ngày ghi nhận']   || r['Ngày báo cáo'] || ''),
    EvidenceLink: String(r['File nguồn']          || r['Evidence Link'] || '').trim(),
    Area:         String(r['Khu vực']             || r['Nhà máy'] || '').trim(),
    Workstream:   String(r['Workstream']           || '').trim(),
    IssueType:    String(r['Loại vấn đề']         || '').trim(),
    Description:  String(r['Mô tả vấn đề']        || r['Mô tả'] || '').trim(),
    RootCause:    String(r['Nguyên nhân']          || '').trim(),
    Impact:       String(r['Ảnh hưởng']           || '').trim(),
    Level:        String(r['Mức độ']              || '').trim(),
    RAG:          rag,
    Status:       status,
    ActionIDRef:  String(r['Action ID liên quan']  || '').trim(),
    Note:         String(r['Ghi chú']             || '').trim()
  };
}

/* ── Normalize TO_ACTIONS row (full version for TO Action Center) ── */
function normalizeTOActionFull_(r, i) {
  var rag      = normalizeRag(r['RAG'] || '');
  var status   = normalizeStatus(r['Trạng thái'] || r['Status'] || '');
  var deadline = formatDate(r['Deadline'] || '');
  return {
    RowId:        String(r['Action ID']                || ('TOA-' + String(i + 1).padStart(4, '0'))),
    ActionID:     String(r['Action ID']                || ('TOA-' + String(i + 1).padStart(4, '0'))),
    Factory:      String(r['Nhà máy']                 || '').trim(),
    Reporter:     String(r['Người ghi nhận']          || r['Nguồn báo cáo'] || '').trim(),
    ReportDate:   formatDate(r['Ngày ghi nhận']       || ''),
    Date:         formatDate(r['Ngày phát sinh']       || ''),
    MeetingID:    String(r['Meeting ID']               || '').trim(),
    IssueIDRef:   String(r['Issue ID liên quan']       || '').trim(),
    Area:         String(r['Khu vực']                 || '').trim(),
    Source:       String(r['Cuộc họp / Bối cảnh']     || r['Bối cảnh'] || '').trim(),
    Action:       String(r['Nội dung action']          || r['Action'] || '').trim(),
    PIC:          String(r['PIC thực hiện']            || r['PIC'] || '').trim(),
    Owner:        String(r['Owner theo dõi']           || r['Owner'] || '').trim(),
    Deadline:     deadline,
    PlanType:     String(r['Loại kế hoạch']           || r['Loại'] || '').trim(),
    Status:       status,
    Priority:     String(r['Mức độ ưu tiên']          || '').trim(),
    RAG:          rag,
    Output:       String(r['Kết quả / Output']         || r['Kết quả'] || '').trim(),
    CEONote:      String(r['CEO cần biết']             || '').trim(),
    EvidenceLink: String(r['File nguồn']              || r['Evidence Link'] || '').trim(),
    RowUrl:       String(r['Link biên bản']            || r['Link'] || '').trim(),
    Note:         String(r['Ghi chú']                 || '').trim(),
    IsOverdue:    isPastDeadline(deadline, status)
  };
}


/* ================================================================
   IMPORT TO REPORT → DATABASE
   importToReportToDatabase(sourceUrl, factory, reporter, reportDate, forceReimport)

   Mở file báo cáo TO (Google Sheets) và ghi vào TO_ISSUES + TO_ACTIONS.
   - Đọc sheet "01_Hien_trang"          → TO_ISSUES
   - Đọc sheet "02_Ke_hoach_hanh_dong"  → TO_ACTIONS
   - Ghi thêm: Nhà máy, Người ghi nhận, Ngày ghi nhận, File nguồn
   - Nếu sheet chưa tồn tại sẽ tự tạo với header chuẩn
   - forceReimport = true để import lại kể cả đã có

   Ví dụ:
     importToReportToDatabase(linkGS1, 'GS1', 'Ms Loan',             '10/06/2026', false)
     importToReportToDatabase(linkGS5, 'GS5', 'Ms Hương / Mr Toản', '10/06/2026', false)
================================================================ */

function importToReportToDatabase(sourceUrl, factory, reporter, reportDate, forceReimport) {
  if (!sourceUrl) return { success: false, error: 'Thiếu sourceUrl. Hãy truyền URL Google Sheets của file báo cáo TO.' };

  var sourceSS;
  try {
    sourceSS = SpreadsheetApp.openByUrl(sourceUrl);
  } catch (e) {
    return { success: false, error: 'Không mở được file: ' + e.message + '. Hãy chắc chắn file đã được share với tài khoản Script.' };
  }

  var date     = reportDate
    ? formatDate(parseDate(reportDate) || new Date())
    : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var nhaMay   = String(factory  || 'Không rõ').trim();
  var nguoiGhi = String(reporter || 'Không rõ').trim();
  var force    = forceReimport === true;

  var result = {
    success:        true,
    sourceFile:     sourceSS.getName(),
    sourceUrl:      sourceUrl,
    factory:        nhaMay,
    reporter:       nguoiGhi,
    reportDate:     date,
    importedAt:     Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    issuesImported: 0, issuesSkipped: 0,
    actionsImported: 0, actionsSkipped: 0,
    errors: []
  };

  try {
    var ir = importHienTrangToIssues_(sourceSS, nhaMay, nguoiGhi, date, sourceUrl, force);
    result.issuesImported = ir.imported;
    result.issuesSkipped  = ir.skipped;
    if (ir.note) result.errors.push('TO_ISSUES: ' + ir.note);
  } catch (e) {
    result.errors.push('TO_ISSUES lỗi: ' + e.message);
    Logger.log('importToReportToDatabase [TO_ISSUES]: ' + e.message);
  }

  try {
    var ar = importKHHDToActions_(sourceSS, nhaMay, nguoiGhi, date, sourceUrl, force);
    result.actionsImported = ar.imported;
    result.actionsSkipped  = ar.skipped;
    if (ar.note) result.errors.push('TO_ACTIONS: ' + ar.note);
  } catch (e) {
    result.errors.push('TO_ACTIONS lỗi: ' + e.message);
    Logger.log('importToReportToDatabase [TO_ACTIONS]: ' + e.message);
  }

  Logger.log('[importToReportToDatabase] ' + JSON.stringify(result));
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      nhaMay + '  |  TO_ISSUES: +' + result.issuesImported + '  |  TO_ACTIONS: +' + result.actionsImported,
      'Import TO Report ✓', 7
    );
  } catch (e) {}

  return result;
}

/* ── Import 01_Hien_trang → TO_ISSUES ── */
function importHienTrangToIssues_(sourceSS, factory, reporter, reportDate, sourceUrl, force) {
  var srcSheet = findSheetByPattern_(sourceSS, ['01_hien_trang', 'hien_trang', 'hiện trạng', 'hien trang', '01']);
  if (!srcSheet) return { imported: 0, skipped: 0, note: 'Không tìm thấy sheet 01_Hien_trang.' };

  var rows = safeReadSheetFull_(srcSheet);
  if (!rows.length) return { imported: 0, skipped: 0, note: 'Sheet 01_Hien_trang trống.' };

  var destSheet = ensureSheetExists_('TO_ISSUES');

  // Kiểm tra duplicate: cùng File nguồn + Nhà máy
  if (!force) {
    var existLinks = getExistingColValues_(destSheet, 'File nguồn');
    if (!existLinks.length) existLinks = getExistingColValues_(destSheet, 'Evidence Link');
    if (existLinks.indexOf(sourceUrl) >= 0) {
      return { imported: 0, skipped: rows.length, note: 'Đã import từ file này (nhà máy: ' + factory + '). Dùng forceReimport=true để import lại.' };
    }
  }

  var prefix   = 'ISS-' + factory + '-' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd') + '-';
  var imported = 0, skipped = 0;

  rows.forEach(function(row, idx) {
    var desc = String(getVal(row, ['Mô tả vấn đề', 'Vấn đề', 'Issue', 'Mô tả', 'Nội dung', 'Hiện trạng', 'Mô tả / Issue']) || '').trim();
    if (!desc) { skipped++; return; }

    var stt = String(getVal(row, ['STT', 'Số TT', 'No', 'TT', '#']) || (idx + 1));
    appendRowByHeaders_(destSheet, {
      'Issue ID'          : prefix + String(stt).padStart(3, '0'),
      'Nhà máy'          : factory,
      'Người ghi nhận'   : reporter,
      'Ngày ghi nhận'    : reportDate,
      'File nguồn'       : sourceUrl,
      'Khu vực'          : String(getVal(row, ['Khu vực', 'Phân xưởng', 'Area', 'Khu vực / Nhà máy']) || '').trim(),
      'Workstream'        : String(getVal(row, ['Workstream', 'Trục', 'Nhóm']) || '').trim(),
      'Loại vấn đề'      : String(getVal(row, ['Loại vấn đề', 'Loại', 'Hạng mục', 'Category', 'Nhóm vấn đề']) || '').trim(),
      'Mô tả vấn đề'     : desc,
      'Nguyên nhân'       : String(getVal(row, ['Nguyên nhân', 'Root cause', 'Nguyên nhân gốc rễ']) || '').trim(),
      'Ảnh hưởng'        : String(getVal(row, ['Ảnh hưởng', 'Impact', 'Tác động']) || '').trim(),
      'Mức độ'           : String(getVal(row, ['Mức độ', 'Severity', 'Độ ưu tiên']) || '').trim(),
      'RAG'               : String(getVal(row, ['RAG', 'Đèn', 'Đèn RAG', 'Màu']) || '').trim(),
      'Trạng thái'        : String(getVal(row, ['Trạng thái', 'Status', 'Tình trạng']) || 'Mới').trim(),
      'Action ID liên quan': '',
      'Ghi chú'           : String(getVal(row, ['Ghi chú', 'Note', 'Nhận xét', 'Comment']) || '').trim()
    });
    imported++;
  });

  SpreadsheetApp.flush();
  return { imported: imported, skipped: skipped };
}

/* ── Import 02_Ke_hoach_hanh_dong → TO_ACTIONS ── */
function importKHHDToActions_(sourceSS, factory, reporter, reportDate, sourceUrl, force) {
  var srcSheet = findSheetByPattern_(sourceSS, ['02_ke_hoach_hanh_dong', 'ke_hoach', 'kế hoạch hành động', 'hanh dong', '02']);
  if (!srcSheet) return { imported: 0, skipped: 0, note: 'Không tìm thấy sheet 02_Ke_hoach_hanh_dong.' };

  var rows = safeReadSheetFull_(srcSheet);
  if (!rows.length) return { imported: 0, skipped: 0, note: 'Sheet 02_Ke_hoach_hanh_dong trống.' };

  var destSheet = ensureSheetExists_('TO_ACTIONS');

  if (!force) {
    var existLinks = getExistingColValues_(destSheet, 'File nguồn');
    if (!existLinks.length) existLinks = getExistingColValues_(destSheet, 'Evidence Link');
    if (existLinks.indexOf(sourceUrl) >= 0) {
      return { imported: 0, skipped: rows.length, note: 'Đã import từ file này (nhà máy: ' + factory + '). Dùng forceReimport=true để import lại.' };
    }
  }

  var prefix   = 'TOA-' + factory + '-' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd') + '-';
  var imported = 0, skipped = 0;

  rows.forEach(function(row, idx) {
    var action = String(getVal(row, ['Nội dung action', 'Việc cần làm', 'Action', 'Hành động', 'Nội dung việc', 'Công việc']) || '').trim();
    if (!action) { skipped++; return; }

    var stt = String(getVal(row, ['STT', 'Số TT', 'No', 'TT', '#']) || (idx + 1));
    var dl  = getVal(row, ['Deadline', 'Hạn hoàn thành', 'Thời hạn', 'Hạn']);
    appendRowByHeaders_(destSheet, {
      'Action ID'         : prefix + String(stt).padStart(3, '0'),
      'Nhà máy'          : factory,
      'Người ghi nhận'   : reporter,
      'Ngày ghi nhận'    : reportDate,
      'Ngày phát sinh'   : reportDate,
      'Meeting ID'        : '',
      'Cuộc họp / Bối cảnh': 'Báo cáo TO — ' + reporter + ' (' + factory + ')',
      'Issue ID liên quan' : '',
      'Khu vực'           : String(getVal(row, ['Khu vực', 'Phân xưởng', 'Area', 'Nhà máy']) || '').trim(),
      'Nội dung action'   : action,
      'PIC thực hiện'     : String(getVal(row, ['PIC', 'PIC thực hiện', 'Người thực hiện', 'Responsible']) || '').trim(),
      'Owner theo dõi'    : String(getVal(row, ['Owner', 'Owner theo dõi', 'Người theo dõi', 'Quản lý']) || '').trim(),
      'Deadline'          : dl ? formatDate(dl) : '',
      'Loại kế hoạch'    : String(getVal(row, ['Loại kế hoạch', 'Loại', 'Ưu tiên xử lý', 'Priority type', '24h/72h', 'Phân loại']) || '').trim(),
      'Trạng thái'        : String(getVal(row, ['Trạng thái', 'Status', 'Tình trạng']) || 'Chưa xử lý').trim(),
      'Mức độ ưu tiên'   : String(getVal(row, ['Mức độ ưu tiên', 'Mức độ', 'Priority', 'Ưu tiên']) || '').trim(),
      'RAG'               : String(getVal(row, ['RAG', 'Đèn']) || '').trim(),
      'Kết quả / Output'  : String(getVal(row, ['Kết quả', 'Output', 'Kết quả / Output']) || '').trim(),
      'CEO cần biết'      : String(getVal(row, ['CEO cần biết', 'CEO', 'CEO note']) || '').trim(),
      'File nguồn'        : sourceUrl,
      'Link biên bản'     : '',
      'Ghi chú'           : String(getVal(row, ['Ghi chú', 'Note', 'Comment']) || '').trim()
    });
    imported++;
  });

  SpreadsheetApp.flush();
  return { imported: imported, skipped: skipped };
}

/* ── Helpers ── */

// Tự tạo sheet nếu chưa tồn tại, dùng config từ getDbSheetConfigs_()
function ensureSheetExists_(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  Logger.log('ensureSheetExists_: tạo mới ' + sheetName);
  var configs = getDbSheetConfigs_();
  var cfg     = null;
  for (var i = 0; i < configs.length; i++) {
    if (configs[i].name === sheetName) { cfg = configs[i]; break; }
  }

  sheet = ss.insertSheet(sheetName);
  if (cfg) {
    sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
    applyDbHeaderFormat_(sheet, cfg.headers.length);
    applyDbValidations_(sheet, cfg.headers, cfg.validations || {});
    try { sheet.autoResizeColumns(1, cfg.headers.length); } catch (e) {}
  }
  return sheet;
}

// Ghi một hàng vào sheet, ánh xạ theo tên cột header (không phụ thuộc vào vị trí cột)
function appendRowByHeaders_(sheet, dataObj) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var row = headers.map(function(h) { return dataObj.hasOwnProperty(h) ? dataObj[h] : ''; });
  sheet.appendRow(row);
}

// Đọc sheet với header dòng đầu có nội dung (flexible)
function safeReadSheetFull_(sheet) {
  try {
    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];
    var headerIdx = 0;
    for (var i = 0; i < Math.min(5, values.length); i++) {
      if (values[i].some(function(c) { return String(c || '').trim() !== ''; })) { headerIdx = i; break; }
    }
    var headers = values[headerIdx].map(function(h) { return String(h || '').trim(); });
    return values.slice(headerIdx + 1)
      .filter(function(row) { return row.some(function(c) { return String(c || '').trim() !== ''; }); })
      .map(function(row) {
        var obj = {};
        headers.forEach(function(h, i) { if (h) obj[h] = (row[i] !== null && row[i] !== undefined) ? row[i] : ''; });
        return obj;
      });
  } catch (e) {
    Logger.log('safeReadSheetFull_ error: ' + e.message);
    return [];
  }
}

// Tìm sheet theo danh sách tên (lowercase, partial match)
function findSheetByPattern_(ss, patterns) {
  var sheets = ss.getSheets();
  for (var i = 0; i < patterns.length; i++) {
    var pat = patterns[i].toLowerCase();
    for (var j = 0; j < sheets.length; j++) {
      var name = sheets[j].getName().toLowerCase();
      if (name === pat || name.replace(/\s/g, '_') === pat || name.includes(pat)) return sheets[j];
    }
  }
  return null;
}

// Lấy danh sách giá trị đã có trong một cột (tránh duplicate)
function getExistingColValues_(sheet, colName) {
  try {
    if (sheet.getLastRow() < 1) return [];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colIdx  = headers.map(function(h) { return String(h || '').trim(); }).indexOf(colName);
    if (colIdx < 0 || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, colIdx + 1, sheet.getLastRow() - 1, 1).getValues()
      .map(function(r) { return String(r[0] || '').trim(); })
      .filter(function(v) { return v; });
  } catch (e) { return []; }
}

/* ── Test functions ── */
function testGetToActionCenterData() {
  var d = getToActionCenterData();
  Logger.log('Issues: ' + d.issuesSummary.total + ' | Actions: ' + d.actionsSummary.total);
  Logger.log('Factories: ' + (d.factories || []).join(', '));
  Logger.log('24h: ' + d.actionsSummary.urgent24hCount + ' | 72h: ' + d.actionsSummary.urgent72hCount);
  Logger.log('No owner: ' + d.actionsSummary.noOwner + ' | CEO: ' + d.actionsSummary.ceoNeeded);
}

function testImportMsLoanGS1TOReport() {
  var result = importToReportToDatabase(
    'https://docs.google.com/spreadsheets/d/1QxYEXdNs1MQCoLtjo0h4ipNFRolk54D713_9hsszZQY/edit?gid=516354085#gid=516354085',
    'GS1',
    'Ms Loan',
    '10/06/2026',
    false
  );
  Logger.log(JSON.stringify(result));
  return result;
}

function testImportHuongToanGS5TOReport() {
  var result = importToReportToDatabase(
    'DAN_LINK_GS5_VAO_DAY',   // ← thay bằng URL file Google Sheets GS5
    'GS5',
    'Ms Hương / Mr Toản',
    '10/06/2026',
    false
  );
  Logger.log(JSON.stringify(result));
  return result;
}