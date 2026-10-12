function doGet(e) {
  var action   = e && e.parameter && e.parameter.action;
  var callback = e && e.parameter && e.parameter.callback;

  function jsonOut(data) {
    var json = JSON.stringify(data);
    var body = callback ? callback + '(' + json + ')' : json;
    var mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
    return ContentService.createTextOutput(body).setMimeType(mime);
  }

  /* ── KIẾN TRÚC MỚI: 8 sheet — 1 module = 1 sheet ── */
  if (action === 'getDashboardData')      return jsonOut(getNewDashboardData());
  if (action === 'getDashboardComments')  return jsonOut(getDashboardComments());
  if (action === 'getToActionCenterData') return jsonOut(getTOModuleData());
  if (action === 'getMeetingCenter')      return jsonOut(getBienBanHopModuleData());
  if (action === 'getRiskSOS')            return jsonOut(getRiskModuleData());
  if (action === 'getCEODecisions')       return jsonOut(getQuyetDinhModuleData());
  if (action === 'getCEOProjects')        return jsonOut(getDuAnCEOModuleData());
  if (action === 'getCEOProjectTracker')  return jsonOut(getDuAnCEOModuleData());
  if (action === 'getPicPmoData')         return jsonOut(getPMOModuleData());
  /* ── Compat routes — Meeting Hub V3 & legacy ── */
  if (action === 'getTOActions') {
    try { return jsonOut(getTOActionsData()); }
    catch(eTOA) { return jsonOut({ exists: false, records: [], total: 0, _error: eTOA.message }); }
  }
  if (action === 'getMeetingHub')        return jsonOut(getMeetingHubData());
  if (action === 'getMeetingHubV3')      return jsonOut(getMeetingHubData());
  if (action === 'getMeetingDetail')     return jsonOut(getMeetingDetail((e.parameter||{}).meetingId||''));
  if (action === 'getMeetingDetailV3')   return jsonOut(getMeetingDetailV3((e.parameter||{}).meetingId||''));
  if (action === 'getMeetingActions')    return jsonOut(getMeetingActions((e.parameter||{}).meetingId||''));
  if (action === 'getMeetingFiles')      return jsonOut(getMeetingFiles((e.parameter||{}).meetingId||''));
  if (action === 'createMeetingDriveFolder') return jsonOut(createMeetingDriveFolder((e.parameter||{}).meetingId||'',(e.parameter||{}).meetingTitle||''));
  if (action === 'createMinutesFromTemplate') return jsonOut(createMinutesFromTemplate((e.parameter||{}).meetingId||'',(e.parameter||{}).meetingTitle||''));
  if (action === 'getMeetingPersonWorkload') return jsonOut(getPersonWorkload((e.parameter||{}).email||''));
  if (action === 'getMeetingPersonDetail')   return jsonOut(getPersonDetail((e.parameter||{}).email||''));
  if (action === 'getMeetingTopicDetail')    return jsonOut(getTopicDetail((e.parameter||{}).topicId||''));
  if (action === 'testMeetingHubConnection') return jsonOut(testMeetingHubConnection());
  if (action === 'setupMeetingHubConfigExample') return jsonOut(setupMeetingHubConfigExample());
  if (action === 'getMeetingMinutes')   return jsonOut(getMeetingMinutesData());
  if (action === 'getDataQualityFull')  return jsonOut(getDataQualityFullData());
  if (action === 'getCommandCenterData') return jsonOut(getCommandCenterData());

  if (action === 'getLosData')           return jsonOut(getLosData());
  if (action === 'getLeadershipCard')    return jsonOut(getLeadershipCard((e.parameter||{}).leaderId||''));
  if (action === 'getLosScorecard')      return jsonOut(getLosScorecard((e.parameter||{}).leaderId||''));
  if (action === 'getLeaderMeData')      return jsonOut(getLeaderMeData((e.parameter||{}).leaderId||''));
  if (action === 'getLosTrend')          return jsonOut(getLosTrend((e.parameter||{}).leaderId||''));
  if (action === 'getLeaderGateReviews') return jsonOut(getLeaderGateReviews((e.parameter||{}).leaderId||''));
  if (action === 'getLeaderCoaching')    return jsonOut(getLeaderCoaching((e.parameter||{}).leaderId||''));
  if (action === 'testLosConnection')    return jsonOut(testLosConnection());
  if (action === 'setupLosDataStore')    return jsonOut(setupLosDataStore());
  if (action === 'getLosDataSourceStatus') return jsonOut(getLosDataSourceStatus());
  if (action === 'getMeetingsForLosLinking') return jsonOut(getMeetingsForLosLinking((e.parameter||{}).leaderId||'',(e.parameter||{}).leaderEmail||''));

  /* ── CLC routes ── */
  if (action === 'getClcDashboard')       return jsonOut(getClcDashboard());
  if (action === 'getClcDataSourceStatus') return jsonOut(getClcDataSourceStatus());

  if (action === 'ping') return jsonOut({ ok: true, ts: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss') });

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
    } else if (action === 'toWriteIssue') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      payload.loaiND = 'Vấn đề';
      result = writeTO_(ss, payload);
    } else if (action === 'toWriteAction') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      payload.loaiND = 'Action';
      result = writeTO_(ss, payload);
    } else if (action === 'toUpdateAction') {
      result = { success: false, error: 'Tính năng cập nhật bằng ID đã bị tắt. Vui lòng cập nhật trực tiếp trong Google Sheets "03_TO_COMMAND_CENTER".' };
    } else if (action === 'saveCEOComment') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      result = writeQuyetDinh_(ss, payload);
    } else if (action === 'uploadMeetingFile') {
      result = uploadMeetingFile(payload.meetingId, payload.folderId, payload.base64data, payload.fileName, payload.mimeType);
    } else if (action === 'renameMeetingFile') {
      result = renameMeetingFile(payload.fileId, payload.newName);
    } else if (action === 'moveMeetingFile') {
      result = moveMeetingFile(payload.fileId, payload.targetFolderId);
    } else if (action === 'trashMeetingFile') {
      result = trashMeetingFile(payload.fileId);
    } else if (action === 'attachEvidenceToAction') {
      result = attachEvidenceToAction(payload.actionId, payload.fileId);
    } else if (action === 'attachMinutesFileToMeeting') {
      result = attachMinutesFileToMeeting(payload.meetingId, payload.fileId);
    } else if (action === 'createMinutesFromTemplate') {
      result = createMinutesFromTemplate(payload.meetingId, payload.meetingTitle);
    } else if (action === 'importMeetingActions') {
      result = importMeetingActionsFromPayload(payload.rows || [], payload.mode || 'batch', payload.meetingId || '');
    } else if (action === 'createMeeting') {
      result = createMeeting(payload);
    } else if (action === 'updateAction') {
      result = updateAction(payload.actionId, payload);
    } else if (action === 'seedMeetingHubPeople') {
      result = seedMeetingHubPeopleFromArray(payload.people || []);
    } else if (action === 'migrateLegacyMeetingHubData') {
      result = migrateLegacyMeetingHubData(payload.dryRun !== false);
    } else if (action === 'setupMeetingHubDataStore') {
      result = setupMeetingHubDataStore();
    } else if (action === 'saveLosVerifiedScore') {
      result = saveLosVerifiedScore(payload);
    } else if (action === 'saveMeVerifiedScore') {
      result = saveMeVerifiedScore(payload);
    } else if (action === 'saveLosReview') {
      result = saveLosReview(payload);
    } else if (action === 'createLosCoachingAction') {
      result = createLosCoachingAction(payload);
    } else if (action === 'updateLosCoachingAction') {
      result = updateLosCoachingAction(payload.coachingId, payload);
    } else if (action === 'linkLosEvidence') {
      result = linkLosEvidence(payload);
    } else if (action === 'setupLosDataStore') {
      result = setupLosDataStore();
    } else if (action === 'importLosLeaders') {
      result = importLosLeaders(payload);
    } else if (action === 'importLosScores') {
      result = importLosScores(payload);
    } else if (action === 'saveLosMeetingLinks') {
      result = saveLosMeetingLinks(payload);
    } else if (action === 'saveClcCommitment') {
      result = saveClcCommitment(payload);
    } else if (action === 'confirmClcCommitment') {
      result = confirmClcCommitment(payload);
    } else if (action === 'saveClcWeeklyReview') {
      result = saveClcWeeklyReview(payload);
    } else if (action === 'setupClcDataStore') {
      result = setupClcDataStore();
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
    var rawTopics    = readSheetSafe('MEETING_TOPICS');
    var rawMeetActs  = readSheetSafe('MEETING_ACTIONS');
    var meetMinutesFull = rawMinutes.map(normalizeMeetingMinutesFull_);
    var ccData = {
      pmoActions:     rows,
      ceoProjects:    rawProjects.map(normalizeCEOProject_),
      toActions:      dashboardData.toActions || [],
      riskSos:        rawRisks.map(normalizeRiskSOS_),
      ceoDecisions:   rawDecisions.map(normalizeCEODecision_),
      meetingTopics:  rawTopics.map(normalizeMeetingTopic_),
      meetingMinutes: meetMinutesFull
    };
    dashboardData.commandCenterSummary = buildCommandCenterSummary(ccData);
    dashboardData.ceoProjects    = ccData.ceoProjects;
    dashboardData.meetingTopics  = ccData.meetingTopics;
    dashboardData.meetingMinutes = meetMinutesFull;
    dashboardData.meetingActions = rawMeetActs.map(normalizeMeetingAction_);
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
   THEO DÕI DỰ ÁN CEO — single-sheet hierarchical model
   Sheet: THEO_DOI_DU_AN_CEO (34 cột)
   Loại bản ghi: DỰ ÁN | HẠNG MỤC | CÔNG VIỆC | CẬP NHẬT | QUYẾT ĐỊNH
===================================================== */

function setupCEOProjectTrackerSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SHEET_NAME = 'THEO_DOI_DU_AN_CEO';
  var COLUMNS = [
    'Loại bản ghi', 'Mã bản ghi', 'Mã cấp trên', 'Mã dự án',
    'Tên dự án/Hạng mục/Công việc', 'Mô tả', 'Bối cảnh', 'Mục tiêu',
    'Phạm vi', 'Sản phẩm đầu ra', 'Người phụ trách', 'Người phối hợp',
    'Trục công việc', 'Giai đoạn triển khai', 'Mức độ ưu tiên',
    'Ngày bắt đầu', 'Thời hạn hoàn thành', 'Trạng thái',
    'Phần trăm hoàn thành', 'RAG', 'Công việc đã thực hiện',
    'Cập nhật mới nhất', 'Kết quả đạt được', 'Vướng mắc/Rủi ro',
    'Hướng xử lý', 'Kế hoạch tiếp theo', 'Nội dung CEO cần biết',
    'Nội dung cần CEO quyết', 'Các phương án đề xuất', 'Quyết định của CEO',
    'Ngày cập nhật', 'Người cập nhật', 'Link bằng chứng/File nguồn', 'Ghi chú',
    'Đề bài CEO', 'Việc đang triển khai', 'Tiến độ hiện tại'
  ];

  var sheet = ss.getSheetByName(SHEET_NAME);
  var isNew = !sheet;
  if (isNew) sheet = ss.insertSheet(SHEET_NAME);

  var existing = isNew ? [] : (sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : []);

  if (isNew) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  } else {
    COLUMNS.forEach(function(col) {
      if (existing.indexOf(col) === -1) {
        var nc = sheet.getLastColumn() + 1;
        sheet.getRange(1, nc).setValue(col);
        Logger.log('Đã thêm cột mới: ' + col);
      }
    });
  }

  var totalHdrCols = sheet.getLastColumn();
  sheet.getRange(1, 1, 1, totalHdrCols)
    .setBackground('#1a2e44').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  var dataEnd = function(col) { return sheet.getRange(2, col, 998, 1); };
  var rule    = function(list) {
    return SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true).setAllowInvalid(false).build();
  };
  dataEnd(1).setDataValidation(rule(['DỰ ÁN','HẠNG MỤC','CÔNG VIỆC','CẬP NHẬT','QUYẾT ĐỊNH']));
  dataEnd(14).setDataValidation(rule(['LÀM NGAY','THIẾT KẾ NGAY','THỬ NGHIỆM SAU','CHỜ XEM XÉT']));
  dataEnd(15).setDataValidation(rule(['Khẩn cấp','Cao','Trung bình','Thấp']));
  dataEnd(18).setDataValidation(rule(['Cần cập nhật','Chưa bắt đầu','Đang triển khai','Chờ phối hợp',
    'Chờ phê duyệt','Tạm dừng','Vướng mắc','Quá hạn','Hoàn thành','Đã hủy']));
  dataEnd(20).setDataValidation(rule(['Xanh','Vàng','Đỏ']));
  dataEnd(28).setDataValidation(rule(['Không','Cần thông báo','Cần quyết định','Đã quyết định']));

  var widths = {1:120,2:100,3:100,4:90,5:220,6:200,7:180,8:180,9:180,10:200,
               11:150,12:150,13:150,14:140,15:120,16:110,17:120,18:140,19:90,
               20:80,21:220,22:220,23:200,24:200,25:200,26:200,27:200,28:150,
               29:200,30:200,31:110,32:130,33:200,34:180,35:220,36:220,37:160};
  Object.keys(widths).forEach(function(c) {
    sheet.setColumnWidth(Number(c), widths[c]);
  });

  // Wrap text toàn bộ vùng dữ liệu
  var totalCols = sheet.getLastColumn();
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), totalCols).setWrap(true);

  // Auto-filter (xóa filter cũ nếu có, tạo lại)
  try {
    var existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
  } catch(eF) {}
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), sheet.getLastColumn()).createFilter();

  Logger.log('setupCEOProjectTrackerSheet: OK — "' + SHEET_NAME + '" ' + (isNew ? 'tạo mới' : 'cập nhật'));
  return { ok: true, sheetName: SHEET_NAME, isNew: isNew, columnsCount: sheet.getLastColumn() };
}

function testSetupCEOProjectTrackerSheet() {
  Logger.log('=== testSetupCEOProjectTrackerSheet ===');
  try {
    var result = setupCEOProjectTrackerSheet();
    Logger.log('Kết quả: ' + JSON.stringify(result));
    Logger.log('Sheet: ' + result.sheetName);
    Logger.log('Trạng thái: ' + (result.isNew ? 'Tạo mới' : 'Cập nhật sheet đã tồn tại'));
    Logger.log('Số cột: ' + result.columnsCount);
    Logger.log('--- OK ---');
  } catch(e) {
    Logger.log('LỖI: ' + e.message + '\n' + e.stack);
  }
}

function importCEOProjectSeedData(seedData) {
  if (!Array.isArray(seedData)) {
    return { added:0, updated:0, skipped:0, errors:1, message:'seedData phải là mảng' };
  }
  var SHEET_NAME = 'THEO_DOI_DU_AN_CEO';
  var COLUMNS = [
    'Loại bản ghi','Mã bản ghi','Mã cấp trên','Mã dự án',
    'Tên dự án/Hạng mục/Công việc','Mô tả','Bối cảnh','Mục tiêu',
    'Phạm vi','Sản phẩm đầu ra','Người phụ trách','Người phối hợp',
    'Trục công việc','Giai đoạn triển khai','Mức độ ưu tiên',
    'Ngày bắt đầu','Thời hạn hoàn thành','Trạng thái',
    'Phần trăm hoàn thành','RAG','Công việc đã thực hiện',
    'Cập nhật mới nhất','Kết quả đạt được','Vướng mắc/Rủi ro',
    'Hướng xử lý','Kế hoạch tiếp theo','Nội dung CEO cần biết',
    'Nội dung cần CEO quyết','Các phương án đề xuất','Quyết định của CEO',
    'Ngày cập nhật','Người cập nhật','Link bằng chứng/File nguồn','Ghi chú',
    'Đề bài CEO','Việc đang triển khai','Tiến độ hiện tại'
  ];
  // Profile columns: safe to fill-in for existing DỰ ÁN rows if currently empty
  var DU_AN_FILL_COLS = [
    'Tên dự án/Hạng mục/Công việc','Mô tả','Bối cảnh','Mục tiêu','Phạm vi',
    'Sản phẩm đầu ra','Người phụ trách','Người phối hợp',
    'Trục công việc','Giai đoạn triển khai','Mức độ ưu tiên',
    'Ghi chú','Đề bài CEO','Việc đang triển khai','Tiến độ hiện tại'
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setupCEOProjectTrackerSheet(); sheet = ss.getSheetByName(SHEET_NAME); }

  var lastRow = sheet.getLastRow();
  var sheetCols = sheet.getLastColumn();
  var allData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheetCols).getValues() : [];

  // Build lookup: Mã bản ghi → { rowIdx (0-based in allData), loai, sheetRow (1-based) }
  var existingMap = {};
  allData.forEach(function(row, idx) {
    var maId = String(row[1] || '').trim();
    if (maId) existingMap[maId] = { rowIdx: idx, loai: String(row[0] || '').trim(), sheetRow: idx + 2 };
  });

  var added = 0, updated = 0, skipped = 0, errors = 0, newRows = [];
  var cellUpdates = []; // { sheetRow, col (1-based), value }

  seedData.forEach(function(item, idx) {
    try {
      var maId = String(item['Mã bản ghi'] || '').trim();
      if (!maId) { errors++; return; }
      var itemLoai = String(item['Loại bản ghi'] || '').trim();

      if (existingMap[maId] !== undefined) {
        // Existing record: for DỰ ÁN rows only, fill empty profile columns
        if (itemLoai === 'DỰ ÁN' && existingMap[maId].loai === 'DỰ ÁN') {
          var existingRow = allData[existingMap[maId].rowIdx];
          var hasUpdate = false;
          DU_AN_FILL_COLS.forEach(function(col) {
            var ci = COLUMNS.indexOf(col);
            if (ci < 0 || ci >= sheetCols) return;
            var existingVal = String(existingRow[ci] || '').trim();
            var newVal = String(item[col] || '').trim();
            if (!existingVal && newVal) {
              cellUpdates.push({ sheetRow: existingMap[maId].sheetRow, col: ci + 1, value: newVal });
              existingRow[ci] = newVal; // update in-memory to avoid double scheduling
              hasUpdate = true;
            }
          });
          if (hasUpdate) updated++; else skipped++;
        } else {
          skipped++;
        }
        return;
      }

      // New record — append
      var row = COLUMNS.map(function(col) {
        var v = item[col];
        return (v === undefined || v === null) ? '' : String(v);
      });
      newRows.push(row);
      existingMap[maId] = { rowIdx: allData.length + newRows.length - 1, loai: itemLoai, sheetRow: 0 };
      added++;
    } catch(e) {
      Logger.log('importCEOProjectSeedData idx=' + idx + ': ' + e.message);
      errors++;
    }
  });

  // Apply cell-level updates for existing DỰ ÁN profile fields
  cellUpdates.forEach(function(u) {
    sheet.getRange(u.sheetRow, u.col).setValue(u.value);
  });

  // Append new rows in one batch
  if (newRows.length > 0) {
    var colCount = Math.max(COLUMNS.length, sheetCols);
    var padded = newRows.map(function(r) {
      while (r.length < colCount) r.push('');
      return r.slice(0, colCount);
    });
    var appendAt = Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(appendAt, 1, padded.length, colCount).setValues(padded);
  }

  var msg = 'Thêm mới: ' + added + ' | Cập nhật hồ sơ DỰ ÁN: ' + updated + ' | Bỏ qua trùng: ' + skipped + ' | Lỗi: ' + errors;
  Logger.log('importCEOProjectSeedData: ' + msg);
  return { added: added, updated: updated, skipped: skipped, errors: errors, message: msg };
}

function testImportCEOProjectSeedData() {
  Logger.log('=== testImportCEOProjectSeedData ===');

  var seedData = [
    // ========== 8 DỰ ÁN ==========
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P1',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P1',
      'Tên dự án/Hạng mục/Công việc': 'GMS – Hệ thống quản lý GSP',
      'Mô tả':           'Xây dựng GMS V1.0 — hệ điều hành quản trị tổng thể của GSP, gồm 12 quy tắc vận hành, 8 phụ lục thực thi, ma trận quyết định, chuỗi KPI, yêu cầu dashboard và yêu cầu logic APS.',
      'Sản phẩm đầu ra': 'GMS V1.0: 12 quy tắc vận hành; 8 phụ lục; Ma trận RACI và Ma trận quyết định; Chuỗi KPI và chính sách trả lương hiệu suất; Checklist Gemba; Quy trình SOS; Yêu cầu dashboard; Yêu cầu logic APS.',
      'Người phụ trách': 'Thành + Sơn',
      'Người phối hợp':  'Thủy, Toản, Lâm',
      'Trục công việc':  '06 Hệ thống / KPI / PMO',
      'Giai đoạn triển khai': 'THIẾT KẾ NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'Sổ tay vận hành GSP; Hệ thống vận hành GSP; 12 quy tắc vận hành; 8 phụ lục GMS; RACI; Ma trận quyết định; Chuỗi KPI; Quy trình leo thang SOS; Checklist Gemba; Trả lương theo hiệu suất và trách nhiệm; Không có người chịu trách nhiệm = Không hành động; Tổ chức phẳng.',
      'Ghi chú':         'Thiết kế trong 60 ngày, làm hệ điều hành mẹ. Trục triển khai: BG + BO. Cơ chế điều phối: PMO + Claude. Thứ tự ưu tiên CEO: 5/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P2',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P2',
      'Tên dự án/Hạng mục/Công việc': 'Kỷ luật nhà máy và nền tảng kho hàng',
      'Mô tả':           'Thiết lập kỷ luật vận hành tại nhà máy — chiến dịch 30 ngày đột phá + hệ thống kho hàng chuẩn hóa: kiểm soát BTP/TP/NVL, quản lý mẫu và phiên bản, tiêu chuẩn nơi làm việc, quản lý trực quan.',
      'Sản phẩm đầu ra': 'Báo cáo kết quả chiến dịch 30 ngày; Sổ tay tiêu chuẩn kho; Quy trình kiểm soát BTP/TP/NVL; Hệ thống quản lý mẫu và phiên bản; Bộ tiêu chuẩn nơi làm việc và quản lý trực quan.',
      'Người phụ trách': 'Thành + Thủy',
      'Người phối hợp':  'Toản, Lâm, quản lý nhà máy',
      'Trục công việc':  '03 Vận hành và Chất lượng',
      'Giai đoạn triển khai': 'LÀM NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      '30 ngày tạo ra trật tự; 30 ngày dẹp hết hỗn loạn; Chuẩn hóa kho hàng; Kiểm soát BTP; Quản lý phiên bản; Không có phiên bản = Không sản xuất; Kho không trật tự = Không nhập thêm; Tiêu chuẩn nơi làm việc.',
      'Ghi chú':         'Cần kết quả rõ trong 30 ngày đầu. Trục triển khai: BO + QA. Cơ chế điều phối: PMO + Quản lý nhà máy. Thứ tự ưu tiên CEO: 2/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P3',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P3',
      'Tên dự án/Hạng mục/Công việc': 'Nền tảng dữ liệu và nguồn dữ liệu duy nhất',
      'Mô tả':           'Xây dựng SSOT (Single Source of Truth) — danh mục chủ sở hữu dữ liệu, quy tắc khóa dữ liệu, nhóm quản trị, quy trình kiểm toán, đánh giá mức độ sẵn sàng dữ liệu cho AI.',
      'Sản phẩm đầu ra': 'Danh mục chủ sở hữu dữ liệu; Quy tắc SSOT; Chính sách khóa dữ liệu; Quy trình kiểm toán dữ liệu; Điều lệ nhóm quản trị dữ liệu; Báo cáo mức độ sẵn sàng dữ liệu.',
      'Người phụ trách': 'Thành',
      'Người phối hợp':  'IT/ERP, tất cả chủ sở hữu dữ liệu',
      'Trục công việc':  '06 Hệ thống / KPI / PMO',
      'Giai đoạn triển khai': 'LÀM NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'Một nguồn dữ liệu duy nhất cho mọi quyết định; Không có chủ sở hữu = Không có dữ liệu sạch; SSOT là nền tảng để AI hoạt động; Dữ liệu sai = quyết định sai.',
      'Ghi chú':         'Tiên quyết cho P7 (AI). Trục triển khai: IT/ERP + tất cả bộ phận. Cơ chế điều phối: PMO + Claude. Thứ tự ưu tiên CEO: 3/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P4',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P4',
      'Tên dự án/Hạng mục/Công việc': 'Lập kế hoạch, quản lý chuỗi cung ứng và MRP',
      'Mô tả':           'Xây dựng hệ thống lập kế hoạch sản xuất và quản lý chuỗi cung ứng: bản đồ chuỗi giá trị, SLA nội bộ, tháp lập kế hoạch S&OP, thiết kế tổ chức SCM, sổ tay MRP V1.0, chính sách tồn kho, yêu cầu logic APS.',
      'Sản phẩm đầu ra': 'Bản đồ chuỗi giá trị hiện tại; Sổ tay SLA nội bộ; Tháp lập kế hoạch S&OP; Thiết kế tổ chức SCM; Sổ tay MRP V1.0; Chính sách tồn kho; Tài liệu yêu cầu logic APS.',
      'Người phụ trách': 'Ly + Thành',
      'Người phối hợp':  'IT/ERP, nhà máy, mua hàng',
      'Trục công việc':  '03 Vận hành và Chất lượng',
      'Giai đoạn triển khai': 'THIẾT KẾ NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'MRP là bộ não của sản xuất; Không có MRP = ước đoán; SLA nội bộ là hợp đồng giữa các bộ phận; Tồn kho = tiền mặt bị khóa; APS thay thế Excel lập kế hoạch.',
      'Ghi chú':         'Cần SLA và dữ liệu nền sẵn sàng trước. Trục triển khai: SCM + BO + IT. Cơ chế điều phối: PMO + CLO. Thứ tự ưu tiên CEO: 4/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P5',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P5',
      'Tên dự án/Hạng mục/Công việc': 'Hệ điều hành chất lượng',
      'Mô tả':           'Xây dựng hệ thống quản lý chất lượng toàn diện: mô hình vận hành QMS, quy trình NCR/CAR/8D, CoQ/COPQ/FPY, quản lý khiếu nại, quản lý mẫu và tiêu chuẩn, lộ trình tích hợp AI.',
      'Sản phẩm đầu ra': 'Mô hình vận hành QMS; Quy trình NCR/CAR/8D; Bảng theo dõi CoQ/COPQ/FPY; Quy trình quản lý khiếu nại và SLA; Hệ thống quản lý mẫu và tiêu chuẩn; Lộ trình sẵn sàng AI cho chất lượng.',
      'Người phụ trách': 'Giang + Đức',
      'Người phối hợp':  'QA, nhà máy, Thành',
      'Trục công việc':  '03 Vận hành và Chất lượng',
      'Giai đoạn triển khai': 'LÀM NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'Chất lượng là trách nhiệm của tất cả; FPY = tỷ lệ đúng ngay từ đầu; Khiếu nại = tín hiệu hệ thống; 8D là công cụ giải quyết vấn đề gốc rễ; CoQ đo chi phí của chất lượng kém.',
      'Ghi chú':         'Ưu tiên xử lý lỗi nghiêm trọng trước, thiết kế hệ thống song song. Trục triển khai: QA + BO. Cơ chế điều phối: PMO + QA Lead. Thứ tự ưu tiên CEO: 7/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P6',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P6',
      'Tên dự án/Hạng mục/Công việc': 'Hệ điều hành thương mại',
      'Mô tả':           'Xây dựng hệ điều hành thương mại: quản trị Sales OS, Sales Story và Playbook, Sales Kit, CRM SSOT, War Room khách hàng chiến lược, quản trị giá và đấu thầu, phát triển kinh doanh mới.',
      'Sản phẩm đầu ra': 'Tài liệu quản trị Sales OS; Sales Story và Playbook chuẩn; Sales Kit đầy đủ; CRM là nguồn dữ liệu duy nhất; Quy trình War Room; Chính sách giá và quy trình đấu thầu; Kế hoạch kinh doanh mới.',
      'Người phụ trách': 'Ly + Quý + Nghĩa',
      'Người phối hợp':  'Marketing, IT/CRM, Thành',
      'Trục công việc':  '02 Sales / Hệ điều hành thương mại',
      'Giai đoạn triển khai': 'LÀM NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'Sales OS là hệ điều hành thương mại; CRM là não bộ bán hàng; War Room = phòng thủ khách hàng chiến lược; Giá = quyết định chiến lược không phải tùy hứng; Sales Story = câu chuyện bán hàng chuẩn.',
      'Ghi chú':         'Trục triển khai: Sales + Marketing + IT. Cơ chế điều phối: PMO + CCO. Thứ tự ưu tiên CEO: 6/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P7',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P7',
      'Tên dự án/Hạng mục/Công việc': 'Ứng dụng công nghệ số và trí tuệ nhân tạo',
      'Mô tả':           'Xây dựng nền tảng AI/số hóa: quản trị AI, Claude PMO, nền tảng dữ liệu AI, trung tâm điều hành, trợ lý AI nghiệp vụ, cơ sở tri thức số. Phụ thuộc vào P3 (SSOT).',
      'Sản phẩm đầu ra': 'Chính sách quản trị AI; Claude PMO vận hành; Nền tảng dữ liệu AI (phụ thuộc P3); Trung tâm điều hành vận hành; Các trợ lý AI nghiệp vụ; Cơ sở tri thức số và cộng tác số.',
      'Người phụ trách': 'Sơn PMO + Thành + IT/Digital/EPS',
      'Người phối hợp':  'Tất cả bộ phận (phụ thuộc P3)',
      'Trục công việc':  '04 Digital & AI',
      'Giai đoạn triển khai': 'LÀM NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'AI không thay con người — AI làm việc thay phần việc không cần óc; Claude PMO là não bộ quản lý dự án; Dữ liệu sạch = AI thông minh; Trung tâm điều hành = tầm nhìn thời gian thực; Tri thức số = tổ chức không phụ thuộc vào người.',
      'Ghi chú':         'LÀM NGAY với Claude PMO; APS/WMS/MES/AI chuyên sâu thử nghiệm sau. Phụ thuộc P3. Trục triển khai: IT/Digital + tất cả bộ phận. Cơ chế điều phối: PMO + CTO/EPS. Thứ tự ưu tiên CEO: 8/8.'
    },
    {
      'Loại bản ghi':                 'DỰ ÁN',
      'Mã bản ghi':                   'P8',
      'Mã cấp trên':                  '',
      'Mã dự án':                     'P8',
      'Tên dự án/Hạng mục/Công việc': 'Quản trị PMO và kiểm soát tiếp nhận',
      'Mô tả':           'Thiết lập văn phòng quản trị dự án (PMO): cổng tiếp nhận, nhật ký yêu cầu CEO, Master Tracker/RAG/SOS, nhật ký quyết định và rủi ro, báo cáo CEO hằng tuần, kiểm soát phụ thuộc con người.',
      'Sản phẩm đầu ra': 'Cổng tiếp nhận PMO vận hành; Nhật ký yêu cầu CEO; Master Tracker với RAG và SOS; Nhật ký quyết định và rủi ro; Báo cáo CEO hằng tuần; Bản đồ phụ thuộc con người.',
      'Người phụ trách': 'Sơn PMO',
      'Người phối hợp':  'Thành, CEO, tất cả quản lý dự án',
      'Trục công việc':  '06 Hệ thống / KPI / PMO',
      'Giai đoạn triển khai': 'LÀM NGAY',
      'Trạng thái':      'Cần cập nhật',
      'Đề bài CEO':      'PMO là hệ thần kinh trung ương; Mỗi yêu cầu CEO phải có ticket; RAG = tín hiệu sớm của vấn đề; Không có quyết định = dự án chết; Báo cáo hằng tuần = nhịp tim của dự án; Con người là điểm phụ thuộc nguy hiểm nhất.',
      'Ghi chú':         'Ưu tiên số 1 — PMO phải vận hành trước khi các dự án khác triển khai. Trục triển khai: PMO + tất cả leads. Cơ chế điều phối: Claude PMO. Thứ tự ưu tiên CEO: 1/8.'
    },

    // ========== HẠNG MỤC P1 (5 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P1.1', 'Mã cấp trên': 'P1', 'Mã dự án': 'P1', 'Tên dự án/Hạng mục/Công việc': 'Bộ quy tắc GMS – 12 quy tắc vận hành',   'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P1.2', 'Mã cấp trên': 'P1', 'Mã dự án': 'P1', 'Tên dự án/Hạng mục/Công việc': 'Bộ 8 phụ lục GMS',                          'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P1.3', 'Mã cấp trên': 'P1', 'Mã dự án': 'P1', 'Tên dự án/Hạng mục/Công việc': 'Ma trận RACI và Ma trận quyết định',        'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P1.4', 'Mã cấp trên': 'P1', 'Mã dự án': 'P1', 'Tên dự án/Hạng mục/Công việc': 'Chuỗi KPI và trả lương theo hiệu suất',     'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P1.5', 'Mã cấp trên': 'P1', 'Mã dự án': 'P1', 'Tên dự án/Hạng mục/Công việc': 'Sổ tay vận hành / Gemba / Hệ thống SOS',   'Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P2 (5 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P2.1', 'Mã cấp trên': 'P2', 'Mã dự án': 'P2', 'Tên dự án/Hạng mục/Công việc': 'Chiến dịch 30 ngày đột phá nhà máy',        'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P2.2', 'Mã cấp trên': 'P2', 'Mã dự án': 'P2', 'Tên dự án/Hạng mục/Công việc': 'Chuẩn hóa kho hàng',                        'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P2.3', 'Mã cấp trên': 'P2', 'Mã dự án': 'P2', 'Tên dự án/Hạng mục/Công việc': 'Kiểm soát BTP / TP / NVL',                  'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P2.4', 'Mã cấp trên': 'P2', 'Mã dự án': 'P2', 'Tên dự án/Hạng mục/Công việc': 'Quản lý mẫu và phiên bản',                  'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P2.5', 'Mã cấp trên': 'P2', 'Mã dự án': 'P2', 'Tên dự án/Hạng mục/Công việc': 'Quản lý trực quan và tiêu chuẩn nơi làm việc','Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P3 (6 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P3.1', 'Mã cấp trên': 'P3', 'Mã dự án': 'P3', 'Tên dự án/Hạng mục/Công việc': 'Danh mục chủ sở hữu dữ liệu',               'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P3.2', 'Mã cấp trên': 'P3', 'Mã dự án': 'P3', 'Tên dự án/Hạng mục/Công việc': 'Quy tắc SSOT',                               'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P3.3', 'Mã cấp trên': 'P3', 'Mã dự án': 'P3', 'Tên dự án/Hạng mục/Công việc': 'Thời gian khóa dữ liệu',                     'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P3.4', 'Mã cấp trên': 'P3', 'Mã dự án': 'P3', 'Tên dự án/Hạng mục/Công việc': 'Quy trình kiểm toán dữ liệu',                'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P3.5', 'Mã cấp trên': 'P3', 'Mã dự án': 'P3', 'Tên dự án/Hạng mục/Công việc': 'Nhóm quản trị dữ liệu',                      'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P3.6', 'Mã cấp trên': 'P3', 'Mã dự án': 'P3', 'Tên dự án/Hạng mục/Công việc': 'Mức độ sẵn sàng dữ liệu',                   'Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P4 (7 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.1', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Bản đồ chuỗi giá trị',                       'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.2', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Sổ tay SLA nội bộ',                          'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.3', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Tháp lập kế hoạch (S&OP)',                   'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.4', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Thiết kế tổ chức SCM',                       'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.5', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Sổ tay MRP V1.0',                            'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.6', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Chính sách tồn kho',                         'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P4.7', 'Mã cấp trên': 'P4', 'Mã dự án': 'P4', 'Tên dự án/Hạng mục/Công việc': 'Yêu cầu logic APS',                          'Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P5 (6 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P5.1', 'Mã cấp trên': 'P5', 'Mã dự án': 'P5', 'Tên dự án/Hạng mục/Công việc': 'Mô hình vận hành QMS',                       'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P5.2', 'Mã cấp trên': 'P5', 'Mã dự án': 'P5', 'Tên dự án/Hạng mục/Công việc': 'NCR / CAR / 8D',                             'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P5.3', 'Mã cấp trên': 'P5', 'Mã dự án': 'P5', 'Tên dự án/Hạng mục/Công việc': 'CoQ / COPQ / FPY',                           'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P5.4', 'Mã cấp trên': 'P5', 'Mã dự án': 'P5', 'Tên dự án/Hạng mục/Công việc': 'Khiếu nại và SLA',                           'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P5.5', 'Mã cấp trên': 'P5', 'Mã dự án': 'P5', 'Tên dự án/Hạng mục/Công việc': 'Quản lý mẫu và tiêu chuẩn',                  'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P5.6', 'Mã cấp trên': 'P5', 'Mã dự án': 'P5', 'Tên dự án/Hạng mục/Công việc': 'Mức sẵn sàng AI cho chất lượng',             'Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P6 (7 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.1', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'Quản trị Hệ điều hành thương mại',           'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.2', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'Sales Story và Playbook',                    'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.3', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'Sales Kit',                                  'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.4', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'CRM – nguồn dữ liệu duy nhất',              'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.5', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'War Room khách hàng chiến lược',             'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.6', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'Quản trị giá và đấu thầu',                   'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P6.7', 'Mã cấp trên': 'P6', 'Mã dự án': 'P6', 'Tên dự án/Hạng mục/Công việc': 'Kinh doanh mới và đầu vào Nhà máy mới',     'Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P7 (6 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P7.1', 'Mã cấp trên': 'P7', 'Mã dự án': 'P7', 'Tên dự án/Hạng mục/Công việc': 'Quản trị AI và mô hình vận hành',            'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P7.2', 'Mã cấp trên': 'P7', 'Mã dự án': 'P7', 'Tên dự án/Hạng mục/Công việc': 'AI PMO / Claude PMO',                        'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P7.3', 'Mã cấp trên': 'P7', 'Mã dự án': 'P7', 'Tên dự án/Hạng mục/Công việc': 'Nền tảng dữ liệu và AI',                     'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P7.4', 'Mã cấp trên': 'P7', 'Mã dự án': 'P7', 'Tên dự án/Hạng mục/Công việc': 'Trung tâm điều hành vận hành',               'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P7.5', 'Mã cấp trên': 'P7', 'Mã dự án': 'P7', 'Tên dự án/Hạng mục/Công việc': 'Trợ lý AI nghiệp vụ',                        'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P7.6', 'Mã cấp trên': 'P7', 'Mã dự án': 'P7', 'Tên dự án/Hạng mục/Công việc': 'Cộng tác số và cơ sở tri thức',              'Trạng thái': 'Cần cập nhật' },

    // ========== HẠNG MỤC P8 (6 hạng mục) ==========
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P8.1', 'Mã cấp trên': 'P8', 'Mã dự án': 'P8', 'Tên dự án/Hạng mục/Công việc': 'Cổng tiếp nhận PMO',                         'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P8.2', 'Mã cấp trên': 'P8', 'Mã dự án': 'P8', 'Tên dự án/Hạng mục/Công việc': 'Nhật ký yêu cầu CEO',                        'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P8.3', 'Mã cấp trên': 'P8', 'Mã dự án': 'P8', 'Tên dự án/Hạng mục/Công việc': 'Master Tracker / RAG / SOS',                 'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P8.4', 'Mã cấp trên': 'P8', 'Mã dự án': 'P8', 'Tên dự án/Hạng mục/Công việc': 'Nhật ký quyết định và rủi ro',               'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P8.5', 'Mã cấp trên': 'P8', 'Mã dự án': 'P8', 'Tên dự án/Hạng mục/Công việc': 'Báo cáo CEO hằng tuần',                      'Trạng thái': 'Cần cập nhật' },
    { 'Loại bản ghi': 'HẠNG MỤC', 'Mã bản ghi': 'P8.6', 'Mã cấp trên': 'P8', 'Mã dự án': 'P8', 'Tên dự án/Hạng mục/Công việc': 'Kiểm soát phụ thuộc con người',              'Trạng thái': 'Cần cập nhật' }
  ];

  var duAnItems = seedData.filter(function(r) { return r['Loại bản ghi'] === 'DỰ ÁN'; });
  var hmItems   = seedData.filter(function(r) { return r['Loại bản ghi'] === 'HẠNG MỤC'; });
  Logger.log('Tổng đầu vào: ' + seedData.length + ' bản ghi (DỰ ÁN: ' + duAnItems.length + ', HẠNG MỤC: ' + hmItems.length + ')');

  var result = importCEOProjectSeedData(seedData);

  Logger.log('--- Kết quả ---');
  Logger.log('Thêm mới:              ' + result.added);
  Logger.log('Cập nhật hồ sơ DỰ ÁN: ' + result.updated);
  Logger.log('Bỏ qua trùng:          ' + result.skipped);
  Logger.log('Lỗi:                   ' + result.errors);

  Logger.log('--- Chi tiết theo dự án ---');
  var projectDefs = [
    {id:'P1', name:'GMS – Hệ thống quản lý GSP'},
    {id:'P2', name:'Kỷ luật nhà máy và nền tảng kho hàng'},
    {id:'P3', name:'Nền tảng dữ liệu và SSOT'},
    {id:'P4', name:'Lập kế hoạch, SCM và MRP'},
    {id:'P5', name:'Hệ điều hành chất lượng'},
    {id:'P6', name:'Hệ điều hành thương mại'},
    {id:'P7', name:'Ứng dụng công nghệ số và AI'},
    {id:'P8', name:'Quản trị PMO và kiểm soát tiếp nhận'}
  ];
  projectDefs.forEach(function(p) {
    var hms = hmItems.filter(function(r) { return r['Mã dự án'] === p.id; });
    var codes = hms.map(function(r) { return r['Mã bản ghi']; }).join(', ');
    Logger.log(p.id + ' [' + p.name + ']: ' + hms.length + ' HM — ' + codes);
  });

  Logger.log('--- ' + (result.errors === 0 ? 'OK' : 'CÓ LỖI') + ' ---');
}

function normalizeCPTRow_(r, i) {
  var deadline = formatDate(r['Thời hạn hoàn thành'] || '');
  var status   = String(r['Trạng thái'] || '').trim();
  return {
    LoaiBanGhi:        String(r['Loại bản ghi'] || '').trim(),
    MaBanGhi:          String(r['Mã bản ghi'] || '').trim(),
    MaCapTren:         String(r['Mã cấp trên'] || '').trim(),
    MaDuAn:            String(r['Mã dự án'] || '').trim(),
    TenBanGhi:         String(r['Tên dự án/Hạng mục/Công việc'] || '').trim(),
    MoTa:              String(r['Mô tả'] || '').trim(),
    BoiCanh:           String(r['Bối cảnh'] || '').trim(),
    MucTieu:           String(r['Mục tiêu'] || '').trim(),
    PhamVi:            String(r['Phạm vi'] || '').trim(),
    SanPhamDauRa:      String(r['Sản phẩm đầu ra'] || '').trim(),
    NguoiPhuTrach:     String(r['Người phụ trách'] || '').trim(),
    NguoiPhoiHop:      String(r['Người phối hợp'] || '').trim(),
    TrucCongViec:      String(r['Trục công việc'] || '').trim(),
    GiaiDoan:          String(r['Giai đoạn triển khai'] || '').trim(),
    MucDoUuTien:       String(r['Mức độ ưu tiên'] || '').trim(),
    NgayBatDau:        formatDate(r['Ngày bắt đầu'] || ''),
    ThoiHanHoanThanh:  deadline,
    TrangThai:         status,
    PhanTramHT:        String(r['Phần trăm hoàn thành'] || '').trim(),
    RAG:               String(r['RAG'] || '').trim(),
    CongViecDaThucHien: String(r['Công việc đã thực hiện'] || '').trim(),
    CapNhatMoiNhat:    String(r['Cập nhật mới nhất'] || '').trim(),
    KetQuaDatDuoc:     String(r['Kết quả đạt được'] || '').trim(),
    VuongMacRuiRo:     String(r['Vướng mắc/Rủi ro'] || '').trim(),
    HuongXuLy:         String(r['Hướng xử lý'] || '').trim(),
    KeHoachTiepTheo:   String(r['Kế hoạch tiếp theo'] || '').trim(),
    NdCEOCanBiet:      String(r['Nội dung CEO cần biết'] || '').trim(),
    NdCanCEOQuyet:     String(r['Nội dung cần CEO quyết'] || '').trim(),
    CacPhuongAn:       String(r['Các phương án đề xuất'] || '').trim(),
    QuyetDinhCEO:      String(r['Quyết định của CEO'] || '').trim(),
    NgayCapNhat:       formatDate(r['Ngày cập nhật'] || ''),
    NguoiCapNhat:      String(r['Người cập nhật'] || '').trim(),
    LinkFile:            String(r['Link bằng chứng/File nguồn'] || '').trim(),
    GhiChu:              String(r['Ghi chú'] || '').trim(),
    DebaiCEO:            String(r['Đề bài CEO'] || '').trim(),
    ViecDangTrienKhai:   String(r['Việc đang triển khai'] || '').trim(),
    TienDoHienTai:       String(r['Tiến độ hiện tại'] || '').trim(),
    IsOverdue:           isPastDeadline(deadline, status)
  };
}

function cptEmptySummary_() {
  return {
    tongDuAn:0, dangTrienKhai:0, dungTienDo:0, coNguyChamTre:0,
    quaHan:0, chuaCoNguoiPhuTrach:0, coRuiRo:0, canCEOQuyetDinh:0,
    duAnCanCanThiep:[], congViecQuaHanNghiemTrong:[], duAnChuaCapNhat7Ngay:[],
    congViecChuaCoNPT:[], canCEOQuyetTrong48h:[],
    sanPhamHoanThanhTuanNay:[], sanPhamDuKienTuanToi:[], uuTienXuLy5Viec:[]
  };
}

function getCEOProjectTrackerData() {
  var rawRows = readSheetSafe('THEO_DOI_DU_AN_CEO');
  if (!rawRows || rawRows.length === 0) {
    return { exists: false, total: 0, projects: [], hangMucs: [], congViecs: [],
             capNhats: [], quyetDinhs: [], summary: cptEmptySummary_() };
  }

  var duAns = [], hangMucs = [], congViecs = [], capNhats = [], quyetDinhs = [];
  rawRows.forEach(function(r, i) {
    var type = String(r['Loại bản ghi'] || '').trim().toUpperCase();
    var rec  = normalizeCPTRow_(r, i);
    if      (type === 'DỰ ÁN')      duAns.push(rec);
    else if (type === 'HẠNG MỤC')   hangMucs.push(rec);
    else if (type === 'CÔNG VIỆC')  congViecs.push(rec);
    else if (type === 'CẬP NHẬT')   capNhats.push(rec);
    else if (type === 'QUYẾT ĐỊNH') quyetDinhs.push(rec);
  });

  var today = new Date();
  var msDay = 864e5;

  var enriched = duAns.map(function(p) {
    var pid   = p.MaBanGhi;
    var hms   = hangMucs.filter(function(h) { return h.MaCapTren === pid || h.MaDuAn === pid; });
    var hmIds = hms.map(function(h) { return h.MaBanGhi; });
    var cvs   = congViecs.filter(function(c) { return c.MaDuAn === pid; });

    var cvDone    = cvs.filter(function(c) { return /hoàn thành/i.test(c.TrangThai); }).length;
    var cvOverdue = cvs.filter(function(c) {
      if (!c.ThoiHanHoanThanh || /hoàn thành|đã hủy/i.test(c.TrangThai)) return false;
      return new Date(c.ThoiHanHoanThanh) < today;
    }).length;
    var soQDCan = quyetDinhs.filter(function(q) {
      return (q.MaDuAn === pid || q.MaCapTren === pid || hmIds.indexOf(q.MaCapTren) !== -1)
        && /cần quyết định/i.test(q.NdCanCEOQuyet);
    }).length;

    // Tất cả CẬP NHẬT của dự án này, mới nhất trước
    var projCNs = capNhats.filter(function(u) {
      return u.MaDuAn === pid || u.MaCapTren === pid || hmIds.indexOf(u.MaCapTren) !== -1;
    }).sort(function(a, b) { return (b.NgayCapNhat || '').localeCompare(a.NgayCapNhat || ''); });
    var lastCN    = projCNs[0] || null;  // Bản ghi CẬP NHẬT mới nhất

    // Ngày cập nhật lần cuối (từ CẬP NHẬT, không lấy từ trường DỰ ÁN)
    var lastDate  = lastCN ? lastCN.NgayCapNhat : '';
    var daysSince = null;
    if (lastDate) {
      try { daysSince = Math.floor((today - new Date(lastDate)) / msDay); } catch(e2) {}
    }

    // Tính RAG hiệu lực: ưu tiên lấy từ DỰ ÁN, nếu trống thì từ CẬP NHẬT mới nhất
    var effectiveRAG = hasText(p.RAG) ? p.RAG : (lastCN && hasText(lastCN.RAG) ? lastCN.RAG : '');

    // Rủi ro: từ DỰ ÁN + CẬP NHẬT mới nhất + HẠNG MỤC + CÔNG VIỆC
    var soRuiRoFull = (hasText(p.VuongMacRuiRo) ? 1 : 0)
      + (lastCN && hasText(lastCN.VuongMacRuiRo) ? 1 : 0)
      + hms.filter(function(h) { return hasText(h.VuongMacRuiRo); }).length
      + cvs.filter(function(c) { return hasText(c.VuongMacRuiRo); }).length;

    return Object.assign({}, p, {
      // Số liệu tổng hợp
      SoHangMuc:          hms.length,
      TongCongViec:       cvs.length,
      CongViecHoanThanh:  cvDone,
      CongViecQuaHan:     cvOverdue,
      SoRuiRo:            soRuiRoFull,
      SoCEOQuyetCan:      soQDCan,
      TienDo:             cvs.length > 0 ? Math.round(cvDone / cvs.length * 100) : null,
      RAG:                effectiveRAG,

      // Cập nhật mới nhất (object đầy đủ — không thay thế hồ sơ dự án)
      LatestUpdate:       lastCN,
      UpdateHistory:      projCNs,
      NgayCapNhatGanNhat: lastDate,
      NgayKhongCapNhat:   daysSince
    });
  });

  var weekStart = new Date(today); weekStart.setDate(today.getDate() - today.getDay());
  var weekEnd   = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 7);
  var nxtWkEnd  = new Date(weekEnd);   nxtWkEnd.setDate(weekEnd.getDate() + 7);
  var priorOrd  = { 'Khẩn cấp':0, 'Cao':1, 'Trung bình':2, 'Thấp':3 };

  function slim(arr) {
    return arr.map(function(x) {
      return { ma: x.MaBanGhi, ten: x.TenBanGhi, nguoiPhuTrach: x.NguoiPhuTrach||'',
               han: x.ThoiHanHoanThanh||'', uuTien: x.MucDoUuTien||'', duAn: x.MaDuAn||'' };
    });
  }

  var canCT   = enriched.filter(function(p) { return p.RAG === 'Đỏ' || p.CongViecQuaHan > 2; });
  var cvQHNT  = congViecs.filter(function(c) {
    if (!c.ThoiHanHoanThanh || /hoàn thành|đã hủy/i.test(c.TrangThai)) return false;
    try { return Math.floor((today - new Date(c.ThoiHanHoanThanh)) / msDay) > 3; } catch(e2) { return false; }
  }).slice(0, 10);
  var chuaCN7 = enriched.filter(function(p) { return p.NgayKhongCapNhat === null || p.NgayKhongCapNhat > 7; });
  var cvNPT   = congViecs.filter(function(c) {
    return !hasText(c.NguoiPhuTrach) && !/hoàn thành|đã hủy/i.test(c.TrangThai);
  }).slice(0, 10);
  var ceoKhan = quyetDinhs.filter(function(q) {
    return /cần quyết định/i.test(q.NdCanCEOQuyet) && /khẩn cấp/i.test(q.MucDoUuTien);
  });
  var spTuanNay = congViecs.filter(function(c) {
    if (!/hoàn thành/i.test(c.TrangThai) || !c.ThoiHanHoanThanh) return false;
    try { var d = new Date(c.ThoiHanHoanThanh); return d >= weekStart && d < weekEnd; } catch(e2) { return false; }
  });
  var spTuanToi = congViecs.filter(function(c) {
    if (/hoàn thành|đã hủy/i.test(c.TrangThai) || !c.ThoiHanHoanThanh) return false;
    try { var d = new Date(c.ThoiHanHoanThanh); return d >= weekEnd && d < nxtWkEnd; } catch(e2) { return false; }
  });
  var uu5 = congViecs
    .filter(function(c) { return !/hoàn thành|đã hủy/i.test(c.TrangThai); })
    .sort(function(a, b) {
      return (priorOrd[a.MucDoUuTien] !== undefined ? priorOrd[a.MucDoUuTien] : 9)
           - (priorOrd[b.MucDoUuTien] !== undefined ? priorOrd[b.MucDoUuTien] : 9);
    }).slice(0, 5);

  return {
    exists:     true,
    total:      rawRows.length,
    projects:   enriched,
    hangMucs:   hangMucs,
    congViecs:  congViecs,
    capNhats:   capNhats,
    quyetDinhs: quyetDinhs,
    summary: {
      tongDuAn:                  enriched.length,
      dangTrienKhai:             enriched.filter(function(p) { return /đang triển khai/i.test(p.TrangThai); }).length,
      dungTienDo:                enriched.filter(function(p) { return p.RAG === 'Xanh'; }).length,
      coNguyChamTre:             enriched.filter(function(p) { return p.RAG === 'Vàng'; }).length,
      quaHan:                    enriched.filter(function(p) { return p.CongViecQuaHan > 0 || /quá hạn/i.test(p.TrangThai); }).length,
      chuaCoNguoiPhuTrach:       enriched.filter(function(p) { return !hasText(p.NguoiPhuTrach); }).length,
      coRuiRo:                   enriched.filter(function(p) { return p.SoRuiRo > 0; }).length,
      canCEOQuyetDinh:           enriched.filter(function(p) { return p.SoCEOQuyetCan > 0; }).length,
      duAnCanCanThiep:           slim(canCT),
      congViecQuaHanNghiemTrong: slim(cvQHNT),
      duAnChuaCapNhat7Ngay:      slim(chuaCN7),
      congViecChuaCoNPT:         slim(cvNPT),
      canCEOQuyetTrong48h:       slim(ceoKhan),
      sanPhamHoanThanhTuanNay:   slim(spTuanNay),
      sanPhamDuKienTuanToi:      slim(spTuanToi),
      uuTienXuLy5Viec:           slim(uu5)
    }
  };
}

function testGetCEOProjectTrackerData() {
  try {
    var data = getCEOProjectTrackerData();
    Logger.log('=== testGetCEOProjectTrackerData ===');
    Logger.log('Exists: ' + data.exists + ' | Tổng: ' + data.total);
    Logger.log('DỰ ÁN=' + data.projects.length + ' | HẠNG MỤC=' + data.hangMucs.length
      + ' | CÔNG VIỆC=' + data.congViecs.length + ' | CẬP NHẬT=' + data.capNhats.length
      + ' | QUYẾT ĐỊNH=' + data.quyetDinhs.length);
    data.projects.forEach(function(p, i) {
      Logger.log('[' + (i+1) + '] ' + p.MaBanGhi + ' | ' + p.TenBanGhi
        + ' | TT=' + p.TrangThai + ' | RAG=' + p.RAG
        + ' | Tiến độ=' + (p.TienDo !== null ? p.TienDo + '%' : 'N/A')
        + ' | CV quá hạn=' + p.CongViecQuaHan);
    });
    Logger.log('Summary: ' + JSON.stringify(data.summary));
    Logger.log('--- OK ---');
  } catch(e) {
    Logger.log('LỖI: ' + e.message + '\n' + e.stack);
  }
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
    Topic:        String(r['Chủ đề họp']           || r['Chủ đề biên bản'] || r['Chủ đề'] || ''),
    Chair:        String(r['Người chủ trì']        || ''),
    Attendees:    String(r['Thành phần tham dự']   || r['Thành phần'] || ''),
    Scope:        String(r['Phạm vi']              || ''),
    Summary:      String(r['Tóm tắt chính']        || r['Tóm tắt nhanh biên bản'] || r['Tóm tắt'] || ''),
    CEODirective: String(r['Chỉ đạo CEO / Nội dung đã chốt'] || r['Chỉ đạo CEO'] || r['Nội dung đã chốt'] || ''),
    Decisions:    String(r['Quyết định đã chốt']   || ''),
    Actions:      String(r['Action phát sinh']      || ''),
    RiskSOS:      String(r['Risk/SOS phát sinh']    || r['Risk SOS phát sinh'] || ''),
    CEODecision:  String(r['CEO cần chốt']          || r['CEO cần quyết tiếp'] || ''),
    FileMinutes:  String(r['File biên bản']         || ''),
    FileAudio:    String(r['File ghi âm/transcript']|| r['File ghi âm'] || r['File transcript'] || ''),
    RecordedBy:   String(r['Người nhập']            || r['Người lập biên bản'] || ''),
    Status:       String(r['Trạng thái biên bản']   || r['Trạng thái'] || 'Nháp'),
    SecretNote:   String(r['Ghi chú bảo mật']       || ''),
    DQ:           calcRowDataQuality_({ Action: r['Chủ đề họp'] || r['Chủ đề biên bản'] || '', Owner: r['Người nhập'] || r['Người lập biên bản'] || '', PIC: r['Người chủ trì'] || '', Deadline: r['Ngày họp'] || '', Status: r['Trạng thái biên bản'] || r['Trạng thái'] || '', RAG: '' })
  };
}

/* ── Normalize MEETING_TOPICS ── */
function normalizeMeetingTopic_(r, i) {
  return {
    TopicID:   String(r['Topic ID']            || ('TOP-' + String(i + 1).padStart(3, '0'))),
    TopicName: String(r['Topic Name']          || r['Tên Topic'] || '').trim(),
    Module:    String(r['Nhóm/Module']         || r['Nhóm'] || '').trim(),
    Desc:      String(r['Mô tả chủ đề']        || r['Mô tả'] || '').trim(),
    Summary:   String(r['Tóm tắt nhanh Topic'] || r['Tóm tắt'] || '').trim(),
    Owner:     String(r['Owner theo dõi']       || '').trim(),
    Status:    String(r['Trạng thái']           || 'Đang diễn ra').trim(),
    Note:      String(r['Ghi chú']             || '').trim(),
    MeetingID: String(r['Meeting ID']          || '').trim()
  };
}

/* ── Normalize MEETING_MINUTES full format ── */
function normalizeMeetingMinutesFull_(r, i) {
  return {
    MeetingID:    String(r['Meeting ID']              || ('MTG-' + String(i + 1).padStart(3, '0'))),
    TopicID:      String(r['Topic ID']                || '').trim(),
    Date:         formatDate(r['Ngày họp']            || ''),
    Topic:        String(r['Chủ đề biên bản']         || r['Chủ đề họp'] || r['Chủ đề'] || '').trim(),
    Module:       String(r['Nhóm/Module']             || r['Nhóm'] || '').trim(),
    Chair:        String(r['Người chủ trì']           || '').trim(),
    Attendees:    String(r['Thành phần tham dự']      || r['Thành phần'] || '').trim(),
    Summary:      String(r['Tóm tắt nhanh biên bản']  || r['Tóm tắt chính'] || r['Tóm tắt'] || '').trim(),
    CEODirective: String(r['Chỉ đạo CEO / Nội dung đã chốt'] || r['Chỉ đạo CEO'] || r['Nội dung đã chốt'] || '').trim(),
    Actions:      String(r['Action phát sinh']         || '').trim(),
    CEODecision:  String(r['CEO cần chốt']             || r['CEO cần quyết tiếp'] || '').trim(),
    Status:       String(r['Trạng thái biên bản']      || r['Trạng thái'] || 'Nháp').trim(),
    FileMinutes:  String(r['File biên bản']            || '').trim(),
    FileAudio:    String(r['File ghi âm/transcript']   || r['File ghi âm'] || r['File transcript'] || '').trim(),
    RecordedBy:   String(r['Người nhập']               || r['Người lập biên bản'] || '').trim(),
    UpdatedAt:    formatDate(r['Ngày cập nhật']        || ''),
    Note:         String(r['Ghi chú']                 || '').trim(),
    FullContent:  String(r['Nội dung biên bản đầy đủ'] || '').trim(),
    DetailURL:    String(r['Full Detail URL']          || r['Link xem chi tiết'] || '').trim()
  };
}

/* ── Normalize MEETING_ACTIONS ── (Cột A–O: giữ nguyên; P–U: mới) */
function normalizeMeetingAction_(r, i) {
  var rag      = normalizeRag(r['RAG'] || '');
  var deadline = formatDate(r['Deadline'] || '');
  var status   = String(r['Trạng thái'] || 'Chưa bắt đầu').trim();
  var pctRaw   = r['% hoàn thành'] || r['% Hoàn thành'] || '';
  var pct      = parseInt(String(pctRaw).replace('%', '').trim(), 10);
  if (isNaN(pct) || pct < 0) pct = 0;
  if (pct > 100) pct = 100;
  return {
    ActionID:       String(r['Action ID']         || ('MACT-' + String(i + 1).padStart(3, '0'))),
    TopicID:        String(r['Topic ID']          || '').trim(),
    MeetingID:      String(r['Meeting ID']        || '').trim(),
    Date:           formatDate(r['Ngày họp']      || ''),
    Module:         String(r['Nhóm/Module']       || '').trim(),
    Action:         String(r['Nội dung action']   || '').trim(),
    Owner:          String(r['Owner/PIC']         || r['PIC'] || '').trim(),
    CoOwner:        String(r['Phối hợp']          || '').trim(),
    Deadline:       deadline,
    Output:         String(r['Đầu ra kỳ vọng']    || '').trim(),
    Status:         status,
    RAG:            rag,
    CEODecision:    String(r['CEO cần chốt']      || '').trim(),
    FileMinutes:    String(r['File biên bản']     || '').trim(),
    Note:           String(r['Ghi chú']           || '').trim(),
    KetQuaThucTe:  String(r['Kết quả thực tế']            || '').trim(),
    PctHoanThanh:  pct,
    NgayCapNhat:   formatDate(r['Ngày cập nhật']           || ''),
    NguoiCapNhat:  String(r['Người cập nhật']              || '').trim(),
    BangChung:     String(r['Bằng chứng/Link']             || '').trim(),
    VuongMac:      String(r['Vướng mắc/Hỗ trợ cần thiết'] || '').trim(),
    IsOverdue:     isPastDeadline(deadline, status)
  };
}

/* ── getMeetingCenterData() — new 3-sheet meeting data ── */
function getMeetingCenterData() {
  var topics  = readSheetSafe('MEETING_TOPICS').map(normalizeMeetingTopic_);
  var minutes = readSheetSafe('MEETING_MINUTES').map(normalizeMeetingMinutesFull_);
  var actions = readSheetSafe('MEETING_ACTIONS').map(normalizeMeetingAction_);
  return { topics: topics, minutes: minutes, actions: actions };
}

/* ── Test function — run manually from Apps Script editor ── */
function testGetMeetingCenterData() {
  try {
    var data = getMeetingCenterData();
    var topics  = data.topics  || [];
    var minutes = data.minutes || [];
    var actions = data.actions || [];

    Logger.log('=== testGetMeetingCenterData ===');
    Logger.log('Topics  (' + topics.length  + '):');
    topics.forEach(function(t, i) {
      Logger.log('  [' + (i+1) + '] ' + t.TopicID + ' | ' + t.TopicName + ' | ' + t.Module + ' | ' + t.Status);
    });
    Logger.log('Minutes (' + minutes.length + '):');
    minutes.slice(0, 5).forEach(function(m, i) {
      Logger.log('  [' + (i+1) + '] ' + m.MeetingID + ' | TopicID=' + m.TopicID + ' | ' + m.Date + ' | ' + m.Topic.substring(0,40));
    });
    if (minutes.length > 5) Logger.log('  … (' + (minutes.length - 5) + ' biên bản còn lại không hiển thị)');
    Logger.log('Actions (' + actions.length + '):');
    actions.slice(0, 5).forEach(function(a, i) {
      Logger.log('  [' + (i+1) + '] ' + a.ActionID + ' | MtgID=' + a.MeetingID + ' | ' + a.Owner + ' | ' + a.Status + ' | ' + a.RAG);
    });
    if (actions.length > 5) Logger.log('  … (' + (actions.length - 5) + ' action còn lại không hiển thị)');

    Logger.log('--- OK ---');
    Logger.log('Kết quả: ' + topics.length + ' topic | ' + minutes.length + ' biên bản | ' + actions.length + ' action');
  } catch(e) {
    Logger.log('LỖI testGetMeetingCenterData: ' + e.message + '\n' + e.stack);
  }
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
  var topics   = data.meetingTopics     || [];

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
  var mmPending     = mm.filter(function(r) { return /chờ|pending/i.test(r.Status || ''); }).length;
  var mmHasCEO      = mm.filter(function(r) { return hasText(r.CEODirective); }).length;
  var mmHasRisk     = mm.filter(function(r) { return hasText(r.RiskSOS); }).length;
  var mmHasActions  = mm.filter(function(r) { return hasText(r.Actions); }).length;
  var mmHasDecision = mm.filter(function(r) { return hasText(r.CEODecision); }).length;

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
    meetingsTopicsTotal:        topics.length,
    meetingsTotal:              mm.length,
    meetingsPendingApproval:    mmPending,
    meetingsWithCEOInstruction: mmHasCEO,
    meetingsWithRisk:           mmHasRisk,
    meetingsWithActions:        mmHasActions,
    meetingsWithCEODecision:    mmHasDecision,

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
  Logger.log('Meetings: Topics=' + (data.meetingTopics||[]).length + ' | Minutes=' + data.meetingMinutes.length + ' | PendingApproval: ' + s.meetingsPendingApproval);
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
  var ss             = SpreadsheetApp.getActiveSpreadsheet();
  var now            = new Date();
  var thirtyDaysAgo  = new Date(now.getTime() - 30 * 24 * 3600 * 1000);

  // ── 5 sheet reads ──
  var issues      = readSheetSafe('TO_ISSUES').map(normalizeToIssue_);
  var meetings    = getTOMeetings_(ss);
  var topics      = readSheetSafe('MEETING_TOPICS').map(normalizeMeetingTopic_);
  var actions     = readSheetSafe('TO_ACTIONS').map(normalizeTOActionFull_);
  var ceoComments = getCEOComments_(ss);

  // ── Hot issues: lọc từ TO_ISSUES (RAG=Red hoặc Mức độ=Cao/High/Critical) ──
  var hotIssues = issues.filter(function(iss) {
    var lvl = String(iss.Level || '').toLowerCase();
    return iss.RAG === 'Red'
      || lvl === 'cao' || lvl === 'high' || lvl === 'critical' || lvl.includes('khẩn');
  });

  // ── Action buckets ──
  var openActions    = [];
  var overdueActions = [];
  var doneActions    = [];
  var noOwner        = [];
  var ceoList        = [];

  actions.forEach(function(a) {
    var s      = String(a.Status || '').toLowerCase();
    var isDone = s === 'done' || s.includes('hoàn') || s.includes('xong');
    var isOv   = a.IsOverdue;

    if (isDone) {
      var updDate = a.NgayCapNhat ? parseDate(a.NgayCapNhat) : null;
      if (!updDate || updDate >= thirtyDaysAgo) doneActions.push(a);
    } else if (isOv) {
      overdueActions.push(a);
    } else {
      openActions.push(a);
    }

    if ((isNeedAssign_(a.PIC) || isNeedAssign_(a.Owner)) && !isDone) noOwner.push(a);
    if (hasText(a.CEODecision) && !isNoValue_(a.CEODecision))         ceoList.push(a);
  });

  // CEO_COMMENTS còn chờ xử lý
  var ceoPending = ceoComments.filter(function(c) {
    var s = String(c.TrangThaiXuLy || '').toLowerCase();
    return s !== 'đã xong' && s !== 'hoàn thành' && s !== 'done';
  });

  // ── Factories từ TO_ISSUES ──
  var factorySet = {};
  issues.forEach(function(x) { if (x.Factory) factorySet[x.Factory] = true; });
  var factories = Object.keys(factorySet).sort();

  // ── Issues summary ──
  var issRAG = { Green: 0, Amber: 0, Red: 0 };
  issues.forEach(function(iss) {
    var r = iss.RAG || 'Amber';
    if (issRAG[r] !== undefined) issRAG[r]++; else issRAG.Amber++;
  });

  return JSON.parse(JSON.stringify({
    updatedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),

    factories:   factories,
    issues:      issues,
    hotIssues:   hotIssues,
    meetings:    meetings,
    topics:      topics,
    actions:     actions,
    ceoComments: ceoComments,

    openActions:    openActions,
    overdueActions: overdueActions,
    doneActions:    doneActions,
    ceoList:        ceoList,
    ceoPending:     ceoPending,
    noOwnerList:    noOwner,

    issuesSummary: {
      total: issues.length,
      open:  issues.filter(function(i) { return i.Status !== 'Đã xử lý' && i.Status !== 'Đóng'; }).length,
      byRAG: issRAG,
      hot:   hotIssues.length
    },

    actionsSummary: {
      total:   actions.length,
      open:    openActions.length,
      overdue: overdueActions.length,
      done:    doneActions.length,
      ceo:     ceoList.length + ceoPending.length,
      noOwner: noOwner.length
    }
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
    Note:         String(r['Ghi chú']             || '').trim(),
    MeetingIDRef: String(r['Meeting ID liên quan'] || '').trim()
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
    TopicID:      String(r['Topic ID']                 || '').trim(),
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
    CoOwner:      String(r['Người phối hợp']           || '').trim(),
    PctHoanThanh: r['% hoàn thành'] !== '' && r['% hoàn thành'] !== undefined ? (Number(r['% hoàn thành']) || 0) : 0,
    KetQuaThucTe: String(r['Kết quả thực tế']          || '').trim(),
    NgayCapNhat:  formatDate(r['Ngày cập nhật']         || ''),
    NguoiCapNhat: String(r['Người cập nhật']            || '').trim(),
    BangChung:    String(r['Bằng chứng/Link']           || '').trim(),
    VuongMac:     String(r['Vướng mắc/Hỗ trợ cần thiết'] || '').trim(),
    CEODecision:  String(r['CEO cần chốt']              || '').trim(),
    IsOverdue:    isPastDeadline(deadline, status)
  };
}


/* ── Normalize TO_MEETINGS row ── */
function normalizeTOMeeting_(r) {
  return {
    MeetingID:   String(r['Meeting ID']    || '').trim(),
    NgayHop:     formatDate(r['Ngày họp']  || ''),
    TieuDe:      String(r['Tiêu đề']       || '').trim(),
    LoaiHop:     String(r['Loại họp']      || '').trim(),
    DiaDiem:     String(r['Địa điểm']      || '').trim(),
    ThanhPhan:   String(r['Thành phần']    || '').trim(),
    ChuTri:      String(r['Chủ trì']       || '').trim(),
    ThuKy:       String(r['Thư ký']        || '').trim(),
    TomTat:      String(r['Tóm tắt']       || '').trim(),
    CanLamRo:    String(r['Cần làm rõ']    || '').trim(),
    CanCEO:      String(r['Cần CEO']       || '').trim(),
    FileBienBan: String(r['File biên bản'] || '').trim(),
    RAG:         normalizeRag(r['RAG']     || ''),
    TrangThai:   String(r['Trạng thái']    || '').trim()
  };
}

/* ── Đọc TO_MEETINGS — tự tạo sheet nếu chưa tồn tại ── */
function getTOMeetings_(ss) {
  var sheet = ss.getSheetByName('TO_MEETINGS');
  if (!sheet) {
    sheet = ss.insertSheet('TO_MEETINGS');
    var headers = [
      'Meeting ID', 'Ngày họp', 'Tiêu đề', 'Loại họp', 'Địa điểm',
      'Thành phần', 'Chủ trì', 'Thư ký', 'Tóm tắt',
      'Cần làm rõ', 'Cần CEO', 'File biên bản', 'RAG', 'Trạng thái'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#064e3b').setFontColor('#ffffff').setFontWeight('bold');
    SpreadsheetApp.flush();
    return [];
  }
  var rows = readSheetSafe('TO_MEETINGS');
  return rows.map(normalizeTOMeeting_).filter(function(r) { return r.MeetingID; });
}

/* ── Normalize CEO_COMMENTS row ── */
function normalizeCeoComment_(r, i) {
  return {
    CommentID:     String(r['Comment ID']          || ('CEO-' + String(i + 1).padStart(3, '0'))),
    Ngay:          formatDate(r['Ngày']             || ''),
    MeetingIDRef:  String(r['Meeting ID liên quan'] || '').trim(),
    IssueIDRef:    String(r['Issue ID liên quan']   || '').trim(),
    ActionIDRef:   String(r['Action ID liên quan']  || '').trim(),
    NoiDung:       String(r['Nội dung CEO']         || '').trim(),
    Loai:          String(r['Loại']                 || '').trim(),
    NguoiGhi:      String(r['Người ghi']            || '').trim(),
    TrangThaiXuLy: String(r['Trạng thái xử lý']    || 'Chờ xử lý').trim(),
    GhiChu:        String(r['Ghi chú']              || '').trim()
  };
}

/* ── Đọc CEO_COMMENTS — tự tạo sheet nếu chưa tồn tại ── */
function getCEOComments_(ss) {
  var sheet = ss.getSheetByName('CEO_COMMENTS');
  if (!sheet) {
    sheet = ss.insertSheet('CEO_COMMENTS');
    var headers = [
      'Comment ID', 'Ngày', 'Meeting ID liên quan', 'Issue ID liên quan',
      'Action ID liên quan', 'Nội dung CEO', 'Loại', 'Người ghi',
      'Trạng thái xử lý', 'Ghi chú'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
    SpreadsheetApp.flush();
    return [];
  }
  return readSheetSafe('CEO_COMMENTS').map(normalizeCeoComment_).filter(function(r) { return r.CommentID; });
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

/* ────────────────────────────────────────────────────────
   TO ACTION CENTER — WRITE FUNCTIONS
   - ensureTOActionNewCols_   : thêm 8 cột mới nếu chưa có
   - generateNextIssueID_     : sinh Issue ID tiếp theo
   - generateNextActionID_    : sinh Action ID tiếp theo
   - toWriteNewIssue_         : ghi dòng mới vào TO_ISSUES
   - toWriteNewAction_        : ghi dòng mới vào TO_ACTIONS
   - toUpdateExistingAction_  : cập nhật dòng có sẵn trong TO_ACTIONS
──────────────────────────────────────────────────────── */

function ensureTOActionNewCols_(ss) {
  var sheet = ss.getSheetByName('TO_ACTIONS');
  if (!sheet) return [];
  var newCols = [
    'Người phối hợp', '% hoàn thành', 'Kết quả thực tế',
    'Ngày cập nhật', 'Người cập nhật', 'Bằng chứng/Link',
    'Vướng mắc/Hỗ trợ cần thiết', 'CEO cần chốt'
  ];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  var added = [];
  newCols.forEach(function(col) {
    if (headers.indexOf(col) < 0) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(col);
      headers.push(col);
      added.push(col);
    }
  });
  if (added.length) SpreadsheetApp.flush();
  return added;
}

function idExistsInSheet_(sheet, id) {
  if (!sheet) return false;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var target = String(id).trim().toLowerCase();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim().toLowerCase() === target) return true;
  }
  return false;
}

function generateNextIssueID_(ss, factory) {
  var sheet = ss.getSheetByName('TO_ISSUES');
  var prefix = 'TOI-' + factory + '-';
  var prefixLc = prefix.toLowerCase();
  var lastRow = sheet ? sheet.getLastRow() : 1;
  var ids = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r){ return String(r[0]||'').trim(); })
    : [];
  var matched = ids.filter(function(id){ return id.toLowerCase().startsWith(prefixLc); });
  var maxSeq = 0;
  matched.forEach(function(id) {
    var parts = id.split('-');
    var last = parseInt(parts[parts.length - 1]) || 0;
    if (last > maxSeq) maxSeq = last;
  });
  var seq = String(maxSeq + 1).padStart(3, '0');
  // Detect pattern: with date component if sample has 4+ parts with 8-digit date
  var hasDate = matched.some(function(id) {
    var parts = id.split('-'); return parts.length >= 4 && /^\d{8}$/.test(parts[2]);
  });
  if (hasDate) {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    return 'TOI-' + factory + '-' + today + '-' + seq;
  }
  return 'TOI-' + factory + '-' + seq;
}

function generateNextActionID_(ss, factory) {
  var sheet = ss.getSheetByName('TO_ACTIONS');
  var prefix = 'TOA-' + factory + '-';
  var prefixLc = prefix.toLowerCase();
  var lastRow = sheet ? sheet.getLastRow() : 1;
  var ids = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r){ return String(r[0]||'').trim(); })
    : [];
  var matched = ids.filter(function(id){ return id.toLowerCase().startsWith(prefixLc); });
  var maxSeq = 0;
  matched.forEach(function(id) {
    var parts = id.split('-');
    var last = parseInt(parts[parts.length - 1]) || 0;
    if (last > maxSeq) maxSeq = last;
  });
  var seq = String(maxSeq + 1).padStart(3, '0');
  var hasDate = matched.some(function(id) {
    var parts = id.split('-'); return parts.length >= 4 && /^\d{8}$/.test(parts[2]);
  });
  if (hasDate) {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    return 'TOA-' + factory + '-' + today + '-' + seq;
  }
  return 'TOA-' + factory + '-' + seq;
}

function toWriteNewIssue_(ss, payload) {
  try {
    var sheet = ss.getSheetByName('TO_ISSUES');
    if (!sheet) return { success: false, error: 'Sheet TO_ISSUES không tồn tại.' };
    var factory = String(payload.factory || '').trim();
    if (!factory) return { success: false, error: 'Thiếu thông tin Nhà máy.' };
    if (!String(payload.description || '').trim()) return { success: false, error: 'Thiếu Mô tả vấn đề.' };
    var issueId = generateNextIssueID_(ss, factory);
    var tries = 0;
    while (idExistsInSheet_(sheet, issueId) && tries < 20) {
      var p = issueId.split('-'); var n = parseInt(p[p.length-1])||0;
      p[p.length-1] = String(n+1).padStart(3,'0'); issueId = p.join('-'); tries++;
    }
    if (tries >= 20) return { success: false, error: 'Không thể sinh Issue ID không trùng.' };
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var data = {
      'Issue ID':            issueId,
      'Nhà máy':             factory,
      'Người ghi nhận':      String(payload.reporter   || '').trim(),
      'Ngày ghi nhận':       String(payload.reportDate || today).trim(),
      'File nguồn':          String(payload.fileSource || '').trim(),
      'Khu vực':             String(payload.area       || '').trim(),
      'Workstream':          String(payload.workstream || '').trim(),
      'Loại vấn đề':         String(payload.issueType  || '').trim(),
      'Mô tả vấn đề':        String(payload.description|| '').trim(),
      'Nguyên nhân':         String(payload.rootCause  || '').trim(),
      'Ảnh hưởng':           String(payload.impact     || '').trim(),
      'Mức độ':              String(payload.level      || '').trim(),
      'RAG':                 String(payload.rag        || '').trim(),
      'Trạng thái':          String(payload.status     || 'Mới').trim(),
      'Action ID liên quan':  '',
      'Meeting ID liên quan': String(payload.meetingIdRef || '').trim(),
      'Ghi chú':              String(payload.note         || '').trim()
    };
    appendRowByHeaders_(sheet, data);
    SpreadsheetApp.flush();
    return { success: true, issueId: issueId, message: 'Đã ghi nhận vấn đề ' + issueId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function toWriteNewAction_(ss, payload) {
  try {
    ensureTOActionNewCols_(ss);
    var sheet = ss.getSheetByName('TO_ACTIONS');
    if (!sheet) return { success: false, error: 'Sheet TO_ACTIONS không tồn tại.' };
    var factory = String(payload.factory || '').trim();
    if (!factory) return { success: false, error: 'Thiếu Nhà máy.' };
    if (!String(payload.action || '').trim()) return { success: false, error: 'Thiếu Nội dung công việc.' };
    if (!String(payload.pic || '').trim()) return { success: false, error: 'Thiếu PIC thực hiện.' };
    if (!String(payload.deadline || '').trim()) return { success: false, error: 'Thiếu Deadline.' };
    var actionId = generateNextActionID_(ss, factory);
    var tries = 0;
    while (idExistsInSheet_(sheet, actionId) && tries < 20) {
      var p = actionId.split('-'); var n = parseInt(p[p.length-1])||0;
      p[p.length-1] = String(n+1).padStart(3,'0'); actionId = p.join('-'); tries++;
    }
    if (tries >= 20) return { success: false, error: 'Không thể sinh Action ID không trùng.' };
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var data = {
      'Action ID':                    actionId,
      'Nhà máy':                      factory,
      'Người ghi nhận':               String(payload.reporter    || '').trim(),
      'Ngày ghi nhận':                String(payload.reportDate  || today).trim(),
      'Ngày phát sinh':               String(payload.issueDate   || '').trim(),
      'Meeting ID':                   String(payload.meetingIdRef || '').trim(),
      'Topic ID':                     String(payload.topicId      || '').trim(),
      'Cuộc họp / Bối cảnh':         String(payload.context      || '').trim(),
      'Issue ID liên quan':           String(payload.issueIdRef  || '').trim(),
      'Khu vực':                      String(payload.area        || '').trim(),
      'Nội dung action':              String(payload.action      || '').trim(),
      'PIC thực hiện':                String(payload.pic         || '').trim(),
      'Owner theo dõi':               String(payload.owner       || '').trim(),
      'Deadline':                     String(payload.deadline    || '').trim(),
      'Loại kế hoạch':                String(payload.planType    || '').trim(),
      'Trạng thái':                   String(payload.status      || 'Chưa xử lý').trim(),
      'Mức độ ưu tiên':               String(payload.priority    || '').trim(),
      'RAG':                          String(payload.rag         || '').trim(),
      'Kết quả / Output':             String(payload.output      || '').trim(),
      'CEO cần biết':                 String(payload.ceoNote     || '').trim(),
      'File nguồn':                   String(payload.fileSource  || '').trim(),
      'Link biên bản':                String(payload.rowUrl      || '').trim(),
      'Ghi chú':                      String(payload.note        || '').trim(),
      'Người phối hợp':               String(payload.coOwner     || '').trim(),
      '% hoàn thành':                 payload.pct !== undefined && payload.pct !== '' ? Number(payload.pct) : '',
      'Kết quả thực tế':              String(payload.ketQua      || '').trim(),
      'Ngày cập nhật':                String(payload.updDate     || '').trim(),
      'Người cập nhật':               String(payload.updBy       || '').trim(),
      'Bằng chứng/Link':              String(payload.bangChung   || '').trim(),
      'Vướng mắc/Hỗ trợ cần thiết':  String(payload.vuongMac    || '').trim(),
      'CEO cần chốt':                 String(payload.ceoDecision || '').trim()
    };
    appendRowByHeaders_(sheet, data);
    SpreadsheetApp.flush();
    return { success: true, actionId: actionId, message: 'Đã tạo công việc ' + actionId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function toUpdateExistingAction_(ss, payload) {
  try {
    ensureTOActionNewCols_(ss);
    var sheet = ss.getSheetByName('TO_ACTIONS');
    if (!sheet) return { success: false, error: 'Sheet TO_ACTIONS không tồn tại.' };
    var actionId = String(payload.actionId || '').trim();
    if (!actionId) return { success: false, error: 'Thiếu Action ID.' };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Không tìm thấy Action ID ' + actionId + '.' };
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var rowIndex = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || '').trim().toLowerCase() === actionId.toLowerCase()) {
        rowIndex = i + 2; break;
      }
    }
    if (rowIndex < 0) return { success: false, error: 'Không tìm thấy Action ID ' + actionId + ' trong TO_ACTIONS.' };
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var allowed = {
      'PIC thực hiện':                payload.pic,
      'Owner theo dõi':               payload.owner,
      'Deadline':                     payload.deadline,
      'Trạng thái':                   payload.status,
      'RAG':                          payload.rag,
      'Kết quả / Output':             payload.output,
      'CEO cần biết':                 payload.ceoNote,
      'Ghi chú':                      payload.note,
      'Người phối hợp':               payload.coOwner,
      '% hoàn thành':                 payload.pct !== undefined && payload.pct !== '' ? Number(payload.pct) : undefined,
      'Kết quả thực tế':              payload.ketQua,
      'Ngày cập nhật':                payload.updDate,
      'Người cập nhật':               payload.updBy,
      'Bằng chứng/Link':              payload.bangChung,
      'Vướng mắc/Hỗ trợ cần thiết':  payload.vuongMac,
      'CEO cần chốt':                 payload.ceoDecision
    };
    var updated = [];
    headers.forEach(function(hdr, ci) {
      if (allowed.hasOwnProperty(hdr) && allowed[hdr] !== undefined && allowed[hdr] !== null) {
        sheet.getRange(rowIndex, ci + 1).setValue(allowed[hdr]);
        updated.push(hdr);
      }
    });
    SpreadsheetApp.flush();
    return { success: true, actionId: actionId, updatedFields: updated, message: 'Đã cập nhật ' + actionId + ' (' + updated.length + ' trường)' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function saveCEOComment_(ss, payload) {
  try {
    getCEOComments_(ss); // ensures sheet + headers exist
    var sheet = ss.getSheetByName('CEO_COMMENTS');
    if (!sheet) return { success: false, error: 'Không thể tạo sheet CEO_COMMENTS.' };
    var noiDung  = String(payload.noiDung  || '').trim();
    var nguoiGhi = String(payload.nguoiGhi || '').trim();
    if (!noiDung)  return { success: false, error: 'Thiếu Nội dung chỉ đạo CEO.' };
    if (!nguoiGhi) return { success: false, error: 'Thiếu Người ghi.' };
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var data = {
      'Comment ID':           '',
      'Ngày':                 String(payload.ngay          || today).trim(),
      'Meeting ID liên quan': String(payload.meetingIdRef  || '').trim(),
      'Issue ID liên quan':   String(payload.issueIdRef    || '').trim(),
      'Action ID liên quan':  String(payload.actionIdRef   || '').trim(),
      'Nội dung CEO':         noiDung,
      'Loại':                 String(payload.loai          || '').trim(),
      'Người ghi':            nguoiGhi,
      'Trạng thái xử lý':    String(payload.trangThaiXuLy || 'Chờ xử lý').trim(),
      'Ghi chú':              String(payload.ghiChu        || '').trim()
    };
    appendRowByHeaders_(sheet, data);
    SpreadsheetApp.flush();
    return { success: true, message: 'Đã ghi nhận chỉ đạo CEO thành công.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
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

/* =====================================================
   KIẾN TRÚC MỚI — 9 SHEET CHÍNH
   Bước 4: Sao lưu sheet cũ
   Bước 5: Tạo cấu trúc 9 sheet mới

   Chạy thủ công từ Apps Script editor:
     setupNewArchitecture_()
===================================================== */

function setupNewArchitecture_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var log = [];

  try {
    log.push('▶ BƯỚC 4 — Sao lưu sheet cũ...');
    log = log.concat(backupExistingSheets_(ss));

    log.push('');
    log.push('▶ BƯỚC 5 — Tạo cấu trúc 9 sheet mới...');
    log = log.concat(createNewSheetStructure_(ss));

    SpreadsheetApp.flush();
    var msg = log.join('\n');
    Logger.log(msg);
    ui.alert('✅ Hoàn thành!\n\n' + msg);
  } catch (err) {
    Logger.log('LỖI setupNewArchitecture_: ' + err.message + '\n' + err.stack);
    ui.alert('❌ Lỗi: ' + err.message + '\n\nXem chi tiết trong Apps Script Logger (Ctrl+Enter).');
  }
}

/* ── Bước 4: Sao lưu các sheet quan trọng ── */
function backupExistingSheets_(ss) {
  var log = [];
  var toBackup = [
    'THEO_DOI_DU_AN_CEO', 'CEO_PROJECTS', 'CEO_PROJECT_UPDATES',
    'PMO_ALL_ACTIONS', 'PMO_90D_ACTIONS',
    'TO_ISSUES', 'TO_ACTIONS', 'TO_MEETINGS',
    'MEETING_MINUTES', 'MEETING_TOPICS', 'MEETING_ACTIONS',
    'CEO_DECISIONS', 'CEO_COMMENTS', 'CEO_Comments',
    'RISK_SOS'
  ];

  toBackup.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      log.push('  ⬜ Không tìm thấy: ' + name);
      return;
    }
    var bakName = name + '_BAK';
    if (ss.getSheetByName(bakName)) {
      log.push('  ⬜ Đã có bản sao rồi: ' + bakName + ' — bỏ qua');
      return;
    }
    sh.copyTo(ss).setName(bakName);
    log.push('  ✅ Đã sao lưu: ' + name + '  →  ' + bakName);
  });

  return log;
}

/* =====================================================
   KIẾN TRÚC MỚI — 8 SHEET CHÍNH
   1 Module = 1 Sheet = 1 Nguồn dữ liệu duy nhất
   Không ID · Không sheet trung gian · Không chuẩn hóa SQL
===================================================== */

/* ── Tự tính RAG từ trạng thái + deadline + vướng mắc ── */
function calcRAG_(status, deadline, vuongMac, ceoCanChot) {
  var s = String(status || '').toLowerCase().trim();
  if (/hoàn thành|done|xong/i.test(s)) return 'Green';
  if (isPastDeadline(deadline, status))  return 'Red';
  if (hasText(ceoCanChot))               return 'Red';
  if (hasText(vuongMac))                 return 'Amber';
  if (deadline) {
    var d = parseDate(deadline);
    if (d) {
      var diff = (d - new Date()) / (1000 * 60 * 60 * 24);
      if (diff >= 0 && diff <= 7) return 'Amber';
    }
  }
  if (/tạm dừng|dừng|blocked/i.test(s)) return 'Amber';
  return 'Green';
}

/* ── Tiến độ text → % số ── */
function tienDoToPct_(tienDo) {
  var t = String(tienDo || '').toLowerCase().trim();
  if (/hoàn thành|done|xong/i.test(t))   return 100;
  if (/gần hoàn thành/i.test(t))          return 80;
  if (/đang thực hiện/i.test(t))          return 50;
  if (/tạm dừng/i.test(t))               return 0;
  return 0;
}

/* ── Normalizer: 01_PMO_GSP_NEXT30 ── */
function normalizePMO_(r) {
  var status  = String(r['Trạng thái']      || '').trim();
  var deadline= formatDate(r['Deadline']     || '');
  var vuong   = String(r['Vướng mắc']       || '').trim();
  var ceoChot = String(r['CEO cần chốt']    || '').trim();
  var tienDo  = String(r['Tiến độ']         || '').trim();
  var rag     = calcRAG_(status, deadline, vuong, ceoChot);
  return {
    PIC:          String(r['PIC / Trục']          || '').trim(),
    NhomViec:     String(r['Nhóm công việc']       || '').trim(),
    NoiDung:      String(r['Nội dung công việc']   || '').trim(),
    NguoiPhuTrach:String(r['Người phụ trách']      || '').trim(),
    Deadline:     deadline,
    TrangThai:    status,
    TienDo:       tienDo,
    PctHT:        tienDoToPct_(tienDo),
    VuongMac:     vuong,
    ViecTiepTheo: String(r['Việc tiếp theo']       || '').trim(),
    CEOCanChot:   ceoChot,
    CapNhat:      String(r['Cập nhật mới nhất']    || '').trim(),
    Link:         String(r['Link tài liệu']        || '').trim(),
    RAG:          rag,
    IsOverdue:    isPastDeadline(deadline, status),
    // Legacy aliases
    Action:       String(r['Nội dung công việc']  || '').trim(),
    Owner:        String(r['Người phụ trách']     || '').trim(),
    PICName:      String(r['PIC / Trục']          || '').trim(),
    Status:       status,
    Workstream:   String(r['PIC / Trục']          || '').trim(),
    CEODecision:  ceoChot,
    CEONote:      ceoChot,
    Risk:         vuong
  };
}

/* ── Normalizer: 02_DU_AN_CEO ── */
function normalizeDuAnCEO_(r) {
  var status  = String(r['Trạng thái']              || '').trim();
  var deadline= formatDate(r['Deadline']             || '');
  var vuong   = String(r['Vướng mắc']               || '').trim();
  var ceoChot = String(r['Nội dung cần CEO chốt']   || '').trim();
  var tienDo  = String(r['Tiến độ']                 || '').trim();
  var rag     = calcRAG_(status, deadline, vuong, ceoChot);
  return {
    MaBanGhi:      String(r['Tên dự án']            || '').trim(),
    TenBanGhi:     String(r['Tên dự án']            || '').trim(),
    TenDuAn:       String(r['Tên dự án']            || '').trim(),
    MoTa:          String(r['Mô tả ngắn']           || '').trim(),
    NguoiPhuTrach: String(r['Người phụ trách']      || '').trim(),
    NguoiPhoiHop:  String(r['Người phối hợp']       || '').trim(),
    UuTien:        String(r['Mức ưu tiên']          || '').trim(),
    Deadline:      deadline,
    TrangThai:     status,
    TienDo:        tienDo,
    PctHT:         tienDoToPct_(tienDo),
    CapNhat:       String(r['Cập nhật mới nhất']    || '').trim(),
    VuongMac:      vuong,
    ViecTiepTheo:  String(r['Việc tiếp theo']       || '').trim(),
    CEOCanChot:    ceoChot,
    QuyetDinhCEO:  String(r['Quyết định của CEO']   || '').trim(),
    Link:          String(r['Link tài liệu']         || '').trim(),
    RAG:           rag,
    IsOverdue:     isPastDeadline(deadline, status),
    // Legacy aliases
    ProjectName:   String(r['Tên dự án']            || '').trim(),
    PICLead:       String(r['Người phụ trách']      || '').trim(),
    Owner:         String(r['Người phụ trách']      || '').trim(),
    Status:        status,
    Issues:        vuong,
    CEODecision:   ceoChot,
    CEONote:       ceoChot,
    Progress:      tienDo,
    Description:   String(r['Mô tả ngắn']           || '').trim()
  };
}

/* ── Normalizer: 03_TO_COMMAND_CENTER ── */
function normalizeTO_(r) {
  var status  = String(r['Trạng thái']             || '').trim();
  var deadline= formatDate(r['Deadline']            || '');
  var vuong   = String(r['Vướng mắc']              || '').trim();
  var ceoChot = String(r['Nội dung cần CEO chốt']  || '').trim();
  var tienDo  = String(r['Tiến độ']                || '').trim();
  var loai    = String(r['Loại nội dung']          || '').trim();
  var tieuDe  = String(r['Tiêu đề']               || '').trim();
  var noiDung = String(r['Nội dung cần xử lý']     || '').trim();
  var nhaMay  = String(r['Nhà máy']               || '').trim();
  var nguoiPT = String(r['Người phụ trách']        || '').trim();
  var rag     = calcRAG_(status, deadline, vuong, ceoChot);
  return {
    LoaiND:          loai,
    NgayGhiNhan:     formatDate(r['Ngày ghi nhận'] || ''),
    NhaMay:          nhaMay,
    NhomVanDe:       String(r['Nhóm vấn đề']       || '').trim(),
    TieuDe:          tieuDe,
    NoiDungXuLy:     noiDung,
    NguoiPhuTrach:   nguoiPT,
    NguoiPhoiHop:    String(r['Người phối hợp']    || '').trim(),
    Deadline:        deadline,
    TrangThai:       status,
    TienDo:          tienDo,
    PctHT:           tienDoToPct_(tienDo),
    VuongMac:        vuong,
    ViecTiepTheo:    String(r['Việc tiếp theo']     || '').trim(),
    CEOCanChot:      ceoChot,
    KetQua:          String(r['Kết quả']            || '').trim(),
    CuocHopLienQuan: String(r['Cuộc họp liên quan']|| '').trim(),
    Link:            String(r['Link bằng chứng']   || '').trim(),
    CapNhat:         String(r['Cập nhật mới nhất'] || '').trim(),
    RAG:             rag,
    IsOverdue:       isPastDeadline(deadline, status),
    // Legacy aliases cho frontend
    IssueID:    '',
    ActionID:   '',
    Factory:    nhaMay,
    Area:       nhaMay,
    Description:tieuDe + (noiDung ? ' — ' + noiDung : ''),
    Action:     noiDung || tieuDe,
    NoiDung:    noiDung || tieuDe,
    PIC:        nguoiPT,
    Owner:      nguoiPT,
    Status:     status,
    Level:      String(r['Nhóm vấn đề'] || '').trim(),
    Reporter:   nguoiPT,
    Loai:       loai,
    NguoiGhi:   nguoiPT,
    Ngay:       formatDate(r['Ngày ghi nhận'] || ''),
    TrangThaiXuLy: status,
    CEONote:    ceoChot,
    CEODecision:ceoChot,
    MeetingIDRef: String(r['Cuộc họp liên quan'] || '').trim()
  };
}

/* ── Normalizer: 04_BIEN_BAN_HOP ── */
function normalizeBienBan_(r) {
  var status   = String(r['Trạng thái']            || 'Nháp').trim();
  var deadline = formatDate(r['Deadline']           || '');
  var ceoChot  = String(r['Nội dung cần CEO chốt'] || '').trim();
  var tenHop   = String(r['Tên cuộc họp']          || '').trim();
  var actionChinh = String(r['Action chính (tóm tắt)'] || '').trim();
  return {
    NgayHop:     formatDate(r['Ngày họp']           || ''),
    TenHop:      tenHop,
    LoaiHop:     String(r['Loại họp']              || '').trim(),
    ChuTri:      String(r['Chủ trì']               || '').trim(),
    ThanhPhan:   String(r['Thành phần tham dự']    || '').trim(),
    TomTat:      String(r['Tóm tắt chính']         || '').trim(),
    KetLuan:     String(r['Kết luận chính']        || '').trim(),
    ActionChinh: actionChinh,
    NguoiAction: String(r['Người phụ trách action']|| '').trim(),
    Deadline:    deadline,
    CEOCanChot:  ceoChot,
    Link:        String(r['Link biên bản']          || '').trim(),
    TrangThai:   status,
    // Legacy aliases cho renderMeetingCenter
    MeetingID:   tenHop,
    Date:        formatDate(r['Ngày họp'] || ''),
    Topic:       tenHop,
    TieuDe:      tenHop,
    Chair:       String(r['Chủ trì'] || '').trim(),
    Attendees:   String(r['Thành phần tham dự'] || '').trim(),
    Summary:     String(r['Tóm tắt chính'] || '').trim(),
    CEODirective:ceoChot,
    CEODecision: ceoChot,
    FileMinutes: String(r['Link biên bản'] || '').trim(),
    FileBienBan: String(r['Link biên bản'] || '').trim(),
    Status:      status,
    Actions:     actionChinh,
    HasCEO:      hasText(ceoChot),
    HasAction:   hasText(actionChinh)
  };
}

/* ── Normalizer: 05_RISK_SOS ── */
function normalizeRiskRow_(r) {
  var status  = String(r['Trạng thái']             || '').trim();
  var deadline= formatDate(r['Deadline']            || '');
  var ceoChot = String(r['Nội dung cần CEO chốt']  || '').trim();
  var mucDo   = String(r['Mức độ']                 || '').trim();
  var vuong   = String(r['Hướng xử lý']            || '').trim();
  var rag = /khẩn cấp/i.test(mucDo) ? 'Red'
           : /cao/i.test(mucDo) ? 'Red'
           : /trung bình/i.test(mucDo) ? 'Amber' : 'Green';
  if (isPastDeadline(deadline, status) || hasText(ceoChot)) rag = 'Red';
  return {
    NgayGhiNhan:   formatDate(r['Ngày ghi nhận']     || ''),
    Module:        String(r['Module']                 || '').trim(),
    NhaMay:        String(r['Nhà máy']               || '').trim(),
    NoiDungRuiRo:  String(r['Nội dung rủi ro']       || '').trim(),
    MucDo:         mucDo,
    AnhHuong:      String(r['Ảnh hưởng']            || '').trim(),
    NguoiPhuTrach: String(r['Người phụ trách']       || '').trim(),
    HuongXuLy:     vuong,
    Deadline:      deadline,
    TrangThai:     status,
    CEOCanChot:    ceoChot,
    KetQua:        String(r['Kết quả']               || '').trim(),
    CapNhat:       String(r['Cập nhật mới nhất']     || '').trim(),
    RAG:           rag,
    IsOverdue:     isPastDeadline(deadline, status),
    IsOpen:        !/đã xử lý|đóng|done/i.test(status),
    // Legacy aliases
    Description:   String(r['Nội dung rủi ro']  || '').trim(),
    Level:         mucDo,
    PIC:           String(r['Người phụ trách']   || '').trim(),
    Action:        vuong,
    Status:        status,
    CEONote:       ceoChot,
    CEODecision:   ceoChot
  };
}

/* ── Normalizer: 06_QUYET_DINH_CEO ── */
function normalizeQuyetDinh_(r) {
  var status  = String(r['Trạng thái']                  || '').trim();
  var deadline= formatDate(r['Deadline cần quyết định'] || '');
  return {
    NgayDeXuat:    formatDate(r['Ngày đề xuất']           || ''),
    Module:        String(r['Module']                     || '').trim(),
    NoiDung:       String(r['Nội dung cần quyết định']    || '').trim(),
    BoiCanh:       String(r['Bối cảnh ngắn']             || '').trim(),
    PhuongAn:      String(r['Phương án đề xuất']         || '').trim(),
    NguoiDeXuat:   String(r['Người đề xuất']             || '').trim(),
    Deadline:      deadline,
    QuyetDinhCEO:  String(r['Quyết định của CEO']        || '').trim(),
    NgayQuyet:     formatDate(r['Ngày quyết định']        || ''),
    TrangThai:     status,
    GhiChu:        String(r['Ghi chú']                   || '').trim(),
    // Legacy
    Content:       String(r['Nội dung cần quyết định'] || '').trim(),
    Decision:      String(r['Quyết định của CEO'] || '').trim(),
    Proposer:      String(r['Người đề xuất'] || '').trim(),
    Status:        status,
    IsPending:     /chờ|pending/i.test(status)
  };
}

/* ── Normalizer: 08_LUONG_CHAT_CHU_TICH ── */
function normalizeLuongChat_(r) {
  var status  = String(r['Trạng thái']              || '').trim();
  var ceoChot = String(r['Nội dung cần CEO chốt']   || '').trim();
  return {
    Ngay:          formatDate(r['Ngày']              || ''),
    ChuDe:         String(r['Chủ đề']              || '').trim(),
    TomTat:        String(r['Tóm tắt']             || '').trim(),
    ChiDao:        String(r['Chỉ đạo của CEO']     || '').trim(),
    NguoiPhuTrach: String(r['Người phụ trách']     || '').trim(),
    Deadline:      formatDate(r['Deadline']          || ''),
    ActionPhatSinh:String(r['Action phát sinh']    || '').trim(),
    TrangThai:     status,
    CEOCanChot:    ceoChot,
    Link:          String(r['Link file']             || '').trim(),
    // Legacy
    Status:        status,
    HasAction:     hasText(String(r['Action phát sinh'] || '')),
    CEONote:       ceoChot
  };
}

/* =====================================================
   DATA GETTERS — 1 HÀM = 1 MODULE = 1 SHEET
===================================================== */

function getPMOModuleData() {
  var rows    = readSheetSafe('01_PMO_GSP_NEXT30').map(normalizePMO_);
  var open    = rows.filter(function(r){ return !/hoàn thành|done/i.test(r.TrangThai); });
  var overdue = rows.filter(function(r){ return r.IsOverdue; });
  var ceoList = rows.filter(function(r){ return hasText(r.CEOCanChot); });
  var ragSum  = { Green: 0, Amber: 0, Red: 0 };
  rows.forEach(function(r){ var g = r.RAG||'Green'; if(ragSum[g]!==undefined) ragSum[g]++; });
  var byPIC = {};
  rows.forEach(function(r) {
    var p = r.PIC || 'Khác';
    if (!byPIC[p]) byPIC[p] = { name: p, total: 0, open: 0, overdue: 0, red: 0, ceo: 0 };
    byPIC[p].total++;
    if (!/hoàn thành|done/i.test(r.TrangThai)) byPIC[p].open++;
    if (r.IsOverdue) byPIC[p].overdue++;
    if (r.RAG === 'Red') byPIC[p].red++;
    if (hasText(r.CEOCanChot)) byPIC[p].ceo++;
  });
  return {
    records:    rows,
    open:       open,
    overdue:    overdue,
    ceoList:    ceoList,
    byPIC:      Object.keys(byPIC).map(function(k){ return byPIC[k]; }),
    summary:    { total: rows.length, open: open.length, overdue: overdue.length, ceo: ceoList.length, byRAG: ragSum },
    pmoActions: rows
  };
}

function getDuAnCEOModuleData() {
  var rawRows = readSheetSafe('02_DU_AN_CEO');

  // Auto-detect schema: nếu có cột 'Loại bản ghi' → extended schema (sau migrate)
  if (rawRows.length > 0 && rawRows[0].hasOwnProperty('Loại bản ghi')) {
    return getDuAnCEO02Extended_(rawRows);
  }

  // Flat schema cũ (14 cột) — bổ sung exists + summary fields đúng tên cho cptRender
  var rows    = rawRows.map(normalizeDuAnCEO_);
  var active  = rows.filter(function(r){ return !/hoàn thành|done/i.test(r.TrangThai); });
  var overdue = rows.filter(function(r){ return r.IsOverdue; });
  var ceoList = rows.filter(function(r){ return hasText(r.CEOCanChot); });
  var risk    = rows.filter(function(r){ return hasText(r.VuongMac); });
  var noPT    = rows.filter(function(r){ return !r.NguoiPhuTrach; });
  var ragSum  = { Green: 0, Amber: 0, Red: 0 };
  rows.forEach(function(r){ var g = r.RAG || 'Green'; if (ragSum[g] !== undefined) ragSum[g]++; });
  return {
    exists:      rows.length > 0,
    records:     rows,
    active:      active,
    overdue:     overdue,
    ceoList:     ceoList,
    hangMucs:    [],
    congViecs:   [],
    capNhats:    [],
    quyetDinhs:  [],
    summary: {
      tongDuAn:             rows.length,
      dangTrienKhai:        active.length,
      dungTienDo:           ragSum.Green,
      coNguyChamTre:        ragSum.Amber,
      quaHan:               overdue.length,
      chuaCoNguoiPhuTrach:  noPT.length,
      coRuiRo:              risk.length,
      canCEOQuyetDinh:      ceoList.length,
      total: rows.length, active: active.length,
      overdue: overdue.length, ceo: ceoList.length, byRAG: ragSum
    },
    ceoProjects: rows,
    projects:    rows
  };
}

function getTOModuleData() {
  var now  = new Date();
  var rows = readSheetSafe('03_TO_COMMAND_CENTER').map(normalizeTO_);
  var issues    = rows.filter(function(r){ return r.LoaiND==='Vấn đề'||r.LoaiND==='Vấn đề nóng'; });
  var hotIssues = rows.filter(function(r){ return r.LoaiND==='Vấn đề nóng'||r.RAG==='Red'; });
  var allActs   = rows.filter(function(r){ return r.LoaiND==='Action'||r.LoaiND==='Việc cần làm rõ'; });
  var openActs = [], overdueActs = [], doneActs = [];
  allActs.forEach(function(a) {
    if (/hoàn thành|done|xong/i.test(a.TrangThai)) doneActs.push(a);
    else if (a.IsOverdue) overdueActs.push(a);
    else openActs.push(a);
  });
  var factorySet = {};
  rows.forEach(function(r){ if(r.NhaMay) factorySet[r.NhaMay]=true; });
  var factories = Object.keys(factorySet).sort();
  // Nhóm action theo cuộc họp liên quan
  var meetingMap = {};
  allActs.forEach(function(a) {
    var key = a.CuocHopLienQuan || '';
    if (!key) return;
    if (!meetingMap[key]) meetingMap[key] = { MeetingID: key, TieuDe: key, NgayHop: '', TrangThai: '', CanCEO: '', actions: [] };
    meetingMap[key].actions.push(a);
  });
  var meetings = Object.keys(meetingMap).map(function(k){ return meetingMap[k]; });
  var ceoList  = rows.filter(function(r){ return hasText(r.CEOCanChot); });
  var ragSum   = { Green: 0, Amber: 0, Red: 0 };
  issues.forEach(function(r){ var g=r.RAG||'Amber'; if(ragSum[g]!==undefined) ragSum[g]++; else ragSum.Amber++; });
  return JSON.parse(JSON.stringify({
    updatedAt:       Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'),
    factories:       factories,
    issues:          issues,
    hotIssues:       hotIssues,
    meetings:        meetings,
    topics:          [],
    actions:         allActs,
    openActions:     openActs,
    overdueActions:  overdueActs,
    doneActions:     doneActs,
    ceoComments:     [],
    ceoPending:      [],
    ceoList:         ceoList,
    noOwnerList:     rows.filter(function(r){ return !r.NguoiPhuTrach; }),
    issuesSummary:   { total: issues.length, open: issues.filter(function(r){ return !/hoàn thành/i.test(r.TrangThai); }).length, byRAG: ragSum, hot: hotIssues.length },
    actionsSummary:  { total: allActs.length, open: openActs.length, overdue: overdueActs.length, done: doneActs.length, ceo: ceoList.length, noOwner: rows.filter(function(r){ return !r.NguoiPhuTrach; }).length }
  }));
}

function getBienBanHopModuleData() {
  var rows    = readSheetSafe('04_BIEN_BAN_HOP').map(normalizeBienBan_);
  var pending = rows.filter(function(r){ return /nháp/i.test(r.TrangThai); });
  var hasCEO  = rows.filter(function(r){ return hasText(r.CEOCanChot); });
  var hasAct  = rows.filter(function(r){ return hasText(r.ActionChinh); });
  return {
    records:   rows,
    pending:   pending,
    hasCEO:    hasCEO,
    hasAction: hasAct,
    summary:   { total: rows.length, pending: pending.length, hasCEO: hasCEO.length, hasAction: hasAct.length },
    minutes:   rows,
    topics:    [],
    actions:   []
  };
}

function getRiskModuleData() {
  var rows  = readSheetSafe('05_RISK_SOS').map(normalizeRiskRow_);
  var open  = rows.filter(function(r){ return r.IsOpen; });
  var red   = rows.filter(function(r){ return r.RAG==='Red'; });
  var ceo   = rows.filter(function(r){ return hasText(r.CEOCanChot); });
  return {
    records: rows,
    open:    open,
    red:     red,
    ceoList: ceo,
    summary: { total: rows.length, open: open.length, red: red.length, ceo: ceo.length },
    riskSos: rows
  };
}

function getQuyetDinhModuleData() {
  var rows    = readSheetSafe('06_QUYET_DINH_CEO').map(normalizeQuyetDinh_);
  var pending = rows.filter(function(r){ return r.IsPending; });
  return {
    records:      rows,
    pending:      pending,
    summary:      { total: rows.length, pending: pending.length },
    decisions:    rows,
    ceoDecisions: rows
  };
}

function getLuongChatModuleData() {
  var rows = readSheetSafe('08_LUONG_CHAT_CHU_TICH').map(normalizeLuongChat_);
  var open = rows.filter(function(r){ return !/đã xử lý/i.test(r.TrangThai); });
  var ceo  = rows.filter(function(r){ return hasText(r.CEOCanChot); });
  return {
    records: rows,
    open:    open,
    ceoList: ceo,
    summary: { total: rows.length, open: open.length, ceo: ceo.length }
  };
}

/* ── Dashboard tổng hợp: đọc từ 7 module ── */
function getNewDashboardData() {
  var pmo     = getPMOModuleData();
  var duAn    = getDuAnCEOModuleData();
  var to      = getTOModuleData();
  var bienBan = getBienBanHopModuleData();
  var risk    = getRiskModuleData();
  var quyet   = getQuyetDinhModuleData();
  var chat    = getLuongChatModuleData();
  var ceoTotal= (to.ceoList||[]).length + (duAn.ceoList||[]).length
              + (pmo.ceoList||[]).length + (risk.ceoList||[]).length
              + (quyet.pending||[]).length + (chat.ceoList||[]).length;
  return {
    pmo:            pmo,
    duAnCEO:        duAn,
    to:             to,
    bienBan:        bienBan,
    risk:           risk,
    quyetDinh:      quyet,
    luongChat:      chat,
    ceoTotal:       ceoTotal,
    // Legacy fields cho dashboard renderer cũ
    pmoActions:     pmo.pmoActions     || [],
    ceoProjects:    duAn.ceoProjects   || [],
    toActions:      to.actions         || [],
    meetingMinutes: bienBan.minutes    || [],
    ceoDecisions:   quyet.decisions    || [],
    riskSos:        risk.riskSos       || [],
    meetingTopics:  [],
    masterPIC:      [],
    masterStatus:   []
  };
}

/* =====================================================
   WRITE HANDLERS — GHI VÀO SHEET MỚI (KHÔNG ID)
===================================================== */

/* ── Ghi vào 03_TO_COMMAND_CENTER (issue + action) ── */
function writeTO_(ss, payload) {
  try {
    var sheet = ss.getSheetByName('03_TO_COMMAND_CENTER');
    if (!sheet) return { success: false, error: 'Sheet 03_TO_COMMAND_CENTER chưa tồn tại. Vui lòng chạy setupNewArchitecture_() từ Apps Script editor.' };
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var data = {
      'Loại nội dung':         String(payload.loaiND || 'Action').trim(),
      'Ngày ghi nhận':         String(payload.ngay || payload.reportDate || today).trim(),
      'Nhà máy':               String(payload.nhaMay || payload.factory || '').trim(),
      'Nhóm vấn đề':           String(payload.nhomVanDe || payload.issueType || payload.workstream || '').trim(),
      'Tiêu đề':               String(payload.tieuDe || payload.description || payload.issueTitle || '').trim(),
      'Nội dung cần xử lý':    String(payload.noiDung || payload.action || payload.rootCause || payload.context || '').trim(),
      'Người phụ trách':       String(payload.nguoiPhuTrach || payload.pic || payload.reporter || '').trim(),
      'Người phối hợp':        String(payload.nguoiPhoiHop || payload.coOwner || '').trim(),
      'Deadline':               String(payload.deadline || '').trim(),
      'Trạng thái':            String(payload.trangThai || payload.status || 'Chưa bắt đầu').trim(),
      'Tiến độ':               String(payload.tienDo || '').trim(),
      'Vướng mắc':             String(payload.vuongMac || '').trim(),
      'Việc tiếp theo':        String(payload.viecTiepTheo || '').trim(),
      'Nội dung cần CEO chốt': String(payload.ceoCanChot || payload.ceoNote || payload.ceoDecision || '').trim(),
      'Kết quả':               String(payload.ketQua || payload.output || '').trim(),
      'Cuộc họp liên quan':    String(payload.cuocHop || payload.meetingIdRef || payload.meeting || '').trim(),
      'Link bằng chứng':       String(payload.link || payload.fileSource || payload.bangChung || '').trim(),
      'Cập nhật mới nhất':     String(payload.capNhat || '').trim()
    };
    if (!data['Tiêu đề'] && !data['Nội dung cần xử lý']) {
      return { success: false, error: 'Cần nhập ít nhất Tiêu đề hoặc Nội dung.' };
    }
    appendRowByHeaders_(sheet, data);
    SpreadsheetApp.flush();
    return { success: true, message: 'Đã ghi nhận thành công vào 03_TO_COMMAND_CENTER.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* ── Ghi vào 06_QUYET_DINH_CEO (thay thế CEO_COMMENTS) ── */
function writeQuyetDinh_(ss, payload) {
  try {
    var sheet = ss.getSheetByName('06_QUYET_DINH_CEO');
    if (!sheet) return { success: false, error: 'Sheet 06_QUYET_DINH_CEO chưa tồn tại. Vui lòng chạy setupNewArchitecture_() từ Apps Script editor.' };
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var data = {
      'Ngày đề xuất':            String(payload.ngay || today).trim(),
      'Module':                  String(payload.module || payload.loai || 'TO').trim(),
      'Nội dung cần quyết định': String(payload.noiDung || payload.content || '').trim(),
      'Bối cảnh ngắn':           String(payload.boiCanh || payload.meetingIdRef || '').trim(),
      'Phương án đề xuất':       String(payload.phuongAn || '').trim(),
      'Người đề xuất':           String(payload.nguoiDeXuat || payload.nguoiGhi || '').trim(),
      'Deadline cần quyết định': String(payload.deadline || '').trim(),
      'Quyết định của CEO':      String(payload.quyetDinh || '').trim(),
      'Ngày quyết định':         String(payload.ngayQuyet || '').trim(),
      'Trạng thái':              String(payload.trangThaiXuLy || payload.trangThai || 'Chờ quyết định').trim(),
      'Ghi chú':                 String(payload.ghiChu || payload.issueIdRef || '').trim()
    };
    if (!data['Nội dung cần quyết định']) {
      return { success: false, error: 'Cần nhập Nội dung cần quyết định.' };
    }
    appendRowByHeaders_(sheet, data);
    SpreadsheetApp.flush();
    return { success: true, message: 'Đã ghi nhận vào 06_QUYET_DINH_CEO.' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/* ── Bước 5: Tạo 9 sheet mới ── */
function createNewSheetStructure_(ss) {
  var log = [];

  var STATUS_MAIN = ['Chưa bắt đầu', 'Đang thực hiện', 'Gần hoàn thành', 'Hoàn thành', 'Tạm dừng'];
  var UU_TIEN     = ['Cao', 'Trung bình', 'Thấp'];
  var NHA_MAY     = ['GS1', 'GS5', 'GS6', 'Khối văn phòng', 'Toàn hệ thống'];
  var MODULE_LIST = ['PMO', 'Dự án CEO', 'TO', 'Risk', 'Biên bản họp', 'Chung'];
  var MUC_DO_RISK = ['Thấp', 'Trung bình', 'Cao', 'Khẩn cấp'];
  var PIC_LIST    = [
    'PIC01 - Kinh Doanh', 'PIC02 - Sales Tech', 'PIC03 - SCM',
    'PIC04 - Mua Hàng',   'PIC05 - Công Nghệ',  'PIC06 - Cơ Điện/QLTB',
    'PIC07 - IT/ERP',     'PIC08 - QA',          'PIC09 - P&C',
    'PIC10 - Giám Sát Tuân Thủ', 'PIC11 - Kiểm Soát Nội Bộ',
    'PIC12 - HCNS', 'PIC13 - Tài Chính/Kế Toán',
    'PIC14 - GS1',  'PIC15 - GS5', 'PIC16 - GS6'
  ];

  var SHEETS = [

    /* 00 */
    { name: '00_TONG_QUAN',
      color: '#1e3a5f',
      autoNote: true,
      headers: ['Module','Tổng bản ghi đang mở','Vấn đề đỏ (RAG Red)',
                'Việc quá hạn','Cần CEO chốt','Tỷ lệ hoàn thành (%)','Cập nhật gần nhất'],
      widths:  [160, 180, 180, 140, 140, 180, 180],
      dropdowns: {}
    },

    /* 01 */
    { name: '01_PMO_GSP_NEXT30',
      color: '#14532d',
      headers: ['PIC / Trục','Nhóm công việc','Nội dung công việc',
                'Người phụ trách','Deadline','Trạng thái','Tiến độ',
                'Vướng mắc','Việc tiếp theo','CEO cần chốt',
                'Cập nhật mới nhất','Link tài liệu'],
      widths:  [180, 160, 300, 160, 100, 150, 150,
                220, 220, 220, 240, 180],
      dropdowns: {
        0: PIC_LIST,
        1: ['Sản xuất','Kinh doanh','Kỹ thuật','Hệ thống','HCNS','Tài chính','Khác'],
        5: STATUS_MAIN,
        6: STATUS_MAIN
      }
    },

    /* 02 */
    { name: '02_DU_AN_CEO',
      color: '#7c2d12',
      headers: ['Tên dự án','Mô tả ngắn','Người phụ trách','Người phối hợp',
                'Mức ưu tiên','Deadline','Trạng thái','Tiến độ',
                'Cập nhật mới nhất','Vướng mắc','Việc tiếp theo',
                'Nội dung cần CEO chốt','Quyết định của CEO','Link tài liệu'],
      widths:  [240, 260, 160, 160, 120, 100, 150, 150,
                240, 220, 220, 220, 220, 180],
      dropdowns: {
        4: UU_TIEN,
        6: STATUS_MAIN,
        7: STATUS_MAIN
      }
    },

    /* 03 */
    { name: '03_TO_COMMAND_CENTER',
      color: '#1e3a5f',
      headers: ['Loại nội dung','Ngày ghi nhận','Nhà máy','Nhóm vấn đề',
                'Tiêu đề','Nội dung cần xử lý','Người phụ trách','Người phối hợp',
                'Deadline','Trạng thái','Tiến độ','Vướng mắc',
                'Việc tiếp theo','Nội dung cần CEO chốt','Kết quả',
                'Cuộc họp liên quan','Link bằng chứng','Cập nhật mới nhất'],
      widths:  [150, 110, 100, 150, 220, 280, 160, 160,
                100, 150, 150, 220, 220, 220, 220, 220, 180, 220],
      dropdowns: {
        0: ['Vấn đề', 'Action', 'Vấn đề nóng', 'Việc cần làm rõ'],
        2: NHA_MAY,
        3: ['Sản xuất','Kinh doanh','Nhân sự','Kỹ thuật','An toàn','Chất lượng','Khác'],
        9: STATUS_MAIN,
        10: STATUS_MAIN
      }
    },

    /* 04 */
    { name: '04_BIEN_BAN_HOP',
      color: '#064e3b',
      headers: ['Ngày họp','Tên cuộc họp','Loại họp','Chủ trì',
                'Thành phần tham dự','Tóm tắt chính','Kết luận chính',
                'Action chính (tóm tắt)','Người phụ trách action','Deadline',
                'Nội dung cần CEO chốt','Link biên bản','Trạng thái'],
      widths:  [100, 240, 150, 150, 220, 280, 280, 220, 160, 100, 220, 180, 110],
      dropdowns: {
        2: ['Họp Ban TO','Họp dự án','Họp CEO','Họp định kỳ','Khác'],
        12: ['Nháp','Đã ký','Đã gửi']
      }
    },

    /* 05 */
    { name: '05_RISK_SOS',
      color: '#7f1d1d',
      headers: ['Ngày ghi nhận','Module','Nhà máy','Nội dung rủi ro',
                'Mức độ','Ảnh hưởng','Người phụ trách','Hướng xử lý',
                'Deadline','Trạng thái','Nội dung cần CEO chốt',
                'Kết quả','Cập nhật mới nhất'],
      widths:  [110, 140, 100, 280, 120, 220, 160, 220, 100, 150, 220, 220, 220],
      dropdowns: {
        1: MODULE_LIST,
        2: NHA_MAY,
        4: MUC_DO_RISK,
        9: ['Mới','Đang xử lý','Đã xử lý','Đóng']
      }
    },

    /* 06 */
    { name: '06_QUYET_DINH_CEO',
      color: '#4c1d95',
      headers: ['Ngày đề xuất','Module','Nội dung cần quyết định',
                'Bối cảnh ngắn','Phương án đề xuất','Người đề xuất',
                'Deadline cần quyết định','Quyết định của CEO',
                'Ngày quyết định','Trạng thái','Ghi chú'],
      widths:  [110, 140, 280, 220, 240, 160, 160, 280, 110, 180, 180],
      dropdowns: {
        1: MODULE_LIST,
        9: ['Chờ quyết định','Đã quyết','Đang triển khai','Hoàn thành']
      }
    },

    /* 07 */
    { name: '07_CHAT_LUONG_DU_LIEU',
      color: '#374151',
      autoNote: true,
      headers: ['Module','Tổng bản ghi','Thiếu người phụ trách',
                'Thiếu Deadline','Quá hạn','Chưa cập nhật >7 ngày',
                'Điểm chất lượng (%)','Cập nhật gần nhất'],
      widths:  [160, 130, 180, 130, 100, 200, 170, 180],
      dropdowns: {}
    },

    /* 08 */
    { name: '08_LUONG_CHAT_CHU_TICH',
      color: '#78350f',
      headers: ['Ngày','Chủ đề','Tóm tắt','Chỉ đạo của CEO',
                'Người phụ trách','Deadline','Action phát sinh',
                'Trạng thái','Nội dung cần CEO chốt','Link file'],
      widths:  [100, 220, 260, 300, 160, 100, 220, 120, 220, 180],
      dropdowns: {
        7: ['Mới','Đang xử lý','Đã xử lý']
      }
    }

  ]; // end SHEETS

  SHEETS.forEach(function(cfg) {
    /* Bỏ qua nếu đã tồn tại */
    if (ss.getSheetByName(cfg.name)) {
      log.push('  ⬜ Đã tồn tại: ' + cfg.name + ' — bỏ qua');
      return;
    }

    var sheet = ss.insertSheet(cfg.name);

    /* ── Header ── */
    var nCols       = cfg.headers.length;
    var headerRange = sheet.getRange(1, 1, 1, nCols);
    headerRange.setValues([cfg.headers])
               .setBackground(cfg.color)
               .setFontColor('#ffffff')
               .setFontWeight('bold')
               .setFontSize(11)
               .setWrap(true)
               .setVerticalAlignment('middle')
               .setHorizontalAlignment('center');
    sheet.setRowHeight(1, 40);
    sheet.setFrozenRows(1);

    /* ── Màu tab ── */
    sheet.setTabColor(cfg.color);

    /* ── Độ rộng cột ── */
    if (cfg.widths) {
      cfg.widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
    }

    /* ── Dropdown (data validation rows 2-1000) ── */
    Object.keys(cfg.dropdowns).forEach(function(idxStr) {
      var colIdx  = parseInt(idxStr, 10);
      var options = cfg.dropdowns[idxStr];
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, colIdx + 1, 999, 1).setDataValidation(rule);
    });

    /* ── Ghi chú cho sheet tự động ── */
    if (cfg.autoNote) {
      sheet.getRange(1, 1).setNote('⚙️ Sheet này do Apps Script tự cập nhật.\nKhông chỉnh sửa tay.');
    }

    log.push('  ✅ Đã tạo: ' + cfg.name + '  (' + nCols + ' cột)');
  });

  return log;
}

/* =====================================================================
   MIGRATION: GSP Project Workbook → 02_DU_AN_CEO (extended schema)
   Schema mở rộng v2 — hỗ trợ phân cấp Dự án / Hạng mục / Công việc
   ===================================================================== */

/* ── Đọc sheet có header không nằm ở hàng 1 ─────────────────────────
   sheetName : tên sheet trong Spreadsheet hiện tại
   headerRow : số thứ tự hàng header (1-based), ví dụ 9 hoặc 4
   Trả về    : array of objects, bỏ qua hàng trống hoàn toàn
   ─────────────────────────────────────────────────────────────────── */
function readSheetFromHeaderRow_(sheetName, headerRow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    Logger.log('[readSheetFromHeaderRow_] Sheet không tìm thấy: ' + sheetName);
    return [];
  }
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < headerRow + 1 || lastCol < 1) return [];
  var headers    = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var numData    = lastRow - headerRow;
  var dataValues = sh.getRange(headerRow + 1, 1, numData, lastCol).getValues();
  return dataValues.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var key = String(h || '').trim();
      if (key) obj[key] = row[i];
    });
    return obj;
  }).filter(function(obj) {
    return Object.keys(obj).some(function(k) {
      var v = obj[k];
      return v !== '' && v !== null && v !== undefined;
    });
  });
}

/* ── Data getter cho 02_DU_AN_CEO extended schema (sau migrate) ────── */
function getDuAnCEO02Extended_(rawRows) {
  var projects   = [];
  var hangMucs   = [];
  var congViecs  = [];
  var capNhats   = [];
  var quyetDinhs = [];

  rawRows.forEach(function(r) {
    var loai = String(r['Loại bản ghi'] || '').trim();
    if      (loai === 'Dự án')      projects.push(normalizeCEO02DuAn_(r));
    else if (loai === 'Hạng mục')   hangMucs.push(normalizeCEO02HangMuc_(r));
    else if (loai === 'Công việc')  congViecs.push(normalizeCEO02CongViec_(r));
    else if (loai === 'Cập nhật')   capNhats.push(normalizeCEO02DuAn_(r));
    else if (loai === 'Quyết định') quyetDinhs.push(normalizeCEO02DuAn_(r));
  });

  // Enrich: tính số công việc + % hoàn thành cho từng dự án
  var cvByDuAn = {};
  congViecs.forEach(function(cv) {
    var maDuAn = cv.MaDuAn;
    if (!maDuAn) return;
    if (!cvByDuAn[maDuAn]) cvByDuAn[maDuAn] = { total: 0, done: 0 };
    cvByDuAn[maDuAn].total++;
    if (/hoàn thành|done|xong/i.test(cv.TrangThai)) cvByDuAn[maDuAn].done++;
  });

  projects.forEach(function(p) {
    var cv = cvByDuAn[p.MaDuAn] || { total: 0, done: 0 };
    p.SoCongViec  = cv.total;
    p.SoHoanThanh = cv.done;
    if (!p.PhanTramHT && cv.total > 0) {
      p.PhanTramHT = Math.round((cv.done / cv.total) * 100);
      p.PctHT      = p.PhanTramHT;
    }
    p.RAG       = calcRAG_(p.TrangThai, p.ThoiHanHoanThanh, p.VuongMacRuiRo, p.NdCanCEOQuyet);
    p.IsOverdue = isPastDeadline(p.ThoiHanHoanThanh, p.TrangThai);
    p.LatestUpdate = null;
  });

  // Summary
  var active   = projects.filter(function(p){ return !/hoàn thành|done/i.test(p.TrangThai); });
  var overdue  = projects.filter(function(p){ return p.IsOverdue; });
  var ceoList  = projects.filter(function(p){ return hasText(p.NdCanCEOQuyet); });
  var risk     = projects.filter(function(p){ return hasText(p.VuongMacRuiRo); });
  var noPT     = projects.filter(function(p){ return !p.NguoiPhuTrach; });
  var ragGreen = projects.filter(function(p){ return p.RAG === 'Green'; }).length;
  var ragAmber = projects.filter(function(p){ return p.RAG === 'Amber'; }).length;
  var ragSum   = { Green: ragGreen, Amber: ragAmber, Red: projects.length - ragGreen - ragAmber };

  return {
    exists:      projects.length > 0,
    projects:    projects,
    hangMucs:    hangMucs,
    congViecs:   congViecs,
    capNhats:    capNhats,
    quyetDinhs:  quyetDinhs,
    summary: {
      tongDuAn:             projects.length,
      dangTrienKhai:        active.length,
      dungTienDo:           ragGreen,
      coNguyChamTre:        ragAmber,
      quaHan:               overdue.length,
      chuaCoNguoiPhuTrach:  noPT.length,
      coRuiRo:              risk.length,
      canCEOQuyetDinh:      ceoList.length,
      tongHangMuc:          hangMucs.length,
      tongCongViec:         congViecs.length,
      total: projects.length, active: active.length,
      overdue: overdue.length, ceo: ceoList.length, byRAG: ragSum
    },
    ceoProjects: projects,
    records:     projects
  };
}

/* ── Normalizer: dòng "Dự án" trong 02_DU_AN_CEO extended ─────────── */
function normalizeCEO02DuAn_(r) {
  var deadline = String(r['Thời hạn']              || '').trim();
  var status   = String(r['Trạng thái']             || '').trim();
  var vuong    = String(r['Vướng mắc']              || '').trim();
  var ceoChot  = String(r['Nội dung cần CEO chốt']  || '').trim();
  var tienDo   = String(r['Tiến độ']               || '').trim();
  var ten      = String(r['Tên']                   || '').trim();
  var maBanGhi = String(r['Mã bản ghi']            || '').trim();
  var maDuAn   = String(r['Mã dự án']              || maBanGhi || '').trim();
  return {
    LoaiBanGhi:       'Dự án',
    MaBanGhi:         maBanGhi || maDuAn,
    MaDuAn:           maDuAn,
    MaCapTren:        '',
    TenBanGhi:        ten,
    TenDuAn:          ten,
    MoTa:             String(r['Mô tả chi tiết']    || '').trim(),
    NhomViec:         String(r['Nhóm việc']          || '').trim(),
    NguoiPhuTrach:    String(r['Người phụ trách']    || '').trim(),
    NguoiPhoiHop:     String(r['Người phối hợp']     || '').trim(),
    MucDoUuTien:      String(r['Mức ưu tiên']        || '').trim(),
    ThoiHanHoanThanh: deadline,
    Deadline:         deadline,
    TrangThai:        status,
    Status:           status,
    TienDo:           tienDo,
    PhanTramHT:       tienDoToPct_(tienDo),
    PctHT:            tienDoToPct_(tienDo),
    CapNhatMoiNhat:   String(r['Cập nhật mới nhất'] || '').trim(),
    CapNhat:          String(r['Cập nhật mới nhất'] || '').trim(),
    VuongMacRuiRo:    vuong,
    VuongMac:         vuong,
    KeTiepTheo:       String(r['Việc tiếp theo']     || '').trim(),
    ViecTiepTheo:     String(r['Việc tiếp theo']     || '').trim(),
    NdCanCEOQuyet:    ceoChot,
    CEOCanChot:       ceoChot,
    NdCEOCanBiet:     String(r['Ghi chú']            || '').trim(),
    QuyetDinhCEO:     String(r['Quyết định của CEO'] || '').trim(),
    CEODecision:      String(r['Quyết định của CEO'] || '').trim(),
    SanPhamDauRa:     String(r['Đầu ra phải có']     || '').trim(),
    LinkFile:         String(r['Link tài liệu']      || '').trim(),
    Link:             String(r['Link tài liệu']      || '').trim(),
    RAG:              calcRAG_(status, deadline, vuong, ceoChot),
    IsOverdue:        isPastDeadline(deadline, status),
    SoCongViec:       0,
    SoHoanThanh:      0,
    LatestUpdate:     null,
    ProjectName:      ten,
    Owner:            String(r['Người phụ trách']    || '').trim(),
    PICLead:          String(r['Người phụ trách']    || '').trim(),
    Issues:           vuong,
    CEONote:          ceoChot,
    Progress:         tienDo,
    Description:      String(r['Mô tả chi tiết']    || '').trim()
  };
}

/* ── Normalizer: dòng "Hạng mục" ────────────────────────────────── */
function normalizeCEO02HangMuc_(r) {
  var base = normalizeCEO02DuAn_(r);
  base.LoaiBanGhi = 'Hạng mục';
  base.MaCapTren  = String(r['Mã dự án'] || '').trim();
  return base;
}

/* ── Normalizer: dòng "Công việc" ───────────────────────────────── */
function normalizeCEO02CongViec_(r) {
  var deadline = String(r['Thời hạn']              || '').trim();
  var status   = String(r['Trạng thái']             || '').trim();
  var vuong    = String(r['Vướng mắc']              || '').trim();
  var ceoChot  = String(r['Nội dung cần CEO chốt']  || '').trim();
  var ten      = String(r['Tên']                   || '').trim();
  var maBanGhi = String(r['Mã bản ghi']            || '').trim();
  var maDuAn   = String(r['Mã dự án']              || '').trim();
  return {
    LoaiBanGhi:       'Công việc',
    MaBanGhi:         maBanGhi,
    MaDuAn:           maDuAn,
    MaCapTren:        maDuAn,
    TenBanGhi:        ten,
    TenDuAn:          ten,
    MoTa:             String(r['Mô tả chi tiết']    || '').trim(),
    NhomViec:         String(r['Nhóm việc']          || '').trim(),
    NguoiPhuTrach:    String(r['Người phụ trách']    || '').trim(),
    NguoiPhoiHop:     String(r['Người phối hợp']     || '').trim(),
    ThoiHanHoanThanh: deadline,
    Deadline:         deadline,
    TrangThai:        status,
    Status:           status,
    TienDo:           '',
    PhanTramHT:       0,
    PctHT:            0,
    VuongMacRuiRo:    vuong,
    VuongMac:         vuong,
    NdCanCEOQuyet:    ceoChot,
    CEOCanChot:       ceoChot,
    SanPhamDauRa:     String(r['Đầu ra phải có']     || '').trim(),
    NdCEOCanBiet:     String(r['Ghi chú']            || '').trim(),
    LinkFile:         String(r['Link tài liệu']      || '').trim(),
    Link:             String(r['Link tài liệu']      || '').trim(),
    RAG:              calcRAG_(status, deadline, vuong, ceoChot),
    IsOverdue:        isPastDeadline(deadline, status),
    ProjectName:      ten,
    Owner:            String(r['Người phụ trách']    || '').trim(),
    Action:           ten,
    Issues:           vuong
  };
}

/* ── DRY-RUN: Đọc + phân tích hai sheet nguồn ─────────────────────
   Cách chạy: Apps Script editor → chọn hàm migrateGSPProjectWorkbookToCEOTracker → Run
   Tham số  : dryRun = true (mặc định) — chỉ đọc và báo cáo, KHÔNG ghi
              dryRun = false — BỊ KHÓA, chưa implement
   Kết quả  : Xem Execution log (View → Logs) sau khi chạy
   ─────────────────────────────────────────────────────────────────── */
function migrateGSPProjectWorkbookToCEOTracker(dryRun) {
  dryRun = dryRun !== false;  // undefined/true/bất kỳ → true; chỉ false tường minh → false
  var log    = [];
  var errors = [];

  var srcDanhMuc  = readSheetFromHeaderRow_('DANH MỤC DỰ ÁN', 9);
  var srcCongViec = readSheetFromHeaderRow_('CÔNG VIỆC CHI TIẾT', 4);

  if (!srcDanhMuc.length)  errors.push('Sheet "DANH MỤC DỰ ÁN" không tồn tại hoặc không có dữ liệu (header dòng 9).');
  if (!srcCongViec.length) errors.push('Sheet "CÔNG VIỆC CHI TIẾT" không tồn tại hoặc không có dữ liệu (header dòng 4).');
  if (errors.length) {
    var fatal = errors.join('\n');
    Logger.log('[migrateGSPProjectWorkbookToCEOTracker] FATAL:\n' + fatal);
    return { ok: false, dryRun: dryRun, errors: errors, log: fatal };
  }

  // Phân tích DANH MỤC DỰ ÁN
  var projectIdSet = {};
  var dupProjMap   = {};
  var projects     = [];

  srcDanhMuc.forEach(function(r) {
    var ma  = String(r['Mã dự án']      || '').trim();
    var ten = String(r['Tên dự án lớn'] || '').trim();
    if (!ma && !ten) return;
    if (!ma) {
      errors.push('Dự án thiếu Mã dự án: "' + ten.slice(0, 60) + '"');
      return;
    }
    if (projectIdSet[ma]) {
      dupProjMap[ma] = (dupProjMap[ma] || 1) + 1;
      errors.push('Mã dự án trùng: ' + ma);
      return;
    }
    projectIdSet[ma] = true;
    projects.push(r);
  });

  // Phân tích CÔNG VIỆC CHI TIẾT
  var maViecSet       = {};
  var dupMaViecMap    = {};
  var cvThieuMaDuAn   = [];
  var cvKhongTonDuAn  = [];
  var cvThieuNhomViec = [];
  var hangMucMap      = {};

  srcCongViec.forEach(function(r) {
    var maViec   = String(r['Mã việc']   || '').trim();
    var maDuAn   = String(r['Mã dự án']  || '').trim();
    var nhomViec = String(r['Nhóm việc'] || '').trim();
    var tenViec  = String(r['Nội dung giao việc theo nguồn'] || '').slice(0, 60);

    if (!maViec) {
      errors.push('Công việc thiếu Mã việc: "' + tenViec + '"');
    } else if (maViecSet[maViec]) {
      dupMaViecMap[maViec] = (dupMaViecMap[maViec] || 1) + 1;
      errors.push('Mã việc trùng: ' + maViec);
    } else {
      maViecSet[maViec] = true;
    }

    if (!maDuAn) {
      cvThieuMaDuAn.push(maViec || '(thiếu Mã việc)');
    } else if (!projectIdSet[maDuAn]) {
      cvKhongTonDuAn.push((maViec || '?') + ' -> "' + maDuAn + '"');
    }

    if (!nhomViec) {
      cvThieuNhomViec.push(maViec || '(thiếu Mã việc)');
    }

    if (maDuAn && nhomViec) {
      var hmKey = maDuAn + '|' + nhomViec;
      if (!hangMucMap[hmKey]) hangMucMap[hmKey] = { maDuAn: maDuAn, nhomViec: nhomViec, soViec: 0 };
      hangMucMap[hmKey].soViec++;
    }
  });

  var hangMucs      = [];
  for (var k in hangMucMap) { if (hangMucMap.hasOwnProperty(k)) hangMucs.push(hangMucMap[k]); }
  var dupProjList   = Object.keys(dupProjMap);
  var dupMaViecList = Object.keys(dupMaViecMap);

  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
  var fmtList = function(arr, cap, icon) {
    if (!arr.length) return String.fromCharCode(9989) + ' Khong co';
    return (icon || String.fromCharCode(10060)) + ' ' + arr.length + ' -- ' + arr.slice(0, cap).join(', ') + (arr.length > cap ? ' ...' : '');
  };

  log.push('================================================================');
  log.push('  DRY-RUN: migrateGSPProjectWorkbookToCEOTracker');
  log.push('  ' + ts);
  log.push('================================================================');
  log.push('');
  log.push('THONG KE NGUON:');
  log.push('  Tong du an (DANH MUC DU AN):           ' + projects.length + ' (/' + srcDanhMuc.length + ' dong)');
  log.push('  Tong hang muc tu suy ra:               ' + hangMucs.length + ' to hop duy nhat (Ma DA + Nhom viec)');
  log.push('  Tong cong viec (CONG VIEC CHI TIET):   ' + srcCongViec.length);
  log.push('  Tong ban ghi se ghi vao 02_DU_AN_CEO:  ' + (projects.length + hangMucs.length + srcCongViec.length));
  log.push('');
  log.push('KIEM TRA MA DU AN:');
  log.push('  Ma du an trung:                ' + fmtList(dupProjList, 10, 'LOI'));
  log.push('');
  log.push('KIEM TRA CONG VIEC (' + srcCongViec.length + ' ban ghi):');
  log.push('  Ma viec trung:                 ' + fmtList(dupMaViecList, 5, 'LOI'));
  log.push('  Thieu Ma du an:                ' + fmtList(cvThieuMaDuAn, 5, 'LOI'));
  log.push('  Ma du an khong ton tai:        ' + fmtList(cvKhongTonDuAn, 5, 'LOI'));
  log.push('  Thieu Nhom viec:               ' + fmtList(cvThieuNhomViec, 5, 'CANH BAO'));
  log.push('');

  if (errors.length) {
    log.push('LOI CAN XU LY (' + errors.length + ' loi):');
    errors.forEach(function(e, i) { log.push('  [' + (i + 1) + '] ' + e); });
  } else {
    log.push('LOI: OK - Khong co loi. Du lieu sach - san sang de migrate.');
  }

  log.push('');
  log.push('CHE DO: DRY-RUN - Khong co gi duoc ghi vao 02_DU_AN_CEO.');
  log.push('================================================================');

  var report = log.join('\n');
  Logger.log(report);

  if (!dryRun) {
    throw new Error('Write mode chua duoc kich hoat. Chay dry-run truoc, xac nhan ket qua, roi moi enable write mode.');
  }

  return {
    ok:      errors.length === 0,
    dryRun:  true,
    stats: {
      tongDuAn:    projects.length,
      tongHangMuc: hangMucs.length,
      tongCongViec:srcCongViec.length,
      tongBanGhi:  projects.length + hangMucs.length + srcCongViec.length
    },
    issues: {
      dupProjectIds:  dupProjList,
      dupMaViec:      dupMaViecList,
      cvThieuMaDuAn:  cvThieuMaDuAn,
      cvKhongTonDuAn: cvKhongTonDuAn,
      cvThieuNhomViec:cvThieuNhomViec
    },
    errors: errors,
    log:    report
  };
}

/* ── Wrapper không tham số — chọn hàm này trong Apps Script editor để chạy dry-run ── */
function testMigrationDryRun() {
  var result = migrateGSPProjectWorkbookToCEOTracker(true);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/* =====================================================
   MEETING HUB — getMeetingHubData() (Legacy compat)
   Đọc từ 3 sheets hiện có:
     MEETING_TOPICS, MEETING_MINUTES, MEETING_ACTIONS
   Map sang data model mới cho frontend mh* namespace.
===================================================== */
function getMeetingHubDataLegacy_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  /* ── Đọc MEETING_TOPICS ── */
  var rawTopics = [];
  try { rawTopics = readSheetSafe('MEETING_TOPICS'); } catch(e) {}
  var topics = rawTopics.map(function(r) {
    return {
      topicId:   String(r['Topic ID'] || r['TopicID'] || r['Mã chủ đề'] || '').trim(),
      topicName: String(r['Topic Name'] || r['TopicName'] || r['Tên chủ đề'] || '').trim(),
      module:    String(r['Module'] || r['Nhóm'] || '').trim(),
      desc:      String(r['Description'] || r['Mô tả'] || r['Desc'] || '').trim(),
      meetingCount:    0,
      openActionCount: 0
    };
  }).filter(function(t) { return t.topicId || t.topicName; });

  /* ── Đọc MEETING_MINUTES ── */
  var rawMinutes = [];
  try { rawMinutes = readSheetSafe('MEETING_MINUTES'); } catch(e) {}
  var today = new Date();
  var meetings = rawMinutes.map(function(r) {
    var dateStr = String(r['Date'] || r['Ngày họp'] || r['Ngày'] || '').trim();
    return {
      meetingId:    String(r['Meeting ID'] || r['MeetingID'] || r['Mã biên bản'] || '').trim(),
      meetingTitle: String(r['Topic'] || r['Chủ đề'] || r['Tiêu đề'] || '').trim(),
      date:         dateStr,
      type:         String(r['Type'] || r['Loại họp'] || r['Loại'] || '').trim(),
      chair:        String(r['Chair'] || r['Chủ trì'] || '').trim(),
      location:     String(r['Location'] || r['Địa điểm'] || '').trim(),
      status:       String(r['Status'] || r['Trạng thái'] || '').trim(),
      topicId:      String(r['Topic ID'] || r['TopicID'] || r['Mã chủ đề'] || '').trim(),
      summary:      String(r['Summary'] || r['Tóm tắt'] || '').trim(),
      ceoDirective: String(r['CEO Directive'] || r['CEODirective'] || r['Chỉ đạo CEO'] || '').trim(),
      ceoDecision:  String(r['CEO Decision'] || r['CEODecision'] || r['CEO cần chốt'] || '').trim(),
      attendees:    String(r['Attendees'] || r['Tham dự'] || '').trim(),
      actionCount:  0
    };
  }).filter(function(m) { return m.meetingId || m.meetingTitle; });

  /* ── Đọc MEETING_ACTIONS ── */
  var rawActions = [];
  try { rawActions = readSheetSafe('MEETING_ACTIONS'); } catch(e) {}
  var actions = rawActions.map(function(r) {
    var dlStr = String(r['Deadline'] || r['Hạn chót'] || '').trim();
    var dl = parseViDate_(dlStr);
    var isOverdue = dl ? (dl < today && (r['Status'] || r['Trạng thái'] || '') && !/hoàn thành|done/i.test(String(r['Status'] || r['Trạng thái'] || ''))) : false;
    return {
      actionId:    String(r['Action ID'] || r['ActionID'] || r['Mã CV'] || '').trim(),
      meetingId:   String(r['Meeting ID'] || r['MeetingID'] || '').trim(),
      topicId:     String(r['Topic ID']   || r['TopicID']   || '').trim(),
      content:     String(r['Content'] || r['Nội dung'] || r['Công việc'] || '').trim(),
      owner:       String(r['Owner'] || r['PIC'] || r['Người phụ trách'] || '').trim(),
      coOwner:     String(r['Co-Owner'] || r['CoOwner'] || r['Phối hợp'] || '').trim(),
      deadline:    dlStr,
      status:      String(r['Status'] || r['Trạng thái'] || '').trim(),
      rag:         String(r['RAG'] || '').trim() || 'Green',
      pct:         Number(r['% HT'] || r['PctHoanThanh'] || r['% Hoàn thành'] || 0),
      result:      String(r['Result'] || r['KetQuaThucTe'] || r['Kết quả thực tế'] || '').trim(),
      updatedDate: String(r['NgayCapNhat'] || r['Ngày CĐ'] || r['Ngày cập nhật'] || '').trim(),
      updatedBy:   String(r['NguoiCapNhat'] || r['Người CĐ'] || r['Người cập nhật'] || '').trim(),
      evidence:    String(r['BangChung'] || r['Bằng chứng'] || r['Evidence'] || '').trim(),
      blocker:     String(r['VuongMac'] || r['Vướng mắc'] || r['Blocker'] || '').trim(),
      ceoNeeded:   String(r['CEO cần chốt'] || r['CEODecision'] || r['CEO Decision'] || '').trim(),
      isOverdue:   isOverdue
    };
  }).filter(function(a) { return a.actionId || a.content; });

  /* ── Tính meetingCount và openActionCount cho từng topic ── */
  topics.forEach(function(t) {
    t.meetingCount    = meetings.filter(function(m) { return m.topicId === t.topicId; }).length;
    t.openActionCount = actions.filter(function(a)  { return a.topicId === t.topicId && a.pct < 100; }).length;
  });

  /* ── Tính actionCount cho từng meeting ── */
  meetings.forEach(function(m) {
    m.actionCount = actions.filter(function(a) { return a.meetingId === m.meetingId; }).length;
  });

  /* ── Build summary ── */
  var done    = actions.filter(function(a){ return a.pct>=100||/hoàn thành|done/i.test(a.status); }).length;
  var inProg  = actions.filter(function(a){ return /đang/i.test(a.status)&&a.pct<100&&!a.isOverdue; }).length;
  var overdue = actions.filter(function(a){ return a.isOverdue; }).length;

  return {
    meetings: meetings,
    actions:  actions,
    topics:   topics,
    people:   [],   // built client-side from actions
    summary: {
      totalMeetings: meetings.length,
      totalActions:  actions.length,
      done: done, inProgress: inProg, overdue: overdue
    },
    /* Legacy fields for backward compat with renderMeetingCenter */
    legacyTopics:   rawTopics.map(normalizeMeetingTopic_),
    legacyMinutes:  rawMinutes.map(normalizeMeetingMinutesFull_),
    legacyActions:  rawActions.map(normalizeMeetingAction_)
  };
}

/* Helper: safe date parse (reuse existing parseViDate_ if available, else inline) */
function parseViDate_(str) {
  if (!str) return null;
  try {
    var s = String(str).trim();
    // dd/MM/yyyy
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
    // yyyy-MM-dd
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  } catch(e) { return null; }
}

/* ══════════════════════════════════════════════════════
   MEETING HUB V3 — Drive & Detail functions
   (Backend prep — chưa deploy. Không gọi hàm này trực tiếp.)
   SECURITY: Chỉ deploy khi đã cấu hình DRIVE_ROOT_FOLDER_ID
══════════════════════════════════════════════════════ */

// Configure via Script Properties (không hardcode trong code):
// DRIVE_ROOT_FOLDER_ID  — ID folder gốc trên Drive
// MINUTES_TEMPLATE_FILE_ID — ID template Google Doc biên bản

var DRIVE_ROOT_FOLDER_ID      = 'CHUA_CAU_HINH'; // set via ScriptProperties
var MINUTES_TEMPLATE_FILE_ID  = 'CHUA_CAU_HINH'; // set via ScriptProperties

// SECURITY: Chỉ cho phép user thuộc domain goldsunpackaging.vn thực hiện thao tác ghi
function assertAuthorizedUser() {
  var email = Session.getActiveUser().getEmail();
  if (!email || !email.toLowerCase().endsWith('@goldsunpackaging.vn')) {
    throw new Error('UNAUTHORIZED: Chỉ tài khoản @goldsunpackaging.vn được phép thực hiện thao tác này. Email: ' + email);
  }
  return email;
}

function _getDriveRootFolder() {
  var rootId = PropertiesService.getScriptProperties().getProperty('DRIVE_ROOT_FOLDER_ID') || DRIVE_ROOT_FOLDER_ID;
  if (rootId === 'CHUA_CAU_HINH') throw new Error('DRIVE_ROOT_FOLDER_ID chưa được cấu hình trong Script Properties.');
  return DriveApp.getFolderById(rootId);
}

function _getMinutesTemplateId() {
  return PropertiesService.getScriptProperties().getProperty('MINUTES_TEMPLATE_FILE_ID') || MINUTES_TEMPLATE_FILE_ID;
}

// Tìm hoặc tạo subfolder theo tên trong folder cha
function _getOrCreateFolder(parentFolder, name) {
  var iter = parentFolder.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parentFolder.createFolder(name);
}

/**
 * Tạo Google Drive folder cho một cuộc họp theo cấu trúc:
 * ROOT/YYYY/MM/MTG-YYYYMMDD-XXX_TEN_CUOC_HOP/01_Bien_ban/02_Tai_lieu_hop/03_Bang_chung_Action/04_Khac
 * Trả về { folderId, folderUrl, subFolders }
 */
function createMeetingDriveFolder(meetingId, meetingTitle) {
  try {
    assertAuthorizedUser();
    var rootFolder = _getDriveRootFolder();
    // Parse YYYY/MM from meetingId (MTG-YYYYMMDD-XXX)
    var parts = (meetingId || '').split('-');
    var dateStr = parts[1] || '';
    var yyyy = dateStr.substring(0,4) || String(new Date().getFullYear());
    var mm   = dateStr.substring(4,6) || String(new Date().getMonth()+1).padStart(2,'0');
    var yearFolder  = _getOrCreateFolder(rootFolder, yyyy);
    var monthFolder = _getOrCreateFolder(yearFolder, mm);
    var safeName = (meetingTitle||meetingId).replace(/[\/\\:*?"<>|]/g,'_').substring(0,80);
    var folderName  = meetingId + '_' + safeName;
    var existing = monthFolder.getFoldersByName(folderName);
    var mtgFolder = existing.hasNext() ? existing.next() : monthFolder.createFolder(folderName);
    // Create sub-folders
    var subNames = ['01_Bien_ban','02_Tai_lieu_hop','03_Bang_chung_Action','04_Khac'];
    var subFolders = {};
    subNames.forEach(function(n){ subFolders[n] = _getOrCreateFolder(mtgFolder, n).getId(); });
    return { folderId: mtgFolder.getId(), folderUrl: mtgFolder.getUrl(), subFolders: subFolders };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Đính kèm file biên bản vào cuộc họp (ghi vào sheet MEETING_MINUTES).
 * fileId: Google Drive file ID
 */
function attachMinutesFileToMeeting(meetingId, fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var fileUrl = file.getUrl();
    var mimeType = file.getMimeType();
    var fileType = mimeType.indexOf('pdf') !== -1 ? 'pdf' : mimeType.indexOf('sheet') !== -1 ? 'gsheet' : 'gdoc';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MEETING_MINUTES');
    if (!sheet) return { error: 'Sheet MEETING_MINUTES không tồn tại' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var midCol = headers.indexOf('MeetingID');
    var urlCol = headers.indexOf('MinutesFileUrl');
    var typeCol = headers.indexOf('MinutesFileType');
    if (midCol < 0) return { error: 'Không tìm thấy cột MeetingID' };
    for (var i = 1; i < data.length; i++) {
      if (data[i][midCol] === meetingId) {
        if (urlCol >= 0) sheet.getRange(i+1, urlCol+1).setValue(fileUrl);
        if (typeCol >= 0) sheet.getRange(i+1, typeCol+1).setValue(fileType);
        return { success: true, fileUrl: fileUrl, fileType: fileType };
      }
    }
    return { error: 'Không tìm thấy cuộc họp ' + meetingId };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Lấy danh sách file trong Drive folder của một cuộc họp.
 */
function getMeetingFiles(meetingId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MEETING_MINUTES');
    if (!sheet) return { files: [] };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var midCol = headers.indexOf('MeetingID');
    var folderCol = headers.indexOf('DriveFolderId');
    if (midCol < 0) return { files: [] };
    var row = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][midCol] === meetingId) { row = data[i]; break; }
    }
    if (!row || folderCol < 0 || !row[folderCol]) return { files: [] };
    var folderId = row[folderCol];
    var folder = DriveApp.getFolderById(folderId);
    var files = [];
    var iter = folder.getFiles();
    while (iter.hasNext()) {
      var f = iter.next();
      files.push({ fileId: f.getId(), name: f.getName(), url: f.getUrl(), mimeType: f.getMimeType(), size: f.getSize(), lastUpdated: f.getLastUpdated().toISOString() });
    }
    return { files: files, folderId: folderId, folderUrl: folder.getUrl() };
  } catch (e) {
    return { error: e.message, files: [] };
  }
}

/**
 * Lấy chi tiết một cuộc họp cụ thể (bao gồm actions + files).
 * Dùng cho API endpoint getMeetingDetail?meetingId=MTG-XXXX
 */
function getMeetingDetail(meetingId) {
  try {
    var allData = getMeetingHubData();
    var meeting = (allData.meetings || []).filter(function(m){ return m.meetingId === meetingId; })[0];
    if (!meeting) return { error: 'Không tìm thấy cuộc họp ' + meetingId };
    var actions = (allData.actions || []).filter(function(a){ return a.meetingId === meetingId; });
    var files = getMeetingFiles(meetingId);
    return { meeting: meeting, actions: actions, files: files.files || [], folderId: files.folderId || '', folderUrl: files.folderUrl || '' };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Lấy các hành động của một cuộc họp cụ thể.
 */
function getMeetingActions(meetingId) {
  try {
    var allData = getMeetingHubData();
    var actions = (allData.actions || []).filter(function(a){ return a.meetingId === meetingId; });
    return { meetingId: meetingId, actions: actions, count: actions.length };
  } catch (e) {
    return { error: e.message, actions: [] };
  }
}

/**
 * Upload file vào folder của cuộc họp (subfolder 02_Tai_lieu_hop mặc định).
 * base64data: chuỗi base64 của file
 */
function uploadMeetingFile(meetingId, folderId, base64data, fileName, mimeType) {
  try {
    assertAuthorizedUser();
    if (!folderId) return { error: 'folderId không được để trống' };
    var folder = DriveApp.getFolderById(folderId);
    var blob = Utilities.newBlob(Utilities.base64Decode(base64data), mimeType || 'application/octet-stream', fileName);
    var file = folder.createFile(blob);
    return { success: true, fileId: file.getId(), name: file.getName(), url: file.getUrl(), mimeType: file.getMimeType(), size: file.getSize() };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Đổi tên file trên Google Drive.
 */
function renameMeetingFile(fileId, newName) {
  try {
    assertAuthorizedUser();
    var file = DriveApp.getFileById(fileId);
    file.setName(newName);
    return { success: true, fileId: fileId, newName: newName };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Chuyển file sang folder khác (không xóa file).
 */
function moveMeetingFile(fileId, targetFolderId) {
  try {
    assertAuthorizedUser();
    var file = DriveApp.getFileById(fileId);
    var targetFolder = DriveApp.getFolderById(targetFolderId);
    // Xóa khỏi các folder hiện tại
    var parents = file.getParents();
    while (parents.hasNext()) { file.removeFromFolder(parents.next()); }
    targetFolder.addFile(file);
    return { success: true, fileId: fileId, targetFolderId: targetFolderId };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Chuyển file vào Trash — KHÔNG xóa vĩnh viễn.
 */
function trashMeetingFile(fileId) {
  try {
    assertAuthorizedUser();
    var file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return { success: true, fileId: fileId, message: 'Đã chuyển "' + file.getName() + '" vào thùng rác.' };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Đính kèm link bằng chứng vào một Action (ghi vào sheet MEETING_ACTIONS).
 */
function attachEvidenceToAction(actionId, fileId) {
  try {
    assertAuthorizedUser();
    var file = DriveApp.getFileById(fileId);
    var fileUrl = file.getUrl();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MEETING_ACTIONS');
    if (!sheet) return { error: 'Sheet MEETING_ACTIONS không tồn tại' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var aidCol = headers.indexOf('ActionID');
    var evCol  = headers.indexOf('Evidence');
    if (aidCol < 0) return { error: 'Không tìm thấy cột ActionID' };
    if (evCol  < 0) return { error: 'Không tìm thấy cột Evidence' };
    for (var i = 1; i < data.length; i++) {
      if (data[i][aidCol] === actionId) {
        sheet.getRange(i+1, evCol+1).setValue(fileUrl);
        return { success: true, actionId: actionId, fileUrl: fileUrl };
      }
    }
    return { error: 'Không tìm thấy action ' + actionId };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Tạo Google Doc biên bản từ template.
 * Copy template vào folder 01_Bien_ban của cuộc họp.
 */
function createMinutesFromTemplate(meetingId, meetingTitle) {
  try {
    assertAuthorizedUser();
    var templateId = _getMinutesTemplateId();
    if (templateId === 'CHUA_CAU_HINH') return { error: 'MINUTES_TEMPLATE_FILE_ID chưa được cấu hình.' };
    // Lấy folderId của cuộc họp từ sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MEETING_MINUTES');
    var folderId = '';
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var headers = data[0];
      var midCol = headers.indexOf('MeetingID');
      var folderCol = headers.indexOf('DriveFolderId');
      if (midCol >= 0 && folderCol >= 0) {
        for (var i = 1; i < data.length; i++) {
          if (data[i][midCol] === meetingId) { folderId = data[i][folderCol]; break; }
        }
      }
    }
    var targetFolder;
    if (folderId) {
      var mtgFolder = DriveApp.getFolderById(folderId);
      var bbIter = mtgFolder.getFoldersByName('01_Bien_ban');
      targetFolder = bbIter.hasNext() ? bbIter.next() : _getOrCreateFolder(mtgFolder, '01_Bien_ban');
    } else {
      targetFolder = _getDriveRootFolder();
    }
    var template = DriveApp.getFileById(templateId);
    var newFileName = 'Biên bản — ' + (meetingTitle || meetingId) + ' — ' + meetingId;
    var newFile = template.makeCopy(newFileName, targetFolder);
    return { success: true, fileId: newFile.getId(), url: newFile.getUrl(), name: newFile.getName() };
  } catch (e) {
    return { error: e.message };
  }
}

/* ══════════════════════════════════════════════════════
   MEETING HUB V3 — doPost handler (merged into main doPost below)
   Giữ lại function này như alias để backward compat
══════════════════════════════════════════════════════ */
function doPostMeetingHub_(e) {
  function jsonOut(data) {
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var action = body.action || (e.parameter && e.parameter.action) || '';
    switch(action) {
      case 'uploadMeetingFile':
        return jsonOut(uploadMeetingFile(body.meetingId, body.folderId, body.base64data, body.fileName, body.mimeType));
      case 'renameMeetingFile':
        return jsonOut(renameMeetingFile(body.fileId, body.newName));
      case 'moveMeetingFile':
        return jsonOut(moveMeetingFile(body.fileId, body.targetFolderId));
      case 'trashMeetingFile':
        return jsonOut(trashMeetingFile(body.fileId));
      case 'attachEvidenceToAction':
        return jsonOut(attachEvidenceToAction(body.actionId, body.fileId));
      case 'createMinutesFromTemplate':
        return jsonOut(createMinutesFromTemplate(body.meetingId, body.meetingTitle));
      case 'attachMinutesFileToMeeting':
        return jsonOut(attachMinutesFileToMeeting(body.meetingId, body.fileId));
      case 'importMeetingActions':
        return jsonOut(importMeetingActionsFromPayload(body.rows || [], body.mode || 'batch', body.meetingId || ''));
      default:
        return jsonOut({ error: 'Unknown action: ' + action });
    }
  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

/**
 * Import actions từ payload (được gọi từ doPost importMeetingActions).
 * Chỉ ghi vào sheet MEETING_ACTIONS — KHÔNG ghi demo data.
 */
function importMeetingActionsFromPayload(rows, mode, defaultMeetingId) {
  assertAuthorizedUser();
  if (!rows || !rows.length) return { error: 'Không có dữ liệu để import' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MEETING_ACTIONS');
  if (!sheet) return { error: 'Sheet MEETING_ACTIONS không tồn tại' };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var imported = 0;
  rows.forEach(function(row) {
    var mid = String(row['MEETING_ID'] || defaultMeetingId || '').trim();
    if (!mid) return;
    var now = new Date().toISOString().substring(0,10);
    var newRow = headers.map(function(h) {
      var map = { 'MeetingID': mid, 'ActionTitle': row['NỘI_DUNG_CÔNG_VIỆC'], 'Content': row['MÔ_TẢ_CHI_TIẾT'],
        'OwnerEmail': row['OWNER_EMAIL'], 'Owner': row['OWNER_NAME'], 'CoOwnerEmails': row['CO_OWNER_EMAILS'],
        'Department': row['PHÒNG_BAN'], 'Deadline': row['DEADLINE'], 'ExpectedOutput': row['ĐẦU_RA_YÊU_CẦU'],
        'Priority': row['ƯU_TIÊN'], 'Status': row['TRẠNG_THÁI'], 'RAG': row['RAG'], 'Pct': row['TIẾN_ĐỘ_%'],
        'Result': row['KẾT_QUẢ_THỰC_TẾ'], 'Evidence': row['LINK_BẰNG_CHỨNG'], 'Blocker': row['VƯỚNG_MẮC'],
        'CEONeeded': row['CẦN_CEO_CHỐT'], 'CEODecision': row['CEO_DECISION'], 'Note': row['GHI_CHÚ'],
        'CreatedDate': now, 'UpdatedDate': now };
      return map[h] !== undefined ? map[h] : '';
    });
    sheet.appendRow(newRow);
    imported++;
  });
  return { success: true, imported: imported, message: 'Đã import ' + imported + ' hành động.' };
}

/* ═══════════════════════════════════════════════════════════════════
   MEETING HUB V3 — REAL DATA BACKEND
   Tất cả functions trong block này đọc/ghi từ Google Sheets & Drive thật.
   Cấu hình qua Script Properties — KHÔNG hardcode ID.
═══════════════════════════════════════════════════════════════════ */

/* ── Config helpers ── */
function getMeetingHubConfig() {
  var p = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: p.getProperty('MEETING_HUB_SPREADSHEET_ID') || '',
    driveRootId:   p.getProperty('DRIVE_ROOT_FOLDER_ID') || '',
    templateId:    p.getProperty('MINUTES_TEMPLATE_FILE_ID') || '',
    allowedDomain: p.getProperty('ALLOWED_DOMAIN') || 'goldsunpackaging.vn'
  };
}

function validateMeetingHubConfig() {
  var c = getMeetingHubConfig();
  var errors = [];
  if (!c.spreadsheetId) errors.push('MEETING_HUB_SPREADSHEET_ID chưa cấu hình');
  if (!c.driveRootId)   errors.push('DRIVE_ROOT_FOLDER_ID chưa cấu hình');
  return { valid: errors.length === 0, errors: errors, config: c };
}

function setupMeetingHubConfigExample() {
  Logger.log('=== Cấu hình Meeting Hub V3 ===');
  Logger.log('Vào Apps Script > Project Settings > Script Properties và thêm:');
  Logger.log('MEETING_HUB_SPREADSHEET_ID = <ID của Google Spreadsheet chứa MEETINGS/ACTIONS/PEOPLE/TOPICS/ACTION_LOG>');
  Logger.log('DRIVE_ROOT_FOLDER_ID = <ID folder gốc trên Google Drive>');
  Logger.log('MINUTES_TEMPLATE_FILE_ID = <ID file Google Doc template biên bản> (tuỳ chọn)');
  Logger.log('ALLOWED_DOMAIN = goldsunpackaging.vn');
  return { message: 'Xem log để biết cách cấu hình Script Properties.' };
}

/* ── assertMhUser_ — alias for assertAuthorizedUser (đã tồn tại) ── */
function assertMhUser_() { return assertAuthorizedUser(); }

/* ── Spreadsheet helper ── */
function getSpreadsheetForMeetingHub_() {
  var cfg = getMeetingHubConfig();
  if (cfg.spreadsheetId) {
    try { return SpreadsheetApp.openById(cfg.spreadsheetId); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* ── Sheet helper functions ── */
function mhSheetRowToObj_(headers, row) {
  var obj = {};
  headers.forEach(function(h, i) { obj[h] = (row[i] !== undefined && row[i] !== null) ? row[i] : ''; });
  return obj;
}

function mhGetSheetData_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  return data.slice(1).map(function(row) { return mhSheetRowToObj_(headers, row); });
}

function mhFindRowIndex_(sheet, colIndex, value) {
  if (sheet.getLastRow() < 2) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex] || '') === String(value)) return i + 1;
  }
  return -1;
}

function mhGetHeaderIndex_(sheet, headerName) {
  if (sheet.getLastRow() < 1) return -1;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName);
}

function mhSetCellByHeader_(sheet, rowIndex, headerName, value) {
  var col = mhGetHeaderIndex_(sheet, headerName);
  if (col >= 0) sheet.getRange(rowIndex, col + 1).setValue(value);
}

function mhListFilesInFolder_(folder, folderLabel) {
  var files = [];
  try {
    var iter = folder.getFiles();
    while (iter.hasNext()) {
      var f = iter.next();
      files.push({
        fileId: f.getId(), name: f.getName(), url: f.getUrl(),
        mimeType: f.getMimeType(), size: f.getSize(),
        modifiedDate: Utilities.formatDate(f.getLastUpdated(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd'),
        modifiedBy: '', folder: folderLabel
      });
    }
  } catch(e) { Logger.log('mhListFilesInFolder_ error: ' + e.message); }
  return files;
}

function mhGenerateMeetingId_(ss, dateStr) {
  var d = dateStr ? new Date(dateStr) : new Date();
  var yyyymmdd = Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'yyyyMMdd');
  var sheet = ss.getSheetByName('MEETINGS');
  var count = 0;
  if (sheet && sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    ids.forEach(function(row) { if (String(row[0]).indexOf('MTG-'+yyyymmdd) === 0) count++; });
  }
  return 'MTG-' + yyyymmdd + '-' + String(count + 1).padStart(3, '0');
}

function mhGenerateActionId_(ss, dateStr) {
  var d = dateStr ? new Date(dateStr) : new Date();
  var yyyymmdd = Utilities.formatDate(d, 'Asia/Ho_Chi_Minh', 'yyyyMMdd');
  var sheet = ss.getSheetByName('ACTIONS');
  var count = 0;
  if (sheet && sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    ids.forEach(function(row) { if (String(row[0]).indexOf('ACT-'+yyyymmdd) === 0) count++; });
  }
  return 'ACT-' + yyyymmdd + '-' + String(count + 1).padStart(4, '0');
}

function mhKpiCalc_(actions) {
  var today = new Date(); today.setHours(0,0,0,0);
  var d7 = new Date(today); d7.setDate(d7.getDate() + 7);
  return {
    overdue:      actions.filter(function(a){ var dl=a.deadline?new Date(a.deadline):null; return dl&&dl<today&&!a._done; }).length,
    dueSoon:      actions.filter(function(a){ var dl=a.deadline?new Date(a.deadline):null; return dl&&dl>=today&&dl<=d7&&!a._done; }).length,
    blocked:      actions.filter(function(a){ return !a._done&&String(a.blocker||'').trim(); }).length,
    noUpdate:     actions.filter(function(a){ var lu=a.lastUpdate||a.updatedDate; if(!lu||a._done) return false; var diff=(today-new Date(lu))/86400000; return diff>7; }).length,
    doneNoResult: actions.filter(function(a){ return a._done&&!String(a.result||'').trim(); }).length
  };
}

/* ── setupMeetingHubDataStore — tạo/chuẩn hóa 5 sheets ── */
function setupMeetingHubDataStore() {
  try {
    assertMhUser_();
    var ss = getSpreadsheetForMeetingHub_();
    var SCHEMAS = {
      'MEETINGS': ['meetingId','meetingDate','meetingTitle','topicId','topicName','meetingType','chairperson','secretary','participants','summary','mainContent','analysis','conclusions','decisions','ceoDirections','needCeoDecision','driveFolderId','driveFolderUrl','minutesFolderId','attachmentsFolderId','evidenceFolderId','otherFolderId','minutesFileId','minutesFileName','minutesFileType','minutesFileUrl','status','createdAt','updatedAt'],
      'ACTIONS':  ['actionId','meetingId','topicId','actionTitle','actionDescription','ownerEmail','ownerName','coOwnerEmails','department','deadline','expectedOutput','priority','status','rag','progress','actualResult','evidenceFileId','evidenceFileName','evidenceFileUrl','blocker','needCeoDecision','ceoDecision','lastUpdate','createdAt','updatedAt'],
      'PEOPLE':   ['email','fullName','department','title','active'],
      'TOPICS':   ['topicId','topicName','description','module','active','sortOrder'],
      'ACTION_LOG': ['logId','actionId','meetingId','updatedAt','updatedBy','oldStatus','newStatus','oldDeadline','newDeadline','progress','note']
    };
    var created = [], updated = [];
    Object.keys(SCHEMAS).forEach(function(sheetName) {
      var headers = SCHEMAS[sheetName];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
        created.push(sheetName);
      } else {
        var existing = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
        var missing = headers.filter(function(h){ return existing.indexOf(h) < 0; });
        if (missing.length > 0) {
          var startCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
          updated.push(sheetName + ' (+' + missing.join(',') + ')');
        }
      }
    });
    return { success: true, created: created, updated: updated };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── getMeetingHubData V3 — replaces legacy, với fallback ── */
function getMeetingHubData() {
  try {
    var validation = validateMeetingHubConfig();
    if (!validation.valid) {
      var legacy = getMeetingHubDataLegacy_();
      legacy._isDemo = true;
      legacy._demoReason = 'Chưa cấu hình Script Properties: ' + validation.errors.join('; ');
      return legacy;
    }
    return getMeetingHubDataFromSheets_();
  } catch(e) {
    try {
      var leg = getMeetingHubDataLegacy_();
      leg._isDemo = true;
      leg._demoReason = 'Lỗi V3: ' + e.message;
      return leg;
    } catch(e2) {
      return { meetings:[], actions:[], topics:[], people:[], _isDemo:true, _demoReason: e.message };
    }
  }
}

function getMeetingHubDataFromSheets_() {
  var ss = getSpreadsheetForMeetingHub_();
  var today = new Date(); today.setHours(0,0,0,0);
  var d7 = new Date(today); d7.setDate(d7.getDate()+7);

  /* ── Topics ── */
  var rawTopics = mhGetSheetData_(ss, 'TOPICS');
  var topics = rawTopics.filter(function(r){ return String(r.active||'').toLowerCase()!=='false'&&(r.topicId||r.topicName); }).map(function(r){
    return { topicId: String(r.topicId||'').trim(), topicName: String(r.topicName||'').trim(), module: String(r.module||'').trim(), desc: String(r.description||'').trim(), meetingCount:0, openActionCount:0 };
  });

  /* ── People ── */
  var rawPeople = mhGetSheetData_(ss, 'PEOPLE');
  var people = rawPeople.filter(function(r){ return r.email; }).map(function(r){
    return { email: String(r.email||'').trim(), fullName: String(r.fullName||'').trim(), department: String(r.department||'').trim(), title: String(r.title||'').trim(), active: String(r.active||'true').toLowerCase()!=='false' };
  });

  /* ── Actions ── */
  var rawActions = mhGetSheetData_(ss, 'ACTIONS');
  var actions = rawActions.filter(function(r){ return r.actionId||r.actionTitle; }).map(function(r){
    var pct = Number(r.progress||0);
    var dlStr = r.deadline ? Utilities.formatDate(r.deadline instanceof Date ? r.deadline : new Date(r.deadline), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd') : '';
    var dl = dlStr ? new Date(dlStr) : null;
    var status = String(r.status||'').trim();
    var isDone = pct>=100 || /hoàn thành|done/i.test(status);
    var isOverdue = dl && dl<today && !isDone;
    var dueSoon = dl && dl>=today && dl<=d7 && !isDone;
    var luStr = r.lastUpdate ? (r.lastUpdate instanceof Date ? Utilities.formatDate(r.lastUpdate,'Asia/Ho_Chi_Minh','yyyy-MM-dd') : String(r.lastUpdate)) : '';
    return {
      actionId:      String(r.actionId||'').trim(),
      meetingId:     String(r.meetingId||'').trim(),
      topicId:       String(r.topicId||'').trim(),
      actionTitle:   String(r.actionTitle||'').trim(),
      content:       String(r.actionDescription||'').trim(),
      owner:         String(r.ownerName||'').trim(),
      ownerEmail:    String(r.ownerEmail||'').trim(),
      coOwner:       String(r.coOwnerEmails||'').trim(),
      department:    String(r.department||'').trim(),
      deadline:      dlStr,
      status:        status,
      rag:           String(r.rag||'Green').trim(),
      pct:           pct,
      priority:      String(r.priority||'').trim(),
      expectedOutput:String(r.expectedOutput||'').trim(),
      result:        String(r.actualResult||'').trim(),
      updatedDate:   luStr,
      updatedBy:     '',
      evidence:      String(r.evidenceFileUrl||'').trim(),
      blocker:       String(r.blocker||'').trim(),
      ceoNeeded:     String(r.needCeoDecision||'').trim(),
      history:       [],
      isOverdue:     isOverdue,
      dueSoon:       dueSoon,
      _done:         isDone
    };
  });

  /* ── Meetings ── */
  var rawMeetings = mhGetSheetData_(ss, 'MEETINGS');
  var meetings = rawMeetings.filter(function(r){ return r.meetingId||r.meetingTitle; }).map(function(r){
    var dateStr = r.meetingDate ? (r.meetingDate instanceof Date ? Utilities.formatDate(r.meetingDate,'Asia/Ho_Chi_Minh','yyyy-MM-dd') : String(r.meetingDate)) : '';
    var mActions = actions.filter(function(a){ return a.meetingId===String(r.meetingId||'').trim(); });
    return {
      meetingId:      String(r.meetingId||'').trim(),
      meetingTitle:   String(r.meetingTitle||'').trim(),
      date:           dateStr,
      type:           String(r.meetingType||'').trim(),
      chair:          String(r.chairperson||'').trim(),
      location:       '',
      status:         String(r.status||'').trim(),
      topicId:        String(r.topicId||'').trim(),
      summary:        String(r.summary||'').trim(),
      decisions:      String(r.decisions||'').trim(),
      ceoDirective:   String(r.ceoDirections||'').trim(),
      ceoDecision:    String(r.needCeoDecision||'').trim(),
      attendees:      String(r.participants||'').trim(),
      driveFolderId:  String(r.driveFolderId||'').trim(),
      minutesFileUrl: String(r.minutesFileUrl||'').trim(),
      minutesFileType:String(r.minutesFileType||'').trim(),
      minutesFileId:  String(r.minutesFileId||'').trim(),
      actionCount:    mActions.length,
      driveFiles:     []
    };
  });

  /* ── Tính aggregates ── */
  topics.forEach(function(t){
    t.meetingCount    = meetings.filter(function(m){ return m.topicId===t.topicId; }).length;
    t.openActionCount = actions.filter(function(a){ return a.topicId===t.topicId&&!a._done; }).length;
  });
  var kpi = mhKpiCalc_(actions);
  var done    = actions.filter(function(a){ return a._done; }).length;
  var inProg  = actions.filter(function(a){ return !a._done&&!a.isOverdue; }).length;

  return {
    meetings: meetings, actions: actions, topics: topics, people: people,
    _isDemo: false,
    summary: {
      totalMeetings: meetings.length, totalActions: actions.length,
      done: done, inProgress: inProg, overdue: kpi.overdue,
      dueSoon: kpi.dueSoon, blocked: kpi.blocked,
      noUpdate: kpi.noUpdate, doneNoResult: kpi.doneNoResult
    }
  };
}

/* ── getMeetingDetailV3 ── */
function getMeetingDetailV3(meetingId) {
  try {
    var allData = getMeetingHubData();
    var meeting = (allData.meetings||[]).filter(function(m){ return m.meetingId===meetingId; })[0];
    if (!meeting) return { error: 'Không tìm thấy cuộc họp: ' + meetingId };
    var mActions = (allData.actions||[]).filter(function(a){ return a.meetingId===meetingId; });
    var files = getMeetingFilesV3_(meetingId, meeting.driveFolderId);
    meeting.driveFiles = files;
    return { meeting: meeting, actions: mActions, files: files };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── getMeetingFilesV3_ — list files từ Drive folder thật ── */
function getMeetingFilesV3_(meetingId, driveFolderId) {
  if (!driveFolderId) return [];
  var files = [];
  try {
    var mtgFolder = DriveApp.getFolderById(driveFolderId);
    var subLabels = { '01_Bien_ban':'Biên bản', '02_Tai_lieu_hop':'Tài liệu họp', '03_Bang_chung_Action':'Bằng chứng', '04_Khac':'Khác' };
    Object.keys(subLabels).forEach(function(subName){
      try {
        var subIter = mtgFolder.getFoldersByName(subName);
        if (subIter.hasNext()) {
          files = files.concat(mhListFilesInFolder_(subIter.next(), subLabels[subName]));
        }
      } catch(e2) {}
    });
    // Cũng list files ở root folder nếu có
    files = files.concat(mhListFilesInFolder_(mtgFolder, 'Root'));
  } catch(e) {
    Logger.log('getMeetingFilesV3_ error for ' + meetingId + ': ' + e.message);
  }
  return files;
}

/* ── createMeeting ── */
function createMeeting(meetingData) {
  try {
    assertMhUser_();
    var ss = getSpreadsheetForMeetingHub_();
    var sheet = ss.getSheetByName('MEETINGS');
    if (!sheet) return { error: 'Sheet MEETINGS không tồn tại. Chạy setupMeetingHubDataStore() trước.' };
    var now = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    var meetingDate = meetingData.meetingDate || meetingData.date || Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    var meetingId = mhGenerateMeetingId_(ss, meetingDate);
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var rowData = {
      meetingId: meetingId,
      meetingDate: meetingDate,
      meetingTitle: meetingData.meetingTitle||meetingData.title||'',
      topicId: meetingData.topicId||'',
      topicName: meetingData.topicName||'',
      meetingType: meetingData.meetingType||meetingData.type||'',
      chairperson: meetingData.chairperson||meetingData.chair||'',
      secretary: meetingData.secretary||'',
      participants: meetingData.participants||meetingData.attendees||'',
      summary: meetingData.summary||'',
      status: meetingData.status||'Planned',
      createdAt: now, updatedAt: now
    };
    var newRow = headers.map(function(h){ return rowData[h]!==undefined ? rowData[h] : ''; });
    sheet.appendRow(newRow);
    var lastRow = sheet.getLastRow();
    // Tạo Drive folder
    var driveResult = {};
    try {
      driveResult = createMeetingDriveFolder(meetingId, rowData.meetingTitle);
      if (driveResult.folderId) {
        mhSetCellByHeader_(sheet, lastRow, 'driveFolderId',  driveResult.folderId);
        mhSetCellByHeader_(sheet, lastRow, 'driveFolderUrl', driveResult.folderUrl);
        if (driveResult.subFolders) {
          mhSetCellByHeader_(sheet, lastRow, 'minutesFolderId',     driveResult.subFolders['01_Bien_ban']||'');
          mhSetCellByHeader_(sheet, lastRow, 'attachmentsFolderId', driveResult.subFolders['02_Tai_lieu_hop']||'');
          mhSetCellByHeader_(sheet, lastRow, 'evidenceFolderId',    driveResult.subFolders['03_Bang_chung_Action']||'');
          mhSetCellByHeader_(sheet, lastRow, 'otherFolderId',       driveResult.subFolders['04_Khac']||'');
        }
        rowData.driveFolderId  = driveResult.folderId;
        rowData.driveFolderUrl = driveResult.folderUrl;
      }
    } catch(de) { Logger.log('Drive folder error: ' + de.message); }
    return { success: true, meetingId: meetingId, meeting: rowData, driveResult: driveResult };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── updateAction ── */
function updateAction(actionId, updateData) {
  try {
    assertMhUser_();
    var ss = getSpreadsheetForMeetingHub_();
    var sheet = ss.getSheetByName('ACTIONS');
    if (!sheet) return { error: 'Sheet ACTIONS không tồn tại' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var aidCol = headers.indexOf('actionId');
    if (aidCol < 0) return { error: 'Cột actionId không tồn tại' };
    var rowIdx = mhFindRowIndex_(sheet, aidCol, actionId);
    if (rowIdx < 0) return { error: 'Không tìm thấy action: ' + actionId };
    var oldRow = sheet.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
    var oldObj = mhSheetRowToObj_(headers, oldRow);
    var now = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    var changed = [];
    var trackedFields = ['status','deadline','ownerEmail','ownerName','progress','actualResult','blocker'];
    trackedFields.forEach(function(f){
      if (updateData[f]!==undefined && String(updateData[f])!==String(oldObj[f]||'')) {
        changed.push({ field:f, old: oldObj[f], new_: updateData[f] });
      }
    });
    // Áp dụng update
    var fieldMap = { status:'status', deadline:'deadline', ownerEmail:'ownerEmail', ownerName:'ownerName',
      progress:'progress', actualResult:'actualResult', blocker:'blocker', rag:'rag', priority:'priority',
      expectedOutput:'expectedOutput', coOwnerEmails:'coOwnerEmails', needCeoDecision:'needCeoDecision' };
    Object.keys(fieldMap).forEach(function(k){
      if (updateData[k]!==undefined) {
        var col = headers.indexOf(fieldMap[k]);
        if (col>=0) sheet.getRange(rowIdx, col+1).setValue(updateData[k]);
      }
    });
    var luCol = headers.indexOf('lastUpdate');
    if (luCol>=0) sheet.getRange(rowIdx, luCol+1).setValue(now);
    var updCol = headers.indexOf('updatedAt');
    if (updCol>=0) sheet.getRange(rowIdx, updCol+1).setValue(now);
    // Ghi ACTION_LOG nếu có thay đổi
    if (changed.length > 0) {
      try {
        var logSheet = ss.getSheetByName('ACTION_LOG');
        if (logSheet) {
          var logHeaders = logSheet.getLastColumn()>0 ? logSheet.getRange(1,1,1,logSheet.getLastColumn()).getValues()[0] : [];
          var logNow = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMddHHmmss');
          changed.forEach(function(ch){
            var logId = 'LOG-' + logNow + '-' + Math.random().toString(36).substr(2,4).toUpperCase();
            var logRow = logHeaders.map(function(h){
              var m = { logId:logId, actionId:actionId, meetingId:oldObj.meetingId||'', updatedAt:now,
                updatedBy: Session.getActiveUser().getEmail()||'', note:ch.field+': '+ch.old+' → '+ch.new_ };
              if (ch.field==='status') { m.oldStatus=ch.old; m.newStatus=ch.new_; }
              if (ch.field==='deadline') { m.oldDeadline=ch.old; m.newDeadline=ch.new_; }
              if (ch.field==='progress') { m.progress=ch.new_; }
              return m[h]!==undefined ? m[h] : '';
            });
            logSheet.appendRow(logRow);
          });
        }
      } catch(le) { Logger.log('ACTION_LOG error: ' + le.message); }
    }
    return { success: true, actionId: actionId, changedFields: changed.length };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── getPersonWorkload ── */
function getPersonWorkload(email) {
  try {
    var ss = getSpreadsheetForMeetingHub_();
    var allActions = mhGetSheetData_(ss, 'ACTIONS');
    var today = new Date(); today.setHours(0,0,0,0);
    var d7 = new Date(today); d7.setDate(d7.getDate()+7);
    var myActions = allActions.filter(function(r){
      return String(r.ownerEmail||'').trim()===email || String(r.coOwnerEmails||'').indexOf(email)>=0;
    }).map(function(r){
      var pct = Number(r.progress||0);
      var dl = r.deadline ? new Date(r.deadline) : null;
      var isDone = pct>=100 || /hoàn thành|done/i.test(String(r.status||''));
      return {
        actionId: r.actionId, actionTitle: r.actionTitle, meetingId: r.meetingId,
        status: r.status, deadline: r.deadline, pct: pct, rag: r.rag, blocker: r.blocker,
        isDone: isDone,
        isOverdue: dl&&dl<today&&!isDone,
        dueSoon: dl&&dl>=today&&dl<=d7&&!isDone
      };
    });
    return {
      email: email,
      total:   myActions.length,
      done:    myActions.filter(function(a){ return a.isDone; }).length,
      open:    myActions.filter(function(a){ return !a.isDone; }).length,
      overdue: myActions.filter(function(a){ return a.isOverdue; }).length,
      dueSoon: myActions.filter(function(a){ return a.dueSoon; }).length,
      blocked: myActions.filter(function(a){ return !a.isDone&&String(a.blocker||'').trim(); }).length,
      actions: myActions
    };
  } catch(e) {
    return { error: e.message, email: email };
  }
}

/* ── getPersonDetail ── */
function getPersonDetail(email) {
  try {
    var ss = getSpreadsheetForMeetingHub_();
    var people = mhGetSheetData_(ss, 'PEOPLE');
    var person = people.filter(function(p){ return String(p.email||'').trim()===email; })[0] || { email: email };
    var workload = getPersonWorkload(email);
    return { person: person, workload: workload, actions: workload.actions || [] };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── getTopicDetail ── */
function getTopicDetail(topicId) {
  try {
    var ss = getSpreadsheetForMeetingHub_();
    var topics = mhGetSheetData_(ss, 'TOPICS');
    var topic = topics.filter(function(t){ return String(t.topicId||'').trim()===topicId; })[0] || { topicId: topicId };
    var meetings = mhGetSheetData_(ss, 'MEETINGS').filter(function(m){ return String(m.topicId||'').trim()===topicId; });
    var actions  = mhGetSheetData_(ss, 'ACTIONS').filter(function(a){  return String(a.topicId||'').trim()===topicId; });
    return { topic: topic, meetings: meetings, actions: actions, meetingCount: meetings.length, actionCount: actions.length };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── testMeetingHubConnection — diagnostic ── */
function testMeetingHubConnection() {
  var report = { ts: Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss'), user: '', config: {}, spreadsheet: null, driveRoot: null, template: null, sheetCounts: {}, errors: [] };
  try { report.user = Session.getActiveUser().getEmail(); } catch(e) { report.errors.push('User: ' + e.message); }
  var v = validateMeetingHubConfig();
  report.config = v.config;
  report.configValid = v.valid;
  if (!v.valid) report.errors = report.errors.concat(v.errors);
  // Spreadsheet
  try {
    var ss = getSpreadsheetForMeetingHub_();
    report.spreadsheet = { name: ss.getName(), id: ss.getId() };
    ['MEETINGS','ACTIONS','PEOPLE','TOPICS','ACTION_LOG'].forEach(function(sn){
      var sh = ss.getSheetByName(sn);
      report.sheetCounts[sn] = sh ? Math.max(0, sh.getLastRow()-1) : 'sheet không tồn tại';
    });
  } catch(e) { report.errors.push('Spreadsheet: ' + e.message); }
  // Drive root
  try {
    var cfg = getMeetingHubConfig();
    if (cfg.driveRootId) {
      var rf = DriveApp.getFolderById(cfg.driveRootId);
      report.driveRoot = { name: rf.getName(), id: rf.getId(), url: rf.getUrl() };
    }
  } catch(e) { report.errors.push('Drive root: ' + e.message); }
  // Template
  try {
    var cfg2 = getMeetingHubConfig();
    if (cfg2.templateId) {
      var tf = DriveApp.getFileById(cfg2.templateId);
      report.template = { name: tf.getName(), id: tf.getId() };
    }
  } catch(e) { report.errors.push('Template: ' + e.message); }
  report.ok = report.errors.length === 0;
  return report;
}

/* ── seedMeetingHubPeopleFromArray ── */
function seedMeetingHubPeopleFromArray(peopleArray) {
  try {
    assertMhUser_();
    if (!peopleArray || !peopleArray.length) return { error: 'Không có dữ liệu' };
    var ss = getSpreadsheetForMeetingHub_();
    var sheet = ss.getSheetByName('PEOPLE');
    if (!sheet) return { error: 'Sheet PEOPLE không tồn tại. Chạy setupMeetingHubDataStore() trước.' };
    var existing = mhGetSheetData_(ss, 'PEOPLE').map(function(r){ return String(r.email||'').trim().toLowerCase(); });
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : ['email','fullName','department','title','active'];
    var added=0, skipped=0;
    peopleArray.forEach(function(p){
      var em = String(p.email||'').trim().toLowerCase();
      if (!em || existing.indexOf(em)>=0) { skipped++; return; }
      var row = headers.map(function(h){ return p[h]!==undefined ? p[h] : (h==='active'?'true':''); });
      sheet.appendRow(row);
      existing.push(em);
      added++;
    });
    return { success: true, added: added, skipped: skipped };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── migrateLegacyMeetingHubData ── */
function migrateLegacyMeetingHubData(dryRun) {
  try {
    assertMhUser_();
    var dryRunMode = (dryRun !== false); // default true
    var ss = getSpreadsheetForMeetingHub_();
    var log = [];
    var migrated = 0, skipped = 0;

    // Migrate MEETING_TOPICS → TOPICS
    var topicsSheet = ss.getSheetByName('TOPICS');
    var topicsHasData = topicsSheet && topicsSheet.getLastRow() > 1;
    if (!topicsHasData) {
      var legTopics = [];
      try { legTopics = readSheetSafe('MEETING_TOPICS'); } catch(e) {}
      if (legTopics.length > 0 && !dryRunMode) {
        if (!topicsSheet) topicsSheet = ss.insertSheet('TOPICS');
        if (topicsSheet.getLastRow() < 1) topicsSheet.appendRow(['topicId','topicName','description','module','active','sortOrder']);
        legTopics.forEach(function(r){
          topicsSheet.appendRow([r['Topic ID']||r['TopicID']||'', r['Topic Name']||r['TopicName']||'', r['Description']||r['Mô tả']||'', r['Module']||'', 'true', '']);
          migrated++;
        });
        log.push('Migrated ' + legTopics.length + ' topics từ MEETING_TOPICS → TOPICS');
      } else {
        log.push((dryRunMode?'[DRY RUN] Would migrate ':'') + legTopics.length + ' topics từ MEETING_TOPICS → TOPICS');
        if (dryRunMode) skipped += legTopics.length;
      }
    } else { log.push('TOPICS đã có dữ liệu — bỏ qua migrate'); skipped++; }

    // Migrate MEETING_MINUTES → MEETINGS
    var meetingsSheet = ss.getSheetByName('MEETINGS');
    var meetingsHasData = meetingsSheet && meetingsSheet.getLastRow() > 1;
    if (!meetingsHasData) {
      var legMinutes = [];
      try { legMinutes = readSheetSafe('MEETING_MINUTES'); } catch(e) {}
      if (legMinutes.length > 0 && !dryRunMode) {
        if (!meetingsSheet) meetingsSheet = ss.insertSheet('MEETINGS');
        if (meetingsSheet.getLastRow() < 1) meetingsSheet.appendRow(['meetingId','meetingDate','meetingTitle','topicId','topicName','meetingType','chairperson','secretary','participants','summary','status','createdAt','updatedAt']);
        var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
        legMinutes.forEach(function(r){
          meetingsSheet.appendRow([r['Meeting ID']||r['MeetingID']||'', r['Date']||r['Ngày họp']||'', r['Topic']||r['Chủ đề']||'', r['Topic ID']||r['TopicID']||'', '', r['Type']||r['Loại họp']||'', r['Chair']||r['Chủ trì']||'', '', r['Attendees']||r['Tham dự']||'', r['Summary']||r['Tóm tắt']||'', r['Status']||r['Trạng thái']||'', now, now]);
          migrated++;
        });
        log.push('Migrated ' + legMinutes.length + ' meetings từ MEETING_MINUTES → MEETINGS');
      } else {
        log.push((dryRunMode?'[DRY RUN] Would migrate ':'') + legMinutes.length + ' meetings từ MEETING_MINUTES → MEETINGS');
        if (dryRunMode) skipped += legMinutes.length;
      }
    } else { log.push('MEETINGS đã có dữ liệu — bỏ qua migrate'); skipped++; }

    // Migrate MEETING_ACTIONS → ACTIONS
    var actionsSheet = ss.getSheetByName('ACTIONS');
    var actionsHasData = actionsSheet && actionsSheet.getLastRow() > 1;
    if (!actionsHasData) {
      var legActions = [];
      try { legActions = readSheetSafe('MEETING_ACTIONS'); } catch(e) {}
      if (legActions.length > 0 && !dryRunMode) {
        if (!actionsSheet) actionsSheet = ss.insertSheet('ACTIONS');
        if (actionsSheet.getLastRow() < 1) actionsSheet.appendRow(['actionId','meetingId','topicId','actionTitle','actionDescription','ownerEmail','ownerName','coOwnerEmails','department','deadline','expectedOutput','priority','status','rag','progress','actualResult','blocker','needCeoDecision','lastUpdate','createdAt','updatedAt']);
        var now2 = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
        legActions.forEach(function(r){
          actionsSheet.appendRow([r['Action ID']||r['ActionID']||r['Mã CV']||'', r['Meeting ID']||r['MeetingID']||'', r['Topic ID']||r['TopicID']||'', r['Content']||r['Nội dung']||r['Công việc']||'', '', r['OwnerEmail']||'', r['Owner']||r['PIC']||r['Người phụ trách']||'', r['Co-Owner']||r['CoOwner']||'', r['Department']||r['Phòng']||'', r['Deadline']||r['Hạn chót']||'', '', r['Priority']||'', r['Status']||r['Trạng thái']||'', r['RAG']||'Green', r['% HT']||r['% Hoàn thành']||0, r['Result']||r['Kết quả thực tế']||'', r['VuongMac']||r['Blocker']||'', r['CEO cần chốt']||'', now2, now2, now2]);
          migrated++;
        });
        log.push('Migrated ' + legActions.length + ' actions từ MEETING_ACTIONS → ACTIONS');
      } else {
        log.push((dryRunMode?'[DRY RUN] Would migrate ':'') + legActions.length + ' actions từ MEETING_ACTIONS → ACTIONS');
        if (dryRunMode) skipped += legActions.length;
      }
    } else { log.push('ACTIONS đã có dữ liệu — bỏ qua migrate'); skipped++; }

    return { success: true, dryRun: dryRunMode, migrated: migrated, skipped: skipped, log: log };
  } catch(e) {
    return { error: e.message };
  }
}

/* ── attachMinutesFileToMeetingV3 — ghi vào MEETINGS sheet mới ── */
function attachMinutesFileToMeetingV3(meetingId, fileId) {
  try {
    assertMhUser_();
    var file = DriveApp.getFileById(fileId);
    var fileUrl = file.getUrl(); var mime = file.getMimeType();
    var fileType = mime.indexOf('pdf')>=0?'pdf':mime.indexOf('sheet')>=0?'gsheet':'gdoc';
    var ss = getSpreadsheetForMeetingHub_();
    var sheet = ss.getSheetByName('MEETINGS');
    if (!sheet) return { error: 'Sheet MEETINGS không tồn tại' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var midCol = headers.indexOf('meetingId');
    if (midCol<0) return { error: 'Cột meetingId không tồn tại' };
    var rowIdx = mhFindRowIndex_(sheet, midCol, meetingId);
    if (rowIdx<0) return { error: 'Không tìm thấy cuộc họp: ' + meetingId };
    mhSetCellByHeader_(sheet, rowIdx, 'minutesFileId',   fileId);
    mhSetCellByHeader_(sheet, rowIdx, 'minutesFileName', file.getName());
    mhSetCellByHeader_(sheet, rowIdx, 'minutesFileType', fileType);
    mhSetCellByHeader_(sheet, rowIdx, 'minutesFileUrl',  fileUrl);
    mhSetCellByHeader_(sheet, rowIdx, 'updatedAt', Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss'));
    return { success: true, fileUrl: fileUrl, fileType: fileType, fileName: file.getName() };
  } catch(e) { return { error: e.message }; }
}

/* ── attachEvidenceToActionV3 — ghi vào ACTIONS sheet mới ── */
function attachEvidenceToActionV3(actionId, fileId) {
  try {
    assertMhUser_();
    var file = DriveApp.getFileById(fileId);
    var ss = getSpreadsheetForMeetingHub_();
    var sheet = ss.getSheetByName('ACTIONS');
    if (!sheet) return { error: 'Sheet ACTIONS không tồn tại' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var aidCol = headers.indexOf('actionId');
    if (aidCol<0) return { error: 'Cột actionId không tồn tại' };
    var rowIdx = mhFindRowIndex_(sheet, aidCol, actionId);
    if (rowIdx<0) return { error: 'Không tìm thấy action: ' + actionId };
    mhSetCellByHeader_(sheet, rowIdx, 'evidenceFileId',   fileId);
    mhSetCellByHeader_(sheet, rowIdx, 'evidenceFileName', file.getName());
    mhSetCellByHeader_(sheet, rowIdx, 'evidenceFileUrl',  file.getUrl());
    mhSetCellByHeader_(sheet, rowIdx, 'updatedAt', Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss'));
    return { success: true, actionId: actionId, fileUrl: file.getUrl() };
  } catch(e) { return { error: e.message }; }
}

/* ── importMeetingActionsFromPayload V3 — đọc/ghi vào ACTIONS sheet mới ── */
function importMeetingActionsV3(rows, mode, defaultMeetingId) {
  try {
    assertMhUser_();
    if (!rows || !rows.length) return { error: 'Không có dữ liệu để import' };
    var ss = getSpreadsheetForMeetingHub_();
    var sheet = ss.getSheetByName('ACTIONS');
    if (!sheet) return { error: 'Sheet ACTIONS không tồn tại. Chạy setupMeetingHubDataStore() trước.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var meetingsSheet = ss.getSheetByName('MEETINGS');
    var validMeetingIds = [];
    if (meetingsSheet && meetingsSheet.getLastRow()>1) {
      validMeetingIds = meetingsSheet.getRange(2,1,meetingsSheet.getLastRow()-1,1).getValues().map(function(r){ return String(r[0]||'').trim(); }).filter(Boolean);
    }
    var peopleSheet = ss.getSheetByName('PEOPLE');
    var validEmails = [];
    if (peopleSheet && peopleSheet.getLastRow()>1) {
      var ph = peopleSheet.getRange(1,1,1,peopleSheet.getLastColumn()).getValues()[0];
      var emCol = ph.indexOf('email');
      if (emCol>=0) validEmails = peopleSheet.getRange(2,1,peopleSheet.getLastRow()-1,peopleSheet.getLastColumn()).getValues().map(function(r){ return String(r[emCol]||'').trim().toLowerCase(); }).filter(Boolean);
    }
    var imported=0, errors=[];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    rows.forEach(function(row, idx){
      var mid = String(row.meetingId || row.MEETING_ID || defaultMeetingId || '').trim();
      var title = String(row.actionTitle || row.NỘI_DUNG_CÔNG_VIỆC || '').trim();
      var ownerEmail = String(row.ownerEmail || row.OWNER_EMAIL || '').trim().toLowerCase();
      var deadline = String(row.deadline || row.DEADLINE || '').trim();
      // Validate
      if (!mid) { errors.push('Row ' + (idx+1) + ': thiếu meetingId'); return; }
      if (!title) { errors.push('Row ' + (idx+1) + ': thiếu actionTitle'); return; }
      if (validMeetingIds.length>0 && validMeetingIds.indexOf(mid)<0) { errors.push('Row ' + (idx+1) + ': meetingId "' + mid + '" không tồn tại'); return; }
      if (ownerEmail && validEmails.length>0 && validEmails.indexOf(ownerEmail)<0) { errors.push('Row ' + (idx+1) + ': ownerEmail "' + ownerEmail + '" không có trong PEOPLE'); return; }
      var actionId = mhGenerateActionId_(ss, null);
      var rowData = { actionId:actionId, meetingId:mid, topicId:row.topicId||'', actionTitle:title,
        actionDescription:row.actionDescription||row.MÔ_TẢ_CHI_TIẾT||'', ownerEmail:ownerEmail,
        ownerName:row.ownerName||row.OWNER_NAME||'', coOwnerEmails:row.coOwnerEmails||row.CO_OWNER_EMAILS||'',
        department:row.department||row.PHÒNG_BAN||'', deadline:deadline,
        expectedOutput:row.expectedOutput||row.ĐẦU_RA_YÊU_CẦU||'',
        priority:row.priority||row.ƯU_TIÊN||'Normal',
        status:row.status||row.TRẠNG_THÁI||'Open',
        rag:row.rag||row.RAG||'Green',
        progress:Number(row.progress||row['TIẾN_ĐỘ_%']||0),
        actualResult:'', blocker:'', needCeoDecision:'', createdAt:now, updatedAt:now, lastUpdate:now };
      var newRow = headers.map(function(h){ return rowData[h]!==undefined ? rowData[h] : ''; });
      sheet.appendRow(newRow);
      imported++;
    });
    return { success: true, imported: imported, errors: errors, message: 'Đã import ' + imported + ' action' + (errors.length?' ('+errors.length+' lỗi bị bỏ qua)':'') + '.' };
  } catch(e) {
    return { error: e.message };
  }
}

/* ═══ END MEETING HUB V3 — REAL DATA BACKEND ═══ */

/* ═══ LOS — LEADERSHIP OPERATING SYSTEM — Backend ═══ */

/* ── Config & Sheet helpers ── */
function getLosConfig() {
  var p = PropertiesService.getScriptProperties();
  return {
    spreadsheetId: p.getProperty('LOS_SPREADSHEET_ID') || p.getProperty('MEETING_HUB_SPREADSHEET_ID') || '',
    allowedDomain: p.getProperty('ALLOWED_DOMAIN') || 'goldsunpackaging.vn'
  };
}

function getLosSpreadsheet_() {
  var cfg = getLosConfig();
  if (cfg.spreadsheetId) {
    try { return SpreadsheetApp.openById(cfg.spreadsheetId); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/* LOS_ELIGIBLE_TITLES — chức danh được phép tham gia LOS */
var LOS_ELIGIBLE_LEVEL_KEYWORDS = [
  'giám đốc', 'director', 'trưởng phòng', 'trưởng bộ phận',
  'quản lý', 'manager', 'phó giám đốc', 'deputy', 'head of',
  'cfo', 'coo', 'ceo', 'cto', 'ban điều hành', 'ban tổng giám đốc'
];
var LOS_INELIGIBLE_KEYWORDS = ['quản đốc', 'trưởng ca', 'tổ trưởng', 'nhóm trưởng', 'nhân viên', 'chuyên viên', 'kỹ sư', 'kỹ thuật viên'];

function isLosEligibleTitle_(title) {
  if (!title) return false;
  var t = title.toLowerCase();
  for (var i=0; i<LOS_INELIGIBLE_KEYWORDS.length; i++) { if (t.indexOf(LOS_INELIGIBLE_KEYWORDS[i])>=0) return false; }
  for (var j=0; j<LOS_ELIGIBLE_LEVEL_KEYWORDS.length; j++) { if (t.indexOf(LOS_ELIGIBLE_LEVEL_KEYWORDS[j])>=0) return true; }
  return false;
}

function getLosEligibleLeaders() {
  try {
    var ss = getLosSpreadsheet_();
    // Ưu tiên 1: LOS_LEADER sheet
    var losSheet = ss.getSheetByName('LOS_LEADER');
    if (losSheet && losSheet.getLastRow() > 1) {
      return mhGetSheetData_(ss, 'LOS_LEADER').filter(function(r){ return String(r.active||'').toLowerCase()!=='false'; });
    }
    // Ưu tiên 2: PEOPLE sheet từ Meeting Hub
    var peopleSheet = ss.getSheetByName('PEOPLE');
    if (peopleSheet && peopleSheet.getLastRow() > 1) {
      return mhGetSheetData_(ss, 'PEOPLE')
        .filter(function(r){ return String(r.active||'true').toLowerCase()!=='false' && isLosEligibleTitle_(r.title||''); })
        .map(function(r, i){ return {
          leaderId: r.email||('L'+String(i+1).padStart(3,'0')),
          email: r.email||'', fullName: r.fullName||'',
          title: r.title||'', department: r.department||'',
          managementLevel: r.title||'', active: true
        }; });
    }
    return [];
  } catch(e) { Logger.log('getLosEligibleLeaders error: ' + e.message); return []; }
}

/* ── Main data function ── */
function getLosData() {
  try {
    var ss = getLosSpreadsheet_();
    var leaders = getLosEligibleLeaders();
    if (!leaders || leaders.length === 0) {
      return { leaders: [], _isDemo: false, _message: 'Chưa có dữ liệu Leader trong LOS_LEADER hoặc PEOPLE sheet.' };
    }

    var scores    = mhGetSheetData_(ss, 'LOS_SCORE') || [];
    var coaching  = mhGetSheetData_(ss, 'LOS_COACHING') || [];
    var reviews   = mhGetSheetData_(ss, 'LOS_REVIEW') || [];
    var losMtgs   = mhGetSheetData_(ss, 'LOS_MEETING') || [];
    var meScores  = mhGetSheetData_(ss, 'LOS_ME_SCORE') || [];

    leaders = leaders.map(function(l) {
      var lid = l.leaderId || l.email || '';
      // Latest score
      var lScores = scores.filter(function(s){ return String(s.leaderId||'')===String(lid); })
        .sort(function(a,b){ return new Date(b.reviewDate||0)-new Date(a.reviewDate||0); });
      var cur  = lScores[0]  || {};
      var prev = lScores[1] || {};
      var losCur  = Number(cur.losVerified||0)  || Number(cur.losAiProposed||0)  || 0;
      var meCur   = Number(cur.meVerified||0)   || Number(cur.meAiProposed||0)   || 0;
      var losPrev = Number(prev.losVerified||0) || Number(prev.losAiProposed||0) || 0;
      var mePrev  = Number(prev.meVerified||0)  || Number(prev.meAiProposed||0)  || 0;

      // Gate reviews
      var lReviews = reviews.filter(function(r){ return String(r.leaderId||'')===lid; })
        .sort(function(a,b){ return new Date(a.reviewDate||0)-new Date(b.reviewDate||0); });

      // Top Gap
      var lCoach = coaching.filter(function(c){ return String(c.leaderId||'')===lid && (c.status||'').toLowerCase()!=='done'; });
      var topGap = lCoach[0] ? lCoach[0].gapTitle||'' : '';

      // Meetings ME
      var lMtgs = losMtgs.filter(function(m){ return String(m.leaderId||'')===lid; });
      var meetingIds = lMtgs.map(function(m){ return m.meetingId; });
      var lMe = meScores.filter(function(s){ return meetingIds.indexOf(String(s.meetingId||''))>=0; });
      var meVals = lMe.map(function(s){ return Number(s.meVerified||0)||Number(s.meAiProposed||0)||0; }).filter(Boolean);
      var meAvg30 = meVals.length>0 ? meVals.reduce(function(a,b){return a+b;},0)/meVals.length : 0;

      return {
        leaderId: lid,
        email:    l.email||'',
        fullName: l.fullName||l.email||lid,
        title:    l.title||'',
        managementLevel: l.managementLevel||l.title||'',
        company:  l.company||'',
        plant:    l.plant||'',
        division: l.division||'',
        department: l.department||'',
        currentGate:  l.currentGate||cur.gate||'Baseline',
        losStartDate: l.losStartDate||'',
        lastReviewDate:cur.reviewDate||l.lastReviewDate||'',
        nextReviewDate:l.nextReviewDate||'',
        losVerified:  Number(cur.losVerified||0),
        losAi:        Number(cur.losAiProposed||0),
        meVerified:   Number(cur.meVerified||0)||meAvg30,
        meAi:         0,
        losPrev:      losPrev,
        mePrev:       mePrev,
        losDelta:     losCur - losPrev,
        meDelta:      meCur  - mePrev,
        topGap:       topGap,
        topGaps:      lCoach.slice(0,3).map(function(c){ return {title:c.gapTitle, desc:c.gapDescription||''}; }),
        gateReviews:  lReviews.map(function(r){ return {gate:r.gate,reviewDate:r.reviewDate,losScore:Number(r.losScore||0),meScore:Number(r.meScore||0),reviewNote:r.reviewNote||''}; }),
        coaching:     lCoach.slice(0,5).map(function(c){ return {gapTitle:c.gapTitle,coachingAction:c.coachingAction,coachName:c.coachName,deadline:c.deadline,status:c.status,beforeScore:c.beforeScore,afterScore:c.afterScore}; }),
        losScoreDetail: cur,
        meetings:     lMtgs.slice(0,20).map(function(m){ return {meetingId:m.meetingId,meetingDate:m.meetingDate,meetingTitle:m.meetingTitle||m.meetingId,leaderRole:m.leaderRole,meScore:0}; }),
        evidenceSummary: { minutes:0, actions:0, dashboard:0, gemba:0 }
      };
    });

    return { leaders: leaders, updatedAt: new Date().toISOString() };
  } catch(e) {
    Logger.log('getLosData error: ' + e.message);
    return { leaders: [], error: e.message };
  }
}

function getLeadershipCard(leaderId) {
  try {
    var data = getLosData();
    var leader = (data.leaders||[]).filter(function(l){ return l.leaderId===leaderId; })[0];
    if (!leader) return { error: 'Không tìm thấy Leader: ' + leaderId };
    return { leader: leader };
  } catch(e) { return { error: e.message }; }
}

function getLosScorecard(leaderId) {
  try {
    var ss = getLosSpreadsheet_();
    var scores = mhGetSheetData_(ss, 'LOS_SCORE') || [];
    var lScores = scores.filter(function(s){ return String(s.leaderId||'')===leaderId; })
      .sort(function(a,b){ return new Date(b.reviewDate||0)-new Date(a.reviewDate||0); });
    return { scorecard: lScores[0] || null, history: lScores };
  } catch(e) { return { error: e.message }; }
}

function getLeaderMeData(leaderId) {
  try {
    var ss = getLosSpreadsheet_();
    var mtgs = (mhGetSheetData_(ss,'LOS_MEETING')||[]).filter(function(m){ return m.leaderId===leaderId; });
    var mids = mtgs.map(function(m){ return m.meetingId; });
    var me   = (mhGetSheetData_(ss,'LOS_ME_SCORE')||[]).filter(function(s){ return mids.indexOf(s.meetingId)>=0; });
    return { meetings: mtgs, meScores: me };
  } catch(e) { return { error: e.message }; }
}

function getMeetingMeScore(meetingId, leaderId) {
  try {
    var ss = getLosSpreadsheet_();
    var me = (mhGetSheetData_(ss,'LOS_ME_SCORE')||[]).filter(function(s){ return s.meetingId===meetingId && (!leaderId||s.leaderId===leaderId); });
    return { meScore: me[0] || null };
  } catch(e) { return { error: e.message }; }
}

function getLosTrend(leaderId) {
  try {
    var ss = getLosSpreadsheet_();
    var reviews = (mhGetSheetData_(ss,'LOS_REVIEW')||[]).filter(function(r){ return r.leaderId===leaderId; })
      .sort(function(a,b){ return new Date(a.reviewDate||0)-new Date(b.reviewDate||0); });
    return { trend: reviews };
  } catch(e) { return { error: e.message }; }
}

function getLeaderGateReviews(leaderId) {
  try {
    var ss = getLosSpreadsheet_();
    var reviews = (mhGetSheetData_(ss,'LOS_REVIEW')||[]).filter(function(r){ return r.leaderId===leaderId; });
    return { gateReviews: reviews };
  } catch(e) { return { error: e.message }; }
}

function getLeaderCoaching(leaderId) {
  try {
    var ss = getLosSpreadsheet_();
    var coaching = (mhGetSheetData_(ss,'LOS_COACHING')||[]).filter(function(c){ return c.leaderId===leaderId; });
    return { coaching: coaching };
  } catch(e) { return { error: e.message }; }
}

function getLeaderEvidence(leaderId, filter) {
  try {
    var ss = getLosSpreadsheet_();
    var ev = (mhGetSheetData_(ss,'LOS_EVIDENCE')||[]).filter(function(e){ return e.leaderId===leaderId && (!filter||e.sourceType===filter); });
    return { evidence: ev };
  } catch(e) { return { error: e.message }; }
}

/* ── Calculations (pure functions) ── */
var LOS_CRITERIA_WEIGHTS = {dataDriven:1.5,planning:1.5,decision:1.5,ownership:1.5,delegation:1.0,problemSolving:1.0,followThrough:1.0,coaching:1.0};
var ME_CRITERIA_WEIGHTS  = {dataReady:1.5,problem:1.5,rootCause:1.5,decision:2.0,actionPicDeadline:2.0,followUp:1.5};

function calculateLosScore_(scoreObj) {
  if (!scoreObj) return 0;
  var total = 0;
  Object.keys(LOS_CRITERIA_WEIGHTS).forEach(function(k) {
    var v = Number(scoreObj[k+'Verified']||0) || Number(scoreObj[k+'Ai']||0) || 0;
    total += v * LOS_CRITERIA_WEIGHTS[k];
  });
  return Math.min(10, Math.round(total*10)/10);
}

function calculateMeScore_(scoreObj) {
  if (!scoreObj) return 0;
  var total = 0;
  Object.keys(ME_CRITERIA_WEIGHTS).forEach(function(k) {
    var v = Number(scoreObj[k+'Verified']||0) || Number(scoreObj[k+'Ai']||0) || 0;
    total += v * ME_CRITERIA_WEIGHTS[k];
  });
  var raw = Math.min(10, Math.round(total*10)/10);
  return applyMeDecisionActionCap_(raw, scoreObj);
}

function applyMeDecisionActionCap_(rawScore, scoreObj) {
  var dec = Number(scoreObj.decisionVerified||0) || Number(scoreObj.decisionAi||0) || 0;
  var act = Number(scoreObj.actionPicDeadlineVerified||0) || Number(scoreObj.actionPicDeadlineAi||0) || 0;
  if (dec < 1.0 || act < 1.0) return Math.min(5.0, rawScore);
  return rawScore;
}

function getLosLeaderStatus_(leaderData) {
  var los = leaderData.losVerified || leaderData.losAi || 0;
  var me  = leaderData.meVerified  || leaderData.meAi  || 0;
  var losDelta = parseFloat(leaderData.losDelta || 0);
  var meDelta  = parseFloat(leaderData.meDelta  || 0);
  if (los === 0 && me === 0) return 'yellow';
  if (los >= 7 && me >= 7 && losDelta >= 0 && meDelta >= 0) return 'green';
  if (los < 5 || me < 5 || losDelta < -1 || meDelta < -1) return 'red';
  return 'yellow';
}

/* ── Write functions ── */
function saveLosVerifiedScore(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_SCORE');
    if (!sheet) return { error: 'Sheet LOS_SCORE không tồn tại. Chạy setupLosDataStore() trước.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var userEmail = Session.getActiveUser().getEmail();
    var scoreId = 'LOS-' + Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss');
    var data = Object.assign({ scoreId:scoreId, createdAt:now, updatedAt:now, verifiedBy:userEmail, verifiedDate:now }, payload);
    var row = headers.map(function(h){ return data[h]!==undefined?data[h]:''; });
    sheet.appendRow(row);
    return { success:true, scoreId:scoreId };
  } catch(e) { return { error: e.message }; }
}

function saveMeVerifiedScore(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_ME_SCORE');
    if (!sheet) return { error: 'Sheet LOS_ME_SCORE không tồn tại.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var userEmail = Session.getActiveUser().getEmail();
    var data = Object.assign({ meScoreId:'ME-'+Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss'), createdAt:now, updatedAt:now, verifiedBy:userEmail, verifiedDate:now }, payload);
    sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
    return { success:true };
  } catch(e) { return { error: e.message }; }
}

function saveLosReview(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_REVIEW');
    if (!sheet) return { error: 'Sheet LOS_REVIEW không tồn tại.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var data = Object.assign({ reviewId:'REV-'+Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss'), createdAt:now }, payload);
    sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
    return { success:true };
  } catch(e) { return { error: e.message }; }
}

function createLosCoachingAction(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_COACHING');
    if (!sheet) return { error: 'Sheet LOS_COACHING không tồn tại.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var data = Object.assign({ coachingId:'CA-'+Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss'), createdAt:now, updatedAt:now }, payload);
    sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
    return { success:true };
  } catch(e) { return { error: e.message }; }
}

function updateLosCoachingAction(coachingId, updateData) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_COACHING');
    if (!sheet) return { error: 'Sheet LOS_COACHING không tồn tại.' };
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var idCol = headers.indexOf('coachingId');
    var rowIdx = mhFindRowIndex_(sheet, idCol, coachingId);
    if (rowIdx < 0) return { error: 'Không tìm thấy coachingId: ' + coachingId };
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    Object.assign(updateData, { updatedAt: now, updatedBy: Session.getActiveUser().getEmail() });
    Object.keys(updateData).forEach(function(k) {
      var col = headers.indexOf(k);
      if (col >= 0) sheet.getRange(rowIdx, col+1).setValue(updateData[k]);
    });
    return { success:true };
  } catch(e) { return { error: e.message }; }
}

function linkLosEvidence(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_EVIDENCE');
    if (!sheet) return { error: 'Sheet LOS_EVIDENCE không tồn tại.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var data = Object.assign({ evidenceId:'EV-'+Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss'), createdAt:now }, payload);
    sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
    return { success:true };
  } catch(e) { return { error: e.message }; }
}

/* ── setupLosDataStore ── */
function setupLosDataStore() {
  try {
    var ss = getLosSpreadsheet_();
    var schemas = {
      LOS_LEADER:   ['leaderId','employeeId','email','fullName','title','managementLevel','company','plant','division','department','photoUrl','losStartDate','currentGate','lastReviewDate','nextReviewDate','active','createdAt','updatedAt'],
      LOS_SCORE:    ['scoreId','leaderId','gate','reviewDate','dataDrivenAi','dataDrivenVerified','planningAi','planningVerified','decisionAi','decisionVerified','ownershipAi','ownershipVerified','delegationAi','delegationVerified','problemSolvingAi','problemSolvingVerified','followThroughAi','followThroughVerified','coachingAi','coachingVerified','losAiProposed','losVerified','verifiedBy','verifiedDate','reviewNote','createdAt','updatedAt'],
      LOS_MEETING:  ['losMeetingId','leaderId','meetingId','leaderRole','meetingDate','meetingTitle','evaluated','createdAt'],
      LOS_ME_SCORE: ['meScoreId','leaderId','meetingId','dataReadyAi','dataReadyVerified','problemAi','problemVerified','rootCauseAi','rootCauseVerified','decisionAi','decisionVerified','actionPicDeadlineAi','actionPicDeadlineVerified','followUpAi','followUpVerified','meAiProposed','meVerified','capApplied','capReason','verifiedBy','verifiedDate','reviewNote','createdAt','updatedAt'],
      LOS_COACHING: ['coachingId','leaderId','sourceType','sourceCriterion','gapTitle','gapDescription','behaviorToChange','coachingAction','nextPractice','coachEmail','coachName','coachingDate','deadline','status','beforeScore','afterScore','nextReviewGate','result','createdAt','updatedAt'],
      LOS_EVIDENCE: ['evidenceId','leaderId','sourceType','sourceId','criterionType','criterionName','title','url','note','evidenceDate','createdAt'],
      LOS_REVIEW:   ['reviewId','leaderId','gate','reviewDate','losScore','meScore','topGap1','topGap2','topGap3','developmentAction','behaviorChanged','gapRemaining','trend','reviewNote','reviewer','nextReviewDate','createdAt']
    };
    var result = {};
    Object.keys(schemas).forEach(function(name) {
      var headers = schemas[name];
      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#042f23').setFontColor('white');
        result[name] = 'created';
      } else {
        var existing = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
        var missing = headers.filter(function(h){ return existing.indexOf(h)<0; });
        if (missing.length>0) {
          missing.forEach(function(h) {
            var col = existing.length + 1;
            sheet.getRange(1, col).setValue(h).setFontWeight('bold').setBackground('#042f23').setFontColor('white');
            existing.push(h);
          });
          result[name] = 'updated (' + missing.length + ' cols added)';
        } else { result[name] = 'ok'; }
      }
    });
    return { success:true, result:result };
  } catch(e) { return { error: e.message }; }
}

/* ── testLosConnection ── */
function testLosConnection() {
  try {
    var ss = getLosSpreadsheet_();
    var cfg = getLosConfig();
    var sheets = ['LOS_LEADER','LOS_SCORE','LOS_MEETING','LOS_ME_SCORE','LOS_COACHING','LOS_EVIDENCE','LOS_REVIEW'];
    var sheetStatus = {};
    sheets.forEach(function(name) {
      var s = ss.getSheetByName(name);
      sheetStatus[name] = s ? 'exists (' + Math.max(0,s.getLastRow()-1) + ' rows)' : 'missing';
    });
    var leaders = getLosEligibleLeaders();
    return {
      config: cfg, spreadsheetName: ss.getName(),
      currentUser: Session.getActiveUser().getEmail(),
      eligibleLeaders: leaders.length,
      sheets: sheetStatus
    };
  } catch(e) { return { error: e.message }; }
}

/* ── importLosLeaders — Upsert leaders from Excel or sync from PEOPLE ── */
function importLosLeaders(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_LEADER');
    if (!sheet) return { error: 'Sheet LOS_LEADER chưa tồn tại. Chạy setupLosDataStore() trước.' };
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');

    var leaders;
    if (payload.mode === 'sync') {
      leaders = getLosEligibleLeaders().map(function(l) { return {
        employeeId: l.employeeId||l.email||'',
        email: l.email||'', fullName: l.fullName||'',
        title: l.title||'', managementLevel: l.managementLevel||l.title||'',
        company: l.company||'', plant: l.plant||'', division: l.division||'',
        department: l.department||'', losStartDate: l.losStartDate||'',
        currentGate: l.currentGate||'Baseline', active: true
      }; });
    } else {
      leaders = payload.leaders || [];
    }

    var existing = mhGetSheetData_(ss, 'LOS_LEADER') || [];
    var imported = 0, updated = 0, skipped = 0;

    leaders.forEach(function(l) {
      if (!l.email && !l.employeeId) { skipped++; return; }
      var existIdx = -1;
      var existRow = -1;
      for (var i=0; i<existing.length; i++) {
        if (l.employeeId && String(existing[i].employeeId||'')===String(l.employeeId)) { existIdx=i; break; }
        if (l.email && String(existing[i].email||'').toLowerCase()===String(l.email||'').toLowerCase()) { existIdx=i; }
      }
      if (existIdx >= 0) {
        var sheetData = sheet.getDataRange().getValues();
        for (var r=1; r<sheetData.length; r++) {
          var eId = String(sheetData[r][headers.indexOf('employeeId')]||'');
          var em  = String(sheetData[r][headers.indexOf('email')]||'').toLowerCase();
          if ((l.employeeId && eId===String(l.employeeId)) || (l.email && em===l.email.toLowerCase())) {
            existRow = r+1; break;
          }
        }
        if (existRow > 0) {
          l.updatedAt = now;
          headers.forEach(function(h,ci) { if (l[h]!==undefined) sheet.getRange(existRow,ci+1).setValue(l[h]); });
          updated++;
        }
      } else {
        if (!isLosEligibleTitle_(l.title||'')) { skipped++; return; }
        var leaderId = 'LDR-' + Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMddHHmmss') + '-' + imported;
        var data = Object.assign({ leaderId:leaderId, createdAt:now, updatedAt:now }, l);
        sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
        imported++;
      }
    });
    return { success:true, imported:imported, updated:updated, skipped:skipped };
  } catch(e) { Logger.log('importLosLeaders error: ' + e.message); return { error: e.message }; }
}

/* ── importLosScores — Bulk import LOS scores from Excel ── */
function importLosScores(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_SCORE');
    if (!sheet) return { error: 'Sheet LOS_SCORE chưa tồn tại.' };
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var rows = payload.scores || [];
    var userEmail = Session.getActiveUser().getEmail();
    var imported = 0;
    var keyMap = {
      'ĐIỀU_HÀNH_BẰNG_DỮ_LIỆU':'dataDrivenVerified','LẬP_KẾ_HOẠCH_ƯU_TIÊN':'planningVerified',
      'RA_QUYẾT_ĐỊNH':'decisionVerified','NHẬN_TRÁCH_NHIỆM':'ownershipVerified',
      'PHÂN_QUYỀN_GIAO_PIC':'delegationVerified','GIẢI_QUYẾT_VẤN_ĐỀ':'problemSolvingVerified',
      'THEO_ĐẾN_KẾT_QUẢ':'followThroughVerified','PHÁT_TRIỂN_ĐỘI_NGŨ':'coachingVerified',
      'NHẬN_XÉT':'reviewNote','NGƯỜI_XÁC_NHẬN':'verifiedBy',
      'MỐC_ĐÁNH_GIÁ':'gate','NGÀY_ĐÁNH_GIÁ':'reviewDate',
      'EMPLOYEE_ID':'employeeId','EMAIL':'email'
    };
    rows.forEach(function(r) {
      var mapped = {};
      Object.keys(r).forEach(function(k) { mapped[keyMap[k]||k] = r[k]; });
      var leaders = mhGetSheetData_(ss, 'LOS_LEADER') || [];
      var ldr = leaders.filter(function(l) {
        return (mapped.employeeId && l.employeeId===String(mapped.employeeId)) ||
               (mapped.email && String(l.email||'').toLowerCase()===String(mapped.email||'').toLowerCase());
      })[0];
      if (!ldr) return;
      var scoreId = 'LOS-' + Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd-HHmmss') + '-' + imported;
      var data = Object.assign({ scoreId:scoreId, leaderId:ldr.leaderId||ldr.email, createdAt:now, updatedAt:now, verifiedBy:mapped.verifiedBy||userEmail }, mapped);
      sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
      imported++;
    });
    return { success:true, imported:imported };
  } catch(e) { return { error: e.message }; }
}

/* ── saveLosMeetingLinks — Lưu nhiều liên kết một lúc ── */
function saveLosMeetingLinks(payload) {
  try {
    assertMhUser_();
    var ss = getLosSpreadsheet_();
    var sheet = ss.getSheetByName('LOS_MEETING');
    if (!sheet) return { error: 'Sheet LOS_MEETING chưa tồn tại.' };
    var headers = sheet.getLastColumn()>0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0] : [];
    var now = Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyy-MM-dd HH:mm:ss');
    var links = payload.links || [];
    var saved = 0;
    var existing = mhGetSheetData_(ss, 'LOS_MEETING') || [];
    links.forEach(function(lk) {
      var dup = existing.filter(function(e){ return e.leaderId===lk.leaderId && e.meetingId===lk.meetingId; })[0];
      if (dup) return;
      var id = 'LM-' + Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMddHHmmss') + '-' + saved;
      var data = Object.assign({ losMeetingId:id, createdAt:now, evaluated:false }, lk);
      sheet.appendRow(headers.map(function(h){ return data[h]!==undefined?data[h]:''; }));
      existing.push(data);
      saved++;
    });
    return { success:true, saved:saved };
  } catch(e) { return { error: e.message }; }
}

/* ── getMeetingsForLosLinking — Gợi ý cuộc họp cho Leader từ Meeting Hub ── */
function getMeetingsForLosLinking(leaderId, leaderEmail) {
  try {
    var ss = getLosSpreadsheet_();
    var meetings = mhGetSheetData_(ss, 'MEETINGS') || [];
    if (!meetings.length) {
      var mhCfg = getMeetingHubConfig();
      if (mhCfg && mhCfg.spreadsheetId) {
        try {
          var mhSs = SpreadsheetApp.openById(mhCfg.spreadsheetId);
          meetings = mhGetSheetData_(mhSs, 'MEETINGS') || [];
        } catch(e2) {}
      }
    }
    var le = (leaderEmail||'').toLowerCase();
    var result = meetings.map(function(m) {
      var chair = String(m.chairpersonEmail||'').toLowerCase();
      var parts = String(m.participantsEmails||'').toLowerCase();
      var sec   = String(m.secretaryEmail||'').toLowerCase();
      var exactMatch = le && (chair.indexOf(le)>=0 || parts.indexOf(le)>=0 || sec.indexOf(le)>=0);
      return {
        meetingId:         m.meetingId||m.id||'',
        title:             m.title||m.meetingTitle||'',
        meetingDate:       m.meetingDate||m.date||'',
        chairpersonEmail:  m.chairpersonEmail||'',
        participantsEmails:m.participantsEmails||'',
        secretaryEmail:    m.secretaryEmail||'',
        minutesFileUrl:    m.minutesFileUrl||'',
        exactMatch:        exactMatch
      };
    }).sort(function(a,b) { return (b.exactMatch?1:0)-(a.exactMatch?1:0) || new Date(b.meetingDate)-new Date(a.meetingDate); });
    return { meetings: result };
  } catch(e) { return { meetings:[], error: e.message }; }
}

/* ── getLosDataSourceStatus — Tổng hợp trạng thái nguồn dữ liệu ── */
function getLosDataSourceStatus() {
  try {
    var ss = getLosSpreadsheet_();
    var leaders  = (mhGetSheetData_(ss,'LOS_LEADER')||[]).filter(function(r){ return String(r.active||'').toLowerCase()!=='false'; });
    var scores   = mhGetSheetData_(ss,'LOS_SCORE')    || [];
    var linked   = mhGetSheetData_(ss,'LOS_MEETING')  || [];
    var me       = mhGetSheetData_(ss,'LOS_ME_SCORE') || [];
    var coaching = mhGetSheetData_(ss,'LOS_COACHING') || [];
    var meetings = mhGetSheetData_(ss,'MEETINGS')     || [];
    var scoredIds = {};
    scores.forEach(function(s){ if (s.leaderId) scoredIds[s.leaderId] = true; });
    var openCoaching = coaching.filter(function(c){ return (c.status||'').toLowerCase()!=='hoàn thành' && (c.status||'').toLowerCase()!=='done'; }).length;
    return {
      leaders:        leaders.length,
      scored:         Object.keys(scoredIds).length,
      mhMeetings:     meetings.length,
      linkedMeetings: linked.length,
      meScored:       me.length,
      openCoaching:   openCoaching
    };
  } catch(e) { return { error: e.message }; }
}

/* ═══ END LOS — LEADERSHIP OPERATING SYSTEM ═══ */

/* ════════════════════════════════════════════════════════════
   CLC — CHI PHÍ LỖI CHẤT LƯỢNG (Quality Defect/Loss Cost)
   Tables: CLC_COMMITMENT, CLC_WEEKLY_PIC_REVIEW
   ════════════════════════════════════════════════════════════ */

/* ── Sheet helpers ── */
var CLC_SS_PROP_ = 'CLC_SPREADSHEET_ID';
var CLC_COMMITMENT_SHEET    = 'CLC_COMMITMENT';
var CLC_WEEKLY_REVIEW_SHEET = 'CLC_WEEKLY_PIC_REVIEW';

function getClcSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty(CLC_SS_PROP_);
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function clcGetSheetData_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  var rows    = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  return rows.map(function(r) {
    var obj = {};
    headers.forEach(function(h, i) { if (h) obj[h] = r[i]; });
    return obj;
  });
}

function clcUpsertRow_(ss, sheetName, idField, idValue, rowData) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return { success: false, error: 'Sheet ' + sheetName + ' không tồn tại' };
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var lastRow = sh.getLastRow();
  /* find existing */
  var idColIdx = headers.indexOf(idField);
  var existingRow = -1;
  if (idColIdx >= 0 && lastRow >= 2) {
    var ids = sh.getRange(2, idColIdx + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(idValue)) { existingRow = i + 2; break; }
    }
  }
  var rowValues = headers.map(function(h) {
    if (rowData.hasOwnProperty(h)) return rowData[h];
    return '';
  });
  if (existingRow > 0) {
    sh.getRange(existingRow, 1, 1, headers.length).setValues([rowValues]);
    return { success: true, action: 'updated', row: existingRow };
  } else {
    sh.appendRow(rowValues);
    return { success: true, action: 'inserted', row: sh.getLastRow() };
  }
}

/* ── getClcDashboard ── main GET endpoint ── */
function getClcDashboard() {
  try {
    var ss = getClcSpreadsheet_();
    var commitments = getClcCommitmentsData_(ss);
    var weeklyReviews = getClcWeeklyReviewData_(ss);
    var pics = buildClcPicList_(ss, commitments, weeklyReviews);
    return {
      ok: true,
      pics: pics,
      cases: [],
      correctiveActions: [],
      preventiveActions: [],
      commitments: commitments,
      weeklyReviews: weeklyReviews,
      _isDemo: false,
      updatedAt: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss')
    };
  } catch(e) { return { ok: false, error: e.message }; }
}

function getClcCommitmentsData_(ss) {
  var rows = clcGetSheetData_(ss, CLC_COMMITMENT_SHEET);
  return rows.filter(function(r) { return r.commitmentId; });
}

function getClcWeeklyReviewData_(ss) {
  var rows = clcGetSheetData_(ss, CLC_WEEKLY_REVIEW_SHEET);
  return rows.filter(function(r) { return r.reviewId; });
}

function buildClcPicList_(ss, commitments, weeklyReviews) {
  /* Build PIC roster from PEOPLE sheet */
  var people = [];
  try {
    var sh = ss.getSheetByName('PEOPLE');
    if (!sh || sh.getLastRow() < 2) sh = ss.getSheetByName('MASTER_PIC');
    if (sh && sh.getLastRow() >= 2) {
      var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      var rows    = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      people = rows.map(function(r) {
        var obj = {};
        headers.forEach(function(h, i) { if (h) obj[String(h).trim()] = r[i]; });
        return obj;
      }).filter(function(p) { return String(p.active || p.Active || 'true').toLowerCase() !== 'false'; });
    }
  } catch(e) { people = []; }

  /* Build PIC objects */
  return people.map(function(p) {
    var email = p.email || p.Email || '';
    var picId = p.picId || p.PicId || email.split('@')[0] || '';
    var picCommits = commitments.filter(function(c) { return c.picId === picId || c.picEmail === email; });
    var picReviews = weeklyReviews.filter(function(r) { return r.picId === picId || r.picEmail === email; });
    return {
      picId:      picId,
      picName:    p.fullName || p.name || p.Name || email,
      picEmail:   email,
      department: p.department || p.Department || '',
      factory:    p.factory    || p.Factory    || '',
      title:      p.title      || p.Title      || '',
      kpiWeekly:  { k1:[0,0,0,0], k2:[0,0,0,0], k3:[0,0,0,0], k4:[0,0,0,0], k5:[0,0,0,0] },
      weeklyScores: [0,0,0,0],
      status: 'N/A', trend: 'flat'
    };
  });
}

/* ── saveClcCommitment ── */
function saveClcCommitment(payload) {
  try {
    if (!payload.commitmentId) return { success: false, error: 'commitmentId bắt buộc' };
    var ss = getClcSpreadsheet_();
    payload._savedAt = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    payload._savedBy = Session.getActiveUser().getEmail();
    return clcUpsertRow_(ss, CLC_COMMITMENT_SHEET, 'commitmentId', payload.commitmentId, payload);
  } catch(e) { return { success: false, error: e.message }; }
}

/* ── confirmClcCommitment — Level 1 electronic signature ── */
function confirmClcCommitment(payload) {
  try {
    if (!payload.commitmentId) return { success: false, error: 'commitmentId bắt buộc' };
    var ss = getClcSpreadsheet_();
    var existing = getClcCommitmentsData_(ss);
    var row = existing.find(function(r) { return r.commitmentId === payload.commitmentId; });
    if (!row) return { success: false, error: 'Không tìm thấy commitmentId: ' + payload.commitmentId };
    row.signed         = true;
    row.signedBy       = payload.signedBy || Session.getActiveUser().getEmail();
    row.signedDate     = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
    row.status         = 'Đã cam kết';
    row._confirmedAt   = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    if (payload.signatureEvidenceUrl) row.signatureEvidenceUrl = payload.signatureEvidenceUrl;
    return clcUpsertRow_(ss, CLC_COMMITMENT_SHEET, 'commitmentId', payload.commitmentId, row);
  } catch(e) { return { success: false, error: e.message }; }
}

/* ── saveClcWeeklyReview ── */
function saveClcWeeklyReview(payload) {
  try {
    if (!payload.reviewId) return { success: false, error: 'reviewId bắt buộc' };
    var ss = getClcSpreadsheet_();
    payload._savedAt = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    payload._savedBy = Session.getActiveUser().getEmail();
    return clcUpsertRow_(ss, CLC_WEEKLY_REVIEW_SHEET, 'reviewId', payload.reviewId, payload);
  } catch(e) { return { success: false, error: e.message }; }
}

/* ── getClcDataSourceStatus ── */
function getClcDataSourceStatus() {
  try {
    var ss = getClcSpreadsheet_();
    var commitments  = (ss.getSheetByName(CLC_COMMITMENT_SHEET)    || { getLastRow: function(){ return 0; } }).getLastRow();
    var reviews      = (ss.getSheetByName(CLC_WEEKLY_REVIEW_SHEET) || { getLastRow: function(){ return 0; } }).getLastRow();
    return { ok: true, commitments: Math.max(0, commitments - 1), weeklyReviews: Math.max(0, reviews - 1) };
  } catch(e) { return { ok: false, error: e.message }; }
}

/* ── setupClcDataStore — tạo 2 sheet nếu chưa có ── */
function setupClcDataStore() {
  try {
    var ss = getClcSpreadsheet_();
    var created = [];

    var commitHeaders = [
      'commitmentId','caseId','actionId','picId','picName','department',
      'commitmentText','commitmentDate','deadline','originalDeadline',
      'expectedOutput','status','signed','signedBy','signedDate',
      'signatureEvidenceUrl','actualResult','evidenceUrl',
      'verifiedBy','verifiedDate','carryForwardFrom','overdueDays',
      '_savedAt','_savedBy','_confirmedAt'
    ];
    if (!ss.getSheetByName(CLC_COMMITMENT_SHEET)) {
      var sh1 = ss.insertSheet(CLC_COMMITMENT_SHEET);
      sh1.getRange(1, 1, 1, commitHeaders.length).setValues([commitHeaders]);
      sh1.getRange(1, 1, 1, commitHeaders.length).setBackground('#0f2d1f').setFontColor('white').setFontWeight('bold');
      sh1.setFrozenRows(1);
      created.push(CLC_COMMITMENT_SHEET);
    }

    var reviewHeaders = [
      'reviewId','picId','year','month','week',
      'kpiMonthlyTarget','weekScore_k1','weekScore_k2','weekScore_k3','weekScore_k4','weekScore_k5',
      'weeklyScore','mtdActual','commitmentStatus',
      'overdueActions','repeatIssues','reviewNote',
      'nextWeekCommitment','reviewedBy','snapshotDate',
      '_savedAt','_savedBy'
    ];
    if (!ss.getSheetByName(CLC_WEEKLY_REVIEW_SHEET)) {
      var sh2 = ss.insertSheet(CLC_WEEKLY_REVIEW_SHEET);
      sh2.getRange(1, 1, 1, reviewHeaders.length).setValues([reviewHeaders]);
      sh2.getRange(1, 1, 1, reviewHeaders.length).setBackground('#0f2d1f').setFontColor('white').setFontWeight('bold');
      sh2.setFrozenRows(1);
      created.push(CLC_WEEKLY_REVIEW_SHEET);
    }

    return { ok: true, created: created, message: created.length > 0 ? 'Đã tạo: ' + created.join(', ') : 'Các sheet đã tồn tại' };
  } catch(e) { return { ok: false, error: e.message }; }
}

/* ═══ END CLC — CHI PHÍ LỖI CHẤT LƯỢNG ═══ */
