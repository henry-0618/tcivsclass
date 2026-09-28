/**
 * 學校教學正常化走廊巡堂系統 - 公版功能選單與初始化模組
 */

// 1. 當試算表開啟時，自動在上方建立功能選單
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏫 巡堂系統功能選單')
    .addItem('🚀 1. 首次系統初始化與工作表檢查', 'initializeSystem')
    .addItem('📋 2. 課表匯入格式規範說明', 'showScheduleGuide')
    .addItem('📱 3. 手機/平板巡堂網址部署教學', 'showDeploymentGuide')
    .addSeparator()
    .addItem('ℹ️ 關於本系統', 'showAbout')
    .addToUi();
}

// 2. 一鍵初始化與架構防呆檢查
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const requiredSheets = [
    { name: '全校課表', headers: ['學期', '班級', '星期', '節次', '科目', '任課教師'] },
    { name: '系統選單設定', headers: ['巡查身分職稱', '上課地點', '吻合狀態', '學習評量', '課堂秩序', '環境衛生', '資訊科技', '通報Email(教學異常)', '通報Email(設備報修)', '學校名稱'] },
    { name: '學期設定', headers: ['學期代碼', '學期名稱', '是否啟用', '建立時間'] },
    { name: '巡查紀錄', headers: ['紀錄ID', '學期', '巡查日期', '星期', '節次', '班級', '巡查身分職稱', '上課地點', '預定科目', '預定教師', '實授科目', '實授教師', '吻合狀態', '評量檢核', '秩序狀況', '環境狀況', '資訊運用', '事由說明', '巡堂觀察', '需維修待處理', '改善備註', '建立時間'] }
  ];
  
  let createdCount = 0;
  let statusReport = [];

  requiredSheets.forEach(item => {
    let sheet = ss.getSheetByName(item.name);
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
      sheet.appendRow(item.headers);
      sheet.getRange(1, 1, 1, item.headers.length)
           .setBackground('#1a73e8')
           .setFontColor('#ffffff')
           .setFontWeight('bold');
      sheet.setFrozenRows(1);
      createdCount++;
      statusReport.push(`➕ 自動補齊工作表：【${item.name}】`);
    } else {
      statusReport.push(`✅ 已存在：【${item.name}】`);
    }
  });

  const optionsSheet = ss.getSheetByName('系統選單設定');
  if (optionsSheet && typeof ensureSchoolNameField_ === 'function') ensureSchoolNameField_(optionsSheet);

  const msg = createdCount > 0 
    ? `系統已自動修復並建立 ${createdCount} 個缺少的工作表！\n\n檢查結果：\n${statusReport.join('\n')}`
    : `🎉 系統檢查全部通過！\n\n所有核心工作表結構健全：\n${statusReport.join('\n')}\n\n提醒：請先在「系統選單設定」的「學校名稱」欄位填寫校名，再確認通報信箱已更新為貴校承辦人信箱。`;

  ui.alert('【巡堂系統初始化結果】', msg, ui.ButtonSet.OK);
}

// 3. 課表格式引導
function showScheduleGuide() {
  const ui = SpreadsheetApp.getUi();
  const guideText = 
    "【全校課表 欄位規範】\n" +
    "請務必保留「全校課表」工作表前 6 欄，順序如下：\n" +
    "1. 學期（例：115-1）\n" +
    "2. 班級（例：一甲、二乙）\n" +
    "3. 星期（例：Mon, Tue, Wed, Thu, Fri）\n" +
    "4. 節次（例：1, 2, 3, 4, 5, 6, 7）\n" +
    "5. 科目（例：國語、數學、體育）\n" +
    "6. 任課教師（例：001 王大明）\n\n" +
    "💡 提示：排課完成後，可直接從排課軟體匯出後複製貼入該表。";
  
  ui.alert('📋 課表匯入格式說明', guideText, ui.ButtonSet.OK);
}

// 4. 手機 Web App 部署指引
function showDeploymentGuide() {
  const ui = SpreadsheetApp.getUi();
  const deployText = 
    "【如何取得手機巡堂網址】\n\n" +
    "注意：本範本不預先部署；請複製到自己的試算表後，再由各校管理者完成以下步驟。\n" +
    "請先在「系統選單設定」的「學校名稱」欄位第 2 列填寫學校名稱。\n\n" +
    "1. 在 Apps Script 編輯器右上角點擊「部署」➜「新增部署作業」。\n" +
    "2. 齒輪圖示選擇「網頁應用程式 (Web App)」。\n" +
    "3. 設定如下：\n" +
    "   - 執行身分：我 (Me)\n" +
    "   - 誰可以存取：所有人 (Anyone)\n" +
    "4. 點擊「部署」並完成初次授權。\n" +
    "5. 複製產生的「網頁應用程式網址」，傳送到手機或轉成 QR Code 即可隨身巡堂填報！";
  
  ui.alert('📱 手機巡堂介面發布教學', deployText, ui.ButtonSet.OK);
}

// 5. 關於系統
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'ℹ️ 關於教學正常化走廊巡堂系統',
    '本系統專為中小學走廊巡堂、教學正常化查核與設備通報設計。\n' +
    '資料完全保存於貴校獨立的 Google 雲端硬碟中，兼顧資安與行政效率。',
    ui.ButtonSet.OK
  );
}

/**
 * 本校 教學正常化與走廊巡堂系統 (Google Apps Script 後端)
 * 功能：
 * 1. 提供 Web App 介面（含行動端 Viewport 宣告）
 * 2. 試算表工作表自動初始化（學期設定、全校課表、巡查紀錄、系統選單設定）
 * 3. 學期動態管理（新增學期、切換啟用學期）
 * 4. 總課表 PDF 矩陣格式智慧解析與匯入（支援教師課表矩陣與班級課表矩陣）
 * 5. 系統選單與下拉變數 Google 試算表動態管理（職稱、地點、評量、常規、環境、媒材、通報Email）
 * 6. 巡查紀錄批次儲存與後端原生 MailApp 自動發信通報
 */

const DEFAULT_SCHOOL_NAME_ = "本校";

function getSchoolName_() {
  try {
    const options = getSystemOptions_();
    return String(options.schoolName || DEFAULT_SCHOOL_NAME_).trim() || DEFAULT_SCHOOL_NAME_;
  } catch (err) {
    return DEFAULT_SCHOOL_NAME_;
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(getSchoolName_() + ' 教學正常化與走廊巡堂系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------- 試算表結構初始化 ----------------------

function getSemesterSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('學期設定');
  if (!sheet) {
    sheet = ss.insertSheet('學期設定');
    sheet.appendRow(["學期代碼", "學期名稱", "是否啟用", "建立時間"]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.appendRow(["115-1", "115學年度第一學期", true, Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm")]);
  }
  return sheet;
}

function getTimetableSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('全校課表');
  if (!sheet) {
    sheet = ss.insertSheet('全校課表');
    sheet.appendRow(["學期", "班級", "星期", "節次", "科目", "任課教師"]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
    seedDefaultMasterSchedule(sheet);
  }
  return sheet;
}

function getInspectionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('巡查紀錄');
  if (!sheet) {
    sheet = ss.insertSheet('巡查紀錄');
    const headers = [
      "紀錄ID", "學期", "巡查日期", "星期", "節次", "班級", "巡查身分職稱", "上課地點",
      "預定科目", "預定教師", "實授科目", "實授教師", "吻合狀態",
      "評量檢核", "秩序狀況", "環境狀況", "資訊運用",
      "事由說明", "巡堂觀察", "需維修待處理", "改善備註", "建立時間"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ---------------------- 系統選單與下拉變數動態設定 ----------------------

function getOptionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('系統選單設定');
  if (!sheet) {
    sheet = ss.insertSheet('系統選單設定');
    const headers = [
      "巡查身分職稱", "上課地點", "吻合狀態", "學習評量", "課堂秩序", "環境衛生", "資訊科技", "通報Email(教學異常)", "通報Email(設備報修)", "學校名稱"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
    sheet.setFrozenRows(1);

    // 預設選項設定
    const defaults = [
      ["校長", "原班教室", "吻合", "正常教學無測驗", "良好", "良好", "觸控大屏/電子白板", "", ""],
      ["教導主任", "多功能教室", "調課", "隨堂形成性小考", "尚可", "尚可", "學生平板載具", "", ""],
      ["總務主任", "E化教室", "公假代理", "違規佔用藝能課考試", "待加強", "待改善", "簡報/投影機", "", ""],
      ["教學組長", "圓形教室", "異常不符", "跨科違規考試", "", "", "板書/傳統教具", "", ""],
      ["訓導組長", "語言教室", "", "", "", "", "無", "", ""],
      ["導護", "多元教室", "", "", "", "", "", "", ""],
      ["", "海涵樓", "", "", "", "", "", "", ""],
      ["", "育樂中心", "", "", "", "", "", "", ""],
      ["", "行政大樓", "", "", "", "", "", "", ""],
      ["", "操場", "", "", "", "", "", "", ""],
      ["", "籃球場", "", "", "", "", "", "", ""],
      ["", "其他", "", "", "", "", "", "", ""]
    ];
    sheet.getRange(2, 1, defaults.length, headers.length).setValues(defaults.map(row => row.concat([""])));
  }
  ensureSchoolNameField_(sheet);
  return sheet;
}

function ensureSchoolNameField_(sheet) {
  const cell = sheet.getRange(1, 10);
  if (String(cell.getValue() || "").trim() !== "學校名稱") {
    cell.setValue("學校名稱").setFontWeight("bold").setBackground("#e2e8f0");
  }
}

function getSystemOptions() {
  const options = getSystemOptions_();
  return {
    inspectorRoles: options.inspectorRoles,
    locations: options.locations,
    matchStatuses: options.matchStatuses,
    evalExams: options.evalExams,
    evalDisciplines: options.evalDisciplines,
    evalEnvs: options.evalEnvs,
    evalIts: options.evalIts,
    emailStatus: buildEmailStatus_(options),
    schoolName: options.schoolName
  };
}

function getSystemOptions_() {
  const defaultOptions = {
    inspectorRoles: ["校長", "教導主任", "總務主任", "教學組長", "訓導組長", "導護"],
    locations: ["原班教室", "多功能教室", "E化教室", "圓形教室", "語言教室", "多元教室", "海涵樓", "育樂中心", "行政大樓", "操場", "籃球場", "其他"],
    matchStatuses: ["吻合", "調課", "公假代理", "異常不符"],
    evalExams: ["正常教學無測驗", "隨堂形成性小考", "違規佔用藝能課考試", "跨科違規考試"],
    evalDisciplines: ["良好", "尚可", "待加強"],
    evalEnvs: ["良好", "尚可", "待改善"],
    evalIts: ["觸控大屏/電子白板", "學生平板載具", "簡報/投影機", "板書/傳統教具", "無"],
    schoolName: "本校",
    emailTeaching: [],
    emailRepair: [],
    invalidEmailTeachingCount: 0,
    invalidEmailRepairCount: 0,
    loadError: false
  };

  try {
    const sheet = getOptionsSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return defaultOptions;

    const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

    const extractCol = (colIdx, fallback) => {
      const list = [];
      for (let i = 0; i < values.length; i++) {
        const val = String(values[i][colIdx] || "").trim();
        if (val) list.push(val);
      }
      return list.length > 0 ? list : fallback;
    };

    const teachingEmails = parseEmailRecipients_(values, 7);
    const repairEmails = parseEmailRecipients_(values, 8);

    return {
      inspectorRoles: extractCol(0, defaultOptions.inspectorRoles),
      locations: extractCol(1, defaultOptions.locations),
      matchStatuses: extractCol(2, defaultOptions.matchStatuses),
      evalExams: extractCol(3, defaultOptions.evalExams),
      evalDisciplines: extractCol(4, defaultOptions.evalDisciplines),
      evalEnvs: extractCol(5, defaultOptions.evalEnvs),
      evalIts: extractCol(6, defaultOptions.evalIts),
      schoolName: values.map(row => String(row[9] || "").trim()).find(Boolean) || defaultOptions.schoolName,
      emailTeaching: teachingEmails.valid,
      emailRepair: repairEmails.valid,
      invalidEmailTeachingCount: teachingEmails.invalid.length,
      invalidEmailRepairCount: repairEmails.invalid.length,
      loadError: false
    };
  } catch (err) {
    Logger.log("讀取選單設定失敗（錯誤細節未回傳前端）");
    return Object.assign({}, defaultOptions, { loadError: true });
  }
}

function parseEmailRecipients_(values, colIdx) {
  const valid = [];
  const invalid = [];
  const seen = new Set();

  values.forEach(row => {
    String(row[colIdx] || "")
      .split(/[\s,;，；]+/)
      .map(item => item.trim().toLowerCase())
      .filter(Boolean)
      .forEach(email => {
        if (seen.has(email)) return;
        seen.add(email);
        if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) valid.push(email);
        else invalid.push(email);
      });
  });

  return { valid, invalid };
}

function buildEmailStatus_(options) {
  const teachingCount = options.emailTeaching.length;
  const repairCount = options.emailRepair.length;
  const invalidCount = options.invalidEmailTeachingCount + options.invalidEmailRepairCount;
  return {
    configured: !options.loadError && teachingCount > 0 && repairCount > 0 && invalidCount === 0,
    teachingCount: teachingCount,
    repairCount: repairCount,
    invalidCount: invalidCount,
    loadError: !!options.loadError
  };
}

function getEmailNotificationStatus() {
  try {
    const options = getSystemOptions_();
    const status = buildEmailStatus_(options);
    status.success = !options.loadError;
    status.remainingDailyQuota = MailApp.getRemainingDailyQuota();
    if (options.loadError) status.message = "無法讀取系統選單設定，Email 通報暫停。";
    return status;
  } catch (err) {
    Logger.log("檢查 Email 通報設定失敗: " + err.toString());
    return {
      success: false,
      configured: false,
      teachingCount: 0,
      repairCount: 0,
      invalidCount: 0,
      remainingDailyQuota: null,
      message: "無法讀取 Email 通報設定，請由指令碼擁有者重新授權 MailApp。"
    };
  }
}

// ---------------------- 學期管理模組 ----------------------

function getSemesterListAndActive() {
  const sheet = getSemesterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { activeSemester: "115-1", list: [{ code: "115-1", name: "115學年度第一學期", active: true }] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const list = [];
  let activeSemester = "";

  values.forEach(r => {
    const code = String(r[0]).trim();
    const name = String(r[1]).trim();
    const active = (r[2] === true || String(r[2]).toUpperCase() === "TRUE");
    list.push({ code, name, active });
    if (active) activeSemester = code;
  });

  if (!activeSemester && list.length > 0) activeSemester = list[0].code;
  return { activeSemester, list };
}

function getSemesterCodes_() {
  const sheet = getSemesterSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Set();
  return new Set(sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat()
    .map(v => String(v).trim()).filter(Boolean));
}

function setActiveSemester(targetCode) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中，請稍候再試" };
    const safeTargetCode = normalizeText_(targetCode, 30);
    const sheet = getSemesterSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "無學期資料" };

    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const targetExists = values.some(row => String(row[0]).trim() === safeTargetCode);
    if (!targetExists) return { success: false, message: "找不到指定學期" };
    const activeValues = values.map(row => [String(row[0]).trim() === safeTargetCode]);
    sheet.getRange(2, 3, activeValues.length, 1).setValues(activeValues);
    return { success: true, activeSemester: safeTargetCode };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function createNewSemester(code, name) {
  const safeCode = normalizeText_(code, 30);
  const safeName = normalizeText_(name, 120);
  if (!safeCode || !safeName) return { success: false, message: "學期代碼與名稱不得為空" };
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中，請稍候再試" };
    const sheet = getSemesterSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const codes = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat().map(c => String(c).trim());
      if (codes.includes(safeCode)) return { success: false, message: "該學期代碼已存在！" };
    }
    sheet.appendRow([
      sanitizeSheetCell_(safeCode),
      sanitizeSheetCell_(safeName),
      false,
      Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm")
    ]);
    return { success: true };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ---------------------- 總課表動態讀取與匯入 ----------------------

function getMasterScheduleBySemester(semesterCode) {
  const sheet = getTimetableSheet();
  const lastRow = sheet.getLastRow();
  const result = { "一甲": {}, "二甲": {}, "三甲": {}, "四甲": {}, "五甲": {}, "六甲": {} };
  const teachersSet = new Set();

  if (lastRow <= 1) return { schedule: result, teachers: ["代課老師"] };

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  values.forEach(r => {
    const sem = String(r[0]).trim();
    if (sem !== semesterCode) return;

    const cls = String(r[1]).trim();
    const day = String(r[2]).trim();
    const period = String(r[3]).trim();
    const sub = String(r[4]).trim();
    const t = String(r[5]).trim();

    if (!result[cls]) result[cls] = {};
    if (!result[cls][day]) result[cls][day] = {};

    result[cls][day][period] = { sub: sub, t: t };
    if (t) teachersSet.add(t);
  });

  const teachers = ["代課老師", ...Array.from(teachersSet).sort()];
  return { schedule: result, teachers: teachers };
}

/**
 * 矩陣課表智慧解析並匯入：
 * 支援輸入包含 [標的, 星期, 第1節, 第2節, 第3節, 第4節, 第5節, 第6節, 第7節] 之二維陣列
 * 格子可填寫「科目/班級」（例如：自然/三甲）或「科目 教師」（例如：數學 005林佳儀）
 */
function importMatrixTimetable(semesterCode, matrixData, mode) {
  try {
    if (!semesterCode || !Array.isArray(matrixData) || matrixData.length === 0 || matrixData.length > 5000) {
      return { success: false, message: "匯入資料為空" };
    }

    const flatRecords = [];
    const parseErrors = [];
    const weekdayMap = {
      "Mon": "Mon", "一": "Mon", "週一": "Mon", "星期一": "Mon",
      "Tue": "Tue", "二": "Tue", "週二": "Tue", "星期二": "Tue",
      "Wed": "Wed", "三": "Wed", "週三": "Wed", "星期三": "Wed",
      "Thu": "Thu", "四": "Thu", "週四": "Thu", "星期四": "Thu",
      "Fri": "Fri", "五": "Fri", "週五": "Fri", "星期五": "Fri"
    };

    const requestedMode = mode == null || mode === "" ? "auto" : String(mode).toLowerCase().trim();
    if (!["auto", "class", "teacher"].includes(requestedMode)) {
      return { success: false, message: "矩陣模式僅支援 auto、class 或 teacher" };
    }

    // 逐列檢查矩陣
    matrixData.forEach((row, index) => {
      if (!Array.isArray(row) || row.length < 3 || (row.length > 9 && row.slice(9).some(v => String(v).trim()))) { parseErrors.push(`第 ${index + 1} 列欄位數不正確`); return; }
      const col0 = String(row[0] || "").trim(); // 可能為教師名 (001 林雅惠) 或 班級 (一甲)
      const col1 = String(row[1] || "").trim(); // 星期
      const day = weekdayMap[col1] || col1;

      // 支援星期為 Mon ~ Fri
      if (!["Mon", "Tue", "Wed", "Thu", "Fri"].includes(day)) { parseErrors.push(`無法辨識星期：${col1}`); return; }

      // 解析第 1 節 ~ 第 7 節 (索引 2 ~ 8)
      for (let p = 1; p <= 7; p++) {
        const cellVal = String(row[p + 1] || "").trim();
        if (!cellVal || cellVal === "-" || cellVal === "—") continue;

        // 格子以斜線、空格或換行拆解
        const parts = cellVal.split(/[\/\s\n]+/);
        if (parts.length >= 2) {
          // 判定 A：教師矩陣模式 (col0 包含數字代碼或為老師名，格子為「科目/班級」)
          const detectedMode = requestedMode === "auto"
            ? ((/\d{3}/.test(col0) || col0.includes("師") || !col0.includes("甲")) ? "teacher" : "class")
            : requestedMode;
          if (detectedMode === "teacher") {
            const subject = parts[0];
            const targetClass = parts[1];
            flatRecords.push({
              targetClass: targetClass,
              weekday: day,
              period: p,
              subject: subject,
              teacher: col0
            });
          } 
          // 判定 B：班級矩陣模式 (col0 為「一甲~六甲」，格子為「科目/教師」)
          else {
            flatRecords.push({
              targetClass: col0,
              weekday: day,
              period: p,
              subject: parts[0],
              teacher: parts.slice(1).join(" ")
            });
          }
        } else {
          parseErrors.push(`第 ${p} 節格式無效`);
        }
      }
    });

    if (parseErrors.length > 0) {
      return { success: false, message: `矩陣資料驗證失敗：${parseErrors[0]}（共 ${parseErrors.length} 項）` };
    }

    if (flatRecords.length === 0) {
      return { success: false, message: "未能成功解析任何配課，請確認格子是否包含「科目/班級」或「科目/教師」（如：自然/三甲）" };
    }

    return importTimetableData(semesterCode, flatRecords);

  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// 寫入課表至試算表
function importTimetableData(semesterCode, scheduleArray) {
  const lock = LockService.getScriptLock();
  let sheet = null;
  let oldRows = [];
  let mutationStarted = false;
  let backupName = "";
  try {
    if (!Array.isArray(scheduleArray) || scheduleArray.length === 0) {
      return { success: false, message: "課表資料為空，未變更既有課表" };
    }
    if (scheduleArray.length > 5000) {
      return { success: false, message: "課表資料超過單次匯入上限" };
    }

    const safeSemester = normalizeText_(semesterCode, 30);
    if (!safeSemester || !getSemesterCodes_().has(safeSemester)) {
      return { success: false, message: "找不到指定學期，未變更既有課表" };
    }
    const rows = scheduleArray.map(item => [
      safeSemester,
      normalizeText_(item.targetClass, 30),
      normalizeText_(item.weekday, 12),
      normalizeText_(item.period, 4),
      normalizeText_(item.subject, 120),
      normalizeText_(item.teacher, 120)
    ]);

    const invalidRow = rows.find(row => !row.every(Boolean) ||
      !["一甲", "二甲", "三甲", "四甲", "五甲", "六甲"].includes(row[1]) ||
      !["Mon", "Tue", "Wed", "Thu", "Fri"].includes(row[2]) ||
      !/^[1-7]$/.test(row[3]));
    if (invalidRow) {
      return { success: false, message: "課表含有無效欄位；目前支援一甲至六甲、週一至週五、第 1 至 7 節，科目及教師必填。未變更既有課表。" };
    }

    if (rows.length === 0) {
      return { success: false, message: "課表資料缺少必要欄位，未變更既有課表" };
    }
    const seenSlots = new Set();
    for (let i = 0; i < rows.length; i++) {
      const slot = `${rows[i][0]}|${rows[i][1]}|${rows[i][2]}|${rows[i][3]}`;
      if (seenSlots.has(slot)) return { success: false, message: "課表含有重複班級／星期／節次，未變更既有課表" };
      seenSlots.add(slot);
    }
    if (!lock.tryLock(10000)) {
      return { success: false, message: "系統正在處理另一筆資料，請稍候再試" };
    }

    sheet = getTimetableSheet();
    const lastRow = sheet.getLastRow();

    oldRows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 6).getValues() : [];
    backupName = createTimetableBackup_(safeSemester, oldRows);
    SpreadsheetApp.flush();
    const retainedRows = oldRows.filter(row => String(row[0]).trim() !== safeSemester);
    const replacementRows = retainedRows.concat(rows);
    mutationStarted = true;
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    if (replacementRows.length > 0) {
      sheet.getRange(2, 1, replacementRows.length, 6).setValues(replacementRows.map(row => row.map(preserveSheetValue_)));
    }
    SpreadsheetApp.flush();

    return { success: true, count: rows.length, backupSheet: backupName };
  } catch (err) {
    try {
      if (mutationStarted && sheet && oldRows) {
        const currentLast = sheet.getLastRow();
        if (currentLast > 1) sheet.getRange(2, 1, currentLast - 1, 6).clearContent();
        if (oldRows.length) sheet.getRange(2, 1, oldRows.length, 6).setValues(oldRows.map(row => row.map(preserveSheetValue_)));
        SpreadsheetApp.flush();
      }
    } catch (rollbackErr) {
      Logger.log("課表回復失敗（細節不回傳）");
      return { success: false, message: `匯入失敗，未能自動回復；請從「${backupName}」回復課表。` };
    }
    return { success: false, message: mutationStarted ? "匯入失敗，已回復原課表。" : "課表檢查或備份失敗，未變更原課表。" };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function createTimetableBackup_(semester, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stamp = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss");
  const base = `課表備份_${semester}_${stamp}`.replace(/[^\w\-一-龥]/g, "_").slice(0, 90);
  let name = base;
  let suffix = 1;
  while (ss.getSheetByName(name)) name = `${base}_${suffix++}`.slice(0, 99);
  const backup = ss.insertSheet(name);
  backup.getRange(1, 1, 1, 6).setValues([["學期", "班級", "星期", "節次", "科目", "任課教師"]]);
  backup.getRange(1, 1, 1, 6).setFontWeight("bold");
  if (rows && rows.length) backup.getRange(2, 1, rows.length, 6).setValues(rows.map(row => row.map(preserveSheetValue_)));
  return name;
}

// ---------------------- 巡查紀錄儲存與後端自動寄信 ----------------------

function buildScheduleMap_(semester) {
  const sheet = getTimetableSheet();
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow <= 1) return map;
  sheet.getRange(2, 1, lastRow - 1, 6).getValues().forEach(row => {
    if (String(row[0]).trim() !== semester) return;
    const cls = String(row[1]).trim();
    const weekday = String(row[2]).trim();
    const period = String(row[3]).trim();
    if (!map[cls]) map[cls] = {};
    if (!map[cls][weekday]) map[cls][weekday] = {};
    map[cls][weekday][period] = { subject: String(row[4]).trim(), teacher: String(row[5]).trim() };
  });
  return map;
}

function expectedWeekdayForDate_(dateText) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) return "";
  const date = new Date(`${dateText}T00:00:00Z`);
  if (isNaN(date.getTime()) || Utilities.formatDate(date, "GMT+8", "yyyy-MM-dd") !== dateText) return "";
  return ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][date.getUTCDay()];
}

function validateInspectionEntry_(entry, semester, scheduleMap, options) {
  if (!entry.date || !entry.targetClass || !entry.inspectorRole) {
    return { ok: false, message: "巡查日期、班級與巡查身分不得空白" };
  }
  const expectedWeekday = expectedWeekdayForDate_(entry.date);
  if (!expectedWeekday || expectedWeekday === "Sat" || expectedWeekday === "Sun") {
    return { ok: false, message: "巡查日期必須為有效的平日 YYYY-MM-DD" };
  }
  if (entry.weekday !== expectedWeekday) {
    return { ok: false, message: "巡查日期與星期不一致" };
  }
  if (!/^[1-7]$/.test(entry.period)) {
    return { ok: false, message: "節次必須為 1 至 7" };
  }
  const expected = scheduleMap[entry.targetClass] && scheduleMap[entry.targetClass][entry.weekday] &&
    scheduleMap[entry.targetClass][entry.weekday][entry.period];
  if (!expected) {
    return { ok: false, message: "找不到指定班級在該星期／節次的課表" };
  }
  const checks = [
    ["巡查身分職稱", options.inspectorRoles], ["上課地點", options.locations],
    ["吻合狀態", options.matchStatuses], ["學習評量", options.evalExams],
    ["課堂秩序", options.evalDisciplines], ["環境衛生", options.evalEnvs],
    ["資訊科技", options.evalIts]
  ];
  for (let i = 0; i < checks.length; i++) {
    const value = [entry.inspectorRole, entry.location, entry.matchStatus, entry.evalExamItem,
      entry.evalDisciplineItem, entry.evalEnvItem, entry.evalItItem][i];
    if (!checks[i][1].includes(value)) return { ok: false, message: `${checks[i][0]}不是目前允許的選項` };
  }
  return { ok: true, expected: expected };
}

function deterministicInspectionId_(semester, entry) {
  const source = [semester, entry.date, entry.weekday, entry.period, entry.targetClass, entry.inspectorRole].join("|");
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8);
  return "INS-" + bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("").slice(0, 32);
}

function saveBatchInspectionRecords(entries, currentSemester) {
  const lock = LockService.getScriptLock();
  let writeStarted = false;
  try {
    if (!Array.isArray(entries) || entries.length === 0) {
      return { success: false, message: "無巡查資料" };
    }
    if (entries.length > 20) {
      return { success: false, message: "單次巡查資料超過系統上限" };
    }
    if (!lock.tryLock(10000)) {
      return { success: false, message: "系統正在處理另一筆巡查，請稍候再試" };
    }

    const sem = normalizeText_(currentSemester || "115-1", 30);
    const options = getSystemOptions_();
    if (options.loadError || !getSemesterCodes_().has(sem)) {
      return { success: false, message: "找不到指定學期或系統選單設定無法讀取" };
    }
    const scheduleMap = buildScheduleMap_(sem);
    const safeEntries = entries.map(normalizeInspectionEntry_);
    let validationError = "";
    safeEntries.forEach(r => {
      if (validationError) return;
      // 不信任前端產生的隨機 ID；以「學期／日期／星期／節次／班級／巡堂身分」建立固定鍵，避免重複寫入。
      r.id = deterministicInspectionId_(sem, r);
      r.createdAt = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
      const validation = validateInspectionEntry_(r, sem, scheduleMap, options);
      if (!validation.ok) { validationError = validation.message; return; }
      r.expectedSubject = validation.expected.subject;
      r.expectedTeacher = validation.expected.teacher;
    });
    if (validationError) return { success: false, message: validationError };

    const sheet = getInspectionSheet();
    const lastRow = sheet.getLastRow();
    // 以業務鍵判斷重複，不只依賴前端傳入的紀錄 ID；可攔截舊版隨機 ID 已存在的重複資料。
    const existingLogicalKeys = lastRow > 1
      ? new Set(sheet.getRange(2, 2, lastRow - 1, 6).getDisplayValues()
          .map(row => row.map(value => String(value).trim()).join("|")))
      : new Set();
    const incomingLogicalKeys = new Set();
    const uniqueEntries = safeEntries.filter(r => {
      const logicalKey = [sem, r.date, r.weekday, r.period, r.targetClass, r.inspectorRole].join("|");
      if (existingLogicalKeys.has(logicalKey) || incomingLogicalKeys.has(logicalKey)) return false;
      incomingLogicalKeys.add(logicalKey);
      return true;
    });

    if (uniqueEntries.length === 0) {
      // 若前次在寫入後、建立 Email journal 前中斷，重送同一批可補建遺失的 journal；
      // 已 sent 或 sending 的鍵仍會被保護而不重寄。
      lock.releaseLock();
      const retryIds = new Set(safeEntries.map(r => r.id));
      const persisted = getAllInspectionRecords(sem).filter(r => retryIds.has(r.id));
      const recoveryEmail = sendInspectionAlertEmails_(persisted, sem);
      return {
        success: true,
        count: 0,
        duplicate: true,
        email: recoveryEmail && recoveryEmail.triggeredCategories ? recoveryEmail : { status: "none", message: "重複送出已略過，未再次寄信。" }
      };
    }

    const rows = uniqueEntries.map(r => [
      r.id,
      sem,
      r.date,
      r.weekday,
      r.period,
      r.targetClass,
      r.inspectorRole,
      r.location,
      r.expectedSubject,
      r.expectedTeacher,
      r.actualSubject,
      r.actualTeacher,
      r.matchStatus,
      r.evalExamItem,
      r.evalDisciplineItem,
      r.evalEnvItem,
      r.evalItItem,
      r.subNotes,
      r.praiseNotes,
      r.repairNotes,
      r.improveNotes,
      r.createdAt
    ].map(sanitizeSheetCell_));

    writeStarted = true;
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
    lock.releaseLock();

    const batchIds = new Set(safeEntries.map(r => r.id));
    const persisted = getAllInspectionRecords(sem).filter(r => batchIds.has(r.id));
    const email = sendInspectionAlertEmails_(persisted, sem);
    return { success: true, count: rows.length, email: email };
  } catch (error) {
    Logger.log("儲存巡查紀錄失敗: " + error.toString());
    if (writeStarted) throw new Error("尚無法確認儲存結果，請保留這批資料並使用原批次重試。");
    return { success: false, message: "巡查紀錄儲存失敗，請稍後再試或聯絡系統管理者。" };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function sendInspectionAlertEmails_(entries, sem) {
  const result = {
    status: "none",
    sentCount: 0,
    failedCount: 0,
    alreadySentCount: 0,
    uncertainCount: 0,
    skippedCategories: 0,
    triggeredCategories: 0,
    message: "本批紀錄無需寄送異常通報。"
  };

  try {
    const opts = getSystemOptions_();
    if (opts.loadError) throw new Error("options-load-failed");
    const abnormalTeaching = entries.filter(r =>
      r.evalEnvItem === "待改善" ||
      (r.evalExamItem && (r.evalExamItem.includes("違規") || r.evalExamItem.includes("佔用"))) ||
      r.evalDisciplineItem === "待加強" ||
      r.matchStatus === "異常不符"
    );
    const repairAlerts = entries.filter(r => r.repairNotes && r.repairNotes.trim().length > 0);
    const sample = entries[0];

    if (abnormalTeaching.length > 0) {
      result.triggeredCategories++;
      const subject = `【${getSchoolName_()}巡堂通報】${sample.date} 第${sample.period}節 教學正常化/常規/環境異常列管通知`;
      let body = buildEmailHeader_(sem, "教學正常化／常規／環境待改善", sample);
      body += `【列管班級現場明細】：\n\n`;
      abnormalTeaching.forEach((r, idx) => {
        body += `[${idx + 1}] 班級：${r.targetClass}（地點：${r.location || "原班教室"}）\n`;
        body += `    • 預定科目／教師：${r.expectedSubject}（${r.expectedTeacher}）\n`;
        body += `    • 實授現場／教師：${r.actualSubject}（${r.actualTeacher}）[${r.matchStatus}]\n`;
        body += `    • 指標檢核：評量[${r.evalExamItem}]／常規[${r.evalDisciplineItem}]／環境[${r.evalEnvItem}]／媒材[${r.evalItItem}]\n`;
        if (r.improveNotes || r.subNotes) body += `    • 說明與改善備註：${r.improveNotes || r.subNotes}\n`;
        body += `\n`;
      });
      body += `========================================================\n此信件由巡堂系統於儲存異常紀錄時自動寄送。`;
      if (opts.emailTeaching.length === 0) result.skippedCategories++;
      mergeEmailSendResult_(result, sendRecipientGroup_(opts.emailTeaching, subject, body, "teaching", abnormalTeaching.map(r => r.id)));
    }

    if (repairAlerts.length > 0) {
      result.triggeredCategories++;
      const subject = `【${getSchoolName_()}設備待修繕通報】${sample.date} 第${sample.period}節 巡堂修繕列管通知`;
      let body = buildEmailHeader_(sem, "設備設施需維修待處理", sample);
      body += `【待修班級與項目明細】：\n\n`;
      repairAlerts.forEach((r, idx) => {
        body += `[${idx + 1}] 班級：${r.targetClass}（地點：${r.location || "原班教室"}）\n`;
        body += `    • 報修項目：${r.repairNotes}\n\n`;
      });
      body += `========================================================\n此信件由巡堂系統自動寄送，請權責單位協助列管處理。`;
      if (opts.emailRepair.length === 0) result.skippedCategories++;
      mergeEmailSendResult_(result, sendRecipientGroup_(opts.emailRepair, subject, body, "repair", repairAlerts.map(r => r.id)));
    }

    if (result.triggeredCategories === 0) return result;
    if (result.uncertainCount > 0) {
      result.status = "partial";
      result.message = `已寄出 ${result.sentCount} 封、先前已寄 ${result.alreadySentCount} 封；另有 ${result.uncertainCount} 封處理中或結果待確認，請查看寄送紀錄。`;
    } else if (result.sentCount + result.alreadySentCount > 0 && result.failedCount === 0 && result.skippedCategories === 0) {
      result.status = "sent";
      result.message = `異常通報已寄送 ${result.sentCount} 封；先前已寄 ${result.alreadySentCount} 封未重寄。`;
    } else if (result.sentCount > 0) {
      result.status = "partial";
      result.message = `異常通報部分完成：寄出 ${result.sentCount} 封；另有 ${result.failedCount} 封失敗、${result.skippedCategories} 類未設定收件人。`;
    } else if (result.failedCount > 0) {
      result.status = "failed";
      result.message = `異常紀錄已儲存，但 ${result.failedCount} 封 Email 寄送失敗。`;
    } else {
      result.status = "skipped";
      result.message = "異常紀錄已儲存，但尚未設定對應的有效收件人。";
    }
    return result;
  } catch (err) {
    Logger.log("發送信件通報失敗: " + err.toString());
    result.status = "failed";
    result.failedCount = Math.max(result.failedCount, 1);
    result.message = "異常紀錄已儲存，但 Email 通報處理失敗。";
    return result;
  }
}

function buildEmailHeader_(sem, alertType, sample) {
  return `【${getSchoolName_()} ${sem} 走廊巡堂通報】\n` +
    `========================================================\n` +
    `通報類型：${alertType}\n` +
    `巡查時段：${sample.date} 第 ${sample.period} 節\n` +
    `巡堂人員：${sample.inspectorRole}\n` +
    `========================================================\n\n`;
}

const EMAIL_MAX_ATTEMPTS_ = 3;
const EMAIL_RETRY_BATCH_LIMIT_ = 50;

function getEmailDeliverySheet_(create) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Email寄送紀錄");
  if (!sheet && create !== false) {
    sheet = ss.insertSheet("Email寄送紀錄");
    sheet.getRange(1, 1, 1, 11).setValues([["唯一鍵", "紀錄ID清單", "類別", "收件人", "主旨", "內容", "狀態", "嘗試次數", "認領時間", "寄出時間", "最後更新"]]);
    sheet.getRange(1, 1, 1, 11).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function emailDeliveryKey_(recordIds, category, recipient) {
  const ids = (recordIds || []).map(String).sort().join(",");
  const source = `${category}|${ids}|${String(recipient).trim().toLowerCase()}`;
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8);
  return "ED-" + bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function findEmailDeliveryRow_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const keys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  const idx = keys.indexOf(key);
  return idx < 0 ? 0 : idx + 2;
}

function claimEmailDelivery_(recipient, subject, body, category, recordIds, retryFailed) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { claimed: false, busy: true };
  try {
    const sheet = getEmailDeliverySheet_(true);
    const key = emailDeliveryKey_(recordIds, category, recipient);
    let row = findEmailDeliveryRow_(sheet, key);
    if (!row) {
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1, 1, 11).setValues([[key, sanitizeSheetCell_((recordIds || []).join(",")), sanitizeSheetCell_(category), sanitizeSheetCell_(recipient), sanitizeSheetCell_(subject), sanitizeSheetCell_(body), "pending", 0, "", "", ""]]);
    }
    const values = sheet.getRange(row, 1, 1, 11).getValues()[0];
    const status = String(values[6] || "pending");
    const attempts = Number(values[7] || 0);
    if (status === "sent" || status === "sending") return { claimed: false, status: status };
    if (status === "failed" && !retryFailed) return { claimed: false, status: status };
    if (attempts >= EMAIL_MAX_ATTEMPTS_) return { claimed: false, status: "failed", exhausted: true };
    sheet.getRange(row, 7, 1, 5).setValues([["sending", attempts + 1, new Date(), values[9] || "", new Date()]]);
    SpreadsheetApp.flush();
    return { claimed: true, row: row, key: key, attempts: attempts + 1 };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function finishEmailDelivery_(row, status) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const sheet = getEmailDeliverySheet_(true);
    if (row > 1 && row <= sheet.getLastRow()) {
      const now = new Date();
      sheet.getRange(row, 7, 1, 5).setValues([[status, sheet.getRange(row, 8).getValue(), "", status === "sent" ? now : "", now]]);
      SpreadsheetApp.flush();
    }
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function sendRecipientGroup_(recipients, subject, body, category, recordIds, retryFailed) {
  const sendResult = { sentCount: 0, failedCount: 0, alreadySentCount: 0, uncertainCount: 0 };
  if (!Array.isArray(recipients) || recipients.length === 0) return sendResult;

  recipients.forEach(recipient => {
    const claim = claimEmailDelivery_(recipient, subject, body, category || "unknown", recordIds || [], !!retryFailed);
    if (!claim.claimed) {
      if (claim.status === "sent") sendResult.alreadySentCount++;
      else if (claim.status === "failed") sendResult.failedCount++;
      else sendResult.uncertainCount++;
      return;
    }
    try {
      MailApp.sendEmail({
        to: recipient,
        subject: subject,
        body: body,
        name: getSchoolName_() + "巡堂系統"
      });
      sendResult.sentCount++;
      try { finishEmailDelivery_(claim.row, "sent"); } catch (finalizeErr) {
        // Mail 已送出但 journal 無法更新時保留 sending，避免自動重寄造成重複。
        Logger.log("Email journal 更新失敗，保留 sending 狀態");
      }
    } catch (err) {
      try { finishEmailDelivery_(claim.row, "failed"); } catch (finalizeErr) {
        Logger.log("Email journal 更新失敗，保留 sending 狀態");
      }
      sendResult.failedCount++;
      Logger.log("Email 寄送失敗（收件人與錯誤細節未記錄）");
    }
  });
  return sendResult;
}

function retryFailedEmailNotifications() {
  const sheet = getEmailDeliverySheet_(false);
  const result = { success: true, attempted: 0, sentCount: 0, failedCount: 0, skippedCount: 0 };
  if (!sheet) return result;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;
  const rows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  for (let i = 0; i < rows.length && result.attempted < EMAIL_RETRY_BATCH_LIMIT_; i++) {
    const row = rows[i];
    if (String(row[6]) !== "failed" || Number(row[7] || 0) >= EMAIL_MAX_ATTEMPTS_) continue;
    result.attempted++;
    const claim = claimEmailDelivery_(row[3], row[4], row[5], row[2], String(row[1] || "").split(",").filter(Boolean), true);
    if (!claim.claimed) { result.skippedCount++; continue; }
    try {
      MailApp.sendEmail({ to: row[3], subject: row[4], body: row[5], name: getSchoolName_() + "巡堂系統" });
      result.sentCount++;
      try { finishEmailDelivery_(claim.row, "sent"); } catch (finalizeErr) { Logger.log("Email journal 更新失敗，保留 sending 狀態"); }
    } catch (err) {
      try { finishEmailDelivery_(claim.row, "failed"); } catch (finalizeErr) { Logger.log("Email journal 更新失敗，保留 sending 狀態"); }
      result.failedCount++;
    }
  }
  return result;
}

function getEmailDeliverySummary() {
  const sheet = getEmailDeliverySheet_(false);
  const summary = { total: 0, pending: 0, sending: 0, sent: 0, failed: 0, retryableFailed: 0, exhausted: 0 };
  if (!sheet) return summary;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return summary;
  const statuses = sheet.getRange(2, 7, lastRow - 1, 1).getDisplayValues().flat();
  const attempts = sheet.getRange(2, 8, lastRow - 1, 1).getDisplayValues().flat();
  statuses.forEach((status, index) => {
    summary.total++;
    if (Object.prototype.hasOwnProperty.call(summary, status)) summary[status]++;
    if (status === "failed") {
      if (Number(attempts[index] || 0) < EMAIL_MAX_ATTEMPTS_) summary.retryableFailed++;
      else summary.exhausted++;
    }
  });
  return summary;
}

function mergeEmailSendResult_(target, source) {
  target.sentCount += source.sentCount;
  target.failedCount += source.failedCount;
  target.alreadySentCount += source.alreadySentCount || 0;
  target.uncertainCount += source.uncertainCount || 0;
}

function normalizeInspectionEntry_(record) {
  const value = record || {};
  return {
    id: normalizeText_(value.id, 120),
    date: normalizeText_(value.date, 20),
    weekday: normalizeText_(value.weekday, 12),
    period: normalizeText_(value.period, 4),
    targetClass: normalizeText_(value.targetClass, 30),
    inspectorRole: normalizeText_(value.inspectorRole, 80),
    location: normalizeText_(value.location || "原班教室", 120),
    expectedSubject: normalizeText_(value.expectedSubject, 120),
    expectedTeacher: normalizeText_(value.expectedTeacher, 120),
    actualSubject: normalizeText_(value.actualSubject, 120),
    actualTeacher: normalizeText_(value.actualTeacher, 120),
    matchStatus: normalizeText_(value.matchStatus || "吻合", 40),
    evalExamItem: normalizeText_(value.evalExamItem || "正常教學無測驗", 200),
    evalDisciplineItem: normalizeText_(value.evalDisciplineItem || "良好", 120),
    evalEnvItem: normalizeText_(value.evalEnvItem || "良好", 120),
    evalItItem: normalizeText_(value.evalItItem || "觸控大屏/電子白板", 200),
    subNotes: normalizeText_(value.subNotes, 2000),
    praiseNotes: normalizeText_(value.praiseNotes, 2000),
    repairNotes: normalizeText_(value.repairNotes, 2000),
    improveNotes: normalizeText_(value.improveNotes, 2000),
    createdAt: normalizeText_(value.createdAt, 30)
  };
}

function normalizeText_(value, maxLength) {
  return String(value == null ? "" : value).trim().slice(0, maxLength || 2000);
}

function sanitizeSheetCell_(value) {
  const text = String(value == null ? "" : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function preserveSheetValue_(value) {
  return typeof value === "string" ? sanitizeSheetCell_(value) : value;
}

function getAllInspectionRecords(filterSemester) {
  try {
    const sheet = getInspectionSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 22).getValues();
    const records = [];

    data.forEach(row => {
      const recSem = String(row[1]).trim();
      if (filterSemester && recSem !== filterSemester) return;

      let dateVal = row[2];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, "GMT+8", "yyyy-MM-dd");
      } else {
        dateVal = String(dateVal);
      }

      records.push({
        id: String(row[0]),
        semester: recSem,
        date: dateVal,
        weekday: String(row[3]),
        period: String(row[4]),
        targetClass: String(row[5]),
        inspectorRole: String(row[6]),
        location: String(row[7] || "原班教室"),
        expectedSubject: String(row[8]),
        expectedTeacher: String(row[9]),
        actualSubject: String(row[10]),
        actualTeacher: String(row[11]),
        matchStatus: String(row[12]),
        evalExamItem: String(row[13] || "正常教學無測驗"),
        evalDisciplineItem: String(row[14] || "良好"),
        evalEnvItem: String(row[15] || "良好"),
        evalItItem: String(row[16] || "觸控大屏/電子白板"),
        subNotes: String(row[17] || ""),
        praiseNotes: String(row[18] || ""),
        repairNotes: String(row[19] || ""),
        improveNotes: String(row[20] || ""),
        createdAt: String(row[21] || "")
      });
    });

    records.sort((a, b) => (b.date + b.period).localeCompare(a.date + a.period));
    return records;
  } catch (err) {
    Logger.log("讀取巡查紀錄失敗: " + err.toString());
    throw new Error("無法讀取巡查紀錄，請稍後再試或聯絡系統管理者。");
  }
}

function deleteRecordById(id) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) return { success: false, message: "系統忙碌中，請稍候再試" };
    const sheet = getInspectionSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: "無紀錄可刪除" };

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        return { success: true };
      }
    }
    return { success: false, message: "找不到指定 ID 之紀錄" };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function seedDefaultMasterSchedule(sheet) {
  const seed = [
    // 一甲
    ["115-1", "一甲", "Mon", "1", "數學", "005 林佳儀"],
    ["115-1", "一甲", "Mon", "2", "本土", "011 勞同光"],
    ["115-1", "一甲", "Mon", "3", "國語", "005 林佳儀"],
    ["115-1", "一甲", "Mon", "4", "健康", "013 林叔玉"],
    ["115-1", "一甲", "Tue", "1", "國語", "005 林佳儀"],
    ["115-1", "一甲", "Tue", "2", "體育", "004 胡定綸"],
    ["115-1", "一甲", "Tue", "3", "數學", "005 林佳儀"],
    ["115-1", "一甲", "Tue", "4", "生藝", "005 林佳儀"],
    ["115-1", "一甲", "Wed", "1", "彈全", "005 林佳儀"],
    ["115-1", "一甲", "Wed", "2", "國語", "005 林佳儀"],
    ["115-1", "一甲", "Wed", "3", "數學", "005 林佳儀"],
    ["115-1", "一甲", "Wed", "4", "生活", "010 蔡螢靜"],
    ["115-1", "一甲", "Thu", "1", "彈舞", "005 林佳儀"],
    ["115-1", "一甲", "Thu", "2", "數學", "005 林佳儀"],
    ["115-1", "一甲", "Thu", "3", "彈英", "010 蔡螢靜"],
    ["115-1", "一甲", "Thu", "4", "生活", "010 蔡螢靜"],
    ["115-1", "一甲", "Thu", "5", "國語", "005 林佳儀"],
    ["115-1", "一甲", "Thu", "6", "國語", "005 林佳儀"],
    ["115-1", "一甲", "Thu", "7", "生藝", "005 林佳儀"],
    ["115-1", "一甲", "Fri", "1", "生藝", "005 林佳儀"],
    ["115-1", "一甲", "Fri", "2", "體育", "004 胡定綸"],
    ["115-1", "一甲", "Fri", "3", "數學", "005 林佳儀"],
    ["115-1", "一甲", "Fri", "4", "國語", "005 林佳儀"],
    // 二甲
    ["115-1", "二甲", "Mon", "1", "國語", "006 蘇英桃"],
    ["115-1", "二甲", "Mon", "2", "彈英", "010 蔡螢靜"],
    ["115-1", "二甲", "Mon", "3", "本土", "011 勞同光"],
    ["115-1", "二甲", "Mon", "4", "數學", "006 蘇英桃"],
    ["115-1", "二甲", "Tue", "1", "國語", "006 蘇英桃"],
    ["115-1", "二甲", "Tue", "2", "體育", "002 陳遠聖"],
    ["115-1", "二甲", "Tue", "3", "生藝", "006 蘇英桃"],
    ["115-1", "二甲", "Tue", "4", "生藝", "006 蘇英桃"],
    ["115-1", "二甲", "Wed", "1", "彈全", "006 蘇英桃"],
    ["115-1", "二甲", "Wed", "2", "數學", "006 蘇英桃"],
    ["115-1", "二甲", "Wed", "3", "國語", "006 蘇英桃"],
    ["115-1", "二甲", "Wed", "4", "健康", "002 陳遠聖"],
    ["115-1", "二甲", "Thu", "1", "彈舞", "006 蘇英桃"],
    ["115-1", "二甲", "Thu", "2", "體育", "002 陳遠聖"],
    ["115-1", "二甲", "Thu", "3", "生活", "013 林叔玉"],
    ["115-1", "二甲", "Thu", "4", "數學", "006 蘇英桃"],
    ["115-1", "二甲", "Thu", "5", "國語", "006 蘇英桃"],
    ["115-1", "二甲", "Thu", "6", "國語", "006 蘇英桃"],
    ["115-1", "二甲", "Thu", "7", "生藝", "006 蘇英桃"],
    ["115-1", "二甲", "Fri", "1", "生藝", "006 蘇英桃"],
    ["115-1", "二甲", "Fri", "2", "生活", "013 林叔玉"],
    ["115-1", "二甲", "Fri", "3", "國語", "006 蘇英桃"],
    ["115-1", "二甲", "Fri", "4", "數學", "006 蘇英桃"],
    // 三甲
    ["115-1", "三甲", "Mon", "1", "社會", "010 蔡螢靜"],
    ["115-1", "三甲", "Mon", "2", "數學", "007 林佑仲"],
    ["115-1", "三甲", "Mon", "3", "國語", "007 林佑仲"],
    ["115-1", "三甲", "Mon", "4", "健康", "003 謝佳宜"],
    ["115-1", "三甲", "Mon", "5", "本土", "011 勞同光"],
    ["115-1", "三甲", "Mon", "6", "自然", "001 林雅惠"],
    ["115-1", "三甲", "Mon", "7", "綜合", "003 謝佳宜"],
    ["115-1", "三甲", "Tue", "1", "數學", "007 林佑仲"],
    ["115-1", "三甲", "Tue", "2", "藝美", "013 林叔玉"],
    ["115-1", "三甲", "Tue", "3", "藝音", "007 林佑仲"],
    ["115-1", "三甲", "Tue", "4", "彈音", "007 林佑仲"],
    ["115-1", "三甲", "Tue", "5", "英語", "004 胡定綸"],
    ["115-1", "三甲", "Tue", "6", "藝美", "013 林叔玉"],
    ["115-1", "三甲", "Tue", "7", "體育", "007 林佑仲"],
    ["115-1", "三甲", "Wed", "1", "彈全", "007 林佑仲"],
    ["115-1", "三甲", "Wed", "2", "國語", "007 林佑仲"],
    ["115-1", "三甲", "Wed", "3", "自然", "001 林雅惠"],
    ["115-1", "三甲", "Wed", "4", "自然", "001 林雅惠"],
    ["115-1", "三甲", "Thu", "1", "彈英", "004 胡定綸"],
    ["115-1", "三甲", "Thu", "2", "數學", "007 林佑仲"],
    ["115-1", "三甲", "Thu", "3", "國語", "007 林佑仲"],
    ["115-1", "三甲", "Thu", "4", "國語", "007 林佑仲"],
    ["115-1", "三甲", "Thu", "5", "社會", "010 蔡螢靜"],
    ["115-1", "三甲", "Thu", "6", "社會", "010 蔡螢靜"],
    ["115-1", "三甲", "Thu", "7", "體育", "007 林佑仲"],
    ["115-1", "三甲", "Fri", "1", "數學", "007 林佑仲"],
    ["115-1", "三甲", "Fri", "2", "國語", "007 林佑仲"],
    ["115-1", "三甲", "Fri", "3", "綜合", "003 謝佳宜"],
    ["115-1", "三甲", "Fri", "4", "彈資", "003 謝佳宜"],
    // 四甲
    ["115-1", "四甲", "Mon", "1", "數學", "008 李淑如"],
    ["115-1", "四甲", "Mon", "2", "國語", "008 李淑如"],
    ["115-1", "四甲", "Mon", "3", "體育", "010 蔡螢靜"],
    ["115-1", "四甲", "Mon", "4", "本土", "011 勞同光"],
    ["115-1", "四甲", "Mon", "5", "自然", "013 林叔玉"],
    ["115-1", "四甲", "Mon", "6", "自然", "013 林叔玉"],
    ["115-1", "四甲", "Mon", "7", "健康", "008 李淑如"],
    ["115-1", "四甲", "Tue", "1", "數學", "008 李淑如"],
    ["115-1", "四甲", "Tue", "2", "國語", "008 李淑如"],
    ["115-1", "四甲", "Tue", "3", "藝音", "001 林雅惠"],
    ["115-1", "四甲", "Tue", "4", "彈音", "001 林雅惠"],
    ["115-1", "四甲", "Tue", "5", "藝美", "008 李淑如"],
    ["115-1", "四甲", "Tue", "6", "藝美", "008 李淑如"],
    ["115-1", "四甲", "Tue", "7", "綜合", "008 李淑如"],
    ["115-1", "四甲", "Wed", "1", "彈全", "008 李淑如"],
    ["115-1", "四甲", "Wed", "2", "社會", "002 陳遠聖"],
    ["115-1", "四甲", "Wed", "3", "數學", "008 李淑如"],
    ["115-1", "四甲", "Wed", "4", "國語", "008 李淑如"],
    ["115-1", "四甲", "Thu", "1", "國語", "008 李淑如"],
    ["115-1", "四甲", "Thu", "2", "彈英", "004 胡定綸"],
    ["115-1", "四甲", "Thu", "3", "數學", "008 李淑如"],
    ["115-1", "四甲", "Thu", "4", "自然", "013 林叔玉"],
    ["115-1", "四甲", "Thu", "5", "社會", "002 陳遠聖"],
    ["115-1", "四甲", "Thu", "6", "社會", "002 陳遠聖"],
    ["115-1", "四甲", "Thu", "7", "體育", "010 蔡螢靜"],
    ["115-1", "四甲", "Fri", "1", "英語", "004 胡定綸"],
    ["115-1", "四甲", "Fri", "2", "國語", "008 李淑如"],
    ["115-1", "四甲", "Fri", "3", "綜合", "008 李淑如"],
    ["115-1", "四甲", "Fri", "4", "彈資", "008 李淑如"],
    // 五甲
    ["115-1", "五甲", "Mon", "1", "數學", "014 陳峻漢"],
    ["115-1", "五甲", "Mon", "2", "國語", "014 陳峻漢"],
    ["115-1", "五甲", "Mon", "3", "綜合", "014 陳峻漢"],
    ["115-1", "五甲", "Mon", "4", "英語", "004 胡定綸"],
    ["115-1", "五甲", "Mon", "5", "社會", "003 謝佳宜"],
    ["115-1", "五甲", "Mon", "6", "社會", "003 謝佳宜"],
    ["115-1", "五甲", "Mon", "7", "本土", "011 勞同光"],
    ["115-1", "五甲", "Tue", "1", "國語", "014 陳峻漢"],
    ["115-1", "五甲", "Tue", "2", "國語", "014 陳峻漢"],
    ["115-1", "五甲", "Tue", "3", "藝音", "003 謝佳宜"],
    ["115-1", "五甲", "Tue", "4", "彈音", "003 謝佳宜"],
    ["115-1", "五甲", "Tue", "5", "自然", "013 林叔玉"],
    ["115-1", "五甲", "Tue", "6", "數學", "014 陳峻漢"],
    ["115-1", "五甲", "Tue", "7", "體育", "014 陳峻漢"],
    ["115-1", "五甲", "Wed", "1", "彈全", "014 陳峻漢"],
    ["115-1", "五甲", "Wed", "2", "社會", "003 謝佳宜"],
    ["115-1", "五甲", "Wed", "3", "國語", "014 陳峻漢"],
    ["115-1", "五甲", "Wed", "4", "英語", "004 胡定綸"],
    ["115-1", "五甲", "Thu", "1", "數學", "014 陳峻漢"],
    ["115-1", "五甲", "Thu", "2", "彈舞", "014 陳峻漢"],
    ["115-1", "五甲", "Thu", "3", "彈英", "004 胡定綸"],
    ["115-1", "五甲", "Thu", "4", "彈英", "004 胡定綸"],
    ["115-1", "五甲", "Thu", "5", "自然", "013 林叔玉"],
    ["115-1", "五甲", "Thu", "6", "自然", "013 林叔玉"],
    ["115-1", "五甲", "Thu", "7", "綜合", "014 陳峻漢"],
    ["115-1", "五甲", "Fri", "1", "彈資", "003 謝佳宜"],
    ["115-1", "五甲", "Fri", "2", "健康", "014 陳峻漢"],
    ["115-1", "五甲", "Fri", "3", "國語", "014 陳峻漢"],
    ["115-1", "五甲", "Fri", "4", "數學", "014 陳峻漢"],
    ["115-1", "五甲", "Fri", "5", "藝美", "013 林叔玉"],
    ["115-1", "五甲", "Fri", "6", "藝美", "013 林叔玉"],
    ["115-1", "五甲", "Fri", "7", "體育", "014 陳峻漢"],
    // 六甲
    ["115-1", "六甲", "Mon", "1", "體育", "009 陳芳妤"],
    ["115-1", "六甲", "Mon", "2", "自然", "013 林叔玉"],
    ["115-1", "六甲", "Mon", "3", "自然", "013 林叔玉"],
    ["115-1", "六甲", "Mon", "4", "國語", "009 陳芳妤"],
    ["115-1", "六甲", "Mon", "5", "數學", "009 陳芳妤"],
    ["115-1", "六甲", "Mon", "6", "本土", "011 勞同光"],
    ["115-1", "六甲", "Mon", "7", "藝美", "013 林叔玉"],
    ["115-1", "六甲", "Tue", "1", "國語", "009 陳芳妤"],
    ["115-1", "六甲", "Tue", "2", "社會", "003 謝佳宜"],
    ["115-1", "六甲", "Tue", "3", "藝音", "009 陳芳妤"],
    ["115-1", "六甲", "Tue", "4", "彈音", "009 陳芳妤"],
    ["115-1", "六甲", "Tue", "5", "數學", "009 陳芳妤"],
    ["115-1", "六甲", "Tue", "6", "英語", "004 胡定綸"],
    ["115-1", "六甲", "Tue", "7", "綜合", "013 林叔玉"],
    ["115-1", "六甲", "Wed", "1", "彈全", "009 陳芳妤"],
    ["115-1", "六甲", "Wed", "2", "國語", "009 陳芳妤"],
    ["115-1", "六甲", "Wed", "3", "自然", "013 林叔玉"],
    ["115-1", "六甲", "Wed", "4", "社會", "003 謝佳宜"],
    ["115-1", "六甲", "Thu", "1", "數學", "009 陳芳妤"],
    ["115-1", "六甲", "Thu", "2", "彈舞", "009 陳芳妤"],
    ["115-1", "六甲", "Thu", "3", "國語", "009 陳芳妤"],
    ["115-1", "六甲", "Thu", "4", "社會", "003 謝佳宜"],
    ["115-1", "六甲", "Thu", "5", "彈英", "004 胡定綸"],
    ["115-1", "六甲", "Thu", "6", "彈英", "004 胡定綸"],
    ["115-1", "六甲", "Thu", "7", "綜合", "013 林叔玉"],
    ["115-1", "六甲", "Fri", "1", "健康", "009 陳芳妤"],
    ["115-1", "六甲", "Fri", "2", "彈資", "003 謝佳宜"],
    ["115-1", "六甲", "Fri", "3", "數學", "009 陳芳妤"],
    ["115-1", "六甲", "Fri", "4", "英語", "004 胡定綸"],
    ["115-1", "六甲", "Fri", "5", "國語", "009 陳芳妤"],
    ["115-1", "六甲", "Fri", "6", "體育", "009 陳芳妤"],
    ["115-1", "六甲", "Fri", "7", "藝美", "013 林叔玉"]
  ];
  sheet.getRange(2, 1, seed.length, seed[0].length).setValues(seed);
}
