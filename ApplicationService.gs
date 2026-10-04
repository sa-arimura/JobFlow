/**
 * 新規登録用サイドバーを表示します。
 * サイドバーはシートを見ながら入力できるため、ダイアログより登録作業に向いています。
 */
function showApplicationRegistrationSidebar() {
  showApplicationSidebar('create', '');
}

/**
 * 応募一覧で選択している行の応募IDを取得し、編集用サイドバーを表示します。
 * 行番号は画面を開くときだけ利用し、保存時は応募IDで対象を再検索します。
 */
function showApplicationEditSidebar() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadsheet.getActiveSheet();

    if (activeSheet.getName() !== SHEET_NAMES.APPLICATIONS) {
      throw new Error(
        '「応募一覧」シートで編集したい応募の行を選択してください。'
      );
    }

    const selectedRow = activeSheet.getActiveRange().getRow();
    if (selectedRow < DATA_START_ROW) {
      throw new Error('見出し行ではなく、編集したい応募の行を選択してください。');
    }

    const applicationId = String(
      activeSheet.getRange(selectedRow, APPLICATION_ID_COLUMN).getDisplayValue()
    ).trim();
    if (applicationId === '') {
      throw new Error(
        '選択した行には応募IDがないため編集できません。' +
        '応募IDのない既存行は今回の編集機能の対象外です。'
      );
    }

    findUniqueApplicationRowById(activeSheet, applicationId);
    showApplicationSidebar('edit', applicationId);
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      APPLICATION_FORM_TITLE,
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showApplicationSidebar(mode, applicationId) {
  const template = HtmlService.createTemplateFromFile('ApplicationForm');
  template.mode = mode;
  template.applicationId = applicationId;

  SpreadsheetApp.getUi().showSidebar(
    template
      .evaluate()
      .setTitle(APPLICATION_FORM_TITLE)
      .setWidth(APPLICATION_SIDEBAR_WIDTH)
  );
}

/**
 * サイドバーの初期表示に必要な情報を返します。
 * 編集時も行番号ではなく応募IDから最新の行を検索します。
 */
function getApplicationFormContext(mode, applicationId) {
  const statuses = getValidSelectionStatuses();
  const context = {
    mode: mode,
    statuses: statuses,
    submissionStatuses: SUBMISSION_STATUSES,
    submissionDefaults: SUBMISSION_DEFAULTS,
    waitingStatuses: WAITING_STATUSES,
    waitingDefaultStatus: WAITING_DEFAULT_STATUS,
    application: null
  };

  if (mode === 'edit') {
    const sheet = getApplicationsSheetForDataOperation();
    const row = findUniqueApplicationRowById(sheet, applicationId);
    context.application = readApplicationFromRow(sheet, row, applicationId);
  }

  return context;
}

/**
 * 新規登録ではDocumentLockを利用し、同時実行による同じ行への書込みを防ぎます。
 */
function registerApplication(formData) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(LOCK_WAIT_MILLISECONDS);

  try {
    const validatedData = validateApplicationInput(formData, false);
    const sheet = getApplicationsSheetForDataOperation();
    const applicationId = Utilities.getUuid();
    const targetRow = findSafeAppendRow(sheet);
    const savedAt = new Date();
    const waitingState = resolveWaitingState(
      '',
      WAITING_STATUS_VALUES.NONE,
      '',
      validatedData.selectionStatus,
      validatedData.waitingStatus,
      savedAt
    );

    ensureSheetSize(sheet, targetRow, APPLICATION_COLUMN_COUNT);
    ensureApplicationStatusValidation(sheet, targetRow, getSettingsSheetForDataOperation());
    sheet
      .getRange(
        targetRow,
        APPLICATION_ID_COLUMN,
        1,
        APPLICATION_WRITE_COLUMN_COUNT
      )
      .setValues([[
        applicationId,
        validatedData.companyName,
        validatedData.jobTitle,
        validatedData.jobUrl,
        validatedData.applicationDate,
        validatedData.applicationSite,
        validatedData.selectionStatus
      ]]);
    writeApplicationActivityData(sheet, targetRow, validatedData);
    sheet
      .getRange(targetRow, APPLICATION_COLUMNS.CONTACT_PERSON)
      .setValue(validatedData.contactPerson);
    sheet
      .getRange(targetRow, APPLICATION_COLUMNS.CREATED_DATE_TIME, 1, 2)
      .setValues([[savedAt, savedAt]]);
    writeSubmissionData(sheet, targetRow, validatedData);
    writeWaitingData(sheet, targetRow, waitingState);

    return {
      success: true,
      applicationId: applicationId,
      message: '応募情報を登録しました。'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 編集では応募IDを変更せず、IDが一意に見つかった場合だけ入力項目を更新します。
 */
function updateApplication(formData) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(LOCK_WAIT_MILLISECONDS);

  try {
    const validatedData = validateApplicationInput(formData, true);
    const sheet = getApplicationsSheetForDataOperation();
    const targetRow = findUniqueApplicationRowById(
      sheet,
      validatedData.applicationId
    );
    const existingSelectionStatus = String(
      sheet.getRange(targetRow, APPLICATION_COLUMNS.SELECTION_STATUS)
        .getDisplayValue()
    ).trim();
    const existingWaitingStatus = String(
      sheet.getRange(targetRow, APPLICATION_COLUMNS.WAITING_STATUS)
        .getDisplayValue()
    ).trim();
    const existingWaitingStartDate = sheet
      .getRange(targetRow, APPLICATION_COLUMNS.WAITING_START_DATE)
      .getValue();
    const savedAt = new Date();
    const waitingState = resolveWaitingState(
      existingSelectionStatus,
      existingWaitingStatus,
      existingWaitingStartDate,
      validatedData.selectionStatus,
      validatedData.waitingStatus,
      savedAt
    );

    sheet
      .getRange(
        targetRow,
        APPLICATION_COLUMNS.COMPANY_NAME,
        1,
        APPLICATION_WRITE_COLUMN_COUNT - 1
      )
      .setValues([[
        validatedData.companyName,
        validatedData.jobTitle,
        validatedData.jobUrl,
        validatedData.applicationDate,
        validatedData.applicationSite,
        validatedData.selectionStatus
      ]]);
    writeApplicationActivityData(sheet, targetRow, validatedData);
    sheet
      .getRange(targetRow, APPLICATION_COLUMNS.CONTACT_PERSON)
      .setValue(validatedData.contactPerson);
    sheet
      .getRange(targetRow, APPLICATION_COLUMNS.UPDATED_DATE_TIME)
      .setValue(savedAt);
    writeSubmissionData(sheet, targetRow, validatedData);
    writeWaitingData(sheet, targetRow, waitingState);

    return {
      success: true,
      applicationId: validatedData.applicationId,
      message: '応募情報を編集しました。'
    };
  } finally {
    lock.releaseLock();
  }
}

function validateApplicationInput(formData, requiresApplicationId) {
  if (!formData || typeof formData !== 'object') {
    throw new Error('入力内容を取得できませんでした。');
  }

  const applicationId = normalizeApplicationText(formData.applicationId);
  const companyName = normalizeApplicationText(formData.companyName);
  const jobTitle = normalizeApplicationText(formData.jobTitle);
  const jobUrl = normalizeApplicationText(formData.jobUrl);
  const applicationDateText = normalizeApplicationText(formData.applicationDate);
  const applicationSite = normalizeApplicationText(formData.applicationSite);
  const contactPerson = normalizeApplicationText(formData.contactPerson);
  const selectionStatus = normalizeApplicationText(formData.selectionStatus);
  const lastReplyDateText = normalizeApplicationText(formData.lastReplyDate);
  const nextScheduleDateText = normalizeApplicationText(formData.nextScheduleDate);
  const interviewDateTimeText = normalizeApplicationText(
    formData.interviewDateTime
  );
  const memo = normalizeApplicationText(formData.memo);
  const resumeStatus = normalizeApplicationText(formData.resumeStatus);
  const workHistoryStatus = normalizeApplicationText(formData.workHistoryStatus);
  const portfolioStatus = normalizeApplicationText(formData.portfolioStatus);
  const otherSubmissionStatus = normalizeApplicationText(formData.otherSubmissionStatus);
  const submissionDeadlineText = normalizeApplicationText(formData.submissionDeadline);
  const submissionDateText = normalizeApplicationText(formData.submissionDate);
  const submissionMemo = normalizeApplicationText(formData.submissionMemo);
  const waitingStatus = normalizeApplicationText(formData.waitingStatus);

  if (requiresApplicationId && applicationId === '') {
    throw new Error('応募IDがないため編集できません。');
  }
  if (companyName === '') {
    throw new Error('会社名を入力してください。');
  }
  if (jobTitle === '') {
    throw new Error('職種を入力してください。');
  }
  if (jobUrl !== '' && !isClearlyValidHttpUrl(jobUrl)) {
    throw new Error(
      '求人URLは「https://」または「http://」から始まる有効なURLを入力してください。'
    );
  }

  const applicationDate = parseOptionalApplicationDate(applicationDateText);
  const validStatuses = getValidSelectionStatuses();
  if (validStatuses.indexOf(selectionStatus) === -1) {
    throw new Error('選考状況は設定シートの一覧から選択してください。');
  }
  [
    ['履歴書', resumeStatus],
    ['職務経歴書', workHistoryStatus],
    ['ポートフォリオ', portfolioStatus],
    ['その他提出物', otherSubmissionStatus]
  ].forEach(function(item) {
    if (requiresApplicationId && item[1] === '') {
      return;
    }
    if (SUBMISSION_STATUSES.indexOf(item[1]) === -1) {
      throw new Error(item[0] + 'は不要・未提出・提出済みから選択してください。');
    }
  });
  if (WAITING_STATUSES.indexOf(waitingStatus) === -1) {
    throw new Error('対応待ちは、なし・企業待ち・自分対応から選択してください。');
  }

  return {
    applicationId: applicationId,
    companyName: companyName,
    jobTitle: jobTitle,
    jobUrl: jobUrl,
    applicationDate: applicationDate,
    applicationSite: applicationSite,
    contactPerson: contactPerson,
    selectionStatus: selectionStatus,
    lastReplyDate: parseOptionalJobFlowDate(
      lastReplyDateText,
      '最終返信日'
    ),
    nextScheduleDate: parseOptionalJobFlowDate(
      nextScheduleDateText,
      '次回予定日'
    ),
    interviewDateTime: parseOptionalJobFlowDateTime(
      interviewDateTimeText,
      '面接日時'
    ),
    memo: memo,
    resumeStatus: resumeStatus,
    workHistoryStatus: workHistoryStatus,
    portfolioStatus: portfolioStatus,
    otherSubmissionStatus: otherSubmissionStatus,
    submissionDeadline: parseOptionalJobFlowDate(
      submissionDeadlineText,
      '提出期限'
    ),
    submissionDate: parseOptionalJobFlowDate(submissionDateText, '提出日'),
    submissionMemo: submissionMemo,
    waitingStatus: waitingStatus
  };
}

function writeApplicationActivityData(sheet, row, validatedData) {
  sheet
    .getRange(
      row,
      APPLICATION_COLUMNS.LAST_REPLY_DATE,
      1,
      APPLICATION_COLUMNS.MEMO -
        APPLICATION_COLUMNS.LAST_REPLY_DATE + 1
    )
    .setValues([[
      validatedData.lastReplyDate,
      validatedData.nextScheduleDate,
      validatedData.interviewDateTime,
      validatedData.memo
    ]]);
}

function writeSubmissionData(sheet, row, validatedData) {
  sheet
    .getRange(
      row,
      APPLICATION_COLUMNS.RESUME_STATUS,
      1,
      APPLICATION_COLUMNS.SUBMISSION_MEMO -
        APPLICATION_COLUMNS.RESUME_STATUS + 1
    )
    .setValues([[
      validatedData.resumeStatus,
      validatedData.workHistoryStatus,
      validatedData.portfolioStatus,
      validatedData.otherSubmissionStatus,
      validatedData.submissionDeadline,
      validatedData.submissionDate,
      validatedData.submissionMemo
    ]]);
}

function writeWaitingData(sheet, row, waitingState) {
  sheet
    .getRange(row, APPLICATION_COLUMNS.WAITING_STATUS, 1, 2)
    .setValues([[
      waitingState.waitingStatus,
      waitingState.waitingStartDate
    ]]);
}

function resolveWaitingState(
  existingSelectionStatus,
  existingWaitingStatus,
  existingWaitingStartDate,
  newSelectionStatus,
  newWaitingStatus,
  savedAt
) {
  if (isTerminalSelectionStatus(newSelectionStatus) ||
      (isTerminalSelectionStatus(existingSelectionStatus) &&
        !isTerminalSelectionStatus(newSelectionStatus))) {
    return {
      waitingStatus: WAITING_STATUS_VALUES.NONE,
      waitingStartDate: ''
    };
  }
  if (newWaitingStatus !== WAITING_STATUS_VALUES.COMPANY) {
    return {
      waitingStatus: newWaitingStatus,
      waitingStartDate: ''
    };
  }
  if (existingWaitingStatus === WAITING_STATUS_VALUES.COMPANY) {
    return {
      waitingStatus: newWaitingStatus,
      waitingStartDate: existingWaitingStartDate
    };
  }
  return {
    waitingStatus: newWaitingStatus,
    waitingStartDate: getJobFlowDateOnly(savedAt)
  };
}

function isTerminalSelectionStatus(selectionStatus) {
  return TERMINAL_SELECTION_STATUSES.indexOf(selectionStatus) !== -1;
}

function getJobFlowDateOnly(dateTime) {
  const timeZone = getJobFlowTimeZone();
  const dateText = Utilities.formatDate(dateTime, timeZone, 'yyyy-MM-dd');
  return parseOptionalJobFlowDate(dateText, '待機開始日');
}

function normalizeApplicationText(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function parseOptionalApplicationDate(dateText) {
  return parseOptionalJobFlowDate(dateText, '応募日');
}

function parseOptionalJobFlowDate(dateText, fieldName) {
  if (dateText === '') {
    return '';
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) {
    throw new Error(fieldName + 'は有効な日付を入力してください。');
  }

  const timeZone = getJobFlowTimeZone();
  let parsedDate;
  try {
    parsedDate = Utilities.parseDate(dateText, timeZone, 'yyyy-MM-dd');
  } catch (error) {
    throw new Error(fieldName + 'は有効な日付を入力してください。');
  }

  if (Utilities.formatDate(parsedDate, timeZone, 'yyyy-MM-dd') !== dateText) {
    throw new Error(fieldName + 'は有効な日付を入力してください。');
  }
  return parsedDate;
}

function parseOptionalJobFlowDateTime(dateTimeText, fieldName) {
  if (dateTimeText === '') {
    return '';
  }
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(dateTimeText)) {
    throw new Error(fieldName + 'は有効な日時を入力してください。');
  }

  const timeZone = getJobFlowTimeZone();
  const dateTimeFormat = "yyyy-MM-dd'T'HH:mm";
  let parsedDateTime;
  try {
    parsedDateTime = Utilities.parseDate(
      dateTimeText,
      timeZone,
      dateTimeFormat
    );
  } catch (error) {
    throw new Error(fieldName + 'は有効な日時を入力してください。');
  }

  if (
    Utilities.formatDate(parsedDateTime, timeZone, dateTimeFormat) !==
    dateTimeText
  ) {
    throw new Error(fieldName + 'は有効な日時を入力してください。');
  }
  return parsedDateTime;
}

function isClearlyValidHttpUrl(url) {
  return /^https?:\/\/(?:[a-z0-9-]+\.)+[a-z]{2,}(?::\d+)?(?:[/?#][^\s]*)?$/i
    .test(url);
}

function getJobFlowTimeZone() {
  const settingsSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.SETTINGS);
  const configuredTimeZone = settingsSheet
    ? String(settingsSheet.getRange('B3').getDisplayValue()).trim()
    : '';

  return configuredTimeZone || Session.getScriptTimeZone();
}

function getValidSelectionStatuses() {
  const settingsSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    throw new Error('設定シートが見つかりません。初期セットアップを実行してください。');
  }

  const statuses = settingsSheet
    .getRange(
      SETTINGS_STATUS_START_ROW,
      SETTINGS_STATUS_COLUMN,
      SELECTION_STATUSES.length,
      1
    )
    .getDisplayValues()
    .map(function(row) {
      return String(row[0]).trim();
    })
    .filter(function(status) {
      return status !== '';
    });

  if (statuses.length === 0) {
    throw new Error('設定シートに選考状況が登録されていません。');
  }
  return statuses;
}

function getApplicationsSheetForDataOperation() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.APPLICATIONS);
  if (!sheet) {
    throw new Error('応募一覧シートが見つかりません。初期セットアップを実行してください。');
  }

  validateApplicationSheetStructure(sheet);
  return sheet;
}

function getSettingsSheetForDataOperation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error('設定シートが見つかりません。初期セットアップを実行してください。');
  }
  return sheet;
}

function validateApplicationSheetStructure(sheet) {
  if (sheet.getMaxColumns() < APPLICATION_COLUMN_COUNT) {
    throw new Error('第8段階の応募一覧列がありません。初期セットアップを実行してください。');
  }

  const headers = sheet
    .getRange(
      APPLICATION_HEADER_ROW,
      1,
      1,
      APPLICATION_COLUMN_COUNT
    )
    .getDisplayValues()[0];
  const invalidHeaders = APPLICATION_HEADERS.filter(function(expected, index) {
    return String(headers[index]).trim() !== expected;
  });

  if (invalidHeaders.length > 0) {
    throw new Error(
      '応募一覧の列構成が想定と異なります。初期セットアップと見出しを確認してください。'
    );
  }
}

/**
 * ARRAYFORMULAの計算列（L～N、R）は追記位置の判定から除外します。
 * それ以外の応募データ列全体を確認し、途中の空欄で既存行を上書きしないようにします。
 */
function findSafeAppendRow(sheet) {
  const lastSheetRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
  const rowCount = lastSheetRow - DATA_START_ROW + 1;
  const mainDataValues = sheet
    .getRange(DATA_START_ROW, 1, rowCount, APPLICATION_COLUMNS.MEMO)
    .getDisplayValues();
  const supplementalDataValues = sheet
    .getRange(
      DATA_START_ROW,
      APPLICATION_COLUMNS.CREATED_DATE_TIME,
      rowCount,
      APPLICATION_COLUMNS.CONTACT_PERSON -
        APPLICATION_COLUMNS.CREATED_DATE_TIME + 1
    )
    .getDisplayValues();
  const extendedDataValues = sheet
    .getRange(
      DATA_START_ROW,
      APPLICATION_COLUMNS.RESUME_STATUS,
      rowCount,
      APPLICATION_COLUMNS.WAITING_START_DATE -
        APPLICATION_COLUMNS.RESUME_STATUS + 1
    )
    .getDisplayValues();

  for (let index = rowCount - 1; index >= 0; index -= 1) {
    const rowValues = mainDataValues[index]
      .concat(supplementalDataValues[index])
      .concat(extendedDataValues[index]);
    const hasApplicationData = rowValues.some(function(value) {
      return String(value).trim() !== '';
    });
    if (hasApplicationData) {
      return DATA_START_ROW + index + 1;
    }
  }

  return DATA_START_ROW;
}

function findUniqueApplicationRowById(sheet, applicationId) {
  const normalizedId = normalizeApplicationText(applicationId);
  if (normalizedId === '') {
    throw new Error('応募IDがないため編集できません。');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    throw new Error('指定された応募IDは存在しません。');
  }

  const idValues = sheet
    .getRange(
      DATA_START_ROW,
      APPLICATION_ID_COLUMN,
      lastRow - DATA_START_ROW + 1,
      1
    )
    .getDisplayValues();
  const matchedRows = [];

  idValues.forEach(function(row, index) {
    if (String(row[0]).trim() === normalizedId) {
      matchedRows.push(DATA_START_ROW + index);
    }
  });

  if (matchedRows.length === 0) {
    throw new Error('指定された応募IDは存在しません。');
  }
  if (matchedRows.length > 1) {
    throw new Error(
      '応募ID「' + normalizedId +
      '」が重複しているため編集できません。データは変更されていません。'
    );
  }
  return matchedRows[0];
}

function readApplicationFromRow(sheet, row, applicationId) {
  const values = sheet
    .getRange(row, 1, 1, APPLICATION_COLUMN_COUNT)
    .getValues()[0];
  const timeZone = getJobFlowTimeZone();
  const applicationDate = values[APPLICATION_COLUMNS.APPLICATION_DATE - 1];
  const lastReplyDate = values[APPLICATION_COLUMNS.LAST_REPLY_DATE - 1];
  const nextScheduleDate = values[APPLICATION_COLUMNS.NEXT_SCHEDULE_DATE - 1];
  const interviewDateTime =
    values[APPLICATION_COLUMNS.INTERVIEW_DATE_TIME - 1];
  const submissionDeadline = values[APPLICATION_COLUMNS.SUBMISSION_DEADLINE - 1];
  const submissionDate = values[APPLICATION_COLUMNS.SUBMISSION_DATE - 1];

  return {
    applicationId: applicationId,
    companyName: values[APPLICATION_COLUMNS.COMPANY_NAME - 1] || '',
    jobTitle: values[APPLICATION_COLUMNS.JOB_TITLE - 1] || '',
    jobUrl: values[APPLICATION_COLUMNS.JOB_URL - 1] || '',
    applicationDate: applicationDate instanceof Date
      ? Utilities.formatDate(applicationDate, timeZone, 'yyyy-MM-dd')
      : '',
    applicationSite: values[APPLICATION_COLUMNS.APPLICATION_SITE - 1] || '',
    contactPerson: values[APPLICATION_COLUMNS.CONTACT_PERSON - 1] || '',
    selectionStatus: values[APPLICATION_COLUMNS.SELECTION_STATUS - 1] || '',
    lastReplyDate: lastReplyDate instanceof Date
      ? Utilities.formatDate(lastReplyDate, timeZone, 'yyyy-MM-dd')
      : '',
    nextScheduleDate: nextScheduleDate instanceof Date
      ? Utilities.formatDate(nextScheduleDate, timeZone, 'yyyy-MM-dd')
      : '',
    interviewDateTime: interviewDateTime instanceof Date
      ? Utilities.formatDate(
          interviewDateTime,
          timeZone,
          "yyyy-MM-dd'T'HH:mm"
        )
      : '',
    memo: values[APPLICATION_COLUMNS.MEMO - 1] || '',
    resumeStatus: values[APPLICATION_COLUMNS.RESUME_STATUS - 1] || '',
    workHistoryStatus: values[APPLICATION_COLUMNS.WORK_HISTORY_STATUS - 1] || '',
    portfolioStatus: values[APPLICATION_COLUMNS.PORTFOLIO_STATUS - 1] || '',
    otherSubmissionStatus:
      values[APPLICATION_COLUMNS.OTHER_SUBMISSION_STATUS - 1] || '',
    submissionDeadline: submissionDeadline instanceof Date
      ? Utilities.formatDate(submissionDeadline, timeZone, 'yyyy-MM-dd')
      : '',
    submissionDate: submissionDate instanceof Date
      ? Utilities.formatDate(submissionDate, timeZone, 'yyyy-MM-dd')
      : '',
    submissionMemo: values[APPLICATION_COLUMNS.SUBMISSION_MEMO - 1] || '',
    waitingStatus: values[APPLICATION_COLUMNS.WAITING_STATUS - 1] || ''
  };
}
