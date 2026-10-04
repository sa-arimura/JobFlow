/**
 * JobFlowの初期セットアップを実行します。
 * try-catchは、途中で問題が起きたときに原因を記録し、成功と誤表示しないための仕組みです。
 */
function setupJobFlow() {
  const warnings = [];

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getOrCreateJobFlowSheets(spreadsheet);

    validateStage8MigrationPreconditions(sheets);

    setupSettingsSheet(sheets.settings);
    setupApplicationsSheet(sheets.applications, sheets.settings, warnings);
    setupDashboardSheet(sheets.dashboard, sheets.settings, sheets.applications);
    setupAnalysisSheet(sheets.analysis, sheets.applications);
    orderJobFlowSheets(spreadsheet, sheets);

    SpreadsheetApp.flush();
    showSetupCompleteDialog(COMPLETED_SETUP_ITEMS, warnings);
  } catch (error) {
    logSetupError(error);
    showSetupErrorDialog(error);
    throw error;
  }
}

function validateStage8MigrationPreconditions(sheets) {
  validateNeedsReviewSettingForMigration(sheets.settings);
  validateContactPersonColumn(sheets.applications);
  validateCalculationColumns(sheets.applications);
  validateSubmissionColumns(sheets.applications);
  validateWaitingColumns(sheets.applications);

  const dashboardTitle = String(
    sheets.dashboard.getRange(DASHBOARD_VALUES.TITLE_CELL).getDisplayValue()
  ).trim();
  if (dashboardTitle !== '') {
    const applicationLastRow = sheets.applications.getMaxRows();
    validateDashboardSelectionStatuses(sheets.settings);
    validateDashboardLayout(sheets.dashboard, applicationLastRow);
  }
}

/**
 * 同名シートがあるかを先に確認し、既存シートを削除または作り直さないようにします。
 */
function getOrCreateJobFlowSheets(spreadsheet) {
  return {
    applications: getOrCreateSheet(spreadsheet, SHEET_NAMES.APPLICATIONS),
    dashboard: getOrCreateSheet(spreadsheet, SHEET_NAMES.DASHBOARD),
    analysis: getOrCreateSheet(spreadsheet, SHEET_NAMES.ANALYSIS),
    settings: getOrCreateSheet(spreadsheet, SHEET_NAMES.SETTINGS)
  };
}

function getOrCreateSheet(spreadsheet, sheetName) {
  const existingSheet = spreadsheet.getSheetByName(sheetName);
  return existingSheet || spreadsheet.insertSheet(sheetName);
}

/**
 * JobFlowの3シートが現在占めている位置だけを使って並べ替えます。
 * そのため、無関係なシートの最終的な位置や内容は変更しません。
 */
function orderJobFlowSheets(spreadsheet, sheets) {
  const activeSheetBeforeOrdering = spreadsheet.getActiveSheet();
  const orderedJobFlowSheets = [
    sheets.applications,
    sheets.dashboard,
    sheets.analysis,
    sheets.settings
  ];
  const currentSheets = spreadsheet.getSheets();
  const targetPositions = orderedJobFlowSheets
    .map(function(sheet) {
      return findSheetPositionById(currentSheets, sheet);
    })
    .sort(function(firstPosition, secondPosition) {
      return firstPosition - secondPosition;
    });

  orderedJobFlowSheets.forEach(function(sheet, index) {
    const targetPosition = targetPositions[index];
    const latestSheets = spreadsheet.getSheets();
    const currentPosition = findSheetPositionById(latestSheets, sheet);
    const sheetAtTargetPosition = latestSheets[targetPosition - 1];

    if (currentPosition === targetPosition) {
      return;
    }

    /*
     * 2枚のJobFlowシートを交換すると、間にある無関係なシートは最終的に元の位置へ戻ります。
     * 最初に目的のシートを移動し、押し出されたシートを空いた位置へ戻します。
     */
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(targetPosition);
    spreadsheet.setActiveSheet(sheetAtTargetPosition);
    spreadsheet.moveActiveSheet(currentPosition);
  });

  const activeSheetToRestore = spreadsheet.getSheets().find(function(candidateSheet) {
    return candidateSheet.getSheetId() === activeSheetBeforeOrdering.getSheetId();
  });
  if (activeSheetToRestore) {
    spreadsheet.setActiveSheet(activeSheetToRestore);
  }
}

function findSheetPositionById(sheetList, targetSheet) {
  const targetSheetId = targetSheet.getSheetId();
  const sheetIndex = sheetList.findIndex(function(candidateSheet) {
    return candidateSheet.getSheetId() === targetSheetId;
  });

  if (sheetIndex === -1) {
    throw new Error(
      'シート位置を取得できませんでした: ' + targetSheet.getName()
    );
  }

  return sheetIndex + 1;
}

function setupApplicationsSheet(sheet, settingsSheet, warnings) {
  const requiresWaitingMigration = validateWaitingColumns(sheet);
  validateContactPersonColumn(sheet);
  validateCalculationColumns(sheet);
  validateSubmissionColumns(sheet);
  ensureSheetSize(sheet, DATA_END_ROW, APPLICATION_COLUMN_COUNT);
  setApplicationHeaders(sheet, warnings);
  if (requiresWaitingMigration) {
    initializeWaitingColumnsForMigration(sheet);
  }
  applyApplicationBasicFormatting(sheet);
  applyApplicationColumnWidthsAndFormats(sheet);
  applyStatusValidation(sheet, settingsSheet, warnings);
  applySubmissionStatusValidation(sheet, warnings);
  applyWaitingStatusValidation(sheet, warnings);
  applyApplicationConditionalFormatting(sheet);
  applyCalculationFormulas(sheet, settingsSheet);
}

/**
 * Z/AAが未導入の第7段階構成か、導入済みの第8段階構成かを判定します。
 * 部分導入や未知の内容がある場合は、既存データを守るため停止します。
 */
function validateWaitingColumns(sheet) {
  const waitingStatusColumn = APPLICATION_COLUMNS.WAITING_STATUS;
  const waitingStartDateColumn = APPLICATION_COLUMNS.WAITING_START_DATE;
  const maxColumns = sheet.getMaxColumns();
  const statusHeader = maxColumns >= waitingStatusColumn
    ? String(sheet.getRange(APPLICATION_HEADER_ROW, waitingStatusColumn)
        .getDisplayValue()).trim()
    : '';
  const startDateHeader = maxColumns >= waitingStartDateColumn
    ? String(sheet.getRange(APPLICATION_HEADER_ROW, waitingStartDateColumn)
        .getDisplayValue()).trim()
    : '';
  const expectedStatusHeader = APPLICATION_HEADERS[waitingStatusColumn - 1];
  const expectedStartDateHeader = APPLICATION_HEADERS[waitingStartDateColumn - 1];

  if (statusHeader === expectedStatusHeader &&
      startDateHeader === expectedStartDateHeader) {
    return false;
  }
  if (statusHeader !== '' || startDateHeader !== '') {
    throw new Error(
      '応募一覧のZ/AA列が第7段階の未導入状態または第8段階の正式構成と' +
      '一致しません。対応待ち管理列は変更していません。'
    );
  }

  const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
  [waitingStatusColumn, waitingStartDateColumn].forEach(function(column) {
    if (column > maxColumns) {
      return;
    }
    const range = sheet.getRange(
      DATA_START_ROW,
      column,
      lastRow - DATA_START_ROW + 1,
      1
    );
    const values = range.getDisplayValues();
    const formulas = range.getFormulas();
    const validations = range.getDataValidations();
    const conflictIndex = values.findIndex(function(row, index) {
      return String(row[0]).trim() !== '' ||
        formulas[index][0] !== '' ||
        validations[index][0] !== null;
    });
    if (conflictIndex !== -1) {
      throw new Error(
        sheet.getRange(DATA_START_ROW + conflictIndex, column).getA1Notation() +
        'に既存の値・数式・入力規則があるため、対応待ち管理列を追加できません。'
      );
    }
  });
  return true;
}

/**
 * 第8段階の初回移行時だけ、応募IDがある既存行を「なし」で初期化します。
 */
function initializeWaitingColumnsForMigration(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    return;
  }
  const rowCount = lastRow - DATA_START_ROW + 1;
  const applicationIds = sheet
    .getRange(DATA_START_ROW, APPLICATION_COLUMNS.APPLICATION_ID, rowCount, 1)
    .getDisplayValues();
  const initialValues = applicationIds.map(function(row) {
    return [String(row[0]).trim() === '' ? '' : WAITING_DEFAULT_STATUS];
  });
  sheet
    .getRange(DATA_START_ROW, APPLICATION_COLUMNS.WAITING_STATUS, rowCount, 1)
    .setValues(initialValues);
}

/**
 * S～Y列に未知の見出しや値がある場合は、既存データを守るため停止します。
 */
function validateSubmissionColumns(sheet) {
  if (sheet.getMaxColumns() < APPLICATION_COLUMNS.RESUME_STATUS) {
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), APPLICATION_HEADER_ROW);
  for (
    let column = APPLICATION_COLUMNS.RESUME_STATUS;
    column <= APPLICATION_COLUMNS.SUBMISSION_MEMO;
    column += 1
  ) {
    if (column > sheet.getMaxColumns()) {
      break;
    }
    const expectedHeader = APPLICATION_HEADERS[column - 1];
    const currentHeader = String(
      sheet.getRange(APPLICATION_HEADER_ROW, column).getDisplayValue()
    ).trim();
    if (currentHeader !== '' && currentHeader !== expectedHeader) {
      throw new Error(
        '応募一覧の' +
        sheet.getRange(APPLICATION_HEADER_ROW, column).getA1Notation() +
        'に既存の見出し「' + currentHeader +
        '」があるため、提出物管理列を追加できません。'
      );
    }
    if (currentHeader === '') {
      const hasExistingData = sheet
        .getRange(DATA_START_ROW, column, lastRow - DATA_START_ROW + 1, 1)
        .getDisplayValues()
        .some(function(row) {
          return String(row[0]).trim() !== '';
        });
      if (hasExistingData) {
        throw new Error(
          '応募一覧の' +
          sheet.getRange(APPLICATION_HEADER_ROW, column).getA1Notation() +
          'より下に既存データがあるため、提出物管理列を追加できません。'
        );
      }
    }
  }
}

/**
 * Q列にJobFlow以外の内容がある場合は、既存データを守るためセットアップを停止します。
 */
function validateContactPersonColumn(sheet) {
  if (sheet.getMaxColumns() < APPLICATION_COLUMNS.CONTACT_PERSON) {
    return;
  }

  const lastRow = Math.max(sheet.getLastRow(), APPLICATION_HEADER_ROW);
  const contactValues = sheet
    .getRange(
      APPLICATION_HEADER_ROW,
      APPLICATION_COLUMNS.CONTACT_PERSON,
      lastRow,
      1
    )
    .getDisplayValues()
    .map(function(row) {
      return String(row[0]).trim();
    });
  const currentHeader = contactValues[0];
  const hasExistingData = contactValues.slice(1).some(function(value) {
    return value !== '';
  });

  if (currentHeader !== '' && currentHeader !== '担当者') {
    throw new Error(
      '応募一覧のQ1に既存の見出し「' + currentHeader +
      '」があるため、担当者列を作成できません。'
    );
  }
  if (currentHeader === '' && hasExistingData) {
    throw new Error(
      '応募一覧のQ列に既存データがあるため、担当者列を作成できません。'
    );
  }
}

/**
 * 計算列に未知の値や数式がある場合は、上書きせずセットアップを停止します。
 * すでにJobFlowのARRAYFORMULAがある再実行では、その数式を安全に再利用します。
 */
function validateCalculationColumns(sheet) {
  const calculationColumns = [
    {
      column: APPLICATION_COLUMNS.ELAPSED_DAYS,
      header: '応募経過日数',
      formula: ELAPSED_DAYS_FORMULA
    },
    {
      column: APPLICATION_COLUMNS.WAITING_DAYS,
      header: '返信待ち日数',
      formula: WAITING_DAYS_FORMULA
    },
    {
      column: APPLICATION_COLUMNS.NEEDS_REVIEW,
      header: '要確認',
      formula: NEEDS_REVIEW_FORMULA
    },
    {
      column: APPLICATION_COLUMNS.DAYS_UNTIL_INTERVIEW,
      header: '面接までの日数',
      formula: DAYS_UNTIL_INTERVIEW_FORMULA
    }
  ];

  calculationColumns.forEach(function(definition) {
    validateCalculationColumn(sheet, definition);
  });
}

function validateCalculationColumn(sheet, definition) {
  if (sheet.getMaxColumns() < definition.column) {
    return;
  }

  const header = String(
    sheet.getRange(APPLICATION_HEADER_ROW, definition.column).getDisplayValue()
  ).trim();
  if (header !== '' && header !== definition.header) {
    throw new Error(
      '応募一覧の' +
      sheet.getRange(APPLICATION_HEADER_ROW, definition.column).getA1Notation() +
      'に既存の見出し「' + header + '」があるため、計算列を作成できません。'
    );
  }

  const formulaCell = sheet.getRange(DATA_START_ROW, definition.column);
  const existingFormula = formulaCell.getFormula();
  const legacyFormulas = getLegacyCalculationFormulas(definition.column);
  const isOwnedFormula = existingFormula === definition.formula ||
    legacyFormulas.indexOf(existingFormula) !== -1;
  if (isOwnedFormula) {
    if (String(formulaCell.getDisplayValue()).indexOf('#REF!') !== -1) {
      throw new Error(
        definition.header +
        'のARRAYFORMULAを展開できません。列内の既存値を確認してください。'
      );
    }
    return;
  }
  if (existingFormula !== '') {
    throw new Error(
      formulaCell.getA1Notation() +
      'に既存の数式があるため、' + definition.header + 'を設定できません。'
    );
  }

  const lastRow = Math.max(sheet.getLastRow(), DATA_START_ROW);
  const values = sheet
    .getRange(DATA_START_ROW, definition.column, lastRow - DATA_START_ROW + 1, 1)
    .getDisplayValues();
  const formulas = sheet
    .getRange(DATA_START_ROW, definition.column, lastRow - DATA_START_ROW + 1, 1)
    .getFormulas();
  const conflictIndex = values.findIndex(function(row, index) {
    return String(row[0]).trim() !== '' || formulas[index][0] !== '';
  });

  if (conflictIndex !== -1) {
    const conflictCell = sheet.getRange(
      DATA_START_ROW + conflictIndex,
      definition.column
    );
    throw new Error(
      conflictCell.getA1Notation() +
      'に既存の値または数式があるため、' + definition.header + 'を設定できません。'
    );
  }
}

function getLegacyCalculationFormulas(column) {
  if (column === APPLICATION_COLUMNS.WAITING_DAYS) {
    return [PREVIOUS_WAITING_DAYS_FORMULA];
  }
  if (column === APPLICATION_COLUMNS.NEEDS_REVIEW) {
    return [
      PREVIOUS_NEEDS_REVIEW_FORMULA,
      PREVIOUS_NEEDS_REVIEW_FORMULA_V2
    ];
  }
  return [];
}

/**
 * 空欄だけへ見出しを入れ、異なる既存値は上書きせず警告として記録します。
 */
function setApplicationHeaders(sheet, warnings) {
  const headerRange = sheet.getRange(
    APPLICATION_HEADER_ROW,
    1,
    1,
    APPLICATION_COLUMN_COUNT
  );
  const currentHeaders = headerRange.getValues()[0];

  APPLICATION_HEADERS.forEach(function(expectedHeader, index) {
    const currentValue = currentHeaders[index];
    if (isEmptyCellValue(currentValue)) {
      sheet.getRange(APPLICATION_HEADER_ROW, index + 1).setValue(expectedHeader);
    } else if (String(currentValue) !== expectedHeader) {
      const cellAddress = sheet.getRange(APPLICATION_HEADER_ROW, index + 1).getA1Notation();
      warnings.push(
        SHEET_NAMES.APPLICATIONS + '!' + cellAddress +
        '（現在値: ' + formatWarningValue(currentValue) + '）'
      );
    }
  });
}

function applyApplicationBasicFormatting(sheet) {
  const headerRange = sheet.getRange(
    APPLICATION_HEADER_ROW,
    1,
    1,
    APPLICATION_COLUMN_COUNT
  );
  const dataRange = sheet.getRange(
    DATA_START_ROW,
    1,
    DATA_END_ROW - DATA_START_ROW + 1,
    APPLICATION_COLUMN_COUNT
  );

  sheet.setFrozenRows(APPLICATION_HEADER_ROW);
  headerRange
    .setFontWeight('bold')
    .setBackground(HEADER_BACKGROUND_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setHorizontalAlignment('center');
  dataRange.setVerticalAlignment('middle');

  sheet
    .getRange(
      DATA_START_ROW,
      APPLICATION_COLUMNS.JOB_URL,
      DATA_END_ROW - DATA_START_ROW + 1,
      1
    )
    .setWrap(true);
  sheet
    .getRange(
      DATA_START_ROW,
      APPLICATION_COLUMNS.MEMO,
      DATA_END_ROW - DATA_START_ROW + 1,
      1
    )
    .setWrap(true);

  applyApplicationFilter(sheet);
}

function applyApplicationFilter(sheet) {
  const managedRowCount = DATA_END_ROW;
  const physicalRowCount = sheet.getMaxRows();
  const allowedRowCounts = [managedRowCount, physicalRowCount];
  const knownColumnCounts = [
    APPLICATION_COLUMNS.UPDATED_DATE_TIME,
    APPLICATION_COLUMNS.SUBMISSION_MEMO,
    APPLICATION_COLUMN_COUNT
  ];
  const expectedRange = sheet.getRange(
    APPLICATION_HEADER_ROW,
    1,
    physicalRowCount,
    APPLICATION_COLUMN_COUNT
  );
  const existingFilter = sheet.getFilter();
  if (!existingFilter) {
    expectedRange.createFilter();
    return;
  }
  if (existingFilter.getRange().getA1Notation() === expectedRange.getA1Notation()) {
    return;
  }

  const existingRange = existingFilter.getRange();
  const isKnownStart = existingRange.getRow() === APPLICATION_HEADER_ROW &&
    existingRange.getColumn() === 1;
  const isKnownRowCount = allowedRowCounts.indexOf(existingRange.getNumRows()) !== -1;
  const isKnownColumnCount = knownColumnCounts.indexOf(existingRange.getNumColumns()) !== -1;
  if (!isKnownStart || !isKnownRowCount || !isKnownColumnCount) {
    throw new Error(
      '応募一覧の既存フィルター範囲がJobFlowの既知の範囲と一致しないため、' +
      '対応待ち管理列まで拡張できません。'
    );
  }

  const criteria = [];
  try {
    for (let column = 1; column <= existingRange.getNumColumns(); column += 1) {
      criteria[column] = existingFilter.getColumnFilterCriteria(column);
    }
  } catch (error) {
    throw new Error(
      '応募一覧の既存フィルター条件を安全に保持できないため、' +
      'フィルター範囲は変更していません。元のエラー: ' + error.message
    );
  }
  existingFilter.remove();
  const newFilter = expectedRange.createFilter();
  criteria.forEach(function(criterion, column) {
    if (criterion) {
      newFilter.setColumnFilterCriteria(column, criterion);
    }
  });
}

function applyApplicationColumnWidthsAndFormats(sheet) {
  APPLICATION_COLUMN_WIDTHS.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  const rowCount = DATA_END_ROW - DATA_START_ROW + 1;
  const dateColumns = [
    APPLICATION_COLUMNS.APPLICATION_DATE,
    APPLICATION_COLUMNS.LAST_REPLY_DATE,
    APPLICATION_COLUMNS.NEXT_SCHEDULE_DATE
  ];
  const dateTimeColumns = [
    APPLICATION_COLUMNS.INTERVIEW_DATE_TIME,
    APPLICATION_COLUMNS.CREATED_DATE_TIME,
    APPLICATION_COLUMNS.UPDATED_DATE_TIME
  ];
  const integerColumns = [
    APPLICATION_COLUMNS.ELAPSED_DAYS,
    APPLICATION_COLUMNS.WAITING_DAYS,
    APPLICATION_COLUMNS.DAYS_UNTIL_INTERVIEW
  ];

  dateColumns.forEach(function(column) {
    sheet.getRange(DATA_START_ROW, column, rowCount, 1).setNumberFormat(DATE_FORMAT);
  });
  dateTimeColumns.forEach(function(column) {
    sheet.getRange(DATA_START_ROW, column, rowCount, 1).setNumberFormat(DATE_TIME_FORMAT);
  });
  integerColumns.forEach(function(column) {
    sheet.getRange(DATA_START_ROW, column, rowCount, 1).setNumberFormat(INTEGER_FORMAT);
  });
  [
    APPLICATION_COLUMNS.SUBMISSION_DEADLINE,
    APPLICATION_COLUMNS.SUBMISSION_DATE,
    APPLICATION_COLUMNS.WAITING_START_DATE
  ].forEach(function(column) {
    sheet.getRange(DATA_START_ROW, column, rowCount, 1).setNumberFormat(DATE_FORMAT);
  });
  sheet
    .getRange(DATA_START_ROW, APPLICATION_COLUMNS.SUBMISSION_MEMO, rowCount, 1)
    .setWrap(true);
}

/**
 * ARRAYFORMULAを2行目へ1つだけ設定し、新しく追加された応募にも自動適用します。
 */
function applyCalculationFormulas(sheet, settingsSheet) {
  validateTerminalSelectionStatuses(settingsSheet);

  const formulas = [
    [APPLICATION_COLUMNS.ELAPSED_DAYS, ELAPSED_DAYS_FORMULA],
    [APPLICATION_COLUMNS.WAITING_DAYS, WAITING_DAYS_FORMULA],
    [APPLICATION_COLUMNS.NEEDS_REVIEW, NEEDS_REVIEW_FORMULA],
    [APPLICATION_COLUMNS.DAYS_UNTIL_INTERVIEW, DAYS_UNTIL_INTERVIEW_FORMULA]
  ];

  formulas.forEach(function(definition) {
    const formulaCell = sheet.getRange(DATA_START_ROW, definition[0]);
    const currentFormula = formulaCell.getFormula();
    const legacyFormulas = getLegacyCalculationFormulas(definition[0]);
    if (currentFormula === '' || legacyFormulas.indexOf(currentFormula) !== -1) {
      formulaCell.setFormula(definition[1]);
    }
  });
}

/**
 * 終了状態は設定シートに実在する値だけを使用します。
 */
function validateTerminalSelectionStatuses(settingsSheet) {
  const configuredStatuses = settingsSheet
    .getRange(
      SETTINGS_STATUS_START_ROW,
      SETTINGS_STATUS_COLUMN,
      SELECTION_STATUSES.length,
      1
    )
    .getDisplayValues()
    .map(function(row) {
      return String(row[0]).trim();
    });
  const missingStatuses = TERMINAL_SELECTION_STATUSES.filter(function(status) {
    return configuredStatuses.indexOf(status) === -1;
  });

  if (missingStatuses.length > 0) {
    throw new Error(
      '設定シートに終了状態が見つかりません: ' + missingStatuses.join('、')
    );
  }
}

/**
 * データ入力規則は、セルへ入力できる値を制限し、入力ミスを防ぐ仕組みです。
 * 無関係な既存規則は保持し、JobFlowの規則または規則がないセルだけを更新します。
 */
function applyStatusValidation(sheet, settingsSheet, warnings) {
  const dataRows = getExistingApplicationDataRows(sheet);
  if (dataRows.length === 0) return;
  const lastDataRow = dataRows[dataRows.length - 1];
  const rowCount = lastDataRow - DATA_START_ROW + 1;
  const targetRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.SELECTION_STATUS,
    rowCount,
    1
  );
  const statusSourceRange = settingsSheet.getRange(
    SETTINGS_STATUS_START_ROW,
    SETTINGS_STATUS_COLUMN,
    SELECTION_STATUSES.length,
    1
  );
  const jobFlowRule = buildJobFlowStatusValidation(statusSourceRange);
  const existingRules = targetRange.getDataValidations();
  let hasValidationChanges = false;
  const dataRowSet = {};
  dataRows.forEach(function(row) { dataRowSet[row] = true; });
  const updatedRules = existingRules.map(function(rowRules, rowIndex) {
    const actualRow = DATA_START_ROW + rowIndex;
    if (!dataRowSet[actualRow]) return rowRules;
    const existingRule = rowRules[0];

    if (!existingRule) {
      hasValidationChanges = true;
      return [jobFlowRule];
    }
    if (isJobFlowStatusValidation(existingRule, statusSourceRange)) {
      return [existingRule];
    }

    const cellAddress = sheet
      .getRange(DATA_START_ROW + rowIndex, APPLICATION_COLUMNS.SELECTION_STATUS)
      .getA1Notation();
    warnings.push(
      SHEET_NAMES.APPLICATIONS + '!' + cellAddress +
      '（既存のデータ入力規則を保持）'
    );
    return [existingRule];
  });

  if (hasValidationChanges) {
    targetRange.setDataValidations(updatedRules);
  }
}

function getExistingApplicationDataRows(sheet) {
  const rowCount = Math.max(sheet.getMaxRows() - DATA_START_ROW + 1, 0);
  if (rowCount === 0) return [];
  const values = sheet.getRange(DATA_START_ROW, 1, rowCount, APPLICATION_COLUMN_COUNT)
    .getDisplayValues();
  const rows = [];
  values.forEach(function(rowValues, index) {
    const realValues = rowValues.slice(0, 11)
      .concat(rowValues.slice(14, 17))
      .concat(rowValues.slice(18, 27));
    if (realValues.some(function(value) { return String(value).trim() !== ''; })) {
      rows.push(DATA_START_ROW + index);
    }
  });
  return rows;
}

function ensureApplicationStatusValidation(sheet, row, settingsSheet) {
  const statusSourceRange = settingsSheet.getRange(
    SETTINGS_STATUS_START_ROW, SETTINGS_STATUS_COLUMN, SELECTION_STATUSES.length, 1
  );
  const cell = sheet.getRange(row, APPLICATION_COLUMNS.SELECTION_STATUS);
  const existingRule = cell.getDataValidation();
  if (!existingRule) {
    cell.setDataValidation(buildJobFlowStatusValidation(statusSourceRange));
    return;
  }
  if (isJobFlowStatusValidation(existingRule, statusSourceRange)) return;
  throw new Error(
    SHEET_NAMES.APPLICATIONS + '!' + cell.getA1Notation() +
    '（既存の未知入力規則があるため登録を停止しました。データは変更されていません。）'
  );
}

function buildJobFlowStatusValidation(statusSourceRange) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(statusSourceRange, true)
    .setAllowInvalid(false)
    .setHelpText(STATUS_HELP_TEXT)
    .build();
}

function isJobFlowStatusValidation(rule, statusSourceRange) {
  if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return false;
  }

  const criteriaValues = rule.getCriteriaValues();
  const sourceRange = criteriaValues[0];
  return sourceRange &&
    sourceRange.getSheet().getSheetId() === statusSourceRange.getSheet().getSheetId() &&
    sourceRange.getA1Notation() === statusSourceRange.getA1Notation();
}

/**
 * 提出物4列へ同じプルダウンを設定します。
 * 同じ規則は再利用し、未知の既存規則は上書きしません。
 */
function applySubmissionStatusValidation(sheet, warnings) {
  const rowCount = DATA_END_ROW - DATA_START_ROW + 1;
  const targetRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.RESUME_STATUS,
    rowCount,
    APPLICATION_COLUMNS.OTHER_SUBMISSION_STATUS -
      APPLICATION_COLUMNS.RESUME_STATUS + 1
  );
  const existingRules = targetRange.getDataValidations();
  const expectedRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SUBMISSION_STATUSES, true)
    .setAllowInvalid(false)
    .setHelpText(SUBMISSION_STATUS_HELP_TEXT)
    .build();
  let hasChanges = false;
  const updatedRules = existingRules.map(function(rowRules, rowIndex) {
    return rowRules.map(function(existingRule, columnIndex) {
      if (!existingRule) {
        hasChanges = true;
        return expectedRule;
      }
      if (isJobFlowSubmissionValidation(existingRule)) {
        return existingRule;
      }
      warnings.push(
        SHEET_NAMES.APPLICATIONS + '!' +
        targetRange.getCell(rowIndex + 1, columnIndex + 1).getA1Notation() +
        '（既存の入力規則を維持）'
      );
      return existingRule;
    });
  });
  if (hasChanges) {
    targetRange.setDataValidations(updatedRules);
  }
}

function applyWaitingStatusValidation(sheet, warnings) {
  const rowCount = DATA_END_ROW - DATA_START_ROW + 1;
  const targetRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.WAITING_STATUS,
    rowCount,
    1
  );
  const expectedRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(WAITING_STATUSES, true)
    .setAllowInvalid(false)
    .setHelpText(WAITING_STATUS_HELP_TEXT)
    .build();
  const existingRules = targetRange.getDataValidations();
  let hasChanges = false;
  const updatedRules = existingRules.map(function(rowRules, rowIndex) {
    const existingRule = rowRules[0];
    if (!existingRule) {
      hasChanges = true;
      return [expectedRule];
    }
    if (isJobFlowWaitingStatusValidation(existingRule)) {
      return [existingRule];
    }
    warnings.push(
      SHEET_NAMES.APPLICATIONS + '!' +
      targetRange.getCell(rowIndex + 1, 1).getA1Notation() +
      '（既存の入力規則を維持）'
    );
    return [existingRule];
  });
  if (hasChanges) {
    targetRange.setDataValidations(updatedRules);
  }
}

function isJobFlowWaitingStatusValidation(rule) {
  if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return false;
  }
  const criteriaValues = rule.getCriteriaValues();
  const values = criteriaValues.length > 0 ? criteriaValues[0] : [];
  return Array.isArray(values) &&
    values.length === WAITING_STATUSES.length &&
    values.every(function(value, index) {
      return String(value) === WAITING_STATUSES[index];
    });
}

function isJobFlowSubmissionValidation(rule) {
  if (rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return false;
  }
  const criteriaValues = rule.getCriteriaValues();
  const values = criteriaValues.length > 0 ? criteriaValues[0] : [];
  return Array.isArray(values) &&
    values.length === SUBMISSION_STATUSES.length &&
    values.every(function(value, index) {
      return String(value) === SUBMISSION_STATUSES[index];
    });
}

/**
 * 条件付き書式は、値に応じて見た目を自動変更する仕組みです。
 * JobFlowのステータス規則だけを差し替え、その他の規則は残します。
 */
function applyStatusConditionalFormatting(sheet) {
  const targetRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.SELECTION_STATUS,
    DATA_END_ROW - DATA_START_ROW + 1,
    1
  );
  const preservedRules = sheet.getConditionalFormatRules().filter(function(rule) {
    return !isJobFlowStatusConditionalRule(rule, targetRange);
  });
  const statusRules = SELECTION_STATUSES.map(function(status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(STATUS_BACKGROUND_COLORS[status])
      .setFontColor(STATUS_TEXT_COLORS[status])
      .setRanges([targetRange])
      .build();
  });

  sheet.setConditionalFormatRules(preservedRules.concat(statusRules));
}

/**
 * G列とN列のJobFlow条件付き書式を一度の書き込みで更新します。
 * 利用者が作成した無関係なルールはそのまま保持します。
 */
function applyApplicationConditionalFormatting(sheet) {
  const statusRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.SELECTION_STATUS,
    DATA_END_ROW - DATA_START_ROW + 1,
    1
  );
  const needsReviewRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.NEEDS_REVIEW,
    DATA_END_ROW - DATA_START_ROW + 1,
    1
  );
  const submissionRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.RESUME_STATUS,
    DATA_END_ROW - DATA_START_ROW + 1,
    APPLICATION_COLUMNS.OTHER_SUBMISSION_STATUS -
      APPLICATION_COLUMNS.RESUME_STATUS + 1
  );
  const preservedRules = sheet.getConditionalFormatRules().filter(function(rule) {
    return !isJobFlowStatusConditionalRule(rule, statusRange) &&
      !isNeedsReviewConditionalRule(rule, needsReviewRange) &&
      !isJobFlowSubmissionConditionalRule(rule, submissionRange);
  });
  const statusRules = SELECTION_STATUSES.map(function(status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(STATUS_BACKGROUND_COLORS[status])
      .setFontColor(STATUS_TEXT_COLORS[status])
      .setRanges([statusRange])
      .build();
  });
  const needsReviewRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(NEEDS_REVIEW_TEXT)
    .setBackground(NEEDS_REVIEW_BACKGROUND_COLOR)
    .setRanges([needsReviewRange])
    .build();
  const submissionRules = SUBMISSION_STATUSES.map(function(status) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(status)
      .setBackground(SUBMISSION_BACKGROUND_COLORS[status])
      .setFontColor('#000000')
      .setRanges([submissionRange])
      .build();
  });

  sheet.setConditionalFormatRules(
    preservedRules
      .concat(statusRules)
      .concat([needsReviewRule])
      .concat(submissionRules)
  );
}

function isJobFlowSubmissionConditionalRule(rule, targetRange) {
  const condition = rule.getBooleanCondition();
  if (!condition ||
      condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) {
    return false;
  }
  const criteriaValues = condition.getCriteriaValues();
  const status = criteriaValues.length > 0 ? String(criteriaValues[0]) : '';
  const ranges = rule.getRanges();
  return SUBMISSION_STATUSES.indexOf(status) !== -1 &&
    ranges.length === 1 &&
    ranges[0].getSheet().getSheetId() === targetRange.getSheet().getSheetId() &&
    ranges[0].getA1Notation() === targetRange.getA1Notation();
}

/**
 * N列が「要確認」のセルだけを淡い赤色にします。
 * 同じJobFlowルールを取り除いてから1件だけ追加し、再実行時の重複を防ぎます。
 */
function applyNeedsReviewConditionalFormatting(sheet) {
  const targetRange = sheet.getRange(
    DATA_START_ROW,
    APPLICATION_COLUMNS.NEEDS_REVIEW,
    DATA_END_ROW - DATA_START_ROW + 1,
    1
  );
  const preservedRules = sheet.getConditionalFormatRules().filter(function(rule) {
    return !isNeedsReviewConditionalRule(rule, targetRange);
  });
  const needsReviewRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(NEEDS_REVIEW_TEXT)
    .setBackground(NEEDS_REVIEW_BACKGROUND_COLOR)
    .setRanges([targetRange])
    .build();

  sheet.setConditionalFormatRules(preservedRules.concat([needsReviewRule]));
}

function isNeedsReviewConditionalRule(rule, targetRange) {
  const condition = rule.getBooleanCondition();
  if (!condition ||
      condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) {
    return false;
  }

  const criteriaValues = condition.getCriteriaValues();
  const ranges = rule.getRanges();
  return criteriaValues.length > 0 &&
    String(criteriaValues[0]) === NEEDS_REVIEW_TEXT &&
    ranges.length === 1 &&
    ranges[0].getSheet().getSheetId() === targetRange.getSheet().getSheetId() &&
    ranges[0].getA1Notation() === targetRange.getA1Notation();
}

function isJobFlowStatusConditionalRule(rule, targetRange) {
  const condition = rule.getBooleanCondition();
  if (!condition ||
      condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) {
    return false;
  }

  const criteriaValues = condition.getCriteriaValues();
  const status = criteriaValues.length > 0 ? String(criteriaValues[0]) : '';
  const ranges = rule.getRanges();

  return SELECTION_STATUSES.indexOf(status) !== -1 &&
    ranges.length === 1 &&
    ranges[0].getSheet().getSheetId() === targetRange.getSheet().getSheetId() &&
    ranges[0].getA1Notation() === targetRange.getA1Notation();
}

function setupSettingsSheet(sheet) {
  ensureSheetSize(sheet, SETTINGS_STATUS_END_ROW, SETTINGS_COLUMN_WIDTHS.length);

  migrateNeedsReviewSetting(sheet);

  const statusSettingsValues = [['選考状況', '背景色', '文字色', '表示順']];
  SELECTION_STATUSES.forEach(function(status, index) {
    statusSettingsValues.push([
      status,
      STATUS_BACKGROUND_COLORS[status],
      STATUS_TEXT_COLORS[status],
      index + 1
    ]);
  });

  setEmptyCellsOnly(sheet.getRange(1, 1, BASIC_SETTINGS_VALUES.length, 3), BASIC_SETTINGS_VALUES);
  setEmptyCellsOnly(
    sheet.getRange(1, SETTINGS_STATUS_COLUMN, statusSettingsValues.length, 4),
    statusSettingsValues
  );

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, SETTINGS_COLUMN_WIDTHS.length)
    .setFontWeight('bold')
    .setBackground(SETTINGS_HEADER_BACKGROUND_COLOR);

  SETTINGS_COLUMN_WIDTHS.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });
}

function migrateNeedsReviewSetting(sheet) {
  const migrationMode = validateNeedsReviewSettingForMigration(sheet);
  if (migrationMode !== 'legacy') {
    return;
  }
  const expectedValues = BASIC_SETTINGS_VALUES[1];
  sheet.getRange('A2:C2').setValues([[
    expectedValues[0],
    expectedValues[1],
    expectedValues[2]
  ]]);
}

function validateNeedsReviewSettingForMigration(sheet) {
  const currentValues = sheet.getRange('A2:C2').getValues()[0];
  const currentLabel = String(currentValues[0]).trim();
  const currentValue = currentValues[1];
  const currentDescription = String(currentValues[2]).trim();
  const expectedValues = BASIC_SETTINGS_VALUES[1];

  if (currentLabel === '' && isEmptyCellValue(currentValue) &&
      currentDescription === '') {
    return 'empty';
  }
  if (currentLabel === expectedValues[0] &&
      currentDescription === expectedValues[2]) {
    return 'current';
  }
  if (currentLabel === PREVIOUS_NEEDS_REVIEW_SETTING.LABEL &&
      Number(currentValue) === PREVIOUS_NEEDS_REVIEW_SETTING.VALUE &&
      currentDescription === PREVIOUS_NEEDS_REVIEW_SETTING.DESCRIPTION) {
    return 'legacy';
  }
  throw new Error(
    '設定!A2:C2が既知の第7段階設定と一致しないため、' +
    '要確認判定日数を移行できません。'
  );
}

/**
 * 設定済みの値を守るため、空欄のセルにだけ初期値を入れます。
 */
function setEmptyCellsOnly(range, initialValues) {
  const currentValues = range.getValues();
  currentValues.forEach(function(row, rowIndex) {
    row.forEach(function(value, columnIndex) {
      if (isEmptyCellValue(value)) {
        range.getCell(rowIndex + 1, columnIndex + 1)
          .setValue(initialValues[rowIndex][columnIndex]);
      }
    });
  });
}

function setupAnalysisSheet(sheet, applicationsSheet) {
  const applicationLastRow = applicationsSheet.getMaxRows();
  ensureSheetSize(
    sheet,
    Math.max(ANALYSIS_INITIAL_ROWS, applicationLastRow + 13),
    ANALYSIS_INITIAL_COLUMNS
  );
  validateAnalysisLayout(sheet);
  applyAnalysisValuesAndFormulas(sheet);
  applyAnalysisFormatting(sheet);
  setupAnalysisCharts(sheet, applicationLastRow);
}

function validateAnalysisLayout(sheet) {
  const managedRanges = [
    sheet.getRange(ANALYSIS_VALUES.TITLE_CELL),
    sheet.getRange(ANALYSIS_VALUES.MONTH_SECTION_RANGE),
    sheet.getRange('A4:B4'),
    sheet.getRange(ANALYSIS_MONTH_LABEL_RANGE),
    sheet.getRange(ANALYSIS_MONTH_COUNT_RANGE),
    sheet.getRange(ANALYSIS_VALUES.SITE_SECTION_RANGE),
    sheet.getRange('A14:B14')
  ];
  const mergedRanges = managedRanges.reduce(function(allRanges, range) {
    return allRanges.concat(range.getMergedRanges());
  }, []);
  if (mergedRanges.length > 0) {
    throw new Error(
      '分析シートのJobFlow管理範囲に結合セルがあるため、集計を設定できません。'
    );
  }

  const fixedValues = [
    [ANALYSIS_VALUES.TITLE_CELL, ANALYSIS_VALUES.TITLE],
    ['A3', ANALYSIS_VALUES.MONTH_SECTION_TITLE],
    ['A4', ANALYSIS_VALUES.MONTH_HEADER],
    ['B4', ANALYSIS_VALUES.MONTH_COUNT_HEADER],
    ['A13', ANALYSIS_VALUES.SITE_SECTION_TITLE],
    ['A14', ANALYSIS_VALUES.SITE_HEADER],
    ['B14', ANALYSIS_VALUES.SITE_COUNT_HEADER]
  ];
  fixedValues.forEach(function(definition) {
    const currentValue = String(sheet.getRange(definition[0]).getDisplayValue()).trim();
    if (currentValue !== '' && currentValue !== definition[1]) {
      throw new Error(
        '分析シート!' + definition[0] +
        'に未知の値「' + currentValue + '」があるため、上書きしません。'
      );
    }
  });

  const expectedMonthlyFormulas = getAnalysisMonthlyFormulaDefinitions();
  expectedMonthlyFormulas.forEach(function(definition) {
    const currentFormula = sheet.getRange(definition.address).getFormula();
    if (currentFormula !== '' && currentFormula !== definition.formula) {
      throw new Error(
        '分析シート!' + definition.address +
        'に未知の数式があるため、上書きしません。'
      );
    }
  });

  const expectedSiteFormula = buildAnalysisSiteSummaryFormula();
  const siteFormula = sheet.getRange(ANALYSIS_SITE_FORMULA_CELL).getFormula();
  const siteAnchorValue = String(
    sheet.getRange(ANALYSIS_SITE_FORMULA_CELL).getDisplayValue()
  ).trim();
  if (siteFormula !== '' && siteFormula !== expectedSiteFormula) {
    throw new Error('分析シート!A15に未知の数式があるため、上書きしません。');
  }
  if (siteFormula === '' && siteAnchorValue !== '') {
    throw new Error('分析シート!A15に未知の値があるため、上書きしません。');
  }

  if (siteFormula === '') {
    const siteAreaValues = sheet.getRange(15, 1, sheet.getMaxRows() - 14, 2)
      .getDisplayValues();
    const siteAreaFormulas = sheet.getRange(15, 1, sheet.getMaxRows() - 14, 2)
      .getFormulas();
    const hasUnknownSiteArea = siteAreaValues.some(function(row, rowIndex) {
      return row.some(function(value, columnIndex) {
        return String(value).trim() !== '' ||
          siteAreaFormulas[rowIndex][columnIndex] !== '';
      });
    });
    if (hasUnknownSiteArea) {
      throw new Error(
        '分析シートの応募サイト別集計予定範囲に未知の値または数式があるため、上書きしません。'
      );
    }
  }
}

function applyAnalysisValuesAndFormulas(sheet) {
  const fixedValues = {
    A1: ANALYSIS_VALUES.TITLE,
    A3: ANALYSIS_VALUES.MONTH_SECTION_TITLE,
    A4: ANALYSIS_VALUES.MONTH_HEADER,
    B4: ANALYSIS_VALUES.MONTH_COUNT_HEADER,
    A13: ANALYSIS_VALUES.SITE_SECTION_TITLE,
    A14: ANALYSIS_VALUES.SITE_HEADER,
    B14: ANALYSIS_VALUES.SITE_COUNT_HEADER
  };
  Object.keys(fixedValues).forEach(function(address) {
    const range = sheet.getRange(address);
    if (String(range.getDisplayValue()).trim() === '') {
      range.setValue(fixedValues[address]);
    }
  });

  getAnalysisMonthlyFormulaDefinitions().forEach(function(definition) {
    const range = sheet.getRange(definition.address);
    if (range.getFormula() !== definition.formula) {
      range.setFormula(definition.formula);
    }
  });

  const siteFormulaCell = sheet.getRange(ANALYSIS_SITE_FORMULA_CELL);
  if (siteFormulaCell.getFormula() !== buildAnalysisSiteSummaryFormula()) {
    siteFormulaCell.setFormula(buildAnalysisSiteSummaryFormula());
  }
}

function getAnalysisMonthlyFormulaDefinitions() {
  const definitions = [];
  for (let index = 0; index < 6; index += 1) {
    const row = 5 + index;
    const monthOffset = 5 - index;
    definitions.push({
      address: 'A' + row,
      formula: '=TEXT(EDATE(TODAY(),-' + monthOffset + '),"yyyy-mm")'
    });
    definitions.push({
      address: 'B' + row,
      formula: '=COUNTIFS(\'応募一覧\'!A2:A,"<>",' +
        '\'応募一覧\'!E2:E,">="&DATE(YEAR(TODAY()),MONTH(TODAY())-' +
        monthOffset + ',1),\'応募一覧\'!E2:E,"<"&DATE(YEAR(TODAY()),MONTH(TODAY())-' +
        (monthOffset - 1) + ',1))'
    });
  }
  return definitions;
}

function buildAnalysisSiteSummaryFormula() {
  return '=IFERROR(CHOOSECOLS(QUERY(FILTER({' +
    'IF(\'応募一覧\'!F2:F="","未入力",\'応募一覧\'!F2:F),' +
    '\'応募一覧\'!A2:A,' +
    'IF(\'応募一覧\'!F2:F="",1,0)},\'応募一覧\'!A2:A<>""),' +
    '"select Col1,count(Col2),max(Col3) group by Col1 ' +
    'order by count(Col2) desc,max(Col3),Col1 asc ' +
    'label Col1 \'\',count(Col2) \'\',max(Col3) \'\'",0),1,2),"")';
}

function applyAnalysisFormatting(sheet) {
  sheet.getRange(ANALYSIS_VALUES.TITLE_CELL)
    .setFontWeight('bold')
    .setFontSize(DASHBOARD_TITLE_FONT_SIZE)
    .setFontColor(DASHBOARD_TITLE_COLOR);
  [ANALYSIS_VALUES.MONTH_SECTION_RANGE, ANALYSIS_VALUES.SITE_SECTION_RANGE]
    .forEach(function(rangeAddress) {
      sheet.getRange(rangeAddress)
        .setBackground(ANALYSIS_SECTION_BACKGROUND_COLOR)
        .setFontWeight('bold');
    });
  ['A4:B4', 'A14:B14'].forEach(function(rangeAddress) {
    sheet.getRange(rangeAddress)
      .setBackground(HEADER_BACKGROUND_COLOR)
      .setFontColor(HEADER_TEXT_COLOR)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });
  ['A5:A10', 'A15:B' + sheet.getMaxRows()].forEach(function(rangeAddress) {
    sheet.getRange(rangeAddress).setVerticalAlignment('middle');
  });
  ['B5:B10', 'B15:B' + sheet.getMaxRows()].forEach(function(rangeAddress) {
    sheet.getRange(rangeAddress)
      .setNumberFormat(INTEGER_FORMAT)
      .setHorizontalAlignment('center');
  });
  sheet.getRange('A5:A10').setNumberFormat('yyyy-mm');
  sheet.getRange('A4:B10').setBorder(
    true, true, true, true, true, true,
    ANALYSIS_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('A14:B' + sheet.getMaxRows()).setBorder(
    true, true, true, true, true, true,
    ANALYSIS_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 30);
  sheet.setColumnWidth(4, 30);
}

function setupDashboardSheet(sheet, settingsSheet, applicationsSheet) {
  const applicationLastRow = applicationsSheet.getMaxRows();
  ensureSheetSize(
    sheet,
    Math.max(30, applicationLastRow + DASHBOARD_SUBMISSION_DEADLINE_START_ROW - 2),
    15
  );
  validateDashboardSelectionStatuses(settingsSheet);
  validateDashboardLayout(sheet, applicationLastRow);
  validateDashboardSubmissionDeadlineLayout(sheet);
  applyDashboardValuesAndFormulas(sheet, applicationLastRow);
  applyDashboardFormatting(sheet);
  applyDashboardSubmissionDeadlineFormatting(sheet);
}

/**
 * ダッシュボードは設定シートに実在する9種類だけを集計対象として使用します。
 */
function validateDashboardSelectionStatuses(settingsSheet) {
  const configuredStatuses = settingsSheet
    .getRange(
      SETTINGS_STATUS_START_ROW,
      SETTINGS_STATUS_COLUMN,
      SELECTION_STATUSES.length,
      1
    )
    .getDisplayValues()
    .map(function(row) {
      return String(row[0]).trim();
    });

  const statusesMatch = configuredStatuses.length === SELECTION_STATUSES.length &&
    configuredStatuses.every(function(status, index) {
      return status === SELECTION_STATUSES[index];
    });
  if (!statusesMatch) {
    throw new Error(
      '設定シートの選考状況一覧がJobFlowの正式な9種類と一致しません。'
    );
  }

  const missingInProgressStatuses = IN_PROGRESS_SELECTION_STATUSES.filter(
    function(status) {
      return configuredStatuses.indexOf(status) === -1;
    }
  );
  if (missingInProgressStatuses.length > 0) {
    throw new Error(
      '選考中の対象状態が設定シートに見つかりません: ' +
      missingInProgressStatuses.join('、')
    );
  }
}

/**
 * ダッシュボードへ利用者の値、未知の数式、グラフ、結合セルがある場合は停止します。
 */
function validateDashboardLayout(sheet, applicationLastRow) {
  const titleValue = String(
    sheet.getRange(DASHBOARD_VALUES.TITLE_CELL).getDisplayValue()
  ).trim();
  if (titleValue !== '' && titleValue !== DASHBOARD_VALUES.TITLE) {
    throw new Error(
      'ダッシュボード!' + DASHBOARD_VALUES.TITLE_CELL +
      'に未知の値「' + titleValue + '」があります。'
    );
  }

  if (sheet.getCharts().length > 0) {
    throw new Error('ダッシュボードに既存のグラフがあるため、集計を設定できません。');
  }

  const interviewRange = getDashboardInterviewRange(sheet, applicationLastRow);
  const mergedRanges = sheet.getRange(DASHBOARD_TARGET_RANGE).getMergedRanges()
    .concat(interviewRange.getMergedRanges());
  if (mergedRanges.length > 0) {
    throw new Error(
      'ダッシュボードの集計予定範囲に結合セルがあるため、集計を設定できません。'
    );
  }

  validateInterviewConditionalFormatting(sheet, interviewRange);

  if (isDashboardAlreadyConfigured(sheet, applicationLastRow)) {
    validateUnusedDashboardCellsAreEmpty(
      sheet,
      getDashboardFormulaDefinitions(applicationLastRow)
    );
    validateDashboardInterviewArea(sheet, applicationLastRow);
    return;
  }

  if (isPreviousDashboardConfigured(sheet, applicationLastRow)) {
    validateUnusedDashboardCellsAreEmpty(
      sheet,
      getStage8DashboardFormulaDefinitions(applicationLastRow),
      STAGE8_DASHBOARD_FIXED_VALUES
    );
    validateDashboardInterviewArea(sheet, applicationLastRow);
    return;
  }

  const initialMessage = String(
    sheet.getRange(DASHBOARD_VALUES.MESSAGE_CELL).getDisplayValue()
  );
  if (initialMessage !== DASHBOARD_VALUES.MESSAGE) {
    throw new Error(
      'ダッシュボード!' + DASHBOARD_VALUES.MESSAGE_CELL +
      'が正規の初期メッセージではないため、上書きしません。'
    );
  }

  const range = sheet.getRange(DASHBOARD_TARGET_RANGE);
  const values = range.getDisplayValues();
  const formulas = range.getFormulas();
  const startRow = range.getRow();
  const startColumn = range.getColumn();

  values.forEach(function(rowValues, rowIndex) {
    rowValues.forEach(function(value, columnIndex) {
      const row = startRow + rowIndex;
      const column = startColumn + columnIndex;
      const cell = sheet.getRange(row, column);
      const isInitialMessageCell =
        cell.getA1Notation() === DASHBOARD_VALUES.MESSAGE_CELL;

      if (!isInitialMessageCell &&
          (String(value).trim() !== '' || formulas[rowIndex][columnIndex] !== '')) {
        throw new Error(
          'ダッシュボード!' + cell.getA1Notation() +
          'に既存の値または数式があるため、集計を設定できません。'
        );
      }
    });
  });

  validateDashboardInterviewArea(sheet, applicationLastRow);
}

function isDashboardAlreadyConfigured(sheet, applicationLastRow) {
  const fixedValuesMatch = Object.keys(DASHBOARD_FIXED_VALUES).every(function(address) {
    return String(sheet.getRange(address).getDisplayValue()) ===
      DASHBOARD_FIXED_VALUES[address];
  });
  const formulasMatch = getDashboardFormulaDefinitions(applicationLastRow)
    .every(function(definition) {
    const formulaCell = sheet.getRange(definition.address);
    const actualFormula = formulaCell.getFormula();
    const isKnownLegacyStatusCountFormula =
      definition.address === 'B10' &&
      actualFormula === PREVIOUS_DASHBOARD_STATUS_COUNTS_FORMULA;
    return (actualFormula === definition.formula ||
      isKnownLegacyStatusCountFormula) &&
      (isKnownLegacyStatusCountFormula ||
        String(formulaCell.getDisplayValue()).indexOf('#') !== 0);
  });

  const interviewValuesMatch = Object.keys(DASHBOARD_INTERVIEW_FIXED_VALUES)
    .every(function(address) {
      return String(sheet.getRange(address).getDisplayValue()) ===
        DASHBOARD_INTERVIEW_FIXED_VALUES[address];
    });

  return fixedValuesMatch && formulasMatch && interviewValuesMatch;
}

function isPreviousDashboardConfigured(sheet, applicationLastRow) {
  const fixedValuesMatch = Object.keys(STAGE8_DASHBOARD_FIXED_VALUES).every(function(address) {
    return String(sheet.getRange(address).getDisplayValue()) ===
      STAGE8_DASHBOARD_FIXED_VALUES[address];
  });
  const formulasMatch = getStage8DashboardFormulaDefinitions(applicationLastRow)
    .every(function(definition) {
      const actualFormula = sheet.getRange(definition.address).getFormula();
      if (definition.address === 'H4') {
        // 一時期のJobFlow正規実装で生成されたSUMPRODUCT式も、
        // Stage8→Stage9移行元として限定的に許容する。
        return actualFormula === definition.formula ||
          actualFormula === buildDashboardUpcomingInterviewsFormula(applicationLastRow);
      }
      return actualFormula === definition.formula;
    });
  return fixedValuesMatch && formulasMatch;
}

function validateUnusedDashboardCellsAreEmpty(
  sheet,
  formulaDefinitions,
  fixedValues
) {
  const range = sheet.getRange(DASHBOARD_TARGET_RANGE);
  const values = range.getDisplayValues();
  const formulas = range.getFormulas();
  const fixedAddresses = Object.keys(fixedValues || DASHBOARD_FIXED_VALUES);
  const formulaAddresses = formulaDefinitions.map(function(definition) {
    return definition.address;
  });
  const statusListRange = sheet.getRange(DASHBOARD_STATUS_LIST_RANGE);

  values.forEach(function(rowValues, rowIndex) {
    rowValues.forEach(function(value, columnIndex) {
      const cell = range.getCell(rowIndex + 1, columnIndex + 1);
      const address = cell.getA1Notation();
      const isFixedCell = fixedAddresses.indexOf(address) !== -1;
      const isFormulaCell = formulaAddresses.indexOf(address) !== -1;
      const isStatusListCell =
        cell.getRow() >= statusListRange.getRow() &&
        cell.getRow() <= statusListRange.getLastRow() &&
        cell.getColumn() >= statusListRange.getColumn() &&
        cell.getColumn() <= statusListRange.getLastColumn();

      if (!isFixedCell && !isFormulaCell && !isStatusListCell &&
          (String(value).trim() !== '' || formulas[rowIndex][columnIndex] !== '')) {
        throw new Error(
          'ダッシュボード!' + address +
          'に未知の値または数式があるため、再設定できません。'
        );
      }
    });
  });
}

/**
 * 面接一覧の範囲に利用者の値や未知の数式がある場合は、削除せず停止します。
 */
function validateDashboardInterviewArea(sheet, applicationLastRow) {
  const range = getDashboardInterviewRange(sheet, applicationLastRow);
  const formulaCell = sheet.getRange(DASHBOARD_INTERVIEW_FORMULA_CELL);
  const expectedFormula = buildDashboardInterviewListFormula(applicationLastRow);
  const currentFormula = formulaCell.getFormula();
  const values = range.getDisplayValues();
  const formulas = range.getFormulas();
  const interviewHeaders = [
    '面接日時',
    '会社名',
    '職種',
    '面接までの日数',
    '求人URL',
    '応募ID'
  ];

  if (currentFormula === expectedFormula) {
    if (String(formulaCell.getDisplayValue()).indexOf('#') === 0) {
      throw new Error(
        'ダッシュボードの面接予定一覧を展開できません。' +
        'J5:Oの既存値を確認してください。既存値は削除していません。'
      );
    }

    /*
     * J列の面接日時または0件時メッセージは必ず値を持つため、
     * J5から連続する行だけをJobFlowのスピル結果として扱います。
     */
    let lastSpillRowIndex = 1;
    for (let rowIndex = 2; rowIndex < values.length; rowIndex += 1) {
      if (String(values[rowIndex][0]).trim() === '') {
        break;
      }
      lastSpillRowIndex = rowIndex;
    }

    values.forEach(function(rowValues, rowIndex) {
      rowValues.forEach(function(value, columnIndex) {
        const actualRow = range.getRow() + rowIndex;
        const actualColumn = range.getColumn() + columnIndex;
        let expectedFixedValue = null;
        if (actualRow === 3 && actualColumn === 10) {
          expectedFixedValue = DASHBOARD_INTERVIEW_FIXED_VALUES.J3;
        } else if (actualRow === 4) {
          expectedFixedValue = interviewHeaders[columnIndex];
        }
        const isKnownFixedValue =
          expectedFixedValue !== null &&
          String(value) === expectedFixedValue &&
          formulas[rowIndex][columnIndex] === '';
        const isSpillCell =
          rowIndex >= 2 && rowIndex <= lastSpillRowIndex;
        const isFormulaOrigin =
          rowIndex === 2 && columnIndex === 0 &&
          formulas[rowIndex][columnIndex] === expectedFormula;
        const isSpillResult =
          isSpillCell &&
          (formulas[rowIndex][columnIndex] === '' || isFormulaOrigin);
        const isEmpty = String(value).trim() === '' &&
          formulas[rowIndex][columnIndex] === '';

        if (!isEmpty && !isKnownFixedValue && !isSpillResult) {
          const address = range
            .getCell(rowIndex + 1, columnIndex + 1)
            .getA1Notation();
          throw new Error(
            'ダッシュボード!' + address +
            'に面接予定一覧の展開を妨げる未知の値または数式があります。' +
            '既存値は削除していません。'
          );
        }
      });
    });
    return;
  }

  values.forEach(function(rowValues, rowIndex) {
    rowValues.forEach(function(value, columnIndex) {
      const actualRow = range.getRow() + rowIndex;
      const actualColumn = range.getColumn() + columnIndex;
      let expectedFixedValue = null;
      if (actualRow === 3 && actualColumn === 10) {
        expectedFixedValue = DASHBOARD_INTERVIEW_FIXED_VALUES.J3;
      } else if (actualRow === 4) {
        expectedFixedValue = interviewHeaders[columnIndex];
      }
      const isKnownFixedValue =
        expectedFixedValue !== null &&
        String(value) === expectedFixedValue &&
        formulas[rowIndex][columnIndex] === '';
      const isEmpty = String(value).trim() === '' &&
        formulas[rowIndex][columnIndex] === '';

      if (!isEmpty && !isKnownFixedValue) {
        const address = range
          .getCell(rowIndex + 1, columnIndex + 1)
          .getA1Notation();
        throw new Error(
          'ダッシュボード!' + address +
          'に未知の値または数式があるため、面接予定一覧を設定できません。'
        );
      }
    });
  });
}

function getDashboardInterviewRange(sheet, applicationLastRow) {
  return sheet.getRange(3, 10, applicationLastRow + 2, 6);
}

/**
 * 面接一覧の予約範囲と重なる既存の条件付き書式は上書きせず停止します。
 */
function validateInterviewConditionalFormatting(sheet, interviewRange) {
  const hasConflict = sheet.getConditionalFormatRules().some(function(rule) {
    return rule.getRanges().some(function(ruleRange) {
      const rowsOverlap =
        ruleRange.getRow() <= interviewRange.getLastRow() &&
        ruleRange.getLastRow() >= interviewRange.getRow();
      const columnsOverlap =
        ruleRange.getColumn() <= interviewRange.getLastColumn() &&
        ruleRange.getLastColumn() >= interviewRange.getColumn();
      return rowsOverlap && columnsOverlap;
    });
  });

  if (hasConflict) {
    throw new Error(
      'ダッシュボードJ3:Oの予約範囲に既存の条件付き書式があるため、' +
      '面接予定一覧を設定できません。'
    );
  }
}

function validateDashboardSubmissionDeadlineLayout(sheet) {
  const managedRange = sheet.getRange(
    DASHBOARD_SUBMISSION_DEADLINE_START_ROW,
    1,
    sheet.getMaxRows() - DASHBOARD_SUBMISSION_DEADLINE_START_ROW + 1,
    8
  );
  const mergedRanges = sheet.getRange('A32:H' + sheet.getMaxRows())
    .getMergedRanges();
  if (mergedRanges.length > 0) {
    throw new Error(
      'ダッシュボードの提出期限管理範囲に結合セルがあるため、' +
      '提出期限一覧を設定できません。'
    );
  }

  const fixedValues = [
    [DASHBOARD_SUBMISSION_DEADLINE_SECTION_RANGE.split(':')[0],
      DASHBOARD_SUBMISSION_DEADLINE_TITLE]
  ];
  DASHBOARD_SUBMISSION_DEADLINE_HEADERS.forEach(function(header, index) {
    fixedValues.push([
      String.fromCharCode('A'.charCodeAt(0) + index) + '34',
      header
    ]);
  });
  fixedValues.forEach(function(definition) {
    const currentValue = String(sheet.getRange(definition[0]).getDisplayValue());
    if (currentValue !== '' && currentValue !== definition[1]) {
      throw new Error(
        'ダッシュボード!' + definition[0] +
        'に未知の値「' + currentValue + '」があるため、上書きしません。'
      );
    }
  });

  const warningCell = sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_WARNING_CELL);
  const warningFormula = buildDashboardSubmissionDeadlineWarningFormula();
  if (warningCell.getFormula() !== '' && warningCell.getFormula() !== warningFormula) {
    throw new Error(
      'ダッシュボード!' + DASHBOARD_SUBMISSION_DEADLINE_WARNING_CELL +
      'に未知の数式があるため、上書きしません。'
    );
  }
  if (warningCell.getFormula() === '' &&
      String(warningCell.getDisplayValue()).trim() !== '') {
    throw new Error(
      'ダッシュボード!' + DASHBOARD_SUBMISSION_DEADLINE_WARNING_CELL +
      'に未知の値があるため、上書きしません。'
    );
  }

  const listCell = sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_FORMULA_CELL);
  const listFormula = buildDashboardSubmissionDeadlineFormula();
  if (listCell.getFormula() !== '' && listCell.getFormula() !== listFormula) {
    throw new Error('ダッシュボード!A35に未知の数式があるため、上書きしません。');
  }
  if (listCell.getFormula() === '') {
    const values = managedRange.getDisplayValues();
    const formulas = managedRange.getFormulas();
    const hasUnknownValues = values.some(function(row, rowIndex) {
      return row.some(function(value, columnIndex) {
        return String(value).trim() !== '' ||
          formulas[rowIndex][columnIndex] !== '';
      });
    });
    if (hasUnknownValues) {
      throw new Error(
        'ダッシュボードの提出期限一覧予定範囲に未知の値または数式があるため、' +
        '上書きしません。'
      );
    }
  }
}

function buildDashboardSubmissionDeadlineWarningFormula() {
  const terminalExclusion = buildTerminalStatusExclusionFormula(
    '\'応募一覧\'!G2:G'
  );
  const unsubmitted = buildSubmissionPendingFormula();
  return '=IF(SUMPRODUCT((\'応募一覧\'!A2:A<>"")*' +
    'IFERROR(\'応募一覧\'!W2:W<>"",TRUE)*' +
    'IFERROR(NOT(ISNUMBER(\'応募一覧\'!W2:W)),TRUE)*' +
    unsubmitted + '*' + terminalExclusion +
    ')=' +
    '0,"",CONCATENATE("' +
    DASHBOARD_SUBMISSION_DEADLINE_WARNING_PREFIX +
    '",SUMPRODUCT((\'応募一覧\'!A2:A<>"")*' +
    'IFERROR(\'応募一覧\'!W2:W<>"",TRUE)*' +
    'IFERROR(NOT(ISNUMBER(\'応募一覧\'!W2:W)),TRUE)*' +
    unsubmitted + '*' + terminalExclusion +
    '),"' + DASHBOARD_SUBMISSION_DEADLINE_WARNING_SUFFIX + '"))';
}

function buildSubmissionPendingFormula() {
  return '((\'応募一覧\'!S2:S="未提出")+' +
    '(\'応募一覧\'!T2:T="未提出")+' +
    '(\'応募一覧\'!U2:U="未提出")+' +
    '(\'応募一覧\'!V2:V="未提出")>0)';
}

function buildSubmissionItemsFormula() {
  return 'REGEXREPLACE(' +
    'IF(\'応募一覧\'!S2:S="未提出","履歴書、","")&' +
    'IF(\'応募一覧\'!T2:T="未提出","職務経歴書、","")&' +
    'IF(\'応募一覧\'!U2:U="未提出","ポートフォリオ、","")&' +
    'IF(\'応募一覧\'!V2:V="未提出","その他提出物","")' +
    ',"、$","")';
}

function buildDashboardSubmissionDeadlineFormula() {
  const terminalExclusion = buildTerminalStatusExclusionFormula(
    '\'応募一覧\'!G2:G'
  );
  const unsubmitted = buildSubmissionPendingFormula();
  const validDeadline = 'IFERROR(ISNUMBER(\'応募一覧\'!W2:W),FALSE)';
  const nearDeadline = 'IFERROR(\'応募一覧\'!W2:W<=TODAY()+7,FALSE)';
  const statusText =
    'IFERROR(IF(\'応募一覧\'!W2:W<TODAY(),' +
      '(TODAY()-INT(\'応募一覧\'!W2:W))&"日超過",' +
      'IF(INT(\'応募一覧\'!W2:W)=TODAY(),"今日",' +
        '"残り"&(INT(\'応募一覧\'!W2:W)-TODAY())&"日")),"")';
  const caution = 'IF(\'応募一覧\'!X2:X<>"","' +
    DASHBOARD_SUBMISSION_DEADLINE_REVIEW_MESSAGE + '","")';
  const source = '{' +
    'IFERROR(\'応募一覧\'!W2:W,""),' +
    statusText + ',' +
    '\'応募一覧\'!B2:B,' +
    '\'応募一覧\'!C2:C,' +
    buildSubmissionItemsFormula() + ',' +
    caution + ',' +
    '\'応募一覧\'!G2:G,' +
    '\'応募一覧\'!D2:D,' +
    '\'応募一覧\'!A2:A' +
    '}';
  const conditions =
    '(\'応募一覧\'!A2:A<>"")*' + validDeadline + '*' +
    unsubmitted + '*' + terminalExclusion + '*' + nearDeadline;
  return '=IFERROR(CHOOSECOLS(SORT(FILTER(' + source + ',' + conditions +
    '),1,TRUE,3,TRUE,9,TRUE),1,2,3,4,5,6,7,8),"' +
    DASHBOARD_SUBMISSION_DEADLINE_NO_ITEMS_MESSAGE + '")';
}

function applyDashboardValuesAndFormulas(sheet, applicationLastRow) {
  setCellIfEmpty(sheet.getRange(DASHBOARD_VALUES.TITLE_CELL), DASHBOARD_VALUES.TITLE);

  Object.keys(DASHBOARD_FIXED_VALUES).forEach(function(address) {
    const range = sheet.getRange(address);
    const currentValue = String(range.getDisplayValue());
    if (address === DASHBOARD_VALUES.MESSAGE_CELL ||
        currentValue === '' ||
        currentValue === DASHBOARD_FIXED_VALUES[address]) {
      range.setValue(DASHBOARD_FIXED_VALUES[address]);
    }
  });

  Object.keys(DASHBOARD_INTERVIEW_FIXED_VALUES).forEach(function(address) {
    const range = sheet.getRange(address);
    const currentValue = String(range.getDisplayValue());
    if (currentValue === '' ||
        currentValue === DASHBOARD_INTERVIEW_FIXED_VALUES[address]) {
      range.setValue(DASHBOARD_INTERVIEW_FIXED_VALUES[address]);
    }
  });

  getDashboardFormulaDefinitions(applicationLastRow).forEach(function(definition) {
    const range = sheet.getRange(definition.address);
    if (range.getFormula() !== definition.formula) {
      range.setFormula(definition.formula);
    }
  });

  const warningCell = sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_WARNING_CELL);
  if (warningCell.getFormula() !== buildDashboardSubmissionDeadlineWarningFormula()) {
    warningCell.setFormula(buildDashboardSubmissionDeadlineWarningFormula());
  }
  const submissionTitle = sheet.getRange('A33');
  if (String(submissionTitle.getDisplayValue()).trim() === '') {
    submissionTitle.setValue(DASHBOARD_SUBMISSION_DEADLINE_TITLE);
  }
  DASHBOARD_SUBMISSION_DEADLINE_HEADERS.forEach(function(header, index) {
    const cell = sheet.getRange(34, index + 1);
    if (String(cell.getDisplayValue()).trim() === '') {
      cell.setValue(header);
    }
  });
  const listCell = sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_FORMULA_CELL);
  if (listCell.getFormula() !== buildDashboardSubmissionDeadlineFormula()) {
    listCell.setFormula(buildDashboardSubmissionDeadlineFormula());
  }
}

function getDashboardFormulaDefinitions(applicationLastRow) {
  return [
    { address: 'B4', formula: DASHBOARD_TOTAL_FORMULA },
    { address: 'D4', formula: buildInProgressDashboardFormula() },
    { address: 'F4', formula: DASHBOARD_NEEDS_REVIEW_FORMULA },
    {
      address: 'H4',
      formula: buildDashboardUpcomingInterviewsFormula(applicationLastRow)
    },
    { address: 'B5', formula: DASHBOARD_COMPANY_WAITING_FORMULA },
    { address: 'D5', formula: DASHBOARD_SELF_ACTION_FORMULA },
    { address: 'B6', formula: DASHBOARD_HIRED_FORMULA },
    { address: 'D6', formula: DASHBOARD_REJECTED_FORMULA },
    { address: 'F6', formula: DASHBOARD_WITHDRAWN_FORMULA },
    { address: 'A10', formula: DASHBOARD_STATUS_NAMES_FORMULA },
    { address: 'B10', formula: DASHBOARD_STATUS_COUNTS_FORMULA },
    {
      address: DASHBOARD_INTERVIEW_FORMULA_CELL,
      formula: buildDashboardInterviewListFormula(applicationLastRow)
    }
  ];
}

function getStage8DashboardFormulaDefinitions(applicationLastRow) {
  return [
    { address: 'B4', formula: DASHBOARD_TOTAL_FORMULA },
    { address: 'D4', formula: buildInProgressDashboardFormula() },
    { address: 'F4', formula: DASHBOARD_NEEDS_REVIEW_FORMULA },
    {
      address: 'H4',
      formula: PREVIOUS_DASHBOARD_UPCOMING_INTERVIEWS_FORMULA
    },
    { address: 'B6', formula: DASHBOARD_HIRED_FORMULA },
    { address: 'D6', formula: DASHBOARD_REJECTED_FORMULA },
    { address: 'F6', formula: DASHBOARD_WITHDRAWN_FORMULA },
    { address: 'A10', formula: DASHBOARD_STATUS_NAMES_FORMULA },
    { address: 'B10', formula: PREVIOUS_DASHBOARD_STATUS_COUNTS_FORMULA }
  ];
}

function buildInProgressDashboardFormula() {
  const countParts = IN_PROGRESS_SELECTION_STATUSES.map(function(status) {
    return 'COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,"' +
      status + '")';
  });
  return '=' + countParts.join('+');
}

function applyDashboardFormatting(sheet) {
  sheet.getRange(DASHBOARD_VALUES.TITLE_CELL)
    .setFontWeight('bold')
    .setFontSize(DASHBOARD_TITLE_FONT_SIZE)
    .setFontColor(DASHBOARD_TITLE_COLOR);

  sheet.getRange('A3:H3')
    .setBackground(DASHBOARD_SECTION_BACKGROUND_COLOR)
    .setFontWeight('bold');
  sheet.getRange('A8:H8')
    .setBackground(DASHBOARD_SECTION_BACKGROUND_COLOR)
    .setFontWeight('bold');
  sheet.getRange('A9:B9')
    .setBackground(HEADER_BACKGROUND_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  ['A4', 'C4', 'E4', 'G4', 'A5', 'C5', 'A6', 'C6', 'E6'].forEach(function(address) {
    sheet.getRange(address)
      .setBackground(DASHBOARD_LABEL_BACKGROUND_COLOR)
      .setFontWeight('bold');
  });
  ['B4', 'D4', 'F4', 'H4', 'B5', 'D5', 'B6', 'D6', 'F6'].forEach(function(address) {
    sheet.getRange(address)
      .setBackground(DASHBOARD_VALUE_BACKGROUND_COLOR)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setNumberFormat(INTEGER_FORMAT);
  });

  sheet.getRange('A10:B30')
    .setVerticalAlignment('middle')
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      DASHBOARD_BORDER_COLOR,
      SpreadsheetApp.BorderStyle.SOLID
    );
  sheet.getRange('B10:B30')
    .setHorizontalAlignment('center')
    .setNumberFormat(INTEGER_FORMAT);

  sheet.getRange(DASHBOARD_INTERVIEW_TITLE_RANGE)
    .setBackground(DASHBOARD_SECTION_BACKGROUND_COLOR)
    .setFontWeight('bold');
  sheet.getRange(DASHBOARD_INTERVIEW_HEADER_RANGE)
    .setBackground(HEADER_BACKGROUND_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      DASHBOARD_BORDER_COLOR,
      SpreadsheetApp.BorderStyle.SOLID
    );
  sheet.getRange('J5:O' + sheet.getMaxRows())
    .setVerticalAlignment('middle');
  sheet.getRange('J5:J' + sheet.getMaxRows())
    .setNumberFormat(DATE_TIME_FORMAT);
  sheet.getRange('M5:M' + sheet.getMaxRows())
    .setNumberFormat(INTEGER_FORMAT)
    .setHorizontalAlignment('center');
  sheet.getRange('K5:L' + sheet.getMaxRows()).setWrap(true);
  sheet.getRange('N5:N' + sheet.getMaxRows()).setWrap(true);

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 90);
  sheet.setColumnWidth(7, 160);
  sheet.setColumnWidth(8, 90);
  sheet.setColumnWidth(10, 150);
  sheet.setColumnWidth(11, 180);
  sheet.setColumnWidth(12, 180);
  sheet.setColumnWidth(13, 130);
  sheet.setColumnWidth(14, 250);
  sheet.setColumnWidth(15, 260);
}

function applyDashboardSubmissionDeadlineFormatting(sheet) {
  sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_WARNING_CELL)
    .setFontColor('#C00000')
    .setFontWeight('bold');
  sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_SECTION_RANGE)
    .setBackground(DASHBOARD_SECTION_BACKGROUND_COLOR)
    .setFontWeight('bold');
  sheet.getRange(DASHBOARD_SUBMISSION_DEADLINE_HEADER_RANGE)
    .setBackground(HEADER_BACKGROUND_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange('A35:H' + sheet.getMaxRows())
    .setVerticalAlignment('middle');
  sheet.getRange('A35:A' + sheet.getMaxRows())
    .setNumberFormat(DATE_FORMAT);
  sheet.getRange('B35:B' + sheet.getMaxRows())
    .setHorizontalAlignment('center');
  sheet.getRange('A34:H' + sheet.getMaxRows()).setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    DASHBOARD_SUBMISSION_DEADLINE_BORDER_COLOR,
    SpreadsheetApp.BorderStyle.SOLID
  );
  sheet.getRange('B35:H' + sheet.getMaxRows()).setWrap(true);
}

function setupAnalysisCharts(sheet, applicationLastRow) {
  validateAnalysisChartPlacementAreas(sheet);
  const charts = sheet.getCharts();
  if (charts.length === 0) {
    createAnalysisMonthlyChart(sheet);
    createAnalysisSiteChart(sheet, applicationLastRow);
    return;
  }

  if (charts.length !== ANALYSIS_CHART_COUNT) {
    throw new Error(
      '分析シートのグラフ数がJobFlow正式仕様と一致しないため、' +
      'グラフを変更せず停止します。'
    );
  }

  const monthlyChart = findExpectedAnalysisChart(
    charts,
    ANALYSIS_MONTHLY_CHART_TITLE,
    ANALYSIS_MONTHLY_CHART_TYPE,
    ANALYSIS_MONTHLY_CHART_ANCHOR,
    'monthly'
  );
  const siteChart = findExpectedAnalysisChart(
    charts,
    ANALYSIS_SITE_CHART_TITLE,
    ANALYSIS_SITE_CHART_TYPE,
    ANALYSIS_SITE_CHART_ANCHOR,
    'site'
  );
  if (!monthlyChart || !siteChart || monthlyChart === siteChart) {
    throw new Error(
      '分析シートの既存グラフをJobFlow管理グラフとして安全に識別できないため、' +
      '変更せず停止します。'
    );
  }

  const expectedSiteRange = getAnalysisSiteChartRange(sheet, applicationLastRow);
  normalizeAnalysisChart(
    sheet,
    monthlyChart,
    Charts.ChartType.COLUMN,
    sheet.getRange('A4:B10'),
    ANALYSIS_MONTHLY_CHART_TITLE,
    ANALYSIS_MONTHLY_CHART_ANCHOR
  );
  normalizeAnalysisChart(
    sheet,
    siteChart,
    Charts.ChartType.BAR,
    expectedSiteRange,
    ANALYSIS_SITE_CHART_TITLE,
    ANALYSIS_SITE_CHART_ANCHOR
  );
}

function validateAnalysisChartPlacementAreas(sheet) {
  const reservedRanges = [
    sheet.getRange('F3:M17'),
    sheet.getRange('F20:M36')
  ];
  const mergedRanges = reservedRanges.reduce(function(allRanges, range) {
    return allRanges.concat(range.getMergedRanges());
  }, []);
  if (mergedRanges.length > 0) {
    throw new Error(
      '分析シートのグラフ配置予約範囲に結合セルがあるため、' +
      'グラフを設定できません。'
    );
  }
}

function createAnalysisMonthlyChart(sheet) {
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange('A4:B10'))
    .setNumHeaders(1)
    .setPosition(
      ANALYSIS_MONTHLY_CHART_ANCHOR.row,
      ANALYSIS_MONTHLY_CHART_ANCHOR.column,
      0,
      0
    )
    .setOption('title', ANALYSIS_MONTHLY_CHART_TITLE)
    .setOption('width', ANALYSIS_CHART_WIDTH)
    .setOption('height', ANALYSIS_CHART_HEIGHT)
    .setOption('legend', { position: 'none' })
    .build();
  sheet.insertChart(chart);
}

function createAnalysisSiteChart(sheet, applicationLastRow) {
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(getAnalysisSiteChartRange(sheet, applicationLastRow))
    .setNumHeaders(1)
    .setPosition(
      ANALYSIS_SITE_CHART_ANCHOR.row,
      ANALYSIS_SITE_CHART_ANCHOR.column,
      0,
      0
    )
    .setOption('title', ANALYSIS_SITE_CHART_TITLE)
    .setOption('width', ANALYSIS_CHART_WIDTH)
    .setOption('height', ANALYSIS_CHART_HEIGHT)
    .setOption('legend', { position: 'none' })
    .build();
  sheet.insertChart(chart);
}

function getAnalysisSiteChartRange(sheet, applicationLastRow) {
  return sheet.getRange(14, 1, applicationLastRow, 2);
}

function findExpectedAnalysisChart(charts, title, chartTypeName, anchor, role) {
  const matches = charts.filter(function(chart) {
    const container = chart.getContainerInfo();
    const options = chart.getOptions();
    const titleOption = options ? options.get('title') : null;
    const chartTitle = titleOption === null || titleOption === undefined
      ? '' : String(titleOption);
    const ranges = chart.getRanges();
    const hasExpectedRoleRange = role === 'monthly'
      ? isExactAnalysisChartRange(ranges, 'A4:B10')
      : isKnownAnalysisSiteChartRange(ranges);
    return container.getAnchorRow() === anchor.row &&
      container.getAnchorColumn() === anchor.column &&
      container.getOffsetX() === 0 &&
      container.getOffsetY() === 0 &&
      chartTitle === title &&
      hasExpectedRoleRange;
  });
  return matches.length === 1 ? matches[0] : null;
}

function isExactAnalysisChartRange(ranges, a1Notation) {
  return ranges.length === 1 &&
    ranges[0].getSheet().getName() === SHEET_NAMES.ANALYSIS &&
    ranges[0].getA1Notation() === a1Notation;
}

function isKnownAnalysisSiteChartRange(ranges) {
  if (ranges.length !== 1 ||
      ranges[0].getSheet().getName() !== SHEET_NAMES.ANALYSIS ||
      ranges[0].getColumn() !== 1 ||
      ranges[0].getNumColumns() !== 2 ||
      ranges[0].getRow() !== 14) {
    return false;
  }
  const rangeEndRow = ranges[0].getLastRow();
  const currentLastRow = getAnalysisSiteChartRange(
    ranges[0].getSheet(),
    ranges[0].getSheet().getParent().getSheetByName(SHEET_NAMES.APPLICATIONS).getMaxRows()
  ).getLastRow();
  return rangeEndRow === currentLastRow;
}

function hasExactAnalysisChartRange(chart, expectedRange) {
  const ranges = chart.getRanges();
  return ranges.length === 1 &&
    ranges[0].getSheet().getSheetId() === expectedRange.getSheet().getSheetId() &&
    ranges[0].getA1Notation() === expectedRange.getA1Notation();
}

function normalizeAnalysisChart(sheet, chart, chartType, expectedRange,
    title, anchor) {
  const modifiedChart = chart.modify()
    .clearRanges()
    .addRange(expectedRange)
    .setNumHeaders(1)
    .setChartType(chartType)
    .setOption('title', title)
    .setPosition(anchor.row, anchor.column, 0, 0)
    .setOption('width', ANALYSIS_CHART_WIDTH)
    .setOption('height', ANALYSIS_CHART_HEIGHT)
    .build();
  sheet.updateChart(modifiedChart);
}

function setCellIfEmpty(range, value) {
  if (isEmptyCellValue(range.getValue())) {
    range.setValue(value);
  }
}

function ensureSheetSize(sheet, requiredRows, requiredColumns) {
  const missingRows = requiredRows - sheet.getMaxRows();
  const missingColumns = requiredColumns - sheet.getMaxColumns();

  if (missingRows > 0) {
    sheet.insertRowsAfter(sheet.getMaxRows(), missingRows);
  }
  if (missingColumns > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), missingColumns);
  }
}

function isEmptyCellValue(value) {
  return value === '' || value === null;
}

function formatWarningValue(value) {
  const text = String(value);
  const maximumLength = 80;
  return text.length > maximumLength
    ? text.substring(0, maximumLength) + '…'
    : text;
}
