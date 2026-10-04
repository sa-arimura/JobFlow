/**
 * JobFlowの設定値を一か所へまとめます。
 * 定数化すると、列や色を変更するときに各処理を探して直す必要がなくなります。
 */
const SHEET_NAMES = Object.freeze({
  APPLICATIONS: '応募一覧',
  DASHBOARD: 'ダッシュボード',
  ANALYSIS: '分析',
  SETTINGS: '設定'
});

const APPLICATION_HEADERS = Object.freeze([
  '応募ID',
  '会社名',
  '職種',
  '求人URL',
  '応募日',
  '応募サイト',
  '選考状況',
  '最終返信日',
  '次回予定日',
  '面接日時',
  'メモ',
  '応募経過日数',
  '返信待ち日数',
  '要確認',
  '作成日時',
  '更新日時',
  '担当者',
  '面接までの日数',
  '履歴書',
  '職務経歴書',
  'ポートフォリオ',
  'その他提出物',
  '提出期限',
  '提出日',
  '提出物メモ',
  '対応待ち',
  '待機開始日'
]);

const APPLICATION_COLUMNS = Object.freeze({
  APPLICATION_ID: 1,
  COMPANY_NAME: 2,
  JOB_TITLE: 3,
  JOB_URL: 4,
  APPLICATION_DATE: 5,
  APPLICATION_SITE: 6,
  SELECTION_STATUS: 7,
  LAST_REPLY_DATE: 8,
  NEXT_SCHEDULE_DATE: 9,
  INTERVIEW_DATE_TIME: 10,
  MEMO: 11,
  ELAPSED_DAYS: 12,
  WAITING_DAYS: 13,
  NEEDS_REVIEW: 14,
  CREATED_DATE_TIME: 15,
  UPDATED_DATE_TIME: 16,
  CONTACT_PERSON: 17,
  DAYS_UNTIL_INTERVIEW: 18,
  RESUME_STATUS: 19,
  WORK_HISTORY_STATUS: 20,
  PORTFOLIO_STATUS: 21,
  OTHER_SUBMISSION_STATUS: 22,
  SUBMISSION_DEADLINE: 23,
  SUBMISSION_DATE: 24,
  SUBMISSION_MEMO: 25,
  WAITING_STATUS: 26,
  WAITING_START_DATE: 27
});

const WAITING_STATUS_VALUES = Object.freeze({
  NONE: 'なし',
  COMPANY: '企業待ち',
  SELF: '自分対応'
});

const WAITING_STATUSES = Object.freeze([
  WAITING_STATUS_VALUES.NONE,
  WAITING_STATUS_VALUES.COMPANY,
  WAITING_STATUS_VALUES.SELF
]);

const WAITING_DEFAULT_STATUS = WAITING_STATUS_VALUES.NONE;

const SUBMISSION_STATUSES = Object.freeze(['不要', '未提出', '提出済み']);

const SUBMISSION_DEFAULTS = Object.freeze({
  RESUME_STATUS: '未提出',
  WORK_HISTORY_STATUS: '未提出',
  PORTFOLIO_STATUS: '不要',
  OTHER_SUBMISSION_STATUS: '不要'
});

const SUBMISSION_BACKGROUND_COLORS = Object.freeze({
  '不要': '#E7E6E6',
  '未提出': '#FCE4D6',
  '提出済み': '#E2F0D9'
});

const SELECTION_STATUSES = Object.freeze([
  '応募準備中',
  '応募済み',
  '書類選考中',
  '書類通過',
  '面接予定',
  '面接済み',
  '採用',
  '不採用',
  '辞退'
]);

const TERMINAL_SELECTION_STATUSES = Object.freeze([
  '採用',
  '不採用',
  '辞退'
]);

const IN_PROGRESS_SELECTION_STATUSES = Object.freeze([
  '応募済み',
  '書類選考中',
  '書類通過',
  '面接予定',
  '面接済み'
]);

const STATUS_BACKGROUND_COLORS = Object.freeze({
  '応募準備中': '#E0E0E0',
  '応募済み': '#FFF2CC',
  '書類選考中': '#FFE599',
  '書類通過': '#D9EAD3',
  '面接予定': '#CFE2F3',
  '面接済み': '#D9D2E9',
  '採用': '#B6D7A8',
  '不採用': '#F4CCCC',
  '辞退': '#CCCCCC'
});

const STATUS_TEXT_COLORS = Object.freeze({
  '応募準備中': '#000000',
  '応募済み': '#000000',
  '書類選考中': '#000000',
  '書類通過': '#000000',
  '面接予定': '#000000',
  '面接済み': '#000000',
  '採用': '#000000',
  '不採用': '#000000',
  '辞退': '#000000'
});

const APPLICATION_COLUMN_WIDTHS = Object.freeze([
  150,
  180,
  180,
  250,
  100,
  130,
  130,
  110,
  110,
  150,
  300,
  120,
  120,
  120,
  150,
  150,
  160,
  130,
  120,
  140,
  140,
  140,
  110,
  110,
  260,
  120,
  110
]);

const DATE_FORMAT = 'yyyy/MM/dd';
const DATE_TIME_FORMAT = 'yyyy/MM/dd HH:mm';
const INTEGER_FORMAT = '0';
const APPLICATION_HEADER_ROW = 1;
const DATA_START_ROW = 2;
const DATA_END_ROW = 1000;
const APPLICATION_COLUMN_COUNT = APPLICATION_HEADERS.length;
const HEADER_BACKGROUND_COLOR = '#1F4E78';
const HEADER_TEXT_COLOR = '#FFFFFF';
const SETTINGS_HEADER_BACKGROUND_COLOR = '#D9EAF7';
const DASHBOARD_TITLE_COLOR = '#1F4E78';
const DASHBOARD_TITLE_FONT_SIZE = 16;
const STATUS_HELP_TEXT = '一覧から選考状況を選択してください。';
const SUBMISSION_STATUS_HELP_TEXT = '不要・未提出・提出済みから選択してください。';
const WAITING_STATUS_HELP_TEXT = 'なし・企業待ち・自分対応から選択してください。';
const APPLICATION_FORM_TITLE = 'JobFlow 応募情報';
const APPLICATION_SIDEBAR_WIDTH = 420;
const APPLICATION_ID_COLUMN = APPLICATION_COLUMNS.APPLICATION_ID;
const APPLICATION_WRITE_COLUMN_COUNT = APPLICATION_COLUMNS.SELECTION_STATUS;
const LOCK_WAIT_MILLISECONDS = 30000;
const NEEDS_REVIEW_TEXT = '要確認';
const NEEDS_REVIEW_BACKGROUND_COLOR = '#F4CCCC';
const ELAPSED_DAYS_FORMULA =
  '=ARRAYFORMULA(IF(E2:E="","",IFERROR(IF(INT(E2:E)>TODAY(),"",TODAY()-INT(E2:E)),"")))';
const PREVIOUS_WAITING_DAYS_FORMULA =
  '=ARRAYFORMULA(IF((E2:E="")+(H2:H<>""),"",IFERROR(IF(INT(E2:E)>TODAY(),"",TODAY()-INT(E2:E)),"")))';
const PREVIOUS_NEEDS_REVIEW_FORMULA =
  '=ARRAYFORMULA(IF(OR(NOT(ISNUMBER(\'設定\'!$B$2)),\'設定\'!$B$2<=0),"",' +
  'IF((E2:E<>"")*(H2:H="")*(M2:M>=\'設定\'!$B$2)*' +
  '(G2:G<>"採用")*(G2:G<>"不採用")*(G2:G<>"辞退"),"要確認","")))';
const PREVIOUS_NEEDS_REVIEW_FORMULA_V2 =
  '=ARRAYFORMULA(IF(OR(NOT(ISNUMBER(\'設定\'!$B$2)),\'設定\'!$B$2<=0),"",' +
  'IF((E2:E<>"")*(H2:H="")*ISNUMBER(M2:M)*(M2:M>=\'設定\'!$B$2)*' +
  '(G2:G<>"採用")*(G2:G<>"不採用")*(G2:G<>"辞退"),"要確認","")))';
const WAITING_DAYS_FORMULA = buildWaitingDaysFormula();
const NEEDS_REVIEW_FORMULA = buildNeedsReviewFormula();
const DAYS_UNTIL_INTERVIEW_FORMULA =
  '=ARRAYFORMULA(IF(J2:J="","",IFERROR(INT(J2:J)-TODAY(),"")))';

const BASIC_SETTINGS_VALUES = Object.freeze([
  ['設定項目', '設定値', '説明'],
  ['要確認判定日数', 7, '企業待ちの開始から何日以上経過した場合に確認対象とするか'],
  ['タイムゾーン', 'Asia/Tokyo', '日付と時刻の計算に使用するタイムゾーン'],
  ['通知機能', '無効', '将来実装する通知機能の有効・無効'],
  ['通知先メールアドレス', '', '将来Gmail通知を送信するメールアドレス']
]);

const PREVIOUS_NEEDS_REVIEW_SETTING = Object.freeze({
  LABEL: '返信なし判定日数',
  VALUE: 5,
  DESCRIPTION: '応募または連絡から何日以上返信がない場合に確認対象とするか'
});

const SETTINGS_STATUS_START_ROW = 2;
const SETTINGS_STATUS_END_ROW = 10;
const SETTINGS_STATUS_COLUMN = 5;
const SETTINGS_COLUMN_WIDTHS = Object.freeze([
  180,
  180,
  420,
  30,
  140,
  100,
  100,
  80
]);

const DASHBOARD_VALUES = Object.freeze({
  TITLE_CELL: 'A1',
  TITLE: 'JobFlow ダッシュボード',
  MESSAGE_CELL: 'A3',
  MESSAGE: '集計機能は今後のステップで実装します。',
  BASIC_SECTION_TITLE: '基本集計',
  STATUS_SECTION_TITLE: '選考状況別件数'
});

const DASHBOARD_TARGET_RANGE = 'A3:H30';
const DASHBOARD_STATUS_LIST_RANGE = 'A10:B30';
const DASHBOARD_INTERVIEW_TITLE_RANGE = 'J3:O3';
const DASHBOARD_INTERVIEW_HEADER_RANGE = 'J4:O4';
const DASHBOARD_INTERVIEW_FORMULA_CELL = 'J5';
const DASHBOARD_SECTION_BACKGROUND_COLOR = '#D9EAF7';
const DASHBOARD_LABEL_BACKGROUND_COLOR = '#EAF2F8';
const DASHBOARD_VALUE_BACKGROUND_COLOR = '#FFFFFF';
const DASHBOARD_BORDER_COLOR = '#B7C9D6';

const DASHBOARD_FIXED_VALUES = Object.freeze({
  A3: '基本集計',
  A4: '応募総数',
  C4: '選考中',
  E4: '要確認',
  G4: '今後の面接',
  A5: '企業待ち',
  C5: '自分対応',
  A6: '採用',
  C6: '不採用',
  E6: '辞退',
  A8: '選考状況別件数',
  A9: '選考状況',
  B9: '件数'
});

// Stage8完成時点のDashboard固定値。Stage9追加項目(A5/C5)は含めません。
const STAGE8_DASHBOARD_FIXED_VALUES = Object.freeze({
  A3: '基本集計',
  A4: '応募総数',
  C4: '選考中',
  E4: '要確認',
  G4: '今後の面接',
  A6: '採用',
  C6: '不採用',
  E6: '辞退',
  A8: '選考状況別件数',
  A9: '選考状況',
  B9: '件数'
});

const DASHBOARD_TOTAL_FORMULA =
  '=COUNTIF(\'応募一覧\'!A2:A,"<>")';
const DASHBOARD_COMPANY_WAITING_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!Z2:Z,"企業待ち")';
const DASHBOARD_SELF_ACTION_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!Z2:Z,"自分対応")';

const ANALYSIS_VALUES = Object.freeze({
  TITLE_CELL: 'A1',
  TITLE: 'JobFlow 分析',
  MONTH_SECTION_RANGE: 'A3:D3',
  MONTH_SECTION_TITLE: '月別応募数',
  MONTH_HEADER: '月',
  MONTH_COUNT_HEADER: '応募数',
  SITE_SECTION_RANGE: 'A13:D13',
  SITE_SECTION_TITLE: '応募サイト別応募数',
  SITE_HEADER: '応募サイト',
  SITE_COUNT_HEADER: '応募数'
});
const ANALYSIS_MONTH_LABEL_RANGE = 'A5:A10';
const ANALYSIS_MONTH_COUNT_RANGE = 'B5:B10';
const ANALYSIS_SITE_FORMULA_CELL = 'A15';
const ANALYSIS_INITIAL_ROWS = 36;
const ANALYSIS_INITIAL_COLUMNS = 13;
const ANALYSIS_SECTION_BACKGROUND_COLOR = '#D9EAF7';
const ANALYSIS_BORDER_COLOR = '#B7C9D6';
const DASHBOARD_SUBMISSION_DEADLINE_WARNING_CELL = 'A32';
const DASHBOARD_SUBMISSION_DEADLINE_SECTION_RANGE = 'A33:H33';
const DASHBOARD_SUBMISSION_DEADLINE_HEADER_RANGE = 'A34:H34';
const DASHBOARD_SUBMISSION_DEADLINE_FORMULA_CELL = 'A35';
const DASHBOARD_SUBMISSION_DEADLINE_START_ROW = 35;
const DASHBOARD_SUBMISSION_DEADLINE_TITLE = '提出期限管理';
const DASHBOARD_SUBMISSION_DEADLINE_HEADERS = Object.freeze([
  '提出期限',
  '残り／超過',
  '会社名',
  '職種',
  '未提出提出物',
  '注意',
  '選考状況',
  '求人URL'
]);
const DASHBOARD_SUBMISSION_DEADLINE_NO_ITEMS_MESSAGE =
  '提出期限が近い未提出案件はありません。';
const DASHBOARD_SUBMISSION_DEADLINE_REVIEW_MESSAGE = '状態要確認';
const DASHBOARD_SUBMISSION_DEADLINE_WARNING_PREFIX =
  '提出期限を確認できない案件が';
const DASHBOARD_SUBMISSION_DEADLINE_WARNING_SUFFIX = '件あります。';
const DASHBOARD_SUBMISSION_DEADLINE_BORDER_COLOR = '#B7C9D6';
const ANALYSIS_CHART_COUNT = 2;
const ANALYSIS_MONTHLY_CHART_TITLE = '月別応募数';
const ANALYSIS_SITE_CHART_TITLE = '応募サイト別応募数';
const ANALYSIS_MONTHLY_CHART_TYPE = 'COLUMN';
const ANALYSIS_SITE_CHART_TYPE = 'BAR';
const ANALYSIS_MONTHLY_CHART_ANCHOR = Object.freeze({
  row: 3,
  column: 6
});
const ANALYSIS_SITE_CHART_ANCHOR = Object.freeze({
  row: 20,
  column: 6
});
const ANALYSIS_CHART_WIDTH = 760;
const ANALYSIS_CHART_HEIGHT = 300;
const DASHBOARD_NEEDS_REVIEW_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!N2:N,"要確認")';
const PREVIOUS_DASHBOARD_UPCOMING_INTERVIEWS_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,"面接予定",' +
  '\'応募一覧\'!J2:J,">="&TODAY())';
const DASHBOARD_HIRED_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,"採用")';
const DASHBOARD_REJECTED_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,"不採用")';
const DASHBOARD_WITHDRAWN_FORMULA =
  '=COUNTIFS(\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,"辞退")';
const DASHBOARD_STATUS_NAMES_FORMULA =
  '=FILTER(\'設定\'!E2:E,\'設定\'!E2:E<>"")';
const PREVIOUS_DASHBOARD_STATUS_COUNTS_FORMULA =
  '=ARRAYFORMULA(IF(A10:A="","",COUNTIFS(' +
  '\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,A10:A)))';
const DASHBOARD_STATUS_COUNTS_FORMULA =
  '=ARRAYFORMULA(IF(A10:A30="","",COUNTIFS(' +
  '\'応募一覧\'!A2:A,"<>",\'応募一覧\'!G2:G,A10:A30)))';

function buildTerminalStatusExclusionFormula(statusRange) {
  return TERMINAL_SELECTION_STATUSES.map(function(status) {
    return '(' + statusRange + '<>"' + status + '")';
  }).join('*');
}

function buildWaitingDaysFormula() {
  return '=ARRAYFORMULA(IF((A2:A<>"")*(Z2:Z="' +
    WAITING_STATUS_VALUES.COMPANY + '")*ISNUMBER(AA2:AA)*' +
    'IFERROR(INT(AA2:AA)<=TODAY(),FALSE)*' +
    buildTerminalStatusExclusionFormula('G2:G') + ',' +
    'IFERROR(TODAY()-INT(AA2:AA),""),""))';
}

function buildNeedsReviewFormula() {
  return '=ARRAYFORMULA(IF(OR(NOT(ISNUMBER(\'設定\'!$B$2)),' +
    '\'設定\'!$B$2<=0),"",IF((A2:A<>"")*(Z2:Z="' +
    WAITING_STATUS_VALUES.COMPANY + '")*ISNUMBER(M2:M)*' +
    '(M2:M>=\'設定\'!$B$2)*' +
    buildTerminalStatusExclusionFormula('G2:G') +
    ',"要確認","")))';
}

const DASHBOARD_INTERVIEW_FIXED_VALUES = Object.freeze({
  J3: '今後の面接予定',
  J4: '面接日時',
  K4: '会社名',
  L4: '職種',
  M4: '面接までの日数',
  N4: '求人URL',
  O4: '応募ID'
});

/**
 * 応募一覧の実際の最大行までを対象にし、固定の小さな件数で切り捨てません。
 */
function buildDashboardUpcomingInterviewsFormula(lastRow) {
  return '=SUMPRODUCT((' +
    '\'応募一覧\'!A2:A' + lastRow + '<>"")*(' +
    '\'応募一覧\'!G2:G' + lastRow + '="面接予定")*' +
    'ISNUMBER(\'応募一覧\'!J2:J' + lastRow + ')*(' +
    'IFERROR(INT(\'応募一覧\'!J2:J' + lastRow + '),0)>=TODAY()))';
}

/**
 * FILTERで対象行を抽出し、SORTで日時、会社名、応募IDの順に並べます。
 * 対象が0件のときは、IFNAで分かりやすい案内文へ置き換えます。
 */
function buildDashboardInterviewListFormula(lastRow) {
  const applicationRange = '\'応募一覧\'!';
  const idRange = applicationRange + 'A2:A' + lastRow;
  const companyRange = applicationRange + 'B2:B' + lastRow;
  const jobTitleRange = applicationRange + 'C2:C' + lastRow;
  const urlRange = applicationRange + 'D2:D' + lastRow;
  const statusRange = applicationRange + 'G2:G' + lastRow;
  const interviewRange = applicationRange + 'J2:J' + lastRow;
  const daysRange = applicationRange + 'R2:R' + lastRow;

  return '=IFNA(SORT(FILTER({' +
    interviewRange + ',' + companyRange + ',' + jobTitleRange + ',' +
    daysRange + ',' + urlRange + ',' + idRange + '},' +
    idRange + '<>"",' +
    statusRange + '="面接予定",' +
    'ISNUMBER(' + interviewRange + '),' +
    'IFERROR(INT(' + interviewRange + '),0)>=TODAY()),' +
    '1,TRUE,2,TRUE,6,TRUE),' +
    '{"今後の面接予定はありません。","","","","",""})';
}

const COMPLETED_SETUP_ITEMS = Object.freeze([
  '必要なシートの確認・作成',
  '応募一覧の見出しと基本書式',
  '選考状況のプルダウン',
  '選考状況の条件付き書式',
  '日数計算と返信なし判定の数式',
  '要確認セルの条件付き書式',
  '提出物管理列・入力規則・条件付き書式',
  '対応待ち管理列・入力規則・待機開始日の自動管理',
  '設定シートの初期値',
  'ダッシュボードの基本集計',
  '今後の面接予定一覧'
]);
