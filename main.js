/**
 * スプレッドシートのconfigシートから、環境変数を取得する
 */
let GITHUB_OWNER = '';
let GITHUB_REPOSITORY = '';
let GITHUB_ACCESS_TOKEN = '';
let GITHUB_URL = '';
let SPREADSHEET_ID = '';

const loadConfiguration = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  const range = configSheet.getDataRange();
  const values = range.getValues();
  const config = {};
  values.forEach((row) => {
    config[row[0]] = row[1];
  });
  return config;
};

const config = loadConfiguration();
GITHUB_OWNER = config['GITHUB_OWNER'] || '';
GITHUB_REPOSITORY = config['GITHUB_REPOSITORY'] || '';
GITHUB_ACCESS_TOKEN = config['GITHUB_ACCESS_TOKEN'] || '';
SPREADSHEET_ID = config['SPREADSHEET_ID'] || '';
GITHUB_URL = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPOSITORY}` || '';

/**
 * スプレッドシートの列管理シートから、列番号を取得する
 */
const loadColumnIndexes = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const columnSheet = ss.getSheetByName('列管理'); // 列管理シートを指定
  const range = columnSheet.getRange(
    2,
    1,
    columnSheet.getLastRow() - 1,
    columnSheet.getLastColumn(),
  );
  const values = range.getValues();
  const columnIndexes = {};

  values.forEach((row) => {
    const columnName = row[1];
    const columnIndex = row[2];
    columnIndexes[columnName] = columnIndex;
  });
  return columnIndexes;
};

const START_ROW = 3;
const START_COLUMN = 1;
const columnIndexes = loadColumnIndexes();
const COLUMN_ISSUE_ID = columnIndexes['ISSUE_ID'];
const COLUMN_TITLE = columnIndexes['TITLE'];
const COLUMN_LABELS = columnIndexes['LABELS'];
const COLUMN_ASSIGNEES = columnIndexes['ASSIGNEES'];
const COLUMN_COMMENT = columnIndexes['COMMENT'];
const COLUMN_SEND_LABELS = columnIndexes['SEND_LABELS'];
const COLUMN_SEND_MENTION = columnIndexes['SEND_MENTION'];
const COLUMN_SEND_COMMENT = columnIndexes['SEND_COMMENT'];
const COLUMN_EXECUTION = columnIndexes['EXECUTION'];
const COLUMN_URL = columnIndexes['URL'];

/**
 * githubからlabelsを取得する
 */
const fetchAllLabels = () => {
  const url = `${GITHUB_URL}/labels`;
  const headers = {
    Authorization: 'token ' + GITHUB_ACCESS_TOKEN,
    Accept: 'application/vnd.github.v3+json',
  };
  const options = {
    method: 'GET',
    headers: headers,
    muteHttpExceptions: true,
  };

  /**
   * UrlFetchAppクラスのfetchメソッド：
   * 引数に指定したURLに対しHTTPリクエストが実行され、サーバーからのHTTPレスポンスが戻り値として取得できる
   */
  const response = UrlFetchApp.fetch(url, options);
  const getlabelsData = JSON.parse(response.getContentText());
  return getlabelsData.map((label) => label.name);
};

/** HTTP通信でgithubから取得したlabelsでスプシにプルダウンリストを作る関数 */
const initializeLabelDropdown = (sheet) => {
  const labels = fetchAllLabels();
  const range = sheet.getRange(START_ROW, COLUMN_SEND_LABELS, 300);

  /** スプシのプルダウンを作る用のコード */
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(labels, true)
    .setAllowInvalid(false)
    .setHelpText('Select a label from the list')
    .build();
  range.setDataValidation(rule);
};

/**
 * githubからissueを取得する
 */
const fetchIssues = () => {
  const url = `${GITHUB_URL}/issues`;
  const options = {
    method: 'GET',
    headers: { Authorization: `token ${GITHUB_ACCESS_TOKEN}` },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  const fetchIssuesData = JSON.parse(response.getContentText());
  return fetchIssuesData
    .filter((issue) => !issue.pull_request)
    .map((issue) => ({
      ...issue,
      assignees: issue.assignees.map((assignee) => assignee.login).join(', '),
    }));
};

/**
 * コメントを取得する
 */
const fetchLatestIssueComment = (issueId) => {
  const url = `${GITHUB_URL}/issues/${issueId}/comments`;
  const options = {
    method: 'GET',
    headers: { Authorization: `token ${GITHUB_ACCESS_TOKEN}` },
    muteHttpExceptions: true,
  };
  const response = UrlFetchApp.fetch(url, options);
  const comments = JSON.parse(response.getContentText());
  // コメントがあれば最新のものを返し、なければbodyを返す
  return comments.length !== 0 ? comments[comments.length - 1].body : issueId.body;
};

/**
 * issue（+コメント）を取得してスプシに記入する
 */
const populateSheetWithIssues = (sheet, logSheet) => {
  try {
    const issues = fetchIssues();
    issues.sort((a, b) => b.number - a.number);
    const lastColumn = sheet.getLastColumn();
    const data = [];

    issues.forEach((issue) => {
      const latestComment = fetchLatestIssueComment(issue.number) || issue.body;
      const titleWithLink = `=HYPERLINK("${issue.html_url}", "${issue.title}")`;
      const rowData = {
        [COLUMN_ISSUE_ID]: issue.number,
        [COLUMN_TITLE]: titleWithLink,
        [COLUMN_LABELS]: issue.labels.map((label) => label.name).join(', '),
        [COLUMN_ASSIGNEES]: issue.assignees,
        [COLUMN_COMMENT]: latestComment,
        [COLUMN_SEND_LABELS]: '',
        [COLUMN_SEND_MENTION]: '',
        [COLUMN_SEND_COMMENT]: '',
        [COLUMN_EXECUTION]: '',
      };
      // 必要な列数の配列を作成し、rowDataの値を設定
      const rowValues = Array.from({ length: lastColumn }, (_, i) => rowData[i + 1] || '');
      data.push(rowValues);
    });

    // 一括でデータを挿入
    if (data.length !== 0) {
      sheet.getRange(START_ROW, START_COLUMN, data.length, lastColumn).setValues(data);
    }
  } catch (error) {
    logSheet.appendRow([new Date(), 'Error populating sheet with issues:', error]);
  }
};

/**
 * スプシからgithubにlabelの追加
 */
/** HTTP通信を行う関数 */
const fetchCurrentIssueLabels = (issueId) => {
  const url = `${GITHUB_URL}/issues/${issueId}`;
  const headers = {
    Authorization: `token ${GITHUB_ACCESS_TOKEN}`,
    Accept: 'application/vnd.github.v3+json',
  };
  const options = {
    method: 'GET',
    headers: headers,
  };

  const response = UrlFetchApp.fetch(url, options);
  const issueData = JSON.parse(response.getContentText());
  return issueData.labels.map((label) => label.name);
};

/** スプシからgithubにlabelsを追加する */
const updateIssueLabelsInGithub = (sheet, editedRow) => {
  const issueId = sheet.getRange(editedRow, COLUMN_ISSUE_ID).getValue();
  const newLabels = sheet.getRange(editedRow, COLUMN_SEND_LABELS).getValue().split(',');

  if (newLabels.length === 0) {
    return; // 新しいラベルがない場合は何もせずに処理を終了
  }

  // 現在のラベルを取得
  const currentLabels = fetchCurrentIssueLabels(issueId);
  const allLabels = new Set([...currentLabels, ...newLabels]);

  // GitHub APIを通じてIssueのラベルを更新
  const url = `${GITHUB_URL}/issues/${issueId}`;
  const options = {
    method: 'PATCH',
    headers: {
      Authorization: `token ${GITHUB_ACCESS_TOKEN}`,
      Accept: 'application/vnd.github.v3+json',
    },
    payload: JSON.stringify({ labels: Array.from(allLabels) }),
  };
  UrlFetchApp.fetch(url, options);

  // スプシにラベルを更新
  sheet.getRange(editedRow, COLUMN_LABELS).setValue(Array.from(allLabels).join(', '));
  sheet.getRange(editedRow, COLUMN_SEND_LABELS).clearContent();
};

/**
 * スプシからgithubにcommentの追加
 */
const postCommentToGithubIssue = (sheet, editedRow) => {
  const issueId = sheet.getRange(editedRow, COLUMN_ISSUE_ID).getValue();
  const mention = sheet.getRange(editedRow, COLUMN_SEND_MENTION).getValue();
  const newComment = sheet.getRange(editedRow, COLUMN_SEND_COMMENT).getValue();
  let submitComment = mention ? `@${mention}\n${newComment}` : newComment;

  const url = `${GITHUB_URL}/issues/${issueId}/comments`;
  const options = {
    method: 'POST',
    headers: {
      Authorization: `token ${GITHUB_ACCESS_TOKEN}`,
      Accept: 'application/vnd.github.v3+json',
    },
    payload: JSON.stringify({ body: submitComment }),
    muteHttpExceptions: true,
  };
  UrlFetchApp.fetch(url, options);

  sheet.getRange(editedRow, COLUMN_SEND_MENTION).clearContent();
  sheet.getRange(editedRow, COLUMN_SEND_COMMENT).clearContent();
};

// Issueが閉じられた時の処理
const handleClosedIssue = (sheet, logSheet, lastRow, issueId) => {
  for (let i = 1; i <= lastRow; i++) {
    const rowIssueId = sheet.getRange(i, 1).getValue();
    if (rowIssueId === issueId) {
      sheet.deleteRow(i);
      break;
    }
  }
};

// Issueが開かれた時の処理
const setRowValues = (sheet, rowNum, issue) => {
  const lastColumn = sheet.getLastColumn();
  sheet
    .getRange(rowNum, COLUMN_TITLE, 1, lastColumn)
    .setValues([
      [
        `=HYPERLINK("${issue.html_url}", "${issue.title}")`,
        issue.labels.map((label) => label.name).join(', '),
        issue.assignees.map((assignee) => assignee.login).join(', '),
        issue.body,
        issue.html_url,
      ],
    ]);
};

const handleOpenedIssue = (sheet, logSheet, lastRow, insertRow, issue, issueId) => {
  // 全データを取得
  const lastColumn = sheet.getLastColumn();
  const data = sheet.getRange(START_ROW, START_COLUMN, lastRow, lastColumn).getValues();
  const rowIndex = data.findIndex((row) => row[0] === issueId);

  // 既存のIssueがあるかを確認: あれば更新、なければ新規追加
  if (rowIndex !== -1) {
    // 更新処理
    setRowValues(sheet, rowIndex + 1, issue);
  } else if (lastRow >= insertRow) {
    // 行を挿入してデータを設定
    logSheet.appendRow([new Date(), 'reopened']);
    sheet.insertRowBefore(insertRow);
    sheet
      .getRange(insertRow, 1, 1, 5)
      .setValues([
        [
          issueId,
          `=HYPERLINK("${issue.html_url}", "${issue.title}")`,
          issue.labels.map((label) => label.name).join(', '),
          issue.assignees.map((assignee) => assignee.login).join(', '),
          issue.body,
        ],
      ]);
  }
};

// コメントが作成された時の処理
const handleCreatedAction = (sheet, lastRow, issueId) => {
  const latestComment = fetchLatestIssueComment(issueId);
  for (let i = 1; i <= lastRow; i++) {
    const rowIssueId = sheet.getRange(i, 1).getValue();
    if (rowIssueId === issueId) {
      sheet.getRange(i, COLUMN_COMMENT).setValue(latestComment);
      break;
    }
  }
};

// labelが追加・削除された時の処理
const handleLabeled = (sheet, lastRow, issueId) => {
  const labels = fetchCurrentIssueLabels(issueId);
  for (let i = 1; i <= lastRow; i++) {
    const rowIssueId = sheet.getRange(i, COLUMN_ISSUE_ID).getValue();
    if (rowIssueId === issueId) {
      sheet.getRange(i, COLUMN_LABELS).setValue(Array.from(labels).join(', '));
      break;
    }
  }
};

/**
 * Issueのassigneesが変更されたときの処理
 */
const handleAssigneesChanged = (sheet, lastRow, issueId) => {
  const issue = fetchIssue(issueId);
  const assignees = issue.assignees.map((assignee) => assignee.login).join(', ');
  for (let i = 1; i <= lastRow; i++) {
    const rowIssueId = sheet.getRange(i, 1).getValue();
    if (rowIssueId === issueId) {
      sheet.getRange(i, COLUMN_ASSIGNEES).setValue(assignees);
      break;
    }
  }
};

// Issueのassigneesが変更されたときの処理
const updateAssigneesInSheet = (sheet, logSheet, lastRow, issueId) => {
  const url = `${GITHUB_URL}/issues/${issueId}`;
  const options = {
    method: 'GET',
    headers: {
      Authorization: `token ${GITHUB_ACCESS_TOKEN}`,
      Accept: 'application/vnd.github.v3+json',
    },
  };
  // GitHubからIssueの詳細を取得する
  const response = UrlFetchApp.fetch(url, options);
  const issue = JSON.parse(response.getContentText());
  // Issueのassigneesを取得する
  const assignees = issue.assignees.map((assignee) => assignee.login).join(', ');

  // スプレッドシートの対応する行にassigneesを反映する
  for (let i = 1; i <= lastRow; i++) {
    const rowIssueId = sheet.getRange(i, 1).getValue();
    if (rowIssueId === issueId) {
      sheet.getRange(i, COLUMN_ASSIGNEES).setValue(assignees); // 8列目にassigneesを反映
      break;
    }
  }
};

/**
 * トリガー関数
 */
const getOrCreateSheet = (sheetName) => {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
  }
  return sheet;
};

const onEdit = (e) => {
  const range = e.range;
  const sheet = range.getSheet();
  const editedCellColumn = range.getColumn();
  const editedRow = range.getRow();
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');

  if (sheet.getName() === 'Issues' && editedCellColumn === COLUMN_EXECUTION && editedRow === 2) {
    // 初期設定（github issueをスプシに反映）
    initializeLabelDropdown(sheet);
    populateSheetWithIssues(sheet, logSheet);
  } else if (
    sheet.getName() === 'Issues' &&
    editedCellColumn === COLUMN_EXECUTION &&
    editedRow > 2
  ) {
    // スプシを更新する度に実行
    postCommentToGithubIssue(sheet, editedRow);
    updateIssueLabelsInGithub(sheet, editedRow);
  }
};

const doPost = (e) => {
  const postData = JSON.parse(e.postData.contents);
  const issue = postData.issue;
  const action = postData.action;
  const issueId = issue.number;
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ログ');

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Issues');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Issues');
  }
  const lastRow = sheet.getLastRow();
  const insertRow = 3;

  if (action === 'closed') {
    handleClosedIssue(sheet, logSheet, lastRow, issueId);
  } else if (action === 'opened' || action === 'reopened') {
    logSheet.appendRow([new Date(), 'reopened']);
    handleOpenedIssue(sheet, logSheet, lastRow, insertRow, issue, issueId);
  } else if (action === 'created') {
    handleCreatedAction(sheet, lastRow, issueId);
  } else if (action === 'labeled' || action === 'unlabeled') {
    handleLabeled(sheet, lastRow, issueId);
  } else if (action === 'assigned' || action === 'unassigned') {
    handleAssigneesChanged(sheet, lastRow, issueId);
  }
};
