const MANAGEMENT_HEADERS = [
  '送信する',

  '送信ステータス_1',
  '送信日時_1',
  'エラー_1',
  'プレビュー件名_1',
  'プレビュー本文_1',

  '送信ステータス_2',
  '送信日時_2',
  'エラー_2',
  'プレビュー件名_2',
  'プレビュー本文_2'
];

const SHARED_DATA_FIELDS = [
  'Session',
  'SessionID',
  'Date',
  'Start',
  'End'
];

const REQUIRED_DATA_FIELDS = [
  'Session',
  'SessionID',
  'Date',
  'Start',
  'End',

  'Recipient_1',
  'Affiliation_1',
  'Name_1',

  'Recipient_2',
  'Affiliation_2',
  'Name_2'
];

const OPTIONAL_DATA_FIELDS = [
  'CC',
  'BCC'
];

const PERSON_NUMBERS = [1, 2];

/**
 * スプレッドシートを開いたときにメニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSJ2026メール送信')
    .addItem('1. 送信用列を準備', 'prepareChairInviteSheet')
    .addItem('2. テスト作成：実際には送らない', 'previewChairInviteEmails')
    .addItem('3. 本番送信', 'sendChairInviteEmails')
    .addToUi();
}

/**
 * 管理列を追加します。
 */
function prepareChairInviteSheet() {
  const sheet = getTargetSheet_();
  ensureManagementColumns_(sheet);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

/**
 * テスト処理。
 * 実際にはメール送信しません。
 */
function previewChairInviteEmails() {
  processChairInviteEmails_(true);
}

/**
 * 本番送信。
 */
function sendChairInviteEmails() {
  processChairInviteEmails_(false);
}

/**
 * メイン処理。
 * 1行につき Recipient_1 と Recipient_2 に個別メールを送ります。
 */
function processChairInviteEmails_(dryRun) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error('別の送信処理が実行中です。少し待ってから再実行してください。');
  }

  try {
    const sheet = getTargetSheet_();
    ensureManagementColumns_(sheet);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) {
      throw new Error('データ行がありません。');
    }

    const rawValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const displayValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    const headers = displayValues[0].map(value => String(value || '').trim());

    const managementIndexes = getManagementIndexes_(headers);
    const dataIndexes = getDataIndexes_(headers);

    validateRequiredColumns_(dataIndexes);
    validateConfigUrls_();

    let processedCount = 0;
    let remainingQuota = MailApp.getRemainingDailyQuota();

    rowLoop:
    for (let r = 1; r < displayValues.length; r++) {
      const rowNumber = r + 1;
      const rawRow = rawValues[r];
      const displayRow = displayValues[r];

      if (isBlankRow_(displayRow)) {
        continue;
      }

      const shouldSend = shouldProcessRow_(rawRow, displayRow, managementIndexes);

      if (!shouldSend) {
        continue;
      }

      const baseRowData = buildBaseRowData_(headers, displayRow, dataIndexes);

      for (const personNo of PERSON_NUMBERS) {
        if (processedCount >= CONFIG.maxEmailsPerRun) {
          break rowLoop;
        }

        const statusHeader = `送信ステータス_${personNo}`;
        const currentStatus = getCellByIndex_(displayRow, managementIndexes[statusHeader]);

        if (String(currentStatus || '').trim() === '送信済み') {
          continue;
        }

        const rowData = buildPersonRowData_(baseRowData, personNo);

        const to = normalizeRecipients_(rowData.Recipient);
        const cc = normalizeRecipients_(rowData.CC);
        const bcc = normalizeRecipients_(rowData.BCC);

        const subject = applyTemplate_(SUBJECT_TEMPLATE, rowData);
        const body = applyTemplate_(BODY_TEMPLATE, rowData);
    
        const missingPlaceholders = getMissingPlaceholders_(
          SUBJECT_TEMPLATE + '\n' + BODY_TEMPLATE,
          rowData
        );

        if (missingPlaceholders.length > 0) {
          setResultForPerson_(
            sheet,
            rowNumber,
            managementIndexes,
            personNo,
            'エラー',
            '',
            `存在しない差し込み項目です: ${missingPlaceholders.join(', ')}`,
            subject,
            body
          );
          continue;
        }

        const validationError = validateMailData_(
          rowData,
          personNo,
          to,
          cc,
          bcc,
          subject,
          body
        );

        if (validationError) {
          setResultForPerson_(
            sheet,
            rowNumber,
            managementIndexes,
            personNo,
            'エラー',
            '',
            validationError,
            subject,
            body
          );
          continue;
        }

        const recipientCount =
          splitRecipients_(to).length +
          splitRecipients_(cc).length +
          splitRecipients_(bcc).length;

        if (!dryRun && recipientCount > remainingQuota) {
          setResultForPerson_(
            sheet,
            rowNumber,
            managementIndexes,
            personNo,
            '未送信',
            '',
            `本日の残り送信可能数が不足しています。必要数: ${recipientCount}, 残り: ${remainingQuota}`,
            subject,
            body
          );
          break rowLoop;
        }

        try {
          if (dryRun) {
            setResultForPerson_(
              sheet,
              rowNumber,
              managementIndexes,
              personNo,
              'テストOK',
              '',
              '実際には送信していません',
              subject,
              body
            );
          } else {
            const options = {};

            if (CONFIG.senderName) {
              options.name = CONFIG.senderName;
            }

            if (CONFIG.replyTo) {
              options.replyTo = CONFIG.replyTo;
            }

            if (cc) {
              options.cc = cc;
            }

            if (bcc) {
              options.bcc = bcc;
            }

            GmailApp.sendEmail(to, subject, body, options);

            remainingQuota -= recipientCount;

            setResultForPerson_(
              sheet,
              rowNumber,
              managementIndexes,
              personNo,
              '送信済み',
              new Date(),
              '',
              subject,
              body
            );

            Utilities.sleep(CONFIG.sleepMs);
          }

          processedCount++;

        } catch (error) {
          setResultForPerson_(
            sheet,
            rowNumber,
            managementIndexes,
            personNo,
            'エラー',
            '',
            error.message,
            subject,
            body
          );
        }
      }
    }

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }
}

/**
 * 対象スプレッドシート・対象シートを取得します。
 */
function getTargetSheet_() {
  const ss = SpreadsheetApp.openByUrl(CONFIG.spreadsheetUrl);
  const sheet = getSheetByGid_(ss, CONFIG.sheetGid);

  if (!sheet) {
    throw new Error(`gid=${CONFIG.sheetGid} のシートが見つかりません。`);
  }

  return sheet;
}

/**
 * gid からシートを取得します。
 */
function getSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    if (String(sheet.getSheetId()) === String(gid)) {
      return sheet;
    }
  }

  return null;
}

/**
 * 管理列をなければ追加します。
 */
function ensureManagementColumns_(sheet) {
  let lastCol = sheet.getLastColumn();
  let headers = [];

  if (lastCol > 0) {
    headers = sheet
      .getRange(1, 1, 1, lastCol)
      .getDisplayValues()[0]
      .map(value => String(value || '').trim());
  }

  MANAGEMENT_HEADERS.forEach(header => {
    if (findHeaderIndex_(headers, header) === -1) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(header);
      headers.push(header);
    }
  });

  const sendFlagIndex = findHeaderIndex_(headers, '送信する');

  if (sendFlagIndex !== -1 && sheet.getMaxRows() >= 2) {
    sheet
      .getRange(2, sendFlagIndex + 1, sheet.getMaxRows() - 1, 1)
      .insertCheckboxes();
  }
}

/**
 * 管理列の位置を取得します。
 */
function getManagementIndexes_(headers) {
  const indexes = {};

  MANAGEMENT_HEADERS.forEach(header => {
    indexes[header] = findHeaderIndex_(headers, header);
  });

  return indexes;
}

/**
 * データ列の位置を取得します。
 */
function getDataIndexes_(headers) {
  const indexes = {};
  const fields = REQUIRED_DATA_FIELDS.concat(OPTIONAL_DATA_FIELDS);

  fields.forEach(field => {
    indexes[field] = findHeaderIndex_(headers, field);
  });

  return indexes;
}

/**
 * 必須列チェック。
 */
function validateRequiredColumns_(dataIndexes) {
  const missing = REQUIRED_DATA_FIELDS.filter(field => dataIndexes[field] === -1);

  if (missing.length > 0) {
    throw new Error(`必須列が見つかりません: ${missing.join(', ')}`);
  }
}

/**
 * 行データの基本部分を作ります。
 */
function buildBaseRowData_(headers, displayRow, dataIndexes) {
  const rowData = {};

  headers.forEach((header, index) => {
    if (header) {
      rowData[header] = String(displayRow[index] || '').trim();
    }
  });

  Object.keys(dataIndexes).forEach(field => {
    const index = dataIndexes[field];

    rowData[field] = index === -1
      ? ''
      : String(displayRow[index] || '').trim();
  });

  rowData.AnswerFormUrl = CONFIG.answerFormUrl;
  rowData.ProgramUrl = CONFIG.programUrl;
  rowData.DeadlineChair1st = CONFIG.deadlineChair1st;

  rowData.Today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy/MM/dd'
  );

  return rowData;
}

/**
 * 1人分の差し込みデータを作ります。
 * Recipient_1 / Affiliation_1 / Name_1 を
 * Recipient / Affiliation / Name に写して本文へ差し込みます。
 */
function buildPersonRowData_(baseRowData, personNo) {
  const rowData = Object.assign({}, baseRowData);

  rowData.PersonNo = String(personNo);
  rowData.Recipient = baseRowData[`Recipient_${personNo}`] || '';
  rowData.Affiliation = baseRowData[`Affiliation_${personNo}`] || '';
  rowData.Name = baseRowData[`Name_${personNo}`] || '';

  return rowData;
}

/**
 * {{Name}} のような差し込みタグを置換します。
 */
function applyTemplate_(template, rowData) {
  return String(template || '').replace(/{{\s*([^{}]+?)\s*}}/g, function(match, key) {
    const normalizedKey = String(key || '').trim();

    if (rowData[normalizedKey] === undefined) {
      return '';
    }

    return String(rowData[normalizedKey]);
  });
}

/**
 * テンプレート内に存在する {{...}} のうち、
 * rowData に存在しないものを返します。
 */
function getMissingPlaceholders_(template, rowData) {
  const missing = [];
  const regex = /{{\s*([^{}]+?)\s*}}/g;
  let match;

  while ((match = regex.exec(String(template || ''))) !== null) {
    const key = String(match[1] || '').trim();

    if (rowData[key] === undefined && !missing.includes(key)) {
      missing.push(key);
    }
  }

  return missing;
}

/**
 * メールデータのチェック。
 */
function validateMailData_(rowData, personNo, to, cc, bcc, subject, body) {
  const emptyFields = [];

  SHARED_DATA_FIELDS.forEach(field => {
    if (!String(rowData[field] || '').trim()) {
      emptyFields.push(field);
    }
  });

  ['Recipient', 'Affiliation', 'Name'].forEach(field => {
    if (!String(rowData[field] || '').trim()) {
      emptyFields.push(`${field}_${personNo}`);
    }
  });

  if (emptyFields.length > 0) {
    return `空欄の必須項目があります: ${emptyFields.join(', ')}`;
  }

  const toRecipients = splitRecipients_(to);

  if (toRecipients.length === 0) {
    return `Recipient_${personNo} のメールアドレスが空です。`;
  }

  if (toRecipients.length > 1) {
    return `Recipient_${personNo} にはメールアドレスを1件だけ入れてください。複数入れると受信者同士が見える可能性があります。`;
  }

  if (!subject) {
    return '件名が空です。';
  }

  if (!body) {
    return '本文が空です。';
  }

  const recipients = [
    ...toRecipients,
    ...splitRecipients_(cc),
    ...splitRecipients_(bcc)
  ];

  const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const invalidEmails = recipients.filter(address => !emailPattern.test(address));

  if (invalidEmails.length > 0) {
    return `メールアドレス形式エラー: ${invalidEmails.join(', ')}`;
  }

  return '';
}

/**
 * 回答フォームURL・プログラムURLの入れ忘れチェック。
 */
function validateConfigUrls_() {
  if (!/^https?:\/\//.test(CONFIG.answerFormUrl)) {
    throw new Error('CONFIG.answerFormUrl に回答フォームURLを設定してください。');
  }

  if (!/^https?:\/\//.test(CONFIG.programUrl)) {
    throw new Error('CONFIG.programUrl に暫定版プログラムURLを設定してください。');
  }
}

/**
 * 1人分の結果を書き込みます。
 */
function setResultForPerson_(sheet, rowNumber, managementIndexes, personNo, status, sentAt, errorMessage, previewSubject, previewBody) {
  sheet
    .getRange(rowNumber, managementIndexes[`送信ステータス_${personNo}`] + 1)
    .setValue(status || '');

  sheet
    .getRange(rowNumber, managementIndexes[`送信日時_${personNo}`] + 1)
    .setValue(sentAt || '');

  sheet
    .getRange(rowNumber, managementIndexes[`エラー_${personNo}`] + 1)
    .setValue(errorMessage || '');

  sheet
    .getRange(rowNumber, managementIndexes[`プレビュー件名_${personNo}`] + 1)
    .setValue(previewSubject || '');

  sheet
    .getRange(rowNumber, managementIndexes[`プレビュー本文_${personNo}`] + 1)
    .setValue(previewBody || '');
}

/**
 * チェックボックス列を見るか、全行送信設定を見るかを判定します。
 */
function shouldProcessRow_(rawRow, displayRow, managementIndexes) {
  if (CONFIG.sendAllRowsIgnoringFlag === true) {
    return true;
  }

  const sendFlagIndex = managementIndexes['送信する'];

  if (sendFlagIndex === -1) {
    return false;
  }

  return isTruthy_(rawRow[sendFlagIndex]) || isTruthy_(displayRow[sendFlagIndex]);
}

/**
 * TRUE / 送信 / yes などを真として扱います。
 */
function isTruthy_(value) {
  if (value === true) {
    return true;
  }

  const text = String(value || '').trim().toLowerCase();

  return ['true', '1', 'yes', 'y', '送信', '送る'].includes(text);
}

/**
 * 宛先をカンマ区切りに正規化します。
 */
function normalizeRecipients_(value) {
  return splitRecipients_(value).join(',');
}

/**
 * 宛先を配列にします。
 */
function splitRecipients_(value) {
  return String(value || '')
    .split(/[,\n;、；]+/)
    .map(address => address.trim())
    .filter(address => address);
}

/**
 * 空行判定。
 */
function isBlankRow_(row) {
  return row.every(value => String(value || '').trim() === '');
}

/**
 * 指定インデックスのセル値を返します。
 */
function getCellByIndex_(row, index) {
  if (index === -1 || index === undefined) {
    return '';
  }

  return row[index];
}

/**
 * ヘッダー位置を探します。
 * 列名の完全一致で探します。
 * ただし、前後の空白だけは無視します。
 */
function findHeaderIndex_(headers, targetHeader) {
  const target = String(targetHeader || '').trim();

  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').trim();

    if (header === target) {
      return i;
    }
  }

  return -1;
}
