// GAS script for RSJ2026 session generation
//
// Usage:
//   1. CONFIG.PAPER_SHEET_NAME に元データのシート名を設定する
//   2. session_template シートを用意する
//   3. 「セッション自動生成」→「セッションシート自動作成」を実行する
//   4. 生成済みシートを削除する場合は
//      「自動生成シート削除」を実行する


/**
 * 自動生成したセッションシートを削除する
 *
 * 対象：
 *   OS_...
 *   IS_...
 *   GS_...
 */
/**
 * GS・IS・OSの生成シートをすべて削除する
 */
function removeGeneratedSheets() {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheets =
    spreadsheet.getSheets();

  let removedCount = 0;

  sheets.forEach(function (sheet) {
    const sheetName =
      sheet.getName();

    if (/^(GS|IS|OS)_/.test(sheetName)) {
      Logger.log(
        "Remove sheet: " +
        sheetName
      );

      spreadsheet.deleteSheet(sheet);

      removedCount++;
    }
  });

  SpreadsheetApp.getUi().alert(
    "全生成シートの削除が完了しました。\n" +
    "削除数: " +
    removedCount
  );
}


/**
 * GSシートを削除する
 */
function removeGeneratedGSSheets() {
  removeGeneratedSheetsByType("GS");
}


/**
 * ISシートを削除する
 */
function removeGeneratedISSheets() {
  removeGeneratedSheetsByType("IS");
}


/**
 * OSシートを削除する
 */
function removeGeneratedOSSheets() {
  removeGeneratedSheetsByType("OS");
}


/**
 * 指定した種類の生成シートを削除する
 *
 * @param {string} sessionType
 *   "GS", "IS", "OS"
 */
function removeGeneratedSheetsByType(
  sessionType
) {
  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheets =
    spreadsheet.getSheets();

  const pattern = new RegExp(
    "^" + sessionType + "_"
  );

  let removedCount = 0;

  sheets.forEach(function (sheet) {
    const sheetName =
      sheet.getName();

    if (pattern.test(sheetName)) {
      Logger.log(
        "Remove sheet: " +
        sheetName
      );

      spreadsheet.deleteSheet(sheet);

      removedCount++;
    }
  });

  SpreadsheetApp.getUi().alert(
    sessionType +
    "シートの削除が完了しました。\n" +
    "削除数: " +
    removedCount
  );
}


/**
 * セッションごとのシートを生成する
 */
function generateSessionSheets() {
  Logger.log("generateSessionSheets()");

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  /*
   * 元データシートを取得
   */
  const paperSheet = spreadsheet.getSheetByName(
    CONFIG.PAPER_SHEET_NAME
  );

  if (!paperSheet) {
    throw new Error(
      "元データシートが見つかりません: " +
      CONFIG.PAPER_SHEET_NAME
    );
  }

  /*
   * テンプレートシートを取得
   */
  const templateSheet = spreadsheet.getSheetByName(
    CONFIG.TEMPLATE_SHEET_NAME
  );

  if (!templateSheet) {
    throw new Error(
      "テンプレートシートが見つかりません: " +
      CONFIG.TEMPLATE_SHEET_NAME
    );
  }

  const lastRow = paperSheet.getLastRow();
  const lastColumn = paperSheet.getLastColumn();

  if (lastRow < CONFIG.DATA_START_ROW) {
    throw new Error(
      "元データに発表情報がありません。"
    );
  }

  Logger.log("Last row: " + lastRow);
  Logger.log("Last column: " + lastColumn);

  /*
   * 元データを一括取得
   */
  const allValues = paperSheet
    .getRange(
      CONFIG.HEADER_ROW,
      1,
      lastRow - CONFIG.HEADER_ROW + 1,
      lastColumn
    )
    .getDisplayValues();

  const headers = allValues[0];

  const dataRows = allValues.slice(
    CONFIG.DATA_START_ROW - CONFIG.HEADER_ROW
  );

  Logger.log(
    "Headers: " +
    JSON.stringify(headers)
  );

  /*
   * 必須列を確認
   */
  const requiredHeaders = [
    "No.",
    "セッションの選択",
    "セッションID",
    "セッション名",
    "必要スロット数",
    "演題名(日本語)"
  ];

  validateRequiredHeaders(
    headers,
    requiredHeaders
  );

  /*
   * 生成シートの表部分へ出力する列
   *
   * session_template の見出しと
   * 同じ順番にする
   *
   * セッション種類、セッションID、セッション名は
   * 表部分には出力しない
   */
  const outputColumns = [
    {
      outputName: "スロット番号",
      sourceHeader: "セッション内のスロット番号",
      defaultValue: ""
    },
    {
      outputName: "発表順",
      sourceHeader: "スロット内順番",
      defaultValue: ""
    },
    {
      outputName: "講演No.",
      sourceHeader: "No.",
      defaultValue: ""
    },
    {
      outputName: "演題名(日本語)",
      sourceHeader: "演題名(日本語)",
      defaultValue: ""
    },
    {
      outputName: "論文副題",
      sourceHeader: "論文副題",
      defaultValue: ""
    },
    {
      outputName: "講演概要文",
      sourceHeader: "講演概要文",
      defaultValue: ""
    },
    {
      outputName: "登録者名(氏名)",
      sourceHeader: "登録者名(氏名)",
      defaultValue: ""
    },
    {
      outputName: "所属機関 (大学 / 勤務先)",
      sourceHeader: "所属機関 (大学 / 勤務先)",
      defaultValue: ""
    },
    {
      outputName: "E-mail",
      sourceHeader: "E-mail",
      defaultValue: ""
    }
  ];

  /*
   * 各出力列について、
   * 元データ上の0始まりインデックスを取得
   */
  const outputColumnIndexes =
    outputColumns.map(function (column) {
      return headers.indexOf(
        column.sourceHeader
      );
    });

  /*
   * 処理に使用する列のインデックス
   *
   * これらは表への出力有無とは無関係に、
   * 元データから直接取得する
   */
  const sessionIdColumnIndex =
    headers.indexOf("セッションID");

  const sessionNameColumnIndex =
    headers.indexOf("セッション名");

  const requiredSlotCountColumnIndex =
    headers.indexOf("必要スロット数");

  const slotNumberColumnIndex =
    headers.indexOf(
      "セッション内のスロット番号"
    );

  const presentationOrderColumnIndex =
    headers.indexOf("スロット内順番");

  /*
   * セッションごとにデータを分類する
   *
   * 例：
   * {
   *   "OS01": [row1, row2, ...],
   *   "GS01": [row3, row4, ...]
   * }
   */
  const sessionDataMap = {};

  dataRows.forEach(function (row, rowIndex) {
    const rawSessionId =
      row[sessionIdColumnIndex];

    const sessionId =
      normalizeSessionId(rawSessionId);

    /*
     * セッションIDが空欄の場合はスキップ
     */
    if (!sessionId) {
      return;
    }

    /*
     * OS、IS、GS以外のIDはスキップ
     */
    if (
      !CONFIG.SESSION_ID_PATTERN.test(
        sessionId
      )
    ) {
      Logger.log(
        "対象外のセッションIDをスキップ: " +
        sessionId +
        " / 元データ行: " +
        (
          rowIndex +
          CONFIG.DATA_START_ROW
        )
      );

      return;
    }

    if (!sessionDataMap[sessionId]) {
      sessionDataMap[sessionId] = [];
    }

    sessionDataMap[sessionId].push(row);
  });

  /*
   * セッションIDを自然順に並べる
   *
   * OS01、OS02、IS01、GS01 ...
   */
  const sessionIds =
    Object.keys(sessionDataMap)
      .sort(compareSessionIds);

  Logger.log(
    "Session IDs: " +
    JSON.stringify(sessionIds)
  );

  if (sessionIds.length === 0) {
    throw new Error(
      "有効なセッションIDが見つかりませんでした。\n" +
      "「セッションID」列を確認してください。"
    );
  }

  /*
   * セッションシートを生成
   */
  let generatedCount = 0;
  let skippedCount = 0;

  sessionIds.forEach(function (sessionId) {
    const sessionRows =
      sessionDataMap[sessionId];

    /*
     * スロット番号とスロット内順番がある場合は
     * その順番で並べる
     */
    sortSessionRows(
      sessionRows,
      slotNumberColumnIndex,
      presentationOrderColumnIndex
    );

    const created = createSessionSheet(
      spreadsheet,
      templateSheet,
      sessionId,
      sessionRows,
      outputColumns,
      outputColumnIndexes,
      sessionNameColumnIndex,
      requiredSlotCountColumnIndex
    );

    if (created) {
      generatedCount++;
    } else {
      skippedCount++;
    }
  });

  SpreadsheetApp.flush();

  Logger.log(
    "Generated session sheets: " +
    generatedCount
  );

  Logger.log(
    "Skipped session sheets: " +
    skippedCount
  );

  SpreadsheetApp.getUi().alert(
    "セッションシートの作成が完了しました。\n\n" +
    "生成数: " +
    generatedCount +
    "\n" +
    "スキップ数: " +
    skippedCount
  );
}


/**
 * 1つのセッションシートを生成する
 *
 * @return {boolean}
 *   true  : シートを生成した
 *   false : 既存シートのためスキップした
 */
function createSessionSheet(
  spreadsheet,
  templateSheet,
  sessionId,
  sessionRows,
  outputColumns,
  outputColumnIndexes,
  sessionNameColumnIndex,
  requiredSlotCountColumnIndex
) {
  const writeSheetName =
    convertSessionIdToSheetName(
      sessionId
    );

  Logger.log(
    "Create session: " +
    sessionId +
    " -> " +
    writeSheetName
  );

  const existingSheet =
    spreadsheet.getSheetByName(
      writeSheetName
    );

  /*
   * 同名シートがすでに存在する場合
   */
  if (existingSheet) {
    if (
      CONFIG.OVERWRITE_EXISTING_SHEETS
    ) {
      spreadsheet.deleteSheet(
        existingSheet
      );

      Logger.log(
        "Existing sheet removed: " +
        writeSheetName
      );
    } else {
      Logger.log(
        "Existing sheet skipped: " +
        writeSheetName
      );

      return false;
    }
  }

  /*
   * セッション名が統一されているか確認する
   */
  validateSessionName(
    sessionId,
    sessionRows,
    sessionNameColumnIndex
  );

  /*
   * 必要スロット数が統一されているか確認する
   */
  validateRequiredSlotCount(
    sessionId,
    sessionRows,
    requiredSlotCountColumnIndex
  );

  /*
   * テンプレートをコピー
   */
  const writeSheet = templateSheet
    .copyTo(spreadsheet)
    .setName(writeSheetName);

  /*
   * 元データから出力用の
   * 2次元配列を作成する
   */
  const writeValues =
    sessionRows.map(
      function (sourceRow, rowIndex) {
        return outputColumns.map(
          function (
            column,
            outputIndex
          ) {
            const sourceColumnIndex =
              outputColumnIndexes[
                outputIndex
              ];

            /*
             * 元データに「スロット内順番」列が
             * 存在しない場合は、
             * 発表順を1から自動採番する
             */
            if (
              column.outputName ===
                "発表順" &&
              sourceColumnIndex === -1
            ) {
              return rowIndex + 1;
            }

            /*
             * 元データに存在しない列は空欄
             */
            if (
              sourceColumnIndex === -1
            ) {
              return column.defaultValue;
            }

            return sourceRow[
              sourceColumnIndex
            ];
          }
        );
      }
    );

  /*
   * 発表情報を一括書き込み
   */
  if (writeValues.length > 0) {
    writeSheet
      .getRange(
        CONFIG.WRITE_SHEET_BASE_ROW,
        1,
        writeValues.length,
        writeValues[0].length
      )
      .setValues(writeValues);
  }

  /*
   * 講演数
   */
  const presentationCount =
    sessionRows.length;

  /*
   * セッション名を元データから直接取得
   *
   * outputColumnsに「セッション名」がなくても
   * 取得できる
   */
  const sessionName =
    getFirstNonEmptyValue(
      sessionRows,
      sessionNameColumnIndex
    );

  /*
   * 必要スロット数を元データから直接取得
   */
  const requiredSlotCount =
    getFirstNonEmptyValue(
      sessionRows,
      requiredSlotCountColumnIndex
    );

  /*
   * テンプレート上の
   * プレースホルダーを置換する
   */
  replaceTemplateInformation(
    writeSheet,
    sessionId,
    sessionName,
    presentationCount,
    requiredSlotCount
  );

  Logger.log(
    "Created: " +
    writeSheetName +
    ", session name: " +
    sessionName +
    ", presentation count: " +
    presentationCount +
    ", required slot count: " +
    requiredSlotCount
  );

  SpreadsheetApp.flush();

  Logger.log(
    "書き込み確認: " +
    writeSheet.getName() +
    " / B1=[" +
    writeSheet.getRange("B1").getDisplayValue() +
    "] / B2=[" +
    writeSheet.getRange("B2").getDisplayValue() +
    "] / B3=[" +
    writeSheet.getRange("B3").getDisplayValue() +
    "] / B4=[" +
    writeSheet.getRange("B4").getDisplayValue() +
    "]"
  );

  return true;
}


/**
 * 指定した列から最初の空欄でない値を取得する
 */
function getFirstNonEmptyValue(
  rows,
  columnIndex
) {
  if (columnIndex === -1) {
    return "";
  }

  for (
    let rowIndex = 0;
    rowIndex < rows.length;
    rowIndex++
  ) {
    const value =
      rows[rowIndex][columnIndex];

    if (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ""
    ) {
      return value;
    }
  }

  return "";
}


/**
 * テンプレート内のセッション情報を置換する
 *
 * session_template：
 *
 * A1 セッションID
 * B1 <<セッションID>>
 *
 * A2 セッション名
 * B2 <<セッション名>>
 *
 * A3 講演数
 * B3 <<講演数>>
 *
 * A4 必要スロット数
 * B4 <<必要スロット数>>
 */
function replaceTemplateInformation(
  sheet,
  sessionId,
  sessionName,
  presentationCount,
  requiredSlotCount
) {
  const replaceRange =
    sheet.getRange(
      CONFIG.TEMPLATE_INFO_START_ROW,
      CONFIG.TEMPLATE_INFO_COL,
      CONFIG.TEMPLATE_INFO_NUM_ROWS,
      1
    );

  const replaceData =
    replaceRange.getValues();

  for (
    let rowIndex = 0;
    rowIndex < replaceData.length;
    rowIndex++
  ) {
    let value =
      replaceData[rowIndex][0];

    if (
      value === null ||
      value === undefined
    ) {
      value = "";
    }

    value = String(value);

    value = value.replace(
      /<<セッションID>>/g,
      String(sessionId)
    );

    value = value.replace(
      /<<セッション名>>/g,
      String(sessionName)
    );

    value = value.replace(
      /<<講演数>>/g,
      String(presentationCount)
    );

    value = value.replace(
      /<<必要スロット数>>/g,
      String(requiredSlotCount)
    );

    replaceData[rowIndex][0] =
      value;
  }

  replaceRange.setValues(
    replaceData
  );
}


/**
 * セッションIDを正規化する
 *
 * 例：
 *   " os01 " -> "OS01"
 *   "ＯＳ０１" -> "OS01"
 */
function normalizeSessionId(value) {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value)
    .normalize("NFKC")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}


/**
 * セッションIDから生成シート名を作る
 *
 * 例：
 *   OS01 -> OS_01
 *   IS2  -> IS_2
 *   GS10 -> GS_10
 */
function convertSessionIdToSheetName(
  sessionId
) {
  const normalizedSessionId =
    normalizeSessionId(sessionId);

  const match =
    normalizedSessionId.match(
      /^(OS|IS|GS)(\d+)$/i
    );

  if (!match) {
    throw new Error(
      "シート名へ変換できないセッションIDです: " +
      sessionId
    );
  }

  const type =
    match[1].toUpperCase();

  const number =
    match[2];

  return type + "_" + number;
}


/**
 * セッションIDの並び順を比較する
 *
 * 並び順：
 *   OS
 *   IS
 *   GS
 *
 * 同じ種類では番号順
 */
function compareSessionIds(a, b) {
  const typeOrder = {
    OS: 1,
    IS: 2,
    GS: 3
  };

  const matchA = a.match(
    /^(OS|IS|GS)(\d+)$/i
  );

  const matchB = b.match(
    /^(OS|IS|GS)(\d+)$/i
  );

  if (!matchA || !matchB) {
    return a.localeCompare(
      b,
      "ja",
      {
        numeric: true
      }
    );
  }

  const typeA =
    matchA[1].toUpperCase();

  const typeB =
    matchB[1].toUpperCase();

  if (
    typeOrder[typeA] !==
    typeOrder[typeB]
  ) {
    return (
      typeOrder[typeA] -
      typeOrder[typeB]
    );
  }

  return (
    Number(matchA[2]) -
    Number(matchB[2])
  );
}


/**
 * セッション内の発表順を並べ替える
 *
 * 優先順位：
 *   1. セッション内のスロット番号
 *   2. スロット内順番
 *
 * 該当列が存在しない場合は、
 * 元データの並び順を維持する
 */
function sortSessionRows(
  rows,
  slotNumberColumnIndex,
  presentationOrderColumnIndex
) {
  if (
    slotNumberColumnIndex === -1 &&
    presentationOrderColumnIndex === -1
  ) {
    return;
  }

  rows.sort(function (rowA, rowB) {
    let slotA =
      Number.MAX_SAFE_INTEGER;

    let slotB =
      Number.MAX_SAFE_INTEGER;

    let orderA =
      Number.MAX_SAFE_INTEGER;

    let orderB =
      Number.MAX_SAFE_INTEGER;

    if (
      slotNumberColumnIndex !== -1
    ) {
      slotA = toSortableNumber(
        rowA[
          slotNumberColumnIndex
        ]
      );

      slotB = toSortableNumber(
        rowB[
          slotNumberColumnIndex
        ]
      );
    }

    if (
      presentationOrderColumnIndex !==
      -1
    ) {
      orderA = toSortableNumber(
        rowA[
          presentationOrderColumnIndex
        ]
      );

      orderB = toSortableNumber(
        rowB[
          presentationOrderColumnIndex
        ]
      );
    }

    if (slotA !== slotB) {
      return slotA - slotB;
    }

    return orderA - orderB;
  });
}


/**
 * ソート用に値を数値化する
 *
 * 空欄や数値でない値は最後に並べる
 */
function toSortableNumber(value) {
  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ""
  ) {
    return Number.MAX_SAFE_INTEGER;
  }

  const number = Number(value);

  if (Number.isFinite(number)) {
    return number;
  }

  return Number.MAX_SAFE_INTEGER;
}


/**
 * 必須列が存在するか確認する
 */
function validateRequiredHeaders(
  headers,
  requiredHeaders
) {
  const missingHeaders =
    requiredHeaders.filter(
      function (headerName) {
        return (
          headers.indexOf(
            headerName
          ) === -1
        );
      }
    );

  if (
    missingHeaders.length > 0
  ) {
    throw new Error(
      "次の必須列が元データにありません。\n" +
      missingHeaders.join("\n")
    );
  }
}


/**
 * 同一セッション内の
 * 「セッション名」が統一されているか確認する
 */
function validateSessionName(
  sessionId,
  sessionRows,
  sessionNameColumnIndex
) {
  if (sessionNameColumnIndex === -1) {
    return;
  }

  const values = sessionRows
    .map(function (row) {
      const value =
        row[sessionNameColumnIndex];

      if (
        value === null ||
        value === undefined
      ) {
        return "";
      }

      return String(value).trim();
    })
    .filter(function (value) {
      return value !== "";
    });

  const uniqueValues =
    Array.from(new Set(values));

  if (uniqueValues.length === 0) {
    throw new Error(
      "セッション " +
      sessionId +
      " の「セッション名」が空欄です。"
    );
  }

  if (uniqueValues.length > 1) {
    throw new Error(
      "セッション " +
      sessionId +
      " の「セッション名」が統一されていません。\n" +
      "入力値: " +
      uniqueValues.join(", ")
    );
  }
}


/**
 * 同一セッション内の
 * 「必要スロット数」が統一されているか確認する
 */
function validateRequiredSlotCount(
  sessionId,
  sessionRows,
  requiredSlotCountColumnIndex
) {
  if (
    requiredSlotCountColumnIndex === -1
  ) {
    return;
  }

  const values = sessionRows
    .map(function (row) {
      const value =
        row[
          requiredSlotCountColumnIndex
        ];

      if (
        value === null ||
        value === undefined
      ) {
        return "";
      }

      return String(value).trim();
    })
    .filter(function (value) {
      return value !== "";
    });

  const uniqueValues =
    Array.from(new Set(values));

  if (uniqueValues.length === 0) {
    throw new Error(
      "セッション " +
      sessionId +
      " の「必要スロット数」が空欄です。"
    );
  }

  if (uniqueValues.length > 1) {
    throw new Error(
      "セッション " +
      sessionId +
      " の「必要スロット数」が統一されていません。\n" +
      "入力値: " +
      uniqueValues.join(", ")
    );
  }
}
