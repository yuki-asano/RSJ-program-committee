const CONFIG = {
  // 元データのシート名
  PAPER_SHEET_NAME: "paper_20260608-1333",

  // セッションシートのテンプレート名
  TEMPLATE_SHEET_NAME: "session_template",

  // 元データのヘッダー行
  HEADER_ROW: 1,

  // 元データの開始行
  DATA_START_ROW: 2,

  // 生成シートへの発表データ書き込み開始行
  // session_templateの25行目から生成データ
  WRITE_SHEET_BASE_ROW: 25,

  // テンプレート左上の置換範囲
  // B1～B4
  TEMPLATE_INFO_START_ROW: 2,
  TEMPLATE_INFO_COL: 2,
  TEMPLATE_INFO_NUM_ROWS: 4,

  // 生成先シートがすでに存在する場合
  // true：削除して再生成
  // false：既存シートをスキップ
  OVERWRITE_EXISTING_SHEETS: false,

  // 対象とするセッションID
  // 例：OS01、IS01、GS01
  SESSION_ID_PATTERN: /^(OS|IS|GS)\d+$/i
};
