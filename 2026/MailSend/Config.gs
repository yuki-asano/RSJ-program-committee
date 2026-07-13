const CONFIG = {
  spreadsheetUrl: 'write_url_here',

  sheetGid: 0,

  // 必ず実際のURLに差し替えてください
  answerFormUrl: 'https://www.google.com/',
  programUrl: 'https://www.google.com/',
  // 回答〆切
  deadlineChair1st: '7/17',

  senderName: 'RSJ2026プログラム委員会',
  replyTo: 'write_email_address_here',

  // 1回の実行で送る最大メール数
  // 1行につき最大2通送るので、30にすると最大15行分です。
  maxEmailsPerRun: 30,

  sleepMs: 500,

  // false 推奨。
  // true にすると「送信する」チェックなしで全行対象になります。
  sendAllRowsIgnoringFlag: false
};
