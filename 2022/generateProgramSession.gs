// RSJ program auto generator
// 
// Author: 
//    Yuki Asano*1, Kohei Kimura*2, Natsuki Miyata*3 and Kei Okada*1
//      *1 東大, *2 電通大, *3 産総研
//
// Log: 
//   - 2022.8.2  avoid the error of 30 min limit of run-time
//   - 2022.7.21 fixed version for RSJ2022
//   - 2022.6.22 initial program for RSJ2022
//
// Usage:
//   - 「paper_...」のシートで発表リストを作成する
//     - ソート用の列(現在はGY列)してリストを発表順に並べる必要がある（スクリプトはリストを上から順に出力していくため）
//   - 「Custom Menu」
//     - 「セッションシート作成」: 発表リストに基づいて各セッションのシートを自動生成(generateSessionSheets)
//     - 「自動生成シート削除」　： すでに生成したシートを削除(removeGeneratedSheets)


//作成したシートを削除
function removeGeneratedSheets(){
  Logger.log("removeGeneratedSheets()")
  let sheets = SpreadsheetApp.getActive().getSheets()
  for(var i in sheets){
    if(sheets[i].getName().indexOf("OS_") != -1 ||
       sheets[i].getName().indexOf("IS_") != -1 ||
       sheets[i].getName().indexOf("GS_") != -1
      ) {
      Logger.log("remove sheet:" + sheets[i].getName())
      SpreadsheetApp.getActive().deleteSheet(sheets[i])
    }
  }
}

// main function
function generateSessionSheets(){
  Logger.log("generateSessionSheets()")
  
  // init
  // 08/02 09:00 k-okada, comment out
  //removeGeneratedSheets()

  // ハイパーパラメータ
  const SESSION_NUM = 66; // 全部で66.
  const WRITE_SHEET_BASE_ROW = 8;  // 生成シートの書き込み初めの行
  const OS_START_NUM = 1;  // オーガナイズドセッション
  const IS_START_NUM = 22; // 国際セッション
  const GS_START_NUM = 25; // 一般セッション
  const SESSION_NAME_ROW_IN_WRITE_SHEET = 8;
  const SESSION_NAME_COL_IN_WRITE_SHEET = 6;
  let sessionBaseRow = 2;

  // シートを設定
  let paperSheet = SpreadsheetApp.getActive().getSheetByName("paper_20220617-0936")  // 発表者リスト
  let paperSheetLastRow = paperSheet.getLastRow(); // -> 748
  Logger.log("paperSheetLaswRow:" + paperSheetLastRow)
  let templateSheet = SpreadsheetApp.getActive().getSheetByName("session_template")  // テンプレート

  // 発表者リストの情報取得
  let paperSheetIndexNames = paperSheet.getRange(1, 1, 1, paperSheet.getLastColumn()).getValues()[0];  // 発表者リストの(1,1)セルから，1行目全てを取得
  // 以下，シートでのセルのカウントは1始まり，indexは0始まりなので，+1してセルに合わせる
  let paperSheetIndex_paperNo          = paperSheetIndexNames.indexOf("No.") + 1; 
  let paperSheetIndex_sessionType      = paperSheetIndexNames.indexOf("セッションの選択") + 1;
  let paperSheetIndex_sessionNumber    = paperSheetIndexNames.indexOf("セッション番号") + 1;
  let paperSheetIndex_sessionName      = paperSheetIndexNames.indexOf("セッション名") + 1;
  let paperSheetIndex_paperTitle       = paperSheetIndexNames.indexOf("演題名(日本語)") + 1;
  let paperSheetIndex_paperTitleSub    = paperSheetIndexNames.indexOf("論文副題") + 1;
  let paperSheetIndex_paperKeyword1    = paperSheetIndexNames.indexOf("キーワード日本語1") + 1;
  let paperSheetIndex_paperKeyword2    = paperSheetIndexNames.indexOf("キーワード日本語2") + 1;
  let paperSheetIndex_paperKeyword3    = paperSheetIndexNames.indexOf("キーワード日本語3") + 1;
  let paperSheetIndex_paperAbst        = paperSheetIndexNames.indexOf("講演概要文") + 1;
  let paperSheetIndex_paperRegistrant  = paperSheetIndexNames.indexOf("登録者名(氏名)") + 1;
  let paperSheetIndex_paperInstitution = paperSheetIndexNames.indexOf("所属機関 (大学 / 勤務先)") + 1;
  let paperSheetIndex_paperEmail       = paperSheetIndexNames.indexOf("E-mail") + 1;
  let paperSheetIndex_slotNum          = paperSheetIndexNames.indexOf("セッション内のスロット番号") + 1;
  let paperSheetIndex_presenOrder      = paperSheetIndexNames.indexOf("スロット内順番") + 1;

  // 書き込みテンプレートのindex順に並べる
  // スロット番号 発表順	講演No.	セッション種類	セッション番号	セッション名	演題名(日本語)	論文副題	キーワード日本語1	キーワード日本語2	キーワード日本語3	講演概要文	登録者名(氏名)	所属機関 (大学 / 勤務先)	E-mail
  let writeContentsIndex = [
    paperSheetIndex_slotNum,
    paperSheetIndex_presenOrder, 
    paperSheetIndex_paperNo,
    paperSheetIndex_sessionType,
    paperSheetIndex_sessionNumber,
    paperSheetIndex_sessionName,
    paperSheetIndex_paperTitle,
    paperSheetIndex_paperTitleSub,
    paperSheetIndex_paperKeyword1,
    paperSheetIndex_paperKeyword2,
    paperSheetIndex_paperKeyword3,
    paperSheetIndex_paperAbst,
    paperSheetIndex_paperRegistrant,
    paperSheetIndex_paperInstitution,
    paperSheetIndex_paperEmail
  ]; 
  Logger.log("paperSheetIndexNames:" + paperSheetIndexNames)
  Logger.log("writeContentsIndex:" + writeContentsIndex)

  // セッションごとに書き出し
  for(var i = 1; i <= SESSION_NUM; i++){
    // i: 現在の作業セッション
    let currentSession = i;

    if(OS_START_NUM <= currentSession && currentSession < IS_START_NUM){
      var sessionType = "OS";
      var sessionNumForName = currentSession;
    }else if(IS_START_NUM <= currentSession && currentSession < GS_START_NUM){
      var sessionType = "IS";
      var sessionNumForName = currentSession - IS_START_NUM + 1;
    }else{
      var sessionType = "GS";
      var sessionNumForName = currentSession - GS_START_NUM + 1;
    }
    
    // セッション内の発表者人数の取得
    let sessionRange = paperSheet.getRange(1, paperSheetIndex_sessionNumber, paperSheetLastRow);  // 「セッション番号」の列を検索. 
    let finder =  sessionRange.createTextFinder(currentSession).matchEntireCell(true)  // [完全に一致するセルを検索]を有効．
    let presenNum = finder.findAll().length;    // セッション内の発表者数をカウント
    Logger.log("presenNum:" + presenNum)

    // 08/02 9:00 k-okada 既にシートが作成されていればスキップする
    if (SpreadsheetApp.getActive().getSheetByName(sessionType + "_" + sessionNumForName.toString())) {
        Logger.log(sessionType + "_" + sessionNumForName.toString() + " is already generated")
        sessionBaseRow = sessionBaseRow + presenNum
        continue
    }

    if(presenNum != 0){  // skip if empty session
      // セッションごとのシート作成
      let writeSheetName = sessionType + "_" + sessionNumForName.toString()
      templateSheet.copyTo(SpreadsheetApp.getActive()).setName(writeSheetName)  // templateをコピー

      // 発表者リストの書き込み
      let writeSheet = SpreadsheetApp.getActive().getSheetByName(writeSheetName);

      for(var j = 0; j < writeContentsIndex.length; j++){
        // j: セッション内　作業のindex
        let writeValues = paperSheet.getRange(sessionBaseRow,writeContentsIndex[j], presenNum).getValues();  // 書き込む内容として，発表リストの(sessionBaseRow, index)セルから，presenNum行 のセル内容を取得
        let writeRange = writeSheet.getRange(WRITE_SHEET_BASE_ROW,j+1, presenNum);  // 書き込む範囲として，writeSheetの(WRITE_SHEET_BASE_ROW, 1)セルから，presenNum行 のセルを選択
        writeRange.setValues(writeValues);
        Logger.log("index:" + j)
        Logger.log("writeValues:" + writeValues)
      }

      // セッション情報の取得
      let sessionID   = sessionType + sessionNumForName.toString();
      let sessionName = writeSheet.getRange(SESSION_NAME_ROW_IN_WRITE_SHEET, SESSION_NAME_COL_IN_WRITE_SHEET).getValue();
      let sessionRoom = currentSession + 100;         // 「セッション室」を取得
      Logger.log("sessionName:" + sessionName);
      Logger.log("sessionRoom:" + sessionRoom);

      // セッション情報の書き込み
      let replaceRange = writeSheet.getRange(1,2, 5); // 書き込みシートのテンプレート部分を取得．(1,2)セルから, 5行．
      let replaceData = replaceRange.getValues();
      Logger.log(replaceData);

      replaceData[0][0] = replaceData[0][0].replace("<<セッションID>>", sessionID);
      replaceData[1][0] = replaceData[1][0].replace("<<セッション名>>", sessionName);
      replaceData[2][0] = replaceData[2][0].replace("<<講演数>>", presenNum);
      // replaceData[3][0] = replaceData[3][0].replace("<<セッション室>>", sessionRoom);
      // replaceData[4][0] = replaceData[4][0].replace("<<座長>>", sessionRoom);
      replaceRange.setValues(replaceData)
      
      // update
      sessionBaseRow = sessionBaseRow + presenNum

    }else{
      Logger.log("Session " + currentSession + " is empty")
    }
  }
} 
