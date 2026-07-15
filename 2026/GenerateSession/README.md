# RSJ2026 セッションシート自動生成

講演申込データの `セッションID` ごとに、OS・IS・GSのセッションシートを自動生成するGoogle Apps Scriptです。

## 準備
- スプレッドシートを用意し、次のシートを作成します。
  ```text
  paper_...
  session_template
  ```
  - `paper_...`：講演申込の元データ．年度ごとにご用意ください．
  - `session_template`：生成シートのひな形 ([session_template](https://docs.google.com/spreadsheets/d/1F7LxEd_s-H5iLM48tggQLKW44PLtHgbKfyCfkguCudY/edit?usp=sharing))

- シートの「拡張機能」「Apps Script」に，プログラムをコピーしてください．(GenerateSession.gs, Menu.gs, Config.gs)
- Config.gsの`PAPER_SHEET_NAME` は、実際の元データシート名に変更してください。
- コピー後は、スプレッドシートを再読み込みすると、スクリプトの実行メニューが現れます。
- 初回実行時に権限を承認する必要があります．

## 実行方法
1. 用意したシートを開きます。
2. 上部メニューの **セッション自動生成** を開きます。
3. **セッションシート自動生成** を実行します。

生成されるシート名の例：

```text
OS_01
IS_01
GS_01
```

## 生成シートの削除

**セッション自動生成** メニューから、次の操作を実行できます。

- **自動生成シート全削除**
- **GSシート削除**
- **ISシート削除**
- **OSシート削除**

## 再生成する場合

既存の生成シートがある場合は、対象シートを削除してから再度 **セッションシート自動生成** を実行します。

`Config.gs` で次のように設定すると、既存シートを自動的に削除して再生成できます。

```javascript
OVERWRITE_EXISTING_SHEETS: true
```

既存シートを残してスキップする場合は、次のように設定します。

```javascript
OVERWRITE_EXISTING_SHEETS: false
```

## 注意事項

- `セッションID` は `OS01`、`IS01`、`GS01` の形式にしてください。
- 同じセッションIDでは、`セッション名` と `必要スロット数` を統一してください。
- スクリプトを修正した場合は、生成済みシートを削除してから再実行してください。

## 謝辞
本プログラムの開発、デバッグ、およびドキュメント作成の支援にChatGPTを使用しました。
