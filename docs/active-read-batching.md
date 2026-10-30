# アクティブシート取得停止：構造調査とバッチ化

> 本文は調査時点の経緯・判断の記録です。「未コミット」「今回はコミットしない」「未検証」等は当時の状態を示します。最新の実機確認範囲と残課題は[安定化作業の到達点](stabilization-status.md)を参照してください。A/B/C独立比較自体は引き続き実機未検証です。

基準HEAD：`7aa14579d171651d37444ddd9a456cbb6450b9ac`。前回までの未コミット修正を保持し、その上で追加実装。コミットしない。Sheaside、リボン、manifest、依存パッケージは変更しない。

## 事実と未特定事項

Excel for the webで、シート切替のみは体感約1秒以内で追従した。一方、名前編集を開始してEscで取消すと、名前変更イベント0回のままアクティブID取得のcontext.syncが未完了となった。metadata取得は完了し、総シート数評価へは進んでいない。今回のログにはRichApi.Errorなし。編集開始/Escの時刻は記録されていないため、API開始との厳密な前後関係は分からない。

**停止経路は特定済みだが、Officeホストが応答を返さない根本原因は未特定。今回のバッチ化で停止を解消したかも未検証。**

公式のdelayForCellEditはセル編集を対象にする。trueでは編集終了まで待機、false（既定）では編集中の対象処理がエラーになると説明されている。SheetDockでは指定していない。この仕様をシート名編集へそのまま適用したり、Esc後の長時間停止を仕様通りとみなしたりはできない。[公式説明](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-delay-in-cell-edit)

名前編集UIとホストのAPI処理が重なった際の不具合・応答配送の停滞などは仮説。参照した公式資料にはこの再現手順の直接原因や確定した回避策は見当たらなかった。Web版だけで観測しており、Windows/Mac未検証なのでWeb固有と断定しない。

## 変更前の構造と分離理由

通常UIはExcelロック内の1回のExcel.runで次の3syncを逐次実行していた。

1. ブック名/readOnly/構成保護、Worksheet一覧のID/名前/位置/表示状態/保護。
2. getActiveWorksheet().load("id")。
3. 引数なしfunctions.sheets()のvalue/error（全種類の総シート数）。

分離した理由はエラー時の情報保持。独立グラフシートなどでactive取得が失敗しても、すでに取得した通常Worksheet一覧を返せるようにしていた。SHEETS失敗時も一覧を保持し、総数不明として変更操作を拒否する。この分離には意味があるが、1と2のロードにはデータ依存がなく、通常時に別の往復を必須とする理由はない。

ただし旧コードでも、active syncがrejectせず未完了のままならloadWorkbookState全体が返らないため、先に成功した一覧をUIへ返せていたわけではない。

## 修正案の比較

| 案 | 実現性・API削減 | 安全性・既存機能への影響 | 判断 |
|---|---|---|---|
| A：一覧とactiveを同一バッチ | 同じContextでloadを積んでsync可能。通常3sync→2sync | getActiveWorksheetの権威的な取得を維持。片方のエラーで統合バッチ全体が失敗するため、別Contextで一覧のみ再取得が必要 | 今回採用。性能改善の根拠あり、停止回避効果は未確認 |
| Aの拡張：総数も含めて1sync | APIモデル上はまとめられる | 任意のSHEETS評価エラーで一覧まで影響。取得失敗の分離がさらに難しくなる | 今回見送り。総数は別取得で保守的に不明扱い |
| B：getSelectedRange().worksheet等 | 別の経路でWorksheetへ到達できる場合はある | 選択の種類・複数範囲・グラフシート等で別の制約を増やす。activeを読む経路の停止を避ける根拠は未確認 | 今回採用しない |
| B：一覧のフラグから判定 | ロード済みの一覧にOfficeのactiveフラグは取得していない。SheetInfo.isActiveはアプリが導出した値 | 自前isActiveを根拠として再利用すると過去の結果で自己確認することになる。getItemOrNullObjectは既知IDの存在確認でありactive発見ではない | 代用不可 |
| C：onActivatedのworksheetIdを利用 | イベント時点のIDを追加syncなしで取得可能（ExcelApi 1.7） | 初期値、登録前/失敗中の変更、ブック世代、未知ID、非対応シートへの遷移、古いイベントを扱う必要がある。表示の暫定更新には候補だが変更直前の再検証は代用できない | 将来の別検証候補。今回は根拠が弱い状態を確定表示しない |
| D：取得済み一覧のみ表示 | active不明でも一覧閲覧は可能 | 全行isActive=false、変更操作は禁止とする。古いactiveを保持しない | A失敗時の限定的な退避として採用 |
| D：待機時間を延ばす/編集待機をON | 呼出時刻をずらせるが停止を防ぐ証拠なし | 名前編集開始はアドインiframe内のイベントではなく、正確に検出できていない。delayForCellEdit=trueは待ちを延ばす可能性 | 今回採用しない |
| D：期限切れでロックを外す | 後続を走らせるだけなら可能 | 旧APIが取消されず並列化し、順序や変更安全性を壊す | 不採用 |

getActiveWorksheetはExcelApi 1.1。取得したIDが同じバッチの通常Worksheet一覧に存在する場合だけactiveとして使う。参照したWorksheetCollection API一覧にはgetActiveWorksheetOrNullObjectはなく、未定義メソッドは導入していない。[WorksheetCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview)

onActivatedにはworksheetIdがあるが、イベント時点の情報を現在状態として永久に扱わない。[WorksheetActivatedEventArgs](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetactivatedeventargs?view=excel-js-preview)

## 今回の実装

通常経路は以下へ変更した。

1. metadataとactive IDのloadを同じContextへ積む → `api.sync.metadataActive`を1回。
2. 総シート数評価 → `api.sync.total`を1回。

結果を読むのはsync後。**通常は2sync（従来3sync）**。Microsoftはloadをまとめてsyncを減らすことを推奨しており、Webでは往復削減に意味がある。これは停止がなくなるという保証でも、必ず33%高速になるという主張でもない。[性能最適化](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/performance)、[syncの往復について](https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/correlated-objects-pattern)

統合取得が**エラーを返して完了した場合だけ**、通常UIの読み取りラッパーは新しいExcel.run/RequestContextで一覧のみを1回再取得する。失敗したContextのプロキシは再利用しない。退避側でもSHEETSを確認し、activeId=null、全行isActive=falseとする。画面に取得失敗と一覧のみである旨を表示し、移動を許可しない。再取得も失敗したら例外をUIへ伝え、それ以上その要求内で再試行しない。

- APIが未応答の間は退避を開始しない。15秒監視・ロック保持・限定回復の既存処理を維持。
- キャンセル/破棄/世代不一致なら退避取得を送らない。総数取得完了後にも中断を確認。
- 通常表示・単一/検索中移動は同じ情報を参照。変更直前のloadWorkbookStateは厳密な統合取得を使い、失敗したら移動しない。読み取り退避を変更操作の承認に使わない。
- 非対応シート混在、readOnly、構成保護、対象保護、順序/状態差異、active不明のガードは維持。
- 各バッチは完全な排他トランザクションではない。取得中イベントによる世代破棄も維持。
- 開発PoCの独立観測は、比較用として従来の分離取得のまま残す。今回の実機比較はPoCイベント観測OFFで行う。

## 変更ファイル（今回追加分）

- `src/excel/workbookState.ts`：統合load/2sync、新Contextによる一覧退避、最後の中断判定。
- `src/domain/workbookState.ts`：読み取り警告readIssue（任意）。安全性用の構成比較には含めない。
- `src/state/useWorksheets.ts`：一覧退避の警告を画面へ伝え、正常取得で解消。
- `tests/workbookState.test.ts`、`tests/useWorksheets.test.tsx`：回帰テスト。
- 本書、README、testing.md、継続調査記録：調査・再検証手順。

## 実機再検証

1. 既存開発サーバーを起動してExcelのブラウザータブ全体を再読み込み。検証ブック、PoCイベント観測OFF。
2. 初期一覧とactive表示を確認し、シート切替だけを試す。診断に`api.sync.metadataActive.start/end`が現れることを確認（通常取得の旧`api.sync.active`はなくなる）。
3. 一覧反映後、前回止まった「名前編集開始 → 約2秒後にEsc」を1回試す。停止したらそこで打ち切る。
4. 正常なら名前変更確定を1回、続いて5～10回。最終名へ追いつくか確認。
5. 追加/削除、ペイン再表示、通常単一移動、検索中単一移動を確認。
6. グラフシートを含む既存の専用コピーがあれば、一覧のみの取得やactive不明時の変更禁止を確認。グラフシートを作り替えない。

DevToolsでlocalhost:3000のフレームを選び、異常直後に以下を保存する。

```javascript
copy(JSON.stringify(window.sheetDockDiagnostics.snapshot(), null, 2))
```

新しい段階：`api.sync.metadataActive`（統合取得）、`normal.read.metadataFallback`（失敗後の退避開始）、`api.run.readMetadataOnly`/`api.sync.metadataOnly`（新Contextの一覧取得）。総数は従来の`api.sync.total`。

統合バッチで止まった場合、ログからmetadataとactiveのどちらのホスト処理かは区別できなくなる。旧ログと合わせ「バッチ化しても未解決」と評価し、別のAPIで並行取得して回避しない。初期ロード時のエラー退避が成功しても、元のactive取得が正常化した意味ではない。

## 未解決

Esc後に応答が戻らないホスト内部の理由、Web限定かどうか、バッチ化の実機効果は不明。既知の制約として扱えるのは公式に説明されたセル編集時の挙動までであり、今回の停止と同一視しない。実機再検証前のコミットは行わない。
