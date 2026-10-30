# シート名変更後の遅延・RichApi.Error調査

> 本文は調査時点の経緯・判断の記録です。「未コミット」「今回はコミットしない」「未検証」等は当時の状態を示します。最新の実機確認範囲と残課題は[安定化作業の到達点](stabilization-status.md)を参照してください。A/B/C独立比較自体は引き続き実機未検証です。

調査日：2026-10-28。基準HEAD：`46dee0351dabb0e91033b7f23679e5bae695d0d6`。開始時作業ツリーはクリーン。Sheasideは変更していない。

## 結論と証拠の限界

**報告されたRichApi.Errorの直接原因は未特定。実Excelでの解消確認は未実施。** スクリーンショットはOffice.js内部の応答処理を示しており、SheetDockの呼び出し箇所、error.code、debugInfo.errorLocationまでは分からない。

一方、独立したコード上の問題は再現できた。以下の修正前テストは2件とも失敗し、修正後は成功した。

1. 初回の状態取得を未完了にしたまま名前変更イベントを2回発火すると、イベントの150ms待機後に2個目の取得が開始された。初回・手動更新がイベントの集約キューを通らず直接refreshを呼んでいたため。実コードのExcelロックはAPIを直列化しているので、これを同時Excel.runの証拠とは扱わない。重複する取得要求・待ち行列を作る経路の証拠である。
2. PoCで自動同期を有効化し、Office.onReady待機中にコンポーネントを破棄すると、解除要求の後にsetEnabled(true)が呼ばれた。toggleAutoがawait後に生存状態を確認していなかったため。ユーザーがPoCのスイッチを有効にしていなければ、この経路は該当しない。

これらが今回の実機症状の原因であると結び付けるログはない。推測による修正コミットは作成せず、修正候補・診断コード・テスト・本書を未コミットで残した。

## 調査した経路

`WorksheetCollection.onNameChanged → SharedEvents → 通常UIのCoalescedRefresh → readWorkbookState → Excel.run → metadata/active/SHEETSのsync → React状態更新`。

PoCイベント観測ONの場合は同じ物理登録からPoCへも通知され、observeBookが別に取得する。旧コードは観測の完了をsnapshotBusへ通知していたため、通常UIにもさらに再取得要求が入った。再帰する循環は見つからないが、1イベントにつき通常取得3sync＋PoC取得3sync＋追加の通常取得3syncになり得る経路があった。実際の回数は集約タイミング次第で、実機で9回発生したと断定しない。

- 7種類のWorksheetイベントとWorkbook.onActivatedの登録を共有する構造、最後の利用者で解除する処理は維持した。
- Worksheet名をイベントハンドラーから再度書き換える処理はない。名前変更を繰り返す再帰ループはコード上確認できない。
- Worksheetプロキシは取得処理内で使い、UIへ渡すのは通常のデータ。異なるRequestContextでWorksheetプロキシを再利用する箇所は確認できない。
- イベント解除は登録時と同じRequestContextで実行しており、公式例と一致。根拠なく別Contextへ変更していない。
- Workbook.onActivatedは別ブックへの移動だけでなくフォーカス復帰等でも発火し得る。再登録回数は診断で観測する。今回、必要なブック再登録を停止する変更はしていない。
- 検索中移動の変更は表示境界→全体境界の変換が中心で、onNameChangedの登録方法を変更していない。PoCには総シート数取得が増えていた。どちらも今回の実機エラーの原因と断定しない。
- React再描画時間/ホスト側負荷は実測できていない。API待機時間とUI負荷を同一視しない。

## 修正候補

- 初回・手動・イベント更新をCoalescedRefreshへ統一。処理中の複数要求は後続1回へ持ち越す。自動同期は維持。
- 通常一覧の取得は、Excelロック待ちの後と取得バッチの間に生存状態を確認。破棄済みなら追加のAPIを送らない。すでに送信済みのsyncを取り消せるとは扱わない。
- PoCの自動同期切替でOffice初期化待ち後・登録完了後の破棄を確認。終了後の新規登録を防ぐ。
- PoCの読み取り専用観測から通常UIへの追加取得通知を除去。PoCの変更操作後の通知は維持する。
- 再登録待ち中の新しいWorkbookイベントを古い完了処理で消さない。
- 通常ドラッグ、検索中ドラッグ、変更直前の再取得、保護/読み取り専用/混在の判定、失敗後再取得は維持。

イベント解除失敗時に重複登録を避ける既存の状態不明扱いを維持。解除を成功したことにしたり、自動で何度も登録し直したりしない。エラーの抑制やwebpackのオーバーレイ無効化も行っていない。

ブラウザーのアンロード時にReactのcleanupや非同期の解除完了が必ず実行されるとは保証しない。公式ではアドインのリフレッシュ/終了時にイベントは破棄される。実ホストがペインを非表示にするだけかランタイムも破棄するかは診断ログで区別する必要がある。

## 開発専用診断

`window.sheetDockDiagnostics`を開発ビルドだけに公開する。ログは直近500行、種別ごとの累積回数を保持。外部送信・ファイル保存はしない。取得したシート名・ブック名・ID・URL・セル内容・イベント引数・生エラーメッセージ・スタックは記録しない。Errorのcodeと許可形式のAPI位置のみ記録する。

| 記録 | 意味 |
|---|---|
| event.onNameChanged等 | 物理イベントの通知回数。通常UI/PoCで二重加算しない |
| normal.refresh.request / refresh.request | 通常UI / デフォルト名の更新要求 |
| normal.refresh.start/end/error | 通常取得処理の実行回数・開始終了・所要時間 |
| poc.refresh.* | PoCイベント更新（手動観測はこの集約処理とは別） |
| api.run.read.* | 通常状態取得のExcel.run |
| api.run.register/unregister.* | イベント登録・解除のExcel.run |
| api.run.poc.metadata/active/total.* | PoCの3回の観測Excel.run |
| api.sync.metadata/active/total.* | 通常取得の各sync |
| api.sync.register/unregister/registrationCleanup.* | 登録・解除のsync |
| api.sync.poc.* | PoC観測/変更内のsync |
| api.queue.request / api.queue.wait | アプリ内Excelロックへの要求と待機時間 |
| api.lock.start/end/error | ロック内処理の回数と所要時間。Excel.runの回数とは別 |
| pane.init/dispose、poc.init/dispose | Reactの初期化・破棄 |
| runtime.pagehide/pageshow | ブラウザーが通知した表示ライフサイクル |
| runtime.error/unhandledrejection | 未処理エラーの観測。preventDefaultは呼ばない |

api.run.*は今回の同期調査対象を計測する。Office.js内部の通信や、移動/切替など別経路のExcel.runまで含む総数ではない。sync.endは成功/失敗いずれでも記録され、失敗時はerror行も記録される。ログ出力の回数から再帰ループがあると即断しない。

## Excel for the webでの再検証

1. SheetDockディレクトリで既存環境の`npm.cmd run dev`を起動。NodeがPATHにない場合は次を利用する。

   ```powershell
   Set-Location .\SheetDock  # 配置先の親ディレクトリから実行
   $env:PATH = "$PWD\.local-tools\node-v24.22.1-win-x64;" + $env:PATH
   npm.cmd run dev
   ```

2. 使い捨ての検証ブックに `対象1、外A、対象2` を用意しSheetDockを開く。manifestは変更していない。古いエラー状態を引き継がない比較のため、最初の1回はExcelのブラウザータブ全体を再読み込みする。
3. ブラウザー開発者ツールのConsoleで、実行対象フレームを`https://localhost:3000/taskpane.html`へ切り替える。`window.sheetDockDiagnostics`があることを確認する。これは本番ビルドやExcel親ページにはない。
4. Consoleの「Preserve log（ログを保持）」をONにし、次を実行する。

   ```javascript
   window.sheetDockDiagnostics.setConsole(true)
   ```

5. 最初はPoCの「イベント自動同期を検証」をOFFのまま、Excel標準タブから対象1の名前を5～10回変更する。入力確定から一覧反映までの体感時間、CPU負荷感、操作不能になったかを記録する。イベント数と取得数が際限なく増え続けないか見る。
6. 待機後、次で名前変更後のログをコピーして保存する。copyはブラウザーDevToolsの補助関数。

   ```javascript
   copy(JSON.stringify(window.sheetDockDiagnostics.snapshot(), null, 2))
   ```

   copyが使えない場合は`JSON.stringify(...)`の結果を手動コピーする。Console全体には他のブック情報が出る可能性があるため、共有するのはSheetDock診断行と必要なエラー部分だけにする。

7. タスクペインを閉じ、再度開く。エラーの有無と一覧を確認。フレームを再選択し、再びsnapshotを保存する。ランタイムが再作成された場合、メモリー内ログはリセットされるため閉じる前後の2つが必要。Preserve logを有効にしていても環境によって破棄フレームのログが残らない場合がある。
8. シート追加・削除・切替を確認。`対象`で検索し、対象2を検索結果先頭へ移動して `対象2 → 対象1 → 外A` になることと検索条件維持を確認する。対象1を改名した場合は先に元の名前へ戻す。
9. 同じ検証をPoCイベント観測ONでも行い、OFF時と比較する。PoCの変更操作の有効化や削除・移動計画の実行は、この不具合調査には不要。
10. 可能なら「起動直後に改名」「手動更新直後に改名」「PoCをONにしてすぐペインを閉じる」も試す。失敗したら連打を止め、ログを保存する。

報告に必要なもの：PoC観測ON/OFF、操作手順、名前変更後/再表示後の2つの診断snapshot、RichApi.Errorのcode/発生時刻。シートやブック内容そのものは不要。

## 自動検証

- 既存197件を含む209テスト、17ファイルが成功。12件追加。
- 修正前に、初回取得の重複要求・PoC終了後登録の2回帰テストが失敗することを確認した。
- 追加対象：初回とイベントの集約、未完了処理の終了、再表示時の登録、再登録中の新イベント、ロック待ち中のキャンセル、取得バッチ間のキャンセル、PoC観測の余分な通知なし、診断情報の機密除外/本番無効化。
- 既存の連続イベント、登録/解除共有、API失敗後復旧、ドラッグ、検索境界、移動安全性のテストも成功。
- TypeScript型チェック成功。
- Development/Productionビルド成功、両方エラー0・警告0。
- マニフェスト検証成功。
- ProductionにPoCモジュールなし。診断用window API・グローバルエラー監視もProduction成果物に含まれない。

テストはモック/React/純粋処理であり、実ホストのRequestContext無効化や今回のRichApi.Errorを再現したものではない。

## 変更ファイル

- `src/diagnostics.ts`：診断リングバッファ、計測、開発用Console API。
- `src/sync/coalescedRefresh.ts`、`src/state/useWorksheets.ts`：初回/手動/自動取得の集約と再登録競合。
- `src/excel/serial.ts`、`src/excel/workbookState.ts`、`src/excel/worksheetEvents.ts`：API計測・取得キャンセル。
- `src/poc/Panel.tsx`、`src/poc/officeAdapter.ts`：終了後登録の抑止・重複観測通知削減・計測。
- `tests/diagnostics.test.ts`、`tests/sync.test.ts`、`tests/useWorksheets.test.tsx`、`tests/workbookState.test.ts`、`tests/sharedEvents.test.ts`、`tests/poc-confirmation.test.tsx`：回帰テスト。
- 本書、README、docs/testing.md：調査と実機手順。

## 根拠・未解決事項

2026-10-28参照：

- [Microsoft：Excelイベント](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-events)：同じRequestContextで解除。ランタイム終了/リフレッシュでイベントが破棄される。
- [Microsoft：Workbookの操作とイベント](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-workbooks)：Workbookのアクティブ化と非アクティブ化。

実機ログで最初の失敗APIを特定するまでは、ホスト固有のコンテキスト失効、登録/解除時の失敗、SHEETS評価や状態取得の失敗、Office内部の未処理エラーを切り分けられない。閉じる/開くの動作、長時間の安定性、共同編集、Windows/Macも未確認。今回の修正候補を実機確認し、原因と妥当性を判断してからコミットする。remote追加・push・タグ・公開・インストールはしていない。


## 継続調査

前回修正後も連続改名で遅延するとの実機報告を受け、[自動同期安定化・リボン調査](sync-stability-ribbon.md)を追加。前回の変更は保持し、15秒の待機監視と変更要求世代による古い取得結果の排除を追加した。直接原因は引き続き未特定、実機確認前のコミットは行わない。
