# 自動同期の安定化とリボン開閉調査（調査履歴）

> 本文は調査時点の経緯・判断の記録です。「未コミット」「今回はコミットしない」「未検証」等は当時の状態を示します。最新の実機確認範囲と残課題は[安定化作業の到達点](stabilization-status.md)を参照してください。A/B/C独立比較自体は引き続き実機未検証です。

作業日：2026-10-28。HEAD：`46dee0351dabb0e91033b7f23679e5bae695d0d6`。開始時から存在した前回の未コミット修正を保持して追加修正した。コミット、remote、push、タグ、リボン/manifestの変更、Sheaside変更は実施しない。

## 実機報告と原因の特定状況

ユーザーの追加報告：前回修正後もExcel for the webで連続したシート名変更後に取得中が長く残り、Excel側の「外A」に対し一覧に「外AA」が残る場合がある。一覧未取得となる場合もあり、ペインを閉じて開き直すと回復することがある。

検索中の単一ドラッグはユーザーの観測で正常に動作しているように見えた。今回の実機報告として記録するが、全境界/全環境の検証完了とはしない。複数移動PoC、長時間安定性、RichApi.Errorの直接原因、開閉トグルは未確認。

**Office APIが応答しなくなる直接原因は依然未特定。** 新しい診断ログは未提供であり、どのsyncが停止したかは断定できない。

コードで確認できた経路：

1. `withExcelLock`は実APIのPromise完了まで後続を待たせる。この排他自体は必要だが、未完了のままならキューも完了しない。
2. 初回/手動更新のbusy解除はPromiseのfinallyに依存していた。従ってホストがresolve/rejectしない場合、永久に取得中になる経路がある。今回未応答Promiseを与える回帰テストを追加した。
3. 取得中の通常の名前変更にはブック切替世代が変わらない。旧版ではその取得結果を一旦適用してから次の取得へ進んでいた。今回は変更要求の世代も照合する。
4. 通常の自動イベント更新だけではbusyをtrueにしない。報告された長いスピナーには、初回/手動更新/切替/移動等の経路が絡んでいる可能性がある。名前変更イベントだけでスピナーが出ると断定しない。

前回の修正（取得要求の集約、PoC終了後の再登録防止、観測による追加要求の削減、破棄済み取得の中断）は撤回していない。今回はそれらで扱っていなかった「ホストの未応答」と「取得中の新しい変更要求」を対象にした。

## 修正と応答しないAPIの扱い

- イベント、手動更新、PoC変更通知などで変更要求の世代を更新。取得開始時の世代と違う結果をUIへ適用しない。
- 実行中に追加された更新要求は従来のdirtyフラグで後続1回に集約。連続変更が止まりAPIが応答するなら、最後の状態を再取得する。
- 正常に完了した切替/移動が新しい一覧へ反映済みの場合、先に始まった取得結果で上書きしない。移動前の再検証や混在ガードは維持。
- 通常取得/操作が15秒を超えたら、スピナーを応答待ちエラーへ切り替え、古い一覧を消し、変更操作と追加の手動取得を無効化する。
- **15秒はUI監視の閾値。Office APIを中止できた意味ではない。Promise.raceでロックを解放したり、未完了APIと並列に次のAPIを始めたりしない。**
- 遅れたAPIが終了したら期限切れの取得結果を破棄し、回復取得を1回要求する。新しいイベント/フォーカス/手動要求がなければ自動回復を無制限に繰り返さない。
- 回復取得も遅延した場合、さらに自動で回し続けず、状態未確認のエラーを残す。手動/次のイベントによる再取得は可能。
- ずっと応答しない場合、アドイン内だけで安全にロックを解放・再実行することはできない。ログ保存後、Excelのブラウザータブ全体の再読み込みを案内する。ブラウザーがJSタイマー自体を停止している間は15秒表示も遅れる可能性がある。
- 正常なAPIエラーは表示し、次のイベント/手動取得で回復可能。取消し・期限切れ・古い世代の破棄と、実APIエラーを診断上区別する。

送信済みの変更操作は期限超過でも取消しを保証しない。旧結果を表示しないことと、Excel側の変更を元に戻すことは別。回復後は実状態を再取得する。複数PoC処理に新たな本番機能は追加していない。

## 再表示で回復する理由

**仮説**：現在はShared Runtime未使用のShowTaskpane方式。ホストがペインのJSランタイムを破棄・再作成した場合、モジュール内のロックPromise、購読状態、React状態が新しくなるので、滞留していた状態が持ち越されない。単に再取得のタイミングが合った可能性もある。

再表示が常に再初期化を意味するとは扱わない。新しい診断snapshotに匿名の`session`識別子を追加した。閉じる前後でsessionが違えばJSモジュール再生成の証拠になる。同じなら、非表示からの復帰/再取得による回復として別に調べる。

今後Shared Runtimeでhide/showを導入した場合、非表示でも状態/イベントが保持されるので、「閉じて開けば復旧する」という期待は成立しない。先に同期安定性を検証する必要がある。

## 診断ログ

前回の`window.sheetDockDiagnostics.snapshot()`を継続使用する。時刻、固定の処理名、所要時間、許可したerror.code/API箇所のみ。ブック名・シート名・ID・セル値・URLは含めない。外部送信なし。

- `session`：JS初期化の識別子。個人/ブックIDではない。
- `normal.refresh.request / coalesced`：要求受付回数と実行中/予約中に集約された回数。
- `normal.refresh.applied`：取得結果を採用したこと。
- `normal.refresh.discarded.revision/generation/disposed/expired`：最新要求/ブック世代/破棄/期限超過による結果破棄。
- `normal.refresh.stalled / normal.operation.stalled`：15秒超過。
- `*.lateSettled`：期限超過していた実処理が終了。
- `normal.refresh.recovery`：期限超過後の回復取得要求。
- 前回の`api.run.*`、`api.sync.*`、`api.queue.wait`、`api.lock.*`、`pane.init/dispose`、`runtime.*`も維持。

リングバッファは直近500件、カウンターは累積。Console出力は初期OFFで、ONでも毎秒最大20行に制限する。Console側で省略した分もカウンター/リングバッファの範囲内には残る。API実行回数の対象範囲は[前回記録](rename-sync-investigation.md)を参照。

## Excel for the webでの実機手順

1. 既存の開発サーバーをSheetDockで起動。新規インストールは不要。

   ```powershell
   Set-Location .\SheetDock  # 配置先の親ディレクトリから実行
   $env:PATH = "$PWD\.local-tools\node-v24.22.1-win-x64;" + $env:PATH
   npm.cmd run dev
   ```

2. 比較開始時にExcelのブラウザータブ全体を再読み込みし、使い捨てブックに `対象1 → 外A → 対象2` を準備。SheetDockを表示。最初はPoCイベント観測OFF。
3. 外Aを外AAへ1回変更し、一覧反映時間を記録する。続けて外A/外AA等を5～10回変更し、最後を外Aにする。改名を止めた後、最終名へ追いつくことを確認する。
4. 待機が15秒を超えた場合、無限のスピナーではなく応答待ちの説明と操作無効化になることを確認する。戻らなければ連打せずログを保存。
5. ペインを閉じ、再表示。初期取得、エラー、sessionの変化を確認する。
6. シート追加・削除・アクティブ切替を確認する。削除は使い捨てブックの追加したシートだけをExcel標準UIで行う。
7. `対象`で検索し、対象2を検索先頭へ移動して `対象2 → 対象1 → 外A` になること、検索条件維持、キャンセルを確認する。
8. PoCイベント観測ONでも改名と開閉を比較する。PoCの変更操作は不要。

### 異常時に保存するもの

DevToolsのConsoleで、実行フレームを`https://localhost:3000/taskpane.html`へ切り替える。名前変更後（閉じる前）と再表示後の両方で次を実行する。

```javascript
copy(JSON.stringify(window.sheetDockDiagnostics.snapshot(), null, 2))
```

copyがないブラウザーではJSON.stringifyの結果を手動コピーする。session、counts、entriesをそのまま保存。任意でConsoleのPreserve logをONにし、`window.sheetDockDiagnostics.setConsole(true)`で時系列出力を有効化できる。初期化時のsnapshotはclearせず残す。

報告には「PoC観測ON/OFF」「改名回数」「最後の変更から何秒待ったか」「スピナー/応答待ち表示のどちらか」「閉じる前後のログ」を添える。特に最後にstartが記録されendのないAPI段階を調べる。error行があればコード/位置を参照する。Consoleの他アプリ出力を一括共有する必要はない。

## リボン開閉トグル（調査のみ）

2026-10-28にMicrosoft公式資料を参照。manifest、Commandsページ、リボン実装は変更していない。

| 調査項目 | 判断 |
|---|---|
| 現在のShowTaskpane | 宣言されたペインを開くアクション。任意の開閉判定コードは実行しないので、この設定のままトグルにはできない |
| Office.addin.hide / showAsTaskpane | SharedRuntime 1.1の正式API。非表示/表示でありアンロード/再初期化ではない |
| 推奨候補（サイドロード等） | Shared Runtime＋ExecuteFunctionのコマンドで状態を判定してhide/show。コマンド多重実行を防ぎ、完了時event.completedを呼ぶ設計 |
| XML変更案 | SharedRuntime要件、Workbook配下のRuntimes/Runtime lifetime=long、FunctionFileを共通ページへ。コマンドハンドラーを共通runtimeへ統合。公式指示に従いShared RuntimeではTaskpaneIdを付けない構成を検討 |
| 開閉状態 | Office.Addinには現在のvisibilityを直接取得するメソッドが列挙されていない。onVisibilityModeChangedとhide/showのPromise完了で状態追跡。初期状態不明/×で閉じる/別操作/イベント欠落を含む実機検証が必要。document.visibilityStateだけでペイン状態を決めない |
| 公開版の推奨 | 当面ShowTaskpane＋Excel標準の×を維持。トグルは別PoCで状態同期と配布条件を確認してから採用判断 |

SharedRuntime 1.1のサポートはWeb版Excel、Windows Excel Version 2002（Build 12527.20092）以降、Mac Excel 16.35（20030802）以降。これはSharedRuntimeだけの最低条件で、SheetDockのExcelApi 1.17等の条件は別途必要。iOSは非対応。[要件表](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/shared-runtime-requirement-sets)

ストア配布について公式資料に差がある。

- [Learnの表示/非表示ガイド](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in)は、2026-02-23以降Marketplace公開アドインのshowAsTaskpane呼び出しは無視され、サイドロード/集中展開のみ適用と記載。
- [Microsoft 365 Developer Blogの2026-02-24告知](https://devblogs.microsoft.com/microsoft365dev/increased-control-over-office-add-in-user-experiences/)は、3月2日からの変更について、明示的ユーザー操作（リボンクリック等）による呼び出しは例外と記載。

そのため「リボンクリックならAppSourceでも確実に開閉できる」とは断定しない。公式への確認と実際の配布経路での検証が必要。自動表示や無断再表示を回避し、ShowTaskpaneを残す方式を公開時の代替とする。Shared Runtime導入自体が公開不可という意味ではない。

公式参考：

- [Office.Addin API](https://learn.microsoft.com/en-us/javascript/api/office/office.addin?view=common-js-preview)：hide、showAsTaskpane、onVisibilityModeChangedとSharedRuntime 1.1。URLのpreviewビューは閲覧指定であり、ここで挙げたAPIは正式Requirement Set。
- [Shared Runtime設定](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)：XML/共通ページ/lifetime/一つのペインの制約。

## 今回変更したファイル（前回未コミット修正に追加）

- `src/state/useWorksheets.ts`：変更要求世代、監視、回復上限、古い結果の不採用。
- `src/sync/pendingWatch.ts`：実APIを取消さない長時間待機監視。
- `src/sync/coalescedRefresh.ts`：集約の診断記録。
- `src/excel/workbookState.ts`：取得中断理由の識別。
- `src/excel/worksheets.ts`：期限切れ/終了後に新しいアクティブ化を送らない。
- `src/taskpane/App.tsx`：応答待ちと操作ブロックの表示分離。
- `src/diagnostics.ts`：sessionとConsole出力制限。
- `tests/pendingWatch.test.ts`、`tests/useWorksheets.test.tsx`：長時間待機・最終状態追従・古い結果・再表示等の回帰テスト。
- 本書、README、前回調査記録、testing.md：追加実機報告と調査結果。

## 残課題

今回の修正で実Excelの無応答/RichApi.Errorが解消したとはまだ扱わない。APIが返ってこない根本原因、ホストのコンテキスト失効、長時間安定性、複数PoC、Mac/Windows、リボントグルは未確認。返答しないOffice呼出しを安全に強制取消しする機能は追加していない。

## 最終自動検証結果

- TypeScript型チェック成功。
- 18テストファイル、218件成功。前回209件を維持し9件追加。
- 名前変更の途中結果を不採用にして最終状態に追従、初回/手動無応答時のスピナー解除、回復取得上限、初回API失敗後復旧、破棄/再表示、ロックを期限切れで解放しないことを検証。
- Development/Productionビルド成功、エラー0・警告0。
- マニフェスト検証成功。
- Webpackモジュール一覧でProductionのPoCは0、Developmentは8モジュール。Productionの診断window APIも不在。
- manifest.xml、public/commands.html、package.json、package-lock.jsonはHEADと一致。
- コミットは作成せず、前回と今回の変更を未コミットで保持。HEADは46dee0351dabb0e91033b7f23679e5bae695d0d6のまま。

これは純粋処理/React/Office.jsモックによる自動検証で、実Excelの連続改名や無応答の解消を保証しない。今回の実機チェックはユーザーによる再検証待ち。

## ユーザー提供ログによる停止箇所の特定

診断session `mvs4yf3v7b0ls`、seq 1～152（欠番なし）を解析。コードの変更・追加の実機操作は行わず、以下を事実として追記する。

- 起動時登録は1回、正常に完了。最初の4回の通常取得は完了し、結果が適用されている。
- seq 91のonActivatedから5回目の取得を開始。metadataのsyncは73msで完了（seq 99）。
- seq 100の`api.sync.active.start`以後、対応するend/errorがない。該当コードは`getActiveWorksheet()`で得たシートの`id`をloadした後の`context.sync()`。同期呼び出しに入ったところまで確認できるが、ホスト内の待機理由は不明。
- この取得ではSHEETS総数評価へ進んでいない。したがって今回の未完了バッチをSHEETS処理の停止とは扱わない。
- 最初の名前変更イベント（seq 101）はactive sync開始の6,145ms後。名前変更イベントから始まった取得が停止したわけではなく、名前変更前から未完了の取得がある。ユーザーが名前編集UIを開始した時刻は記録されておらず、編集モードとの関係までは判断できない。
- seq 113で15秒監視が発動。その後もイベントと要求受付は継続するが、実取得は増えず集約待機となっている。
- 最終ログ時点でactive syncの未完了区間は少なくとも745,324ms（12分25.324秒）。ログ取得そのものの時刻は別なので、この値は下限。
- APIロックは6回開始・5回終了。要求8回のうち実行開始6回で、後続2件はロック待ち。イベントの再帰でAPIが無制限に並列実行された証拠はない。
- PoCの更新開始は停止後。PoCのAPI観測開始ログはなく、ロック待ちである。少なくとも今回停止したAPIはPoC観測ではない。
- error/runtime.unhandledrejection等の記録はない。今回のログから以前のRichApi.Errorと同一原因とは結び付けられない。
- pane.initは1回、pane.dispose/pagehideは記録されていない。このログ単体ではペイン再表示・ランタイム再作成による回復は検証できない。

特定できたのは「通常取得のactive sync未完了 → ロック保持 → 後続同期の待機」という停止経路。Office側がactive syncへ応答しない根本原因は未特定。名前変更イベントの重複登録や更新集約を原因と決めつけない。

次の切り分けは、専用ブックでページ全体を再読み込みして新しいsessionを作り、PoC観測OFFのまま、まず改名せずシート切替だけを試す。停止する場合は同じactive syncか比較する。切替だけなら正常な場合に、名前変更UIを開いて確定する操作を加えて比較する。APIロックを強制解放する修正は行わない。今回のログ受領時点では追加コード修正・コミットは実施していない。

## 比較実機結果：改名しないシート切替

ユーザー報告：Excel for the web、PoCイベント観測OFF、名前変更なしのシート切替では長時間停止を再現しなかった。体感の反映時間は約1秒またはそれ未満、最終アクティブ表示は正しく、操作上の大きな遅延は感じなかった。RichApi.Errorも発生しなかった。計測ログによる厳密なレイテンシ値ではない。

この結果は「通常の切替だけで常に停止する」という仮説を弱める。一方、再現しなかった1試行だけで切替処理やホストの問題を除外はしない。前回ログの未完了active syncは改名イベントの6.145秒前に開始しているため、名前編集UIに入ったタイミングと取得処理との重なりを次に比較する。onNameChangedの発火時刻は名前編集開始時刻ではない。

次の比較（専用ブック・PoC観測OFF。停止したらその時点のログを保存して以後は行わず、ページ再読み込み）：

1. すでにアクティブなシートで一覧の反映完了を待ち、名前編集を開始して約2秒後にEscで取消す。名前確定イベントなしでも止まるか確認。
2. 同じく反映完了を待ってから名前を変更し、Enterで確定。連打せず1回ずつ確認。止まらなければ5回ほど反復。
3. 1・2が正常なら、別シートへ切り替えてすぐ名前編集・確定を試す。取得中の編集開始との重なりを比較。

記録はケース番号、取消し/確定、最後の改名直後に止まったか、診断snapshot。編集開始時刻はログにないので、操作順序の報告も必要。無理に同じケースを連打して再現させない。

公式の[セル編集中の実行待機](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-delay-in-cell-edit)も確認したが、delayForCellEditはセル編集を対象とする。シート名編集に同じ挙動が適用される根拠ではなく、今回の対策として有効化していない。既存コードでもこのオプションは指定していない。

今回の比較結果受領ではコードの追加変更・自動テスト再実行・コミットはせず、既存未コミット修正を維持して記録だけ更新。通常の切替の約1秒以内という体感値を、今回の長時間停止と同じ不具合として扱わない。

## ケース①実機結果：名前編集を開始してEscで取消し

ユーザー報告：名前の確定はせず、編集開始の約2秒後にEsc。15秒応答待ち警告が表示された。ケース②③は未実施。警告スクリーンショットとsession `mvs677fn1vhmw`（seq 1～62）の診断ログを照合した。

### ログ上の時系列

最初の`event.workbookActivated`（1793195348053ms）を相対時刻0とする。これは編集開始を示すマーカーではない。

| 相対時刻 | 確認できた処理 |
|---|---|
| −17.995秒 | 初回一覧取得の適用完了 |
| 0秒 | Workbookアクティブ化イベント |
| ＋0.076秒 | Worksheetアクティブ化イベント |
| ＋0.152秒 | 2回目の更新開始、旧イベント解除開始 |
| ＋0.303秒 | 解除完了、新イベント登録開始 |
| ＋0.416秒 | 再登録完了、一覧取得開始 |
| ＋0.518秒 | metadata取得完了 |
| ＋0.519秒 | active IDをloadした後のcontext.sync開始 |
| ＋15.153秒 | 更新全体の15秒監視が発動 |
| ＋160.469秒 | 最終の更新要求。active syncのend/errorは依然なし |

active syncはログの最終行まで少なくとも159.950秒（約2分40秒）未完了。警告は更新全体の開始から15秒であり、active sync開始から15秒ではない。SHEETS評価へは進んでいない。

onNameChangedは0回。登録start/endは各2回、解除start/endは各1回で、2回目はWorkbookアクティブ化に伴う順序通りの再登録。これを重複登録の証拠とは扱わない。PoCは初期化のみで、観測API実行の記録はない。APIエラー/未処理例外の記録もない。

### 判断

- 名前の変更確定・onNameChanged通知がなくても、前回と同じactive sync未完了が発生した。したがって「確定後の名前変更イベント処理が停止の必要条件」という仮説は否定できる。
- 切替のみの比較では長時間停止なし、編集開始→取消しでは停止したというユーザー報告から、名前編集UIとの関連を強く疑う。ただし、ホスト内部の原因や、編集モードがAPIを直接止めたことまでは未確定。
- **編集開始とEscの正確な時刻は現在の診断に記録されていない。** Workbook/Worksheetアクティブ化時刻を編集開始時刻へ読み替えてはならない。active syncが編集開始の前・途中・終了後のどこに入ったかは、このログだけでは厳密に照合できない。
- ユーザーの「約2秒後Esc」と警告/ログを合わせると、編集を取消しても待機が速やかに解消しなかったケースである。通常の短時間の編集待ちだけでは説明できない。
- 今回は再登録があったが、前回停止ログでは初回登録後に再登録なしで同じactive syncが停止した。再登録だけを共通原因とは扱わない。

根本原因が特定できていないため、delayForCellEditの推測適用、APIロックの強制解除、アクティブシートの未検証キャッシュで安全判定を代用する変更は行っていない。ケース②③を追加でユーザーへ要求せず、次の実装調査対象をactive ID取得経路と名前編集UIの重なりへ限定する。今回の対応はログ解析と文書追記のみ。既存未コミット修正を維持し、追加コード変更・コミットなし。


## アクティブ取得経路の改善

[バッチ化の構造調査と比較](active-read-batching.md)を追加。通常取得をmetadata＋activeの同一syncと総数syncへ変更（3→2回）。エラー時の一覧保持は新Contextで1回だけ退避する。未応答解消は未検証。PoC観測は従来構造を保持。

## 独立したAPI比較実験の追加

[読み取り比較A/B/C](active-read-comparison.md)を追加。通常の同期・再登録処理は変更せず、開発専用の別起動モードで一覧のみ・既知ID・アクティブIDを比較する。比較結果は実機未検証、コミットなし。

## 通常表示と変更前検証の分離

通常UIログ`mvusvtuijciw1`でmetadataActive完了後のtotal同期が254.542秒以上未完了だったことを受け、[通常表示の取得見直し](normal-display-read.md)を実施。表示用の取得からSHEETSを分離し、移動前後の厳密な確認には残す。実機での停止解消は未確認、コミットなし。
