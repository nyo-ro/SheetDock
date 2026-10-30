# SheetDock

TypeScript・React・Office.jsで新規開発する、Excel向けOffice Web Add-inです。日本語のタスクペインからワークシートを確認・切り替えできます。自動同期と単一シートのドラッグ移動を通常UIへ統合しています。AppSource公開用の構成は未整備です。

## 実装済み

- 通常ワークシートのID・名前・位置・表示状態・アクティブ状態を取得し、位置順に表示
- 表示シートへの切り替え（直前に対象の存在・表示状態を再確認し、切り替え結果も確認）
- 名前の部分一致検索（大文字小文字を区別しない）とクリア
- シートのアクティブ化・追加・削除・名前・位置・表示状態・シート保護変更をイベントで自動反映
- 単一シートのドラッグ移動（挿入線、先頭/途中/末尾、同位置/キャンセル対応）
- 手動更新、読み込み中表示、API失敗・未対応環境のエラー表示
- Hidden／VeryHiddenは状態のみ表示し、クリックによる切り替え・再表示は行わない

## 開発

Node.js 24 LTS、npm、VS Code、ExcelApi 1.17に対応するExcelが必要です。

```sh
npm ci
npm run typecheck
npm test
npm run build
npm run validate
```

HTTPS証明書の準備後、`npm run dev`で`https://localhost:3000`を起動します。環境を変更する証明書の信頼登録・サイドロードは明示的に実行する手順に分けています。

- [安定化作業の到達点・実機結果・次の作業](docs/stabilization-status.md)
- [アクティブ取得停止の調査・バッチ化（履歴）](docs/active-read-batching.md)
- [自動同期の継続安定化・リボン開閉調査（履歴）](docs/sync-stability-ribbon.md)
- [名前変更後の遅延・RichApi.Error調査（履歴・根本原因未特定）](docs/rename-sync-investigation.md)
- [検索中移動・複数移動PoCの最新手順](docs/filtered-reorder-multi-poc.md)
- [自動同期・単一ドラッグの実装とWeb実機チェック](docs/worksheet-sync-reorder.md)
- [開発環境・HTTPS・サイドロード](docs/development.md)
- [テスト手順・既知の制約](docs/testing.md)
- [移行記録・撤去一覧](docs/sheetdock-web-phase1a.md)
- [Web版の調査要点と引き継ぎ](docs/web-migration-notes.md)

## 構成

```text
src/taskpane/    React UIとエントリーポイント
src/excel/       Office初期化・対応API判定・Excel操作
src/state/       一覧・検索・読み込み・エラーのReact状態
src/domain/      データ型・検索・移動計画・安全判定
src/sync/        更新集約・イベント共有・未応答監視
src/dev/         開発専用の独立読み取り比較
src/poc/         開発専用の高度な操作検証
public/         コマンドページ・開発用ヘルプ
assets/         公式テンプレート由来の仮アイコン
manifest.xml    Excel専用の開発用XMLマニフェスト
tests/         TypeScript単体テスト
```

通常UIでの追加・削除・複製・名前変更、表示状態変更、タブ色、複数選択/複数移動、一括操作、設定保存、自動起動は未実装です。検索中も、表示シートの境界をブック全体の位置へ変換して単一ドラッグ移動できます。グラフシート等の混在、または総シート数を確認できない場合も移動できません。複数ウィンドウの実機検証は未完了です。

ユーザーの最新のExcel for the web実機検証では、名前編集開始から約2秒後のEsc、名前変更5〜10回、通常・検索中の単一ドラッグ移動がすべて成功し、今回の再現手順ではAPI応答停止が発生しませんでした。根本原因の完全特定や全条件での解消は意味しません。以前の一覧・切替・検索・自動同期・ドラッグキャンセル・Hidden/VeryHidden表示・PoC単一移動の成功記録も維持します。

複数移動PoC、構成保護・読み取り専用・グラフシート混在時の拒否、途中失敗、PoCキャンセル・結果表示、独立A/B/C比較、長時間使用、Windows Microsoft 365・Macは未検証です。グラフシートが一覧に表示されないことは確認済みです。詳しい確認範囲は[現在の状況](docs/stabilization-status.md)を参照してください。

履歴資料の`docs/sheetdock-phase1.md`と`docs/sheetdock-phase2a.md`は、採用を終了した旧VSTO構成の記録です。現在の開発手順には使用しません。

## 開発専用の技術検証UI（Phase 1-B②）

`npm run dev`で開発配信した場合のみ、基本UIの下に「Office.js技術検証」が表示されます。PoCの変更操作は初期状態では無効です。通常UIの自動同期は自動で開始され、PoCのイベント観測スイッチとは独立しています。

[検証結果・Web版Excelのチェックリスト](docs/sheetdock-web-phase1b.md)を読んで、使い捨ての検証ブックで操作してください。混在グラフシートのブックは観測専用。VeryHidden削除の復元は保証できず、本番採用不可としています。

`npm run build`のproduction成果物には`src/poc`・`src/dev`のコードを含めません。通常UIの自動同期と単一ドラッグはproductionにも含まれます。

## 独立した読み取り比較（開発専用）

通常画面の「読み取り比較モードを開く（再読み込み）」から、通常UI・PoCを起動しないA/B/C比較へ移れます。1実行環境につき1試験、診断JSONは比較画面内でコピーできます。[手順・制約・結果の読み方](docs/active-read-comparison.md)を参照してください。実機結果は未確認です。

通常表示の取得と移動前の安全確認の分離については[最新の変更・実機手順](docs/normal-display-read.md)を参照してください。表示更新では総数を取得せず、ドロップ後の厳密な確認に成功した場合だけ移動します。
