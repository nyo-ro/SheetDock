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
src/domain/      データ型・検索・位置順ソート
public/         コマンドページ・開発用ヘルプ
assets/         公式テンプレート由来の仮アイコン
manifest.xml    Excel専用の開発用XMLマニフェスト
tests/         TypeScript単体テスト
```

通常UIでの追加・削除・複製・名前変更、表示状態変更、タブ色、複数選択/複数移動、一括操作、設定保存、自動起動は未実装です。検索中も、表示シートの境界をブック全体の位置へ変換して単一ドラッグ移動できます。グラフシート等の混在、または総シート数を確認できない場合も移動できません。複数ウィンドウの実機検証は未完了です。

ユーザーによるWeb版Excelの確認で、通常UIの自動同期・単一ドラッグ・キャンセル・旧版の検索中ドラッグ無効化は成功しています。グラフシートが一覧に表示されないことも確認済みで、別課題です。今回追加した検索中移動・複数移動PoC・混在ガードの詳細・構成保護・PoC結果表示は実機未検証です。以前の一覧/切替、名前検索、Hidden/VeryHiddenの表示状態、PoC単一移動の成功記録は維持します。Windows Microsoft 365・Macは未検証です。

履歴資料の`docs/sheetdock-phase1.md`と`docs/sheetdock-phase2a.md`は、採用を終了した旧VSTO構成の記録です。現在の開発手順には使用しません。

## 開発専用の技術検証UI（Phase 1-B②）

`npm run dev`で開発配信した場合のみ、基本UIの下に「Office.js技術検証」が表示されます。PoCの変更操作は初期状態では無効です。通常UIの自動同期は自動で開始され、PoCのイベント観測スイッチとは独立しています。

[検証結果・Web版Excelのチェックリスト](docs/sheetdock-web-phase1b.md)を読んで、使い捨ての検証ブックで操作してください。混在グラフシートのブックは観測専用。VeryHidden削除の復元は保証できず、本番採用不可としています。

`npm run build`のproduction成果物には`src/poc`のコードを含めません。通常UIの自動同期と単一ドラッグはproductionにも含まれます。
