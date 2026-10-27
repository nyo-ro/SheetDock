# SheetDock

TypeScript・React・Office.jsで新規開発する、Excel向けOffice Web Add-inです。日本語のタスクペインからワークシートを確認・切り替えできます。現在は開発基盤を整えるPhase 1-Aで、AppSource公開用ではありません。

## 実装済み

- 通常ワークシートのID・名前・位置・表示状態・アクティブ状態を取得し、位置順に表示
- 表示シートへの切り替え（直前に対象の存在・表示状態を再確認し、切り替え結果も確認）
- 名前の部分一致検索（大文字小文字を区別しない）とクリア
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

独立したグラフシート、追加・削除・複製・名前変更、表示状態変更、タブ色、複数選択、ドラッグ、一括操作、設定保存、複数ウィンドウ、自動更新・自動起動は未実装です。Excel側の変更後は手動更新してください。アクティブ表示は直近の取得時点の状態です。

Windows Microsoft 365・Mac・WebのExcel実機検証は未実施です。今回の環境にはOffice 2013があり、必要なAPI要件を満たしません。通常のブラウザーだけで開いてもExcelの代わりにはなりません。

履歴資料の`docs/sheetdock-phase1.md`と`docs/sheetdock-phase2a.md`は、採用を終了した旧VSTO構成の記録です。現在の開発手順には使用しません。
