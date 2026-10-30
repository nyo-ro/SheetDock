# テストと制約

現在の確認済み範囲・未確認範囲は[安定化作業の到達点](stabilization-status.md)を参照する。過去の19件、87件、197件等のテスト結果は各Phaseの履歴であり、現在の件数ではない。

## 自動検証

Node.js 24環境で以下を実行する。Development出力を別ディレクトリにし、Productionの検査対象を上書きしない。

```powershell
npm run typecheck
npm test
npm run build -- --json .local-tools/final-stabilization-production-stats.json
npx --no-install webpack --mode development --output-path .local-tools/final-stabilization-dev-build --json .local-tools/final-stabilization-development-stats.json
npm run validate
```

最終再検証の結果は[記録](stabilization-status.md#自動検証)を参照。テストは純粋関数、React UI、Office.jsモック、取得・変更・キューの結合を含む。Excelホストの動作やAPI応答時間を保証しない。

Productionのwebpack統計に`src/dev`・`src/poc`モジュールがなく、生成JavaScriptに比較モードや診断公開APIの識別文字列がないことを確認する。Developmentでは該当モジュールが含まれることも確認する。

## Excel for the webでの再確認

使い捨ての検証用ブックを使う。既存の証明書・サイドロード設定で`npm run dev`を起動し、Excelのブラウザータブ全体を再読み込みする。通常画面を使用し、PoCのイベント観測はOFFにする。

1. 一覧とアクティブ表示を確認し、Excel側でシートを切り替える。
2. 名前編集を開始し約2秒後にEsc。続いて名前変更を5〜10回確定し、最終名に追いつくことを確認する。
3. 通常・検索中の単一シート移動を行い、Excelの順序と照合する。検索条件が維持されることも確認する。
4. ドラッグのキャンセル、ペインを閉じて再表示、シート追加・削除を確認する。今回の最新報告には再表示・追加・削除の個別再確認は含まれていない。
5. 後続の安全性検証では構成保護・読み取り専用・グラフシート混在の専用ブックで、変更が拒否されることを確認する。

未応答になったら追加の変更を止め、DevToolsで`localhost:3000`のSheetDockフレームを選び、次の読み取りコマンドでログを保存する。PowerShellやExcel本体のtopフレームでは実行しない。

```javascript
copy(JSON.stringify(window.sheetDockDiagnostics.snapshot(), null, 2))
```

発生手順・操作時刻・PoC観測ON/OFFを添える。未応答APIのロックは強制解除しない。保存後にExcelのブラウザータブ全体を再読み込みする。

## 高度な操作と調査用モード

- [通常表示と安全確認の分離](normal-display-read.md)：総数未確認の表示から書き込みを許可しない。
- [複数移動PoCの具体例](filtered-reorder-multi-poc.md)：通常UIの複数選択は未実装。
- [PoCチェックリスト](sheetdock-web-phase1b.md)：明示的な有効化・確認を経て専用ブックだけで使用する。
- [A/B/C独立比較](active-read-comparison.md)：開発専用、1実行環境1試験。実機比較は未実施。
- [初期環境・監査の履歴](sheetdock-web-phase1a.md)：依存の監査結果は当時の値。今回npm監査・依存更新は行っていない。

Windows Microsoft 365・Mac、長時間使用、共同編集、複数ウィンドウ、アクセシビリティ・高DPI等の品質確認は残る。公開構成とAppSource準備も未完了。
