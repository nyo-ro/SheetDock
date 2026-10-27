# Web版への引き継ぎ

## 製品方針

SheetDockはTypeScript・Office.js・Web UIで新規開発する。Sheaside v1.0.0（c34996234f413f64847f3ddaedcc47851331eec4）の機能・UX・安全性を参考にする。VSTOコード、登録情報、ユーザー設定は移植しない。最終的にWindows・Mac・WebとAppSource公開を目指す。

## 残した旧作業記録

- [sheetdock-phase1.md](sheetdock-phase1.md)：元ソースの版、初期分離時のコピー/除外方針、元リポジトリ保護の記録。
- [sheetdock-phase2a.md](sheetdock-phase2a.md)：旧名称変更で確認したリボン、設定保存、リソース、保護すべきOffice識別子の記録。

両資料は旧VSTO案の履歴として内容をそのまま保存している。記載されたファイル名やコマンドの多くは今回撤去済みで、現在のWeb構成への指示ではない。旧Phase 2-Bを進める予定もない。

## Phase 0調査のうち今回引き継ぐ判断

- 通常ワークシートの一覧、ID、名前、位置、アクティブ取得・切替はOffice.jsで扱える。
- 独立したグラフシートとワークシート上のグラフは別物。今回の一覧は通常ワークシートのみ。
- UI内の選択とExcelのアクティブ状態は将来も区別する。今回は複数選択状態を導入しない。
- 複数操作は原子的なトランザクションを前提にしない。操作直前にID・状態を再検証し、失敗を成功に見せない。
- 名前変更・移動等のイベントを今後検証する前提でExcelApi 1.17を製品の最低要件に置く。今回の基本APIだけならより低い要件も可能だが、対象を広げる判断は別途行う。
- 並び替え、VeryHidden操作、混在グラフシート、複数ウィンドウ、イベント漏れ、共同編集、AppSource配布は実機検証を経て仕様化する。
- XMLマニフェストを採用。今回使用するAPIは正式公開APIで、Preview APIを使わない。

これはPhase 0の会話内調査から、現在の最小実装に関係する論点を整理したもの。全機能の移植可否を再確認・確定した文書ではない。

## 公式参照

- [Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview)：id/name/position/visibility/activate。ページのPreview表示と個々のAPIの正式Requirement Setは区別する。
- [WorksheetCollection](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetcollection?view=excel-js-preview)：getActiveWorksheet/getItemOrNullObject。
- [Requirement Sets](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)：ExcelApi 1.17のホスト対応を確認する。

記録日：2026-10-27。実機での互換性は[テスト表](testing.md)に記録する。
