# SheetDock Phase 1 作業記録

実施日: 2026-10-27（日本時間）

## 基準と結果

- 基準: Sheaside v1.0.0 / c34996234f413f64847f3ddaedcc47851331eec4。
- 作業前後でHEADとv1.0.0タグの一致、未コミット変更なしを確認。
- 管理対象80ファイルから79ファイルを内容変更なしでコピー。全79ファイルのSHA-256一致を確認。
- 除外した管理対象: .vscode/settings.json（エディターの選好設定）。
- 管理対象外の.git、.vs、bin、obj、artifacts、既存MSI/ZIP、開発証明書・秘密鍵、個人設定はコピーしない。
- .gitignoreは元の内容を維持。IDEキャッシュ、bin/obj、*.user、*.suo、*.pfx、artifactsを除外する。
- Gitはmainで新規初期化。元の履歴・タグ・remoteを持ち込まず、ステージング・初回コミットも未実施。
- 本記録を1ファイル追加。コピー79ファイルと合わせて80ファイル（.git内部を除く）。
- 名称・実装の変更、ビルド、テスト、インストール、コミット、プッシュ、GitHub操作は実施していない。

## 検証と引き継ぎ事項

- コピー直後および最終検証でコピー全79ファイルのSHA-256一致を確認。
- コピー先一覧は管理対象一覧から明示除外1件を引いた集合と一致（本記録のみ追加）。
- 元リポジトリのHEAD・タグ・作業ツリー・remoteを再確認。本記録作成前後には無視対象と.git内部を含む全ファイルのSHA-256と一覧も比較。
- 秘密鍵・認証トークン・個人メール・ユーザーパスの候補を検査。候補は伏字の説明、検査パターン、テスト用値で、実秘密情報は確認されなかった。検査は既知パターンと内容確認の範囲。
- プロジェクト内の証明書参照名・thumbprintは元のまま。証明書ファイル・秘密鍵は持ち込んでいない。新しい署名設定はPhase 2以降で準備する。
- ユーザー設定データはコピーしていない。設定保存先はまだSheasideのままであり、初期値から開始するための保存先分離はPhase 2で行う。

## Phase 2以降の変更対象一覧（未着手）

以下は現行ファイル名。ソリューション・プロジェクト・名前空間・アセンブリ・リボン・VSTO登録・MSI識別子・スクリプト・現行文書のSheaside参照を更新する対象。

重点対象:
- Sheaside/Services/SheetPaneManager.cs: LocalApplicationData配下の設定保存先を分離。既存設定は移行しない。
- Sheaside/Services/UserPreferences.cs: 設定XMLの製品名。
- Sheaside/Sheaside.csproj、tools/New-DevelopmentCertificate.ps1: 新しい署名設定。
- installer/Sheaside.wxs: 製品識別子、登録先、表示名、日本語化。共存は保証しない。

全対象:

- README.md
- Sheaside.slnx
- Sheaside/Assets/README.md
- Sheaside/Controls/BatchDetailsDialog.cs
- Sheaside/Controls/ColorPaletteDialog.cs
- Sheaside/Controls/ModernButton.cs
- Sheaside/Controls/RenameSheetDialog.cs
- Sheaside/Controls/ReorderableSheetListBox.cs
- Sheaside/Controls/RibbonArtwork.cs
- Sheaside/Controls/SheetListControl.cs
- Sheaside/Models/SheetBatch.cs
- Sheaside/Models/SheetSelectionState.cs
- Sheaside/Models/TabColorChoice.cs
- Sheaside/Models/WorksheetDragEventArgs.cs
- Sheaside/Models/WorksheetItem.cs
- Sheaside/Properties/AssemblyInfo.cs
- Sheaside/Properties/Resources.Designer.cs
- Sheaside/Properties/Settings.Designer.cs
- Sheaside/Services/ExcelSheet.cs
- Sheaside/Services/ExcelSheetBatchService.cs
- Sheaside/Services/ExcelWorksheetService.cs
- Sheaside/Services/OfficeThemeReader.cs
- Sheaside/Services/SheetBatchRunner.cs
- Sheaside/Services/SheetPaneManager.cs
- Sheaside/Services/UserPreferences.cs
- Sheaside/Services/WorksheetMoveRules.cs
- Sheaside/Services/WorksheetNameRules.cs
- Sheaside/Sheaside.csproj
- Sheaside/SheasideRibbon.cs
- Sheaside/SheasideRibbon.xml
- Sheaside/ThisAddIn.Designer.cs
- Sheaside/ThisAddIn.Designer.xml
- Sheaside/ThisAddIn.cs
- docs/deployment.md
- docs/development.md
- docs/publication.md
- docs/testing.md
- docs/usage.md
- installer/Sheaside.wxs
- tests/AppearanceTests.cs
- tests/BatchTests.cs
- tests/ChartTests.cs
- tests/ReleaseTests.cs
- tests/ReorderTests.cs
- tests/SelectionTests.cs
- tests/SmokeTests.cs
- tools/Build-Msi.ps1
- tools/New-DevelopmentCertificate.ps1
- tools/Publish-ReleaseCandidate.ps1
- tools/Test-Distribution.ps1
- tools/Test-ExcelUi.ps1
- tools/Test-Msi.ps1
- tools/Test-Sheaside.ps1
- tools/Verify-Phase.ps1

### 関連リソース・ファイル名の確認対象

- Sheaside/Assets/sheaside-16.png
- Sheaside/Assets/sheaside-32.png
- Sheaside/Assets/sheaside-256.png
- Sheaside/Properties/Resources.resx
- Sheaside/Properties/Settings.settings

### 過去文書の扱いを確認する対象

以下はSheasideの開発記録。SheetDockの実績として扱わず、由来の明示方法をPhase 2で決める。過去の事実を機械的に製品名置換しない。

- docs/archive/rc.2-checklist.md
- docs/archive/rc.2-progress.md
- docs/archive/rc.3-checklist.md
- docs/archive/rc.3-progress.md
- docs/archive/rc.4-checklist.md
- docs/archive/rc.4-progress.md
- docs/archive/rc.5-checklist.md
- docs/archive/rc.5-distribution.md
- docs/archive/rc.5-install-checklist.md
- docs/archive/rc.5-progress.md
- docs/archive/rc.5-release-notes.md
- docs/archive/v0.4.1.md
- docs/archive/v0.5.1.md
- docs/archive/v0.5.md
- docs/archive/v0.6.md
- docs/archive/v1.0-checklist.md
- docs/archive/v1.0-progress.md
- docs/archive/v1.0-rc.md
