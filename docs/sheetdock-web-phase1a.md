# SheetDock Web Phase 1-A 作業記録

作業日：2026-10-27

既存のSheetDockディレクトリと.gitを継続使用した。開始時はmainの初回コミット前でremoteなし。全ファイル未追跡という状態はPhase 1/2-Aの作業記録と一致した。旧ソースはSheasideの対応ファイルに対する名称・バージョン変更と比較し、想定外の変更やWeb版の作業ファイルがないことを確認して撤去した。

## 撤去と保持

撤去した旧ファイルは106件。C#・WinForms・VSTOプロジェクト、ソリューション、WiX、MSI/VSTOビルド/署名/検証スクリプト、旧テスト、生成物、旧配布/利用文書を撤去。READMEと.gitignoreはWeb用へ置き換えた。

旧資料docs/sheetdock-phase1.md、docs/sheetdock-phase2a.mdは内容を保持した。用途と今後の調査要点はweb-migration-notes.mdへ整理した。Sheasideへ書き込み操作は行っていない。旧.gitを削除・再作成していない。

## 新構成と判断

UI、Excel API、React状態、検索/ソートを分離。ExcelApi 1.17、XMLマニフェスト、日本語、localhost HTTPS。新規ID：4e8952c2-160e-41ac-91a7-0469ab02c117。Office.jsは公式CDN。通常ワークシートのみを対象とし、非表示シートはAPI層でも拒否する。

Microsoftの公式React Task Paneテンプレートを参考に必要な構成を直接作成した。webpack、ts-loader、HTML/静的ファイルコピーを使用。テンプレートの不要なサンプル処理・ポリフィル等は導入していない。仮アイコンと構造の出典はTHIRD_PARTY_NOTICES.mdに記載した。

Node.jsは公式配布のポータブル24.22.1を.local-toolsへ配置し、公式SHA256と照合した。システムへのインストール・PATHの永続変更はなし。証明書信頼登録・Officeサイドロード・AppSource申請は行っていない。

## 撤去ファイルの全一覧

- `SheetDock.slnx`
- `SheetDock/Assets/README.md`
- `SheetDock/Assets/sheetdock-16.png`
- `SheetDock/Assets/sheetdock-256.png`
- `SheetDock/Assets/sheetdock-32.png`
- `SheetDock/Controls/BatchDetailsDialog.cs`
- `SheetDock/Controls/ColorPaletteDialog.cs`
- `SheetDock/Controls/ModernButton.cs`
- `SheetDock/Controls/RenameSheetDialog.cs`
- `SheetDock/Controls/ReorderableSheetListBox.cs`
- `SheetDock/Controls/RibbonArtwork.cs`
- `SheetDock/Controls/SheetListControl.cs`
- `SheetDock/Models/SheetBatch.cs`
- `SheetDock/Models/SheetSelectionState.cs`
- `SheetDock/Models/TabColorChoice.cs`
- `SheetDock/Models/WorksheetDragEventArgs.cs`
- `SheetDock/Models/WorksheetItem.cs`
- `SheetDock/Properties/AssemblyInfo.cs`
- `SheetDock/Properties/Resources.Designer.cs`
- `SheetDock/Properties/Resources.resx`
- `SheetDock/Properties/Settings.Designer.cs`
- `SheetDock/Properties/Settings.settings`
- `SheetDock/Services/ExcelSheet.cs`
- `SheetDock/Services/ExcelSheetBatchService.cs`
- `SheetDock/Services/ExcelWorksheetService.cs`
- `SheetDock/Services/OfficeThemeReader.cs`
- `SheetDock/Services/SheetBatchRunner.cs`
- `SheetDock/Services/SheetPaneManager.cs`
- `SheetDock/Services/UserPreferences.cs`
- `SheetDock/Services/WorksheetMoveRules.cs`
- `SheetDock/Services/WorksheetNameRules.cs`
- `SheetDock/SheetDock.csproj`
- `SheetDock/SheetDockRibbon.cs`
- `SheetDock/SheetDockRibbon.xml`
- `SheetDock/ThisAddIn.Designer.cs`
- `SheetDock/ThisAddIn.Designer.xml`
- `SheetDock/ThisAddIn.cs`
- `SheetDock/obj/Debug/.NETFramework,Version=v4.8.AssemblyAttributes.cs`
- `SheetDock/obj/Debug/Interop.Microsoft.Office.Core.dll`
- `SheetDock/obj/Debug/Interop.Microsoft.Office.Interop.Excel.dll`
- `SheetDock/obj/Debug/SheetDock.Properties.Resources.resources`
- `SheetDock/obj/Debug/SheetDock.csproj.AssemblyReference.cache`
- `SheetDock/obj/Debug/SheetDock.csproj.GenerateResource.cache`
- `SheetDock/obj/Debug/SheetDock.csproj.ResolveComReference.cache`
- `SheetDock/obj/Debug/SheetDock.dll`
- `SheetDock/obj/Debug/SheetDock.pdb`
- `SheetDock/obj/Release/.NETFramework,Version=v4.8.AssemblyAttributes.cs`
- `SheetDock/obj/Release/Interop.Microsoft.Office.Core.dll`
- `SheetDock/obj/Release/Interop.Microsoft.Office.Interop.Excel.dll`
- `SheetDock/obj/Release/SheetDock.Properties.Resources.resources`
- `SheetDock/obj/Release/SheetDock.csproj.AssemblyReference.cache`
- `SheetDock/obj/Release/SheetDock.csproj.GenerateResource.cache`
- `SheetDock/obj/Release/SheetDock.csproj.ResolveComReference.cache`
- `SheetDock/obj/Release/SheetDock.dll`
- `SheetDock/obj/SmokeTests/SmokeTests.exe`
- `SheetDock/obj/SmokeTests/SmokeTests.pdb`
- `SheetDock/obj/SmokeTests/rc5-palette.png`
- `SheetDock/obj/SmokeTests/rc5-pane.png`
- `artifacts/phase2a/baseline.json`
- `artifacts/phase2a/changed-files.json`
- `artifacts/phase2a/compile-Debug.log`
- `artifacts/phase2a/compile-Release.log`
- `artifacts/phase2a/file-map.json`
- `artifacts/phase2a/original-snapshot.json`
- `artifacts/phase2a/resources-Debug.log`
- `artifacts/phase2a/resources-Release.log`
- `docs/archive/rc.2-checklist.md`
- `docs/archive/rc.2-progress.md`
- `docs/archive/rc.3-checklist.md`
- `docs/archive/rc.3-progress.md`
- `docs/archive/rc.4-checklist.md`
- `docs/archive/rc.4-progress.md`
- `docs/archive/rc.5-checklist.md`
- `docs/archive/rc.5-distribution.md`
- `docs/archive/rc.5-install-checklist.md`
- `docs/archive/rc.5-progress.md`
- `docs/archive/rc.5-release-notes.md`
- `docs/archive/v0.4.1.md`
- `docs/archive/v0.5.1.md`
- `docs/archive/v0.5.md`
- `docs/archive/v0.6.md`
- `docs/archive/v1.0-checklist.md`
- `docs/archive/v1.0-progress.md`
- `docs/archive/v1.0-rc.md`
- `docs/deployment.md`
- `docs/development.md`
- `docs/publication.md`
- `docs/testing.md`
- `docs/usage.md`
- `installer/Sheaside.wxs`
- `tests/AppearanceTests.cs`
- `tests/BatchTests.cs`
- `tests/ChartTests.cs`
- `tests/ReleaseTests.cs`
- `tests/ReorderTests.cs`
- `tests/SelectionTests.cs`
- `tests/SmokeTests.cs`
- `tools/Build-Msi.ps1`
- `tools/New-DevelopmentCertificate.ps1`
- `tools/Publish-ReleaseCandidate.ps1`
- `tools/Test-Distribution.ps1`
- `tools/Test-ExcelUi.ps1`
- `tools/Test-Msi.ps1`
- `tools/Test-MsiPrerequisites.ps1`
- `tools/Test-SheetDock.ps1`
- `tools/Verify-Phase.ps1`

## 最終検証結果

- TypeScript型チェック：成功。
- Vitest 5.0.5：3ファイル・19テスト成功。検索/位置順ソート、非表示/消失対象の拒否、API失敗伝播、切替後ID確認、Office初期化・対応要件・タイムアウトを確認。
- webpack productionビルド：成功。タスクペインJSは約239 KiB。
- Microsoft公式office-addin-manifest 3.0.1：XMLスキーマ・HTTPS・リソース等の検証成功。開発マニフェストとしてvalid。ストア審査や実機互換性を保証するものではない。
- マニフェストのresid・ローカルURL参照：全件解決。
- HTTPS開発サーバー：起動成功。taskpane HTML、生成JS、commands/help、manifest、4種類のアイコンがHTTP 200。一時自己署名証明書を.local-tools内に生成し、検証クライアントへ明示的にCAを渡した。OSの信頼ストアは変更せず、サーバーは検証後に停止。
- Git管理候補に旧VSTOソース/ビルド成果物・秘密鍵・既知トークン形式がないことを確認。.local-tools、node_modules、distは除外。
- Sheasideは.gitを含む1,211ファイルをSHA256比較し、パスと内容が作業前と完全一致。
- Excel実機：WindowsにあるExcelはOffice 2013（15.0.5603.1000）のためExcelApi 1.17非対応。Microsoft 365、Mac、Webは未検証。サイドロード・信頼登録・既存ブックの操作なし。

## 採用バージョン

| パッケージ | バージョン |
|---|---|
| Node.js | 24.22.1 |
| React / React DOM | 18.3.1 |
| Fluent UI React | 9.74.11 |
| TypeScript | 5.9.3 |
| webpack | 5.112.2 |
| webpack-dev-server | 6.1.0 |
| Vitest | 5.0.5 |
| @types/office-js | 1.0.614 |
| office-addin-debugging | 7.0.1 |
| office-addin-manifest | 3.0.1 |
| office-addin-dev-certs | 2.0.10 |

Office.js本体はMicrosoft公式CDNの正式公開チャネルを使用し、npmへ同梱していない。具体的な依存解決はpackage-lock.jsonに固定。

## 依存関係の監査と残課題

初回監査の13件から、VitestとMicrosoft開発ツールを修正版へ更新し再検証した。`npm audit --omit=dev`は0件。開発依存を含む監査はHigh 8件が残る。主因はサイドロードCLIの推移依存である`@azure/msal-node`、`adm-zip`とその親パッケージ。`npm audit fix`では解消しなかった。互換性未確認の強制overrideは行っていない。

これらはタスクペインのproductionバンドルには含まれない。今回、CLIによる認証・サイドロードや外部ZIPの展開は実行していない。Microsoftツールの依存更新を確認し、実機サイドロード前に監査を再実行する。CLIの`start --help`起動と引数だけは確認済み。

npm 11のインストールスクリプト保護により、一部の認証関連ネイティブ依存（keytar等）のスクリプトは未承認。ビルド・テスト・マニフェスト検証は動作したが、将来の認証/サイドロードで必要な場合は内容を確認してから許可する。全パッケージのスクリプトを一括承認しない。

## Git

既存の.gitを保持して、指定された初回コミットメッセージ`chore: initialize SheetDock Office Add-in`で記録する。remote追加・push・タグ・リリースは行わない。コミットハッシュは完了報告と`git log -1`で確認できる。
