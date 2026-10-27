# SheetDock Phase 2-A 作業記録

実施日: 2026-10-27（日本時間）

## 作業前の確認

- Phase 1記録を確認。初回コミット前のmain、remoteなし、全ソース未追跡。
- コピー79ファイルが元SheasideとSHA-256一致。記録を含む80ファイルの一覧が一致し、想定外の追加・編集なし。
- 変更前ハッシュ・パス対応表・検証ログはartifacts/phase2a/に保存（ローカル専用、Git無視対象）。

## 本体の構成と変更

- SheetDock.slnx → SheetDock/SheetDock.csproj。
- AssemblyName / RootNamespace / 名前空間: SheetDock。
- Title / Product: SheetDock。Description: SheetDock - Excelシート管理アドイン。
- AssemblyVersion / AssemblyFileVersion: 1.0.0.0。InformationalVersion: 1.0.0。
- リボン: SheetDockRibbon.cs / SheetDockRibbon.xml、埋め込み名 SheetDock.SheetDockRibbon.xml。
- Resources.Designer.cs: SheetDock.Properties.Resources。VSTO生成コード・XMLも名前空間を統一。
- リボンID、UIタイトル・メッセージ内製品名、Trace、ドラッグ形式をSheetDockへ変更。
- ドラッグ形式: SheetDock.LocalWorksheetMove。
- 設定保存先: %LOCALAPPDATA%\SheetDock\settings.xml。XMLルート: SheetDock。
- ShowHidden / ShowVeryHiddenの初期値はともにfalse。旧設定インポート処理は追加していない。
- SheetPaneManagerの単一UserPreferencesとChanged通知を共有する仕組みを維持。
- PNGはsheetdock-16.png / sheetdock-32.png / sheetdock-256.pngへ改名。画像バイト列と描画ロジックは維持。
- tests内の参照・期待値とTest-SheetDock.ps1 / Verify-Phase.ps1のパスを追従。

## 変更ファイル一覧

フォルダー移動によるパス変更を含む。

| 変更前 | 変更後 | 内容 |
| --- | --- | --- |
| Sheaside.slnx | SheetDock.slnx | 更新 |
| Sheaside/Assets/README.md | SheetDock/Assets/README.md | 更新 |
| Sheaside/Assets/sheaside-16.png | SheetDock/Assets/sheetdock-16.png | 内容不変・移動のみ |
| Sheaside/Assets/sheaside-256.png | SheetDock/Assets/sheetdock-256.png | 内容不変・移動のみ |
| Sheaside/Assets/sheaside-32.png | SheetDock/Assets/sheetdock-32.png | 内容不変・移動のみ |
| Sheaside/Controls/BatchDetailsDialog.cs | SheetDock/Controls/BatchDetailsDialog.cs | 更新 |
| Sheaside/Controls/ColorPaletteDialog.cs | SheetDock/Controls/ColorPaletteDialog.cs | 更新 |
| Sheaside/Controls/ModernButton.cs | SheetDock/Controls/ModernButton.cs | 更新 |
| Sheaside/Controls/RenameSheetDialog.cs | SheetDock/Controls/RenameSheetDialog.cs | 更新 |
| Sheaside/Controls/ReorderableSheetListBox.cs | SheetDock/Controls/ReorderableSheetListBox.cs | 更新 |
| Sheaside/Controls/RibbonArtwork.cs | SheetDock/Controls/RibbonArtwork.cs | 更新 |
| Sheaside/Controls/SheetListControl.cs | SheetDock/Controls/SheetListControl.cs | 更新 |
| Sheaside/Models/SheetBatch.cs | SheetDock/Models/SheetBatch.cs | 更新 |
| Sheaside/Models/SheetSelectionState.cs | SheetDock/Models/SheetSelectionState.cs | 更新 |
| Sheaside/Models/TabColorChoice.cs | SheetDock/Models/TabColorChoice.cs | 更新 |
| Sheaside/Models/WorksheetDragEventArgs.cs | SheetDock/Models/WorksheetDragEventArgs.cs | 更新 |
| Sheaside/Models/WorksheetItem.cs | SheetDock/Models/WorksheetItem.cs | 更新 |
| Sheaside/Properties/AssemblyInfo.cs | SheetDock/Properties/AssemblyInfo.cs | 更新 |
| Sheaside/Properties/Resources.Designer.cs | SheetDock/Properties/Resources.Designer.cs | 更新 |
| Sheaside/Properties/Resources.resx | SheetDock/Properties/Resources.resx | 内容不変・移動のみ |
| Sheaside/Properties/Settings.Designer.cs | SheetDock/Properties/Settings.Designer.cs | 更新 |
| Sheaside/Properties/Settings.settings | SheetDock/Properties/Settings.settings | 内容不変・移動のみ |
| Sheaside/Services/ExcelSheet.cs | SheetDock/Services/ExcelSheet.cs | 更新 |
| Sheaside/Services/ExcelSheetBatchService.cs | SheetDock/Services/ExcelSheetBatchService.cs | 更新 |
| Sheaside/Services/ExcelWorksheetService.cs | SheetDock/Services/ExcelWorksheetService.cs | 更新 |
| Sheaside/Services/OfficeThemeReader.cs | SheetDock/Services/OfficeThemeReader.cs | 更新 |
| Sheaside/Services/SheetBatchRunner.cs | SheetDock/Services/SheetBatchRunner.cs | 更新 |
| Sheaside/Services/SheetPaneManager.cs | SheetDock/Services/SheetPaneManager.cs | 更新 |
| Sheaside/Services/UserPreferences.cs | SheetDock/Services/UserPreferences.cs | 更新 |
| Sheaside/Services/WorksheetMoveRules.cs | SheetDock/Services/WorksheetMoveRules.cs | 更新 |
| Sheaside/Services/WorksheetNameRules.cs | SheetDock/Services/WorksheetNameRules.cs | 更新 |
| Sheaside/Sheaside.csproj | SheetDock/SheetDock.csproj | 更新 |
| Sheaside/SheasideRibbon.cs | SheetDock/SheetDockRibbon.cs | 更新 |
| Sheaside/SheasideRibbon.xml | SheetDock/SheetDockRibbon.xml | 更新 |
| Sheaside/ThisAddIn.cs | SheetDock/ThisAddIn.cs | 更新 |
| Sheaside/ThisAddIn.Designer.cs | SheetDock/ThisAddIn.Designer.cs | 更新 |
| Sheaside/ThisAddIn.Designer.xml | SheetDock/ThisAddIn.Designer.xml | 更新 |
| tests/AppearanceTests.cs | tests/AppearanceTests.cs | 更新 |
| tests/BatchTests.cs | tests/BatchTests.cs | 更新 |
| tests/ChartTests.cs | tests/ChartTests.cs | 更新 |
| tests/ReleaseTests.cs | tests/ReleaseTests.cs | 更新 |
| tests/ReorderTests.cs | tests/ReorderTests.cs | 更新 |
| tests/SelectionTests.cs | tests/SelectionTests.cs | 更新 |
| tests/SmokeTests.cs | tests/SmokeTests.cs | 更新 |
| tools/Test-Sheaside.ps1 | tools/Test-SheetDock.ps1 | 更新 |
| tools/Verify-Phase.ps1 | tools/Verify-Phase.ps1 | 更新 |
| （新規） | docs/sheetdock-phase2a.md | 本記録 |

## 検証結果

- ソリューション参照、csprojのCompile / EmbeddedResource / None参照は存在確認済み（保留中の証明書参照を除く）。
- 元ファイルとの比較で、対象内の差分は製品名置換と指定アセンブリメタデータのみ。製品名以外の日本語メッセージや機能ロジックに差分なし。
- Office COM GUID、VSTO種別GUIDを含む既存GUIDは全ファイルで維持。ProjectGuidとアセンブリGUIDも今回は変更なし。
- Debug / ReleaseのソリューションCompile成功、警告0・エラー0。
- Compile単独ではリソース準備が省略されたため、プロジェクトをResolveReferences;PrepareResources;Compileで再検証。両構成とも成功、警告0・エラー0。検証ターゲット選択の問題で、本体コードの不具合ではない。
- Release DLLの実リソース名SheetDock.SheetDockRibbon.xml / SheetDock.Properties.Resources.resources、アセンブリ名、3種類のバージョンを確認。
- リボンXML内のコールバックがクラスに存在することを確認。
- tools/Test-SheetDock.ps1: PASS 430 assertions (managed only)。設定初期値・永続化・複数ビュー同期・検索・選択・ドラッグ・一括操作・テーマなど既存テストを実行。
- Excel統合・実機リボン操作・VSTO読み込み・MSI/ZIPテストは未実施。
- 署名・配布マニフェスト生成をコマンド引数で無効化し、登録ターゲットを含まないCompile系列のみ実行。プロジェクトの配布設定は変更していない。
- 元Sheasideは.git・無視対象を含む全1211ファイルの一覧・SHA-256が作業前と一致。元リポジトリへの書き込み操作なし。
- SheetDockにcommit / stage / remote / tag / push / release操作なし。新機能・証明書・秘密情報の追加なし。

検証コマンド（MSBuild.exeはローカルVisual Studioのもの。両コマンドをConfiguration=Releaseでも実行）:

    MSBuild.exe SheetDock.slnx /t:Compile /p:Configuration=Debug "/p:Platform=Any CPU" /p:VisualStudioVersion=18.0 /p:SignManifests=false /p:GenerateManifests=false /p:RegisterForComInterop=false
    MSBuild.exe SheetDock/SheetDock.csproj "/t:ResolveReferences;PrepareResources;Compile" /p:Configuration=Debug /p:Platform=AnyCPU /p:VisualStudioVersion=18.0 /p:SignManifests=false /p:GenerateManifests=false /p:RegisterForComInterop=false
    ./tools/Test-SheetDock.ps1

## 残存するSheaside表記とPhase 2-B対象

| ファイル | 保留理由・必要な対応 |
| --- | --- |
| SheetDock/SheetDock.csproj | ManifestKeyFileとNoneのSheaside_TemporaryKey.pfx参照2件、既存thumbprintを維持。新署名・マニフェスト配布設定を準備。証明書実体はなし |
| installer/Sheaside.wxs | Excelアドイン登録キー、製品名、配置先、MSI ProductCode生成方針・UpgradeCode・Component GUID、DLL/VSTO名、日本語化 |
| tools/New-DevelopmentCertificate.ps1 | 旧プロジェクトパス、証明書名、subject、プロジェクト書き換え先を更新し新規証明書を管理 |
| tools/Build-Msi.ps1 | ソリューション、payload、DLL、manifest、VSTO、MSI/ZIP名、配布文書と公開URLへの依存を更新 |
| tools/Publish-ReleaseCandidate.ps1 | プロジェクト・出力先・配布アセンブリ名とVSTOマニフェストを更新 |
| tools/Test-Distribution.ps1 | 新payload、埋め込みリボン名、バージョン、ZIP/MSI名、旧配布物の固定ハッシュ前提を見直す |
| tools/Test-Msi.ps1 | 新DLL・マニフェスト・登録情報に期待値を合わせる |
| tools/Test-MsiPrerequisites.ps1 | 今回内容変更なし。Phase 2-BのMSI・前提条件・日本語化の検証に使用する |
| tools/Test-ExcelUi.ps1 | 旧登録キー・旧設定パス・旧UI名に依存。Phase 2-Bで一緒に更新してから実行する |
| tools/Verify-Phase.ps1 | ソリューションとテスト起動名のみ更新済み。通常Buildは署名・登録を伴うため、Phase 2-B完了後に総合実行する |
| README.md、docs/deployment.md、docs/development.md、docs/publication.md、docs/testing.md、docs/usage.md | 配布・導入・開発手順・公開URLは旧製品のまま。新製品の公開設定確定後に更新 |
| docs/archive/*.md、docs/sheetdock-phase1.md | 過去記録と分離時点の事実として保持。機械的に置換しない |
| docs/sheetdock-phase2a.md | 変更前との比較・保留事項の説明に旧名を使用 |

## 未確認事項・制約

- 配布可能な正式ビルドではなく本体とリソースのコンパイル確認。旧証明書参照と配布・登録設定の整合はPhase 2-B待ち。
- 実際のExcel上でのリボン・複数ウィンドウ・起動時の設定分離は未確認。設定同期は既存管理コードテストとロジック同一性で確認。
- Sheasideとの同時インストールは保証しない。旧登録や旧ユーザー設定は変更・削除していない。
- 初回コミット前のため通常のgit diffには変更一覧が出ない。本記録の一覧とartifacts/phase2a/内baseline.json / file-map.json / changed-files.jsonで比較可能。
