# GitHub Pages 配信（開発版）

この構成は、Office Web Add-inの**静的なタスクペインをホスティング**するものです。Excelへの自動インストールやAppSource配布ではありません。GitHub Pagesの配信先は `https://nyo-ro.github.io/SheetDock/` です。

## 構成

- ローカル開発用の `manifest.xml` は引き続き `https://localhost:3000` を参照する。
- `manifest.pages.xml` はGitHub Pages用で、アイコン・ヘルプ・タスクペイン・コマンドページのURLが `https://nyo-ro.github.io/SheetDock/` を参照する。
- `.github/workflows/pages.yml` は `main` 更新時または手動実行で、Node.js 24による `npm ci` → `npm run build` → Pagesへのデプロイを行う。
- 公開用成果物の `dist/manifest.xml` を `manifest.pages.xml` で上書きする。開発用のルート `manifest.xml` には影響しない。
- 配信するコードはproductionビルドで、開発専用PoC・独立比較モードは含まれない。

## GitHub管理画面での設定

GitHubリポジトリの **Settings → Pages → Build and deployment → Source** を **GitHub Actions** に設定する。GitHub連携からこの設定を直接変更できない場合は、リポジトリ管理者が操作する必要がある。

**Actions** の `Deploy SheetDock to GitHub Pages` ワークフローでビルド・デプロイが正常終了したことを確認する。初回のデプロイでは、Pagesの有効化や組織・アカウントの権限設定が必要な場合がある。

確認先：

- 配信ページ：`https://nyo-ro.github.io/SheetDock/`
- タスクペイン：`https://nyo-ro.github.io/SheetDock/taskpane.html`
- サイドロード用マニフェスト：`https://nyo-ro.github.io/SheetDock/manifest.xml`

## Excel for the webで試す

1. 公開URLの `taskpane.html` が開けること、およびマニフェストURLがXMLを返すことを確認する。
2. 公開マニフェストをファイルとしてダウンロードする。GitHubのソース `manifest.pages.xml` も内容は同じ。
3. Excel for the webで使い捨ての検証ブックを開き、アドイン管理画面から公開マニフェストをアップロードする。
4. `SheetDockを開く` からタスクペインを表示し、一覧・シート切替・検索・自動同期を確認する。シート並び替えは重要でないテストブックで確認する。

**注意：** 同じOfficeアドインIDを使っているため、以前のlocalhost開発版をサイドロード済みの場合は競合を避けるため開発版を解除してから試す。ブラウザーやExcelのマニフェストキャッシュが残ることがある。組織のポリシーでサイドロードできない場合は管理者設定が必要。ExcelApi 1.17を要求し、Windows/Macは未検証。

## 既知の制約

- この仕組みは静的ファイルの配信であり、Excelへの登録・配布を自動化しない。
- GitHub Pages上に公開するものは誰でも閲覧可能。Excelブックの内容はPagesのサーバーへ保存する設計ではないが、Excelアクセス権限を持つアドインなので重要なブックでは実機確認が完了するまで使用しない。
- 動作確認はExcel for the webのローカル配信で行ったもので、Pages配信・長時間使用・Windows/Macでの安定性は別途確認が必要。
- ローカルの証明書、`npm run dev`、開発用マニフェストは従来どおり利用できる。
