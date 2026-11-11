# GitHub Pages配信（開発版）

- 案内ページ：[https://nyo-ro.github.io/SheetDock/](https://nyo-ro.github.io/SheetDock/)
- 登録用XML：[manifest.xml](https://nyo-ro.github.io/SheetDock/manifest.xml)
- タスクペイン：[taskpane.html](https://nyo-ro.github.io/SheetDock/taskpane.html)
- ビルドと配信結果：[Actions](https://github.com/nyo-ro/SheetDock/actions/workflows/pages.yml)

静的なタスクペインの開発版配信です。正式リリース、AppSource公開、Excelへの自動インストールではありません。

## ローカルサーバーなしで利用する

1. 案内ページからXMLをダウンロードする。Node.js、npm、localhost用証明書、`npm run dev`は不要。
2. Excel for the webで使い捨ての検証ブックを開く。
3. ホーム → アドイン → その他のアドインからOfficeアドインの管理画面へ進み、「自分のアドインのアップロード」で保存したXMLを選択する。画面名称・配置は環境で異なる。
4. ホームの「SheetDockを開く」から起動する。ブラウザーで案内ページやtaskpane.htmlを開くだけではExcelに登録されない。

localhost開発版とPages版は同じアドインIDを使用する。登録済みの開発版を外してからPages版をアップロードする。開発版を再び検証するときは、Pages版の登録を外してルートの`manifest.xml`を再登録する。複数登録の共存は想定しない。

マニフェストをブラウザーで開くとXMLが表示される場合は、案内ページのダウンロードボタンから保存する。ファイルがHTMLのエラーページになっていないか確認する。

組織のポリシーによってアップロードを許可されない場合は管理者へ確認する。ブラウザー変更・キャッシュ削除等で再登録が必要になることがある。[Microsoftの公式サイドロード手順](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)を参照。

## 実機確認済み／未検証

localhost配信を使うExcel for the webでは一覧・切替・検索・自動同期・通常と検索中の単一ドラッグ、名前編集のEsc取消、連続した名前変更を確認済み。

**Pages配信でのExcel内動作は未検証。** HTTP配信・HTML/JS/画像の取得が成功しても、Excel APIの成功とは扱わない。Windows Microsoft 365・Mac、長時間使用、共同編集・複数ウィンドウ、保護や混在時の異常系、複数移動PoCも実機未検証のまま。

ExcelApi 1.17が必要。独立グラフシートは一覧対象外。Hidden/VeryHiddenは表示のみ。通常UIの複数選択・複数移動・一括操作は未実装。構成保護・読み取り専用・構成不明などでは移動を拒否する。共同編集との完全な排他、原子的な変更は保証しない。

## Excel for the webの最小チェック

1. ローカルサーバーを停止した状態でPages用XMLを登録し、タスクペインを開く。
2. Sheet1 → Sheet2 → Sheet3の検証ブックで一覧とアクティブ表示を照合し、SheetDockから切り替える。
3. 検索・クリアを試す。Excel側でシート切替・追加・名前変更・並び替えを行い、自動同期を確認する。
4. 名前編集開始から約2秒後にEsc、続いて名前変更を5〜10回行い、最終名に追いつくことを確認する。
5. Sheet2を先頭へドラッグしてSheet2 → Sheet1 → Sheet3になることを確認。キャンセル時は順序が変わらないことを確認する。
6. 名前に共通部分があるシートを用意し、検索中の単一移動後も検索条件とExcelの実順序が整合することを確認する。
7. ペインを閉じて再表示し、エラーや取得中のまま停止する現象がないか確認する。

異常時は操作手順・時刻・ブラウザー・Excel環境を記録する。PagesはProductionなのでPoC・独立比較・開発診断APIは存在しない。診断ログが必要な場合だけlocalhostのDevelopment配信で再現を試す。ブック内容や個人情報を公開Issueへ貼らない。

## ビルドの分離と検査

| コマンド | 用途 |
|---|---|
| `npm run dev` | localhostのHTTPS開発配信。PoC・診断・独立比較を維持 |
| `npm run build` | localhost用マニフェストを含む通常Productionビルド |
| `npm run build:pages` | Pages用Productionビルドと成果物検査 |
| `npm run validate` | 開発用マニフェストの検証 |
| `npm run validate:pages` | Pages用マニフェストの検証 |

Pagesビルドは`manifest.pages.xml`を`dist/manifest.xml`へコピーし、案内ページをindex.htmlとhelp.htmlへ配置する。ルートの開発用manifest.xmlは変更しない。AppDomainはPagesのorigin、各ファイルのURLは`/SheetDock/`配下。webpackのpublicPathも`/SheetDock/`へ固定する。

HTML、XMLが参照するローカル成果物の存在を検査し、webpackモジュール一覧と生成JSからPoC・独立比較・診断公開APIの混入を検査する。CSSはstyle-loaderでJSに含まれ、ページ起動時に適用される。案内ページのCSSはHTML内に含む。独立CSSファイルが存在しないこと自体は欠落ではない。

## GitHub設定とActions

Settings → Pages → Build and deployment → Sourceは**GitHub Actions**。2026-11-11にREST APIでworkflow方式とHTTPS有効を確認した。

main更新またはworkflow_dispatchで、Node.js 24のnpm ci → 型チェック → 全テスト → Pagesビルド／成果物検査 → configure-pages → artifactアップロード → deploy-pagesを実行する。ビルドはcontents/pagesのread、デプロイはcontents read・pages write・id-token writeを使用する。github-pages環境へ配信する。

失敗時はActionsのbuild/deployを別々に確認する。Pages未有効化なら管理画面のSourceを確認する。公開URLがまだ404の場合はデプロイ完了を待ち、マニフェストとJSへのアクセスを再確認する。設定や権限を無条件に広げない。

[GitHub公式のカスタムワークフロー](https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages)を参照。ローカルのマニフェスト検証成功はAppSource審査成功を意味しない。

## 今回の開始状態と履歴修正

2026-11-11開始時はローカルmainとorigin/mainが`8d891f3`で一致、未コミット変更なし。追加3コミットは指定の3ファイルのみ。以前の8コミットの先端は`f90ced1`。

開始時のActions（run 42940174993）はbuild成功、deployはPages未有効化により404。Pages APIも404、has_pages=falseだった。

| 元のコミット | メール修正後 |
|---|---|
| 47d0da2 | 5235fee |
| a2477ea | 6c7e6a5 |
| 8d891f3 | ffc8c07 |

この3件だけのAuthor/Committerを既存の安全なコミットと同じGitHub noreplyへ修正し、ツリー・メッセージ・日時を維持した。以前の8コミットはIDも内容も変更していない。事前のbundleはローカルのGit無視対象に保存し、公開しない。

pushは指定ブランチmainと開始時のリモートSHAを使うforce-with-leaseで行う。リモートに想定外の更新があれば上書きしない。古いコミットへの直接URL、Actionsの記録、キャッシュ、他者のclone等に旧情報が残る可能性があり、履歴修正だけで完全削除を保証しない。必要ならGitHub Supportへの削除相談が別途必要。

2026-11-11の監査は開発依存High 8件・Moderate 3件・Low 1件、実行用依存（--omit=dev）は0件。今回は依存の一括更新を行っていない。開発用のCLI等に残る指摘と、Pagesへ配信される実行用依存を区別する。

最初のローカルテストではGit無視対象の旧バックアップまで検出し、300件が二重に実行された。Vitestの対象をtests配下へ限定して、正規の300件だけを実行するよう修正した。Pages成果物検査はHTML最適化で省略される引用符にも対応するようDOMで参照を読み取る。

## ローカル最終検証（2026-11-11）

型チェック、全300テスト／24ファイル、Development／通常Production／Pages Productionビルド、開発用とPages用のマニフェスト検証が成功。各ビルドはエラー・警告0。通常ProductionとPagesの開発専用モジュール混入なし、DevelopmentではPoC・独立比較を含むことを確認した。HTML・XMLのリンク先ファイル、マニフェストのダウンロードリンク、ExcelApi 1.17、localhostマニフェスト不変も確認した。

公開後は上記Actionsのbuild/deploy成功と、配信されたマニフェスト・HTML・JS・画像のHTTP応答を照合する。これはExcel実機検証とは別の確認である。
