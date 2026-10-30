# 開発環境とローカル実行

## 前提

- Node.js 24 LTS（Microsoft公式テンプレートのNode >=20 <25と整合）、npm、VS Code。
- ExcelApi 1.17に対応するMicrosoft 365 Excel。公式対応表ではWindows Version 2302（Build 16130.20332）以上、Mac 16.70以上、Webは対応。実際にはランタイムのisSetSupportedで判定する。Windows・Macは実機未検証。Webの確認済み範囲は[現在の状況](stabilization-status.md)を参照。
- Office.jsはMicrosoftの公式CDNから実行時に読み込む。CDNへの通信が必要。
- このマニフェストはlocalhost開発用。公開用URL・AppSource登録は未設定。

`npm ci`でpackage-lock.jsonに固定された依存関係を導入する。グローバルなYeomanインストールは不要。公式React Task Paneテンプレートのwebpack・XML構造を参考に構成し、生成ツールそのものは実行していない。

今回の環境では、OSへインストールせず、公式配布のNode.js 24.22.1を`.local-tools`へ配置して検証した。これはGit管理対象外で、他の環境は通常のNode.js 24を利用できる。この環境のPowerShellではセッション中のみ以下を設定できる。

```powershell
$env:PATH = "$PWD\.local-tools\node-v24.22.1-win-x64;" + $env:PATH
npm.cmd ci
```

## HTTPS証明書（環境変更を伴うため明示実行）

`npm run build`、`npm test`、`npm run dev`は証明書を自動作成・信頼登録しない。サーバーは既存の証明書だけを読む。

環境所有者が信頼ストアへの登録を了承した場合のみ、次を実行する。

```sh
npm run certs
```

これはMicrosoftの`office-addin-dev-certs install`を実行し、ユーザーの信頼ストアと`~/.office-addin-dev-certs`を変更する。今回の作業では実行していない。

既に別途準備した証明書を利用する場合は、`localhost.key`と`localhost.crt`を置いたディレクトリを`SHEETDOCK_CERT_DIR`に設定する。localhostのSAN、期限、Excel内WebViewでの信頼を確認する。秘密鍵をコミットしない。

```powershell
$env:SHEETDOCK_CERT_DIR = "$PWD\.certificates"
npm.cmd run dev
```

標準では`~/.office-addin-dev-certs/localhost.key`と`localhost.crt`を読む。証明書がなければ説明付きエラーで停止する。ホストはlocalhost、ポート3000固定。別ポートを使う場合はマニフェストのURLも合わせて変更する。

## Windows Excelへのサイドロード

環境所有者がOfficeへの登録変更を了承した後、対応するMicrosoft 365 Excelで実施する。

1. 専用の空のテストブックを用意する。業務・個人ブックは開かない。
2. `npm run dev`でHTTPSサーバーを起動し、ブラウザーで`https://localhost:3000/help.html`へ警告なしでアクセスできることを確認。
3. 別のターミナルで`npm run sideload`を実行する。開発用登録・Office起動・環境によってWebViewのループバック設定等を求められるため、表示内容を確認する。
4. ホームタブの「SheetDockを開く」を選択する。
5. [テスト手順](testing.md)の項目を確認する。
6. 終了時は`npm run stop`で開発用登録を解除し、devサーバーをCtrl+Cで停止する。

CLIはMicrosoft公式`office-addin-debugging`を使用する。サーバーは別途起動する。サイドロードとstopはOffice設定を変更するので、今回の作業ではどちらも実行していない。

## Web・Mac

Webは専用の新規テストブックで、Excelのアドイン管理から「自分のアドインのアップロード」を選び`manifest.xml`を指定する。テナントのポリシーにより利用できないことがある。ブラウザーがlocalhost HTTPSへアクセスできる必要がある。

MacはMicrosoftのMac向けサイドロード手順を利用する。証明書の信頼・マニフェスト配置先はWindowsと異なる。Macは実機未検証。Webはユーザーによるサイドロードと基本機能の動作確認済み。

## 検証コマンド

```sh
npm run typecheck
npm test
npm run build
npm run validate
```

`dist/manifest.xml`もlocalhostのまま。productionビルドはJSの最適化を意味し、配布可能なURLへの変更を意味しない。`npm run validate`は開発用マニフェストの検証で、AppSource審査を代替しない。

## 参照（2026-10-27確認）

- [公式React Excelクイックスタート](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-react)
- [公式テンプレート](https://github.com/OfficeDev/Office-Addin-TaskPane-React)
- [マニフェスト検証](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest)
- [Webへのサイドロード](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- [Macへのサイドロード](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)
- [Excel API要件](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-requirement-sets)

## 開発ツールの既知の注意点

今回の検証時点で、公式サイドロードCLIの推移依存にnpm監査のHigh 8件が残る。タスクペインの実行用依存は0件。サイドロード前にツール更新と監査を再確認する。認証関連依存のインストールスクリプトはnpmの保護により未実行のものがあり、認証で必要になった場合は対象と処理内容を確認してから許可する。詳細は[移行記録](sheetdock-web-phase1a.md)を参照。

## Phase 1-B②の技術検証

`npm run dev`は基本UIの下に開発専用PoCを追加します。起動・画面表示だけでブックを書き換えることはありません。前回と同じmanifest.xmlを利用でき、ID・権限・API要件は変更していません。

[検証手順と制約](sheetdock-web-phase1b.md)に従って専用ブックを用意してください。productionビルドにはPoCを含めません。
