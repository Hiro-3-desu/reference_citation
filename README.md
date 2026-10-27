# RefCite

Word で使う文献管理アドインです。EndNote の「Cite While You Write」のように、Word の右側のパネルから引用を挿入し、番号の振り直しと文献リストの作成を自動で行います。

## できること

- **文献の追加**: PubMed 検索、PMID / DOI の入力、RIS・BibTeX・CSL-JSON ファイルの読み込み（EndNote は RIS で書き出し）
- **引用の挿入**: 文献を選んで「引用を挿入」。番号は文書中の初出順に自動で振られ、範囲はまとめて表示（例: 1–3,5）
- **文献リスト**: BMJ、Vancouver（NLM / Elsevier）、AMA、NEJM、Lancet、AHA、JACC、ESC、Circulation Journal、Nature
- **引用の編集**: 引用から文献を外す。カーソルを引用の中に置いて「引用を挿入」すると文献を追加
- **投稿用の変換**: 引用と文献リストをふつうの文字に変換

## データの保存場所

- **ライブラリ**: そのパソコンの Word アドインの中（ブラウザの localStorage）。「管理」タブから JSON / RIS で書き出せます。
- **文書で使った文献と引用**: Word 文書の中（document settings）。文書だけをほかのパソコンに渡しても、番号の更新や文献リストの再作成ができます。

外部に送るのは、PubMed（NCBI E-utilities）と Crossref への検索語・PMID・DOI だけです。

## 仕組み

- 引用は Word のコンテンツコントロール（タグ `rmc:<key>`）、文献リストはタグ `rmbib` のコンテンツコントロールです。
- 文献の整形は [citation-js](https://citation.js.org/)（citeproc-js）と、[CSL スタイル](https://github.com/citation-style-language/styles)（`src/styles/`、CC BY-SA 3.0）を使っています。

| ファイル | 内容 |
|---|---|
| `src/core.js` | 文献データの変換、番号付け、文献リストの整形（Word に依存しない） |
| `src/word.js` | Word とのやりとり（引用の挿入、番号の更新、文献リスト） |
| `src/sources.js` | PubMed / Crossref からの取得、ライブラリの保存 |
| `src/app.js`, `src/taskpane.html`, `src/taskpane.css` | 作業ウィンドウの画面 |
| `manifest.template.xml` | アドインのマニフェスト（`{{BASE_URL}}` を公開先に置き換えて使う） |

## ビルドとテスト

```sh
npm install
BASE_URL=https://<ユーザー名>.github.io/<リポジトリ名> npm run build   # dist/ に出力
npm test                                                             # 番号付け・取り込み・整形のテスト
node test/ui.e2e.mjs                                                 # Word の模擬環境で画面を通しで操作（Playwright が必要）
```

`dist/` を HTTPS で公開し、`dist/manifest.xml` を Word に登録すると使えます。
`main` に push すると、GitHub Actions（`.github/workflows/pages.yml`）がテストとビルドをして GitHub Pages に公開します。

公開先: https://hiro-3-desu.github.io/reference_citation/
マニフェスト: https://hiro-3-desu.github.io/reference_citation/manifest.xml

## Word への登録（Windows の Microsoft 365）

1. `manifest.xml` を https://hiro-3-desu.github.io/reference_citation/manifest.xml からダウンロードし、`C:\RefCite\manifest.xml` に置きます（同期していないフォルダに置く）。
2. Word をすべて閉じ、ターミナル（PowerShell）で次を実行します。

   ```
   reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v RefCite /t REG_SZ /d "C:\RefCite\manifest.xml" /f
   ```

3. Word を開くと「ホーム」タブの右端に「RefCite」ボタンが出ます。

登録をやめるときは `reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v RefCite /f` を実行します。

「信頼できるアドイン カタログ」に共有フォルダーを登録する方法もありますが、環境によっては「共有フォルダー」タブが表示されません。

Word on the web では「挿入」→「アドイン」→「その他のアドイン」→「マイ アドイン」→「マイ アドインのアップロード」から `manifest.xml` を選ぶだけで使えます。

## パソコンを買い替えるとき

ライブラリはパソコンごとに保存されます。「管理」タブの「ライブラリを書き出す（JSON）」で書き出したファイルを OneDrive などに保存しておき、新しいパソコンで「追加」タブの「ファイルから読み込む」で戻します。
