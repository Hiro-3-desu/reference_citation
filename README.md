# RefCite

Word で使う文献管理アドインです。EndNote の「Cite While You Write」のように、Word の右側のパネルから引用を挿入し、番号の振り直しと文献リストの作成を自動で行います。

## できること

- **AI の回答から追加**: Claude や Codex の回答をそのまま貼り付けると、PMID・DOI・PubMed / DOI のリンク・表の PMID 列を拾って PubMed / Crossref で確認する。ID のない文献はタイトルで検索して候補を出す。AI が書いた内容と ID の論文が食い違うもの（ID の間違い）や、見つからないもの（実在しない可能性）は印を付けて知らせる
- **文献の追加**: PubMed 検索、RIS・BibTeX・CSL-JSON ファイルの読み込み（EndNote は RIS で書き出し）
- **引用の挿入**: 文献を選んで「引用を挿入」。番号は文書中の初出順に自動で振られ、範囲はまとめて表示（例: 1–3,5）
- **文献リスト**: BMJ、Vancouver（NLM / Elsevier）、AMA、NEJM、Lancet、AHA、JACC、ESC、Circulation Journal、Nature
- **形式の追加**: 一覧にない雑誌は「文献リスト」タブで雑誌名を検索して追加（CSL スタイル集のうち番号式の約 5,500 誌）。Zotero Style Repository などの .csl ファイルも読み込める。追加した形式はそのパソコンに保存され、使った文書の中にも保存されるので、ほかのパソコンで文書を開いても同じ形式で更新できる。著者名と年で引用する形式や、文献リストを著者名順に並べる形式は、番号と合わないので受け付けない
- **論文を開く**: 各文献に PubMed・出版社（DOI）・全文（PMC）のリンク。「PDF を探す」で OpenAlex に DOI / PMID を問い合わせ、無料で読める版（オープンアクセス）があれば開く
- **引用の編集**: 引用から文献を外す。カーソルを引用の中に置いて「引用を挿入」すると文献を追加
- **リンクを外す（投稿前）**: 引用と文献リストの RefCite とのリンクを外し、ただの文字にする

## データの保存場所

- **ライブラリ**: そのパソコンの Word アドインの中（ブラウザの localStorage）。「管理」タブから JSON / RIS で書き出せます。
- **文書で使った文献と引用**: Word 文書の中（document settings）。文書だけをほかのパソコンに渡しても、番号の更新や文献リストの再作成ができます。

外部に送るのは、PubMed（NCBI E-utilities）・Crossref・OpenAlex への検索語・PMID・DOI だけです。

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
BASE_URL=https://<ユーザー名>.github.io/<リポジトリ名> CSL_STYLES_DIR=<citation-style-language/styles のフォルダ> npm run build   # dist/ に出力
npm test                                                             # 番号付け・取り込み・整形のテスト
node test/ui.e2e.mjs                                                 # Word の模擬環境で画面を通しで操作（Playwright が必要）
node test/ai.e2e.mjs                                                 # 「AI の回答から追加」を PubMed の模擬応答で通しで操作
CSL_STYLES_DIR=<styles のフォルダ> node test/styles.e2e.mjs          # 「形式を追加」を通しで操作
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
