// 使い方:
//   node build.mjs                    … dist/ に出力（manifest の URL は BASE_URL 環境変数、既定は https://localhost:3000）
//   BASE_URL=https://example.github.io/refcite node build.mjs
import { readFileSync, writeFileSync, mkdirSync, cpSync, readdirSync } from 'node:fs'
import { build } from 'esbuild'
import { execFileSync } from 'node:child_process'

// 文献スタイル（表示名・言語・既定の引用番号の形）
const STYLE_META = [
  ['bmj', 'BMJ', 'superscript'],
  ['nlm-citation-sequence', 'Vancouver（NLM / Citing Medicine）', 'brackets'],
  ['elsevier-vancouver', 'Vancouver（Elsevier）', 'brackets'],
  ['american-medical-association', 'AMA（JAMA など）', 'superscript'],
  ['the-new-england-journal-of-medicine', 'NEJM', 'superscript'],
  ['the-lancet', 'Lancet', 'superscript'],
  ['american-heart-association', 'AHA（Circulation など）', 'superscript'],
  ['journal-of-the-american-college-of-cardiology', 'JACC', 'parens'],
  ['european-society-of-cardiology', 'ESC（European Heart Journal など）', 'superscript'],
  ['circulation-journal', 'Circulation Journal', 'superscript'],
  ['nature', 'Nature', 'superscript']
]

const styles = STYLE_META.map(([id, title, marker]) => {
  const xml = readFileSync(`src/styles/${id}.csl`, 'utf8')
  const locale = /default-locale="([^"]+)"/.exec(xml)?.[1] || 'en-US'
  return { id, title, marker, locale, xml }
})
const locales = { 'en-GB': readFileSync('src/styles/locales-en-GB.xml', 'utf8') }
writeFileSync('src/styles.generated.js',
  '// build.mjs が src/styles/ から生成するファイル（編集しない）\n' +
  `export const STYLES = ${JSON.stringify(styles)}\n` +
  `export const LOCALES = ${JSON.stringify(locales)}\n`)

if (process.argv.includes('--styles-only')) process.exit(0)

const BASE_URL = (process.env.BASE_URL || 'https://localhost:3000').replace(/\/$/, '')
mkdirSync('dist', { recursive: true })
await build({
  entryPoints: ['src/app.js'],
  bundle: true,
  minify: true,
  format: 'iife',
  target: 'es2020',
  outfile: 'dist/taskpane.js',
  legalComments: 'none'
})
cpSync('src/taskpane.html', 'dist/taskpane.html')
cpSync('src/taskpane.css', 'dist/taskpane.css')
cpSync('assets', 'dist/assets', { recursive: true })
// 形式の一覧（CSL_STYLES_DIR に citation-style-language/styles を置いたときだけ作る）
if (process.env.CSL_STYLES_DIR) {
  execFileSync(process.execPath, ['scripts/style-index.mjs', process.env.CSL_STYLES_DIR, 'dist/styles-index.json'], { stdio: 'inherit' })
} else {
  console.warn('CSL_STYLES_DIR がないので、形式の一覧（styles-index.json）は作りません')
}
writeFileSync('dist/manifest.xml', readFileSync('manifest.template.xml', 'utf8').replaceAll('{{BASE_URL}}', BASE_URL))
console.log('built dist/ for', BASE_URL, readdirSync('dist'))
