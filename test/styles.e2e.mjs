// 「形式を追加」を通しで試す。CSL の取得は CSL_STYLES_DIR のローカルの styles から返す。
import { createRequire } from 'node:module'
import { readFileSync, existsSync } from 'node:fs'
import http from 'node:http'
import path from 'node:path'
const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright')
const STYLES = process.env.CSL_STYLES_DIR

const dist = path.resolve('dist')
const server = http.createServer((req, res) => {
  try { res.end(readFileSync(path.join(dist, decodeURIComponent(req.url.split('?')[0])))) } catch { res.statusCode = 404; res.end() }
}).listen(0)
const base = `http://127.0.0.1:${server.address().port}`
const shots = process.env.SHOTS || '.'
const browser = await chromium.launch()
const ctx = await browser.newContext({ viewport: { width: 360, height: 1000 } })
const errors = []
const fetched = []
await ctx.route('**/office.js', r => r.fulfill({ body: '', contentType: 'text/javascript' }))
await ctx.route('https://raw.githubusercontent.com/citation-style-language/**', r => {
  const m = /\/(styles|locales)\/[^/]+\/(.+)$/.exec(new URL(r.request().url()).pathname)
  const file = m[1] === 'styles' ? path.join(STYLES, m[2]) : null
  fetched.push(m[2])
  if (file && existsSync(file)) return r.fulfill({ body: readFileSync(file, 'utf8'), headers: { 'access-control-allow-origin': '*' } })
  return r.fulfill({ status: 404, headers: { 'access-control-allow-origin': '*' } })
})
const page = await ctx.newPage()
page.on('pageerror', e => errors.push(e.message))
await page.addInitScript({ path: 'test/word-mock.js' })
await page.goto(base + '/taskpane.html')
const docText = () => page.evaluate(() => window.__mock.text())
const bibFirst = async () => (await docText()).split('\n').find(l => l.startsWith('  ')).trim()

// 文献を 2 件引用して文献リストを作る
await page.click('[data-tab=add]')
await page.setInputFiles('#fileInput', 'test/sample.ris')
await page.click('[data-tab=library]')
const items = page.locator('#libList .ref')
await items.nth(0).click(); await items.nth(1).click()
await page.click('#btnCite')
await page.evaluate(() => window.__mock.cursorAtEndOf(2))
await page.click('[data-tab=format]')
await page.click('#btnBib')
console.log('BMJ       :', await bibFirst())

// 雑誌名で検索して追加（独立した形式）
await page.fill('#styleTerm', 'europace')
await page.click('#styleForm button')
await page.waitForTimeout(200)
console.log('search    :', await page.textContent('#styleStatus'), await page.locator('#styleResults .name').first().textContent())
await page.locator('#styleResults button').first().click()
await page.waitForTimeout(400)
console.log('toast     :', await page.textContent('#toast'))
console.log('Europace  :', await bibFirst(), '| cite:', (await docText()).split('\n')[0])

// 雑誌ごとの形式（元の形式を取りに行く）
await page.fill('#styleTerm', 'Circulation')
await page.click('#styleForm button')
await page.waitForTimeout(200)
const names = await page.locator('#styleResults .name').allTextContents()
console.log('search    :', names.slice(0, 3))
await page.locator('#styleResults button').first().click()
await page.waitForTimeout(400)
console.log('Circulation:', await bibFirst())
console.log('select    :', await page.locator('#styleSelect optgroup[label="追加した形式"] option').allTextContents(), '=', await page.$eval('#styleSelect', s => s.selectedOptions[0].text))
await page.screenshot({ path: `${shots}/styles.png`, fullPage: true })

// 使えない形式（著者名と年）の CSL ファイル
await page.setInputFiles('#styleFile', path.join(STYLES, 'apa.csl'))
await page.waitForTimeout(200)
console.log('apa file  :', await page.textContent('#toast'))
// 雑誌ごとの CSL ファイル
await page.setInputFiles('#styleFile', path.join(STYLES, 'dependent', 'american-heart-journal.csl'))
await page.waitForTimeout(400)
console.log('AHJ file  :', await page.textContent('#toast'))
console.log('AHJ       :', await bibFirst())

// 別のパソコン（ライブラリも追加した形式もない）で同じ文書を開いた想定
const settings = await page.evaluate(() => window.__mock.settings())
const page2 = await (await browser.newContext()).newPage()
await page2.route('**/office.js', r => r.fulfill({ body: '', contentType: 'text/javascript' }))
await page2.addInitScript({ path: 'test/word-mock.js' })
await page2.addInitScript(s => { window.__preset = s }, settings)
await page2.addInitScript(() => {
  const set = window.Office.context.document.settings.set
  for (const [k, v] of Object.entries(window.__preset)) set(k, JSON.parse(v))
})
await page2.goto(base + '/taskpane.html')
await page2.waitForTimeout(200)
await page2.click('[data-tab=format]')
console.log('other PC  :', await page2.$eval('#styleSelect', s => s.selectedOptions[0]?.text))

// 削除
await page.locator('#customStyles button').first().click()
console.log('after del :', await page.locator('#customStyles .name').allTextContents())
console.log('fetched   :', fetched)
console.log('errors    :', errors)
await browser.close(); server.close()
