// dist/ をブラウザで開き、Word の模擬環境で 文献の読み込み → 引用 → 文献リスト → 編集 を通しで試す
import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import http from 'node:http'
import path from 'node:path'
const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright')

const dist = path.resolve('dist')
const server = http.createServer((req, res) => {
  try {
    const p = path.join(dist, decodeURIComponent(req.url.split('?')[0]))
    res.end(readFileSync(p))
  } catch { res.statusCode = 404; res.end() }
}).listen(0)
const base = `http://127.0.0.1:${server.address().port}`
const shots = process.env.SHOTS || '.'

const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 360, height: 760 } })
const errors = []
page.on('pageerror', e => errors.push(e.message))
page.on('console', m => { if (m.type() === 'error') errors.push(m.text()) })
await page.route('**/office.js', r => r.fulfill({ body: '', contentType: 'text/javascript' }))
await page.addInitScript({ path: 'test/word-mock.js' })
await page.goto(base + '/taskpane.html')
await page.waitForTimeout(300)
const step = async name => { await page.screenshot({ path: `${shots}/${name}.png` }); }
const docText = () => page.evaluate(() => window.__mock.text())

// 1. RIS を読み込む
await page.click('[data-tab=add]')
await page.setInputFiles('#fileInput', 'test/sample.ris')
await page.waitForTimeout(200)
console.log('import:', await page.textContent('#fileStatus'))
await page.setInputFiles('#fileInput', 'test/sample.ris')
await page.waitForTimeout(200)
console.log('reimport:', await page.textContent('#fileStatus'))
await step('1-add')

// 2. 1 段落目に文献 2 件、2 段落目に 1 件（うち 1 件は重複）を引用
await page.click('[data-tab=library]')
await step('2-library')
const items = page.locator('#libList .ref')
console.log('library items:', await items.count())
await items.nth(2).click(); await items.nth(0).click()
await page.click('#btnCite')
await page.waitForTimeout(200)
await page.evaluate(() => window.__mock.cursorAtEndOf(1))
await items.nth(1).click(); await items.nth(0).click()
await page.click('#btnCite')
await page.waitForTimeout(200)
console.log('toast:', await page.textContent('#toast'))
await step('3-cited')

// 3. 文献リストを 3 段落目の後ろに挿入
await page.evaluate(() => window.__mock.cursorAtEndOf(2))
await page.click('[data-tab=format]')
await page.click('#btnBib')
await page.waitForTimeout(200)
console.log('status:', await page.textContent('#fmtStatus'))
console.log('--- doc (BMJ) ---\n' + await docText())

// 4. スタイル変更
await page.selectOption('#styleSelect', 'nlm-citation-sequence')
await page.waitForTimeout(200)
console.log('--- doc (Vancouver) ---\n' + await docText())
await step('4-format')

// 5. 1 つ目の引用を編集して 1 件外す
await page.evaluate(() => window.__mock.cursorInsideCitation(0))
await page.click('#btnEditCite')
await page.waitForTimeout(200)
console.log('editor refs:', await page.locator('#citeRefs .ref').count())
await step('5-edit')
await page.locator('#citeRefs .remove').first().click()
await page.waitForTimeout(200)
console.log('--- doc (after remove) ---\n' + await docText())

// 6. 引用中のみ表示
await page.evaluate(() => window.__mock.clearInside())
await page.click('[data-tab=library]')
await page.check('#onlyCited')
console.log('cited only:', await items.count(), await page.locator('#libList .badge').allTextContents())

// 7. 投稿用に変換
await page.click('[data-tab=manage]')
await page.click('#btnPlain'); await page.click('#btnPlainYes')
await page.waitForTimeout(200)
console.log('unlink:', await page.textContent('#plainStatus'))
console.log('--- doc (plain) ---\n' + await docText())

console.log('errors:', errors)
await browser.close(); server.close()
