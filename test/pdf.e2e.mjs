// 「自分の PDF」の登録・開く・変更・削除と、文書やバックアップでの扱いを試す
import { createRequire } from 'node:module'
import { readFileSync, writeFileSync } from 'node:fs'
import http from 'node:http'
import path from 'node:path'
import os from 'node:os'
const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright')

const dist = path.resolve('dist')
const server = http.createServer((req, res) => {
  try { res.end(readFileSync(path.join(dist, decodeURIComponent(req.url.split('?')[0])))) } catch { res.statusCode = 404; res.end() }
}).listen(0)
const base = `http://127.0.0.1:${server.address().port}/taskpane.html`
const browser = await chromium.launch()
const errors = []
async function open () {
  const page = await (await browser.newContext({ viewport: { width: 360, height: 900 } })).newPage()
  page.on('pageerror', e => errors.push(e.message))
  await page.route('**/office.js', r => r.fulfill({ body: '', contentType: 'text/javascript' }))
  await page.addInitScript({ path: 'test/word-mock.js' })
  await page.addInitScript(() => { window.__opened = []; window.Office.context.ui = { openBrowserWindow: u => window.__opened.push(u) } })
  await page.goto(base)
  return page
}
const page = await open()
await page.click('[data-tab=add]')
await page.setInputFiles('#fileInput', 'test/sample.ris')
await page.click('[data-tab=library]')
const row = i => page.locator('#libList .ref').nth(i)

// 登録：パソコン内のパス → エラー、https → 保存
await row(1).locator('.linkbtn', { hasText: 'PDF を登録' }).click()
await page.fill('.pdfedit input', 'C:\\Users\\me\\paper.pdf')
await page.click('.pdfedit button.primary')
console.log('local path :', await page.textContent('.pdfedit .err'))
await page.fill('.pdfedit input', 'https://onedrive.live.com/?id=ABC&cid=XYZ')
await page.click('.pdfedit button.primary')
console.log('saved      :', await row(1).locator('.links').innerText())
console.log('selected   :', await page.textContent('#selCount'))
await row(1).locator('.links a.mine').click()
console.log('opened     :', await page.evaluate(() => window.__opened))
await page.screenshot({ path: (process.env.SHOTS || '.') + '/pdf.png' })

// 引用しても、文書には PDF のリンクが入らない
await row(1).click()
await page.click('#btnCite')
const settings = await page.evaluate(() => window.__mock.settings())
console.log('in document:', JSON.stringify(settings).includes('onedrive'))

// バックアップ（JSON）に入り、別のパソコンで読み込むと戻る
await page.click('[data-tab=manage]')
await page.click('#btnExportJson')
const json = await page.inputValue('#exportBox')
console.log('in backup  :', json.includes('onedrive'))
const tmp = path.join(os.tmpdir(), 'refcite-backup.json')
writeFileSync(tmp, json)
const page2 = await open()
await page2.click('[data-tab=add]')
await page2.setInputFiles('#fileInput', tmp)
console.log('restore    :', await page2.textContent('#fileStatus'))
await page2.click('[data-tab=library]')
console.log('restored   :', (await page2.locator('#libList .links a.mine').count()) === 1)

// 削除
await page.click('[data-tab=library]')
await row(1).locator('.linkbtn', { hasText: 'PDF を変更' }).click()
await page.locator('.pdfedit button', { hasText: '削除' }).click()
console.log('deleted    :', await row(1).locator('.links a.mine').count() === 0)
console.log('errors     :', errors)
await browser.close(); server.close()
