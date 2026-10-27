// 文献のリンク（PubMed・出版社・PMC）と「PDF を探す」を試す。OpenAlex の応答は模擬。
import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import http from 'node:http'
import path from 'node:path'
const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright')

const dist = path.resolve('dist')
const server = http.createServer((req, res) => {
  try { res.end(readFileSync(path.join(dist, decodeURIComponent(req.url.split('?')[0])))) } catch { res.statusCode = 404; res.end() }
}).listen(0)
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 360, height: 800 } })
const errors = []
page.on('pageerror', e => errors.push(e.message))
await page.route('**/office.js', r => r.fulfill({ body: '', contentType: 'text/javascript' }))
const openalex = []
await page.route('https://api.openalex.org/**', r => {
  const url = decodeURIComponent(r.request().url()).toLowerCase()
  openalex.push(url.replace('https://api.openalex.org/works/', ''))
  const cors = { 'access-control-allow-origin': '*' }
  if (url.includes('nejmoa025142')) return r.fulfill({ headers: cors, json: { open_access: { is_oa: true, oa_url: 'https://example.org/oa' }, best_oa_location: { pdf_url: 'https://example.org/danami2.pdf' } } })
  if (url.includes('12113-7')) return r.fulfill({ headers: cors, json: { open_access: { is_oa: false }, best_oa_location: null } })
  return r.fulfill({ status: 404, headers: cors })
})
await page.addInitScript({ path: 'test/word-mock.js' })
await page.addInitScript(() => {
  window.__opened = []
  window.Office.context.ui = { openBrowserWindow: url => window.__opened.push(url) }
})
await page.goto(`http://127.0.0.1:${server.address().port}/taskpane.html`)
await page.click('[data-tab=add]')
await page.setInputFiles('#fileInput', 'test/sample.ris')
await page.click('[data-tab=library]')

const rows = page.locator('#libList .ref')
for (let i = 0; i < 3; i++) console.log('links', i, await rows.nth(i).locator('.links').innerText())
// 出版社リンク（選択は切り替わらない）
await rows.nth(1).locator('.links a', { hasText: '出版社' }).click()
console.log('selected after link click:', await page.textContent('#selCount'))
// PDF を探す：見つかる → 開く
await rows.nth(1).locator('.linkbtn').first().click()
await page.waitForTimeout(200)
console.log('found  :', await rows.nth(1).locator('.linkbtn').first().textContent())
// 見つからない
await rows.nth(2).locator('.linkbtn').first().click()
await page.waitForTimeout(200)
console.log('none   :', await rows.nth(2).locator('.linkbtn').first().textContent(), '|', await page.textContent('#toast'))
// もう一度押しても問い合わせない（覚えている）
await rows.nth(1).locator('.linkbtn').first().click()
console.log('opened :', await page.evaluate(() => window.__opened))
console.log('openalex requests:', openalex)
await page.screenshot({ path: (process.env.SHOTS || '.') + '/links.png' })
console.log('errors :', errors)
await browser.close(); server.close()
