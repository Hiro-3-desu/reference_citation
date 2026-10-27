// 「AI の回答から追加」を、PubMed の応答を模したデータで通しで試す
import { createRequire } from 'node:module'
import { readFileSync } from 'node:fs'
import http from 'node:http'
import path from 'node:path'
const require = createRequire(import.meta.url)
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright')

const DB = {
  12517460: { title: 'Primary angioplasty versus intravenous thrombolytic therapy for acute myocardial infarction: a quantitative review of 23 randomised trials.', authors: ['Keeley EC', 'Boura JA', 'Grines CL'], source: 'Lancet', pubdate: '2003 Jan 4', volume: '361', issue: '9351', pages: '13-20', doi: '10.1016/S0140-6736(03)12113-7' },
  // AI が間違った PMID を書いた想定（別の論文）
  29138293: { title: 'Dapagliflozin in patients with heart failure and reduced ejection fraction.', authors: ['McMurray JJV'], source: 'N Engl J Med', pubdate: '2019 Nov 21', volume: '381', issue: '21', pages: '1995-2008', doi: '10.1056/NEJMoa1911303' },
  12930925: { title: 'A comparison of coronary angioplasty with fibrinolytic therapy in acute myocardial infarction.', authors: ['Andersen HR', 'Nielsen TT', 'Rasmussen K'], source: 'N Engl J Med', pubdate: '2003 Aug 21', volume: '349', issue: '8', pages: '733-42', doi: '10.1056/NEJMoa025142' },
  12684343: { title: 'Long distance transport for primary angioplasty vs immediate thrombolysis in acute myocardial infarction. Final results of the randomized national multicentre trial--PRAGUE-2.', authors: ['Widimský P'], source: 'Eur Heart J', pubdate: '2003 Jan', volume: '24', issue: '1', pages: '94-104', doi: '10.1016/s0195-668x(02)00468-2' },
  29999999: { title: 'Impact of regionalization of ST-segment-elevation myocardial infarction care on treatment times and outcomes for emergency medical services-transported patients.', authors: ['Jollis JG'], source: 'Circulation', pubdate: '2018 Jan 23', volume: '137', issue: '4', pages: '376-387', doi: '10.1161/CIRCULATIONAHA.117.032446' },
  31111111: { title: 'Coronary angioplasty in elderly patients.', authors: ['Other A'], source: 'Am J Cardiol', pubdate: '2001', volume: '1', issue: '1', pages: '1-2', doi: '' }
}
const summary = id => {
  const r = DB[id]
  return { uid: String(id), title: r.title, authors: r.authors.map(name => ({ name, authtype: 'Author' })), source: r.source, fulljournalname: r.source, pubdate: r.pubdate, sortpubdate: '', volume: r.volume, issue: r.issue, pages: r.pages, articleids: r.doi ? [{ idtype: 'doi', value: r.doi }] : [] }
}
const requests = []
function fakeEutils (url) {
  const u = new URL(url)
  requests.push(u.pathname.split('/').pop() + ' ' + (u.searchParams.get('term') || u.searchParams.get('id')))
  if (u.pathname.endsWith('esearch.fcgi')) {
    const term = u.searchParams.get('term').toLowerCase()
    let ids = []
    if (term.includes('[doi]')) ids = Object.keys(DB).filter(id => DB[id].doi && term.includes(DB[id].doi.toLowerCase()))
    else if (term.includes('regionalization[ti]')) ids = ['29999999']
    else if (term.includes('angioplasty[ti]') && term.includes('fibrinolytic[ti]')) ids = ['31111111', '12930925']
    return { esearchresult: { count: String(ids.length), idlist: ids } }
  }
  const ids = u.searchParams.get('id').split(',').filter(id => DB[id])
  return { result: { uids: ids, ...Object.fromEntries(ids.map(id => [id, summary(id)])) } }
}

const dist = path.resolve('dist')
const server = http.createServer((req, res) => {
  try { res.end(readFileSync(path.join(dist, decodeURIComponent(req.url.split('?')[0])))) } catch { res.statusCode = 404; res.end() }
}).listen(0)
const shots = process.env.SHOTS || '.'
const browser = await chromium.launch()
const page = await browser.newPage({ viewport: { width: 360, height: 900 } })
const errors = []
page.on('pageerror', e => errors.push(e.message))
await page.route('**/office.js', r => r.fulfill({ body: '', contentType: 'text/javascript' }))
await page.route('https://eutils.ncbi.nlm.nih.gov/**', r => r.fulfill({ json: fakeEutils(r.request().url()), headers: { 'access-control-allow-origin': '*' } }))
await page.route('https://api.crossref.org/**', r => r.fulfill({ status: 404, headers: { 'access-control-allow-origin': '*' } }))
await page.addInitScript({ path: 'test/word-mock.js' })
await page.goto(`http://127.0.0.1:${server.address().port}/taskpane.html`)
await page.click('[data-tab=add]')

for (const name of ['claude', 'codex']) {
  await page.fill('#aiInput', readFileSync(`test/fixtures/${name}.md`, 'utf8'))
  await page.click('#aiRun')
  await page.waitForFunction(() => !document.getElementById('aiRun').disabled)
  console.log(`[${name}]`, await page.textContent('#aiStatus'))
  for (const li of await page.locator('#aiList .rv').all()) {
    const chip = await li.locator('.chip').textContent()
    const checked = await li.locator('input:checked').evaluateAll(els => els.map(e => e.closest('.opt').querySelector('div:not(.meta):not(.jr)')?.textContent || 'どれでもない'))
    console.log('  ', chip.padEnd(8), '| 選択:', checked.join(' / ').slice(0, 60))
  }
  await page.screenshot({ path: `${shots}/ai-${name}.png`, fullPage: true })
  console.log('  button:', await page.textContent('#aiAdd'))
  await page.click('#aiAdd')
  console.log('  toast:', await page.textContent('#toast'))
}
await page.click('[data-tab=library]')
console.log('library:', await page.locator('#libList .ref .title').allTextContents())
console.log('requests:', requests.length, requests)
console.log('errors:', errors)
await browser.close(); server.close()
