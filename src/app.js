/* global Office */
import { Cite } from '@citation-js/core'
import { STYLE_LIST, MARKERS, refKey, shortAuthor, year, journal, parseImportText } from './core.js'
import { pubmedSearch, lookupIds, loadLibrary, saveLibrary, addToLibrary } from './sources.js'
import * as word from './word.js'

const $ = id => document.getElementById(id)

const app = {
  library: loadLibrary(),
  doc: word.loadDocState(),
  selected: new Set(), // ライブラリで選んだ文献の id
  numbers: new Map(), // refKey → この文書での番号
  pmResults: [],
  pmSelected: new Set(),
  editing: null // 編集中の引用 { key, refIds }
}

// ---------- 共通 ----------

let toastTimer
function toast (msg, isError = false) {
  const t = $('toast')
  t.textContent = msg
  t.classList.toggle('error', isError)
  t.hidden = false
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => { t.hidden = true }, isError ? 6000 : 3000)
}

function errorMessage (e) {
  if (e?.message === 'Failed to fetch') return 'インターネットに接続できませんでした。'
  return e?.message || String(e)
}

function persistLibrary () {
  if (!saveLibrary(app.library)) toast('ライブラリを保存できませんでした。容量が足りない可能性があります。', true)
}

function setNumbers (usedIds) {
  app.numbers = new Map(usedIds.map((id, i) => [refKey(app.doc.refs[id]), i + 1]))
}

function requireWord () {
  if (word.inWord()) return true
  toast('この操作は Word の中でだけ使えます。', true)
  return false
}

// 文献 1 件分の <li> を作る（テキストはすべて textContent で入れる）
function refItem (r, { checkbox = false, checked = false, badge = '', button = null } = {}) {
  const li = document.createElement('li')
  li.className = 'ref' + (checked ? ' checked' : '')
  if (checkbox) {
    const cb = document.createElement('input')
    cb.type = 'checkbox'
    cb.checked = checked
    cb.tabIndex = -1
    li.append(cb)
  }
  const meta = document.createElement('div')
  meta.className = 'meta'
  const b = document.createElement('b')
  b.textContent = shortAuthor(r)
  const y = document.createElement('span')
  y.textContent = year(r)
  meta.append(b, y)
  if (badge) {
    const s = document.createElement('span')
    s.className = 'badge'
    s.textContent = badge
    meta.append(s)
  }
  const title = document.createElement('div')
  title.className = 'title'
  title.textContent = r.title || '（タイトルなし）'
  const jr = document.createElement('div')
  jr.className = 'jr'
  jr.textContent = [journal(r), r.PMID ? `PMID ${r.PMID}` : ''].filter(Boolean).join(' · ')
  li.append(meta, title, jr)
  if (button) li.append(button)
  return li
}

// ---------- タブ ----------

for (const btn of document.querySelectorAll('[data-tab]')) {
  btn.addEventListener('click', () => {
    for (const b of document.querySelectorAll('[data-tab]')) b.setAttribute('aria-selected', String(b === btn))
    for (const s of document.querySelectorAll('.tab')) s.hidden = s.id !== 'tab-' + btn.dataset.tab
  })
}

// ---------- 文献タブ ----------

function renderLibrary () {
  const q = $('libSearch').value.trim().toLowerCase()
  const onlyCited = $('onlyCited').checked
  const list = $('libList')
  list.replaceChildren()
  let shown = 0
  const items = app.library
    .map(r => ({ r, n: app.numbers.get(refKey(r)) }))
    .filter(({ r, n }) => {
      if (onlyCited && !n) return false
      if (!q) return true
      const hay = [r.title, journal(r), r['container-title'], year(r), r.PMID, r.DOI,
        ...(r.author || []).map(a => a.family || a.literal || '')].join(' ').toLowerCase()
      return q.split(/\s+/).every(w => hay.includes(w))
    })
  if (onlyCited) items.sort((a, b) => a.n - b.n)
  for (const { r, n } of items) {
    const li = refItem(r, { checkbox: true, checked: app.selected.has(r.id), badge: n ? `[${n}]` : '' })
    li.addEventListener('click', () => {
      if (app.selected.has(r.id)) app.selected.delete(r.id)
      else app.selected.add(r.id)
      renderLibrary()
    })
    list.append(li)
    shown++
  }
  $('libEmpty').hidden = app.library.length > 0
  $('libCount').textContent = app.library.length ? `${shown} / ${app.library.length} 件` : ''
  const n = app.selected.size
  $('selCount').textContent = `${n} 件選択`
  $('btnCite').disabled = n === 0
  $('btnDelete').disabled = n === 0
}

$('libSearch').addEventListener('input', renderLibrary)
$('onlyCited').addEventListener('change', renderLibrary)

$('btnCite').addEventListener('click', async () => {
  if (!requireWord()) return
  // 同じ論文がすでに文書に入っていれば、その id を使う（番号が二重にならないように）
  const docByKey = new Map(Object.values(app.doc.refs).map(r => [refKey(r), r]))
  const refs = app.library
    .filter(r => app.selected.has(r.id))
    .map(r => {
      const { _added, ...csl } = r
      return docByKey.get(refKey(r)) || csl
    })
  $('btnCite').disabled = true
  try {
    const res = await word.insertCitation(app.doc, refs)
    setNumbers(res.usedIds)
    app.selected.clear()
    toast(`引用を挿入しました（文献 ${res.references} 件）`)
  } catch (e) {
    toast('引用を挿入できませんでした: ' + errorMessage(e), true)
  }
  renderLibrary()
})

$('btnDelete').addEventListener('click', () => { $('deleteConfirm').hidden = false })
$('btnDeleteNo').addEventListener('click', () => { $('deleteConfirm').hidden = true })
$('btnDeleteYes').addEventListener('click', () => {
  const n = app.selected.size
  app.library = app.library.filter(r => !app.selected.has(r.id))
  app.selected.clear()
  persistLibrary()
  $('deleteConfirm').hidden = true
  renderLibrary()
  toast(`${n} 件を削除しました`)
})

// ---------- 追加タブ ----------

function addItems (items) {
  const added = addToLibrary(app.library, items)
  persistLibrary()
  renderLibrary()
  return added
}

function renderPubmed () {
  const list = $('pmList')
  list.replaceChildren()
  const inLib = new Set(app.library.map(refKey))
  for (const r of app.pmResults) {
    const have = inLib.has(refKey(r))
    const li = refItem(r, { checkbox: !have, checked: app.pmSelected.has(r.id), badge: have ? '追加済み' : '' })
    if (!have) {
      li.addEventListener('click', () => {
        if (app.pmSelected.has(r.id)) app.pmSelected.delete(r.id)
        else app.pmSelected.add(r.id)
        renderPubmed()
      })
    }
    list.append(li)
  }
  $('pmActions').hidden = app.pmResults.length === 0
  $('pmAdd').disabled = app.pmSelected.size === 0
  $('pmAdd').textContent = app.pmSelected.size ? `選択した ${app.pmSelected.size} 件を追加` : '選択した文献を追加'
}

$('pmForm').addEventListener('submit', async e => {
  e.preventDefault()
  const term = $('pmTerm').value.trim()
  if (!term) return
  $('pmStatus').textContent = '検索しています…'
  try {
    const { total, items } = await pubmedSearch(term)
    app.pmResults = items
    app.pmSelected.clear()
    $('pmStatus').textContent = total
      ? `${total.toLocaleString()} 件中、関連の高い ${items.length} 件を表示しています。`
      : '見つかりませんでした。'
  } catch (err) {
    app.pmResults = []
    $('pmStatus').textContent = '検索できませんでした: ' + errorMessage(err)
  }
  renderPubmed()
})

$('pmAdd').addEventListener('click', () => {
  const items = app.pmResults.filter(r => app.pmSelected.has(r.id))
  const added = addItems(items)
  app.pmSelected.clear()
  renderPubmed()
  toast(`${added} 件をライブラリに追加しました`)
})

$('idAdd').addEventListener('click', async () => {
  const text = $('idInput').value.trim()
  if (!text) return
  $('idStatus').textContent = '取得しています…'
  try {
    const { found, failed } = await lookupIds(text)
    const added = addItems(found)
    const dup = found.length - added
    $('idStatus').textContent = [
      `${added} 件を追加しました。`,
      dup ? `${dup} 件はすでにライブラリにありました。` : '',
      failed.length ? `見つからなかったもの: ${failed.join(', ')}` : ''
    ].filter(Boolean).join(' ')
    if (!failed.length) $('idInput').value = ''
  } catch (err) {
    $('idStatus').textContent = '取得できませんでした: ' + errorMessage(err)
  }
})

$('fileInput').addEventListener('change', async e => {
  const file = e.target.files[0]
  e.target.value = ''
  if (!file) return
  try {
    const items = parseImportText(await file.text())
    const added = addItems(items)
    $('fileStatus').textContent = `${file.name}: ${items.length} 件を読み込み、${added} 件を追加しました（重複 ${items.length - added} 件）。`
  } catch (err) {
    $('fileStatus').textContent = `${file.name} を読み込めませんでした。RIS・BibTeX・CSL-JSON のどれかか確認してください。（${errorMessage(err)}）`
  }
})

// ---------- 文献リストタブ ----------

function renderFormat () {
  const ss = $('styleSelect')
  ss.replaceChildren(...STYLE_LIST.map(s => new Option(s.title, s.id)))
  ss.value = app.doc.style
  const ms = $('markerSelect')
  ms.replaceChildren(...Object.entries(MARKERS).map(([k, m]) => new Option(m.label, k)))
  ms.value = app.doc.marker
}

async function applyFormat (message) {
  if (!word.inWord()) return
  try {
    await word.saveDocState(app.doc)
    const res = await word.refresh(app.doc)
    setNumbers(res.usedIds)
    renderLibrary()
    $('fmtStatus').textContent = message(res)
  } catch (e) {
    $('fmtStatus').textContent = '更新できませんでした: ' + errorMessage(e)
  }
}

$('styleSelect').addEventListener('change', () => {
  app.doc.style = $('styleSelect').value
  app.doc.marker = STYLE_LIST.find(s => s.id === app.doc.style)?.marker || app.doc.marker
  $('markerSelect').value = app.doc.marker
  applyFormat(res => `${$('styleSelect').selectedOptions[0].text} に変更しました（引用 ${res.citations} か所・文献 ${res.references} 件）。`)
})

$('markerSelect').addEventListener('change', () => {
  app.doc.marker = $('markerSelect').value
  applyFormat(() => '引用番号の形を変更しました。')
})

$('btnBib').addEventListener('click', async () => {
  if (!requireWord()) return
  try {
    const res = await word.insertBibliography(app.doc)
    setNumbers(res.usedIds)
    renderLibrary()
    $('fmtStatus').textContent = res.existed
      ? 'この文書にはすでに文献リストがあるので、更新しました。'
      : `文献リストを挿入しました（${res.references} 件）。`
  } catch (e) {
    $('fmtStatus').textContent = '文献リストを挿入できませんでした: ' + errorMessage(e)
  }
})

$('btnRefresh').addEventListener('click', () => {
  if (!requireWord()) return
  applyFormat(res => `更新しました（引用 ${res.citations} か所・文献 ${res.references} 件${res.bibliographies ? '' : '。文献リストはまだ挿入されていません'}）。`)
})

function renderCiteEditor () {
  const list = $('citeRefs')
  list.replaceChildren()
  if (!app.editing) {
    $('citeEditor').hidden = true
    return
  }
  for (const id of app.editing.refIds) {
    const r = app.doc.refs[id]
    if (!r) continue
    const btn = document.createElement('button')
    btn.className = 'remove'
    btn.textContent = '外す'
    btn.addEventListener('click', async () => {
      const rest = app.editing.refIds.filter(x => x !== id)
      try {
        const res = await word.setCitationRefs(app.doc, app.editing.key, rest)
        setNumbers(res.usedIds)
        app.editing = rest.length ? { ...app.editing, refIds: rest } : null
        renderCiteEditor()
        renderLibrary()
        toast(rest.length ? '文献を外しました' : '引用を削除しました')
      } catch (e) {
        toast('変更できませんでした: ' + errorMessage(e), true)
      }
    })
    list.append(refItem(r, { badge: app.numbers.get(refKey(r)) ? `[${app.numbers.get(refKey(r))}]` : '', button: btn }))
  }
  $('citeEditor').hidden = false
}

$('btnEditCite').addEventListener('click', async () => {
  if (!requireWord()) return
  try {
    app.editing = await word.getSelectedCitation(app.doc)
    if (!app.editing) toast('カーソルが引用番号の中にありません。番号をクリックしてから押してください。', true)
  } catch (e) {
    toast(errorMessage(e), true)
  }
  renderCiteEditor()
})

// ---------- 管理タブ ----------

$('btnImportDoc').addEventListener('click', () => {
  const refs = Object.values(app.doc.refs)
  const added = addItems(refs.map(r => ({ ...r })))
  $('docStatus').textContent = refs.length
    ? `文書の文献 ${refs.length} 件のうち、${added} 件をライブラリに追加しました。`
    : 'この文書にはまだ文献が入っていません。'
})

function exportText (text, filename, type) {
  const box = $('exportBox')
  box.value = text
  box.hidden = false
  box.select()
  try {
    const a = document.createElement('a')
    a.href = URL.createObjectURL(new Blob([text], { type }))
    a.download = filename
    document.body.append(a)
    a.click()
    a.remove()
  } catch { /* 保存できない環境では、下の欄からコピーしてもらう */ }
  toast('書き出しました。保存されない場合は、表示された内容をコピーしてファイルに保存してください。')
}

const stamp = () => new Date().toISOString().slice(0, 10)

$('btnExportJson').addEventListener('click', () => {
  exportText(JSON.stringify(app.library.map(({ _added, ...r }) => r), null, 2), `refcite-library-${stamp()}.json`, 'application/json')
})

$('btnExportRis').addEventListener('click', () => {
  const ris = new Cite(app.library.map(({ _added, ...r }) => r)).format('ris')
  exportText(ris, `refcite-library-${stamp()}.ris`, 'application/x-research-info-systems')
})

$('btnPlain').addEventListener('click', () => {
  if (requireWord()) $('plainConfirm').hidden = false
})
$('btnPlainNo').addEventListener('click', () => { $('plainConfirm').hidden = true })
$('btnPlainYes').addEventListener('click', async () => {
  $('plainConfirm').hidden = true
  try {
    await word.refresh(app.doc)
    const n = await word.convertToPlainText()
    $('plainStatus').textContent = `${n} か所をふつうの文字に変換しました。`
  } catch (e) {
    $('plainStatus').textContent = '変換できませんでした: ' + errorMessage(e)
  }
})

// ---------- 起動 ----------

async function start () {
  app.doc = word.loadDocState()
  renderFormat()
  renderLibrary()
  if (!word.inWord()) {
    $('notWord').hidden = false
    return
  }
  try {
    setNumbers(await word.scan(app.doc))
    renderLibrary()
  } catch { /* 番号の表示がなくても使える */ }
}

if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(() => start())
else start()
