/* global Office */
import { Cite } from '@citation-js/core'
import { styleList, getStyle, registerStyle, unregisterStyle, analyzeCsl, hasLocale, registerLocale, MARKERS, refKey, shortAuthor, year, journal, parseImportText, parseAiText, linksFor } from './core.js'
import { pubmedSearch, resolveAiItems, findOpenAccess, loadLibrary, saveLibrary, addToLibrary,
  loadStyleIndex, searchStyles, fetchStyleXml, fetchLocaleXml, savedStyles, saveStyles, savedLocales, saveLocales } from './sources.js'
import * as word from './word.js'

const $ = id => document.getElementById(id)

const app = {
  library: loadLibrary(),
  doc: word.loadDocState(),
  selected: new Set(), // ライブラリで選んだ文献の id
  numbers: new Map(), // refKey → この文書での番号
  pmResults: [],
  pmSelected: new Set(),
  aiResults: [], // resolveAiItems の結果
  aiPick: new Map(), // aiResults の番号 → 追加する文献（追加しないなら null）
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

// ---------- 論文を開く ----------

// いつも使っているブラウザで開く（Word の中では Office の API、それ以外はふつうのリンク）
function openUrl (url) {
  try {
    if (typeof Office !== 'undefined' && Office.context?.ui?.openBrowserWindow) {
      Office.context.ui.openBrowserWindow(url)
      return
    }
  } catch { /* 下の方法で開く */ }
  window.open(url, '_blank', 'noopener')
}

const oaCache = new Map() // refKey → 無料版の { url, pdf } または null

// ライブラリにある同じ文献から、自分の PDF のリンクを探す
function ownPdf (r) {
  if (r._pdf) return r._pdf
  const k = refKey(r)
  return app.library.find(x => refKey(x) === k)?._pdf || ''
}

// 「PubMed・出版社・全文（PMC）・PDF を探す・自分の PDF」の並び。pdfEdit ならリンクの登録もできる
function linkRow (r, { pdfEdit = false } = {}) {
  const row = document.createElement('div')
  row.className = 'links'
  // クリックしても、文献の選択が切り替わらないようにする
  row.addEventListener('click', e => e.stopPropagation())
  const mine = ownPdf(r)
  if (mine) {
    const a = document.createElement('a')
    a.href = mine
    a.className = 'mine'
    a.textContent = '自分の PDF'
    a.addEventListener('click', e => { e.preventDefault(); openUrl(mine) })
    row.append(a)
  }
  for (const l of linksFor(r)) {
    const a = document.createElement('a')
    a.href = l.url
    a.target = '_blank'
    a.rel = 'noopener'
    a.textContent = l.label
    a.addEventListener('click', e => { e.preventDefault(); openUrl(l.url) })
    row.append(a)
  }
  if (r.DOI || r.PMID) {
    const btn = document.createElement('button')
    btn.className = 'linkbtn'
    const key = refKey(r)
    const show = oa => {
      btn.textContent = oa === undefined ? 'PDF を探す' : oa ? (oa.pdf ? '無料 PDF' : '無料で読む') : '無料版なし'
      btn.disabled = oa === null
      btn.title = oa === null ? '無料で読める版は見つかりませんでした。「出版社」から開き、所属機関のアカウントで読んでください。' : ''
    }
    show(oaCache.get(key))
    btn.addEventListener('click', async e => {
      e.preventDefault()
      let oa = oaCache.get(key)
      if (oa === undefined) {
        btn.disabled = true
        btn.textContent = '探しています…'
        try {
          oa = await findOpenAccess(r)
          oaCache.set(key, oa)
        } catch (err) {
          show(undefined)
          toast('PDF を探せませんでした: ' + errorMessage(err), true)
          return
        }
        show(oa)
        if (!oa) {
          toast('無料で読める版は見つかりませんでした。「出版社」から開き、所属機関のアカウントで読んでください。')
          return
        }
      }
      openUrl(oa.url)
    })
    row.append(btn)
  }
  if (pdfEdit) {
    const btn = document.createElement('button')
    btn.className = 'linkbtn'
    btn.textContent = mine ? 'PDF を変更' : 'PDF を登録'
    btn.addEventListener('click', e => {
      e.preventDefault()
      row.after(pdfEditor(r))
      btn.disabled = true
    })
    row.append(btn)
  }
  return row
}

// 自分の PDF のリンクを登録・変更・削除する欄
function pdfEditor (r) {
  const box = document.createElement('div')
  box.className = 'pdfedit'
  box.addEventListener('click', e => e.stopPropagation())
  const help = document.createElement('p')
  help.textContent = 'OneDrive などで PDF の「リンクをコピー」をして貼り付けてください。共有範囲は「自分だけ」にしてください。'
  const input = document.createElement('input')
  input.type = 'url'
  input.placeholder = 'https://…'
  input.value = r._pdf || ''
  const msg = document.createElement('p')
  msg.className = 'err'
  const row = document.createElement('div')
  row.className = 'row end'
  const mk = (text, cls, fn) => {
    const b = document.createElement('button')
    b.textContent = text
    if (cls) b.className = cls
    b.addEventListener('click', fn)
    return b
  }
  const save = url => {
    const item = app.library.find(x => x.id === r.id)
    if (url) item._pdf = url
    else delete item._pdf
    persistLibrary()
    renderLibrary()
  }
  row.append(
    ...(r._pdf ? [mk('削除', 'ghost', () => { save(''); toast('PDF のリンクを削除しました') })] : []),
    mk('やめる', 'ghost', () => renderLibrary()),
    mk('保存', 'primary', () => {
      const url = input.value.trim()
      if (/^[a-z]:\\|^\\\\|^file:/i.test(url)) {
        msg.textContent = 'パソコン内のファイルは開けません。OneDrive などに保存して、その共有リンクを貼り付けてください。'
        return
      }
      if (!/^https:\/\/\S+$/i.test(url)) {
        msg.textContent = '「https://」で始まるリンクを貼り付けてください。'
        return
      }
      save(url)
      toast('PDF のリンクを登録しました')
    })
  )
  input.addEventListener('keydown', e => { if (e.key === 'Enter') row.lastChild.click() })
  box.append(help, input, msg, row)
  setTimeout(() => input.focus(), 0)
  return box
}

// 文献 1 件分の <li> を作る（テキストはすべて textContent で入れる）
function refItem (r, { checkbox = false, checked = false, badge = '', button = null, pdfEdit = false } = {}) {
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
  li.append(meta, title, jr, linkRow(r, { pdfEdit }))
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
    const li = refItem(r, { checkbox: true, checked: app.selected.has(r.id), badge: n ? `[${n}]` : '', pdfEdit: true })
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
      // 自分の PDF のリンクは文書に入れない（文書を人に渡しても見えないように）
      const { _added, _pdf, ...csl } = r
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

// ---------- AI の回答から追加 ----------

const AI_STATUS = {
  ok: ['ok', '確認できました'],
  mismatch: ['warn', '要確認'],
  candidates: ['info', 'タイトルで検索'],
  notfound: ['bad', '見つかりません']
}

function aiOption (r, { type, name, checked, onPick }) {
  const label = document.createElement('label')
  label.className = 'opt'
  const input = document.createElement('input')
  input.type = type
  input.name = name
  input.checked = checked
  input.addEventListener('change', () => onPick(input.checked))
  const meta = document.createElement('div')
  meta.className = 'meta'
  const b = document.createElement('b')
  b.textContent = shortAuthor(r)
  const y = document.createElement('span')
  y.textContent = year(r)
  meta.append(b, y)
  const title = document.createElement('div')
  title.textContent = r.title || '（タイトルなし）'
  const jr = document.createElement('div')
  jr.className = 'jr'
  jr.textContent = [journal(r), r.PMID ? `PMID ${r.PMID}` : (r.DOI ? `DOI ${r.DOI}` : '')].filter(Boolean).join(' · ')
  label.append(input, meta, title, jr, linkRow(r))
  return label
}

function renderAi () {
  const list = $('aiList')
  list.replaceChildren()
  const inLib = new Set(app.library.map(refKey))
  app.aiResults.forEach((it, i) => {
    const li = document.createElement('li')
    li.className = 'rv'
    const owned = it.ref && inLib.has(refKey(it.ref))
    const [cls, text] = owned ? ['ok', '追加済み'] : AI_STATUS[it.status]
    const chip = document.createElement('span')
    chip.className = 'chip ' + cls
    chip.textContent = text
    const src = document.createElement('div')
    src.className = 'src'
    src.textContent = 'AI の回答: ' + it.label
    li.append(chip, src)

    const note = msg => {
      const p = document.createElement('div')
      p.className = 'note'
      p.textContent = msg
      li.append(p)
    }
    const pick = ref => checked => {
      app.aiPick.set(i, checked ? ref : null)
      renderAiButton()
    }
    const radio = (r, label) => {
      const opt = aiOption(r, { type: 'radio', name: 'ai' + i, checked: app.aiPick.get(i) === r, onPick: pick(r) })
      if (inLib.has(refKey(r))) {
        opt.querySelector('input').disabled = true
        opt.querySelector('.meta').append(Object.assign(document.createElement('span'), { className: 'badge', textContent: '追加済み' }))
      } else if (label) {
        opt.querySelector('.meta').append(Object.assign(document.createElement('span'), { className: 'badge', textContent: label }))
      }
      return opt
    }
    const noneOption = () => {
      const none = document.createElement('label')
      none.className = 'opt'
      const r = Object.assign(document.createElement('input'), { type: 'radio', name: 'ai' + i, checked: !app.aiPick.get(i) })
      r.addEventListener('change', () => { app.aiPick.set(i, null); renderAiButton() })
      none.append(r, Object.assign(document.createElement('div'), { textContent: 'どれでもない（追加しない）' }))
      return none
    }
    if (it.status === 'ok') {
      if (!owned) li.append(aiOption(it.ref, { type: 'checkbox', name: 'ai' + i, checked: !!app.aiPick.get(i), onPick: pick(it.ref) }))
    } else if (it.status === 'mismatch') {
      note(it.doiConflict
        ? 'PMID と DOI が別の論文を指しています。正しいものを選んでください。'
        : 'AI が書いた内容と、この ID の論文が一致しません。ID が間違っている可能性があります。正しいものを選んでください。')
      li.append(radio(it.ref, 'ID の論文'))
      for (const c of it.candidates || []) li.append(radio(c.ref, 'タイトルで検索'))
      li.append(noneOption())
    } else if (it.status === 'candidates') {
      note('ID が書かれていないので、タイトルで探した候補です。正しいものを選んでください。')
      for (const c of it.candidates) li.append(radio(c.ref))
      li.append(noneOption())
    } else {
      note('PubMed にも Crossref にも見つかりませんでした。AI が実在しない論文を挙げた可能性があります。')
    }
    list.append(li)
  })
  $('aiReview').hidden = app.aiResults.length === 0
  renderAiButton()
}

function renderAiButton () {
  const n = [...app.aiPick.values()].filter(Boolean).length
  $('aiAdd').disabled = n === 0
  $('aiAdd').textContent = n ? `選択した ${n} 件を追加` : '選択した文献を追加'
}

$('aiRun').addEventListener('click', async () => {
  const text = $('aiInput').value
  const items = parseAiText(text)
  app.aiResults = []
  app.aiPick.clear()
  renderAi()
  if (!items.length) {
    $('aiStatus').textContent = '文献が見つかりませんでした。PMID・DOI・論文タイトルを含む部分を貼り付けてください。'
    return
  }
  $('aiRun').disabled = true
  try {
    app.aiResults = await resolveAiItems(items, msg => { $('aiStatus').textContent = msg })
    const inLib = new Set(app.library.map(refKey))
    app.aiResults.forEach((it, i) => {
      // 確認できたものと、タイトルがよく一致する候補は、最初から選んでおく
      if (it.status === 'ok' && !inLib.has(refKey(it.ref))) app.aiPick.set(i, it.ref)
      // タイトルがよく一致する候補は選んでおく（ID が食い違っていたものも、正しい論文が見つかれば）
      const top = it.candidates?.[0]
      if ((it.status === 'candidates' || it.status === 'mismatch') && top && top.score >= 0.8 && !inLib.has(refKey(top.ref))) app.aiPick.set(i, top.ref)
    })
    const count = s => app.aiResults.filter(it => it.status === s).length
    $('aiStatus').textContent = [
      `${app.aiResults.length} 件を確認しました。`,
      count('ok') ? `確認できた ${count('ok')} 件` : '',
      count('mismatch') ? `要確認 ${count('mismatch')} 件` : '',
      count('candidates') ? `タイトルで検索 ${count('candidates')} 件` : '',
      count('notfound') ? `見つからない ${count('notfound')} 件` : ''
    ].filter(Boolean).join(' / ')
  } catch (err) {
    $('aiStatus').textContent = '確認できませんでした: ' + errorMessage(err)
  }
  $('aiRun').disabled = false
  renderAi()
})

$('aiAdd').addEventListener('click', () => {
  const refs = [...app.aiPick.values()].filter(Boolean)
  const added = addItems(refs)
  app.aiPick.clear()
  renderAi()
  toast(`${added} 件をライブラリに追加しました`)
})

$('aiClear').addEventListener('click', () => {
  $('aiInput').value = ''
  $('aiStatus').textContent = ''
  app.aiResults = []
  app.aiPick.clear()
  renderAi()
})

$('aiPromptCopy').addEventListener('click', async () => {
  const box = $('aiPrompt')
  try {
    await navigator.clipboard.writeText(box.value)
    toast('コピーしました')
  } catch {
    box.select()
    toast('選択しました。Ctrl+C でコピーしてください。')
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
  const list = styleList()
  const group = (label, items) => {
    const g = document.createElement('optgroup')
    g.label = label
    g.append(...items.map(s => new Option(s.title, s.id)))
    return g
  }
  const custom = list.filter(s => s.custom)
  ss.replaceChildren(group('よく使う形式', list.filter(s => !s.custom)), ...(custom.length ? [group('追加した形式', custom)] : []))
  ss.value = getStyle(app.doc.style) ? app.doc.style : 'bmj'
  const ms = $('markerSelect')
  ms.replaceChildren(...Object.entries(MARKERS).map(([k, m]) => new Option(m.label, k)))
  ms.value = app.doc.marker
  renderCustomStyles()
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

// 文書で使う形式を切り替える。追加した形式なら、ほかのパソコンでも開けるように文書にも保存する。
function useStyle (id) {
  const style = getStyle(id)
  app.doc.style = id
  app.doc.marker = style?.marker || app.doc.marker
  if (style && !styleList().find(s => s.id === id)?.custom) delete app.doc.customStyle
  else if (style) app.doc.customStyle = { id: style.id, title: style.title, xml: style.xml, locale: style.locale, marker: style.marker }
  renderFormat()
  applyFormat(res => `${style?.title || id} に変更しました（引用 ${res.citations} か所・文献 ${res.references} 件）。`)
}

$('styleSelect').addEventListener('change', () => useStyle($('styleSelect').value))

// ---------- 形式の追加 ----------

async function ensureLocale (lang) {
  if (!lang || hasLocale(lang)) return
  try {
    const xml = await fetchLocaleXml(lang)
    registerLocale(lang, xml)
    const all = savedLocales()
    all[lang] = xml
    saveLocales(all)
  } catch { /* 言語ファイルがなければ英語（米国）で整形される */ }
}

// 形式を登録し、このパソコンに保存する
async function addStyle (style) {
  await ensureLocale(style.locale)
  registerStyle(style)
  const all = savedStyles()
  all[style.id] = style
  if (!saveStyles(all)) toast('形式をこのパソコンに保存できませんでした。この文書でだけ使えます。', true)
}

function renderCustomStyles () {
  const list = $('customStyles')
  list.replaceChildren()
  const saved = Object.values(savedStyles())
  for (const st of saved) {
    const li = document.createElement('li')
    li.className = 'srow'
    const name = document.createElement('div')
    name.className = 'name'
    name.textContent = st.title
    const del = document.createElement('button')
    del.className = 'ghost'
    del.textContent = '削除'
    del.addEventListener('click', () => {
      const all = savedStyles()
      delete all[st.id]
      saveStyles(all)
      // 開いている文書で使っている形式は、文書に保存した分で使い続けられる
      if (app.doc.style !== st.id) unregisterStyle(st.id)
      renderFormat()
      toast(`「${st.title}」を一覧から削除しました`)
    })
    li.append(name, del)
    list.append(li)
  }
  $('customBox').hidden = saved.length === 0
}

function renderStyleResults (hits) {
  const list = $('styleResults')
  list.replaceChildren()
  for (const h of hits) {
    const li = document.createElement('li')
    li.className = 'srow'
    const name = document.createElement('div')
    name.className = 'name'
    name.textContent = h.t
    if (h.p) {
      const sub = document.createElement('small')
      sub.textContent = `元の形式: ${h.p}`
      name.append(sub)
    }
    const btn = document.createElement('button')
    const have = !!getStyle(h.id)
    btn.textContent = have ? '使う' : '追加'
    btn.className = have ? '' : 'primary'
    btn.addEventListener('click', async () => {
      btn.disabled = true
      try {
        if (!have) {
          const index = await loadStyleIndex()
          const xml = await fetchStyleXml(h, index.commit)
          const a = analyzeCsl(xml)
          if (!a.ok) throw new Error(a.reason)
          await addStyle({ id: h.id, title: h.t, xml, locale: a.locale, marker: a.marker })
        }
        useStyle(h.id)
        renderStyleResults(hits)
        toast(`「${h.t}」を追加して、この文書の形式にしました`)
      } catch (e) {
        btn.disabled = false
        toast('追加できませんでした: ' + errorMessage(e), true)
      }
    })
    li.append(name, btn)
    list.append(li)
  }
}

$('styleForm').addEventListener('submit', async e => {
  e.preventDefault()
  const q = $('styleTerm').value.trim()
  if (!q) return
  $('styleStatus').textContent = '探しています…'
  try {
    const hits = searchStyles(await loadStyleIndex(), q)
    $('styleStatus').textContent = hits.length
      ? `${hits.length} 件見つかりました${hits.length >= 30 ? '（上位 30 件）' : ''}。`
      : '見つかりませんでした。別の書き方（略さない雑誌名など）で試すか、下の「CSL ファイルから読み込む」を使ってください。'
    renderStyleResults(hits)
  } catch (err) {
    $('styleStatus').textContent = errorMessage(err)
  }
})

$('styleFile').addEventListener('change', async e => {
  const file = e.target.files[0]
  e.target.value = ''
  if (!file) return
  try {
    let xml = await file.text()
    let a = analyzeCsl(xml)
    if (a.ok && a.dependent) {
      // 雑誌ごとの形式ファイルは、元になる形式を取りに行く
      xml = await fetchStyleXml({ id: a.id, p: a.parent })
      const title = a.title
      a = analyzeCsl(xml)
      a.title = title
    }
    if (!a.ok) throw new Error(a.reason)
    const id = 'file-' + a.id
    await addStyle({ id, title: a.title, xml, locale: a.locale, marker: a.marker })
    useStyle(id)
    toast(`「${a.title}」を追加して、この文書の形式にしました`)
  } catch (err) {
    toast(`${file.name} を読み込めませんでした: ${errorMessage(err)}`, true)
  }
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
  const ris = new Cite(app.library.map(({ _added, _pdf, ...r }) => r)).format('ris')
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
    $('plainStatus').textContent = `${n} か所のリンクを外しました。`
  } catch (e) {
    $('plainStatus').textContent = 'リンクを外せませんでした: ' + errorMessage(e)
  }
})

// ---------- 起動 ----------

/* global __VERSION__ */
$('version').textContent = typeof __VERSION__ !== 'undefined' ? __VERSION__ : '開発版'

async function start () {
  for (const [lang, xml] of Object.entries(savedLocales())) registerLocale(lang, xml)
  for (const st of Object.values(savedStyles())) registerStyle(st)
  app.doc = word.loadDocState()
  // ほかのパソコンで追加した形式を使っている文書なら、文書に保存された形式を使う
  const cs = app.doc.customStyle
  if (cs && cs.id === app.doc.style && !getStyle(cs.id)) {
    registerStyle(cs)
    ensureLocale(cs.locale)
  }
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
