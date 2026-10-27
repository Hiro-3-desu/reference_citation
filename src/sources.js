// 文献データの取得（PubMed / Crossref）と、このパソコンに保存するライブラリ

import { pubmedSummaryToCsl, crossrefToCsl, refKey, titleSimilarity } from './core.js'

const EUTILS = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils'
const TOOL = 'tool=refcite'

// NCBI は API キーなしだと 1 秒に 3 回まで。間隔を空けて順番に送る
let ncbiQueue = Promise.resolve()
function ncbiSlot () {
  const slot = ncbiQueue.then(() => new Promise(resolve => setTimeout(resolve, 380)))
  ncbiQueue = slot
  return ncbiQueue
}

async function getJson (url) {
  if (url.startsWith(EUTILS)) await ncbiSlot()
  const res = await fetch(url)
  if (!res.ok) throw new Error(`通信エラー（${res.status}）`)
  return res.json()
}

export async function pubmedSearch (term, retmax = 20) {
  const q = encodeURIComponent(term)
  const s = await getJson(`${EUTILS}/esearch.fcgi?db=pubmed&retmode=json&sort=relevance&retmax=${retmax}&${TOOL}&term=${q}`)
  const ids = s.esearchresult?.idlist || []
  return { total: +(s.esearchresult?.count || 0), items: await pubmedByIds(ids) }
}

export async function pubmedByIds (ids) {
  if (!ids.length) return []
  const s = await getJson(`${EUTILS}/esummary.fcgi?db=pubmed&retmode=json&${TOOL}&id=${ids.join(',')}`)
  return (s.result?.uids || []).filter(uid => !s.result[uid].error).map(uid => pubmedSummaryToCsl(s.result[uid]))
}

// PubMed の検索語を作る（タイトルの主な語を [ti] で AND）
function titleQuery (title) {
  const ws = title.toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .split(/[^a-z0-9]+/).filter(w => w.length >= 3)
    .filter(w => !['the', 'and', 'for', 'with', 'from', 'versus', 'among', 'after', 'into', 'their', 'that', 'this'].includes(w))
  return [...new Set(ws)].slice(0, 12).map(w => `${w}[ti]`).join(' AND ')
}

async function esearchIds (term, retmax = 20) {
  const s = await getJson(`${EUTILS}/esearch.fcgi?db=pubmed&retmode=json&${TOOL}&retmax=${retmax}&term=${encodeURIComponent(term)}`)
  return s.esearchresult?.idlist || []
}

async function crossrefDoi (doi) {
  const res = await fetch(`https://api.crossref.org/works/${encodeURIComponent(doi)}/transform/application/vnd.citationstyles.csl+json`)
  if (!res.ok) return null
  return crossrefToCsl(await res.json())
}

// parseAiText の結果を PubMed / Crossref で確かめる。
// 各項目に status を付けて返す:
//   ok        … ID で見つかり、AI の書いた内容とも合っている
//   mismatch  … ID で見つかったが、AI の書いたタイトルと食い違う（ID の間違いの可能性）
//   candidates… ID がなく、タイトルで検索した候補がある
//   notfound  … 見つからない（実在しない可能性）
export async function resolveAiItems (items, onProgress = () => {}) {
  const pmids = new Set(items.flatMap(it => it.pmids))
  const dois = [...new Set(items.flatMap(it => it.dois))]
  const pmcids = [...new Set(items.flatMap(it => it.pmcids))]

  // DOI と PMCID は PubMed でまとめて PMID に変換する
  onProgress('PubMed で確認しています…')
  const extra = []
  for (let i = 0; i < dois.length; i += 40) extra.push(...await esearchIds(dois.slice(i, i + 40).map(d => `"${d}"[doi]`).join(' OR '), 100))
  for (let i = 0; i < pmcids.length; i += 40) extra.push(...await esearchIds(pmcids.slice(i, i + 40).map(d => `${d}[pmcid]`).join(' OR '), 100))
  extra.forEach(id => pmids.add(id))

  // タイトルしかないものは、タイトルで検索する
  const titleOnly = items.filter(it => !it.pmids.length && !it.dois.length && !it.pmcids.length && it.title)
  const candidateIds = new Map()
  let n = 0
  for (const it of titleOnly) {
    onProgress(`タイトルで検索しています（${++n} / ${titleOnly.length}）…`)
    const q = titleQuery(it.title)
    let ids = q ? await esearchIds(q, 5) : []
    if (!ids.length && q && it.year) ids = await esearchIds(q.split(' AND ').slice(0, 6).join(' AND ') + ` AND ${it.year}[dp]`, 5)
    candidateIds.set(it, ids)
    ids.forEach(id => pmids.add(id))
  }

  const all = [...pmids]
  const records = []
  for (let i = 0; i < all.length; i += 150) records.push(...await pubmedByIds(all.slice(i, i + 150)))
  const byPmid = new Map(records.map(r => [r.PMID, r]))
  const byDoi = new Map(records.filter(r => r.DOI).map(r => [r.DOI.toLowerCase(), r]))
  const byPmc = new Map(records.filter(r => r.PMCID).map(r => [r.PMCID.toUpperCase(), r]))

  const out = []
  for (const it of items) {
    if (candidateIds.has(it)) {
      const cands = candidateIds.get(it).map(id => byPmid.get(id)).filter(Boolean)
        .map(r => ({ ref: r, score: titleSimilarity(r.title, it.label) }))
        .sort((a, b) => b.score - a.score)
        .slice(0, 3)
      out.push({ ...it, status: cands.length ? 'candidates' : 'notfound', candidates: cands })
      continue
    }
    let ref = it.pmids.map(p => byPmid.get(p)).find(Boolean) ||
      it.dois.map(d => byDoi.get(d)).find(Boolean) ||
      it.pmcids.map(p => byPmc.get(p)).find(Boolean)
    if (!ref && it.dois.length) {
      onProgress('Crossref で確認しています…')
      for (const d of it.dois) {
        try { ref = await crossrefDoi(d) } catch { /* 次の DOI を試す */ }
        if (ref) break
      }
    }
    if (!ref) {
      out.push({ ...it, status: 'notfound' })
      continue
    }
    const score = it.checkable ? titleSimilarity(ref.title, it.label) : 1
    // タイトルが書かれていなくても、筆頭著者と発行年が合っていれば同じ論文とみなす
    const first = (ref.author?.[0]?.family || '').toLowerCase()
    const y = String(ref.issued?.['date-parts']?.[0]?.[0] || '')
    const authorYear = first.length >= 2 && y && it.label.toLowerCase().includes(first) && it.label.includes(y)
    // PMID と DOI が別の論文を指していないかも見る
    const doiConflict = it.pmids.length && it.dois.length && ref.DOI && !it.dois.includes(ref.DOI.toLowerCase())
    out.push({ ...it, ref, score, status: (score >= 0.6 || authorYear) && !doiConflict ? 'ok' : 'mismatch', doiConflict: !!doiConflict })
  }

  // ID が食い違っていたものは、AI が書いたタイトルでも探して、正しい論文の候補を出す
  const mismatched = out.filter(it => it.status === 'mismatch' && it.title)
  const extraIds = new Map()
  n = 0
  for (const it of mismatched) {
    onProgress(`正しい論文を探しています（${++n} / ${mismatched.length}）…`)
    const q = titleQuery(it.title)
    extraIds.set(it, q ? (await esearchIds(q, 3)).filter(id => id !== it.ref.PMID) : [])
  }
  const need = [...new Set([...extraIds.values()].flat())].filter(id => !byPmid.has(id))
  if (need.length) (await pubmedByIds(need)).forEach(r => byPmid.set(r.PMID, r))
  for (const [it, ids] of extraIds) {
    it.candidates = ids.map(id => byPmid.get(id)).filter(Boolean)
      .map(r => ({ ref: r, score: titleSimilarity(r.title, it.label) }))
      .filter(c => c.score >= 0.5)
      .sort((a, b) => b.score - a.score)
  }
  return out
}

// ---------- ライブラリ（このパソコンの Word アドイン内に保存） ----------

const LIB_KEY = 'refcite.library.v1'

export function loadLibrary () {
  try {
    return JSON.parse(localStorage.getItem(LIB_KEY))?.items || []
  } catch {
    return []
  }
}

export function saveLibrary (items) {
  try {
    localStorage.setItem(LIB_KEY, JSON.stringify({ v: 1, items }))
    return true
  } catch {
    return false
  }
}

// 重複を除いて追加する。戻り値は実際に追加した件数
export function addToLibrary (library, items) {
  const keys = new Set(library.map(refKey))
  let added = 0
  for (const it of items) {
    const k = refKey(it)
    if (keys.has(k)) continue
    keys.add(k)
    library.unshift({ ...it, _added: new Date().toISOString() })
    added++
  }
  return added
}

// ---------- 引用の形式（CSL スタイル）の追加 ----------

const CSL_RAW = 'https://raw.githubusercontent.com/citation-style-language'
let styleIndex = null

// ビルド時に作った一覧（番号式の形式だけ、約 5,500 件）を読み込む
export async function loadStyleIndex () {
  if (!styleIndex) {
    const res = await fetch('styles-index.json')
    if (!res.ok) throw new Error('形式の一覧を読み込めませんでした。')
    styleIndex = await res.json()
  }
  return styleIndex
}

export function searchStyles (index, query, limit = 30) {
  const q = query.toLowerCase().normalize('NFKC').split(/\s+/).filter(Boolean)
  if (!q.length) return []
  const hits = []
  for (const s of index.styles) {
    const hay = `${s.t} ${s.s || ''} ${s.id}`.toLowerCase()
    if (!q.every(w => hay.includes(w))) continue
    // 雑誌名がそのまま一致するもの、前方一致するものを上に
    const t = s.t.toLowerCase()
    const full = q.join(' ')
    const rank = t === full || (s.s || '').toLowerCase() === full ? 0 : t.startsWith(full) ? 1 : 2
    hits.push({ ...s, rank })
  }
  return hits.sort((a, b) => a.rank - b.rank || a.t.length - b.t.length).slice(0, limit)
}

async function getText (url) {
  const res = await fetch(url)
  if (!res.ok) throw new Error(`形式のファイルを取得できませんでした（${res.status}）`)
  return res.text()
}

// 形式の CSL を取得する。雑誌ごとの形式（dependent）は、元になる形式の中身を使う。
export async function fetchStyleXml (entry, commit = 'master') {
  if (!entry.p) return getText(`${CSL_RAW}/styles/${commit}/${entry.id}.csl`)
  return getText(`${CSL_RAW}/styles/${commit}/${entry.p}.csl`)
}

export async function fetchLocaleXml (lang) {
  return getText(`${CSL_RAW}/locales/master/locales-${lang}.xml`)
}

const STYLES_KEY = 'refcite.styles.v1'
const LOCALES_KEY = 'refcite.locales.v1'

function loadJson (key) {
  try { return JSON.parse(localStorage.getItem(key)) || {} } catch { return {} }
}

function saveJson (key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value))
    return true
  } catch {
    return false
  }
}

export const savedStyles = () => loadJson(STYLES_KEY)
export const saveStyles = styles => saveJson(STYLES_KEY, styles)
export const savedLocales = () => loadJson(LOCALES_KEY)
export const saveLocales = locales => saveJson(LOCALES_KEY, locales)
