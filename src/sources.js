// 文献データの取得（PubMed / Crossref）と、このパソコンに保存するライブラリ

import { pubmedSummaryToCsl, crossrefToCsl, refKey } from './core.js'

const EUTILS = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils'
const TOOL = 'tool=refcite'

async function getJson (url) {
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

export async function lookupDoi (doi) {
  // PubMed にあればそちらを優先（雑誌の略名が正確なため）
  const s = await getJson(`${EUTILS}/esearch.fcgi?db=pubmed&retmode=json&${TOOL}&term=${encodeURIComponent(doi + '[doi]')}`)
  const ids = s.esearchresult?.idlist || []
  if (ids.length === 1) return (await pubmedByIds(ids))[0]
  const c = await getJson(`https://api.crossref.org/works/${encodeURIComponent(doi)}/transform/application/vnd.citationstyles.csl+json`)
  return crossrefToCsl(c)
}

// "PMID や DOI を空白・改行区切りで並べたもの" を読み取って文献を取得する
export async function lookupIds (text) {
  const tokens = text.split(/[\s,;]+/).map(t => t.trim()).filter(Boolean)
  const pmids = []
  const dois = []
  const unknown = []
  for (const t of tokens) {
    const doi = /(10\.\d{4,9}\/\S+)/.exec(t)
    if (doi) dois.push(doi[1].replace(/[.)]+$/, ''))
    else if (/^(PMID:?)?\d{1,9}$/i.test(t)) pmids.push(t.replace(/\D/g, ''))
    else unknown.push(t)
  }
  const found = await pubmedByIds(pmids)
  const failed = [...unknown, ...pmids.filter(id => !found.some(f => f.PMID === id))]
  for (const d of dois) {
    try {
      found.push(await lookupDoi(d))
    } catch {
      failed.push(d)
    }
  }
  return { found, failed }
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
