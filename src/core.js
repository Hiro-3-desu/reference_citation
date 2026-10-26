// Word に依存しない処理（文献データの変換・番号付け・文献リストの整形）。
// Node のテストからもそのまま読み込めるようにしている。

import { Cite, plugins } from '@citation-js/core'
import '@citation-js/plugin-csl'
import '@citation-js/plugin-ris'
import '@citation-js/plugin-bibtex'
import { STYLES, LOCALES } from './styles.generated.js'

const csl = plugins.config.get('@csl')
for (const s of STYLES) csl.templates.add(s.id, s.xml)
for (const [lang, xml] of Object.entries(LOCALES)) csl.locales.add(lang, xml)

export const STYLE_LIST = STYLES.map(({ id, title, locale, marker }) => ({ id, title, locale, marker }))

export function newId (prefix = 'r') {
  return prefix + Math.random().toString(36).slice(2, 10)
}

// ---------- 文献データの取り込み ----------

const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 }

function stripTrailingPeriod (s) {
  return (s || '').trim().replace(/\.$/, '')
}

// PubMed の "Smith JA" を CSL の名前に変換
function pubmedName (a) {
  if (a.authtype && a.authtype !== 'Author') return { literal: a.name }
  const m = /^(.+?)\s+([A-Z]{1,4})$/.exec(a.name.trim())
  return m ? { family: m[1], given: m[2] } : { literal: a.name }
}

function pubmedDate (r) {
  // sortpubdate: "2020/10/21 00:00"
  const m = /^(\d{4})\/(\d{2})\/(\d{2})/.exec(r.sortpubdate || '')
  if (m) {
    // pubdate が年だけなら月日を付けない
    const parts = (r.pubdate || '').split(/\s+/)
    if (parts.length === 1) return { 'date-parts': [[+m[1]]] }
    const mon = MONTHS[(parts[1] || '').slice(0, 3).toLowerCase()]
    return { 'date-parts': [mon ? [+m[1], mon] : [+m[1]]] }
  }
  const y = /\d{4}/.exec(r.pubdate || '')
  return y ? { 'date-parts': [[+y[0]]] } : undefined
}

// NCBI esummary（retmode=json）の 1 件を CSL-JSON に変換
export function pubmedSummaryToCsl (r) {
  const ids = Object.fromEntries((r.articleids || []).map(a => [a.idtype, a.value]))
  const item = {
    id: newId(),
    type: 'article-journal',
    title: stripTrailingPeriod(r.title),
    author: (r.authors || []).map(pubmedName),
    'container-title': r.fulljournalname || r.source,
    'container-title-short': r.source,
    volume: r.volume || undefined,
    issue: r.issue || undefined,
    page: r.pages || undefined,
    issued: pubmedDate(r),
    DOI: ids.doi,
    PMID: String(r.uid),
    PMCID: ids.pmc
  }
  // 電子版のみ（ページなし）の論文は article number を使う
  if (!item.page && r.elocationid) {
    const m = /pii:\s*([^\s;]+?)\.?(?:\s|;|$)/.exec(r.elocationid)
    if (m) item.page = m[1]
  }
  return clean(item)
}

// Crossref の CSL-JSON（/transform/application/vnd.citationstyles.csl+json）を整える
export function crossrefToCsl (c) {
  const item = {
    id: newId(),
    type: c.type === 'journal-article' ? 'article-journal' : (c.type || 'article-journal'),
    title: Array.isArray(c.title) ? c.title[0] : c.title,
    author: (c.author || []).map(a => a.literal ? { literal: a.literal } : { family: a.family, given: a.given }),
    'container-title': Array.isArray(c['container-title']) ? c['container-title'][0] : c['container-title'],
    'container-title-short': Array.isArray(c['container-title-short']) ? c['container-title-short'][0] : (c['container-title-short'] || c.journalAbbreviation),
    volume: c.volume,
    issue: c.issue,
    page: c.page || c['article-number'],
    issued: c.issued || c['published-print'] || c['published-online'],
    DOI: c.DOI,
    URL: c.URL,
    publisher: c.publisher
  }
  return clean(item)
}

// RIS / BibTeX / CSL-JSON のテキストを CSL-JSON の配列に変換
export function parseImportText (text) {
  text = text.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n').trim()
  let forceType
  if (/^TY {2}- /m.test(text)) forceType = '@ris/file'
  else if (/^@\w+\s*\{/m.test(text)) forceType = '@biblatex/text'
  const data = new Cite(text, forceType ? { forceType } : {}).data
  // RIS の雑誌略名（JO / J2 / JA）は citation-js が拾わないので自分で読む
  const abbrevs = forceType === '@ris/file'
    ? text.split(/^ER {2}-.*$/m).filter(r => /^TY {2}- /m.test(r)).map(r => /^(?:JO|J2|JA) {2}- (.+)$/m.exec(r)?.[1]?.trim())
    : []
  return data.map((d, i) => {
    const { _graph, ...rest } = d
    const item = clean({ ...rest, id: newId() })
    if (abbrevs[i] && !item['container-title-short']) item['container-title-short'] = abbrevs[i]
    return item
  })
}

function clean (o) {
  for (const k of Object.keys(o)) {
    if (o[k] === undefined || o[k] === null || o[k] === '' || (Array.isArray(o[k]) && o[k].length === 0)) delete o[k]
  }
  return o
}

// 同じ文献かどうか（DOI → PMID → タイトル＋年）
export function refKey (r) {
  if (r.DOI) return 'doi:' + r.DOI.toLowerCase()
  if (r.PMID) return 'pmid:' + r.PMID
  const y = r.issued?.['date-parts']?.[0]?.[0] || ''
  return 'title:' + (r.title || '').toLowerCase().replace(/[^a-z0-9]/g, '') + ':' + y
}

// ---------- 表示用 ----------

export function shortAuthor (r) {
  const a = r.author || []
  if (!a.length) return '著者不明'
  const first = a[0].family || a[0].literal || ''
  return a.length > 1 ? `${first} ほか` : first
}

export function year (r) {
  return r.issued?.['date-parts']?.[0]?.[0] || ''
}

export function journal (r) {
  return r['container-title-short'] || r['container-title'] || ''
}

// ---------- 番号付け ----------

// citations: 文書中の順番に並んだ [{ key, refIds }]
// 戻り値: 文献の番号順の並び（refId の配列）と、各引用の番号の配列
export function computeNumbering (citations) {
  const order = []
  const number = new Map()
  const labels = []
  for (const c of citations) {
    const nums = []
    for (const id of c.refIds) {
      if (!number.has(id)) {
        order.push(id)
        number.set(id, order.length)
      }
      nums.push(number.get(id))
    }
    labels.push([...new Set(nums)].sort((a, b) => a - b))
  }
  return { order, labels }
}

// [1,2,3,5] → "1–3,5"（3 つ以上続くときだけ範囲にする）
export function compressNumbers (nums) {
  const parts = []
  let i = 0
  while (i < nums.length) {
    let j = i
    while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++
    if (j - i >= 2) parts.push(`${nums[i]}–${nums[j]}`)
    else for (let k = i; k <= j; k++) parts.push(String(nums[k]))
    i = j + 1
  }
  return parts.join(',')
}

export const MARKERS = {
  superscript: { label: '上付き数字　¹⁻³', wrap: s => s, superscript: true },
  brackets: { label: '角かっこ　[1–3]', wrap: s => `[${s}]`, superscript: false },
  parens: { label: '丸かっこ　(1–3)', wrap: s => `(${s})`, superscript: false }
}

export function citationLabel (nums, marker) {
  const m = MARKERS[marker] || MARKERS.superscript
  return { text: m.wrap(compressNumbers(nums)), superscript: m.superscript }
}

// ---------- 文献リスト ----------

// 引用順に並んだ CSL-JSON を受け取り、1 件ずつの整形済みテキストを返す。
// 各要素は [{ text, italic, bold, superscript }] の配列（Word に書式付きで書き込むため）。
export function formatBibliography (items, styleId) {
  if (!items.length) return []
  const style = STYLES.find(s => s.id === styleId) || STYLES[0]
  const entries = new Cite(items).format('bibliography', {
    format: 'html',
    template: style.id,
    lang: style.locale,
    asEntryArray: true
  })
  return entries.map(([, html]) => htmlToRuns(html))
}

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', '#38': '&' }

function decode (s) {
  return s.replace(/&(#x?[0-9a-f]+|\w+);/gi, (m, e) => {
    if (e[0] === '#') return String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10))
    return ENTITIES[e] ?? m
  })
}

// citeproc の HTML（<div class="csl-left-margin"> などを含む）を書式付きの文字列片に分解する
export function htmlToRuns (html) {
  const runs = []
  const stack = []
  const fmt = () => ({
    italic: stack.includes('i'),
    bold: stack.includes('b'),
    superscript: stack.includes('sup'),
    subscript: stack.includes('sub')
  })
  const re = /<(\/?)([a-z]+)([^>]*)>|([^<]+)/gi
  let m
  while ((m = re.exec(html))) {
    if (m[4] !== undefined) {
      const text = decode(m[4]).replace(/\s+/g, ' ')
      if (text) runs.push({ text, ...fmt() })
      continue
    }
    const [, close, tag, attrs] = m
    const t = tag.toLowerCase()
    let kind = null
    if (t === 'i' || t === 'em' || (t === 'span' && /font-style:\s*italic/.test(attrs))) kind = 'i'
    else if (t === 'b' || t === 'strong' || (t === 'span' && /font-weight:\s*bold/.test(attrs))) kind = 'b'
    else if (t === 'sup') kind = 'sup'
    else if (t === 'sub') kind = 'sub'
    else if (t === 'span' || t === 'div') kind = 'x'
    if (!kind) continue
    if (close) {
      const idx = stack.lastIndexOf(kind)
      if (idx >= 0) stack.splice(idx, 1)
    } else {
      stack.push(kind)
      // 番号（csl-left-margin）と本文の間に空白を入れる
      if (t === 'div' && /csl-right-inline/.test(attrs) && runs.length && !/\s$/.test(runs[runs.length - 1].text)) {
        runs.push({ text: ' ', ...fmt() })
      }
    }
  }
  // 同じ書式の隣り合う片をまとめ、先頭・末尾の空白を落とす
  const merged = []
  for (const r of runs) {
    const p = merged[merged.length - 1]
    if (p && p.italic === r.italic && p.bold === r.bold && p.superscript === r.superscript && p.subscript === r.subscript) p.text += r.text
    else merged.push({ ...r })
  }
  for (const r of merged) r.text = r.text.replace(/ {2,}/g, ' ')
  while (merged.length && !(merged[0].text = merged[0].text.replace(/^\s+/, ''))) merged.shift()
  while (merged.length && !(merged[merged.length - 1].text = merged[merged.length - 1].text.replace(/\s+$/, ''))) merged.pop()
  return merged
}

export function runsToText (runs) {
  return runs.map(r => r.text).join('')
}
