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

// あとから追加した形式（雑誌名で検索して追加したもの、CSL ファイルから読み込んだもの）
const extraStyles = new Map()

export function styleList () {
  return [
    ...STYLES.map(({ id, title, locale, marker }) => ({ id, title, locale, marker, custom: false })),
    ...[...extraStyles.values()].filter(s => !STYLES.some(b => b.id === s.id))
      .map(({ id, title, locale, marker }) => ({ id, title, locale, marker, custom: true }))
  ]
}

export function getStyle (id) {
  return STYLES.find(s => s.id === id) || extraStyles.get(id) || null
}

export function registerStyle (style) {
  csl.templates.add(style.id, style.xml)
  extraStyles.set(style.id, style)
}

export function unregisterStyle (id) {
  extraStyles.delete(id)
}

export function hasLocale (lang) {
  return csl.locales.has(lang)
}

export function registerLocale (lang, xml) {
  csl.locales.add(lang, xml)
}

const attr = (xml, re) => re.exec(xml)?.[1]

// CSL ファイルの中身を調べる。このアドインは番号で引用するので、番号式で、
// 文献リストを引用順に並べる形式だけを受け付ける。
export function analyzeCsl (xml) {
  if (!/<style\b[^>]*xmlns="http:\/\/purl\.org\/net\/xbiblio\/csl"/.test(xml)) {
    return { ok: false, reason: 'CSL の形式ファイルではありません。' }
  }
  const info = /<info>([\s\S]*?)<\/info>/.exec(xml)?.[1] || ''
  const title = (attr(info, /<title>([\s\S]*?)<\/title>/) || '名前のない形式').replace(/&amp;/g, '&').trim()
  const idUrl = attr(info, /<id>([\s\S]*?)<\/id>/) || ''
  const id = (idUrl.split('/').filter(Boolean).pop() || title).toLowerCase().replace(/[^a-z0-9-]+/g, '-')
  const parent = attr(info, /href="https?:\/\/www\.zotero\.org\/styles\/([^"]+)"\s+rel="independent-parent"/) ||
    attr(info, /rel="independent-parent"\s+href="https?:\/\/www\.zotero\.org\/styles\/([^"]+)"/)
  const locale = attr(xml, /<style\b[^>]*default-locale="([^"]+)"/) || 'en-US'
  if (parent && !/<citation\b/.test(xml)) return { ok: true, dependent: true, parent, id, title, locale }

  const format = attr(info, /citation-format="([^"]+)"/)
  if (format && format !== 'numeric') {
    return { ok: false, id, title, reason: `「${title}」は番号ではなく ${format === 'author-date' ? '著者名と年' : '脚注など'} で引用する形式なので、このアドインでは使えません。` }
  }
  const bib = /<bibliography\b[\s\S]*?<\/bibliography>/.exec(xml)?.[0]
  if (!bib) return { ok: false, id, title, reason: `「${title}」には文献リストの書き方が含まれていません。` }
  const sortKeys = [...(/<sort>([\s\S]*?)<\/sort>/.exec(bib)?.[1] || '').matchAll(/<key\b([^>]*)>/g)].map(m => m[1])
  if (sortKeys.some(k => !/variable="citation-number"/.test(k))) {
    return { ok: false, id, title, reason: `「${title}」は文献リストを著者名の順に並べる形式なので、引用順の番号と合わず、このアドインでは使えません。` }
  }
  const cit = /<citation\b[\s\S]*?<\/citation>/.exec(xml)?.[0] || ''
  const marker = /vertical-align="sup"/.test(cit) ? 'superscript' : /prefix="\("/.test(cit) ? 'parens' : 'brackets'
  return { ok: true, dependent: false, id, title, locale, marker }
}

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
  // RefCite で書き出した JSON は、そのまま読む（自分の PDF のリンクも残す）
  if (text.startsWith('[')) {
    try {
      const arr = JSON.parse(text)
      if (Array.isArray(arr) && arr.every(o => o && typeof o === 'object' && (o.title || o.type))) {
        return arr.map(({ _added, ...o }) => clean({ ...o, id: newId() }))
      }
    } catch { /* CSL-JSON として下で読む */ }
  }
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
  const style = getStyle(styleId) || STYLES[0]
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

// ---------- AI の回答からの読み取り ----------

const STOPWORDS = new Set(('a an and are as at be by for from has have in into is it its of on or that the their this to was were with ' +
  'versus vs among after before during between within without using use study trial patients').split(' '))

function words (s) {
  return (s || '').toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '')
    .split(/[^a-z0-9]+/).filter(w => w.length >= 3 && !STOPWORDS.has(w))
}

// 論文タイトルの語のうち、AI の回答の文章に含まれている割合（0〜1）
export function titleSimilarity (title, text) {
  const t = [...new Set(words(title))]
  if (!t.length) return 0
  const have = new Set(words(text))
  return t.filter(w => have.has(w)).length / t.length
}

// Markdown の装飾とリンクを外した、読める文字列にする
function plain (s) {
  return s
    .replace(/!?\[([^\]]*)\]\(([^)]*)\)/g, (m, label, url) => `${label} ${url}`)
    .replace(/<\/?[a-z][^>]*>/gi, ' ')
    .replace(/[*_`]+/g, '')
    .replace(/^\s*(\d{1,3}[.)]|\[\d{1,3}\]|[-*•+])\s+/, '')
    .replace(/\s*\|\s*/g, ' | ')
    .replace(/\s+/g, ' ')
    .trim()
}

const DOI_RE = /\b(10\.\d{4,9}\/[^\s"'<>|\]\[]+)/g
const PMID_RES = [
  /\bPMIDs?\s*[:#：]?\s*(\d{1,9})\b/gi,
  /\bPMIDs?\s*[:#：]?\s*(?:\d{1,9}\s*[,、;]\s*)+(\d{1,9})\b/gi, // "PMIDs: 1, 2, 3" の 2 つ目以降
  /pubmed\.ncbi\.nlm\.nih\.gov\/(\d{1,9})/gi,
  /ncbi\.nlm\.nih\.gov\/pubmed\/(\d{1,9})/gi
]
const PMC_RE = /\b(PMC\d{3,9})\b/g

function cleanDoi (d) {
  try { d = decodeURIComponent(d) } catch { /* そのまま使う */ }
  d = d.replace(/[.,;:。、）]+$/, '')
  // 末尾の ")" は、DOI の中に "(" がなければ文章側のかっこ
  while (d.endsWith(')') && (d.match(/\(/g) || []).length < (d.match(/\)/g) || []).length) d = d.slice(0, -1)
  return d.replace(/[.,;:]+$/, '')
}

function extractIds (raw) {
  const pmids = new Set()
  const dois = new Set()
  const pmcids = new Set()
  for (const re of PMID_RES) for (const m of raw.matchAll(re)) pmids.add(m[1])
  for (const m of raw.matchAll(/\bPMIDs?\s*[:#：]?\s*((?:\d{1,9}\s*[,、;]\s*)+\d{1,9})/gi)) {
    for (const id of m[1].split(/[,、;\s]+/)) if (id) pmids.add(id)
  }
  for (const m of raw.matchAll(DOI_RE)) dois.add(cleanDoi(m[1]).toLowerCase())
  for (const m of raw.matchAll(PMC_RE)) pmcids.add(m[1].toUpperCase())
  // 行全体が数字だけなら PMID とみなす（PMID を並べて貼った場合）
  const bare = raw.trim()
  if (/^\d{4,9}$/.test(bare)) pmids.add(bare)
  return { pmids: [...pmids], dois: [...dois], pmcids: [...pmcids] }
}

// AI の回答から、検索に使うタイトルらしい部分を取り出す
export function guessTitle (raw) {
  const segs = []
  for (const m of raw.matchAll(/\*\*([^*]+)\*\*|__([^_]+)__|[“"「]([^”"」]+)[”"」]|(?<![*\w])\*([^*]+)\*(?!\*)/g)) {
    segs.push((m[1] || m[2] || m[3] || m[4]).trim())
  }
  const text = plain(raw).replace(/https?:\/\/\S+/g, ' ')
  const good = s => s.split(/\s+/).length >= 4 && /[a-z]{3,}/.test(s) && !/^(doi|pmid|https?)/i.test(s)
  const quoted = segs.filter(good).sort((a, b) => b.length - a.length)[0]
  if (quoted) return quoted.replace(/[.。]$/, '')
  // ". " で区切った中で、小文字の単語を多く含む一番長い部分（著者名の並びは除く）
  const parts = text.split(/\.\s+|\s+\|\s+|—|–\s/).map(s => s.trim())
  const scored = parts
    .filter(good)
    .filter(s => !/^(doi|pmid|journal|authors?)\b/i.test(s))
    .map(s => ({ s, n: (s.match(/\b[a-z]{3,}\b/g) || []).length }))
    .sort((a, b) => b.n - a.n || b.s.length - a.s.length)
  return scored[0]?.s.replace(/[.。]$/, '') || ''
}

// AI の回答を、文献 1 件ずつのまとまりに分ける
export function parseAiText (text) {
  const lines = text.replace(/\r\n?/g, '\n').split('\n')
  const items = []
  let cur = null
  let table = null // 表の見出しから分かった PMID・DOI の列
  const startRe = /^(\s*)(\d{1,3}[.)]|\[\d{1,3}\]|[-*•+])\s+/
  for (const line of lines) {
    if (/^\s*```/.test(line)) { cur = null; continue }
    if (!line.trim() || /^\s*#/.test(line)) { cur = null; table = null; continue }
    if (/^\s*\|?\s*:?-{3,}/.test(line)) continue // 表の区切り行
    if (/^\s*\|/.test(line)) { // 表は 1 行 1 件
      cur = null
      const cells = line.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim())
      if (!table) {
        table = { pmid: cells.findIndex(c => /^pmid/i.test(c)), doi: cells.findIndex(c => /^doi/i.test(c)) }
        if (table.pmid >= 0 || table.doi >= 0) continue // 見出し行
      }
      let raw = line
      if (table.pmid >= 0 && /^\d{1,9}$/.test(cells[table.pmid] || '')) raw += ` PMID ${cells[table.pmid]}`
      if (table.doi >= 0 && cells[table.doi]) raw += ` ${cells[table.doi]}`
      items.push({ raw })
      continue
    }
    table = null
    const m = startRe.exec(line)
    const indent = (/^\s*/.exec(line)[0]).length
    // 箇条書きの中で、より深く字下げされた行はその項目の続き
    if (cur && cur.marker && indent > cur.indent) {
      cur.raw += '\n' + line
    } else {
      cur = { raw: line, indent, marker: !!m }
      items.push(cur)
    }
  }

  const out = []
  for (const { raw } of items) {
    const ids = extractIds(raw)
    const label = plain(raw.split('\n').map(l => l.trim()).join(' '))
    const year = /\b(19[5-9]\d|20[0-4]\d)\b/.exec(label)?.[1]
    const nIds = ids.pmids.length + ids.dois.length + ids.pmcids.length
    const textWords = label.replace(/https?:\/\/\S+|\b10\.\d{4,9}\/\S+|\bPMID\S*\s*\d+|\bPMC\d+/gi, ' ').split(/\s+/).filter(w => /[a-z]{2,}/i.test(w))
    // ID がない場合は、文献らしい書き方（et al.、巻;ページ、太字や引用符のタイトル）があるものだけ残す
    const citationLike = /\bet al\b|\d{4};\s*\d|\d+\(\d+\)|:\s*\d+[-–]\d+|\b(N Engl J Med|NEJM|Lancet|JAMA|BMJ|Circulation|JACC|Eur Heart J)\b/i.test(label) || /\*\*[^*]{20,}\*\*|[“"][^”"]{20,}[”"]/.test(raw)
    if (!nIds && !(year && textWords.length >= 6 && citationLike)) continue
    // 1 つのまとまりに PMID が複数（ID の羅列）なら 1 件ずつに分ける
    if (ids.pmids.length > 1 && !ids.dois.length) {
      for (const p of ids.pmids) out.push({ label: `PMID ${p}`, pmids: [p], dois: [], pmcids: [], title: '', year: '', checkable: false })
      continue
    }
    if (!ids.pmids.length && ids.dois.length > 1) {
      for (const d of ids.dois) out.push({ label: d, pmids: [], dois: [d], pmcids: [], title: '', year: '', checkable: false })
      continue
    }
    out.push({
      label: label.length > 300 ? label.slice(0, 300) + '…' : label,
      ...ids,
      title: guessTitle(raw),
      year: year || '',
      // ID 以外の文章が十分あれば、見つかった論文と食い違っていないか確かめる
      checkable: textWords.length >= 5
    })
  }
  // 同じ ID がすでに前の項目に出ていれば、その項目は省く（回答の最後の ID 一覧など）
  const seen = new Set()
  const unique = out.filter(it => {
    const ids = [...it.pmids.map(p => 'p' + p), ...it.dois.map(d => 'd' + d), ...it.pmcids.map(p => 'c' + p)]
    const dup = ids.length > 0 && ids.every(id => seen.has(id))
    ids.forEach(id => seen.add(id))
    return !dup
  })
  return unique.slice(0, 100)
}

// ---------- 論文を開くリンク ----------

export function linksFor (r) {
  const links = []
  if (r.PMID) links.push({ label: 'PubMed', url: `https://pubmed.ncbi.nlm.nih.gov/${r.PMID}/` })
  if (r.DOI) links.push({ label: '出版社', url: `https://doi.org/${r.DOI}` })
  else if (r.URL) links.push({ label: '論文のページ', url: r.URL })
  if (r.PMCID) links.push({ label: '全文（PMC）', url: `https://pmc.ncbi.nlm.nih.gov/articles/${r.PMCID}/` })
  return links
}
