/* global Office, Word */
// Word とのやりとり。
// 引用はコンテンツコントロール（タグ "rmc:<key>"）、文献リストはタグ "rmbib" で文書に埋め込む。
// どの引用がどの文献を指すかと文献データは、文書の設定（document.settings）に保存するので、
// 文書ファイルだけで番号の振り直しや文献リストの再作成ができる。

import { computeNumbering, citationLabel, formatBibliography, newId } from './core.js'

const SETTINGS_KEY = 'refcite'
const CITE_PREFIX = 'rmc:'
const BIB_TAG = 'rmbib'

export function inWord () {
  return typeof Office !== 'undefined' && typeof Word !== 'undefined' && Office.context?.document
}

export function loadDocState () {
  const s = inWord() ? Office.context.document.settings.get(SETTINGS_KEY) : null
  return { v: 1, style: 'bmj', marker: 'superscript', refs: {}, cites: {}, ...(s || {}) }
}

export function saveDocState (state) {
  return new Promise((resolve, reject) => {
    if (!inWord()) return resolve()
    const settings = Office.context.document.settings
    settings.set(SETTINGS_KEY, state)
    settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error))
  })
}

function citeKeyOf (cc) {
  return cc.tag && cc.tag.startsWith(CITE_PREFIX) ? cc.tag.slice(CITE_PREFIX.length) : null
}

// カーソル位置に引用を挿入する。カーソルが既存の引用の中なら、その引用に文献を追加する。
export async function insertCitation (state, refs) {
  for (const r of refs) state.refs[r.id] = r
  await Word.run(async ctx => {
    const sel = ctx.document.getSelection()
    const parent = sel.parentContentControlOrNullObject
    parent.load('tag')
    await ctx.sync()

    if (!parent.isNullObject && parent.tag === BIB_TAG) {
      throw new Error('文献リストの中には引用を挿入できません。本文にカーソルを置いてください。')
    }
    const existing = parent.isNullObject ? null : citeKeyOf(parent)
    if (existing) {
      const ids = state.cites[existing] || []
      state.cites[existing] = [...new Set([...ids, ...refs.map(r => r.id)])]
      return
    }

    const key = newId('c')
    state.cites[key] = refs.map(r => r.id)
    const range = sel.insertText('[…]', 'End')
    const cc = range.insertContentControl()
    cc.tag = CITE_PREFIX + key
    cc.title = '引用'
    cc.appearance = 'BoundingBox'
    cc.getRange('After').select()
    await ctx.sync()
  })
  await saveDocState(state)
  return refresh(state)
}

// カーソル位置の引用を返す（なければ null）
export async function getSelectedCitation (state) {
  return Word.run(async ctx => {
    const parent = ctx.document.getSelection().parentContentControlOrNullObject
    parent.load('tag')
    await ctx.sync()
    const key = parent.isNullObject ? null : citeKeyOf(parent)
    return key ? { key, refIds: state.cites[key] || [] } : null
  })
}

// 引用の文献を差し替える。空になったら引用ごと消す。
export async function setCitationRefs (state, key, refIds) {
  await Word.run(async ctx => {
    if (refIds.length) {
      state.cites[key] = refIds
      return
    }
    const ccs = ctx.document.body.contentControls.getByTag(CITE_PREFIX + key)
    ccs.load('items')
    await ctx.sync()
    for (const cc of ccs.items) cc.delete(false)
    delete state.cites[key]
    await ctx.sync()
  })
  await saveDocState(state)
  return refresh(state)
}

// カーソルのある段落の後ろに文献リストを挿入する（すでにあれば更新だけ）
export async function insertBibliography (state) {
  const exists = await Word.run(async ctx => {
    const bibs = ctx.document.body.contentControls.getByTag(BIB_TAG)
    bibs.load('items')
    await ctx.sync()
    if (bibs.items.length) return true
    const para = ctx.document.getSelection().paragraphs.getLast()
    const p = para.insertParagraph('文献リスト', 'After')
    const cc = p.insertContentControl()
    cc.tag = BIB_TAG
    cc.title = '文献リスト'
    cc.appearance = 'BoundingBox'
    await ctx.sync()
    return false
  })
  const result = await refresh(state)
  return { ...result, existed: exists }
}

// 文書中の引用の番号を振り直し、文献リストを作り直す
export async function refresh (state) {
  return Word.run(async ctx => {
    const ccs = ctx.document.body.contentControls
    ccs.load('items/tag')
    await ctx.sync()

    const cites = ccs.items.filter(cc => citeKeyOf(cc))
    const bibs = ccs.items.filter(cc => cc.tag === BIB_TAG)
    const citations = cites.map(cc => {
      const key = citeKeyOf(cc)
      return { key, refIds: (state.cites[key] || []).filter(id => state.refs[id]) }
    })
    const { order, labels } = computeNumbering(citations)

    cites.forEach((cc, i) => {
      const { text, superscript } = labels[i].length
        ? citationLabel(labels[i], state.marker)
        : { text: '[?]', superscript: false }
      const range = cc.insertText(text, 'Replace')
      range.font.superscript = superscript
    })

    const entries = formatBibliography(order.map(id => state.refs[id]), state.style)
    for (const bib of bibs) writeBibliography(bib, entries)
    await ctx.sync()
    return { citations: cites.length, references: order.length, bibliographies: bibs.length, usedIds: order }
  })
}

function writeBibliography (cc, entries) {
  cc.clear()
  if (!entries.length) {
    cc.insertText('（まだ引用がありません）', 'Start')
    return
  }
  entries.forEach((runs, i) => {
    // 1 件目は clear() の後に残る空の段落に書き、2 件目からは段落を足す
    const p = i === 0 ? cc.paragraphs.getFirst() : cc.insertParagraph('', 'End')
    for (const run of runs) {
      const r = p.insertText(run.text, 'End')
      r.font.italic = run.italic
      r.font.bold = run.bold
      r.font.superscript = run.superscript
      r.font.subscript = run.subscript
    }
  })
}

// 投稿用に、引用と文献リストをふつうの文字に変換する（元に戻せない）
export async function convertToPlainText () {
  return Word.run(async ctx => {
    const ccs = ctx.document.body.contentControls
    ccs.load('items/tag')
    await ctx.sync()
    let n = 0
    for (const cc of ccs.items) {
      if (citeKeyOf(cc) || cc.tag === BIB_TAG) {
        cc.delete(true)
        n++
      }
    }
    await ctx.sync()
    return n
  })
}

// 文書を変更せずに、引用されている文献を番号順に返す
export async function scan (state) {
  return Word.run(async ctx => {
    const ccs = ctx.document.body.contentControls
    ccs.load('items/tag')
    await ctx.sync()
    const citations = ccs.items.filter(cc => citeKeyOf(cc)).map(cc => ({
      refIds: (state.cites[citeKeyOf(cc)] || []).filter(id => state.refs[id])
    }))
    return computeNumbering(citations).order
  })
}
