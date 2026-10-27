import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import * as c from '../src/core.js'

test('番号は初出順に振り、引用ごとに並べ替えて範囲にまとめる', () => {
  const { order, labels } = c.computeNumbering([
    { refIds: ['a', 'b'] }, { refIds: ['c', 'a'] }, { refIds: ['d', 'e', 'f', 'b'] }
  ])
  assert.deepEqual(order, ['a', 'b', 'c', 'd', 'e', 'f'])
  assert.deepEqual(labels, [[1, 2], [1, 3], [2, 4, 5, 6]])
  assert.equal(c.compressNumbers([2, 4, 5, 6]), '2,4–6')
  assert.equal(c.compressNumbers([1, 2]), '1,2')
  assert.equal(c.citationLabel([1, 2, 3], 'brackets').text, '[1–3]')
  assert.equal(c.citationLabel([1, 2, 3], 'superscript').superscript, true)
})

test('RIS を読み込み、雑誌の略名も拾う', () => {
  const items = c.parseImportText(readFileSync(new URL('./sample.ris', import.meta.url), 'utf8'))
  assert.equal(items.length, 3)
  assert.equal(items[1]['container-title-short'], 'N Engl J Med')
  assert.equal(items[1].author[0].family, 'Andersen')
})

test('PubMed の要約を CSL に変換する', () => {
  const r = c.pubmedSummaryToCsl({
    uid: '1', pubdate: '2020 Oct 21', sortpubdate: '2020/10/21 00:00', source: 'BMJ', title: 'A title.',
    authors: [{ name: 'Smith JA', authtype: 'Author' }, { name: 'ABC Group', authtype: 'CollectiveName' }],
    articleids: [{ idtype: 'doi', value: '10.1/x' }], pages: '', elocationid: 'pii: m3339. doi: 10.1/x'
  })
  assert.equal(r.title, 'A title')
  assert.deepEqual(r.author, [{ family: 'Smith', given: 'JA' }, { literal: 'ABC Group' }])
  assert.deepEqual(r.issued, { 'date-parts': [[2020, 10]] })
  assert.equal(r.page, 'm3339')
  assert.equal(r.DOI, '10.1/x')
})

test('どのスタイルでも文献リストを作れる', () => {
  const items = c.parseImportText(readFileSync(new URL('./sample.ris', import.meta.url), 'utf8'))
  for (const s of c.STYLE_LIST) {
    const bib = c.formatBibliography(items, s.id)
    assert.equal(bib.length, 3, s.id)
    assert.match(c.runsToText(bib[0]), /Keeley/, s.id)
  }
})

test('citeproc の HTML を書式付きの文字列片にする', () => {
  const runs = c.htmlToRuns('<div class="csl-entry"><div class="csl-left-margin">1. </div><div class="csl-right-inline">A <i>B</i> &amp; <b>C</b>.</div></div>')
  assert.equal(c.runsToText(runs), '1. A B & C.')
  assert.deepEqual(runs.filter(r => r.italic).map(r => r.text), ['B'])
})

const fixture = name => readFileSync(new URL(`./fixtures/${name}.md`, import.meta.url), 'utf8')

test('Claude の回答から、ID・リンク・ID なしの文献を拾う', () => {
  const items = c.parseAiText(fixture('claude'))
  assert.equal(items.length, 3) // 最後の英語の文章は文献として拾わない
  assert.deepEqual(items[0].pmids, ['12517460'])
  assert.deepEqual(items[0].dois, ['10.1016/s0140-6736(03)12113-7'])
  assert.deepEqual(items[1].pmids, ['29138293']) // 字下げした箇条書きのリンクも同じ文献
  assert.deepEqual(items[2].pmids, [])
  assert.equal(items[2].title, 'A comparison of coronary angioplasty with fibrinolytic therapy in acute myocardial infarction')
})

test('Codex の回答から、DOI リンク・表の PMID 列・ID 一覧を拾い、重複は省く', () => {
  const items = c.parseAiText(fixture('codex'))
  assert.deepEqual(items.map(it => [...it.pmids, ...it.dois]), [
    ['10.1016/s0140-6736(03)12113-7'],
    ['10.1056/nejmoa025142'],
    ['12930925'],
    ['12684343'],
    ['12517460']
  ])
})

test('ID だけを並べて貼った場合は 1 行 1 件', () => {
  assert.equal(c.parseAiText('33087347\n10.1136/bmj.m3339\nPMID: 12517460').length, 3)
  assert.equal(c.parseAiText('PMIDs: 1111111, 2222222, 3333333').length, 3)
})

test('タイトルの一致度', () => {
  assert.ok(c.titleSimilarity('A comparison of coronary angioplasty with fibrinolytic therapy', 'Andersen HR. A comparison of coronary angioplasty with fibrinolytic therapy. NEJM') > 0.9)
  assert.ok(c.titleSimilarity('Dapagliflozin in heart failure with reduced ejection fraction', 'Keeley EC. Primary angioplasty versus thrombolysis. Lancet') < 0.2)
})
