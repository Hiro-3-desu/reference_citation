// ブラウザでの動作確認用に、Word JavaScript API のうち RefCite が使う部分だけを真似たもの。
// 本物の Word の細かい挙動までは再現していない。
(() => {
  let uid = 0
  class Font { constructor () { this.superscript = false; this.italic = false; this.bold = false; this.subscript = false } }
  class Run { constructor (text) { this.text = text; this.font = new Font() } }
  class Para {
    constructor (text = '') { this.runs = text ? [new Run(text)] : [] }
    insertText (text, loc) {
      const r = new Run(text)
      if (loc === 'Start') this.runs.unshift(r)
      else this.runs.push(r)
      return { font: r.font }
    }
    insertParagraph (text, loc) {
      const p = new Para(text)
      const list = this.owner()
      const i = list.indexOf(this)
      list.splice(loc === 'Before' ? i : i + 1, 0, p)
      return p
    }
    owner () {
      if (doc.blocks.includes(this)) return doc.blocks
      return doc.blocks.find(b => b instanceof CC && b.paragraphs.includes(this)).paragraphs
    }
    insertContentControl () {
      const cc = new CC(true)
      cc.paragraphs = [this]
      const i = doc.blocks.indexOf(this)
      doc.blocks.splice(i, 1, cc)
      return cc
    }
  }
  class CC {
    constructor (block) { this.id = ++uid; this.block = block; this.tag = ''; this.title = ''; this.runs = []; this.paragraphs = [] }
    load () {}
    get isNullObject () { return false }
    insertText (text, loc) {
      if (this.block) return this.paragraphs[0].insertText(text, loc)
      const r = new Run(text)
      if (loc === 'Replace') this.runs = [r]
      else this.runs.push(r)
      return { font: r.font }
    }
    getRange (loc) {
      return { select: () => { if (loc === 'After') sel = { para: findPara(this), after: this } } }
    }
    clear () { this.paragraphs = [new Para()] }
    insertParagraph (text, loc) {
      const p = new Para(text)
      if (loc === 'End') this.paragraphs.push(p)
      else this.paragraphs.unshift(p)
      return p
    }
    get paragraphsCollection () { return this.paragraphs }
    delete (keep) {
      if (this.block) {
        const i = doc.blocks.indexOf(this)
        doc.blocks.splice(i, 1, ...(keep ? this.paragraphs : []))
      } else {
        const p = findPara(this)
        const i = p.runs.indexOf(this)
        p.runs.splice(i, 1, ...(keep ? this.runs : []))
      }
    }
  }
  // cc.paragraphs.getFirst() に対応
  Object.defineProperty(Array.prototype, 'getFirst', { value: function () { return this[0] }, enumerable: false })
  Object.defineProperty(Array.prototype, 'getLast', { value: function () { return this[this.length - 1] }, enumerable: false })

  function findPara (cc) {
    for (const b of doc.blocks) if (b instanceof Para && b.runs.includes(cc)) return b
  }
  function allCCs () {
    const out = []
    for (const b of doc.blocks) {
      if (b instanceof CC) out.push(b)
      else for (const r of b.runs) if (r instanceof CC) out.push(r)
    }
    return out
  }

  const doc = { blocks: [new Para('Primary PCI is preferred'), new Para('Transfer delays matter'), new Para('References')] }
  let sel = { para: doc.blocks[0], after: null } // after: この要素の後ろにカーソル（null なら段落末尾）

  const settingsStore = {}
  window.Office = {
    onReady: cb => Promise.resolve().then(() => cb({ host: 'Word' })),
    AsyncResultStatus: { Succeeded: 'succeeded' },
    context: {
      document: {
        settings: {
          get: k => settingsStore[k] ? JSON.parse(settingsStore[k]) : null,
          set: (k, v) => { settingsStore[k] = JSON.stringify(v) },
          saveAsync: cb => cb({ status: 'succeeded' })
        }
      }
    }
  }
  window.Word = {
    run: async cb => cb({
      sync: async () => {},
      document: {
        getSelection: () => ({
          get parentContentControlOrNullObject () {
            const cc = sel.inside
            return cc ? Object.assign(cc, {}) : { isNullObject: true, load () {} }
          },
          insertText (text) {
            const r = new Run(text)
            const runs = sel.para.runs
            const i = sel.after ? runs.indexOf(sel.after) + 1 : runs.length
            runs.splice(i, 0, r)
            return {
              insertContentControl () {
                const cc = new CC(false)
                cc.runs = [r]
                runs.splice(runs.indexOf(r), 1, cc)
                return cc
              }
            }
          },
          paragraphs: { getLast: () => sel.para }
        }),
        body: {
          get contentControls () {
            const items = allCCs()
            return { items, load () {}, getByTag: tag => ({ items: items.filter(c => c.tag === tag), load () {} }) }
          }
        }
      }
    })
  }
  // テストから使う補助
  window.__mock = {
    cursorAtEndOf: i => { sel = { para: doc.blocks.filter(b => b instanceof Para)[i], after: null } },
    cursorInsideCitation: n => { sel = { ...sel, inside: allCCs().filter(c => c.tag.startsWith('rmc:'))[n] } },
    clearInside: () => { delete sel.inside },
    text: () => doc.blocks.map(b => {
      const fmtRun = r => r instanceof CC
        ? `«${r.runs.map(fmtRun).join('')}»`
        : (r.font.superscript ? `^{${r.text}}` : r.font.italic ? `_${r.text}_` : r.text)
      if (b instanceof CC) return `[BIB]\n${b.paragraphs.map(p => '  ' + p.runs.map(fmtRun).join('')).join('\n')}\n[/BIB]`
      return b.runs.map(fmtRun).join('')
    }).join('\n'),
    settings: () => settingsStore
  }
})()
