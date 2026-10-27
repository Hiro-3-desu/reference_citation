// CSL スタイル集（citation-style-language/styles）から、検索用の一覧を作る。
// 使い方: node scripts/style-index.mjs <styles のフォルダ> <出力する JSON>
import { readFileSync, readdirSync, writeFileSync } from 'node:fs'
import { execSync } from 'node:child_process'
import path from 'node:path'

const [dir, out] = process.argv.slice(2)
const tag = (xml, re) => re.exec(xml)?.[1]?.trim()
const unescape = s => s && s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")

function read (file, dependent) {
  const xml = readFileSync(file, 'utf8')
  const info = tag(xml, /<info>([\s\S]*?)<\/info>/) || ''
  const entry = {
    id: path.basename(file, '.csl'),
    t: unescape(tag(info, /<title>([\s\S]*?)<\/title>/)),
    f: tag(info, /citation-format="([^"]+)"/)
  }
  const short = unescape(tag(info, /<title-short>([\s\S]*?)<\/title-short>/))
  if (short) entry.s = short
  if (dependent) entry.p = tag(info, /href="https?:\/\/www\.zotero\.org\/styles\/([^"]+)"\s+rel="independent-parent"/) || tag(info, /rel="independent-parent"\s+href="https?:\/\/www\.zotero\.org\/styles\/([^"]+)"/)
  return entry
}

const styles = [
  ...readdirSync(dir).filter(f => f.endsWith('.csl')).map(f => read(path.join(dir, f), false)),
  ...readdirSync(path.join(dir, 'dependent')).filter(f => f.endsWith('.csl')).map(f => read(path.join(dir, 'dependent', f), true))
].filter(s => s.t && s.f === 'numeric' && (!('p' in s) || s.p)) // 番号で引用する形式だけ
let commit = 'master'
try { commit = execSync('git rev-parse HEAD', { cwd: dir }).toString().trim() } catch { /* git でなければ master */ }
writeFileSync(out, JSON.stringify({ commit, styles }))
console.log(`${styles.length} styles (commit ${commit.slice(0, 7)}) -> ${out}`)
