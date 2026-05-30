#!/usr/bin/env node
import { convert } from './convert.js'

function parseArgs(argv: string[]) {
  const a = { input: '', out: '', themeSet: [] as string[], allowLocalFiles: false }
  for (let i = 0; i < argv.length; i++) {
    const t = argv[i]
    if (t === '-o' || t === '--output') a.out = argv[++i]
    else if (t === '--theme-set') a.themeSet.push(argv[++i])
    else if (t === '--allow-local-files') a.allowLocalFiles = true
    else if (!t.startsWith('-')) a.input = t
  }
  return a
}

async function main() {
  const a = parseArgs(process.argv.slice(2))
  if (!a.input || !a.out) {
    console.error('usage: marp-native-pptx <input.md> -o <out.pptx> [--theme-set <dir>] [--allow-local-files]')
    process.exit(1)
  }
  await convert(a.input, a.out, { themeSet: a.themeSet, allowLocalFiles: a.allowLocalFiles })
  console.log(`wrote ${a.out}`)
}
main().catch((e) => { console.error(e); process.exit(1) })
