#!/usr/bin/env node
import { validatePptx } from './validate.js'
import { schemaValidate } from './schema-validate.js'

async function main(): Promise<void> {
  const argv = process.argv.slice(2)
  const file = argv.find((a) => !a.startsWith('-'))
  const noSchema = argv.includes('--no-schema')
  if (!file) {
    console.error('usage: marp-native-pptx-validate <file.pptx> [--no-schema]')
    process.exit(2)
  }

  const findings = await validatePptx(file)
  if (findings.length === 0) {
    console.log('✓ PowerPoint-clean rules: 0 violations')
  } else {
    console.log(`✗ PowerPoint-clean rules: ${findings.length} violation(s)`)
    for (const f of findings) console.log(`    [${f.rule}] ${f.part ? f.part + ' — ' : ''}${f.detail}`)
  }

  let schemaErrors = 0
  if (!noSchema) {
    const s = schemaValidate(file)
    if (!s.available) {
      console.log(`• schema (OpenXmlValidator): skipped — ${s.note}`)
    } else if (s.errors === 0) {
      console.log('✓ schema (OpenXmlValidator): 0 errors')
    } else {
      schemaErrors = s.errors
      console.log(`✗ schema (OpenXmlValidator): ${s.errors} error(s)`)
      for (const l of s.sample) console.log(`    ${l}`)
    }
  }

  process.exit(findings.length + schemaErrors > 0 ? 1 : 0)
}

main().catch((e) => {
  console.error(e)
  process.exit(2)
})
