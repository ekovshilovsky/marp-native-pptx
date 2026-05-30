// Optional OOXML *schema* validation via the bundled Open XML SDK validator
// (tools/ooxml-schema). Requires .NET; degrades gracefully when it's absent so
// the pure-Node PowerPoint-clean validator (validate.ts) still runs everywhere.
import { spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export interface SchemaResult {
  available: boolean
  errors: number
  sample: string[]
  note?: string
}

export function schemaValidate(pptxPath: string): SchemaResult {
  // tools/ is a sibling of src/ and dist/ (one level up from this module).
  const projDir = join(dirname(fileURLToPath(import.meta.url)), '..', 'tools', 'ooxml-schema')
  if (!existsSync(join(projDir, 'OoxmlSchema.csproj'))) {
    return { available: false, errors: 0, sample: [], note: 'bundled schema validator (tools/ooxml-schema) not present' }
  }
  const probe = spawnSync('dotnet', ['--version'], { encoding: 'utf8' })
  if (probe.status !== 0) {
    return { available: false, errors: 0, sample: [], note: 'dotnet not on PATH — install .NET to enable schema validation' }
  }
  const run = spawnSync('dotnet', ['run', '--project', projDir, '-c', 'Release', '--', pptxPath], {
    encoding: 'utf8',
    maxBuffer: 32 * 1024 * 1024,
  })
  const out = `${run.stdout ?? ''}${run.stderr ?? ''}`
  const m = out.match(/TOTAL ERRORS:\s*(\d+)/)
  if (!m) {
    return {
      available: false,
      errors: 0,
      sample: out.trim().split('\n').slice(-6),
      note: 'schema validator produced no result (build/runtime error — see sample)',
    }
  }
  return {
    available: true,
    errors: Number(m[1]),
    sample: out.split('\n').filter((l) => /^\[\d+\]/.test(l)).slice(0, 15),
  }
}
