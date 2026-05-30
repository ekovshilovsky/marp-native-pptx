import { existsSync } from 'node:fs'
import puppeteer, { type Browser } from 'puppeteer-core'

/** Discover a Chrome/Chromium executable from env or common locations. */
export function findChrome(): string {
  const fromEnv = process.env.PUPPETEER_EXECUTABLE_PATH || process.env.CHROME_PATH
  if (fromEnv) return fromEnv
  const candidates = [
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  ]
  for (const c of candidates) {
    if (existsSync(c)) return c
  }
  throw new Error(
    'No Chrome/Chromium found. Set CHROME_PATH or PUPPETEER_EXECUTABLE_PATH to a browser executable.',
  )
}

export async function launch(): Promise<Browser> {
  const args = ['--allow-file-access-from-files']
  if (process.env.CHROME_NO_SANDBOX === '1') args.push('--no-sandbox', '--disable-setuid-sandbox')
  return puppeteer.launch({ executablePath: findChrome(), args, headless: true })
}
