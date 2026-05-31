// Generated cover/tile art for the showcase.
//
// We can't ship stock photography, so these are cohesive per-theme gradient
// compositions (a base gradient + soft palette "blobs") that read as real
// brand imagery, sized to the layout's image regions. They're ordinary raster
// images once baked to PNG — drop in real photos of the same dimensions and
// every image-driven layout works unchanged.

// A full-bleed 16:9 cover with a left-side dark scrim baked in, so white
// overlay text stays legible on any theme's colors.
export function coverSvg(palette: string[]): string {
  const a = palette[0]
  const b = palette[1] ?? palette[0]
  const c = palette[2] ?? palette[0]
  const d = palette[3] ?? palette[1] ?? palette[0]
  return (
    `<svg viewBox="0 0 1600 900" xmlns="http://www.w3.org/2000/svg">` +
    `<defs>` +
    `<linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#${a}"/><stop offset="1" stop-color="#${b}"/></linearGradient>` +
    `<radialGradient id="r1" cx="0.72" cy="0.28" r="0.65"><stop offset="0" stop-color="#${c}" stop-opacity="0.85"/><stop offset="1" stop-color="#${c}" stop-opacity="0"/></radialGradient>` +
    `<radialGradient id="r2" cx="0.9" cy="0.85" r="0.55"><stop offset="0" stop-color="#${d}" stop-opacity="0.8"/><stop offset="1" stop-color="#${d}" stop-opacity="0"/></radialGradient>` +
    `<linearGradient id="scrim" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="#000" stop-opacity="0.66"/><stop offset="0.6" stop-color="#000" stop-opacity="0"/></linearGradient>` +
    `</defs>` +
    `<rect width="1600" height="900" fill="url(#g)"/>` +
    `<rect width="1600" height="900" fill="url(#r1)"/>` +
    `<rect width="1600" height="900" fill="url(#r2)"/>` +
    `<rect width="1600" height="900" fill="url(#scrim)"/>` +
    `</svg>`
  )
}

// An adversarial calibration "test card" at an arbitrary aspect. The centered
// circle is the tell: a correct cover/contain fit keeps it round, while the old
// stretch-to-box path squashes it into an ellipse. Corner brackets show whether
// the crop took from the right edges, and the label states the source aspect.
// Use it to eyeball that a non-square source fits a differently-shaped region.
export function calibrationSvg(w: number, h: number, label = `${w}x${h}`): string {
  const r = Math.min(w, h) * 0.36
  const cx = w / 2
  const cy = h / 2
  const arm = Math.min(w, h) * 0.12
  const m = Math.min(w, h) * 0.06 // bracket inset from each corner
  const bracket = (x: number, y: number, dx: number, dy: number): string =>
    `<path d="M ${x + dx * arm} ${y} H ${x} V ${y + dy * arm}" fill="none" stroke="#fff" stroke-width="6"/>`
  return (
    `<svg viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg">` +
    `<defs><linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#11131a"/><stop offset="1" stop-color="#2b3350"/></linearGradient></defs>` +
    `<rect width="${w}" height="${h}" fill="url(#bg)"/>` +
    // concentric rings — eccentricity reveals any non-uniform scaling
    `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#ff5d73" stroke-width="8"/>` +
    `<circle cx="${cx}" cy="${cy}" r="${r * 0.6}" fill="none" stroke="#ffd166" stroke-width="6"/>` +
    `<line x1="${cx}" y1="${cy - r}" x2="${cx}" y2="${cy + r}" stroke="#fff" stroke-width="3"/>` +
    `<line x1="${cx - r}" y1="${cy}" x2="${cx + r}" y2="${cy}" stroke="#fff" stroke-width="3"/>` +
    bracket(m, m, 1, 1) +
    bracket(w - m, m, -1, 1) +
    bracket(m, h - m, 1, -1) +
    bracket(w - m, h - m, -1, -1) +
    `<text x="${cx}" y="${h - m * 1.5}" fill="#fff" font-family="Arial" font-size="${Math.min(w, h) * 0.08}" text-anchor="middle">${label}</text>` +
    `</svg>`
  )
}

// A square tile (no scrim) for image-feature panels / galleries. `seed` varies
// the composition so a row of tiles looks distinct.
export function tileSvg(palette: string[], seed: number): string {
  const a = palette[seed % palette.length]
  const b = palette[(seed + 2) % palette.length]
  const c = palette[(seed + 4) % palette.length]
  const cx = (0.3 + 0.4 * ((seed % 3) / 2)).toFixed(3)
  const cy = (0.3 + 0.3 * (seed % 2)).toFixed(3)
  return (
    `<svg viewBox="0 0 560 560" xmlns="http://www.w3.org/2000/svg">` +
    `<defs>` +
    `<linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#${a}"/><stop offset="1" stop-color="#${b}"/></linearGradient>` +
    `<radialGradient id="r" cx="${cx}" cy="${cy}" r="0.6"><stop offset="0" stop-color="#${c}" stop-opacity="0.85"/><stop offset="1" stop-color="#${c}" stop-opacity="0"/></radialGradient>` +
    `</defs>` +
    `<rect width="560" height="560" fill="url(#g)"/>` +
    `<rect width="560" height="560" fill="url(#r)"/>` +
    `</svg>`
  )
}
