import { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } from 'docx'
import { saveAs } from 'file-saver'
import { latexToOmml, extractLatex, isKatexDisplay, isInsideKatex } from './latex2omml'
import type { StyleConfig } from './style'

function docxFont(cssFontFamily: string): string {
  // Map CSS font-family values to docx-friendly Windows font names
  if (cssFontFamily.includes('Songti') || cssFontFamily.includes('Noto Serif')) return 'SimSun'
  if (cssFontFamily.includes('KaiTi') || cssFontFamily.includes('STKaiti')) return 'KaiTi'
  return 'Microsoft YaHei'
}

// ---- COPY support ----

// Registry for OMML strings during clipboard build — avoids HTML attribute escaping issues
const copyOmmlRegistry: string[] = []

/**
 * Build clipboard HTML and plain text from the preview DOM.
 *
 * Strategy:
 * - HTML: clone DOM, replace KaTeX with OMML placeholder markers (index-based),
 *   serialize to HTML string, then replace markers with raw OMML XML.
 *   Word/WPS paste this and render native math formulas.
 * - Plain text: extract readable Unicode text, replace KaTeX with Unicode
 *   approximation (from textContent, but without annotation duplicates),
 *   replace Mermaid with placeholder.
 */
function buildClipboardData(docEl: Element, style: string): { html: string; plain: string } {
  copyOmmlRegistry.length = 0
  const clone = docEl.cloneNode(true) as Element

  // === Build HTML: replace KaTeX with OMML markers ===
  // Step 1: Extract LaTeX from annotations and generate OMML BEFORE removing mathml.
  // The annotation elements live inside .katex-mathml, so we must read them first.

  // Replace display math with OMML markers
  const displayBlocks = clone.querySelectorAll('.katex-display')
  displayBlocks.forEach(block => {
    const latexInfo = extractLatex(block)
    if (latexInfo) {
      try {
        const omml = latexToOmml(latexInfo.latex, true)
        const idx = copyOmmlRegistry.length
        copyOmmlRegistry.push(omml)
        const placeholder = document.createElement('span')
        placeholder.textContent = `@@PASTIFY_OMML_${idx}@@`
        block.replaceWith(placeholder)
      } catch { /* keep original */ }
    }
  })

  // Replace inline math with OMML markers
  const inlineBlocks = clone.querySelectorAll('.katex')
  inlineBlocks.forEach(block => {
    if (block.closest('.katex-display')) return

    const latexInfo = extractLatex(block)
    if (latexInfo) {
      try {
        const omml = latexToOmml(latexInfo.latex, false)
        const idx = copyOmmlRegistry.length
        copyOmmlRegistry.push(omml)
        const placeholder = document.createElement('span')
        placeholder.textContent = `@@PASTIFY_OMML_${idx}@@`
        block.replaceWith(placeholder)
      } catch { /* keep original */ }
    }
  })

  // Step 2: Now that LaTeX has been extracted, remove any remaining .katex-mathml
  // (for formulas where OMML conversion failed and the original KaTeX HTML remains).
  // .katex-mathml contains <math> MathML + <annotation> with raw LaTeX source —
  // keeping it causes DUPLICATE formula text in both innerHTML and textContent.
  clone.querySelectorAll('.katex-mathml').forEach(m => m.remove())

  // Clean Mermaid: remove SVG and inline styles
  clone.querySelectorAll('pre.mermaid svg').forEach(svg => svg.remove())
  clone.querySelectorAll('pre.mermaid style').forEach(s => s.remove())

  // Get innerHTML, then replace OMML markers with raw OMML XML
  let innerHtml = clone.innerHTML
  for (let i = 0; i < copyOmmlRegistry.length; i++) {
    const marker = `@@PASTIFY_OMML_${i}@@`
    // innerHTML will have the marker as plain text content inside a <span>
    // It may appear as: <span>@@PASTIFY_OMML_0@@</span>
    innerHtml = innerHtml.replace(
      new RegExp(`<span[^>]*>${marker.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')}</span>`, 'g'),
      copyOmmlRegistry[i]
    )
    // Also handle case where marker is not wrapped in span (shouldn't happen, but just in case)
    innerHtml = innerHtml.replace(marker, copyOmmlRegistry[i])
  }

  const html = `<!DOCTYPE html>
<html xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
<head><meta charset="UTF-8"><style>${style}</style>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]-->
</head>
<body>${innerHtml}</body></html>`

  // === Build plain text: readable Unicode text ===
  const plainClone = docEl.cloneNode(true) as Element
  // Remove KaTeX MathML layer (contains <math> + <annotation> with raw LaTeX)
  // Only keep .katex-html which has Unicode math symbols (∫ ∑ α β)
  plainClone.querySelectorAll('.katex-mathml').forEach(m => m.remove())
  // Remove Mermaid SVG and styles
  plainClone.querySelectorAll('pre.mermaid svg').forEach(svg => svg.remove())
  plainClone.querySelectorAll('pre.mermaid style').forEach(s => s.remove())
  plainClone.querySelectorAll('style').forEach(s => s.remove())
  plainClone.querySelectorAll('script').forEach(s => s.remove())
  // Get text content — KaTeX renders Unicode math symbols in its spans,
  // so textContent gives readable formulas like ∫, ∑, α, β etc.
  const plain = (plainClone.textContent || '')
    .replace(/\n{3,}/g, '\n\n')
    .trim()

  return { html, plain }
}

export function copyRichText(docEl: Element, style: string): void {
  const { html, plain } = buildClipboardData(docEl, style)

  // Primary: Clipboard API with HTML + plain text
  try {
    const htmlBlob = new Blob([html], { type: 'text/html' })
    const plainBlob = new Blob([plain], { type: 'text/plain' })
    navigator.clipboard.write([
      new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': plainBlob }),
    ])
    return
  } catch (_) { /* fallthrough to backup */ }

  // Backup: contentEditable + execCommand (for browsers without ClipboardItem)
  const el = document.createElement('div')
  el.contentEditable = 'true'
  el.innerHTML = html
  el.style.cssText = 'position:fixed;left:-9999px'
  document.body.appendChild(el)
  el.focus()
  const sel = window.getSelection()!
  const range = document.createRange()
  range.selectNodeContents(el)
  sel.removeAllRanges()
  sel.addRange(range)
  document.execCommand('copy')
  sel.removeAllRanges()
  document.body.removeChild(el)
}

// ---- DOCX export ----

const H_LEVEL: Record<string,any> = {
  h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2,
  h3: HeadingLevel.HEADING_3, h4: HeadingLevel.HEADING_4,
  h5: HeadingLevel.HEADING_5, h6: HeadingLevel.HEADING_6,
}

function t(n: Node): string { return (n.textContent || '').trim() }

// ---- Mermaid SVG → PNG conversion ----

interface MermaidEntry {
  svg: SVGElement
  fallbackText: string
  paragraphIndex: number
}

let _mermaidEntries: MermaidEntry[] = []

async function svgToPngBlob(svgElement: SVGElement): Promise<Blob> {
  // Deep clone to avoid mutating the original DOM
  const clone = svgElement.cloneNode(true) as SVGElement

  // Strip <style> tags that may contain @import / @font-face referencing external URLs.
  // External references taint the canvas, causing SecurityError on toBlob().
  clone.querySelectorAll('style').forEach(s => s.remove())

  // Ensure proper SVG namespace
  if (!clone.getAttribute('xmlns')) {
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
  }

  const svgData = new XMLSerializer().serializeToString(clone)
  const svgBlob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' })
  const url = URL.createObjectURL(svgBlob)

  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => {
      URL.revokeObjectURL(url)
      const rect = svgElement.getBoundingClientRect()
      const w = rect.width || img.naturalWidth || 600
      const h = rect.height || img.naturalHeight || 400
      const scale = 2
      const canvas = document.createElement('canvas')
      canvas.width = w * scale
      canvas.height = h * scale
      const ctx = canvas.getContext('2d')!
      ctx.scale(scale, scale)
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, w, h)
      try {
        ctx.drawImage(img, 0, 0, w, h)
        canvas.toBlob(blob => {
          if (blob) resolve(blob)
          else reject(new Error('canvas.toBlob returned null'))
        }, 'image/png')
      } catch (e) {
        reject(new Error(`Canvas tainted or draw failed: ${e}`))
      }
    }
    img.onerror = (e) => { URL.revokeObjectURL(url); reject(new Error(`SVG image load failed: ${e}`)) }
    img.src = url
  })
}

// OMML placeholder marker — pure ASCII to avoid XML encoding issues.
// The docx library may encode non-ASCII chars as XML character references,
// so we use a simple ASCII prefix that survives round-trip serialization.
const OMML_MARKER_PREFIX = '__PASTIFY_OMML_'
const ommlRegistry: string[] = []

function registerOmml(omml: string): string {
  const id = ommlRegistry.length
  ommlRegistry.push(omml)
  return `${OMML_MARKER_PREFIX}${id}__`
}

function mermaidFallbackPara(entry: MermaidEntry): Paragraph {
  return new Paragraph({
    children: [new TextRun({ text: '[流程图] ' + entry.fallbackText.slice(0, 120), italics: true, size: 20, color: '666666' })],
    spacing: { after: 120 },
    shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' },
  })
}

export async function downloadDocx(docEl: Element, _style: string, filename: string, cfg?: StyleConfig): Promise<void> {
  // Reset registries
  ommlRegistry.length = 0
  _mermaidEntries = []

  let children: (Paragraph | Table)[]
  try {
    children = walk(docEl, cfg)
  } catch (e) {
    console.error('[Pastify DOCX] walk() failed:', e)
    throw new Error(`DOM walk 失败: ${e}`)
  }

  // Process Mermaid SVGs into ImageRun paragraphs
  if (_mermaidEntries.length > 0) {
    const pngResults = await Promise.allSettled(
      _mermaidEntries.map(e => svgToPngBlob(e.svg).then(blob => ({ blob, entry: e })))
    )
    const replacements: Promise<void>[] = []
    pngResults.forEach((result, i) => {
      const entry = _mermaidEntries[i]
      if (result.status === 'fulfilled') {
        const { blob } = result.value
        const svgEl = entry.svg
        const rect = svgEl.getBoundingClientRect()
        const w = rect.width || 600
        const h = rect.height || 400
        const maxW = 5500
        const scale = Math.min(maxW / (w * 9525 / 72), 1)
        replacements.push(
          blob.arrayBuffer().then(buf => {
            children[entry.paragraphIndex] = new Paragraph({
              children: [new ImageRun({
                type: 'png',
                data: buf,
                transformation: { width: Math.round(w * scale * 9525 / 72), height: Math.round(h * scale * 9525 / 72) },
              })],
              spacing: { before: 120, after: 120 },
              alignment: AlignmentType.CENTER,
            })
          }).catch(() => {
            children[entry.paragraphIndex] = mermaidFallbackPara(entry)
          })
        )
      } else {
        children[entry.paragraphIndex] = mermaidFallbackPara(entry)
      }
    })
    await Promise.allSettled(replacements)
  }

  const doc = new Document({
    sections: [{ properties: {}, children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })] }],
  })

  let blob: Blob
  try {
    blob = await Packer.toBlob(doc)
  } catch (e) {
    console.error('[Pastify DOCX] Packer.toBlob failed:', e)
    throw new Error(`Packer.toBlob 失败: ${e}`)
  }

  let finalBlob: Blob
  try {
    finalBlob = await replaceOmmlMarkers(blob)
  } catch (e) {
    console.error('[Pastify DOCX] replaceOmmlMarkers failed:', e)
    finalBlob = blob
  }

  saveAs(finalBlob, filename)
}

async function replaceOmmlMarkers(blob: Blob): Promise<Blob> {
  if (ommlRegistry.length === 0) return blob

  try {
    const JSZip = (await import('jszip')).default
    const zip = await JSZip.loadAsync(blob)

    const docFile = zip.file('word/document.xml')
    if (!docFile) {
      console.warn('Pastify: word/document.xml not found in docx zip')
      return blob
    }

    let docXml = await docFile.async('string')

    // Ensure math namespace is declared (required for OMML elements)
    if (!docXml.includes('xmlns:m=')) {
      docXml = docXml.replace(
        /<w:document\s/,
        '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      )
    }

    // Use DOM parser to safely find and replace OMML markers
    // Regex can span across w:r boundaries and consume adjacent text runs
    const parser = new DOMParser()
    const xmlDoc = parser.parseFromString(docXml, 'application/xml')

    const parseError = xmlDoc.querySelector('parsererror')
    if (parseError) {
      console.warn('Pastify: XML parse error, document may be malformed')
      // Fall back to regex as last resort
      return replaceOmmlMarkersRegex(docXml, zip)
    }

    // Find all w:t elements whose text content is an OMML marker
    const wtNodes = xmlDoc.querySelectorAll('w\\:t, t')
    const replacements: { wr: Element; omml: string }[] = []

    for (const wt of wtNodes) {
      const text = wt.textContent || ''
      const match = text.match(/^__PASTIFY_OMML_(\d+)__$/)
      if (!match) continue
      const idx = parseInt(match[1], 10)
      if (idx >= ommlRegistry.length) continue

      // Get the parent w:r element
      const wr = wt.parentElement
      if (!wr || (wr.tagName !== 'w:r' && wr.tagName !== 'r')) continue

      replacements.push({ wr, omml: ommlRegistry[idx] })
    }

    // Replace each w:r with the OMML XML
    for (const { wr, omml } of replacements) {
      try {
        const ommlDoc = parser.parseFromString(omml, 'application/xml')
        const ommlEl = ommlDoc.documentElement
        if (wr.parentElement) {
          wr.parentElement.replaceChild(ommlEl, wr)
        }
      } catch (e) {
        console.warn('Pastify: failed to inject OMML element:', e)
      }
    }

    docXml = new XMLSerializer().serializeToString(xmlDoc)

    zip.file('word/document.xml', docXml)

    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    })
  } catch (err) {
    console.error('Pastify: OMML post-processing failed, returning original blob:', err)
    return blob
  }
}

// Regex fallback kept for XML parse error scenarios
async function replaceOmmlMarkersRegex(docXml: string, zip: any): Promise<Blob> {
  for (let i = 0; i < ommlRegistry.length; i++) {
    const marker = `${OMML_MARKER_PREFIX}${i}__`
    const esc = marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const re = new RegExp(
      `<w:r[^>]*>(?:<w:rPr[^>]*(?:/>|>.*?</w:rPr>))?<w:t[^>]*>${esc}</w:t></w:r>`,
      'gs'
    )
    if (re.test(docXml)) {
      docXml = docXml.replace(re, ommlRegistry[i])
    }
  }
  zip.file('word/document.xml', docXml)
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  })
}

function walk(el: Element, cfg?: StyleConfig): (Paragraph | Table)[] {
  // Compute sizes dynamically from user-selected font size (px → docx half-points)
  const baseHalfPt = Math.round((cfg?.fontSize || 16) * 1.5)
  const H_SIZES: Record<string,number> = {
    h1: Math.round(baseHalfPt * 1.55), h2: Math.round(baseHalfPt * 1.3),
    h3: Math.round(baseHalfPt * 1.12), h4: Math.round(baseHalfPt * 1.05),
    h5: Math.round(baseHalfPt * 1.05), h6: Math.round(baseHalfPt * 1.05),
  }
  const bodySize = baseHalfPt
  const font = cfg ? docxFont(cfg.fontFamily) : 'Microsoft YaHei'

  const out: (Paragraph | Table)[] = []
  for (const n of Array.from(el.childNodes)) {
    if (n.nodeType === Node.COMMENT_NODE) continue
    if (n.nodeType === Node.TEXT_NODE) {
      const s = t(n); if (s) out.push(para(s, bodySize, font))
      continue
    }
    if (n.nodeType !== Node.ELEMENT_NODE) continue
    const e = n as Element
    const tag = e.tagName.toLowerCase()

    if (tag === 'style' || tag === 'script' || tag === 'svg') continue

    // Mermaid pre — capture SVG for ImageRun conversion
    if (tag === 'pre' && e.classList.contains('mermaid')) {
      const svg = e.querySelector('svg')
      const text = t(e)
      if (svg) {
        _mermaidEntries.push({ svg: svg as SVGElement, fallbackText: text, paragraphIndex: out.length })
        out.push(new Paragraph({ children: [new TextRun({ text: '', size: 1 })], spacing: { after: 0 } }))
      } else if (text) {
        out.push(new Paragraph({ children: [new TextRun({ text: '[流程图] ' + text.slice(0, 120), italics: true, size: bodySize, color: '666666', font })], spacing: { after: 120 }, shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' } }))
      }
      continue
    }

    // Heading — may contain inline KaTeX
    if (H_SIZES[tag]) {
      const runs = buildRunsWithMath(e, { bold: true, size: H_SIZES[tag], color: '000000', font })
      out.push(new Paragraph({ children: runs, heading: H_LEVEL[tag], spacing: { before: 400 - parseInt(tag[1])*50, after: 160 } }))
      continue
    }

    // Paragraph — may contain inline KaTeX
    if (tag === 'p') {
      const runs = buildRunsWithMath(e, { size: bodySize, font })
      if (runs.length) {
        out.push(new Paragraph({ children: runs, spacing: { after: 120 } }))
      } else {
        const s = t(e); if (s) out.push(para(s, bodySize, font))
      }
      continue
    }

    // Code block — left-aligned, preserve indent, dark theme, no line numbers
    if (tag === 'pre') {
      const rawText = (e.textContent || '').replace(/^\n+|\n+$/g, '')
      if (rawText) {
        const lines = rawText.split('\n')
        const runs: TextRun[] = []
        lines.forEach((line, i) => {
          runs.push(new TextRun({
            text: line,
            font: 'Courier New',
            size: Math.max(16, bodySize - 4),
            color: 'd4d4d4',
            ...(i < lines.length - 1 ? { break: 1 as any } : {})
          }))
        })
        out.push(new Paragraph({ children: runs, spacing: { after: 120 }, shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' } }))
      }
      continue
    }

    // Lists
    if (tag === 'ul' || tag === 'ol') {
      e.querySelectorAll(':scope > li').forEach(li => {
        const runs = buildRunsWithMath(li as Element, { size: bodySize, font })
        if (runs.length) {
          out.push(new Paragraph({ children: runs, spacing: { after: 80 }, bullet: { level: 0 } }))
        } else {
          const s = t(li); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, size: bodySize, font })], spacing: { after: 80 }, bullet: { level: 0 } }))
        }
      })
      continue
    }

    // Blockquote
    if (tag === 'blockquote') {
      const runs = buildRunsWithMath(e, { italics: true, size: bodySize, font })
      if (runs.length) {
        out.push(new Paragraph({ children: runs, spacing: { after: 120 }, indent: { left: 480 }, border: { left: { style: 'single', size: 10, color: 'dddddd', space: 8 } } }))
      } else {
        const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, italics: true, size: bodySize, font })], spacing: { after: 120 }, indent: { left: 480 }, border: { left: { style: 'single', size: 10, color: 'dddddd', space: 8 } } }))
      }
      continue
    }

    // Table — cells may contain inline KaTeX
    if (tag === 'table') {
      const trs = e.querySelectorAll('tr')
      const cols = Math.max(...Array.from(trs).map(tr => tr.querySelectorAll('th,td').length), 1)
      const tdSize = Math.round(bodySize * 0.9)
      const rows: TableRow[] = []
      trs.forEach((tr, ri) => {
        const isHead = ri === 0 && !!tr.querySelector('th')
        rows.push(new TableRow({ children: Array.from(tr.querySelectorAll('th,td')).map(td => {
          const runs = buildRunsWithMath(td as Element, { bold: isHead, size: tdSize, font })
          return new TableCell({
            children: [new Paragraph({ children: runs.length ? runs : [new TextRun({ text: t(td), bold: isHead, size: tdSize, font })] })],
            width: { size: Math.floor(8000/cols), type: WidthType.DXA },
            shading: isHead ? { type: 'solid', color: 'f5f4f0', fill: 'f5f4f0' } : undefined
          })
        }) }))
      })
      out.push(new Table({ rows, width: { size: 8000, type: WidthType.DXA } }))
      continue
    }

    // HR
    if (tag === 'hr') { out.push(new Paragraph({ children: [new TextRun({ text: '—'.repeat(40), color: 'cccccc', size: 16 })], spacing: { before: 160, after: 160 }, alignment: AlignmentType.CENTER })); continue }

    // KaTeX display math (block) — convert to OMML
    if (isKatexDisplay(e)) {
      const latexInfo = extractLatex(e)
      if (latexInfo) {
        try {
          const omml = latexToOmml(latexInfo.latex, true)
          const marker = registerOmml(omml)
          out.push(new Paragraph({
            children: [new TextRun({ text: marker, size: bodySize, font })],
            spacing: { before: 80, after: 80 },
            alignment: AlignmentType.CENTER
          }))
        } catch {
          const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, size: bodySize, font })], spacing: { before: 80, after: 80 }, alignment: AlignmentType.CENTER }))
        }
      }
      continue
    }

    // Container — recurse
    if (tag === 'div' || tag === 'section') { out.push(...walk(e, cfg)); continue }

    // Skip KaTeX internals — already handled
    if (isInsideKatex(e)) continue

    // Fallback
    const s = t(e); if (s) out.push(para(s, bodySize, font))
  }
  return out
}

function para(text: string, size?: number, font?: string): Paragraph {
  return new Paragraph({ children: [new TextRun({ text, size: size || 22, font: font || 'Microsoft YaHei' })], spacing: { after: 120 } })
}

/**
 * Build docx runs from an element that may contain inline KaTeX formulas.
 * Non-math text becomes TextRun, inline formulas become OMML markers.
 */
function buildRunsWithMath(el: Element, baseStyle: any = {}): any[] {
  const runs: any[] = []

  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === Node.TEXT_NODE) {
      const text = (child.textContent || '')
      if (text.trim()) {
        runs.push(new TextRun({ ...baseStyle, text: text.trim() }))
      }
      continue
    }

    if (child.nodeType !== Node.ELEMENT_NODE) continue
    const childEl = child as Element

    // Check if this is a KaTeX inline formula
    if (childEl.classList.contains('katex')) {
      const latexInfo = extractLatex(childEl)
      if (latexInfo) {
        try {
          const omml = latexToOmml(latexInfo.latex, false)
          const marker = registerOmml(omml)
          runs.push(new TextRun({ ...baseStyle, text: marker }))
        } catch {
          runs.push(new TextRun({ ...baseStyle, text: t(childEl) }))
        }
      }
      continue
    }

    // Regular inline elements (strong, em, code, a, etc.)
    const tag = childEl.tagName.toLowerCase()
    if (tag === 'style' || tag === 'script') continue

    // Merge formatting
    const childStyle = { ...baseStyle }
    if (tag === 'strong' || tag === 'b') childStyle.bold = true
    if (tag === 'em' || tag === 'i') childStyle.italics = true
    if (tag === 'code') { childStyle.font = 'Courier New'; childStyle.size = 20; childStyle.color = 'c7254e' }

    // Recurse into nested inline elements
    const nestedRuns = buildRunsWithMath(childEl, childStyle)
    runs.push(...nestedRuns)
  }

  return runs
}
