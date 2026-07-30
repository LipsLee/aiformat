import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } from 'docx'
import { saveAs } from 'file-saver'
import { latexToOmml, extractLatex, isKatexDisplay, isInsideKatex } from './latex2omml'

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

  // === Build HTML: replace KaTeX with OMML markers, clean up Mermaid ===
  // Remove KaTeX annotation elements (they contain raw LaTeX source)
  clone.querySelectorAll('.katex annotation').forEach(a => a.remove())

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
  // Remove annotations (raw LaTeX source)
  plainClone.querySelectorAll('.katex annotation').forEach(a => a.remove())
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

const H_SIZES: Record<string,number> = { h1:44, h2:36, h3:30, h4:26, h5:22, h6:20 }
const H_LEVEL: Record<string,any> = {
  h1: HeadingLevel.HEADING_1, h2: HeadingLevel.HEADING_2,
  h3: HeadingLevel.HEADING_3, h4: HeadingLevel.HEADING_4,
  h5: HeadingLevel.HEADING_5, h6: HeadingLevel.HEADING_6,
}

function t(n: Node): string { return (n.textContent || '').trim() }

// OMML placeholder marker — we insert this into the docx as text, then
// post-process the zip to replace it with real OMML XML.
const OMML_MARKER_PREFIX = '\uFEFF__OMML_'
const ommlRegistry: string[] = []

function registerOmml(omml: string): string {
  const id = ommlRegistry.length
  ommlRegistry.push(omml)
  return `${OMML_MARKER_PREFIX}${id}__`
}

export async function downloadDocx(docEl: Element, _style: string, filename: string): Promise<void> {
  // Reset registry
  ommlRegistry.length = 0

  const children = walk(docEl)
  const doc = new Document({
    sections: [{ properties: {}, children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })] }],
  })

  const blob = await Packer.toBlob(doc)

  // Post-process: replace OMML markers with actual OMML XML
  const finalBlob = await replaceOmmlMarkers(blob)

  saveAs(finalBlob, filename)
}

async function replaceOmmlMarkers(blob: Blob): Promise<Blob> {
  if (ommlRegistry.length === 0) return blob

  const JSZip = (await import('jszip')).default
  const zip = await JSZip.loadAsync(blob)

  let docXml = await zip.file('word/document.xml')!.async('string')

  // Ensure math namespace is declared
  if (!docXml.includes('xmlns:m=')) {
    docXml = docXml.replace(
      /<w:document\s/,
      '<w:document xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    )
  }

  // Replace each marker with OMML XML
  for (let i = 0; i < ommlRegistry.length; i++) {
    const marker = `${OMML_MARKER_PREFIX}${i}__`
    const omml = ommlRegistry[i]

    const escapedMarker = marker.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')
    const runRegex = new RegExp(
      `<w:r[^>]*>(?:<w:rPr[^>]*(?:/>|>.*?</w:rPr>))?<w:t[^>]*>${escapedMarker}</w:t></w:r>`,
      'gs'
    )
    docXml = docXml.replace(runRegex, omml)
  }

  zip.file('word/document.xml', docXml)

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE'
  })
}

function walk(el: Element): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = []
  for (const n of Array.from(el.childNodes)) {
    if (n.nodeType === Node.COMMENT_NODE) continue
    if (n.nodeType === Node.TEXT_NODE) {
      const s = t(n); if (s) out.push(para(s))
      continue
    }
    if (n.nodeType !== Node.ELEMENT_NODE) continue
    const e = n as Element
    const tag = e.tagName.toLowerCase()

    if (tag === 'style' || tag === 'script' || tag === 'svg') continue

    // Mermaid pre
    if (tag === 'pre' && e.classList.contains('mermaid')) {
      const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: '[流程图] ' + s.slice(0, 120), italics: true, size: 20, color: '666666' })], spacing: { after: 120 }, shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' } }))
      continue
    }

    // Heading — may contain inline KaTeX
    if (H_SIZES[tag]) {
      const runs = buildRunsWithMath(e, { bold: true, size: H_SIZES[tag] })
      out.push(new Paragraph({ children: runs, heading: H_LEVEL[tag], spacing: { before: 400 - parseInt(tag[1])*50, after: 160 } }))
      continue
    }

    // Paragraph — may contain inline KaTeX
    if (tag === 'p') {
      const runs = buildRunsWithMath(e, { size: 22 })
      if (runs.length) {
        out.push(new Paragraph({ children: runs, spacing: { after: 120 } }))
      } else {
        const s = t(e); if (s) out.push(para(s))
      }
      continue
    }

    // Code block
    if (tag === 'pre') { const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, font: 'Courier New', size: 18 })], spacing: { after: 120 }, shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' } })); continue }

    // Lists
    if (tag === 'ul' || tag === 'ol') {
      e.querySelectorAll(':scope > li').forEach(li => {
        const runs = buildRunsWithMath(li as Element, { size: 22 })
        if (runs.length) {
          out.push(new Paragraph({ children: runs, spacing: { after: 80 }, bullet: { level: 0 } }))
        } else {
          const s = t(li); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, size: 22 })], spacing: { after: 80 }, bullet: { level: 0 } }))
        }
      })
      continue
    }

    // Blockquote
    if (tag === 'blockquote') {
      const runs = buildRunsWithMath(e, { italics: true, size: 22 })
      if (runs.length) {
        out.push(new Paragraph({ children: runs, spacing: { after: 120 }, indent: { left: 480 }, border: { left: { style: 'single', size: 10, color: 'ddd', space: 8 } } }))
      } else {
        const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, italics: true, size: 22 })], spacing: { after: 120 }, indent: { left: 480 }, border: { left: { style: 'single', size: 10, color: 'ddd', space: 8 } } }))
      }
      continue
    }

    // Table — cells may contain inline KaTeX
    if (tag === 'table') {
      const trs = e.querySelectorAll('tr')
      const cols = Math.max(...Array.from(trs).map(tr => tr.querySelectorAll('th,td').length), 1)
      const rows: TableRow[] = []
      trs.forEach((tr, ri) => {
        const isHead = ri === 0 && !!tr.querySelector('th')
        rows.push(new TableRow({ children: Array.from(tr.querySelectorAll('th,td')).map(td => {
          const runs = buildRunsWithMath(td as Element, { bold: isHead, size: 20 })
          return new TableCell({
            children: [new Paragraph({ children: runs.length ? runs : [new TextRun({ text: t(td), bold: isHead, size: 20 })] })],
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
            children: [new TextRun({ text: marker, size: 22 })],
            spacing: { before: 80, after: 80 },
            alignment: AlignmentType.CENTER
          }))
        } catch {
          const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, size: 22 })], spacing: { before: 80, after: 80 }, alignment: AlignmentType.CENTER }))
        }
      }
      continue
    }

    // Container — recurse
    if (tag === 'div' || tag === 'section') { out.push(...walk(e)); continue }

    // Skip KaTeX internals — already handled
    if (isInsideKatex(e)) continue

    // Fallback
    const s = t(e); if (s) out.push(para(s))
  }
  return out
}

function para(text: string): Paragraph {
  return new Paragraph({ children: [new TextRun({ text, size: 22 })], spacing: { after: 120 } })
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
