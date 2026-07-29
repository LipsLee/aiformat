import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx'
import { saveAs } from 'file-saver'

// ---- COPY support ----
function getKatexCssTag(): string {
  const link = document.querySelector('link[href*="katex.min.css"]') as HTMLLinkElement | null
  return link ? link.outerHTML : ''
}

export function copyRichText(innerHtml: string, style: string): void {
  // Build complete HTML document WITH KaTeX CSS so Word/Google Docs can render formulas
  const katexCss = getKatexCssTag()
  const docHtml = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">${katexCss}<style>${style}</style></head>
<body><div class="doc">${innerHtml}</div></body></html>`

  // Extract plain text (KaTeX textContent gives readable Unicode math symbols)
  const d = document.createElement('div')
  d.innerHTML = innerHtml
  const plain = (d.textContent || '').trim()

  // Primary: Clipboard API with HTML + plain text
  try {
    const htmlBlob = new Blob([docHtml], { type: 'text/html' })
    const plainBlob = new Blob([plain], { type: 'text/plain' })
    navigator.clipboard.write([
      new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': plainBlob }),
    ])
    return
  } catch (_) { /* fallthrough to backup */ }

  // Backup: contentEditable + execCommand (for browsers without ClipboardItem)
  const el = document.createElement('div')
  el.contentEditable = 'true'
  el.innerHTML = `<style>${style}</style><div class="doc">${innerHtml}</div>`
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

export async function downloadDocx(docEl: Element, _style: string, filename: string): Promise<void> {
  const children = walk(docEl)
  const doc = new Document({
    sections: [{ properties: {}, children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })] }],
  })
  Packer.toBlob(doc).then(blob => saveAs(blob, filename))
}

// Walk ONLY direct children of root. Recurse only into <div>/<section>/<span> containers.
// For <p><span class="katex">...</span></p>, extract text from <p> via textContent.
// This avoids descending into KaTeX internals which would produce garbled output.
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

    // Heading
    if (H_SIZES[tag]) {
      out.push(new Paragraph({ children: [new TextRun({ text: t(e), bold: true, size: H_SIZES[tag] })], heading: H_LEVEL[tag], spacing: { before: 400 - parseInt(tag[1])*50, after: 160 } }))
      continue
    }

    // Paragraph
    if (tag === 'p') { const s = t(e); if (s) out.push(para(s)); continue }

    // Code block
    if (tag === 'pre') { const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, font: 'Courier New', size: 18 })], spacing: { after: 120 }, shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' } })); continue }

    // Lists
    if (tag === 'ul' || tag === 'ol') {
      e.querySelectorAll(':scope > li').forEach(li => { const s = t(li); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, size: 22 })], spacing: { after: 80 }, bullet: { level: 0 } })) })
      continue
    }

    // Blockquote
    if (tag === 'blockquote') { const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, italics: true, size: 22 })], spacing: { after: 120 }, indent: { left: 480 }, border: { left: { style: 'single', size: 10, color: 'ddd', space: 8 } } })); continue }

    // Table
    if (tag === 'table') {
      const trs = e.querySelectorAll('tr')
      const cols = Math.max(...Array.from(trs).map(tr => tr.querySelectorAll('th,td').length), 1)
      const rows: TableRow[] = []
      trs.forEach((tr, ri) => {
        const isHead = ri === 0 && !!tr.querySelector('th')
        rows.push(new TableRow({ children: Array.from(tr.querySelectorAll('th,td')).map(td => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t(td), bold: isHead, size: 20 })] })], width: { size: Math.floor(8000/cols), type: WidthType.DXA }, shading: isHead ? { type: 'solid', color: 'f5f4f0', fill: 'f5f4f0' } : undefined })) }))
      })
      out.push(new Table({ rows, width: { size: 8000, type: WidthType.DXA } }))
      continue
    }

    // HR
    if (tag === 'hr') { out.push(new Paragraph({ children: [new TextRun({ text: '—'.repeat(40), color: 'cccccc', size: 16 })], spacing: { before: 160, after: 160 }, alignment: 'center' })); continue }

    // KaTeX display math (block)
    if (e.classList.contains('katex-display') || e.classList.contains('math-block')) { const s = t(e); if (s) out.push(new Paragraph({ children: [new TextRun({ text: s, size: 22 })], spacing: { before: 80, after: 80 }, alignment: 'center' })); continue }

    // Container — recurse
    if (tag === 'div' || tag === 'section') { out.push(...walk(e)); continue }

    // KaTeX inline math — skip (text already in parent <p>)
    if (e.closest('.katex')) continue

    // Fallback
    const s = t(e); if (s) out.push(para(s))
  }
  return out
}

function para(text: string): Paragraph {
  return new Paragraph({ children: [new TextRun({ text, size: 22 })], spacing: { after: 120 } })
}
