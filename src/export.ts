import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx'
import { saveAs } from 'file-saver'

export function copyRichText(innerHtml: string, style: string): void {
  const doc = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>${style}</style></head>
<body><div class="doc">${innerHtml}</div></body></html>`

  const plain = extractText(innerHtml)

  const fallback = () => {
    const el = document.createElement('div')
    el.contentEditable = 'true'
    el.innerHTML = `<style>${style}</style><div class="doc">${innerHtml}</div>`
    el.style.position = 'fixed'; el.style.left = '-9999px'
    document.body.appendChild(el); el.focus()
    const sel = window.getSelection()!, range = document.createRange()
    range.selectNodeContents(el); sel.removeAllRanges(); sel.addRange(range)
    document.execCommand('copy')
    sel.removeAllRanges(); document.body.removeChild(el)
  }

  try {
    const htmlBlob = new Blob([doc], { type: 'text/html' })
    const plainBlob = new Blob([plain], { type: 'text/plain' })
    navigator.clipboard.write([new ClipboardItem({
      'text/html': htmlBlob,
      'text/plain': plainBlob,
    })]).catch(fallback)
  } catch {
    fallback()
  }
}

function extractText(html: string): string {
  const d = document.createElement('div')
  d.innerHTML = html
  return (d.textContent || '').trim()
}

// ---- DOCX ----

const H_SIZES = [0,44,36,30,26,22,20]
const H_MAP  = [HeadingLevel.HEADING_1,HeadingLevel.HEADING_1,HeadingLevel.HEADING_2,HeadingLevel.HEADING_3,HeadingLevel.HEADING_4,HeadingLevel.HEADING_5,HeadingLevel.HEADING_6]

function textOf(el: Node): string { return el.textContent || '' }
function isOnlyWhitespace(n: Node): boolean {
  return n.nodeType === Node.TEXT_NODE && !(n.textContent || '').trim()
}

export async function downloadDocx(docEl: Element, _style: string, filename: string): Promise<void> {
  const children = buildFlat(docEl)
  const doc = new Document({
    sections: [{ properties: {}, children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })] }],
  })
  Packer.toBlob(doc).then(blob => saveAs(blob, filename))
}

/** Walk DIRECT children of root only (not recursively descend into p→span etc) */
function buildFlat(root: Element): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = []

  for (const n of Array.from(root.childNodes)) {
    if (isOnlyWhitespace(n)) continue

    if (n.nodeType === Node.TEXT_NODE) {
      const t = (n.textContent || '').trim()
      if (t) out.push(p(t))
      continue
    }

    if (n.nodeType !== Node.ELEMENT_NODE) continue
    const el = n as Element
    const tag = el.tagName.toLowerCase()

    // Skip style/script
    if (tag === 'style' || tag === 'script') continue
    // Skip mermaid SVG (already rendered in DOM, not useful for docx)
    if (tag === 'svg') continue

    // Mermaid pre block
    if (tag === 'pre' && el.classList.contains('mermaid')) {
      const t = textOf(el).trim()
      if (t) out.push(new Paragraph({
        children: [new TextRun({ text: '[流程图] ' + t.slice(0, 120), italics: true, size: 20, color: '666666' })],
        spacing: { after: 120 },
        shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' },
      }))
      continue
    }

    // Headings
    const hm = tag.match(/^h([1-6])$/)
    if (hm) {
      const lv = parseInt(hm[1])
      out.push(new Paragraph({
        children: [new TextRun({ text: textOf(el), bold: true, size: H_SIZES[lv] })],
        heading: H_MAP[lv],
        spacing: { before: 400 - lv*50, after: 160 },
      }))
      continue
    }

    // Paragraph — extract text, skip inline markup
    if (tag === 'p') {
      const t = textOf(el).trim()
      if (t) out.push(p(t))
      continue
    }

    // Code block
    if (tag === 'pre') {
      const code = el.querySelector('code')
      const t = textOf(code || el).trim()
      if (t) out.push(new Paragraph({
        children: [new TextRun({ text: t, font: 'Courier New', size: 18 })],
        spacing: { after: 120 },
        shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' },
      }))
      continue
    }

    // Lists
    if (tag === 'ul' || tag === 'ol') {
      el.querySelectorAll(':scope > li').forEach(li => {
        const t = textOf(li).trim()
        if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 80 }, bullet: { level: 0 } }))
      })
      continue
    }

    // Blockquote
    if (tag === 'blockquote') {
      const t = textOf(el).trim()
      if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, italics: true, size: 22 })], spacing: { after: 120 }, indent: { left: 480 },
        border: { left: { style: 'single', size: 10, color: 'ddd', space: 8 } } }))
      continue
    }

    // Table
    if (tag === 'table') {
      const trs = el.querySelectorAll('tr')
      const colCount = Math.max(...Array.from(trs).map(tr => tr.querySelectorAll('th,td').length), 1)
      const rows: TableRow[] = []
      trs.forEach((tr, ri) => {
        const isHead = ri === 0 && !!tr.querySelector('th')
        const cells = Array.from(tr.querySelectorAll('th,td')).map(td => new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: textOf(td), bold: isHead, size: 20 })] })],
          width: { size: Math.floor(8000 / colCount), type: WidthType.DXA },
          shading: isHead ? { type: 'solid', color: 'f5f4f0', fill: 'f5f4f0' } : undefined,
        }))
        rows.push(new TableRow({ children: cells }))
      })
      out.push(new Table({ rows, width: { size: 8000, type: WidthType.DXA } }))
      continue
    }

    // HR
    if (tag === 'hr') {
      out.push(new Paragraph({ children: [new TextRun({ text: '—'.repeat(40), color: 'cccccc', size: 16 })], spacing: { before: 160, after: 160 }, alignment: 'center' }))
      continue
    }

    // KaTeX block math
    if (el.classList.contains('katex-display') || el.classList.contains('math-block')) {
      const t = textOf(el).trim()
      if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { before: 80, after: 80 }, alignment: 'center' }))
      continue
    }

    // Nested div — recurse
    if (tag === 'div' || tag === 'section') {
      out.push(...buildFlat(el))
      continue
    }

    // KaTeX inline — skip (content already extracted by parent <p>)
    if (el.closest('.katex')) continue
    if (el.classList.contains('math-inline')) {
      const t = textOf(el).trim()
      if (t) out.push(p(t))
      continue
    }

    // Fallback
    const t = textOf(el).trim()
    if (t) out.push(p(t))
  }

  return out
}

function p(text: string): Paragraph {
  return new Paragraph({ children: [new TextRun({ text, size: 22 })], spacing: { after: 120 } })
}
