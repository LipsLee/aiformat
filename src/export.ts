import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx'
import { saveAs } from 'file-saver'

export function copyRichText(innerHtml: string, style: string): void {
  const fullHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${style}</style></head><body><div class="doc">${innerHtml}</div></body></html>`

  function fallback() {
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
    const div = document.createElement('div')
    div.innerHTML = innerHtml
    const plain = (div.textContent || '').trim()
    const blob = new Blob([fullHtml], { type: 'text/html' })
    navigator.clipboard.write([new ClipboardItem({
      'text/html': blob,
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    })]).catch(fallback)
  } catch { fallback() }
}

export function copyPlainText(text: string): void {
  navigator.clipboard.writeText(text).catch(() => {
    const ta = document.createElement('textarea')
    ta.value = text; ta.style.position = 'fixed'; ta.style.left = '-9999px'
    document.body.appendChild(ta); ta.select(); document.execCommand('copy')
    document.body.removeChild(ta)
  })
}

// ---- DOCX ----
function txt(el: Node): string { return el.textContent || '' }

function walkDom(root: Element): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = []

  for (const el of Array.from(root.childNodes)) {
    if (el.nodeType === Node.TEXT_NODE) {
      const t = txt(el).trim()
      if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 120 } }))
      continue
    }
    if (el.nodeType !== Node.ELEMENT_NODE) continue

    const elem = el as Element
    const tag = elem.tagName.toLowerCase()

    // Skip style/script
    if (tag === 'style' || tag === 'script') continue

    // Skip KaTeX-display wrapper, handled inside
    if (elem.classList.contains('katex-display')) {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(elem), size: 22 })],
        spacing: { before: 120, after: 120 }, alignment: 'center',
      }))
      continue
    }

    // Mermaid (rendered as SVG in DOM)
    if (tag === 'pre' && elem.classList.contains('mermaid')) {
      out.push(new Paragraph({
        children: [new TextRun({ text: '[流程图] ' + txt(elem).slice(0, 120), italics: true, size: 20, color: '666666' })],
        spacing: { after: 120 },
        shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' },
      }))
      continue
    }

    // Headings
    if (/^h[1-6]$/.test(tag)) {
      const lv = parseInt(tag[1]), sizes = [0,44,36,30,26,22,20]
      const hMap = [HeadingLevel.HEADING_1,HeadingLevel.HEADING_1,HeadingLevel.HEADING_2,HeadingLevel.HEADING_3,HeadingLevel.HEADING_4,HeadingLevel.HEADING_5,HeadingLevel.HEADING_6]
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(elem), bold: true, size: sizes[lv] })],
        heading: hMap[lv],
        spacing: { before: 400 - lv*50, after: 160 },
      }))
      continue
    }

    // Paragraph
    if (tag === 'p') {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(elem), size: 22 })],
        spacing: { after: 120 },
      }))
      continue
    }

    // Code block
    if (tag === 'pre') {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(elem), font: 'Courier New', size: 18 })],
        spacing: { after: 120 },
        shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' },
      }))
      continue
    }

    // Lists
    if (tag === 'ul' || tag === 'ol') {
      elem.querySelectorAll('li').forEach(li => {
        out.push(new Paragraph({
          children: [new TextRun({ text: txt(li), size: 22 })],
          spacing: { after: 80 }, bullet: { level: 0 },
        }))
      })
      continue
    }

    // Blockquote
    if (tag === 'blockquote') {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(elem), italics: true, size: 22 })],
        spacing: { after: 120 }, indent: { left: 480 },
        border: { left: { style: 'single', size: 10, color: '2563eb', space: 8 } },
      }))
      continue
    }

    // Table
    if (tag === 'table') {
      const trs = elem.querySelectorAll('tr')
      const colCount = Math.max(...Array.from(trs).map(tr => tr.querySelectorAll('th,td').length), 1)
      const colWidth = Math.floor(8000 / colCount)
      const rows: TableRow[] = []
      trs.forEach((tr, ri) => {
        const isHead = ri === 0 && !!tr.querySelector('th')
        const cells = Array.from(tr.querySelectorAll('th,td')).map(td => new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: txt(td), bold: isHead, size: 20 })] })],
          width: { size: colWidth, type: WidthType.DXA },
          shading: isHead ? { type: 'solid', color: 'f5f4f0', fill: 'f5f4f0' } : undefined,
        }))
        rows.push(new TableRow({ children: cells }))
      })
      out.push(new Table({ rows, width: { size: 8000, type: WidthType.DXA } }))
      continue
    }

    // Horizontal rule
    if (tag === 'hr') {
      out.push(new Paragraph({
        children: [new TextRun({ text: '—'.repeat(40), color: 'cccccc', size: 16 })],
        spacing: { before: 160, after: 160 }, alignment: 'center',
      }))
      continue
    }

    // Div or other container — recurse
    if (tag === 'div' || tag === 'section' || tag === 'article') {
      out.push(...walkDom(elem))
      continue
    }

    // KaTeX inline math span — extract text
    if (elem.classList.contains('katex')) {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(elem), size: 22 })],
        spacing: { after: 80 },
      }))
      continue
    }

    // Fallback: treat as paragraph
    const t = txt(elem).trim()
    if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 120 } }))
  }

  return out
}

export async function downloadDocx(docEl: Element, _style: string, filename: string): Promise<void> {
  const children = walkDom(docEl)
  const doc = new Document({
    sections: [{
      properties: {},
      children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })],
    }],
  })
  Packer.toBlob(doc).then(blob => saveAs(blob, filename))
}
