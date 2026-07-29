import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } from 'docx'
import { saveAs } from 'file-saver'

export function copyRichText(fullHtml: string, style: string): void {
  const fallback = () => {
    const el = document.createElement('div')
    el.contentEditable = 'true'
    el.innerHTML = `<style>${style}</style><div class="doc">${fullHtml.replace(/<style[^>]*>[\s\S]*?<\/style>/gi,'').replace(/^[\s\S]*?<div class="doc">/, '').replace(/<\/div>\s*$/,'')}</div>`
    el.style.position = 'fixed'; el.style.left = '-9999px'
    document.body.appendChild(el); el.focus()
    const sel = window.getSelection()!, range = document.createRange()
    range.selectNodeContents(el); sel.removeAllRanges(); sel.addRange(range)
    document.execCommand('copy')
    sel.removeAllRanges(); document.body.removeChild(el)
  }

  try {
    const div = document.createElement('div')
    div.innerHTML = fullHtml
    const plain = (div.textContent || '').trim()
    const blob = new Blob([fullHtml], { type: 'text/html' })
    navigator.clipboard.write([new ClipboardItem({
      'text/html': blob,
      'text/plain': new Blob([plain], { type: 'text/plain' }),
    })]).catch(fallback)
  } catch { fallback() }
}

// ---- DOCX: real DOM walker ----

function txt(n: Node): string { return n.textContent || '' }

function buildDocx(root: Element): (Paragraph | Table)[] {
  const out: (Paragraph | Table)[] = []

  const hSizes = [0,44,36,30,26,22,20]
  const hMap = [HeadingLevel.HEADING_1,HeadingLevel.HEADING_1,HeadingLevel.HEADING_2,HeadingLevel.HEADING_3,HeadingLevel.HEADING_4,HeadingLevel.HEADING_5,HeadingLevel.HEADING_6]

  function process(el: Element) {
    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType === Node.TEXT_NODE) {
        const t = txt(child).trim()
        if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 120 } }))
        continue
      }
      if (child.nodeType !== Node.ELEMENT_NODE) continue

      const c = child as Element
      const tag = c.tagName.toLowerCase()

      // Skip style/script
      if (tag === 'style' || tag === 'script') continue

      // Mermaid rendered SVG — skip, can't embed SVG in DOCX easily
      if (tag === 'svg' && (c.closest('pre.mermaid') || c.parentElement?.classList.contains('mermaid'))) continue
      if (tag === 'pre' && c.classList.contains('mermaid')) {
        const t = txt(c).trim()
        if (t) out.push(new Paragraph({
          children: [new TextRun({ text: '[流程图] ' + t.slice(0, 120), italics: true, size: 20, color: '666666' })],
          spacing: { after: 120 },
          shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' },
        }))
        continue
      }

      // KaTeX rendered math — strip to plain text
      if (c.classList.contains('katex') || c.closest('.katex-display')) {
        continue // handled by parent block element
      }
      if (c.classList.contains('katex-display') || (tag === 'div' && c.classList.contains('math-block'))) {
        const t = txt(c).trim()
        if (t) out.push(new Paragraph({
          children: [new TextRun({ text: t, size: 22 })],
          spacing: { before: 80, after: 80 }, alignment: 'center',
        }))
        continue
      }

      // Headings
      const hm = tag.match(/^h([1-6])$/)
      if (hm) {
        const lv = parseInt(hm[1])
        out.push(new Paragraph({
          children: [new TextRun({ text: txt(c), bold: true, size: hSizes[lv] })],
          heading: hMap[lv],
          spacing: { before: 400 - lv*50, after: 160 },
        }))
        continue
      }

      // Paragraph
      if (tag === 'p') {
        const t = txt(c).trim()
        if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 120 } }))
        continue
      }

      // Code block
      if (tag === 'pre') {
        const t = txt(c).trim()
        if (t) out.push(new Paragraph({
          children: [new TextRun({ text: t, font: 'Courier New', size: 18 })],
          spacing: { after: 120 },
          shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' },
        }))
        continue
      }

      // Lists
      if (tag === 'ul' || tag === 'ol') {
        c.querySelectorAll(':scope > li').forEach(li => {
          const t = txt(li).trim()
          if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 80 }, bullet: { level: 0 } }))
        })
        continue
      }

      // Blockquote
      if (tag === 'blockquote') {
        const t = txt(c).trim()
        if (t) out.push(new Paragraph({
          children: [new TextRun({ text: t, italics: true, size: 22 })],
          spacing: { after: 120 }, indent: { left: 480 },
          border: { left: { style: 'single', size: 10, color: 'ddd', space: 8 } },
        }))
        continue
      }

      // Table
      if (tag === 'table') {
        const trs = c.querySelectorAll('tr')
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

      // HR
      if (tag === 'hr') {
        out.push(new Paragraph({
          children: [new TextRun({ text: '—'.repeat(40), color: 'cccccc', size: 16 })],
          spacing: { before: 160, after: 160 }, alignment: 'center',
        }))
        continue
      }

      // KaTeX inline math span
      if (c.classList.contains('math-inline') || c.querySelector('.katex')) {
        const t = txt(c).trim()
        if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 80 } }))
        continue
      }

      // Div / section / article — recurse
      if (tag === 'div' || tag === 'section' || tag === 'article' || tag === 'span') {
        process(c)
        continue
      }

      // Fallback
      const t = txt(c).trim()
      if (t) out.push(new Paragraph({ children: [new TextRun({ text: t, size: 22 })], spacing: { after: 120 } }))
    }
  }

  process(root)
  return out
}

export async function downloadDocx(docEl: Element, _style: string, filename: string): Promise<void> {
  const children = buildDocx(docEl)
  const doc = new Document({
    sections: [{
      properties: {},
      children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })],
    }],
  })
  Packer.toBlob(doc).then(blob => saveAs(blob, filename))
}
