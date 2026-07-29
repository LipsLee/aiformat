import { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ImageRun } from 'docx'
import { saveAs } from 'file-saver'
import { makeStyles } from './parser'
import { getStyleConfig } from './style'

export function copyRichText(html: string): void {
  const config = getStyleConfig()
  const clean = html.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
  const fullHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${makeStyles(config)}</style></head><body>${clean}</body></html>`

  function fallback() {
    const el = document.createElement('div')
    el.contentEditable = 'true'
    el.innerHTML = `<style>${makeStyles(config)}</style>${clean}`
    el.style.position = 'fixed'; el.style.left = '-9999px'
    document.body.appendChild(el); el.focus()
    const sel = window.getSelection()!, range = document.createRange()
    range.selectNodeContents(el); sel.removeAllRanges(); sel.addRange(range)
    document.execCommand('copy')
    sel.removeAllRanges(); document.body.removeChild(el)
  }

  try {
    const blob = new Blob([fullHtml], { type: 'text/html' })
    const div = document.createElement('div')
    div.innerHTML = clean
    const plain = (div.textContent || '').trim()
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

// ---- Mermaid → Image ----
async function renderMermaidImage(code: string): Promise<Uint8Array | null> {
  try {
    const mermaid = (window as any).mermaid
    if (!mermaid) return null
    const id = 'mdocx-' + Math.random().toString(36).slice(2, 8)
    const { svg } = await mermaid.render(id, code)

    return new Promise(resolve => {
      const img = new Image()
      img.onload = () => {
        const c = document.createElement('canvas')
        c.width = img.width * 2; c.height = img.height * 2
        const ctx = c.getContext('2d')!
        ctx.scale(2, 2); ctx.drawImage(img, 0, 0)
        c.toBlob(blob => {
          if (blob) blob.arrayBuffer().then(buf => resolve(new Uint8Array(buf)))
          else resolve(null)
        }, 'image/png')
      }
      img.onerror = () => resolve(null)
      img.src = 'data:image/svg+xml;base64,' + btoa(String.fromCharCode(...new TextEncoder().encode(svg)))
    })
  } catch {
    return null
  }
}

// ---- DOCX ----
function txt(el: Element): string { return el.textContent || '' }

async function buildDocx(root: Element): Promise<(Paragraph | Table)[]> {
  const out: (Paragraph | Table)[] = []
  for (const el of Array.from(root.children)) {
    const tag = el.tagName.toLowerCase()
    if (tag === 'style') continue

    if (tag === 'pre' && el.classList.contains('mermaid')) {
      const code = txt(el)
      const imgData = await renderMermaidImage(code)
      if (imgData) {
        out.push(new Paragraph({
          children: [new ImageRun({ data: imgData, transformation: { width: 500, height: 300 }, type: 'png' })],
          spacing: { before: 120, after: 120 }, alignment: 'center',
        }))
      } else {
        out.push(new Paragraph({
          children: [new TextRun({ text: '[流程图] ' + code.slice(0, 120), italics: true, size: 20, color: '666666' })],
          spacing: { after: 120 },
          shading: { type: 'solid', color: 'fafbfc', fill: 'fafbfc' },
        }))
      }
      continue
    }

    if (tag === 'div' && (el.classList.contains('katex-display') || el.classList.contains('math-block'))) {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(el).slice(0, 200), font: 'Courier New', size: 20 })],
        spacing: { before: 120, after: 120 }, alignment: 'center',
      }))
      continue
    }

    if (/^h[1-6]$/.test(tag)) {
      const lv = parseInt(tag[1]), sizes = [0,44,36,30,26,22,20]
      const hMap = [HeadingLevel.HEADING_1,HeadingLevel.HEADING_1,HeadingLevel.HEADING_2,HeadingLevel.HEADING_3,HeadingLevel.HEADING_4,HeadingLevel.HEADING_5,HeadingLevel.HEADING_6]
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(el), bold: true, size: sizes[lv] })],
        heading: hMap[lv],
        spacing: { before: 400 - lv*50, after: 160 },
      }))
      continue
    }
    if (tag === 'p') {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(el), size: 22 })],
        spacing: { after: 120 },
      }))
      continue
    }
    if (tag === 'pre') {
      const code = el.querySelector('code')
      out.push(new Paragraph({
        children: [new TextRun({ text: code ? txt(code) : txt(el), font: 'Courier New', size: 18 })],
        spacing: { after: 120 },
        shading: { type: 'solid', color: '1e1e1e', fill: '1e1e1e' },
      }))
      continue
    }
    if (tag === 'ul' || tag === 'ol') {
      el.querySelectorAll('li').forEach(li => {
        out.push(new Paragraph({
          children: [new TextRun({ text: txt(li), size: 22 })],
          spacing: { after: 80 }, bullet: { level: 0 },
        }))
      })
      continue
    }
    if (tag === 'blockquote') {
      out.push(new Paragraph({
        children: [new TextRun({ text: txt(el), italics: true, size: 22 })],
        spacing: { after: 120 }, indent: { left: 480 },
        border: { left: { style: 'single', size: 10, color: '2563eb', space: 8 } },
      }))
      continue
    }
    if (tag === 'table') {
      const trs = el.querySelectorAll('tr')
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
    if (tag === 'hr') {
      out.push(new Paragraph({
        children: [new TextRun({ text: '—'.repeat(40), color: 'cccccc', size: 16 })],
        spacing: { before: 160, after: 160 }, alignment: 'center',
      }))
      continue
    }
    if (tag === 'div') out.push(...(await buildDocx(el)))
  }
  return out
}

export async function downloadDocx(html: string, filename: string): Promise<void> {
  const div = document.createElement('div')
  div.innerHTML = html
  const root = div.querySelector('.doc') || div
  const children = await buildDocx(root)
  const doc = new Document({
    sections: [{
      properties: {},
      children: children.length ? children : [new Paragraph({ children: [new TextRun('')] })],
    }],
  })
  Packer.toBlob(doc).then(blob => saveAs(blob, filename))
}
