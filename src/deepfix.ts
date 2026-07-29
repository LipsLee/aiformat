// Deep correction — walk DOM text nodes, find mangled or unrendered LaTeX AND markdown, re-render
import { Marked } from 'marked'

const md = new Marked({ gfm: true, breaks: false })

interface TextPatch {
  node: Text
  parts: (string | Element)[]
}

export async function deepCorrect(root: Element): Promise<number> {
  const patches = collectPatches(root)
  if (patches.length === 0) return 0

  for (const patch of patches) {
    const parent = patch.node.parentNode
    if (!parent) continue
    for (const part of patch.parts) {
      parent.insertBefore(
        typeof part === 'string' ? document.createTextNode(part) : part,
        patch.node
      )
    }
    parent.removeChild(patch.node)
  }

  // Re-run KaTeX auto-render on the updated DOM
  const w = window as any
  if (w.renderMathInElement) {
    try {
      w.renderMathInElement(root, {
        delimiters: [
          { left: '$$', right: '$$', display: true },
          { left: '$', right: '$', display: false },
          { left: '\\(', right: '\\)', display: false },
          { left: '\\[', right: '\\]', display: true },
        ],
        throwOnError: false,
      })
    } catch (e) { /* ignore */ }
  }

  return patches.length
}

function collectPatches(root: Element): TextPatch[] {
  const patches: TextPatch[] = []
  walk(root, patches)
  return patches
}

function walk(node: Node, patches: TextPatch[]) {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent || ''
    if (hasIssues(text)) {
      const parts = processText(text)
      if (parts && parts.some(p => typeof p !== 'string')) {
        patches.push({ node: node as Text, parts })
      }
    }
    return
  }

  if (node.nodeType !== Node.ELEMENT_NODE) return
  const el = node as Element
  const tag = el.tagName.toLowerCase()
  if (
    tag === 'script' || tag === 'style' ||
    el.closest('.katex') || el.closest('pre') || el.closest('code') ||
    el.closest('.mermaid') || el.closest('[data-deep-fixed]')
  ) return

  for (const child of Array.from(el.childNodes)) {
    walk(child, patches)
  }
}

// ========== DETECTION ==========

const LATEX_SIGNAL = /\\frac|\\sum|\\int|\\prod|\\sqrt|\\alpha|\\beta|\\gamma|\\delta|\\epsilon|\\theta|\\lambda|\\mu|\\pi|\\sigma|\\omega|\\infty|\\partial|\\nabla|\\times|\\div|\\pm|\\cdot|\\leq|\\geq|\\neq|\\approx|\\equiv|\\rightarrow|\\Rightarrow|\\leftarrow|\\Leftarrow|\\subset|\\supset|\\in|\\notin|\\forall|\\exists|\\mathbb|\\mathbf|\\mathcal|\\text|\\begin|\\end|\\lim|\\log|\\ln|\\sin|\\cos|\\tan|\\det|\\max|\\min|\\bigg|\\Bigg|\\big|\\Big|\\hbar|\\hat|\\dot|\\ddot|\\widehat|\\overrightarrow|\\overleftarrow|\\mapsto|\\longrightarrow|\\longleftarrow|\\left|\\right|\\middle|\\langle|\\rangle|\\lVert|\\rVert|\\binom/

const MD_SIGNAL = /\*\*[^*]+\*\*|\*[^*]+\*|~~[^~]+~~|`{1,2}[^`]+`{1,2}|\[[^\]]+\]\([^)]+\)|^#{1,6}\s+|^\s*[-*+]\s+|^\s*\d+\.\s+|^>\s+|^\|.+\|$/m

function hasIssues(text: string): boolean {
  return hasLatex(text) || hasMarkdown(text)
}

function hasLatex(text: string): boolean {
  return LATEX_SIGNAL.test(text) || /\$[^$]+\$/.test(text) || /\$\$[\s\S]*?\$\$/.test(text) || /\\\(/.test(text) || /\\\[/.test(text)
}

function hasMarkdown(text: string): boolean {
  return MD_SIGNAL.test(text)
}

// ========== PROCESSING ==========

function processText(text: string): (string | Element)[] {
  // Step 1: extract and protect LaTeX blocks
  const latexBlocks: string[] = []
  let processed = text
    .replace(/\\\[([\s\S]*?)\\\]/g, (_, m) => { latexBlocks.push('$$' + m.trim() + '$$'); return `␟L${latexBlocks.length - 1}␟` })
    .replace(/\\\((.+?)\\\)/g,  (_, m) => { latexBlocks.push('$' + m.trim() + '$');    return `␟L${latexBlocks.length - 1}␟` })
    .replace(/\$\$([\s\S]+?)\$\$/g,  (_, m) => { latexBlocks.push('$$' + m.trim() + '$$'); return `␟L${latexBlocks.length - 1}␟` })
    .replace(/\$([^$]+?)\$/g,       (_, m) => { latexBlocks.push('$' + m.trim() + '$');    return `␟L${latexBlocks.length - 1}␟` })

  // Step 2: parse remaining text with marked to convert markdown → HTML
  let html = md.parse(processed) as string

  // Step 3: restore latex blocks
  for (let i = 0; i < latexBlocks.length; i++) {
    html = html.replace(`␟L${i}␟`, latexBlocks[i])
  }

  // Step 4: strip <p> wrapper if text is a single paragraph with only inline content
  html = html.replace(/^<p>(.+?)<\/p>\n*$/s, '$1')

  // Step 5: build elements from HTML
  const div = document.createElement('div')
  div.setAttribute('data-deep-fixed', '1')
  div.innerHTML = html

  const parts: (string | Element)[] = []
  while (div.firstChild) {
    parts.push(div.firstChild as Element)
  }

  // Step 6: render any remaining LaTeX in the new elements
  const katex = (window as any).katex
  if (katex) {
    renderLatexInParts(parts, katex)
  }

  return parts.length > 0 ? parts : [text]
}

function renderLatexInParts(parts: (string | Element)[], katex: any) {
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]
    if (part instanceof Text || typeof part === 'string') continue

    // Walk element to find text nodes with LaTeX
    const walkEl = (el: Element) => {
      for (const child of Array.from(el.childNodes)) {
        if (child.nodeType === Node.TEXT_NODE) {
          const txt = child.textContent || ''
          if (hasLatex(txt)) {
            const span = document.createElement('span')
            try {
              span.innerHTML = txt.replace(/\$\$([\s\S]+?)\$\$/g, (_, m) => katex.renderToString(m.trim(), { displayMode: true, throwOnError: false, strict: false }))
                .replace(/\$([^$]+?)\$/g, (_, m) => katex.renderToString(m.trim(), { displayMode: false, throwOnError: false, strict: false }))
              child.parentNode?.replaceChild(span, child)
            } catch { /* keep original */ }
          }
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          walkEl(child as Element)
        }
      }
    }
    walkEl(part)
  }
}
