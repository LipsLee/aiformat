// Deep correction — walk DOM text nodes, find unrendered LaTeX, re-render

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
  if ((window as any).renderMathInElement) {
    try {
      (window as any).renderMathInElement(root, {
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
    if (hasLatex(text)) {
      const parts = parseMixedContent(text)
      if (parts && parts.some(p => typeof p !== 'string')) {
        patches.push({ node: node as Text, parts })
      }
    }
    return
  }

  // Skip already-rendered KaTeX nodes, code blocks, pre blocks, script, style
  if (node.nodeType !== Node.ELEMENT_NODE) return
  const el = node as Element
  const tag = el.tagName.toLowerCase()
  if (
    tag === 'script' || tag === 'style' ||
    el.classList.contains('katex') || el.classList.contains('katex-display') ||
    el.closest('.katex') || el.closest('pre') || el.closest('code') ||
    el.closest('.math-block') || el.closest('.math-inline') ||
    el.closest('[data-deep-fixed]')
  ) return

  for (const child of Array.from(el.childNodes)) {
    walk(child, patches)
  }
}

const LATEX_SIGNAL = /\\frac|\\sum|\\int|\\prod|\\sqrt|\\alpha|\\beta|\\gamma|\\delta|\\epsilon|\\theta|\\lambda|\\mu|\\pi|\\sigma|\\omega|\\infty|\\partial|\\nabla|\\times|\\div|\\pm|\\cdot|\\leq|\\geq|\\neq|\\approx|\\equiv|\\rightarrow|\\Rightarrow|\\leftarrow|\\Leftarrow|\\subset|\\supset|\\in|\\notin|\\forall|\\exists|\\mathbb|\\mathbf|\\mathcal|\\text|\\begin|\\end|\\lim|\\log|\\ln|\\sin|\\cos|\\tan|\\det|\\max|\\min/

function hasLatex(text: string): boolean {
  return LATEX_SIGNAL.test(text) || /\$[^$]+\$/.test(text) || /\$\$[\s\S]*?\$\$/.test(text) || /\\\(/.test(text) || /\\\[/.test(text)
}

function parseMixedContent(text: string): (string | Element)[] {
  const parts: (string | Element)[] = []

  // Build a regex that captures:
  // - $$ ... $$ (block math)
  // - $ ... $ (inline math)
  // - \( ... \) (inline math)
  // - \[ ... \] (block math)
  // - Standalone \frac{...}{...} patterns (broken formulas)
  const pattern = /(\$\$[\s\S]+?\$\$)|(\$[^$]+\$)|(\\\(.+?\\\))|(\\\[[\s\S]+?\\\])/g

  let lastIndex = 0
  let match = pattern.exec(text)

  while (match) {
    // Text before match
    if (match.index > lastIndex) {
      parts.push(text.slice(lastIndex, match.index))
    }

    const raw = match[0]
    try {
      let latex: string
      let displayMode: boolean

      if (raw.startsWith('$$')) {
        latex = raw.slice(2, -2).trim()
        displayMode = true
      } else if (raw.startsWith('\\[')) {
        latex = raw.slice(2, -2).trim()
        displayMode = true
      } else if (raw.startsWith('\\(')) {
        latex = raw.slice(2, -2).trim()
        displayMode = false
      } else if (raw.startsWith('$')) {
        latex = raw.slice(1, -1).trim()
        displayMode = false
      } else {
        parts.push(raw)
        lastIndex = pattern.lastIndex
        match = pattern.exec(text)
        continue
      }

      const katex = (window as any).katex
      if (katex) {
        const span = document.createElement('span')
        span.setAttribute('data-deep-fixed', '1')
        span.innerHTML = katex.renderToString(latex, {
          displayMode,
          throwOnError: false,
          strict: false,
        })
        span.classList.add(displayMode ? 'math-block' : 'math-inline')
        if (displayMode) span.style.display = 'block'
        parts.push(span)
      } else {
        parts.push(raw)
      }
    } catch {
      parts.push(raw)
    }

    lastIndex = pattern.lastIndex
    match = pattern.exec(text)
  }

  // Remaining text
  if (lastIndex < text.length) {
    parts.push(text.slice(lastIndex))
  }

  return parts.length > 0 ? parts : [text]
}
