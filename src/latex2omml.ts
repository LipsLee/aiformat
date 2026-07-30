// LaTeX → OMML (Office Math Markup Language) converter
// Generates native Word/WPS math formulas — no fonts or CSS needed.

/**
 * Convert a LaTeX math string to an OMML XML string.
 * Supports: fractions, super/subscripts, roots, integrals, sums, Greek letters,
 * common operators, matrices, and basic text.
 */

interface MathElement {
  toOmml(): string
}

export function latexToOmml(latex: string, display: boolean): string {
  const parser = new LatexParser(latex)
  const elements = parser.parse()
  const inner = elements.map(e => e.toOmml()).join('')

  if (display) {
    return `<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>${inner}</m:oMath></m:oMathPara>`
  }
  return `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${inner}</m:oMath>`
}

// ---- Tokenizer ----

interface Token {
  type: 'command' | 'group' | 'char' | 'space' | 'subscript' | 'superscript' | 'eof'
  value: string
  children?: Token[]
}

class Tokenizer {
  private pos = 0
  private src: string

  constructor(src: string) {
    this.src = src.trim()
  }

  private peek(): string {
    return this.pos < this.src.length ? this.src[this.pos] : ''
  }

  private next(): string {
    return this.pos < this.src.length ? this.src[this.pos++] : ''
  }

  tokenize(): Token[] {
    const tokens: Token[] = []
    while (this.pos < this.src.length) {
      const ch = this.peek()

      if (ch === ' ' || ch === '\t' || ch === '\n') {
        this.pos++
        tokens.push({ type: 'space', value: ' ' })
        continue
      }

      if (ch === '\\') {
        this.pos++
        let cmd = ''
        while (this.pos < this.src.length && /[a-zA-Z]/.test(this.src[this.pos])) {
          cmd += this.src[this.pos++]
        }
        if (!cmd) {
          const c = this.next()
          cmd = c
        }
        tokens.push({ type: 'command', value: cmd })
        continue
      }

      if (ch === '{') {
        this.pos++
        const children: Token[] = []
        let depth = 1
        let subSrc = ''
        while (this.pos < this.src.length && depth > 0) {
          const c = this.next()
          if (c === '{') depth++
          else if (c === '}') {
            depth--
            if (depth === 0) break
          }
          subSrc += c
        }
        const subTokenizer = new Tokenizer(subSrc)
        children.push(...subTokenizer.tokenize())
        tokens.push({ type: 'group', value: '', children })
        continue
      }

      if (ch === '_') {
        this.pos++
        tokens.push({ type: 'subscript', value: '_' })
        continue
      }

      if (ch === '^') {
        this.pos++
        tokens.push({ type: 'superscript', value: '^' })
        continue
      }

      if (ch === '$') {
        this.pos++
        continue
      }

      this.pos++
      tokens.push({ type: 'char', value: ch })
    }
    tokens.push({ type: 'eof', value: '' })
    return tokens
  }
}

// ---- AST Nodes ----

class MathRun implements MathElement {
  text: string
  italic: boolean

  constructor(text: string, italic = true) {
    this.text = text
    this.italic = italic
  }
  toOmml(): string {
    const escaped = escapeXml(this.text)
    if (this.italic) {
      return `<m:r><m:t>${escaped}</m:t></m:r>`
    }
    return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>${escaped}</m:t></m:r>`
  }
}

class MathFraction implements MathElement {
  num: MathElement[]
  den: MathElement[]

  constructor(num: MathElement[], den: MathElement[]) {
    this.num = num
    this.den = den
  }
  toOmml(): string {
    const numXml = this.num.map(e => e.toOmml()).join('')
    const denXml = this.den.map(e => e.toOmml()).join('')
    return `<m:f><m:num>${numXml}</m:num><m:den>${denXml}</m:den></m:f>`
  }
}

class MathRoot implements MathElement {
  degree: MathElement[] | null
  radicand: MathElement[]

  constructor(degree: MathElement[] | null, radicand: MathElement[]) {
    this.degree = degree
    this.radicand = radicand
  }
  toOmml(): string {
    const radXml = this.radicand.map(e => e.toOmml()).join('')
    if (this.degree && this.degree.length > 0) {
      const degXml = this.degree.map(e => e.toOmml()).join('')
      return `<m:rad><m:deg>${degXml}</m:deg><m:e>${radXml}</m:e></m:rad>`
    }
    return `<m:rad><m:deg/><m:e>${radXml}</m:e></m:rad>`
  }
}

class MathSubSup implements MathElement {
  base: MathElement[]
  sub: MathElement[] | null
  sup: MathElement[] | null
  type: 'subsup' | 'sub' | 'sup'

  constructor(
    base: MathElement[],
    sub: MathElement[] | null,
    sup: MathElement[] | null,
    type: 'subsup' | 'sub' | 'sup' = 'subsup'
  ) {
    this.base = base
    this.sub = sub
    this.sup = sup
    this.type = type
  }
  toOmml(): string {
    const baseXml = this.base.map(e => e.toOmml()).join('')
    if (this.type === 'sub' || (this.sub && !this.sup)) {
      const subXml = this.sub!.map(e => e.toOmml()).join('')
      return `<m:sSub><m:e>${baseXml}</m:e><m:sub>${subXml}</m:sub></m:sSub>`
    }
    if (this.type === 'sup' || (this.sup && !this.sub)) {
      const supXml = this.sup!.map(e => e.toOmml()).join('')
      return `<m:sSup><m:e>${baseXml}</m:e><m:sup>${supXml}</m:sup></m:sSup>`
    }
    const subXml = this.sub!.map(e => e.toOmml()).join('')
    const supXml = this.sup!.map(e => e.toOmml()).join('')
    return `<m:sSubSup><m:e>${baseXml}</m:e><m:sub>${subXml}</m:sub><m:sup>${supXml}</m:sup></m:sSubSup>`
  }
}

class MathNary implements MathElement {
  op: string
  sub: MathElement[] | null
  sup: MathElement[] | null
  body: MathElement[]

  constructor(op: string, sub: MathElement[] | null, sup: MathElement[] | null, body: MathElement[]) {
    this.op = op
    this.sub = sub
    this.sup = sup
    this.body = body
  }
  toOmml(): string {
    const chr = escapeXml(this.op)
    const subXml = this.sub ? this.sub.map(e => e.toOmml()).join('') : ''
    const supXml = this.sup ? this.sup.map(e => e.toOmml()).join('') : ''
    const bodyXml = this.body.map(e => e.toOmml()).join('')
    return `<m:nary><m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="undOvr"/><m:subHide m:val="${this.sub ? 0 : 1}"/><m:supHide m:val="${this.sup ? 0 : 1}"/></m:naryPr><m:sub>${subXml}</m:sub><m:sup>${supXml}</m:sup><m:e>${bodyXml}</m:e></m:nary>`
  }
}

class MathDelim implements MathElement {
  beg: string
  end: string
  body: MathElement[]

  constructor(beg: string, end: string, body: MathElement[]) {
    this.beg = beg
    this.end = end
    this.body = body
  }
  toOmml(): string {
    const bodyXml = this.body.map(e => e.toOmml()).join('')
    const begChr = this.beg === '' ? '' : `m:val="${escapeXml(this.beg)}"`
    const endChr = this.end === '' ? '' : `m:val="${escapeXml(this.end)}"`
    return `<m:d><m:dPr>${begChr ? `<m:begChr ${begChr}/>` : ''}${endChr ? `<m:endChr ${endChr}/>` : ''}</m:dPr><m:e>${bodyXml}</m:e></m:d>`
  }
}

class MathMatrix implements MathElement {
  rows: MathElement[][]

  constructor(rows: MathElement[][]) {
    this.rows = rows
  }
  toOmml(): string {
    const rowsXml = this.rows.map(row => {
      const cellsXml = row.map(cell => `<m:e>${cell.toOmml()}</m:e>`).join('')
      return `<m:mr>${cellsXml}</m:mr>`
    }).join('')
    return `<m:m>${rowsXml}</m:m>`
  }
}

class MathGroup implements MathElement {
  elements: MathElement[]

  constructor(elements: MathElement[]) {
    this.elements = elements
  }
  toOmml(): string {
    return this.elements.map(e => e.toOmml()).join('')
  }
}

// ---- Symbol maps ----

const GREEK_LOWER: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ϵ',
  zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ',
  mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ', varrho: 'ϱ', sigma: 'σ',
  varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'ϕ', chi: 'χ', psi: 'ψ',
  omega: 'ω',
}

const GREEK_UPPER: Record<string, string> = {
  Alpha: 'Α', Beta: 'Β', Gamma: 'Γ', Delta: 'Δ', Epsilon: 'Ε', Zeta: 'Ζ',
  Eta: 'Η', Theta: 'Θ', Iota: 'Ι', Kappa: 'Κ', Lambda: 'Λ', Mu: 'Μ', Nu: 'Ν',
  Xi: 'Ξ', Pi: 'Π', Rho: 'Ρ', Sigma: 'Σ', Tau: 'Τ', Upsilon: 'Υ', Phi: 'Φ',
  Chi: 'Χ', Psi: 'Ψ', Omega: 'Ω',
}

const OPERATORS: Record<string, string> = {
  times: '×', div: '÷', cdot: '·', pm: '±', mp: '∓', infty: '∞',
  partial: '∂', nabla: '∇', forall: '∀', exists: '∃',
  leq: '≤', geq: '≥', neq: '≠', approx: '≈', equiv: '≡', propto: '∝',
  rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftarrow: '⇐',
  leftrightarrow: '↔', Leftrightarrow: '⇔', mapsto: '↦',
  longrightarrow: '⟶', longleftarrow: '⟵',
  subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
  in: '∈', notin: '∉', cup: '∪', cap: '∩', emptyset: '∅',
  angle: '∠', perp: '⊥', parallel: '∥', cdots: '⋯', ldots: '…', vdots: '⋮',
  prime: '′', backslash: '\\', vert: '|', Vert: '‖',
  sum: '∑', prod: '∏', coprod: '∐', int: '∫', oint: '∮', iint: '∬', iiint: '∭',
  bigcup: '⋃', bigcap: '⋂', bigoplus: '⨁', bigotimes: '⨂',
  langle: '⟨', rangle: '⟩', lvert: '|', rvert: '|', lVert: '‖', rVert: '‖',
  to: '→', gets: '←', ast: '∗', star: '⋆', circ: '∘', bullet: '•',
  dagger: '†', ddagger: '‡', ell: 'ℓ', Re: 'ℜ', Im: 'ℑ', aleph: 'ℵ',
  hbar: 'ℏ', imath: 'ı', jmath: 'ȷ', S: '§', P: '¶',
  degree: '°', neg: '¬', land: '∧', lor: '∨', lnot: '¬',
  vartriangle: '△', triangle: '△', square: '□', blacksquare: '■',
}

const FUNCTIONS = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
  'arcsin', 'arccos', 'arctan',
  'sinh', 'cosh', 'tanh', 'coth',
  'log', 'ln', 'lg', 'exp',
  'lim', 'limsup', 'liminf',
  'max', 'min', 'sup', 'inf',
  'det', 'dim', 'arg', 'deg', 'gcd', 'hom', 'ker', 'Pr',
  'mod', 'bmod', 'pmod',
])

const BLACKBOARD_BOLD: Record<string, string> = {
  R: 'ℝ', Z: 'ℤ', Q: 'ℚ', C: 'ℂ', N: 'ℕ', P: 'ℙ', A: '𝔸', B: '𝔹', D: '𝔻', E: '𝔼', F: '𝔽',
  G: '𝔾', H: 'ℍ', I: '𝕀', J: '𝕁', K: '𝕂', L: '𝕃', M: '𝕄', O: '𝕆', S: '𝕊', T: '𝕋', U: '𝕌', V: '𝕍', W: '𝕎', X: '𝕏', Y: '𝕐',
}

// ---- Parser ----

class LatexParser {
  private tokens: Token[]
  private pos = 0

  constructor(src: string) {
    this.tokens = new Tokenizer(src).tokenize()
  }

  private peek(): Token {
    return this.tokens[this.pos] || { type: 'eof', value: '' }
  }

  private next(): Token {
    return this.tokens[this.pos++] || { type: 'eof', value: '' }
  }

  private skipSpaces(): void {
    while (this.peek().type === 'space') this.pos++
  }

  parse(): MathElement[] {
    return this.parseUntil()
  }

  private parseUntil(): MathElement[] {
    const elements: MathElement[] = []

    while (this.peek().type !== 'eof') {
      const tok = this.peek()

      if (tok.type === 'space') {
        this.pos++
        continue
      }

      if (tok.type === 'char') {
        this.pos++
        const base: MathElement[] = [new MathRun(tok.value)]

        const next = this.peek()
        if (next.type === 'subscript' || next.type === 'superscript') {
          this.pos++
          const isSub = next.type === 'subscript'
          const firstEl = this.parseOne()
          const after = this.peek()

          if (isSub && after.type === 'superscript') {
            this.pos++
            const supEl = this.parseOne()
            elements.push(new MathSubSup(base, firstEl, supEl))
          } else if (!isSub) {
            const supEl = firstEl
            const afterSub = this.peek()
            if (afterSub.type === 'subscript') {
              this.pos++
              const actualSub = this.parseOne()
              elements.push(new MathSubSup(base, actualSub, supEl))
            } else {
              elements.push(new MathSubSup(base, null, supEl))
            }
          } else {
            elements.push(new MathSubSup(base, firstEl, null))
          }
          continue
        }

        elements.push(...base)
        continue
      }

      if (tok.type === 'group') {
        this.pos++
        const groupParser = new LatexParser('')
        groupParser.tokens = tok.children || []
        groupParser.pos = 0
        const groupElements = groupParser.parse()
        elements.push(new MathGroup(groupElements))
        continue
      }

      if (tok.type === 'subscript' || tok.type === 'superscript') {
        this.pos++
        const isSub = tok.type === 'subscript'
        const firstEl = this.parseOne()
        const after = this.peek()

        if (isSub && after.type === 'superscript') {
          this.pos++
          const supEl = this.parseOne()
          const last = elements.pop()
          if (last) {
            elements.push(new MathSubSup([last], firstEl, supEl))
          }
        } else if (!isSub) {
          const supEl = firstEl
          const afterSub = this.peek()
          if (afterSub.type === 'subscript') {
            this.pos++
            const actualSub = this.parseOne()
            const last = elements.pop()
            if (last) {
              elements.push(new MathSubSup([last], actualSub, supEl))
            }
          } else {
            const last = elements.pop()
            if (last) {
              elements.push(new MathSubSup([last], null, supEl))
            }
          }
        } else {
          const last = elements.pop()
          if (last) {
            elements.push(new MathSubSup([last], firstEl, null))
          }
        }
        continue
      }

      if (tok.type === 'command') {
        this.pos++
        const cmd = tok.value
        const result = this.handleCommand(cmd)
        if (result) {
          elements.push(result)
        }
        continue
      }

      this.pos++
    }

    return elements
  }

  private parseOne(): MathElement[] {
    this.skipSpaces()
    const tok = this.peek()

    if (tok.type === 'char') {
      this.pos++
      return [new MathRun(tok.value)]
    }

    if (tok.type === 'group') {
      this.pos++
      const groupParser = new LatexParser('')
      groupParser.tokens = tok.children || []
      groupParser.pos = 0
      return groupParser.parse()
    }

    if (tok.type === 'command') {
      this.pos++
      const result = this.handleCommand(tok.value)
      return result ? [result] : []
    }

    this.pos++
    return []
  }

  private handleCommand(cmd: string): MathElement | null {
    // Greek letters
    if (GREEK_LOWER[cmd]) return new MathRun(GREEK_LOWER[cmd])
    if (GREEK_UPPER[cmd]) return new MathRun(GREEK_UPPER[cmd], false)

    // Operators and symbols
    if (OPERATORS[cmd]) return new MathRun(OPERATORS[cmd], false)

    // Functions (sin, cos, etc.)
    if (FUNCTIONS.has(cmd)) return new MathRun(cmd, false)

    // \frac{a}{b}
    if (cmd === 'frac' || cmd === 'dfrac' || cmd === 'tfrac' || cmd === 'cfrac') {
      const num = this.parseOne()
      const den = this.parseOne()
      return new MathFraction(num, den)
    }

    // \sqrt{x} or \sqrt[n]{x}
    if (cmd === 'sqrt') {
      this.skipSpaces()
      const next = this.peek()
      let degree: MathElement[] | null = null

      if (next.type === 'char' && next.value === '[') {
        this.pos++
        let degreeSrc = ''
        let depth = 1
        while (this.peek().type !== 'eof' && depth > 0) {
          const t = this.next()
          if (t.type === 'char' && t.value === '[') depth++
          else if (t.type === 'char' && t.value === ']') {
            depth--
            if (depth === 0) break
          }
          degreeSrc += t.value
        }
        const degParser = new LatexParser(degreeSrc)
        degree = degParser.parse()
      }

      const radicand = this.parseOne()
      return new MathRoot(degree, radicand)
    }

    // \root{n}{x}
    if (cmd === 'root') {
      const degree = this.parseOne()
      const radicand = this.parseOne()
      return new MathRoot(degree, radicand)
    }

    // N-ary operators
    const NARY_OPS: Record<string, string> = {
      int: '∫', oint: '∮', iint: '∬', iiint: '∭',
      sum: '∑', prod: '∏', coprod: '∐',
      bigcup: '⋃', bigcap: '⋂', bigoplus: '⨁', bigotimes: '⨂',
      bigvee: '⋁', bigwedge: '⋀',
    }

    if (NARY_OPS[cmd]) {
      this.skipSpaces()
      let sub: MathElement[] | null = null
      let sup: MathElement[] | null = null

      const next = this.peek()
      if (next.type === 'subscript') {
        this.pos++
        sub = this.parseOne()
      }
      const after = this.peek()
      if (after.type === 'superscript') {
        this.pos++
        sup = this.parseOne()
      }

      if (!sup && !sub) {
        if (next.type === 'superscript') {
          this.pos++
          sup = this.parseOne()
          const afterSub = this.peek()
          if (afterSub.type === 'subscript') {
            this.pos++
            sub = this.parseOne()
          }
        }
      }

      const body = this.parseOne()
      return new MathNary(NARY_OPS[cmd], sub, sup, body)
    }

    // \lim — function with optional subscript
    if (cmd === 'lim') {
      this.skipSpaces()
      let sub: MathElement[] | null = null
      const next = this.peek()
      if (next.type === 'subscript') {
        this.pos++
        sub = this.parseOne()
      }
      const limEl = new MathRun('lim', false)
      if (sub) {
        return new MathSubSup([limEl], sub, null)
      }
      return limEl
    }

    // \left and \right — delimiters
    if (cmd === 'left') {
      this.skipSpaces()
      const delimTok = this.next()
      let begChar = delimTok.value
      if (delimTok.type === 'command' && delimTok.value === 'langle') begChar = '⟨'
      if (delimTok.type === 'command' && delimTok.value === 'vert') begChar = '|'
      if (delimTok.type === 'command' && delimTok.value === 'Vert') begChar = '‖'
      if (delimTok.type === 'char' && delimTok.value === '.') begChar = ''

      const body: MathElement[] = []
      while (this.peek().type !== 'eof') {
        const t = this.peek()
        if (t.type === 'command' && t.value === 'right') break
        if (t.type === 'space') { this.pos++; continue }
        if (t.type === 'char') {
          this.pos++
          body.push(new MathRun(t.value))
        } else if (t.type === 'group') {
          this.pos++
          const gp = new LatexParser('')
          gp.tokens = t.children || []
          gp.pos = 0
          body.push(new MathGroup(gp.parse()))
        } else if (t.type === 'command') {
          this.pos++
          const r = this.handleCommand(t.value)
          if (r) body.push(r)
        } else {
          this.pos++
        }
      }

      if (this.peek().type === 'command' && this.peek().value === 'right') {
        this.pos++
        this.skipSpaces()
        const endDelimTok = this.next()
        let endChar = endDelimTok.value
        if (endDelimTok.type === 'command' && endDelimTok.value === 'rangle') endChar = '⟩'
        if (endDelimTok.type === 'command' && endDelimTok.value === 'vert') endChar = '|'
        if (endDelimTok.type === 'command' && endDelimTok.value === 'Vert') endChar = '‖'
        if (endDelimTok.type === 'char' && endDelimTok.value === '.') endChar = ''

        return new MathDelim(begChar, endChar, body)
      }

      return new MathGroup(body)
    }

    if (cmd === 'right') {
      return null
    }

    // \begin{matrix} ... \end{matrix}
    if (cmd === 'begin') {
      this.skipSpaces()
      const envTok = this.next()
      let envName = ''
      if (envTok.type === 'group' && envTok.children) {
        for (const ct of envTok.children) {
          if (ct.type === 'char') envName += ct.value
        }
      }

      const rows: MathElement[][] = []
      let currentRow: MathElement[] = []
      let cellSrc = ''

      while (this.peek().type !== 'eof') {
        const t = this.peek()
        if (t.type === 'command' && t.value === 'end') {
          this.pos++
          this.skipSpaces()
          this.next()
          break
        }
        if (t.type === 'char' && t.value === '&') {
          this.pos++
          const cellParser = new LatexParser(cellSrc)
          currentRow.push(...cellParser.parse())
          cellSrc = ''
          continue
        }
        if (t.type === 'command' && (t.value === '\\' || t.value === 'cr')) {
          this.pos++
          const cellParser = new LatexParser(cellSrc)
          currentRow.push(...cellParser.parse())
          cellSrc = ''
          rows.push(currentRow)
          currentRow = []
          continue
        }
        if (t.type === 'space') {
          this.pos++
          cellSrc += ' '
          continue
        }
        this.pos++
        if (t.type === 'char') cellSrc += t.value
        else if (t.type === 'group') {
          cellSrc += '{'
          if (t.children) {
            for (const ct of t.children) {
              if (ct.type === 'char') cellSrc += ct.value
              else if (ct.type === 'command') cellSrc += '\\' + ct.value
              else if (ct.type === 'subscript') cellSrc += '_'
              else if (ct.type === 'superscript') cellSrc += '^'
              else if (ct.type === 'space') cellSrc += ' '
            }
          }
          cellSrc += '}'
        } else if (t.type === 'command') {
          cellSrc += '\\' + t.value
        }
      }

      if (cellSrc.trim()) {
        const cellParser = new LatexParser(cellSrc)
        currentRow.push(...cellParser.parse())
      }
      if (currentRow.length > 0) rows.push(currentRow)

      if (envName.includes('matrix') || envName.includes('array') || envName.includes('bmatrix') || envName.includes('pmatrix') || envName.includes('vmatrix')) {
        return new MathMatrix(rows)
      }
      return new MathGroup(rows.flat())
    }

    // \text{...} and variants
    if (cmd === 'text' || cmd === 'textrm' || cmd === 'mathrm' || cmd === 'mathit' || cmd === 'mathsf' || cmd === 'textit') {
      const content = this.parseOne()
      const isItalic = (cmd === 'mathit' || cmd === 'textit')
      const result: MathElement[] = []
      for (const el of content) {
        if (el instanceof MathRun) {
          result.push(new MathRun(el.text, isItalic))
        } else {
          result.push(el)
        }
      }
      return new MathGroup(result)
    }

    // \mathbf{...} or \textbf{...}
    if (cmd === 'mathbf' || cmd === 'textbf') {
      const content = this.parseOne()
      const result: MathElement[] = []
      for (const el of content) {
        if (el instanceof MathRun) {
          result.push(new MathRun(el.text, false))
        } else {
          result.push(el)
        }
      }
      return new MathGroup(result)
    }

    // \mathbb{...}
    if (cmd === 'mathbb') {
      const content = this.parseOne()
      const result: MathElement[] = []
      for (const el of content) {
        if (el instanceof MathRun) {
          const bb = BLACKBOARD_BOLD[el.text] || el.text
          result.push(new MathRun(bb, false))
        } else {
          result.push(el)
        }
      }
      return new MathGroup(result)
    }

    // \overline, \underline, \overrightarrow, etc.
    if (cmd === 'overline' || cmd === 'underline' || cmd === 'overrightarrow' || cmd === 'overleftarrow' || cmd === 'widehat' || cmd === 'widetilde' || cmd === 'overbrace' || cmd === 'underbrace') {
      const content = this.parseOne()
      let opChar = '̄'
      if (cmd === 'underline') opChar = '̱'
      else if (cmd === 'overrightarrow') opChar = '→'
      else if (cmd === 'overleftarrow') opChar = '←'
      else if (cmd === 'widehat') opChar = '̂'
      else if (cmd === 'widetilde') opChar = '̃'
      else if (cmd === 'overbrace') opChar = '⏞'
      else if (cmd === 'underbrace') opChar = '⏟'

      return new MathGroup([
        new MathGroup(content),
        new MathRun(opChar, false)
      ])
    }

    // \quad, \qquad, \,, \;, \!, \:
    if (cmd === 'quad' || cmd === 'qquad' || cmd === ',' || cmd === ';' || cmd === ':' || cmd === '!') {
      const spaces: Record<string, string> = {
        quad: '\u2003\u2003', qquad: '\u2003\u2003\u2003\u2003',
        ',': '\u2009', ';': '\u2004', ':': '\u2005', '!': '',
      }
      return new MathRun(spaces[cmd] || ' ', false)
    }

    // \not — negation
    if (cmd === 'not') {
      const next = this.parseOne()
      return new MathGroup(next)
    }

    // Special characters
    if (cmd === '{') return new MathRun('{', false)
    if (cmd === '}') return new MathRun('}', false)
    if (cmd === '%') return new MathRun('%', false)
    if (cmd === '#') return new MathRun('#', false)
    if (cmd === '&') return new MathRun('&', false)
    if (cmd === '|') return new MathRun('|', false)
    if (cmd === '\\') return new MathRun('\\', false)

    // \operatorname{...}
    if (cmd === 'operatorname') {
      const content = this.parseOne()
      const result: MathElement[] = []
      for (const el of content) {
        if (el instanceof MathRun) {
          result.push(new MathRun(el.text, false))
        } else {
          result.push(el)
        }
      }
      return new MathGroup(result)
    }

    // \binom{n}{k}
    if (cmd === 'binom' || cmd === 'dbinom' || cmd === 'tbinom') {
      const n = this.parseOne()
      const k = this.parseOne()
      return new MathDelim('(', ')', [new MathFraction(n, k)])
    }

    // \big, \Big, \bigg, \Bigg — just ignore, treat next as normal
    if (cmd === 'big' || cmd === 'Big' || cmd === 'bigg' || cmd === 'Bigg' || cmd === 'bigl' || cmd === 'bigr' || cmd === 'Bigl' || cmd === 'Bigr' || cmd === 'biggl' || cmd === 'biggr' || cmd === 'Biggl' || cmd === 'Biggr') {
      this.skipSpaces()
      const next = this.parseOne()
      return new MathGroup(next)
    }

    // Unknown command — render as text
    return new MathRun('\\' + cmd, false)
  }
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

// ---- DOM helpers ----

/**
 * Extract original LaTeX from a KaTeX-rendered element.
 * KaTeX stores the original LaTeX in annotation elements.
 */
export function extractLatex(el: Element): { latex: string; display: boolean } | null {
  const annotation = el.querySelector('.katex annotation[encoding="application/x-tex"]') ||
                     el.querySelector('annotation[encoding="application/x-tex"]')
  if (annotation) {
    const latex = annotation.textContent || ''
    const display = el.classList.contains('katex-display') ||
                    el.closest('.katex-display') !== null
    return { latex: latex.trim(), display }
  }
  return null
}

/**
 * Check if an element is a KaTeX math block
 */
export function isKatexDisplay(el: Element): boolean {
  return el.classList.contains('katex-display') || el.classList.contains('math-block')
}

/**
 * Check if an element is inside KaTeX inline math
 */
export function isInsideKatex(el: Element): boolean {
  return el.closest('.katex') !== null
}
