import { Marked } from 'marked'
import type { StyleConfig } from './style'

const marked = new Marked({ gfm: true, breaks: true })

export function parseMarkdown(text: string, c: StyleConfig): string {
  const body = marked.parse(text) as string
  return `<style>${makeStyles(c)}</style><div class="doc">${body}</div>`
}

export function makeStyles(c: StyleConfig): string {
  return `
.doc{font-family:${c.fontFamily};font-size:${c.fontSize}px;line-height:${c.lineHeight};color:#1a1a1a;word-break:break-word;}
.doc h1{font-size:1.55em;font-weight:700;margin:1.4em 0 0.5em;padding-bottom:8px;border-bottom:2px solid #1a1a1a;}
.doc h2{font-size:1.3em;font-weight:700;margin:1.2em 0 0.4em;}
.doc h3{font-size:1.12em;font-weight:600;margin:1.1em 0 0.3em;}
.doc h4,.doc h5,.doc h6{font-size:1.05em;font-weight:600;margin:1em 0 0.25em;color:#333;}
.doc p{margin:0 0 ${c.paraSpacing}px 0;}
.doc li p,.doc blockquote p,.doc td p,.doc th p{margin:4px 0;}
.doc strong,.doc b{font-weight:600;}
.doc em,.doc i{font-style:italic;}
.doc a{color:#2563eb;text-decoration:none;}
.doc ul,.doc ol{margin:0.5em 0;padding-left:1.8em;}
.doc li{margin-bottom:0.25em;}
.doc blockquote{margin:0.9em 0;padding:10px 18px;border-left:3px solid #ddd;color:#555;font-size:.95em;}
.doc code{font-family:"SF Mono","Cascadia Code","JetBrains Mono",Consolas,monospace;font-size:.88em;background:#f2f1ef;padding:2px 6px;border-radius:3px;color:#c7254e;}
.doc pre{margin:0.9em 0;padding:16px 20px;overflow-x:auto;background:#1e1e1e;border-radius:6px;}
.doc pre code{font-size:.84em;line-height:1.55;color:#d4d4d4;background:none;padding:0;white-space:pre;}
.doc table{width:100%;border-collapse:collapse;font-size:.92em;margin:0.9em 0;}
.doc thead{border-bottom:2px solid #ddd;}
.doc th{padding:9px 14px;text-align:left;font-weight:600;color:#555;font-size:.88em;letter-spacing:.3px;}
.doc td{padding:9px 14px;border-bottom:1px solid #eee;}
.doc tbody tr:hover td{background:#fafaf9;}
.doc hr{border:none;border-top:1px solid #ddd;margin:2em 0;}
.doc img{max-width:100%;height:auto;margin:.8em 0;}
`.trim()
}
