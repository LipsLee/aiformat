// Simulate what happens when KaTeX renders a formula and we extract textContent
// KaTeX renders \int_{a}^{b} into HTML like:
// <span class="katex">
//   <span class="katex-mathml">
//     <math xmlns="...">
//       <semantics>
//         <mrow>...</mrow>
//         <annotation encoding="application/x-tex">\int_{a}^{b} u(x)v'(x)dx = ...</annotation>
//       </semantics>
//     </math>
//   </span>
//   <span class="katex-html" aria-hidden="true">
//     <span class="base">...∫...</span>
//   </span>
// </span>

// The key issue: KaTeX has TWO rendering paths:
// 1. .katex-mathml contains <math> + <annotation> — MathML for accessibility
// 2. .katex-html contains the visual HTML spans with Unicode chars

// When we do textContent on the whole thing, we get BOTH:
// - The MathML text content (which includes the annotation LaTeX source)
// - The HTML text content (Unicode math symbols)

// Our code removes .katex annotation, but does NOT remove .katex-mathml
// So textContent still gets:
// 1. MathML's <mrow> text (which may duplicate the Unicode output)
// 2. HTML's text (the Unicode symbols)

// Let's verify:
const html = `
<span class="katex">
  <span class="katex-mathml">
    <math xmlns="http://www.w3.org/1998/Math/MathML">
      <semantics>
        <mrow>
          <msubsup><mo>∫</mo><mi>a</mi><mi>b</mi></msubsup>
          <mi>u</mi><mo>(</mo><mi>x</mi><mo>)</mo>
        </mrow>
        <annotation encoding="application/x-tex">\\int_{a}^{b} u(x)v'(x)dx</annotation>
      </semantics>
    </math>
  </span>
  <span class="katex-html" aria-hidden="true">
    <span class="base">
      <span class="strut"></span>
      <span class="mop">∫</span>
      <span class="msupsub">
        <span class="vlist">
          <span style="top:-2.5em"><span>b</span></span>
          <span style="top:-3.2em"><span>a</span></span>
        </span>
      </span>
      <span class="mord mathnormal">u</span>
      <span class="mopen">(</span>
      <span class="mord mathnormal">x</span>
      <span class="mclose">)</span>
    </span>
  </span>
</span>
`

// Simulate DOM
const div = document.createElement('div')
div.innerHTML = html

console.log('=== Full textContent (BEFORE cleanup) ===')
console.log(div.textContent.replace(/\s+/g, ' ').trim())

// Remove annotation only (what our current code does)
const clone1 = div.cloneNode(true)
clone1.querySelectorAll('.katex annotation').forEach(a => a.remove())
console.log('\n=== After removing annotation only ===')
console.log(clone1.textContent.replace(/\s+/g, ' ').trim())

// Remove katex-mathml entirely (what we SHOULD do)
const clone2 = div.cloneNode(true)
clone2.querySelectorAll('.katex-mathml').forEach(a => a.remove())
console.log('\n=== After removing .katex-mathml ===')
console.log(clone2.textContent.replace(/\s+/g, ' ').trim())

// Remove BOTH annotation AND katex-mathml
const clone3 = div.cloneNode(true)
clone3.querySelectorAll('.katex-mathml').forEach(a => a.remove())
clone3.querySelectorAll('.katex annotation').forEach(a => a.remove())
console.log('\n=== After removing both .katex-mathml and annotation ===')
console.log(clone3.textContent.replace(/\s+/g, ' ').trim())
