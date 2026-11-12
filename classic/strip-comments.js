// Removes comments from JavaScript (and the <script>/<style>/HTML comments of an .html page) — the published panel
// files carry no notes (the originals have real names from past chats as examples). Code is otherwise unchanged.
//   node strip-comments.js in.js > out.js        node strip-comments.js --html in.html > out.html
const fs = require('fs');

const KEYWORDS_BEFORE_REGEX = new Set(['return', 'typeof', 'instanceof', 'in', 'of', 'new', 'delete', 'void', 'throw', 'case', 'do', 'else', 'yield', 'await']);
function stripJs(src) {
  let out = '', i = 0, prev = ''; // prev: last significant token kind/text, to tell a regex from a division
  const n = src.length;
  const regexAllowed = () => {
    if (!prev) return true;
    if (/^[A-Za-z_$][\w$]*$/.test(prev)) return KEYWORDS_BEFORE_REGEX.has(prev);
    if (/^[\d.]/.test(prev)) return false;
    return !/^[)\]}]$/.test(prev) && prev !== 'str';
  };
  function skipString(q) { // returns end index (exclusive) of a '..' / ".." string starting at i
    let j = i + 1;
    while (j < n && src[j] !== q) { if (src[j] === '\\') j++; j++; }
    return j + 1;
  }
  function skipTemplate(start) { // `...${ ... }...`
    let j = start + 1;
    while (j < n && src[j] !== '`') {
      if (src[j] === '\\') { j += 2; continue; }
      if (src[j] === '$' && src[j + 1] === '{') {
        // copy the expression through the same stripper, tracking braces
        let depth = 1, k = j + 2;
        while (k < n && depth) {
          const ch = src[k];
          if (ch === '{') depth++;
          else if (ch === '}') depth--;
          else if (ch === "'" || ch === '"') { let m = k + 1; while (m < n && src[m] !== ch) { if (src[m] === '\\') m++; m++; } k = m; }
          else if (ch === '`') { k = skipTemplate(k) - 1; }
          k++;
        }
        j = k;
        continue;
      }
      j++;
    }
    return j + 1;
  }
  while (i < n) {
    const c = src[i], d = src[i + 1];
    if (c === '/' && d === '/') { while (i < n && src[i] !== '\n') i++; continue; }
    if (c === '/' && d === '*') { const e = src.indexOf('*/', i + 2); const body = src.slice(i, e < 0 ? n : e + 2); i = e < 0 ? n : e + 2; if (body.includes('\n')) out += '\n'; else out += ' '; continue; }
    if (c === "'" || c === '"') { const e = skipString(c); out += src.slice(i, e); i = e; prev = 'str'; continue; }
    if (c === '`') { const e = skipTemplate(i); out += src.slice(i, e); i = e; prev = 'str'; continue; }
    if (c === '/' && regexAllowed()) { // regex literal
      let j = i + 1, inClass = false;
      while (j < n) {
        const ch = src[j];
        if (ch === '\\') { j += 2; continue; }
        if (ch === '[') inClass = true; else if (ch === ']') inClass = false;
        else if (ch === '/' && !inClass) break;
        else if (ch === '\n') break;
        j++;
      }
      j++;
      while (j < n && /[a-z]/i.test(src[j])) j++;
      out += src.slice(i, j); i = j; prev = 'str'; continue;
    }
    if (/[A-Za-z_$]/.test(c)) { let j = i; while (j < n && /[\w$]/.test(src[j])) j++; prev = src.slice(i, j); out += prev; i = j; continue; }
    if (/\d/.test(c)) { let j = i; while (j < n && /[\w.]/.test(src[j])) j++; prev = src.slice(i, j); out += prev; i = j; continue; }
    if (!/\s/.test(c)) prev = c;
    out += c; i++;
  }
  return out.replace(/[ \t]+$/gm, '').replace(/\n{3,}/g, '\n\n');
}
function stripHtml(html) {
  return html
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/(<style[^>]*>)([\s\S]*?)(<\/style>)/g, (m, a, css, b) => a + css.replace(/\/\*[\s\S]*?\*\//g, '') + b)
    .replace(/(<script[^>]*>)([\s\S]*?)(<\/script>)/g, (m, a, js, b) => a + stripJs(js) + b);
}
module.exports = { stripJs, stripHtml };
if (require.main === module) {
  const html = process.argv[2] === '--html', file = process.argv[html ? 3 : 2];
  const src = fs.readFileSync(file, 'utf8');
  process.stdout.write(html ? stripHtml(src) : stripJs(src));
}
