(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const SENTINEL = String.fromCharCode(0);

  function splitTableRow(line) {
    let text = String(line || '').trim();
    if (text.startsWith('|')) text = text.slice(1);
    if (text.endsWith('|')) text = text.slice(0, -1);
    return text.split('|').map(cell => cell.trim());
  }

  function isTableDivider(line) {
    const cells = splitTableRow(line);
    return cells.length > 1 && cells.every(cell => /^:?-{3,}:?$/.test(cell));
  }

  function renderInline(text) {
    let html = App.escapeHtml(text || '');
    const codeParts = [];
    html = html.replace(/`([^`]+)`/g, (_, code) => {
      const key = `${SENTINEL}CODE${codeParts.length}${SENTINEL}`;
      codeParts.push(`<code>${code}</code>`);
      return key;
    });
    html = html.replace(/\[([^\]]+)\]\(#cite:([^\)]+)\)/g, (_, label, ref) => `<a href="#cite:${ref}" class="citation">${label}</a>`);
    html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^\s\)]+)\)/g, (_, label, url) => `<a href="${url}" target="_blank" rel="noreferrer">${label}</a>`);
    html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/__([^_]+)__/g, '<strong>$1</strong>');
    html = html.replace(/~~([^~]+)~~/g, '<del>$1</del>');
    html = html.replace(/(^|\s)\*([^*\n]+)\*(?=\s|$|[,.!?;:，。！？；：])/g, '$1<em>$2</em>');
    html = html.replace(/(^|\s)_([^_\n]+)_(?=\s|$|[,.!?;:，。！？；：])/g, '$1<em>$2</em>');
    codeParts.forEach((part, i) => { html = html.replace(`${SENTINEL}CODE${i}${SENTINEL}`, part); });
    return html;
  }

  function renderMarkdown(md) {
    const src = String(md || '').replace(/\r\n/g, '\n');
    const lines = src.split('\n');
    const out = [];
    let i = 0;

    const isBlockStart = (line, next) => {
      const tline = line || '';
      return /^```/.test(tline.trim()) || /^\s{0,3}#{1,6}\s+/.test(tline) || /^\s*>\s?/.test(tline) || /^\s*[-*+]\s+/.test(tline) || /^\s*\d+\.\s+/.test(tline) || (tline.includes('|') && next && isTableDivider(next));
    };

    while (i < lines.length) {
      const line = lines[i];
      if (!line.trim()) { i++; continue; }

      const fence = line.trim().match(/^```\s*([\w-]+)?/);
      if (fence) {
        i++;
        const code = [];
        while (i < lines.length && !/^```/.test(lines[i].trim())) code.push(lines[i++]);
        if (i < lines.length) i++;
        const lang = fence[1] ? ` data-lang="${App.escapeHtml(fence[1])}"` : '';
        out.push(`<pre class="md-code"${lang}><code>${App.escapeHtml(code.join('\n'))}</code></pre>`);
        continue;
      }

      const heading = line.match(/^\s{0,3}(#{1,6})\s+(.+)$/);
      if (heading) {
        const level = heading[1].length;
        out.push(`<h${level}>${renderInline(heading[2].trim())}</h${level}>`);
        i++;
        continue;
      }

      if (/^\s*>\s?/.test(line)) {
        const quote = [];
        while (i < lines.length && /^\s*>\s?/.test(lines[i])) quote.push(lines[i++].replace(/^\s*>\s?/, ''));
        out.push(`<blockquote>${renderMarkdown(quote.join('\n'))}</blockquote>`);
        continue;
      }

      if (/^\s*[-*+]\s+/.test(line) || /^\s*\d+\.\s+/.test(line)) {
        const ordered = /^\s*\d+\.\s+/.test(line);
        const items = [];
        const re = ordered ? /^\s*\d+\.\s+(.+)$/ : /^\s*[-*+]\s+(.+)$/;
        while (i < lines.length && re.test(lines[i])) {
          const item = lines[i].match(re)[1];
          items.push(`<li>${renderInline(item)}</li>`);
          i++;
        }
        out.push(`<${ordered ? 'ol' : 'ul'}>${items.join('')}</${ordered ? 'ol' : 'ul'}>`);
        continue;
      }

      if (line.includes('|') && lines[i + 1] && isTableDivider(lines[i + 1])) {
        const headers = splitTableRow(line);
        i += 2;
        const rows = [];
        while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(splitTableRow(lines[i++]));
        const thead = headers.map(h => `<th>${renderInline(h)}</th>`).join('');
        const tbody = rows.map(row => `<tr>${headers.map((_, idx) => `<td>${renderInline(row[idx] || '')}</td>`).join('')}</tr>`).join('');
        out.push(`<div class="md-table-wrap"><table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table></div>`);
        continue;
      }

      const para = [];
      while (i < lines.length && lines[i].trim() && !isBlockStart(lines[i], lines[i + 1])) para.push(lines[i++]);
      out.push(`<p>${para.map(x => renderInline(x.trim())).join('<br>')}</p>`);
    }
    return out.join('');
  }

  App.splitTableRow = splitTableRow;
  App.isTableDivider = isTableDivider;
  App.renderInline = renderInline;
  App.renderMarkdown = renderMarkdown;
})();
