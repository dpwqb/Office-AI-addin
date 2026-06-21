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

  function colAlign(cell) {
    const c = String(cell || '').trim();
    const l = c.startsWith(':'), r = c.endsWith(':');
    return l && r ? 'center' : r ? 'right' : l ? 'left' : '';
  }

  function isHorizontalRule(line) {
    return /^\s{0,3}([-*_])\s*(\1\s*){2,}$/.test(line || '');
  }

  // Allow only safe URL schemes; returns the (escaped) url or null when blocked.
  function sanitizeUrl(url, allowData) {
    const u = String(url || '').trim();
    if (/^(https?:|mailto:)/i.test(u)) return u;
    if (/^#/.test(u)) return u;
    if (/^\.?\.?\//.test(u)) return u;                  // relative path
    if (allowData && /^data:image\//i.test(u)) return u;
    return null;                                        // block javascript:, vbscript:, data:text/html, ...
  }

  function renderInline(text) {
    let html = App.escapeHtml(text || '');
    const parts = [];      // protected fragments (code spans, anchors) restored last-but-one
    const escs = [];       // backslash-escaped literals restored very last
    const protect = (fragment) => {
      const key = `${SENTINEL}P${parts.length}${SENTINEL}`;
      parts.push(fragment);
      return key;
    };

    // 1) Backslash escapes — protect before any other rule can touch them.
    html = html.replace(/\\([\\`*_{}\[\]()#+\-.!~>])/g, (_, ch) => {
      const key = `${SENTINEL}E${escs.length}${SENTINEL}`;
      escs.push(App.escapeHtml(ch));
      return key;
    });

    // 2) Inline code — double backtick before single.
    html = html.replace(/``(.+?)``/g, (_, code) => protect(`<code>${code.trim()}</code>`));
    html = html.replace(/`([^`]+)`/g, (_, code) => protect(`<code>${code}</code>`));

    // 3) Images — before links so the link rule doesn't partially match.
    html = html.replace(/!\[([^\]]*)\]\(([^\s)]+)(?:\s+"[^"]*")?\)/g, (_, alt, url) => {
      const safe = sanitizeUrl(url, true);
      return safe ? protect(`<img src="${safe}" alt="${alt}" loading="lazy">`) : alt;
    });

    // 4) Citation links.
    html = html.replace(/\[([^\]]+)\]\(#cite:([^\)]+)\)/g, (_, label, ref) => protect(`<a href="#cite:${ref}" class="citation">${label}</a>`));

    // 5) Standard links — sanitize scheme; blocked => keep just the label.
    html = html.replace(/\[([^\]]+)\]\(([^\s)]+)(?:\s+"[^"]*")?\)/g, (_, label, url) => {
      const safe = sanitizeUrl(url, false);
      return safe ? protect(`<a href="${safe}" target="_blank" rel="noreferrer">${label}</a>`) : label;
    });

    // 6) Autolinks <url> then bare URLs (< > are already &lt; &gt; here).
    html = html.replace(/&lt;((?:https?:\/\/|mailto:)[^\s>]+)&gt;/g, (_, url) => protect(`<a href="${url}" target="_blank" rel="noreferrer">${url}</a>`));
    html = html.replace(/(^|[\s(])((?:https?:\/\/)[^\s<>()]+)/g, (_, pre, url) => {
      const trail = (url.match(/[.,)]+$/) || [''])[0];
      const clean = url.slice(0, url.length - trail.length);
      return `${pre}${protect(`<a href="${clean}" target="_blank" rel="noreferrer">${clean}</a>`)}${trail}`;
    });

    // 7) Bold+italic before bold before italic.
    html = html.replace(/\*\*\*([^*]+)\*\*\*/g, '<strong><em>$1</em></strong>');
    html = html.replace(/___([^_]+)___/g, '<strong><em>$1</em></strong>');
    html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/__([^_]+)__/g, '<strong>$1</strong>');
    html = html.replace(/~~([^~]+)~~/g, '<del>$1</del>');
    html = html.replace(/(^|\s)\*([^*\n]+)\*(?=\s|$|[,.!?;:，。！？；：])/g, '$1<em>$2</em>');
    html = html.replace(/(^|\s)_([^_\n]+)_(?=\s|$|[,.!?;:，。！？；：])/g, '$1<em>$2</em>');

    // Restore protected fragments, then escaped literals (last, so they never feed a regex).
    parts.forEach((part, idx) => { html = html.replace(`${SENTINEL}P${idx}${SENTINEL}`, () => part); });
    escs.forEach((ch, idx) => { html = html.replace(`${SENTINEL}E${idx}${SENTINEL}`, () => ch); });
    return html;
  }

  const LIST_RE = /^(\s*)([-*+]|\d+\.)\s+(.*)$/;

  function renderListItemInner(text) {
    const task = String(text).match(/^\[([ xX])\]\s+(.*)$/);
    if (task) {
      const checked = task[1].toLowerCase() === 'x' ? ' checked' : '';
      return { inner: `<input type="checkbox" disabled${checked}> ${renderInline(task[2])}`, task: true };
    }
    return { inner: renderInline(text), task: false };
  }

  // Stack-based, indentation-aware list parser. Returns rendered html + next line index.
  function renderList(lines, start) {
    const tokens = [];
    let i = start;
    while (i < lines.length) {
      const line = lines[i];
      if (!line.trim()) {
        // a blank line is tolerated only if a list item follows immediately
        if (i + 1 < lines.length && LIST_RE.test(lines[i + 1])) { i++; continue; }
        break;
      }
      const m = line.match(LIST_RE);
      if (m) {
        const indent = m[1].replace(/\t/g, '  ').length;
        tokens.push({ indent, ordered: /\d/.test(m[2]), text: m[3] });
        i++;
      } else if (/^\s+\S/.test(line) && tokens.length) {
        // indented continuation line -> append to previous item
        tokens[tokens.length - 1].text += '\n' + line.trim();
        i++;
      } else {
        break;
      }
    }

    // Walk tokens with a stack of open lists; each level tracks whether its <li> is open.
    let html = '';
    const stack = [];                       // [{ indent, ordered, liOpen }]
    const top = () => stack[stack.length - 1];
    const openList = (tok, task) => {
      html += `<${tok.ordered ? 'ol' : 'ul'}${task ? ' class="task-list"' : ''}>`;
      stack.push({ indent: tok.indent, ordered: tok.ordered, liOpen: false });
    };
    const closeTopList = () => {
      const lvl = stack.pop();
      if (lvl.liOpen) html += '</li>';
      html += `</${lvl.ordered ? 'ol' : 'ul'}>`;
    };

    tokens.forEach(tok => {
      const { inner, task } = renderListItemInner(tok.text);
      // Close levels deeper than the current indent (their parent <li> stays open).
      while (stack.length && tok.indent < top().indent) closeTopList();

      if (!stack.length || tok.indent > top().indent) {
        openList(tok, task);
      } else {
        if (top().liOpen) { html += '</li>'; top().liOpen = false; }
        if (tok.ordered !== top().ordered) { closeTopList(); openList(tok, task); }
      }
      html += `<li${task ? ' class="task-item"' : ''}>${inner}`;
      top().liOpen = true;
    });
    while (stack.length) closeTopList();
    return { html, nextIndex: i };
  }

  function renderMarkdown(md) {
    const src = String(md || '').replace(/\r\n/g, '\n');
    const lines = src.split('\n');
    const out = [];
    let i = 0;

    const isBlockStart = (line, next) => {
      const tline = line || '';
      return /^```/.test(tline.trim()) || /^\s{0,3}#{1,6}\s+/.test(tline) || /^\s*>\s?/.test(tline) || /^\s*[-*+]\s+/.test(tline) || /^\s*\d+\.\s+/.test(tline) || isHorizontalRule(tline) || (tline.includes('|') && next && isTableDivider(next));
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
        const langLabel = fence[1] ? App.escapeHtml(fence[1]) : '';
        out.push(
          `<div class="md-codeblock"${langLabel ? ` data-lang="${langLabel}"` : ''}>` +
            `<div class="md-codebar"><span class="md-lang">${langLabel || 'text'}</span>` +
            `<button type="button" class="md-copy" data-action="copy-code">Copy</button></div>` +
            `<pre class="md-code"><code>${App.escapeHtml(code.join('\n'))}</code></pre>` +
          `</div>`
        );
        continue;
      }

      // Setext heading: text line followed by an underline (===/---).
      const next = lines[i + 1] || '';
      if (line.trim() && !isBlockStart(line, next) && !isTableDivider(next) && /^\s{0,3}(=+|-+)\s*$/.test(next)) {
        const level = next.trim()[0] === '=' ? 1 : 2;
        out.push(`<h${level}>${renderInline(line.trim())}</h${level}>`);
        i += 2;
        continue;
      }

      // Horizontal rule.
      if (isHorizontalRule(line)) { out.push('<hr>'); i++; continue; }

      const heading = line.match(/^\s{0,3}(#{1,6})\s+(.*)$/);
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

      if (LIST_RE.test(line)) {
        const { html, nextIndex } = renderList(lines, i);
        out.push(html);
        i = nextIndex;
        continue;
      }

      if (line.includes('|') && lines[i + 1] && isTableDivider(lines[i + 1])) {
        const headers = splitTableRow(line);
        const aligns = splitTableRow(lines[i + 1]).map(colAlign);
        const styleFor = idx => (aligns[idx] ? ` style="text-align:${aligns[idx]}"` : '');
        i += 2;
        const rows = [];
        while (i < lines.length && lines[i].includes('|') && lines[i].trim()) rows.push(splitTableRow(lines[i++]));
        const thead = headers.map((h, idx) => `<th${styleFor(idx)}>${renderInline(h)}</th>`).join('');
        const tbody = rows.map(row => `<tr>${headers.map((_, idx) => `<td${styleFor(idx)}>${renderInline(row[idx] || '')}</td>`).join('')}</tr>`).join('');
        out.push(`<div class="md-table-wrap"><table><thead><tr>${thead}</tr></thead><tbody>${tbody}</tbody></table></div>`);
        continue;
      }

      const para = [];
      while (i < lines.length && lines[i].trim() && !isBlockStart(lines[i], lines[i + 1]) && !/^\s{0,3}(=+|-+)\s*$/.test(lines[i + 1] || '')) para.push(lines[i++]);
      if (!para.length) { out.push(`<p>${renderInline(lines[i].trim())}</p>`); i++; continue; }
      out.push(`<p>${para.map(x => renderInline(x.trim())).join('<br>')}</p>`);
    }
    return out.join('');
  }

  App.splitTableRow = splitTableRow;
  App.isTableDivider = isTableDivider;
  App.renderInline = renderInline;
  App.renderMarkdown = renderMarkdown;
})();
