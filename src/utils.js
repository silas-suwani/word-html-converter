import DOMPurify from 'dompurify';
import TurndownService from 'turndown';

const VOID_TAGS = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr']);
const PRESERVE_CONTENT_TAGS = 'img,iframe,video,audio,svg,table,ul,ol,li,input,embed,object,canvas,source';

export function sanitizeHtml(html = '') {
  return DOMPurify.sanitize(String(html), {
    USE_PROFILES: { html: true },
    ADD_ATTR: ['target', 'style', 'class', 'id', 'width', 'height', 'colspan', 'rowspan', 'start', 'align'],
  });
}

function unwrap(element) {
  const parent = element.parentNode;
  if (!parent) return;
  while (element.firstChild) parent.insertBefore(element.firstChild, element);
  parent.removeChild(element);
}

export function cleanHtml(html = '', options = {}) {
  const safeHtml = sanitizeHtml(html);
  const parser = new DOMParser();
  const doc = parser.parseFromString(safeHtml, 'text/html');
  const body = doc.body;

  if (options.removeComments) {
    const walker = doc.createTreeWalker(body, NodeFilter.SHOW_COMMENT);
    const comments = [];
    while (walker.nextNode()) comments.push(walker.currentNode);
    comments.forEach((node) => node.parentNode?.removeChild(node));
  }

  if (options.removeInlineStyles) {
    body.querySelectorAll('[style]').forEach((node) => node.removeAttribute('style'));
  }

  if (options.removeClassesIds) {
    body.querySelectorAll('[class], [id]').forEach((node) => {
      node.removeAttribute('class');
      node.removeAttribute('id');
    });
  }

  if (options.removeBold) body.querySelectorAll('b, strong').forEach(unwrap);
  if (options.removeItalic) body.querySelectorAll('i, em').forEach(unwrap);
  if (options.removeUnderline) body.querySelectorAll('u').forEach(unwrap);
  if (options.removeLinks) body.querySelectorAll('a').forEach(unwrap);
  if (options.removeFontTags) body.querySelectorAll('font, span, small, big, mark, sub, sup, s, strike, del, ins').forEach(unwrap);

  if (options.stripFormatting) {
    body.querySelectorAll('b, strong, i, em, u, s, strike, del, ins, font, span, small, big, mark, sub, sup').forEach(unwrap);
    body.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach((heading) => {
      const paragraph = doc.createElement('p');
      while (heading.firstChild) paragraph.appendChild(heading.firstChild);
      heading.replaceWith(paragraph);
    });
    body.querySelectorAll('[style], [class], [id], [align]').forEach((node) => {
      node.removeAttribute('style');
      node.removeAttribute('class');
      node.removeAttribute('id');
      node.removeAttribute('align');
    });
  }

  if (options.removeEmptyTags) {
    // Repeat because removing an empty inline element can leave its parent empty.
    let changed = true;
    while (changed) {
      changed = false;
      [...body.querySelectorAll('*')].reverse().forEach((node) => {
        if (VOID_TAGS.has(node.tagName.toLowerCase())) return;
        if (node.querySelector(PRESERVE_CONTENT_TAGS)) return;
        const hasText = node.textContent.trim().length > 0;
        const hasMeaningfulChild = [...node.children].some((child) => child.tagName.toLowerCase() !== 'br');
        if (!hasText && !hasMeaningfulChild) {
          node.remove();
          changed = true;
        }
      });
    }
  }

  // Do not let a pasted HTML fragment smuggle executable URLs into the output.
  body.querySelectorAll('a[href], img[src]').forEach((node) => {
    const attr = node.tagName.toLowerCase() === 'a' ? 'href' : 'src';
    const value = node.getAttribute(attr) || '';
    if (/^\s*javascript:/i.test(value)) node.removeAttribute(attr);
  });

  return body.innerHTML.trim();
}

/** Pretty-print block elements while keeping inline tags together. Each heading gets one blank source line above it. */
export function formatHtml(html = '') {
  const original = sanitizeHtml(html).trim();
  let output = original;
  if (!output) return '';
  const startsWithHeading = /^<h[1-6]\b/i.test(original);

  const blockTag = '(?:address|article|aside|blockquote|div|dl|dt|dd|fieldset|figcaption|figure|footer|form|h[1-6]|header|hr|li|main|nav|ol|p|pre|section|table|caption|colgroup|thead|tbody|tfoot|tr|td|th|ul)';
  const blockRegex = new RegExp(`<\\/?${blockTag}\\b[^>]*>`, 'gi');
  output = output.replace(blockRegex, '\n$&');
  output = output.replace(/\n*<h([1-6])\b/gi, '\n\n<h$1');
  output = output.replace(/[\t ]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  // Even a first-element heading receives the requested blank source line above it.
  return startsWithHeading ? `\n\n${output}` : output;
}

export function htmlToText(html = '') {
  const safeHtml = sanitizeHtml(html);
  const doc = new DOMParser().parseFromString(safeHtml, 'text/html');
  doc.body.querySelectorAll('br').forEach((node) => node.replaceWith('\n'));
  doc.body.querySelectorAll('img').forEach((node) => {
    const alt = node.getAttribute('alt');
    node.replaceWith(alt ? `[Image: ${alt}]` : '[Image]');
  });
  doc.body.querySelectorAll('p,div,section,article,h1,h2,h3,h4,h5,h6,li,blockquote,tr').forEach((node) => {
    node.append('\n');
  });
  return doc.body.textContent.replace(/[\t ]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

export function htmlToMarkdown(html = '') {
  const turndown = new TurndownService({
    headingStyle: 'atx',
    bulletListMarker: '-',
    codeBlockStyle: 'fenced',
  });
  return turndown.turndown(sanitizeHtml(html));
}

export function wrapHtmlDocument(fragment, title = 'Converted document') {
  return `<!doctype html>\n<html lang="en">\n<head>\n  <meta charset="utf-8">\n  <meta name="viewport" content="width=device-width, initial-scale=1">\n  <title>${escapeHtml(title)}</title>\n</head>\n<body>\n${formatHtml(fragment).split('\n').map((line) => `  ${line}`).join('\n')}\n</body>\n</html>`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[char]));
}

export async function embedImagesAsBase64(html = '') {
  // Parsing in an inert document lets us detect file:/// image sources in Word clipboard HTML;
  // the final fragment is sanitized after image embedding attempts.
  const doc = new DOMParser().parseFromString(String(html), 'text/html');
  const images = [...doc.body.querySelectorAll('img[src]')];
  const results = await Promise.all(images.map(async (img) => {
    const src = img.getAttribute('src') || '';
    if (!src || src.startsWith('data:')) return { embedded: false, failed: false };
    if (/^(?:blob:|file:|javascript:)/i.test(src)) return { embedded: false, failed: true };
    try {
      const response = await fetch(src, { mode: 'cors' });
      if (!response.ok) throw new Error(`Image fetch failed: ${response.status}`);
      const blob = await response.blob();
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
      img.setAttribute('src', dataUrl);
      return { embedded: true, failed: false };
    } catch {
      return { embedded: false, failed: true };
    }
  }));
  return {
    html: sanitizeHtml(doc.body.innerHTML).trim(),
    embeddedCount: results.filter((result) => result.embedded).length,
    failedCount: results.filter((result) => result.failed).length,
  };
}

export function downloadFile(filename, content, mimeType = 'text/plain;charset=utf-8') {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

export async function copyText(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();
  const success = document.execCommand('copy');
  textarea.remove();
  if (!success) throw new Error('Clipboard access is not available in this browser.');
}

export function countWords(html = '') {
  const text = htmlToText(html);
  return text ? text.split(/\s+/).filter(Boolean).length : 0;
}
