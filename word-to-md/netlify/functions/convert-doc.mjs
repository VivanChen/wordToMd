import WordExtractor from 'word-extractor';

/**
 * Netlify Function: convert-doc
 * 
 * Receives a .doc file as base64, extracts content using word-extractor,
 * and returns structured HTML that the client can convert to Markdown.
 * 
 * POST /api/convert-doc
 * Body: { filename: string, fileBase64: string }
 * Returns: { html: string, warnings: string[] }
 */
export async function handler(event) {
  // CORS headers
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Content-Type': 'application/json',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({ error: 'Method not allowed' }),
    };
  }

  try {
    const { fileBase64, filename } = JSON.parse(event.body);

    if (!fileBase64) {
      return {
        statusCode: 400,
        headers,
        body: JSON.stringify({ error: 'Missing fileBase64 in request body' }),
      };
    }

    // Decode base64 to buffer
    const buffer = Buffer.from(fileBase64, 'base64');
    const warnings = [];

    // Extract content from .doc
    const extractor = new WordExtractor();
    const doc = await extractor.extract(buffer);

    // Build HTML from extracted content
    const html = buildHtmlFromDoc(doc, warnings);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ html, warnings, filename }),
    };
  } catch (err) {
    console.error('Conversion error:', err);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({
        error: `Conversion failed: ${err.message}`,
        details: err.stack,
      }),
    };
  }
}

/**
 * Convert word-extractor Document to HTML
 */
function buildHtmlFromDoc(doc, warnings) {
  const parts = [];

  // Get body text - word-extractor returns paragraphs
  const body = doc.getBody();
  if (!body) {
    warnings.push('Document body is empty');
    return '<p>(Empty document)</p>';
  }

  // word-extractor getBody() returns full text
  // We need to process it paragraph by paragraph
  const paragraphs = body.split('\n');

  let inList = false;
  let listType = null;

  for (const para of paragraphs) {
    const trimmed = para.trim();
    if (!trimmed) {
      if (inList) {
        parts.push(listType === 'ul' ? '</ul>' : '</ol>');
        inList = false;
        listType = null;
      }
      continue;
    }

    // Detect heading patterns (all caps, short lines, or numbered patterns)
    if (isLikelyHeading(trimmed, paragraphs)) {
      if (inList) {
        parts.push(listType === 'ul' ? '</ul>' : '</ol>');
        inList = false;
        listType = null;
      }
      const level = detectHeadingLevel(trimmed);
      parts.push(`<h${level}>${escapeHtml(trimmed)}</h${level}>`);
      continue;
    }

    // Detect list items
    const bulletMatch = trimmed.match(/^[\u2022\u2023\u25E6\u2043\u2219•●○◦‣⁃\-\*]\s+(.+)/);
    const numberedMatch = trimmed.match(/^(\d+)[.)]\s+(.+)/);

    if (bulletMatch) {
      if (!inList || listType !== 'ul') {
        if (inList) parts.push(listType === 'ul' ? '</ul>' : '</ol>');
        parts.push('<ul>');
        inList = true;
        listType = 'ul';
      }
      parts.push(`<li>${escapeHtml(bulletMatch[1])}</li>`);
      continue;
    }

    if (numberedMatch) {
      if (!inList || listType !== 'ol') {
        if (inList) parts.push(listType === 'ul' ? '</ul>' : '</ol>');
        parts.push('<ol>');
        inList = true;
        listType = 'ol';
      }
      parts.push(`<li>${escapeHtml(numberedMatch[2])}</li>`);
      continue;
    }

    // Close list if no longer in list context
    if (inList) {
      parts.push(listType === 'ul' ? '</ul>' : '</ol>');
      inList = false;
      listType = null;
    }

    // Detect table-like content (tab-separated)
    if (trimmed.includes('\t')) {
      const tableLines = [trimmed];
      // Look ahead for more table rows - handled inline below
      parts.push(buildTableRow(trimmed));
      continue;
    }

    // Regular paragraph
    parts.push(`<p>${escapeHtml(trimmed)}</p>`);
  }

  if (inList) {
    parts.push(listType === 'ul' ? '</ul>' : '</ol>');
  }

  // Post-process: merge consecutive table rows into proper tables
  let html = parts.join('\n');
  html = mergeTableRows(html);

  // Try to extract and append any tables from the document
  try {
    const tables = doc.getTableData?.();
    if (tables && tables.length > 0) {
      for (const table of tables) {
        html += buildTable(table);
      }
    }
  } catch {
    // getTableData might not be available in all versions
  }

  // Add headers/footers info if present
  try {
    const headers = doc.getHeaders?.({ includeFooters: false });
    const footers = doc.getFooters?.({ includeHeaders: false });
    if (headers) {
      const headerText = typeof headers === 'string' ? headers.trim() : '';
      if (headerText) {
        warnings.push(`Document has header: "${headerText.substring(0, 50)}..."`);
      }
    }
    if (footers) {
      const footerText = typeof footers === 'string' ? footers.trim() : '';
      if (footerText) {
        warnings.push(`Document has footer: "${footerText.substring(0, 50)}..."`);
      }
    }
  } catch {}

  // Add endnotes/footnotes if present
  try {
    const endnotes = doc.getEndnotes?.();
    if (endnotes && endnotes.trim()) {
      html += `<hr/><h3>Notes</h3>`;
      for (const note of endnotes.split('\n').filter(n => n.trim())) {
        html += `<p>${escapeHtml(note.trim())}</p>`;
      }
    }
  } catch {}

  try {
    const footnotes = doc.getFootnotes?.();
    if (footnotes && footnotes.trim()) {
      html += `<hr/><h3>Footnotes</h3>`;
      for (const note of footnotes.split('\n').filter(n => n.trim())) {
        html += `<p>${escapeHtml(note.trim())}</p>`;
      }
    }
  } catch {}

  return html;
}

function isLikelyHeading(text, allParagraphs) {
  // Short text that looks like a heading
  if (text.length <= 80 && text.length >= 2) {
    // All caps
    if (text === text.toUpperCase() && /[A-Z]/.test(text) && text.length <= 60) return true;
    // Numbered section: "1. Title" or "1.1 Title" or "Chapter 1"
    if (/^(\d+\.?\d*\.?\s+|Chapter\s+\d+|第\s*[一二三四五六七八九十\d]+\s*[章節篇])/i.test(text)) return true;
    // Chinese heading patterns
    if (/^[一二三四五六七八九十]+[、.．]\s*/.test(text) && text.length <= 40) return true;
    if (/^[（(][一二三四五六七八九十\d]+[)）]\s*/.test(text) && text.length <= 40) return true;
  }
  return false;
}

function detectHeadingLevel(text) {
  if (/^(Chapter|第\s*[一二三四五六七八九十]+\s*章)/i.test(text)) return 1;
  if (/^[一二三四五六七八九十]+[、.．]/.test(text)) return 2;
  if (/^\d+\.\s+/.test(text) && text === text.toUpperCase()) return 2;
  if (/^\d+\.\d+/.test(text)) return 3;
  if (/^[（(][一二三四五六七八九十\d]+[)）]/.test(text)) return 3;
  if (text === text.toUpperCase() && /[A-Z]/.test(text)) return 2;
  return 3;
}

function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildTableRow(line) {
  const cells = line.split('\t').map(c => c.trim());
  const tds = cells.map(c => `<td>${escapeHtml(c)}</td>`).join('');
  return `<tr__temp>${tds}</tr__temp>`;
}

function mergeTableRows(html) {
  // Merge consecutive <tr__temp> rows into proper <table>
  return html.replace(
    /(<tr__temp>.*?<\/tr__temp>\n?)+/g,
    (match) => {
      const rows = match.replace(/tr__temp>/g, 'tr>');
      return `<table>${rows}</table>`;
    }
  );
}

function buildTable(tableData) {
  if (!Array.isArray(tableData) || tableData.length === 0) return '';
  let html = '<table>';
  for (let i = 0; i < tableData.length; i++) {
    const row = Array.isArray(tableData[i]) ? tableData[i] : [tableData[i]];
    const tag = i === 0 ? 'th' : 'td';
    html += '<tr>';
    for (const cell of row) {
      html += `<${tag}>${escapeHtml(String(cell || ''))}</${tag}>`;
    }
    html += '</tr>';
  }
  html += '</table>';
  return html;
}
