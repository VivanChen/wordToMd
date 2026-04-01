import { useState, useCallback, useRef, useEffect, useMemo } from "react";
import * as mammoth from "mammoth";

/* ═══════════════════════════════════════════════════════════
   Phase 1: HTML → Raw Markdown
   mammoth HTML has: <a id="bookmark"></a>, <p> inside <td>,
   colspan, nested bold/italic, Word TOC links, etc.
   ═══════════════════════════════════════════════════════════ */
function htmlToMarkdown(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  return convertNode(doc.body).replace(/\n{3,}/g, "\n\n").trim();
}

function convertNode(node) {
  if (node.nodeType === 3) return node.textContent;
  if (node.nodeType !== 1) return "";
  const tag = node.tagName.toLowerCase();
  const ch = () => Array.from(node.childNodes).map(convertNode).join("");

  switch (tag) {
    case "h1": case "h2": case "h3":
    case "h4": case "h5": case "h6": {
      const lvl = +tag[1];
      const text = ch().trim();
      if (!text) return "\n"; // skip empty headings (from bookmarks)
      return `\n${"#".repeat(lvl)} ${text}\n\n`;
    }
    case "p": {
      const text = ch().trim();
      if (!text) return "\n";
      return `\n${text}\n\n`;
    }
    case "br": return " ";
    case "strong": case "b": {
      const c = ch();
      return c.trim() ? `**${c}**` : "";
    }
    case "em": case "i": {
      const c = ch();
      return c.trim() ? `*${c}*` : "";
    }
    case "u": return `<u>${ch()}</u>`;
    case "s": case "del": case "strike": return `~~${ch()}~~`;
    case "code": return `\`${ch()}\``;
    case "pre": return `\n\`\`\`\n${ch().trim()}\n\`\`\`\n\n`;
    case "blockquote": return `\n> ${ch().trim().replace(/\n/g, "\n> ")}\n\n`;
    case "a": {
      const href = node.getAttribute("href") || "";
      const id = node.getAttribute("id") || "";
      const text = ch().trim();
      // Word bookmark anchor: <a id="_Toc..."></a>
      if (id && !text) return "";
      if (!href && !text) return "";
      if (href) return `[${text}](${href})`;
      return text;
    }
    case "img": {
      const src = node.getAttribute("src") || "";
      if (src.startsWith("data:")) return "";
      return `![image](${src})`;
    }
    case "ul": {
      return "\n" + Array.from(node.children)
        .filter(c => c.tagName?.toLowerCase() === "li")
        .map(li => `- ${convertNode(li).trim()}`).join("\n") + "\n\n";
    }
    case "ol": {
      return "\n" + Array.from(node.children)
        .filter(c => c.tagName?.toLowerCase() === "li")
        .map((li, i) => `${i + 1}. ${convertNode(li).trim()}`).join("\n") + "\n\n";
    }
    case "li": return ch();
    case "table": return convertTable(node);
    case "hr": return "\n---\n\n";
    case "sup": return `<sup>${ch()}</sup>`;
    case "sub": return `<sub>${ch()}</sub>`;
    default: return ch();
  }
}

function convertTable(tableNode) {
  const rows = Array.from(tableNode.querySelectorAll("tr"));
  if (!rows.length) return "";

  const data = rows.map(row =>
    Array.from(row.querySelectorAll("th, td")).map(cell => {
      // Get all <p> content inside cell, join with space (NOT newline)
      const ps = Array.from(cell.querySelectorAll("p"));
      let text;
      if (ps.length > 0) {
        text = ps.map(p => convertNode(p).replace(/^\n+|\n+$/g, "").trim()).filter(Boolean).join(" ");
      } else {
        text = convertNode(cell).trim();
      }
      // Collapse whitespace
      return text.replace(/\n+/g, " ").replace(/\s{2,}/g, " ").trim();
    })
  );

  const cols = Math.max(...data.map(r => r.length));
  data.forEach(r => { while (r.length < cols) r.push(""); });

  const widths = Array.from({ length: cols }, (_, i) =>
    Math.max(3, ...data.map(r => (r[i] || "").length))
  );
  const fmtRow = r => "| " + Array.from({ length: cols }, (_, i) => (r[i] || "").padEnd(widths[i])).join(" | ") + " |";
  const sep = "| " + widths.map(w => "-".repeat(w)).join(" | ") + " |";
  return "\n" + [fmtRow(data[0]), sep, ...data.slice(1).map(fmtRow)].join("\n") + "\n\n";
}

/* ═══════════════════════════════════════════════════════════
   Phase 2: Post-Processing Pipeline
   Fixes Word-specific artifacts in the raw MD
   ═══════════════════════════════════════════════════════════ */
function postProcess(md, opts = {}) {
  let r = md;

  // 1. Remove empty bookmark links: []() or []()[]()...
  r = r.replace(/\[\s*\]\([^)]*\)/g, "");

  // 2. Unwrap bold/italic that wraps an entire line (Word header formatting)
  r = r.replace(/^\*{1,3}([^*\n]+)\*{1,3}$/gm, "$1");

  // 3. Remove empty headings (## with no text)
  r = r.replace(/^#{1,6}\s*$/gm, "");

  // 4. Remove Edition History
  if (opts.removeEditionHistory !== false) {
    r = removeEditionHistory(r);
  }

  // 5. Fix TOC
  r = fixTOC(r);

  // 6. Convert metadata tables → key-value
  r = fixMetadataTables(r);

  // 7. Fix field data tables
  r = fixDataTables(r);

  // 8. Add section numbers
  r = addSectionNumbers(r);

  // 9. Build document header
  r = addDocHeader(r);

  // 10. Remove stray "Index" lines (leftover from metadata tables)
  r = r.replace(/^\n*Index\n*$/gm, "");

  // 11. Cleanup
  r = r.replace(/\n{3,}/g, "\n\n").trim();

  return r;
}

function removeEditionHistory(md) {
  const lines = md.split("\n");
  let skip = false;
  const out = [];
  for (const line of lines) {
    if (/Edition History|修改紀錄/.test(line) && !skip) { skip = true; continue; }
    if (skip) {
      if (/^#{1,6}\s/.test(line)) { skip = false; out.push(line); }
      continue;
    }
    out.push(line);
  }
  return out.join("\n");
}

function fixTOC(md) {
  const lines = md.split("\n");
  const tocIdx = lines.findIndex(l => /^##?\s*目錄/.test(l.trim()));
  if (tocIdx === -1) return md;

  const entries = [];
  let endIdx = tocIdx + 1;

  for (let i = tocIdx + 1; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) { endIdx = i + 1; continue; }
    if (line.startsWith("##") || line.startsWith("---")) { endIdx = i; break; }

    // [1.\tAP_ATM  銀行名稱資料檔\t5](#_Toc...)
    const m = line.match(/^\[(\d+)\.\s*(.+?)[\s\t]+\d+\]\(#[^)]*\)$/);
    if (m) {
      entries.push({ num: +m[1], title: m[2].replace(/[\s\t]+/g, " ").trim() });
      endIdx = i + 1;
    } else if (/^\[(?:\d+\.\s*\d+|目錄[\s\t]*\d+)\]\(#[^)]*\)$/.test(line)) {
      endIdx = i + 1; // skip garbage
    } else if (line.startsWith("[")) {
      endIdx = i + 1; // skip other TOC lines
    } else {
      endIdx = i; break;
    }
  }

  if (!entries.length) return md;

  const toc = entries.map(e => {
    const slug = `${e.num}-${e.title.toLowerCase().replace(/[()（）]/g, "").replace(/\s+/g, "-")}`;
    return `${e.num}. [${e.title}](#${slug})`;
  });

  return [...lines.slice(0, tocIdx + 1), ...toc, "", "---", "", ...lines.slice(endIdx)].join("\n");
}

function fixMetadataTables(md) {
  const lines = md.split("\n");
  const out = [];
  let i = 0;

  while (i < lines.length) {
    if (/^\|.*(?:Table\s*Name|TableName)/i.test(lines[i]) && !/FieldName|Field Name/i.test(lines[i])) {
      // Collect table
      const tbl = [lines[i]];
      let j = i + 1;
      while (j < lines.length && lines[j].trimStart().startsWith("|")) { tbl.push(lines[j]); j++; }

      const m = parseMeta(tbl);
      if (m) {
        out.push(`**Table Name:** ${m.tableName}`);
        out.push(`**DB Name:** ${m.dbName}`);
        out.push(`**System:** ${m.system}`);
        out.push(`**Primary Key:** ${m.primaryKey || "-"}`);
        out.push(`**Index:** ${m.index || "-"}`);
        out.push("");
        i = j;
        continue;
      }
    }
    out.push(lines[i]);
    i++;
  }
  return out.join("\n");
}

function parseMeta(tableLines) {
  const rows = tableLines.filter(l => !/^\|\s*-+/.test(l));
  const m = { tableName: "", dbName: "", system: "", primaryKey: "", index: "" };

  for (const row of rows) {
    const cells = row.split("|").slice(1, -1).map(c => c.trim());
    for (let c = 0; c < cells.length; c++) {
      const cell = cells[c].replace(/\[([^\]]*)\]\([^)]*\)/g, "$1").replace(/\*+/g, "").trim();
      const next = (cells[c + 1] || "").replace(/\[([^\]]*)\]\([^)]*\)/g, "$1").replace(/\*+/g, "").trim();
      const k = cell.toLowerCase().replace(/\s/g, "");

      if (k === "tablename" || k === "table name") m.tableName = next;
      else if (k === "dbname" || k === "db name") m.dbName = next;
      else if (k === "system") m.system = next;
      else if (k === "primarykey" || k === "primary key") {
        m.primaryKey = cells.slice(c + 1).map(s => s.trim()).filter(Boolean).join(", ") || "-";
      }
      else if (k === "index") {
        m.index = cells.slice(c + 1).map(s => s.trim()).filter(Boolean).join(", ") || "-";
      }
    }
  }
  return m.tableName ? m : null;
}

function fixDataTables(md) {
  const lines = md.split("\n");
  const out = [];
  let i = 0;

  while (i < lines.length) {
    if (/^\|.*(?:FieldName|Field\s*Name)/i.test(lines[i])) {
      const tbl = [lines[i]];
      let j = i + 1;
      while (j < lines.length && lines[j].startsWith("|")) { tbl.push(lines[j]); j++; }
      out.push(...fixFieldTable(tbl));
      i = j;
      continue;
    }
    out.push(lines[i]);
    i++;
  }
  return out.join("\n");
}

function fixFieldTable(tableLines) {
  const headerCells = tableLines[0].split("|").slice(1, -1).map(c => c.trim());
  const cm = {};
  headerCells.forEach((h, idx) => {
    const k = h.toLowerCase().replace(/\s+/g, "").replace(/\*/g, "");
    if (k === "no") cm.no = idx;
    else if (k.includes("fieldname") || k.includes("field name")) cm.field = idx;
    else if (k === "type") cm.type = idx;
    else if (k === "description") cm.desc = idx;
    else if (k === "default") cm.def = idx;
    else if (k === "null") cm.nullable = idx;
    else if (k === "remark") cm.remark = idx;
  });

  const data = tableLines.slice(2).filter(l => l.startsWith("|") && !/^\|\s*-+/.test(l));
  const rows = [];
  let n = 1;

  for (const row of data) {
    const cells = row.split("|").slice(1, -1).map(c => c.trim());
    const field = (cells[cm.field ?? 1] || "").trim();
    if (!field) continue;

    rows.push({
      no: String(n++),
      field,
      type: (cells[cm.type ?? 2] || "").trim(),
      desc: (cells[cm.desc ?? 3] || "").trim(),
      def: (cells[cm.def ?? 4] || "").trim() || "-",
      nullable: (cells[cm.nullable ?? 5] || "").trim(),
      remark: (cells[cm.remark ?? 6] || "").replace(/\s+/g, " ").trim() || "-",
    });
  }

  if (!rows.length) return tableLines;

  const hdrs = ["No", "FieldName", "Type", "Description", "Default", "Null", "Remark"];
  const keys = ["no", "field", "type", "desc", "def", "nullable", "remark"];
  const w = hdrs.map((h, i) => Math.max(h.length, ...rows.map(r => (r[keys[i]] || "").length)));

  const lines = [];
  lines.push("| " + hdrs.map((h, i) => h.padEnd(w[i])).join(" | ") + " |");
  lines.push("| " + w.map(v => "-".repeat(v)).join(" | ") + " |");
  for (const row of rows) {
    lines.push("| " + keys.map((k, i) => (row[k] || "").padEnd(w[i])).join(" | ") + " |");
  }
  return lines;
}

function addSectionNumbers(md) {
  let num = 0;
  return md.split("\n").map(line => {
    const m = line.match(/^##\s+(.+)$/);
    if (!m) return line;
    const t = m[1].trim();
    if (t === "目錄") return line;
    if (/^\d+\.\s+/.test(t)) { num = +t.match(/^(\d+)/)[1]; return line; }
    return `## ${++num}. ${t}`;
  }).join("\n");
}

function addDocHeader(md) {
  const lines = md.split("\n");
  let code = "", part = "", title = "";
  const clean = [];
  let foundHdr = false;

  for (const line of lines) {
    const pm = line.match(/專案代號\s*[：:（(\s]\s*([^)）\n]+)/);
    if (pm) { code = pm[1].replace(/[)）]/g, "").trim(); foundHdr = true; continue; }
    const cm = line.match(/共同部分\s*[：:（(\s]\s*([^)）\n]+)/);
    if (cm) { part = cm[1].replace(/[)）]/g, "").trim(); foundHdr = true; continue; }
    if (/資料庫規格報告書/.test(line) && !line.startsWith("#")) { title = "資料庫規格報告書"; foundHdr = true; continue; }
    if (!foundHdr && !clean.length && !line.trim()) continue;
    clean.push(line);
  }

  if (!foundHdr) return md;

  const hdr = [
    `# ${part ? part + "_DB_V1" : ""} ${title || "資料庫規格報告書"}`.trim(),
    "",
    ...(code ? [`**專案代號**：${code}  `] : []),
    ...(part ? [`**共同部分**：${part}  `] : []),
    "**系統**：消企金授信管理應用系統",
    "",
  ];

  return [...hdr, ...clean].join("\n");
}

/* ═══════════════════════════════════════════════════════════
   Full Pipeline
   ═══════════════════════════════════════════════════════════ */
function convertDocx(html, opts) {
  return postProcess(htmlToMarkdown(html), opts);
}

/* ═══════════════════════════════════════════════════════════
   Markdown → HTML Preview Renderer
   ═══════════════════════════════════════════════════════════ */
function renderMd(md) {
  if (!md) return "";
  let h = md;
  h = h.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  // Code blocks
  h = h.replace(/```\n?([\s\S]*?)```/g, '<pre class="cb"><code>$1</code></pre>');
  h = h.replace(/`([^`]+)`/g, '<code class="ci">$1</code>');

  // Headings — add id for anchor navigation
  for (let i = 6; i >= 1; i--) {
    const re = new RegExp(`^#{${i}}\\s+(.+)$`, "gm");
    h = h.replace(re, (_, text) => {
      const id = text.trim().toLowerCase()
        .replace(/\*+/g, "")
        .replace(/[()（）]/g, "")
        .replace(/\s+/g, "-")
        .replace(/[^\w\u4e00-\u9fff-]/g, "");
      return `<h${i} class="mh mh${i}" id="${id}">${text}</h${i}>`;
    });
  }

  h = h.replace(/^---$/gm, '<hr class="mhr"/>');
  h = h.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  h = h.replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, "<em>$1</em>");
  h = h.replace(/~~(.+?)~~/g, "<del>$1</del>");
  h = h.replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img alt="$1" src="$2" class="mi"/>');
  h = h.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" class="ma">$1</a>');

  // Lists
  h = h.replace(/^\s*[-*]\s+(.+)$/gm, "<li>$1</li>");
  h = h.replace(/(<li>.*<\/li>\n?)+/g, m => `<ul class="mul">${m}</ul>`);
  h = h.replace(/^\d+\.\s+(.+)$/gm, "<oli>$1</oli>");
  h = h.replace(/(<oli>.*<\/oli>\n?)+/g, m => `<ol class="mol">${m.replace(/oli>/g, "li>")}</ol>`);

  // Tables
  h = h.replace(/^\|(.+)\|$/gm, (match) => {
    const cells = match.slice(1, -1).split("|").map(c => c.trim());
    if (cells.every(c => /^-+$/.test(c))) return "<!--sep-->";
    return `<tr>${cells.map(c => `<td class="mtd">${c}</td>`).join("")}</tr>`;
  });
  h = h.replace(/((<tr>.*<\/tr>\n?)+)/g, m => {
    let t = m.replace(/<!--sep-->\n?/g, "");
    const fr = t.match(/<tr>(.*?)<\/tr>/);
    if (fr) t = t.replace(fr[0], fr[0].replace(/<td class="mtd">/g, '<td class="mth">'));
    return `<table class="mt">${t}</table>`;
  });

  h = h.replace(/^&gt;\s?(.+)$/gm, '<blockquote class="mbq">$1</blockquote>');

  // Paragraphs — handle consecutive <strong> lines as metadata block with <br>
  h = h.split("\n").map((l, i, arr) => {
    const t = l.trim();
    if (!t || /^<(h[1-6]|ul|ol|li|pre|table|tr|td|hr|blockquote|div|p|img)/.test(t) || /^<!--/.test(t)) return t;
    // Check if this is a bold metadata line followed by another bold line
    const isBoldLine = /^<strong>[^<]+<\/strong>/.test(t);
    const nextIsBold = arr[i + 1] && /^<strong>[^<]+<\/strong>/.test(arr[i + 1]?.trim());
    if (isBoldLine && nextIsBold) return `<p class="mp mb0">${t}</p>`;
    return `<p class="mp">${t}</p>`;
  }).join("\n");

  return h;
}

/* Extract TOC entries from markdown */
function extractToc(md) {
  if (!md) return [];
  const entries = [];
  const lines = md.split("\n");
  for (const line of lines) {
    const m = line.match(/^(#{1,6})\s+(.+)$/);
    if (m) {
      const level = m[1].length;
      const text = m[2].trim();
      const id = text.toLowerCase()
        .replace(/\*+/g, "")
        .replace(/[()（）]/g, "")
        .replace(/\s+/g, "-")
        .replace(/[^\w\u4e00-\u9fff-]/g, "");
      entries.push({ level, text, id });
    }
  }
  return entries;
}

/* ═══════════════════════════════════════════════════════════
   Icons
   ═══════════════════════════════════════════════════════════ */
const I = {
  upload: <svg width="44" height="44" fill="none" viewBox="0 0 48 48"><path d="M24 32V16m0 0l-8 8m8-8l8 8" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/><path d="M8 32v4a4 4 0 004 4h24a4 4 0 004-4v-4" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"/></svg>,
  file: <svg width="18" height="18" fill="none" viewBox="0 0 20 20"><path d="M4 2h8l5 5v11a1 1 0 01-1 1H4a1 1 0 01-1-1V3a1 1 0 011-1z" stroke="currentColor" strokeWidth="1.5"/><path d="M12 2v5h5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>,
  dl: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><path d="M9 3v8m0 0l-3-3m3 3l3-3M3 13v1a1 1 0 001 1h10a1 1 0 001-1v-1" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  eye: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><path d="M1.5 9s3-5.5 7.5-5.5S16.5 9 16.5 9s-3 5.5-7.5 5.5S1.5 9 1.5 9z" stroke="currentColor" strokeWidth="1.5"/><circle cx="9" cy="9" r="2.5" stroke="currentColor" strokeWidth="1.5"/></svg>,
  code: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><path d="M5.5 5L2 9l3.5 4M12.5 5L16 9l-3.5 4M10.5 3l-3 12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  ok: <svg width="14" height="14" fill="none" viewBox="0 0 16 16"><path d="M3 8.5l3 3 7-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  x: <svg width="14" height="14" fill="none" viewBox="0 0 16 16"><path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>,
  warn: <svg width="14" height="14" fill="none" viewBox="0 0 16 16"><path d="M8 2l6.5 11H1.5L8 2z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round"/><path d="M8 7v2.5M8 11.5v.01" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>,
  copy: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><rect x="6" y="6" width="10" height="10" rx="1.5" stroke="currentColor" strokeWidth="1.4"/><path d="M12 6V3.5A1.5 1.5 0 0010.5 2h-7A1.5 1.5 0 002 3.5v7A1.5 1.5 0 003.5 12H6" stroke="currentColor" strokeWidth="1.4"/></svg>,
  gear: <svg width="16" height="16" fill="none" viewBox="0 0 20 20"><circle cx="10" cy="10" r="3" stroke="currentColor" strokeWidth="1.5"/><path d="M10 1v2m0 14v2M3.93 3.93l1.41 1.41m9.32 9.32l1.41 1.41M1 10h2m14 0h2M3.93 16.07l1.41-1.41m9.32-9.32l1.41-1.41" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>,
};

/* ═══════════════════════════════════════════════════════════
   Styles
   ═══════════════════════════════════════════════════════════ */
const A = "#7F77DD";
const AL = "rgba(127,119,221,0.08)";
const AB = "rgba(127,119,221,0.2)";

/* ═══════════════════════════════════════════════════════════
   App Component
   ═══════════════════════════════════════════════════════════ */
export default function App() {
  const [page, setPage] = useState("convert");
  const [results, setResults] = useState([]);
  const [idx, setIdx] = useState(0);
  const [view, setView] = useState("preview");
  const [busy, setBusy] = useState(false);
  const [prog, setProg] = useState({ d: 0, t: 0, n: "" });
  const [toast, setToast] = useState(null);
  const [drag, setDrag] = useState(false);
  const [opts, setOpts] = useState({ removeEditionHistory: true });
  const [rdFile, setRdFile] = useState(null);
  const [rdContent, setRdContent] = useState("");
  const ref = useRef(null);

  const notify = useCallback((m, ty = "success") => {
    setToast({ m, ty }); setTimeout(() => setToast(null), 3500);
  }, []);

  const handleFiles = useCallback(async (fl) => {
    const valid = Array.from(fl).filter(f => /\.docx$/i.test(f.name));
    if (!valid.length) { notify("請選擇 .docx 檔案", "error"); return; }
    setIdx(0); setBusy(true);

    const res = [];
    for (let i = 0; i < valid.length; i++) {
      setProg({ d: i, t: valid.length, n: valid[i].name });
      try {
        const buf = await valid[i].arrayBuffer();
        const r = await mammoth.convertToHtml({ arrayBuffer: buf });
        const md = convertDocx(r.value, opts);
        res.push({ file: valid[i], md, html: r.value, ok: true, warn: r.messages?.filter(m => m.type === "warning").length || 0 });
      } catch (e) {
        res.push({ file: valid[i], md: "", html: "", ok: false, err: e.message });
      }
    }
    setResults(res);
    setProg({ d: valid.length, t: valid.length, n: "" });
    setBusy(false);

    const ok = res.filter(r => r.ok).length, fail = res.length - ok;
    if (fail && ok) notify(`${ok} 成功 / ${fail} 失敗`, "warning");
    else if (fail) notify(`轉換失敗: ${res[0].err}`, "error");
    else notify(`${ok} 個檔案轉換完成！`);
  }, [opts]);

  const cur = results[idx];
  const setMd = v => setResults(p => p.map((r, i) => i === idx ? { ...r, md: v } : r));

  const dl = (r) => {
    const nm = r.file.name.replace(/\.\w+$/, "") + ".md";
    const u = URL.createObjectURL(new Blob([r.md], { type: "text/markdown;charset=utf-8" }));
    Object.assign(document.createElement("a"), { href: u, download: nm }).click();
    URL.revokeObjectURL(u);
    notify(`已下載 ${nm}`);
  };
  const cp = async (t) => { try { await navigator.clipboard.writeText(t); notify("已複製"); } catch { notify("複製失敗", "error"); } };

  return (
    <div style={{ fontFamily: '"Noto Sans TC", system-ui, sans-serif', minHeight: "100vh", background: "var(--color-background-tertiary)", color: "var(--color-text-primary)" }}>
      {toast && <div style={{ position: "fixed", top: 20, right: 20, zIndex: 999, padding: "12px 22px", borderRadius: 10, fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 8, background: toast.ty === "error" ? "#E24B4A" : toast.ty === "warning" ? "#BA7517" : A, color: "#fff", boxShadow: "0 8px 32px rgba(0,0,0,0.18)", animation: "slideIn .3s ease" }}>
        {toast.ty === "error" ? I.x : toast.ty === "warning" ? I.warn : I.ok}{toast.m}
      </div>}

      <header style={{ background: "var(--color-background-primary)", borderBottom: `1px solid ${AB}`, padding: "0 24px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 56, position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: `linear-gradient(135deg, ${A}, #AFA9EC)`, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontWeight: 800, fontSize: 14 }}>M</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>Word → Markdown</div>
            <div style={{ fontSize: 10.5, color: "var(--color-text-tertiary)", marginTop: -1 }}>高品質 docx 轉換器</div>
          </div>
        </div>
        <nav style={{ display: "flex", gap: 3 }}>
          {[{ id: "convert", l: "轉換", i: I.file }, { id: "reader", l: "MD 閱讀器", i: I.eye }].map(n => (
            <button key={n.id} onClick={() => setPage(n.id)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 14px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 13, fontWeight: 600, fontFamily: "inherit", background: page === n.id ? AL : "transparent", color: page === n.id ? A : "var(--color-text-secondary)" }}>
              {n.i}{n.l}
            </button>
          ))}
        </nav>
      </header>

      <main style={{ maxWidth: 1320, margin: "0 auto", padding: "24px 20px" }}>
        {page === "convert" && (<div>
          {/* Upload */}
          <div onDrop={e => { e.preventDefault(); setDrag(false); handleFiles(e.dataTransfer.files); }} onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onClick={() => ref.current?.click()} style={{ border: `2px dashed ${drag ? A : AB}`, borderRadius: 14, padding: "34px 20px", textAlign: "center", cursor: "pointer", background: drag ? AL : "var(--color-background-primary)", transition: "all 0.2s", marginBottom: 14 }}>
            <input ref={ref} type="file" accept=".docx" multiple style={{ display: "none" }} onChange={e => handleFiles(e.target.files)} />
            <div style={{ color: A, marginBottom: 8 }}>{I.upload}</div>
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 4 }}>拖曳 .docx 檔案到這裡</div>
            <div style={{ fontSize: 13, color: "var(--color-text-tertiary)" }}>mammoth.js 轉換 + 後處理引擎自動修正 Word 格式瑕疵</div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 16, marginBottom: 14, fontSize: 13, color: "var(--color-text-secondary)" }}>
            {I.gear}
            <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
              <input type="checkbox" checked={opts.removeEditionHistory} onChange={e => setOpts({ ...opts, removeEditionHistory: e.target.checked })} style={{ accentColor: A }} />
              移除修改歷史表 (Edition History)
            </label>
          </div>

          {busy && <div style={{ background: "var(--color-background-primary)", borderRadius: 12, border: `1px solid ${AB}`, padding: "28px 24px", marginBottom: 16, textAlign: "center" }}>
            <div style={{ width: 40, height: 40, border: `3px solid ${AB}`, borderTopColor: A, borderRadius: "50%", animation: "spin .7s linear infinite", margin: "0 auto 12px" }}/>
            <div style={{ fontWeight: 600 }}>轉換中 ({prog.d}/{prog.t})</div>
            <div style={{ fontSize: 13, color: "var(--color-text-tertiary)", marginTop: 4 }}>{prog.n}</div>
          </div>}

          {results.length > 0 && !busy && (<div>
            {results.length > 1 && <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 12 }}>
              {results.map((r, i) => (
                <button key={i} onClick={() => setIdx(i)} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 12px", borderRadius: 8, border: `1px solid ${idx === i ? A : AB}`, background: idx === i ? AL : "var(--color-background-primary)", color: idx === i ? A : "var(--color-text-primary)", cursor: "pointer", fontSize: 12, fontWeight: 600, fontFamily: "inherit" }}>
                  {I.file}<span style={{ maxWidth: 160, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.file.name}</span>
                  {r.ok ? <span style={{ color: "#1D9E75" }}>{I.ok}</span> : <span style={{ color: "#E24B4A" }}>{I.x}</span>}
                </button>
              ))}
            </div>}

            {cur && !cur.ok && <div style={{ background: "rgba(226,75,74,0.06)", border: "1px solid rgba(226,75,74,0.2)", borderRadius: 12, padding: "24px 28px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#E24B4A", fontWeight: 600, marginBottom: 8 }}>{I.x} 轉換失敗</div>
              <div style={{ fontSize: 14, color: "var(--color-text-secondary)" }}>{cur.err}</div>
            </div>}

            {cur?.ok && (<div>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
                <div style={{ display: "flex", background: "var(--color-background-secondary)", borderRadius: 8, padding: 3 }}>
                  {[{ id: "preview", l: "預覽", i: I.eye }, { id: "source", l: "Markdown", i: I.code }, { id: "html", l: "Raw HTML", i: I.file }].map(m => (
                    <button key={m.id} onClick={() => setView(m.id)} style={{ display: "flex", alignItems: "center", gap: 5, padding: "5px 12px", borderRadius: 6, border: "none", fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", background: view === m.id ? "var(--color-background-primary)" : "transparent", color: view === m.id ? A : "var(--color-text-secondary)", boxShadow: view === m.id ? "0 1px 3px rgba(0,0,0,0.06)" : "none" }}>
                      {m.i}{m.l}
                    </button>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  <Btn i={I.copy} l="複製" o={() => cp(cur.md)} outline />
                  <Btn i={I.dl} l="下載 .md" o={() => dl(cur)} />
                </div>
              </div>

              <div style={{ position: "relative" }}>
                {view === "preview" && <MdViewer md={cur.md} maxH={650} />}
                {view === "source" && <div style={{ background: "var(--color-background-primary)", borderRadius: 12, border: `1px solid ${AB}`, overflow: "hidden" }}><textarea value={cur.md} onChange={e => setMd(e.target.value)} style={{ width: "100%", minHeight: 550, padding: "22px 28px", border: "none", resize: "vertical", fontFamily: '"JetBrains Mono", monospace', fontSize: 13, lineHeight: 1.7, background: "transparent", color: "var(--color-text-primary)", outline: "none" }}/></div>}
                {view === "html" && <div style={{ background: "var(--color-background-primary)", borderRadius: 12, border: `1px solid ${AB}`, overflow: "hidden" }}><pre style={{ padding: "22px 28px", fontSize: 12, lineHeight: 1.6, overflowX: "auto", maxHeight: 550, fontFamily: '"JetBrains Mono", monospace', color: "var(--color-text-secondary)", margin: 0, whiteSpace: "pre-wrap", wordBreak: "break-all" }}>{cur.html}</pre></div>}
              </div>

              <div style={{ display: "flex", gap: 14, marginTop: 10, fontSize: 12, color: "var(--color-text-tertiary)", flexWrap: "wrap" }}>
                <span>{cur.md.length.toLocaleString()} 字元</span>
                <span>{cur.md.split(/\s+/).filter(Boolean).length.toLocaleString()} 詞</span>
                <span>{cur.md.split("\n").length} 行</span>
                {cur.warn > 0 && <span style={{ color: "#BA7517" }}>{cur.warn} 個警告</span>}
              </div>
            </div>)}
          </div>)}
        </div>)}

        {page === "reader" && (<div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
            <h2 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>Markdown 閱讀器</h2>
            <label style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 16px", borderRadius: 8, border: `1px solid ${AB}`, background: "var(--color-background-primary)", cursor: "pointer", fontSize: 13, fontWeight: 600, color: "var(--color-text-secondary)", fontFamily: "inherit" }}>
              {I.upload} 開啟 .md 檔案
              <input type="file" accept=".md,.markdown,.txt" style={{ display: "none" }} onChange={async e => { const f = e.target.files[0]; if (!f) return; setRdFile(f.name); setRdContent(await f.text()); }}/>
            </label>
          </div>
          {!rdFile ? (
            <div style={{ background: "var(--color-background-primary)", borderRadius: 12, padding: 60, textAlign: "center", border: `1px solid ${AB}` }}>
              <div style={{ fontWeight: 600, marginBottom: 8 }}>開啟 .md 檔案來預覽</div>
              <div style={{ fontSize: 13, color: "var(--color-text-tertiary)" }}>支援 Markdown 即時渲染</div>
            </div>
          ) : (<div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
              {I.file}<span style={{ fontWeight: 600, fontSize: 14 }}>{rdFile}</span>
              <div style={{ marginLeft: "auto", display: "flex", gap: 6 }}>
                <Btn i={I.copy} l="複製" o={() => cp(rdContent)} outline sm />
                <Btn i={I.x} l="關閉" o={() => { setRdFile(null); setRdContent(""); }} outline sm />
              </div>
            </div>
            <MdViewer md={rdContent} maxH={700} />
          </div>)}
        </div>)}
      </main>

      <style>{`
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes slideIn{from{transform:translateX(100px);opacity:0}to{transform:translateX(0);opacity:1}}
        *{box-sizing:border-box;margin:0}
        button,input,textarea{font-family:inherit}
        ::selection{background:${AL}}
        ::-webkit-scrollbar{width:5px}
        ::-webkit-scrollbar-thumb{background:${AB};border-radius:3px}
        button:hover{filter:brightness(1.05)}
        .toc-sidebar button:hover{background:${AL}!important}
        .toc-sidebar::-webkit-scrollbar{width:3px}
        .md-body{line-height:1.8;font-size:15px;color:var(--color-text-primary);scroll-behavior:smooth}
        .md-body h1[id],.md-body h2[id],.md-body h3[id],.md-body h4[id]{scroll-margin-top:16px}
        .md-body .mh{color:var(--color-text-primary);margin-top:28px;margin-bottom:12px}
        .md-body .mh1{font-size:26px;font-weight:700}
        .md-body .mh2{font-size:21px;font-weight:700;border-bottom:1px solid rgba(127,119,221,0.15);padding-bottom:6px}
        .md-body .mh3{font-size:18px;font-weight:600}
        .md-body .mh4{font-size:16px;font-weight:600}
        .md-body .mh5{font-size:15px;font-weight:600}
        .md-body .mh6{font-size:14px;font-weight:600}
        .md-body .mhr{border:none;border-top:1px solid rgba(127,119,221,0.2);margin:24px 0}
        .md-body .mp{margin:5px 0;line-height:1.8}
        .md-body .mp.mb0{margin-bottom:0;padding-bottom:0}
        .md-body .ma{color:#7F77DD;text-decoration:none}
        .md-body .ma:hover{text-decoration:underline}
        .md-body .mi{max-width:100%;border-radius:8px;margin:8px 0}
        .md-body .mul,.md-body .mol{padding-left:24px;margin:8px 0}
        .md-body .mt{border-collapse:collapse;margin:12px 0;width:100%;font-size:13px}
        .md-body .mtd{border:1px solid rgba(127,119,221,0.12);padding:8px 12px;white-space:nowrap}
        .md-body .mth{border:1px solid rgba(127,119,221,0.12);padding:8px 12px;font-weight:600;background:rgba(127,119,221,0.04);white-space:nowrap}
        .md-body .mbq{border-left:3px solid #7F77DD;padding:8px 20px;margin:12px 0;color:var(--color-text-secondary);background:rgba(127,119,221,0.04);border-radius:0 6px 6px 0}
        .md-body .cb{background:rgba(127,119,221,0.06);padding:16px 20px;border-radius:10px;overflow-x:auto;font-size:0.85em;line-height:1.7;font-family:"JetBrains Mono",monospace;border:1px solid rgba(127,119,221,0.12)}
        .md-body .ci{background:rgba(127,119,221,0.1);padding:2px 6px;border-radius:4px;font-size:0.88em;font-family:"JetBrains Mono",monospace}
      `}</style>
    </div>
  );
}

function Btn({ i, l, o, outline, sm, disabled }) {
  return <button onClick={o} disabled={disabled} style={{ display: "flex", alignItems: "center", gap: 6, padding: sm ? "4px 12px" : "8px 16px", borderRadius: 8, fontSize: sm ? 12 : 13, fontWeight: 600, fontFamily: "inherit", cursor: disabled ? "not-allowed" : "pointer", border: outline ? `1px solid ${AB}` : "none", background: outline ? "var(--color-background-primary)" : A, color: outline ? "var(--color-text-primary)" : "#fff", opacity: disabled ? .5 : 1 }}>{i}{l}</button>;
}

/* ═══════════════════════════════════════════════════════════
   MdViewer: Content panel + TOC sidebar + anchor scroll
   ═══════════════════════════════════════════════════════════ */
function MdViewer({ md, maxH = 700 }) {
  const toc = useMemo(() => extractToc(md), [md]);
  const html = useMemo(() => renderMd(md), [md]);
  const contentRef = useRef(null);
  const [activeId, setActiveId] = useState("");
  const [tocOpen, setTocOpen] = useState(true);

  // Scroll spy: track which heading is visible
  useEffect(() => {
    const el = contentRef.current;
    if (!el) return;
    const onScroll = () => {
      const headings = el.querySelectorAll("h1[id],h2[id],h3[id],h4[id],h5[id],h6[id]");
      let current = "";
      for (const h of headings) {
        if (h.offsetTop - el.scrollTop <= 60) current = h.id;
      }
      setActiveId(current);
    };
    el.addEventListener("scroll", onScroll, { passive: true });
    return () => el.removeEventListener("scroll", onScroll);
  }, [html]);

  // Handle anchor clicks inside the content
  useEffect(() => {
    const el = contentRef.current;
    if (!el) return;
    const onClick = (e) => {
      const a = e.target.closest("a[href^='#']");
      if (!a) return;
      e.preventDefault();
      const id = decodeURIComponent(a.getAttribute("href").slice(1));
      const target = el.querySelector(`[id="${CSS.escape(id)}"]`) || el.querySelector(`[id*="${id}"]`);
      if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
    };
    el.addEventListener("click", onClick);
    return () => el.removeEventListener("click", onClick);
  }, [html]);

  const scrollTo = (id) => {
    const el = contentRef.current;
    if (!el) return;
    const target = el.querySelector(`[id="${CSS.escape(id)}"]`) || el.querySelector(`[id*="${id}"]`);
    if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
  };

  const hasToc = toc.length > 2;

  return (
    <div style={{ display: "flex", gap: 0, background: "var(--color-background-primary)", borderRadius: 12, border: `1px solid ${AB}`, overflow: "hidden" }}>
      {/* TOC Sidebar */}
      {hasToc && (
        <div className="toc-sidebar" style={{
          width: tocOpen ? 260 : 0, minWidth: tocOpen ? 260 : 0,
          borderRight: tocOpen ? `1px solid ${AB}` : "none",
          transition: "all 0.2s",
          overflow: "hidden", flexShrink: 0, position: "relative",
        }}>
          <div style={{ padding: "14px 0", height: maxH, overflowY: "auto", overflowX: "hidden" }}>
            <div style={{ padding: "0 14px 10px", fontSize: 12, fontWeight: 700, color: "var(--color-text-tertiary)", textTransform: "uppercase", letterSpacing: 0.5 }}>目錄</div>
            {toc.map((entry, i) => {
              const isActive = activeId === entry.id;
              const indent = Math.max(0, entry.level - 1) * 14;
              return (
                <button key={i} onClick={() => scrollTo(entry.id)} style={{
                  display: "block", width: "100%", textAlign: "left",
                  padding: `4px 14px 4px ${14 + indent}px`,
                  border: "none", cursor: "pointer", fontFamily: "inherit",
                  fontSize: entry.level <= 1 ? 13 : 12,
                  fontWeight: entry.level <= 2 ? 600 : 400,
                  lineHeight: 1.5,
                  color: isActive ? A : "var(--color-text-secondary)",
                  background: isActive ? AL : "transparent",
                  borderLeft: isActive ? `2px solid ${A}` : "2px solid transparent",
                  transition: "all 0.15s",
                  whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
                }}>
                  {entry.text.replace(/\*+/g, "")}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {/* TOC toggle */}
      {hasToc && (
        <button onClick={() => setTocOpen(!tocOpen)} title={tocOpen ? "隱藏目錄" : "顯示目錄"} style={{
          position: "absolute", zIndex: 10, left: tocOpen ? 252 : 4, top: 8,
          width: 24, height: 24, borderRadius: 6, border: `1px solid ${AB}`,
          background: "var(--color-background-primary)", cursor: "pointer",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 12, color: "var(--color-text-tertiary)", fontFamily: "inherit",
          transition: "left 0.2s",
        }}>
          {tocOpen ? "◀" : "▶"}
        </button>
      )}

      {/* Content */}
      <div ref={contentRef} className="md-body" style={{
        flex: 1, padding: "28px 36px", maxHeight: maxH, overflowY: "auto",
        position: "relative",
      }} dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
}
