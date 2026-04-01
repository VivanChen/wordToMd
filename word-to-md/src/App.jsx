import { useState, useCallback, useRef, useEffect } from "react";
import * as mammoth from "mammoth";

/* ═══════════════════════════════════════════
   HTML → Markdown Conversion Engine
   ═══════════════════════════════════════════ */
function htmlToMarkdown(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  return convertNode(doc.body).replace(/\n{3,}/g, "\n\n").trim();
}

function convertNode(node) {
  if (node.nodeType === 3) return node.textContent;
  if (node.nodeType !== 1) return "";
  const tag = node.tagName.toLowerCase();
  const children = Array.from(node.childNodes).map(convertNode).join("");

  switch (tag) {
    case "h1": return `\n# ${children.trim()}\n\n`;
    case "h2": return `\n## ${children.trim()}\n\n`;
    case "h3": return `\n### ${children.trim()}\n\n`;
    case "h4": return `\n#### ${children.trim()}\n\n`;
    case "h5": return `\n##### ${children.trim()}\n\n`;
    case "h6": return `\n###### ${children.trim()}\n\n`;
    case "p": return `\n${children.trim()}\n\n`;
    case "br": return "\n";
    case "strong": case "b": return `**${children}**`;
    case "em": case "i": return `*${children}*`;
    case "u": return `<u>${children}</u>`;
    case "s": case "del": case "strike": return `~~${children}~~`;
    case "code": return `\`${children}\``;
    case "pre": return `\n\`\`\`\n${children.trim()}\n\`\`\`\n\n`;
    case "blockquote": return `\n> ${children.trim().replace(/\n/g, "\n> ")}\n\n`;
    case "a": {
      const href = node.getAttribute("href") || "";
      return `[${children}](${href})`;
    }
    case "img": {
      const src = node.getAttribute("src") || "";
      const alt = node.getAttribute("alt") || "image";
      return `![${alt}](${src})`;
    }
    case "ul": {
      const items = Array.from(node.children)
        .filter(c => c.tagName?.toLowerCase() === "li")
        .map(li => `- ${convertNode(li).trim()}`);
      return `\n${items.join("\n")}\n\n`;
    }
    case "ol": {
      const items = Array.from(node.children)
        .filter(c => c.tagName?.toLowerCase() === "li")
        .map((li, i) => `${i + 1}. ${convertNode(li).trim()}`);
      return `\n${items.join("\n")}\n\n`;
    }
    case "li": return children;
    case "table": {
      const rows = Array.from(node.querySelectorAll("tr"));
      if (rows.length === 0) return children;
      const tableData = rows.map(row =>
        Array.from(row.querySelectorAll("th, td")).map(cell => convertNode(cell).trim())
      );
      const colCount = Math.max(...tableData.map(r => r.length));
      const colWidths = Array.from({ length: colCount }, (_, i) =>
        Math.max(3, ...tableData.map(r => (r[i] || "").length))
      );
      const formatRow = row =>
        "| " + Array.from({ length: colCount }, (_, i) => (row[i] || "").padEnd(colWidths[i])).join(" | ") + " |";
      const separator = "| " + colWidths.map(w => "-".repeat(w)).join(" | ") + " |";
      const lines = [formatRow(tableData[0]), separator, ...tableData.slice(1).map(formatRow)];
      return `\n${lines.join("\n")}\n\n`;
    }
    case "hr": return "\n---\n\n";
    case "sup": return `<sup>${children}</sup>`;
    case "sub": return `<sub>${children}</sub>`;
    default: return children;
  }
}

/* ═══════════════════════════════════════════
   Markdown → HTML Renderer
   ═══════════════════════════════════════════ */
function renderMarkdown(md) {
  if (!md) return "";
  let html = md
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/^#{6}\s+(.+)$/gm, '<h6 style="font-size:14px;font-weight:600;margin:16px 0 8px">$1</h6>')
    .replace(/^#{5}\s+(.+)$/gm, '<h5 style="font-size:15px;font-weight:600;margin:16px 0 8px">$1</h5>')
    .replace(/^#{4}\s+(.+)$/gm, '<h4 style="font-size:16px;font-weight:600;margin:18px 0 8px">$1</h4>')
    .replace(/^#{3}\s+(.+)$/gm, '<h3 style="font-size:18px;font-weight:600;margin:20px 0 10px">$1</h3>')
    .replace(/^#{2}\s+(.+)$/gm, '<h2 style="font-size:21px;font-weight:700;margin:24px 0 12px">$1</h2>')
    .replace(/^#{1}\s+(.+)$/gm, '<h1 style="font-size:26px;font-weight:700;margin:28px 0 14px">$1</h1>')
    .replace(/^---$/gm, '<hr style="border:none;border-top:1px solid rgba(127,119,221,0.2);margin:20px 0"/>')
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, "<em>$1</em>")
    .replace(/~~(.+?)~~/g, "<del>$1</del>")
    .replace(/`([^`]+)`/g, '<code style="background:rgba(127,119,221,0.1);padding:2px 6px;border-radius:4px;font-size:0.88em;font-family:JetBrains Mono,monospace">$1</code>')
    .replace(/```\n?([\s\S]*?)```/g, '<pre style="background:rgba(127,119,221,0.06);padding:16px 20px;border-radius:10px;overflow-x:auto;font-size:0.85em;line-height:1.7;font-family:JetBrains Mono,monospace;border:1px solid rgba(127,119,221,0.12)"><code>$1</code></pre>')
    .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img alt="$1" src="$2" style="max-width:100%;border-radius:8px;margin:8px 0"/>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" style="color:#7F77DD;text-decoration:none;border-bottom:1px solid rgba(127,119,221,0.3)">$1</a>')
    .replace(/^\s*[-*]\s+(.+)$/gm, "<li>$1</li>")
    .replace(/(<li>.*<\/li>\n?)+/g, m => `<ul style="padding-left:24px;margin:8px 0">${m}</ul>`)
    .replace(/^\d+\.\s+(.+)$/gm, "<oli>$1</oli>")
    .replace(/(<oli>.*<\/oli>\n?)+/g, m => `<ol style="padding-left:24px;margin:8px 0">${m.replace(/oli>/g, "li>")}</ol>`)
    .replace(/^\|(.+)\|$/gm, (match) => {
      const cells = match.slice(1, -1).split("|").map(c => c.trim());
      if (cells.every(c => /^-+$/.test(c))) return "<!--sep-->";
      return `<tr>${cells.map(c => `<td style="border:1px solid rgba(127,119,221,0.15);padding:10px 14px;font-size:14px">${c}</td>`).join("")}</tr>`;
    })
    .replace(/((<tr>.*<\/tr>\n?)+)/g, m => {
      const cleaned = m.replace(/<!--sep-->\n?/g, "");
      return `<table style="border-collapse:collapse;margin:16px 0;width:100%;border-radius:8px;overflow:hidden">${cleaned}</table>`;
    })
    .replace(/^&gt;\s?(.+)$/gm, '<blockquote style="border-left:3px solid #7F77DD;padding:8px 20px;margin:12px 0;color:var(--color-text-secondary);background:rgba(127,119,221,0.04);border-radius:0 6px 6px 0">$1</blockquote>');

  html = html.split("\n").map(line => {
    const trimmed = line.trim();
    if (!trimmed) return "";
    if (/^<(h[1-6]|ul|ol|li|pre|table|tr|td|th|hr|blockquote|div|p)/.test(trimmed)) return trimmed;
    if (/^<!--/.test(trimmed)) return "";
    return `<p style="margin:6px 0;line-height:1.8">${trimmed}</p>`;
  }).join("\n");

  return html;
}

/* ═══════════════════════════════════════════
   GitHub API
   ═══════════════════════════════════════════ */
async function githubSaveFile(token, owner, repo, path, content, message) {
  const base = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json" };
  let sha;
  try {
    const existing = await fetch(base, { headers });
    if (existing.ok) sha = (await existing.json()).sha;
  } catch {}
  const body = {
    message,
    content: btoa(unescape(encodeURIComponent(content))),
    ...(sha ? { sha } : {}),
  };
  const res = await fetch(base, { method: "PUT", headers, body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`GitHub API ${res.status}: ${(await res.json()).message || "Unknown error"}`);
  return res.json();
}

async function githubListFiles(token, owner, repo, path = "") {
  const base = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const res = await fetch(base, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`GitHub API ${res.status}`);
  return res.json();
}

async function githubGetFile(token, owner, repo, path) {
  const base = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
  const res = await fetch(base, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`GitHub API ${res.status}`);
  const data = await res.json();
  return decodeURIComponent(escape(atob(data.content)));
}

/* ═══════════════════════════════════════════
   SVG Icons
   ═══════════════════════════════════════════ */
const I = {
  upload: <svg width="44" height="44" fill="none" viewBox="0 0 48 48"><path d="M24 32V16m0 0l-8 8m8-8l8 8" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"/><path d="M8 32v4a4 4 0 004 4h24a4 4 0 004-4v-4" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round"/></svg>,
  file: <svg width="18" height="18" fill="none" viewBox="0 0 20 20"><path d="M4 2h8l5 5v11a1 1 0 01-1 1H4a1 1 0 01-1-1V3a1 1 0 011-1z" stroke="currentColor" strokeWidth="1.5"/><path d="M12 2v5h5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>,
  gh: <svg width="18" height="18" fill="currentColor" viewBox="0 0 20 20"><path d="M10 1.667A8.333 8.333 0 001.667 10c0 3.683 2.388 6.808 5.702 7.913.416.075.569-.18.569-.402v-1.405c-2.32.504-2.81-1.12-2.81-1.12-.38-.963-.926-1.22-.926-1.22-.757-.517.057-.507.057-.507.837.059 1.278.86 1.278.86.744 1.274 1.95.907 2.426.693.075-.54.29-.907.528-1.115-1.851-.21-3.798-.925-3.798-4.12 0-.91.325-1.654.86-2.238-.087-.21-.373-1.06.081-2.208 0 0 .7-.224 2.292.855a7.99 7.99 0 014.166 0c1.59-1.08 2.289-.855 2.289-.855.455 1.148.17 1.998.083 2.208.537.584.858 1.328.858 2.238 0 3.204-1.95 3.907-3.808 4.113.3.258.568.768.568 1.548v2.294c0 .224.152.48.574.399A8.337 8.337 0 0010 1.667z"/></svg>,
  dl: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><path d="M9 3v8m0 0l-3-3m3 3l3-3M3 13v1a1 1 0 001 1h10a1 1 0 001-1v-1" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  arr: <svg width="18" height="18" fill="none" viewBox="0 0 20 20"><path d="M3 10h14m0 0l-4-4m4 4l-4 4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  eye: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><path d="M1.5 9s3-5.5 7.5-5.5S16.5 9 16.5 9s-3 5.5-7.5 5.5S1.5 9 1.5 9z" stroke="currentColor" strokeWidth="1.5"/><circle cx="9" cy="9" r="2.5" stroke="currentColor" strokeWidth="1.5"/></svg>,
  code: <svg width="16" height="16" fill="none" viewBox="0 0 18 18"><path d="M5.5 5L2 9l3.5 4M12.5 5L16 9l-3.5 4M10.5 3l-3 12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  gear: <svg width="18" height="18" fill="none" viewBox="0 0 20 20"><circle cx="10" cy="10" r="3" stroke="currentColor" strokeWidth="1.5"/><path d="M10 1v2m0 14v2M3.93 3.93l1.41 1.41m9.32 9.32l1.41 1.41M1 10h2m14 0h2M3.93 16.07l1.41-1.41m9.32-9.32l1.41-1.41" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>,
  folder: <svg width="18" height="18" fill="none" viewBox="0 0 20 20"><path d="M2 5a2 2 0 012-2h4l2 2h6a2 2 0 012 2v8a2 2 0 01-2 2H4a2 2 0 01-2-2V5z" stroke="currentColor" strokeWidth="1.5"/></svg>,
  ok: <svg width="14" height="14" fill="none" viewBox="0 0 16 16"><path d="M3 8.5l3 3 7-7" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/></svg>,
  x: <svg width="14" height="14" fill="none" viewBox="0 0 16 16"><path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" strokeWidth="2" strokeLinecap="round"/></svg>,
  warn: <svg width="14" height="14" fill="none" viewBox="0 0 16 16"><path d="M8 2l6.5 11H1.5L8 2z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round"/><path d="M8 7v2.5M8 11.5v.01" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/></svg>,
  doc: <svg width="18" height="18" fill="none" viewBox="0 0 20 20"><rect x="3" y="1" width="14" height="18" rx="2" stroke="currentColor" strokeWidth="1.5"/><path d="M7 6h6M7 9h6M7 12h4" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>,
  batch: <svg width="16" height="16" fill="none" viewBox="0 0 20 20"><rect x="4" y="3" width="12" height="14" rx="1.5" stroke="currentColor" strokeWidth="1.3"/><rect x="2" y="1" width="12" height="14" rx="1.5" stroke="currentColor" strokeWidth="1.3" opacity="0.4"/></svg>,
};

/* ═══════════════════════════════════════════
   Shared Styles
   ═══════════════════════════════════════════ */
const C = {
  accent: "#7F77DD",
  accentLt: "rgba(127,119,221,0.08)",
  accentBd: "rgba(127,119,221,0.2)",
  accentMd: "rgba(127,119,221,0.35)",
};

const btn = (bg, color, border) => ({
  display: "flex", alignItems: "center", gap: 7, padding: "9px 18px", borderRadius: 9,
  border: border || "none", background: bg, color, cursor: "pointer",
  fontSize: 13, fontWeight: 600, fontFamily: "inherit", transition: "all 0.15s",
});

/* ═══════════════════════════════════════════
   Main App
   ═══════════════════════════════════════════ */
export default function App() {
  const [page, setPage] = useState("convert");
  const [files, setFiles] = useState([]);         // uploaded File objects
  const [results, setResults] = useState([]);      // { file, markdown, html, status, error }
  const [activeIdx, setActiveIdx] = useState(0);
  const [previewMode, setPreviewMode] = useState("preview");
  const [converting, setConverting] = useState(false);
  const [convProgress, setConvProgress] = useState({ done: 0, total: 0, current: "" });
  const [toast, setToast] = useState(null);
  const [ghCfg, setGhCfg] = useState({ token: "", owner: "", repo: "", path: "docs" });
  const [ghFiles, setGhFiles] = useState([]);
  const [loadingGH, setLoadingGH] = useState(false);
  const [savingGH, setSavingGH] = useState(false);
  const [viewFile, setViewFile] = useState(null);
  const [viewContent, setViewContent] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  const showToast = useCallback((msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  }, []);

  // Load GitHub config
  useEffect(() => {
    try {
      const s = localStorage?.getItem?.("wmd_gh");
      if (s) setGhCfg(JSON.parse(s));
    } catch {}
  }, []);
  const saveGhCfg = (c) => { setGhCfg(c); try { localStorage?.setItem?.("wmd_gh", JSON.stringify(c)); } catch {} };

  /* ─── File type detection ─── */
  const isDoc = (f) => /\.doc$/i.test(f.name);
  const isDocx = (f) => /\.docx$/i.test(f.name);
  const isWord = (f) => isDoc(f) || isDocx(f);

  /* ─── Convert a single file ─── */
  const convertOne = async (file) => {
    if (isDocx(file)) {
      // Client-side mammoth.js
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer });
      const html = result.value;
      const md = htmlToMarkdown(html);
      const warnings = result.messages?.filter(m => m.type === "warning").length || 0;
      return { file, markdown: md, html, status: "ok", warnings };
    }

    if (isDoc(file)) {
      // Server-side via Netlify Function
      const base64 = await fileToBase64(file);
      const res = await fetch("/api/convert-doc", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ fileBase64: base64, filename: file.name }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }));
        throw new Error(err.error || `Server error ${res.status}`);
      }
      const data = await res.json();
      const md = htmlToMarkdown(data.html);
      return { file, markdown: md, html: data.html, status: "ok", warnings: data.warnings?.length || 0 };
    }

    throw new Error("Unsupported file type");
  };

  const fileToBase64 = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = () => reject(new Error("Failed to read file"));
    reader.readAsDataURL(file);
  });

  /* ─── Handle file upload ─── */
  const handleFiles = useCallback(async (fileList) => {
    const valid = Array.from(fileList).filter(isWord);
    if (valid.length === 0) {
      showToast("請選擇 .doc 或 .docx 檔案", "error");
      return;
    }
    setFiles(valid);
    setActiveIdx(0);
    setConverting(true);
    setConvProgress({ done: 0, total: valid.length, current: valid[0].name });

    const newResults = [];
    for (let i = 0; i < valid.length; i++) {
      setConvProgress({ done: i, total: valid.length, current: valid[i].name });
      try {
        const r = await convertOne(valid[i]);
        newResults.push(r);
      } catch (err) {
        newResults.push({ file: valid[i], markdown: "", html: "", status: "error", error: err.message });
      }
    }

    setResults(newResults);
    setConvProgress({ done: valid.length, total: valid.length, current: "" });
    setConverting(false);

    const ok = newResults.filter(r => r.status === "ok").length;
    const fail = newResults.filter(r => r.status === "error").length;
    if (fail > 0 && ok > 0) showToast(`${ok} 個成功，${fail} 個失敗`, "warning");
    else if (fail > 0) showToast(`轉換失敗：${newResults[0].error}`, "error");
    else showToast(`${ok} 個檔案轉換完成！`);
  }, []);

  /* ─── Current active result ─── */
  const active = results[activeIdx] || null;

  /* ─── Edit markdown ─── */
  const updateMarkdown = (val) => {
    setResults(prev => prev.map((r, i) => i === activeIdx ? { ...r, markdown: val } : r));
  };

  /* ─── Download ─── */
  const downloadMd = (r) => {
    if (!r?.markdown) return;
    const name = r.file.name.replace(/\.\w+$/, "") + ".md";
    const blob = new Blob([r.markdown], { type: "text/markdown;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    Object.assign(document.createElement("a"), { href: url, download: name }).click();
    URL.revokeObjectURL(url);
    showToast(`已下載 ${name}`);
  };

  const downloadAll = () => {
    results.filter(r => r.status === "ok").forEach(r => downloadMd(r));
  };

  /* ─── GitHub push ─── */
  const pushToGithub = async (r) => {
    const { token, owner, repo, path } = ghCfg;
    if (!token || !owner || !repo) { showToast("請先設定 GitHub", "error"); setPage("settings"); return; }
    setSavingGH(true);
    try {
      const name = r.file.name.replace(/\.\w+$/, "") + ".md";
      const fp = path ? `${path}/${name}` : name;
      await githubSaveFile(token, owner, repo, fp, r.markdown, `Add ${name}`);
      showToast(`已推送至 GitHub: ${fp}`);
    } catch (err) {
      showToast(`GitHub 推送失敗: ${err.message}`, "error");
    }
    setSavingGH(false);
  };

  const pushAllToGithub = async () => {
    const { token, owner, repo, path } = ghCfg;
    if (!token || !owner || !repo) { showToast("請先設定 GitHub", "error"); setPage("settings"); return; }
    setSavingGH(true);
    let ok = 0;
    for (const r of results.filter(r => r.status === "ok")) {
      try {
        const name = r.file.name.replace(/\.\w+$/, "") + ".md";
        const fp = path ? `${path}/${name}` : name;
        await githubSaveFile(token, owner, repo, fp, r.markdown, `Add ${name}`);
        ok++;
      } catch {}
    }
    showToast(`已推送 ${ok} 個檔案至 GitHub`);
    setSavingGH(false);
  };

  /* ─── GitHub file browser ─── */
  const loadGhFiles = async () => {
    const { token, owner, repo, path } = ghCfg;
    if (!token || !owner || !repo) { showToast("請先設定 GitHub", "error"); return; }
    setLoadingGH(true);
    try {
      const list = await githubListFiles(token, owner, repo, path);
      setGhFiles((Array.isArray(list) ? list : []).filter(f => f.name.endsWith(".md")));
    } catch (err) { showToast(`載入失敗: ${err.message}`, "error"); }
    setLoadingGH(false);
  };

  const openGhFile = async (f) => {
    const { token, owner, repo } = ghCfg;
    try {
      const content = await githubGetFile(token, owner, repo, f.path);
      setViewFile(f.name); setViewContent(content); setPage("reader");
    } catch (err) { showToast(`載入失敗: ${err.message}`, "error"); }
  };

  /* ─── Drag & Drop ─── */
  const onDrop = useCallback(e => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }, [handleFiles]);
  const onDragOver = e => { e.preventDefault(); setDragOver(true); };
  const onDragLeave = () => setDragOver(false);

  /* ─── Nav items ─── */
  const nav = [
    { id: "convert", label: "轉換", icon: I.arr },
    { id: "reader", label: "閱讀", icon: I.eye },
    { id: "github", label: "GitHub", icon: I.gh },
    { id: "settings", label: "設定", icon: I.gear },
  ];

  return (
    <div style={{ fontFamily: '"Noto Sans TC", system-ui, sans-serif', minHeight: "100vh", background: "var(--color-background-tertiary)", color: "var(--color-text-primary)" }}>

      {/* ── Toast ── */}
      {toast && (
        <div style={{
          position: "fixed", top: 20, right: 20, zIndex: 999, padding: "12px 22px",
          borderRadius: 10, fontSize: 13, fontWeight: 600, display: "flex", alignItems: "center", gap: 8,
          background: toast.type === "error" ? "#E24B4A" : toast.type === "warning" ? "#BA7517" : C.accent,
          color: "#fff", boxShadow: "0 8px 32px rgba(0,0,0,0.18)", animation: "slideIn .3s ease",
        }}>
          {toast.type === "error" ? I.x : toast.type === "warning" ? I.warn : I.ok}
          {toast.msg}
        </div>
      )}

      {/* ── Header ── */}
      <header style={{
        background: "var(--color-background-primary)",
        borderBottom: `1px solid ${C.accentBd}`,
        padding: "0 24px", display: "flex", alignItems: "center", justifyContent: "space-between",
        height: 58, position: "sticky", top: 0, zIndex: 50,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{
            width: 34, height: 34, borderRadius: 9,
            background: `linear-gradient(135deg, ${C.accent}, #AFA9EC)`,
            display: "flex", alignItems: "center", justifyContent: "center",
            color: "#fff", fontWeight: 800, fontSize: 15,
          }}>M</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 16, letterSpacing: -0.3 }}>Word → Markdown</div>
            <div style={{ fontSize: 10.5, color: "var(--color-text-tertiary)", marginTop: -1 }}>支援 .doc + .docx｜Netlify + GitHub</div>
          </div>
        </div>
        <nav style={{ display: "flex", gap: 3 }}>
          {nav.map(n => (
            <button key={n.id} onClick={() => setPage(n.id)} style={{
              ...btn("transparent", page === n.id ? C.accent : "var(--color-text-secondary)"),
              padding: "7px 13px", fontSize: 13, background: page === n.id ? C.accentLt : "transparent",
              borderRadius: 8,
            }}>
              {n.icon}{n.label}
            </button>
          ))}
        </nav>
      </header>

      {/* ── Main ── */}
      <main style={{ maxWidth: 1120, margin: "0 auto", padding: "24px 20px" }}>

        {/* ═══ CONVERT ═══ */}
        {page === "convert" && (<div>

          {/* Upload zone */}
          <div
            onDrop={onDrop} onDragOver={onDragOver} onDragLeave={onDragLeave}
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `2px dashed ${dragOver ? C.accent : C.accentBd}`,
              borderRadius: 14, padding: "36px 20px", textAlign: "center", cursor: "pointer",
              background: dragOver ? C.accentLt : "var(--color-background-primary)",
              transition: "all 0.2s", marginBottom: 20,
            }}
          >
            <input ref={fileInputRef} type="file" accept=".doc,.docx" multiple style={{ display: "none" }}
              onChange={e => handleFiles(e.target.files)} />
            <div style={{ color: C.accent, marginBottom: 10 }}>{I.upload}</div>
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 6 }}>拖曳 Word 檔案到這裡，或點擊選擇</div>
            <div style={{ fontSize: 13, color: "var(--color-text-tertiary)" }}>
              支援 <span style={{ color: C.accent, fontWeight: 600 }}>.docx</span> (瀏覽器端 mammoth.js) 和{" "}
              <span style={{ color: "#D85A30", fontWeight: 600 }}>.doc</span> (Netlify Function server-side 轉換)
            </div>
            <div style={{ display: "flex", gap: 8, justifyContent: "center", marginTop: 14 }}>
              <span style={{ ...tag("#7F77DD"), }}>.docx → client-side</span>
              <span style={{ ...tag("#D85A30"), }}>.doc → server-side</span>
            </div>
          </div>

          {/* Converting progress */}
          {converting && (
            <div style={{
              background: "var(--color-background-primary)", borderRadius: 12,
              border: `1px solid ${C.accentBd}`, padding: "28px 24px", marginBottom: 20,
              textAlign: "center",
            }}>
              <div style={{
                width: 44, height: 44, border: `3px solid ${C.accentBd}`, borderTopColor: C.accent,
                borderRadius: "50%", animation: "spin .7s linear infinite", margin: "0 auto 14px",
              }}/>
              <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 6 }}>
                轉換中... ({convProgress.done}/{convProgress.total})
              </div>
              <div style={{ fontSize: 13, color: "var(--color-text-tertiary)" }}>{convProgress.current}</div>
              <div style={{
                height: 4, borderRadius: 2, background: C.accentLt, marginTop: 14,
                overflow: "hidden",
              }}>
                <div style={{
                  height: "100%", borderRadius: 2, background: C.accent,
                  width: `${convProgress.total ? (convProgress.done / convProgress.total) * 100 : 0}%`,
                  transition: "width 0.3s",
                }}/>
              </div>
            </div>
          )}

          {/* Results area */}
          {results.length > 0 && !converting && (<div>

            {/* File tabs + batch actions */}
            <div style={{
              display: "flex", alignItems: "center", justifyContent: "space-between",
              marginBottom: 12, flexWrap: "wrap", gap: 8,
            }}>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", flex: 1 }}>
                {results.map((r, i) => (
                  <button key={i} onClick={() => setActiveIdx(i)} style={{
                    ...btn(
                      activeIdx === i ? C.accentLt : "var(--color-background-primary)",
                      activeIdx === i ? C.accent : "var(--color-text-primary)",
                      `1px solid ${activeIdx === i ? C.accent : C.accentBd}`,
                    ),
                    padding: "6px 12px", fontSize: 12, position: "relative",
                  }}>
                    {isDoc(r.file) ? I.doc : I.file}
                    <span style={{ maxWidth: 140, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {r.file.name}
                    </span>
                    {r.status === "error" && <span style={{ color: "#E24B4A" }}>{I.x}</span>}
                    {r.status === "ok" && <span style={{ color: "#1D9E75" }}>{I.ok}</span>}
                  </button>
                ))}
              </div>
              {results.filter(r => r.status === "ok").length > 1 && (
                <div style={{ display: "flex", gap: 6 }}>
                  <button onClick={downloadAll} style={btn("var(--color-background-primary)", "var(--color-text-primary)", `1px solid ${C.accentBd}`)}>
                    {I.batch} 全部下載
                  </button>
                  <button onClick={pushAllToGithub} disabled={savingGH} style={{ ...btn(C.accent, "#fff"), opacity: savingGH ? 0.5 : 1 }}>
                    {I.gh} 全部推送
                  </button>
                </div>
              )}
            </div>

            {/* Active file panel */}
            {active && active.status === "error" && (
              <div style={{
                background: "rgba(226,75,74,0.06)", border: "1px solid rgba(226,75,74,0.2)",
                borderRadius: 12, padding: "24px 28px",
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#E24B4A", fontWeight: 600, marginBottom: 8 }}>
                  {I.x} 轉換失敗
                </div>
                <div style={{ fontSize: 14, color: "var(--color-text-secondary)", lineHeight: 1.7 }}>
                  {active.error}
                </div>
                {isDoc(active.file) && (
                  <div style={{ fontSize: 13, color: "var(--color-text-tertiary)", marginTop: 12, padding: "12px 16px", background: "rgba(127,119,221,0.05)", borderRadius: 8 }}>
                    提示：.doc 需要 Netlify Function 支援。本地開發請先用 <code style={{ background: C.accentLt, padding: "1px 5px", borderRadius: 4 }}>netlify dev</code> 啟動。<br/>
                    或者先將 .doc 轉存為 .docx 後再上傳。
                  </div>
                )}
              </div>
            )}

            {active && active.status === "ok" && (<div>
              {/* View mode toggle + actions */}
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
                <div style={{ display: "flex", background: "var(--color-background-secondary)", borderRadius: 8, padding: 3 }}>
                  {[
                    { id: "preview", label: "預覽", icon: I.eye },
                    { id: "source", label: "Markdown", icon: I.code },
                    { id: "html", label: "HTML", icon: I.file },
                  ].map(m => (
                    <button key={m.id} onClick={() => setPreviewMode(m.id)} style={{
                      ...btn(
                        previewMode === m.id ? "var(--color-background-primary)" : "transparent",
                        previewMode === m.id ? C.accent : "var(--color-text-secondary)",
                      ),
                      padding: "5px 12px", fontSize: 12,
                      boxShadow: previewMode === m.id ? "0 1px 4px rgba(0,0,0,0.06)" : "none",
                      borderRadius: 6,
                    }}>
                      {m.icon}{m.label}
                    </button>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  <button onClick={() => downloadMd(active)} style={btn("var(--color-background-primary)", "var(--color-text-primary)", `1px solid ${C.accentBd}`)}>
                    {I.dl} 下載 .md
                  </button>
                  <button onClick={() => pushToGithub(active)} disabled={savingGH} style={{ ...btn(C.accent, "#fff"), opacity: savingGH ? 0.5 : 1 }}>
                    {I.gh} {savingGH ? "推送中..." : "推送 GitHub"}
                  </button>
                </div>
              </div>

              {/* Content */}
              <div style={{
                background: "var(--color-background-primary)", borderRadius: 12,
                border: `1px solid ${C.accentBd}`, overflow: "hidden",
              }}>
                {previewMode === "preview" && (
                  <div style={{ padding: "28px 36px", lineHeight: 1.8, fontSize: 15, maxHeight: 620, overflowY: "auto" }}
                    dangerouslySetInnerHTML={{ __html: renderMarkdown(active.markdown) }} />
                )}
                {previewMode === "source" && (
                  <textarea
                    value={active.markdown}
                    onChange={e => updateMarkdown(e.target.value)}
                    style={{
                      width: "100%", minHeight: 500, padding: "22px 28px", border: "none", resize: "vertical",
                      fontFamily: '"JetBrains Mono", monospace', fontSize: 13, lineHeight: 1.7,
                      background: "transparent", color: "var(--color-text-primary)", outline: "none",
                    }}
                  />
                )}
                {previewMode === "html" && (
                  <pre style={{
                    padding: "22px 28px", fontSize: 12, lineHeight: 1.6, overflowX: "auto", maxHeight: 500,
                    fontFamily: '"JetBrains Mono", monospace', color: "var(--color-text-secondary)", margin: 0,
                    whiteSpace: "pre-wrap", wordBreak: "break-all",
                  }}>
                    {active.html}
                  </pre>
                )}
              </div>

              {/* Stats bar */}
              <div style={{ display: "flex", gap: 14, marginTop: 10, fontSize: 12, color: "var(--color-text-tertiary)", flexWrap: "wrap" }}>
                <span>{active.markdown.length.toLocaleString()} 字元</span>
                <span>{active.markdown.split(/\s+/).filter(Boolean).length.toLocaleString()} 詞</span>
                <span>{active.markdown.split("\n").length} 行</span>
                <span>來源：{isDoc(active.file) ? "Netlify Function (.doc)" : "mammoth.js (.docx)"}</span>
                {active.warnings > 0 && <span style={{ color: "#BA7517" }}>{active.warnings} 個警告</span>}
              </div>
            </div>)}
          </div>)}
        </div>)}

        {/* ═══ READER ═══ */}
        {page === "reader" && (<div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
            <h2 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>Markdown 閱讀器</h2>
            <label style={{ ...btn("var(--color-background-primary)", "var(--color-text-secondary)", `1px solid ${C.accentBd}`), cursor: "pointer" }}>
              {I.upload} 開啟 .md 檔案
              <input type="file" accept=".md,.markdown,.txt" style={{ display: "none" }} onChange={async e => {
                const f = e.target.files[0]; if (!f) return;
                setViewFile(f.name); setViewContent(await f.text());
              }}/>
            </label>
          </div>

          {!viewFile && (
            <div style={{
              background: "var(--color-background-primary)", borderRadius: 12, padding: 60,
              textAlign: "center", border: `1px solid ${C.accentBd}`,
            }}>
              <div style={{ color: "var(--color-text-tertiary)", marginBottom: 12 }}>{I.folder}</div>
              <div style={{ fontWeight: 600, marginBottom: 8 }}>尚未開啟檔案</div>
              <div style={{ fontSize: 13, color: "var(--color-text-tertiary)" }}>
                上傳 .md 檔案，或從 GitHub 頁面載入
              </div>
            </div>
          )}

          {viewFile && (<div>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
              {I.file}
              <span style={{ fontWeight: 600, fontSize: 14 }}>{viewFile}</span>
              <button onClick={() => { setViewFile(null); setViewContent(""); }} style={{
                ...btn("transparent", "var(--color-text-secondary)", `1px solid ${C.accentBd}`),
                marginLeft: "auto", padding: "4px 12px", fontSize: 12,
              }}>關閉</button>
              <button onClick={() => {
                const blob = new Blob([viewContent], { type: "text/markdown" });
                const url = URL.createObjectURL(blob);
                Object.assign(document.createElement("a"), { href: url, download: viewFile }).click();
                URL.revokeObjectURL(url);
                showToast(`已下載 ${viewFile}`);
              }} style={{ ...btn(C.accent, "#fff"), padding: "4px 14px", fontSize: 12 }}>
                {I.dl} 下載
              </button>
            </div>
            <div style={{
              background: "var(--color-background-primary)", borderRadius: 12,
              border: `1px solid ${C.accentBd}`, padding: "28px 36px", lineHeight: 1.8, fontSize: 15,
              maxHeight: 700, overflowY: "auto",
            }} dangerouslySetInnerHTML={{ __html: renderMarkdown(viewContent) }}/>
          </div>)}
        </div>)}

        {/* ═══ GITHUB ═══ */}
        {page === "github" && (<div>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
            <h2 style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>GitHub 檔案庫</h2>
            <button onClick={loadGhFiles} disabled={loadingGH} style={{ ...btn(C.accent, "#fff"), opacity: loadingGH ? 0.5 : 1 }}>
              {loadingGH ? "載入中..." : "重新整理"}
            </button>
          </div>

          {!ghCfg.token ? (
            <div style={{
              background: "var(--color-background-primary)", borderRadius: 12, padding: 40,
              textAlign: "center", border: `1px solid ${C.accentBd}`,
            }}>
              <div style={{ fontWeight: 600, marginBottom: 8 }}>尚未設定 GitHub</div>
              <div style={{ fontSize: 13, color: "var(--color-text-tertiary)", marginBottom: 16 }}>前往設定頁面填入 Token 及 Repo 資訊</div>
              <button onClick={() => setPage("settings")} style={btn(C.accent, "#fff")}>前往設定</button>
            </div>
          ) : ghFiles.length === 0 && !loadingGH ? (
            <div style={{
              background: "var(--color-background-primary)", borderRadius: 12, padding: 40,
              textAlign: "center", border: `1px solid ${C.accentBd}`,
            }}>
              {I.folder}
              <div style={{ fontWeight: 600, margin: "10px 0 6px" }}>沒有找到 .md 檔案</div>
              <div style={{ fontSize: 13, color: "var(--color-text-tertiary)" }}>
                點擊「重新整理」從 {ghCfg.owner}/{ghCfg.repo}/{ghCfg.path} 載入
              </div>
            </div>
          ) : (
            <div style={{ display: "grid", gap: 6 }}>
              {ghFiles.map(f => (
                <button key={f.sha} onClick={() => openGhFile(f)} style={{
                  display: "flex", alignItems: "center", gap: 12, padding: "13px 18px", borderRadius: 10,
                  border: `1px solid ${C.accentBd}`, background: "var(--color-background-primary)",
                  cursor: "pointer", textAlign: "left", width: "100%", fontFamily: "inherit",
                  transition: "border-color 0.15s",
                }}>
                  <div style={{ color: C.accent }}>{I.file}</div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{f.name}</div>
                    <div style={{ fontSize: 12, color: "var(--color-text-tertiary)" }}>{(f.size / 1024).toFixed(1)} KB · {f.path}</div>
                  </div>
                  <span style={{ color: C.accent, fontSize: 12, fontWeight: 600 }}>開啟</span>
                </button>
              ))}
            </div>
          )}
        </div>)}

        {/* ═══ SETTINGS ═══ */}
        {page === "settings" && (<div style={{ maxWidth: 580 }}>
          <h2 style={{ fontSize: 20, fontWeight: 700, marginBottom: 20 }}>設定</h2>

          {/* GitHub config */}
          <div style={{
            background: "var(--color-background-primary)", borderRadius: 12,
            border: `1px solid ${C.accentBd}`, padding: "24px 28px",
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 18, fontWeight: 700, fontSize: 15 }}>
              {I.gh} GitHub 連線設定
            </div>
            {[
              { key: "token", label: "Personal Access Token", type: "password", ph: "ghp_xxxxxxxxxxxx" },
              { key: "owner", label: "Owner（用戶名或組織）", ph: "your-username" },
              { key: "repo", label: "Repository 名稱", ph: "my-docs" },
              { key: "path", label: "資料夾路徑（選填）", ph: "docs" },
            ].map(f => (
              <div key={f.key} style={{ marginBottom: 14 }}>
                <label style={{ display: "block", fontSize: 13, fontWeight: 500, marginBottom: 5, color: "var(--color-text-secondary)" }}>{f.label}</label>
                <input
                  type={f.type || "text"} value={ghCfg[f.key]} placeholder={f.ph}
                  onChange={e => saveGhCfg({ ...ghCfg, [f.key]: e.target.value })}
                  style={{
                    width: "100%", padding: "10px 14px", borderRadius: 8, fontSize: 14,
                    border: `1px solid ${C.accentBd}`, background: "var(--color-background-secondary)",
                    color: "var(--color-text-primary)", outline: "none", boxSizing: "border-box",
                    fontFamily: "inherit",
                  }}
                  onFocus={e => e.target.style.borderColor = C.accent}
                  onBlur={e => e.target.style.borderColor = C.accentBd}
                />
              </div>
            ))}
            <div style={{ fontSize: 12, color: "var(--color-text-tertiary)", lineHeight: 1.7, marginTop: 4 }}>
              Token 需要 <code style={{ background: C.accentLt, padding: "1px 5px", borderRadius: 4 }}>repo</code> 權限。<br/>
              前往 GitHub → Settings → Developer settings → Personal access tokens 建立。
            </div>
          </div>

          {/* Architecture info */}
          <div style={{
            background: "var(--color-background-primary)", borderRadius: 12,
            border: `1px solid ${C.accentBd}`, padding: "24px 28px", marginTop: 14,
          }}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14 }}>系統架構</div>
            <div style={{ fontSize: 13, color: "var(--color-text-secondary)", lineHeight: 2 }}>
              {archItem("#7F77DD", ".docx 轉換", "mammoth.js 瀏覽器端直接處理")}
              {archItem("#D85A30", ".doc 轉換", "Netlify Function + word-extractor server-side")}
              {archItem("#1D9E75", "MD 儲存", "GitHub API push to repository")}
              {archItem("#BA7517", "部署", "Netlify 靜態 SPA + Serverless Functions")}
            </div>
          </div>

          {/* Deploy guide */}
          <div style={{
            background: "var(--color-background-primary)", borderRadius: 12,
            border: `1px solid ${C.accentBd}`, padding: "24px 28px", marginTop: 14,
          }}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14 }}>部署到 Netlify</div>
            <div style={{ fontSize: 13, color: "var(--color-text-secondary)", lineHeight: 2 }}>
              <div>1. 推送專案到 GitHub Repository</div>
              <div>2. 前往 Netlify → New site from Git</div>
              <div>3. Build command: <code style={{ background: C.accentLt, padding: "1px 5px", borderRadius: 4 }}>npm run build</code></div>
              <div>4. Publish directory: <code style={{ background: C.accentLt, padding: "1px 5px", borderRadius: 4 }}>dist</code></div>
              <div>5. Functions directory: <code style={{ background: C.accentLt, padding: "1px 5px", borderRadius: 4 }}>netlify/functions</code>（已在 netlify.toml 設定）</div>
              <div style={{ marginTop: 8 }}>本地開發：<code style={{ background: C.accentLt, padding: "1px 5px", borderRadius: 4 }}>npx netlify dev</code></div>
            </div>
          </div>
        </div>)}
      </main>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes slideIn { from { transform: translateX(100px); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
        * { box-sizing: border-box; margin: 0; }
        button, input, textarea { font-family: inherit; }
        ::selection { background: ${C.accentLt}; }
        ::-webkit-scrollbar { width: 5px; }
        ::-webkit-scrollbar-thumb { background: ${C.accentBd}; border-radius: 3px; }
        button:hover { filter: brightness(1.05); }
      `}</style>
    </div>
  );
}

/* helper: colored tag for upload zone */
function tag(color) {
  return {
    fontSize: 11, fontWeight: 600, padding: "3px 10px", borderRadius: 6,
    background: `${color}12`, color, border: `1px solid ${color}30`,
  };
}

/* helper: architecture item */
function archItem(color, title, desc) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
      <span style={{ width: 8, height: 8, borderRadius: 4, background: color, flexShrink: 0 }}/>
      <span><strong>{title}</strong> — {desc}</span>
    </div>
  );
}
