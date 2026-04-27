import { useState, useRef, useEffect } from "react";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import {
  Table, TableBody, TableCell, TableContainer, TableHead, TableRow,
  Paper, IconButton, Tooltip, Box, Typography, Chip, FormControl,
  InputLabel, Select, MenuItem, Button, Dialog, DialogTitle,
  DialogContent, DialogActions, TextField, LinearProgress, Alert,
  CircularProgress,
} from "@mui/material";
import DownloadIcon from "@mui/icons-material/Download";
import SaveIcon from "@mui/icons-material/Save";
import CheckCircleIcon from "@mui/icons-material/CheckCircle";
import ErrorIcon from "@mui/icons-material/Error";
import {
  Document, Packer, Paragraph, TextRun,
  Table as DocxTable, TableRow as DocxTableRow, TableCell as DocxTableCell,
  WidthType, HeadingLevel, BorderStyle, ShadingType,
} from "docx";
import { saveAs } from "file-saver";


const API = import.meta.env.VITE_API_URL;
const CMS_URL = import.meta.env.VITE_CMS_URL;
const TOKEN = import.meta.env.VITE_TOKEN;

const SAMPLE_QUESTIONS = [
  "Which companies have entered or expanded into the United States market in 2025, and what are the key details of their expansion (industry, investment size, location, and strategic intent)?",
  "Which non-US companies entered the US market in 2025?",
  "Which existing companies expanded their US operations in 2025?",
  "What industries are seeing the highest number of new entrants?",
  "Provide a structured analysis of companies entering or expanding into the US market in 2025, including company profile, expansion type, investment size, location, strategic rationale, competitive impact, and available contact information for business outreach.",
];

const INDUSTRIES = [
  { value: "", label: "All Industries" },
  { value: "Pharmaceuticals", label: "Pharmaceuticals / Biotech" },
  { value: "Food and Beverage", label: "Food & Beverage" },
  { value: "Technology", label: "Technology" },
  { value: "Automotive", label: "Automotive / EV" },
  { value: "Renewable Energy", label: "Renewable Energy" },
  { value: "Manufacturing", label: "Manufacturing" },
  { value: "Consumer Goods", label: "Consumer Goods" },
  { value: "Healthcare", label: "Healthcare" },
  { value: "Chemicals", label: "Chemicals" },
];

// ─────────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────────
const stripBold = (str) =>
  String(str).replace(/\*\*(.*?)\*\*/g, "$1").replace(/__(.*?)__/g, "$1").trim();

const stripMarkdown = (str) =>
  stripBold(str).replace(/^#+\s*/, "").replace(/`/g, "").trim();

const esc = (v) => String(v ?? "").replace(/'/g, "''");

const confidenceColor = (val = "") => {
  const v = val.toLowerCase();
  if (v.includes("confirmed")) return { bg: "#d1fae5", color: "#065f46" };
  if (v.includes("high")) return { bg: "#dbeafe", color: "#1e40af" };
  if (v.includes("medium")) return { bg: "#fef3c7", color: "#92400e" };
  return { bg: "#fee2e2", color: "#991b1b" };
};

const inferTableTitle = (headers) => {
  const hs = headers.map(h => h.toLowerCase()).join(" ");
  if (hs.includes("channel") || hs.includes("email") || hs.includes("phone")) return "Contact Details";
  if (hs.includes("field") && hs.includes("detail")) return "Expansion Data";
  if (hs.includes("company") || hs.includes("date") || hs.includes("confidence")) return "Market Expansion Data";
  return "Table";
};

// ─────────────────────────────────────────────
// MAP a lead record → CMS field names
// Adjust keys here if the Directus collection uses different field names
// ─────────────────────────────────────────────
function leadToCmsPayload(lead) {
  return {
    company: lead.company !== "—" ? lead.company : null,
    country: lead.country !== "—" ? lead.country : null,
    industry: lead.industry !== "—" ? lead.industry : null,
    date: lead.date !== "—" ? lead.date : null,
    source: lead.source !== "—" ? lead.source : null,
    expansion_type: lead.expansion_type !== "—" ? lead.expansion_type : null,
    location: lead.location !== "—" ? lead.location : null,
    investment_amount: lead.investment_amount !== "—" ? lead.investment_amount : null,
    confidence: lead.confidence !== "—" ? lead.confidence : null,
    investment_status: lead.status !== "—" ? lead.status : null,
    website: lead.website !== "—" ? lead.website : null,
    email: lead.email !== "—" ? lead.email : null,
    phone: lead.phone !== "—" ? lead.phone : null,
    twitter: lead.twitter !== "—" ? lead.twitter : null,
    linkedin: lead.linkedin !== "—" ? lead.linkedin : null,
    summary: lead.summary !== "—" ? lead.summary : null,
  };
}

// ─────────────────────────────────────────────
// PUSH A SINGLE LEAD TO CMS
// ─────────────────────────────────────────────
// ✅ Corrected version - uses TOKEN from .env automatically
async function pushLeadToCms(lead) {
  const res = await fetch(CMS_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${TOKEN}`,   // ← Uses the global TOKEN
    },
    body: JSON.stringify(leadToCmsPayload(lead)),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => res.statusText);
    throw new Error(`HTTP ${res.status}: ${errText}`);
  }
  return await res.json();
}

// ─────────────────────────────────────────────
// SAVE LEADS DIALOG
// Shows token input → progress bar → per-lead results
// ─────────────────────────────────────────────
function SaveLeadsDialog({ open, onClose, leads, label }) {
  const [phase, setPhase] = useState("pushing");
  const [results, setResults] = useState([]);
  const [current, setCurrent] = useState(0);

  useEffect(() => {
    if (open) {
      setPhase("pushing");
      setResults([]);
      setCurrent(0);
      startPushing();
    }
  }, [open]);

  const startPushing = async () => {
    const res = [];
    for (let i = 0; i < leads.length; i++) {
      setCurrent(i + 1);
      try {
        await pushLeadToCms(leads[i]);   // ← uses TOKEN from .env automatically
        res.push({ company: leads[i].company, ok: true });
      } catch (e) {
        res.push({ company: leads[i].company, ok: false, error: e.message });
      }
      setResults([...res]);
      await new Promise(r => setTimeout(r, 120));
    }
    setPhase("done");
  };

  const successCount = results.filter(r => r.ok).length;
  const failCount = results.filter(r => !r.ok).length;

  return (
    <Dialog open={open} onClose={phase === "pushing" ? undefined : onClose} maxWidth="sm" fullWidth
      PaperProps={{ sx: { borderRadius: "16px", border: "1px solid rgba(201,168,76,0.25)" } }}>

      <DialogTitle sx={{ background: "linear-gradient(135deg, #1e3a5f 0%, #2a4f7c 100%)", color: "#c9a84c", fontFamily: "'Playfair Display', serif", fontSize: "16px", py: 2 }}>
        💾 Saving Leads to Database
        <Typography sx={{ fontSize: "11px", color: "#94a3b8", mt: 0.3 }}>
          {label} · {leads.length} record{leads.length !== 1 ? "s" : ""} → {CMS_URL}
        </Typography>
      </DialogTitle>

      <DialogContent sx={{ pt: 3 }}>
        {phase === "pushing" && (
          <Box>
            <Box sx={{ display: "flex", alignItems: "center", gap: 1.5, mb: 2 }}>
              <CircularProgress size={18} sx={{ color: "#c9a84c" }} />
              <Typography sx={{ fontSize: "13px", color: "#374151" }}>
                Saving lead {current} of {leads.length}…
              </Typography>
            </Box>
            <LinearProgress variant="determinate" value={(current / leads.length) * 100}
              sx={{ borderRadius: 4, height: 6, mb: 2, backgroundColor: "#f1f5f9", "& .MuiLinearProgress-bar": { backgroundColor: "#c9a84c" } }} />
            <ResultsList results={results} />
          </Box>
        )}

        {phase === "done" && (
          <Box>
            <Alert severity={failCount === 0 ? "success" : "warning"} sx={{ mb: 2, borderRadius: "10px" }}>
              {failCount === 0
                ? `All ${successCount} lead${successCount !== 1 ? "s" : ""} saved successfully! ✅`
                : `${successCount} succeeded · ${failCount} failed`}
            </Alert>
            <ResultsList results={results} />
          </Box>
        )}
      </DialogContent>

      <DialogActions sx={{ px: 3, pb: 2.5 }}>
        {phase !== "pushing" && (
          <Button onClick={onClose} variant="contained" size="small"
            sx={{ borderRadius: "10px", background: "#1e3a5f", color: "#c9a84c", textTransform: "none" }}>
            Close
          </Button>
        )}
      </DialogActions>
    </Dialog>
  );
}

// Small helper: renders the per-lead result list
function ResultsList({ results }) {
  if (!results.length) return null;
  return (
    <Box sx={{ maxHeight: 280, overflowY: "auto", display: "flex", flexDirection: "column", gap: 0.5 }}>
      {results.map((r, i) => (
        <Box key={i} sx={{ display: "flex", alignItems: "flex-start", gap: 1, px: 1.5, py: 0.8, borderRadius: "8px", background: r.ok ? "#f0fdf4" : "#fef2f2", border: `1px solid ${r.ok ? "#86efac" : "#fca5a5"}` }}>
          {r.ok
            ? <CheckCircleIcon sx={{ fontSize: 15, color: "#16a34a", mt: 0.15, flexShrink: 0 }} />
            : <ErrorIcon sx={{ fontSize: 15, color: "#dc2626", mt: 0.15, flexShrink: 0 }} />}
          <Box>
            <Typography sx={{ fontSize: "12px", fontWeight: 600, color: r.ok ? "#15803d" : "#dc2626" }}>
              {r.company}
            </Typography>
            {!r.ok && (
              <Typography sx={{ fontSize: "10px", color: "#9ca3af", mt: 0.2 }}>{r.error}</Typography>
            )}
          </Box>
        </Box>
      ))}
    </Box>
  );
}

// ─────────────────────────────────────────────
// COMPREHENSIVE SUMMARY DATA BUILDER
// ─────────────────────────────────────────────
function buildSummaryData(sections) {
  return sections.map(section => {
    const rec = {
      company: section.name || "—", country: section.country || "—", industry: section.industry || "—",
      date: "—", source: "—", expansion_type: "—", location: "—", investment_amount: "—",
      confidence: "—", status: "—", website: "—", email: "—", phone: "—", twitter: "—",
      linkedin: "—", summary: "—",
    };

    for (const block of section.blocks) {
      if (block.type === "text") {
        const text = block.content;
        if (rec.summary === "—") {
          const firstLine = text.split("\n").map(l => stripMarkdown(l)).find(l => l.length > 30);
          if (firstLine) rec.summary = firstLine;
        }
        if (rec.status === "—") { const m = text.match(/\b(Planning|Announced|In Progress|Launched)\b/i); if (m) rec.status = m[1]; }
        if (rec.confidence === "—") { const m = text.match(/confidence[:\s]+([A-Za-z]+)/i); if (m) rec.confidence = m[1].trim(); }
        if (rec.source === "—") { const m = text.match(/source[:\s]+([^\n|]+)/i); if (m) rec.source = m[1].trim(); }
        continue;
      }

      const lh = block.headers.map(h => h.toLowerCase().trim());
      const isKV = block.headers.length === 2 && (lh[0].includes("field") || lh[0].includes("channel"));
      const isContact = lh.some(h => h.includes("channel") || h.includes("email") || h.includes("phone") || h.includes("website") || (h.includes("twitter") && !h.includes("company")));

      if (isKV) {
        for (const row of block.rows) {
          const key = (row[0] || "").toLowerCase().trim();
          const val = (row[1] || "").trim();
          if (!val || val === "—" || val === "N/A") continue;
          if (key.includes("date")) rec.date = val;
          if (key.includes("source")) rec.source = val;
          if (key.includes("expansion") || key.includes("type")) rec.expansion_type = val;
          if (key.includes("location") || key.includes("city")) rec.location = val;
          if (key.includes("amount") || key.includes("investment")) rec.investment_amount = val;
          if (key.includes("confidence")) rec.confidence = val;
          if (key.includes("status")) rec.status = val;
          if (key.includes("website") || key.includes("url")) rec.website = val;
          if (key.includes("email")) rec.email = val;
          if (key.includes("phone")) rec.phone = val;
          if (key.includes("twitter") || key.includes("x handle")) rec.twitter = val;
          if (key.includes("linkedin")) rec.linkedin = val;
        }
      } else if (isContact) {
        for (const row of block.rows) {
          const ch = (row[0] || "").toLowerCase().trim();
          const val = (row[1] || "").trim();
          if (!val || val === "N/A" || val === "—") continue;
          if (ch.includes("website") || ch.includes("url")) rec.website = val;
          if (ch.includes("email")) rec.email = val;
          if (ch.includes("phone")) rec.phone = val;
          if (ch.includes("twitter") || ch.includes("x")) rec.twitter = val;
          if (ch.includes("linkedin")) rec.linkedin = val;
        }
      } else {
        for (const row of block.rows) {
          lh.forEach((h, i) => {
            const val = (row[i] || "").trim();
            if (!val || val === "—") return;
            if (h.includes("date")) rec.date = val;
            if (h.includes("source")) rec.source = val;
            if (h.includes("expansion") || h.includes("type")) rec.expansion_type = val;
            if (h.includes("location") || h.includes("city")) rec.location = val;
            if (h.includes("amount") || h.includes("investment")) rec.investment_amount = val;
            if (h.includes("confidence")) rec.confidence = val;
            if (h.includes("company")) rec.company = val;
            if (h.includes("country")) rec.country = val;
            if (h.includes("industry") || h.includes("sector")) rec.industry = val;
          });
        }
      }
    }
    return rec;
  });
}

// ─────────────────────────────────────────────
// SQL GENERATOR (kept for reference / export)
// ─────────────────────────────────────────────
// function generateSQL(leads) {
//   const ddl = `-- ════════════════════════════════════════════
// -- USA Market Expansion Leads
// -- Generated: ${new Date().toISOString()}
// -- ════════════════════════════════════════════

// CREATE TABLE IF NOT EXISTS expansion_leads (
//   id                 SERIAL PRIMARY KEY,
//   company            TEXT, country TEXT, industry TEXT, date TEXT, source TEXT,
//   expansion_type     TEXT, location TEXT, investment_amount TEXT, confidence TEXT,
//   status             TEXT, website TEXT, email TEXT, phone TEXT, twitter TEXT,
//   linkedin           TEXT, summary TEXT, created_at TIMESTAMP DEFAULT NOW()
// );

// `;
//   const inserts = leads.map((r, idx) =>
//     `-- Lead ${idx + 1}: ${r.company}\n` +
//     `INSERT INTO expansion_leads (company,country,industry,date,source,expansion_type,location,investment_amount,confidence,status,website,email,phone,twitter,linkedin,summary) VALUES (\n` +
//     `  '${esc(r.company)}','${esc(r.country)}','${esc(r.industry)}','${esc(r.date)}',\n` +
//     `  '${esc(r.source)}','${esc(r.expansion_type)}','${esc(r.location)}','${esc(r.investment_amount)}',\n` +
//     `  '${esc(r.confidence)}','${esc(r.status)}','${esc(r.website)}','${esc(r.email)}',\n` +
//     `  '${esc(r.phone)}','${esc(r.twitter)}','${esc(r.linkedin)}','${esc(r.summary)}'\n);\n`
//   ).join("\n");
//   return ddl + inserts;
// }

// ─────────────────────────────────────────────
// SINGLE-TABLE DOCX DOWNLOAD
// ─────────────────────────────────────────────
async function downloadSingleTableDocx(headers, rows, title = "Table") {
  try {
    const MARGIN = 720;
    const PAGE_WIDTH = 15840; // A4 Landscape width in DXA
    const EFF_W = PAGE_WIDTH - MARGIN * 2;

    // Better column width calculation for wide tables
    const minColWidth = 800; // minimum width per column
    const availableWidth = EFF_W;

    // Calculate dynamic widths - give more space to important columns
    const importantCols = ["Company", "Summary", "Expansion Details", "Contact", "Investment"];

    const colWidths = headers.map((header, index) => {
      const isImportant = importantCols.some(imp =>
        header.toLowerCase().includes(imp.toLowerCase())
      );

      if (isImportant) {
        return Math.floor(availableWidth * 0.18); // 18% for important columns
      } else {
        return Math.floor(availableWidth / headers.length);
      }
    });

    // Ensure total doesn't exceed available width
    const totalWidth = colWidths.reduce((sum, w) => sum + w, 0);
    if (totalWidth > availableWidth) {
      const scale = availableWidth / totalWidth;
      colWidths.forEach((w, i) => colWidths[i] = Math.floor(w * scale));
    }

    const border = { style: BorderStyle.SINGLE, size: 1, color: "D1D5DB" };
    const borders = { top: border, bottom: border, left: border, right: border };
    const margins = { top: 40, bottom: 40, left: 60, right: 60 };

    const hRow = new DocxTableRow({
      tableHeader: true,
      children: headers.map((h, i) =>
        new DocxTableCell({
          borders,
          margins,
          width: { size: colWidths[i], type: WidthType.DXA },
          shading: { fill: "1E3A5F", type: ShadingType.CLEAR },
          children: [new Paragraph({
            children: [new TextRun({ text: String(h), bold: true, color: "C9A84C", size: 14 })],
          })],
        })
      ),
    });

    const dRows = rows.map((row, ri) =>
      new DocxTableRow({
        children: row.map((cell, i) =>
          new DocxTableCell({
            borders,
            margins,
            width: { size: colWidths[i], type: WidthType.DXA },
            shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F8FAFC", type: ShadingType.CLEAR },
            children: [new Paragraph({
              children: [new TextRun({
                text: String(cell ?? "—"),
                size: 11,
                font: "Arial"
              })],
            })],
          })
        ),
      })
    );

    const table = new DocxTable({
      width: { size: availableWidth, type: WidthType.DXA },
      columnWidths: colWidths,
      rows: [hRow, ...dRows],
    });

    const doc = new Document({
      styles: { default: { document: { run: { font: "Arial", size: 18 } } } },
      sections: [{
        properties: {
          page: {
            size: { width: PAGE_WIDTH, height: 12240 }, // Landscape
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            spacing: { after: 160 },
            children: [new TextRun({ text: title, bold: true, size: 24, color: "1E3A5F" })],
          }),
          new Paragraph({
            spacing: { after: 200 },
            children: [new TextRun({
              text: `Generated: ${new Date().toLocaleString()}`,
              size: 14, color: "6B7280", italics: true,
            })],
          }),
          table,
          new Paragraph({
            spacing: { before: 300 },
            children: [new TextRun({
              text: "All information sourced from verified public reports. Facts only.",
              size: 12, color: "9CA3AF", italics: true,
            })],
          }),
        ],
      }],
    });

    const slug = title.replace(/\s+/g, "_").replace(/[^a-zA-Z0-9_]/g, "");
    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${slug}_${new Date().toISOString().slice(0, 10)}.docx`);

  } catch (err) {
    console.error("DOCX download error:", err);
    alert("Download failed: " + err.message);
  }
}

// ─────────────────────────────────────────────
// MARKDOWN TABLE PARSER
// ─────────────────────────────────────────────
function parseOneTable(lines) {
  const dataLines = lines.filter(l => !l.trim().match(/^\|[\s\-:|]+\|$/));
  if (dataLines.length < 2) return null;
  const parseRow = (row) => row.replace(/^\||\|$/g, "").split("|").map(c => stripBold(c.trim()));
  const headers = parseRow(dataLines[0]);
  const rows = dataLines.slice(1).map(parseRow);
  return { headers, rows };
}

function parseResponseBlocks(content) {
  if (!content) return [{ type: "text", content: "" }];
  const lines = content.split("\n"); const blocks = []; let i = 0;
  while (i < lines.length) {
    if (lines[i].trim().startsWith("|")) {
      const tableLines = [];
      while (i < lines.length && (lines[i].trim().startsWith("|") || lines[i].trim() === "")) {
        if (lines[i].trim().startsWith("|")) tableLines.push(lines[i]); i++;
      }
      const parsed = parseOneTable(tableLines);
      if (parsed) blocks.push({ type: "table", ...parsed });
    } else {
      const textLines = [];
      while (i < lines.length && !lines[i].trim().startsWith("|")) { textLines.push(lines[i]); i++; }
      const text = textLines.join("\n").trim();
      if (text) blocks.push({ type: "text", content: text });
    }
  }
  return blocks.length ? blocks : [{ type: "text", content: content }];
}

const COMPANY_HEADING_RE = /^#{1,4}\s+([^—\n]{2,60}?)\s*—\s*([^—\n]{2,60}?)\s*—\s*([^—\n]{2,80}?)\s*$/;
const OVERALL_HEADING_RE = /overall analysis/i;

function parseCompanySections(blocks) {
  const sections = []; let preamble = []; let analysisBlocks = [];
  let current = null; let inAnalysis = false;
  for (const block of blocks) {
    if (block.type === "text") {
      const lines = block.content.split("\n"); let consumed = false;
      for (let li = 0; li < lines.length; li++) {
        const line = lines[li].trim();
        if (OVERALL_HEADING_RE.test(line)) { inAnalysis = true; const rem = lines.slice(li + 1).join("\n").trim(); if (rem) analysisBlocks.push({ type: "text", content: rem }); consumed = true; break; }
        const m = line.match(COMPANY_HEADING_RE);
        if (m) { current = { heading: stripMarkdown(line), name: m[1].trim(), country: m[2].trim(), industry: m[3].trim(), blocks: [] }; sections.push(current); const rem = lines.slice(li + 1).join("\n").trim(); if (rem) current.blocks.push({ type: "text", content: rem }); consumed = true; break; }
      }
      if (!consumed) { if (inAnalysis) analysisBlocks.push(block); else if (current) current.blocks.push(block); else preamble.push(block); }
    } else { if (inAnalysis) analysisBlocks.push(block); else if (current) current.blocks.push(block); else preamble.push(block); }
  }
  return { preamble, sections, analysisBlocks };
}

function tableType(headers) {
  const hs = headers.map(h => h.toLowerCase());
  return hs.some(h => h.includes("channel") || h.includes("website") || h.includes("email") || h.includes("phone") || (h.includes("twitter") && !h.includes("company"))) ? "contact" : "expansion";
}

// ─────────────────────────────────────────────
// TABLE WRAPPER with per-table download icon
// ─────────────────────────────────────────────
function TableWithDownload({ headers, rows, title, children }) {
  return (
    <Box sx={{ mb: 1.5 }}>
      <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", px: 1, py: 0.3, background: "rgba(30,58,95,0.04)", borderRadius: "6px 6px 0 0", border: "1px solid rgba(201,168,76,0.15)", borderBottom: "none" }}>
        <Typography sx={{ fontSize: "10px", color: "#6b7280", letterSpacing: "0.06em", textTransform: "uppercase", fontWeight: 600 }}>{title}</Typography>
        <Tooltip title={`Download "${title}" as .docx`} placement="top">
          <IconButton size="small" onClick={() => downloadSingleTableDocx(headers, rows, title)} sx={{ color: "#c9a84c", p: 0.4, "&:hover": { color: "#e8c87a", background: "rgba(201,168,76,0.1)" } }}>
            <DownloadIcon sx={{ fontSize: 15 }} />
          </IconButton>
        </Tooltip>
      </Box>
      <Box sx={{ "& .MuiTableContainer-root": { borderRadius: "0 0 6px 6px", mb: "0 !important" } }}>{children}</Box>
    </Box>
  );
}

// ─────────────────────────────────────────────
// SUMMARY TABLE COLUMNS
// ─────────────────────────────────────────────
const SUMMARY_COLS = [
  { key: "company", label: "Company", width: 110, required: true },
  { key: "country", label: "Country", width: 90, required: true },
  { key: "industry", label: "Industry", width: 110, required: true },
  { key: "status", label: "Status", width: 90 },
  { key: "expansion_type", label: "Expansion Type", width: 110 },
  { key: "date", label: "Date", width: 90 },
  { key: "location", label: "Location (USA)", width: 130 },
  { key: "investment_amount", label: "Investment", width: 100, required: true },
  { key: "confidence", label: "Confidence", width: 90 },
  { key: "website", label: "Website", width: 160 },
  { key: "email", label: "Email", width: 170 },
  { key: "phone", label: "Phone", width: 120 },
  { key: "twitter", label: "Twitter/X", width: 110 },
  { key: "linkedin", label: "LinkedIn", width: 150 },
  { key: "source", label: "Source", width: 150, required: true },
  { key: "summary", label: "Summary", width: 260 },
];

const REQUIRED_KEYS = SUMMARY_COLS.filter(c => c.required).map(c => c.key);
const hasValue = (v) => v && v !== "—" && v.toLowerCase() !== "not disclosed" && v.toLowerCase() !== "n/a";
const isComplete = (lead) => REQUIRED_KEYS.every(k => hasValue(lead[k]));

// ─────────────────────────────────────────────
// LEADS TABLE  (shared renderer)
// ─────────────────────────────────────────────
function LeadsTable({ leads, cols, title }) {
  if (!leads.length) return null;
  const headers = cols.map(c => c.label);
  const rows = leads.map(r => cols.map(c => r[c.key] ?? "—"));

  const renderCell = (col, val) => {
    if (col.key === "confidence" && hasValue(val)) { const c = confidenceColor(val); return <TableCell key={col.key} sx={{ py: 0.6, whiteSpace: "nowrap" }}><Chip label={val} size="small" sx={{ background: c.bg, color: c.color, fontWeight: 600, fontSize: "10px" }} /></TableCell>; }
    if (col.key === "status" && hasValue(val)) { const sc = { "launched": { bg: "#d1fae5", color: "#065f46" }, "in progress": { bg: "#dbeafe", color: "#1e40af" }, "announced": { bg: "#fef3c7", color: "#92400e" }, "planning": { bg: "#f3e8ff", color: "#6b21a8" } }[val.toLowerCase()] || { bg: "#f1f5f9", color: "#475569" }; return <TableCell key={col.key} sx={{ py: 0.6, whiteSpace: "nowrap" }}><Chip label={val} size="small" sx={{ background: sc.bg, color: sc.color, fontWeight: 600, fontSize: "10px" }} /></TableCell>; }
    if ((col.key === "website" || col.key === "linkedin") && val.startsWith("http")) return <TableCell key={col.key} sx={{ fontSize: "11px", py: 0.6, maxWidth: col.width, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}><a href={val} target="_blank" rel="noreferrer" style={{ color: "#c9a84c", textDecoration: "underline" }}>{val}</a></TableCell>;
    if (col.key === "email" && val.includes("@")) return <TableCell key={col.key} sx={{ fontSize: "11px", py: 0.6, whiteSpace: "nowrap" }}><a href={`mailto:${val}`} style={{ color: "#c9a84c" }}>{val}</a></TableCell>;
    if (col.key === "company") return <TableCell key={col.key} sx={{ fontSize: "11px", py: 0.6, fontWeight: 700, color: "#1e3a5f", whiteSpace: "nowrap" }}>{val}</TableCell>;
    if (col.key === "summary") return <TableCell key={col.key} sx={{ fontSize: "10.5px", py: 0.6, maxWidth: col.width, whiteSpace: "normal", lineHeight: 1.4, color: "#374151" }}>{val}</TableCell>;
    if (col.required && !hasValue(val)) return <TableCell key={col.key} sx={{ fontSize: "11px", py: 0.6, background: "#fef2f2", color: "#dc2626", fontWeight: 600, whiteSpace: "nowrap" }}>Missing</TableCell>;
    return <TableCell key={col.key} sx={{ fontSize: "11px", py: 0.6, maxWidth: col.width, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{val}</TableCell>;
  };

  return (
    <TableWithDownload headers={headers} rows={rows} title={title}>
      <TableContainer component={Paper} sx={{ border: "1px solid rgba(201,168,76,0.15)", maxHeight: 460, overflow: "auto" }}>
        <Table size="small" stickyHeader>
          <TableHead>
            <TableRow>
              {cols.map(col => <TableCell key={col.key} sx={{ background: "#1e3a5f", color: "#c9a84c", fontWeight: "bold", fontSize: "10px", py: 0.7, whiteSpace: "nowrap", minWidth: col.width }}>{col.label}{col.required ? " *" : ""}</TableCell>)}
            </TableRow>
          </TableHead>
          <TableBody>
            {leads.map((row, ri) => (
              <TableRow key={ri} sx={{ background: ri % 2 === 0 ? "#fff" : "#f8fafc", verticalAlign: "top" }}>
                {cols.map(col => renderCell(col, row[col.key] ?? "—"))}
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>
    </TableWithDownload>
  );
}

// ─────────────────────────────────────────────
// SUMMARY TABLE  (complete + incomplete split)
// ─────────────────────────────────────────────
function SummaryTable({ sections }) {
  const allLeads = buildSummaryData(sections);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [dialogLeads, setDialogLeads] = useState([]);
  const [dialogLabel, setDialogLabel] = useState("");

  if (!allLeads.length) return null;

  const completeLeads = allLeads.filter(l => isComplete(l));
  const incompleteLeads = allLeads.filter(l => !isComplete(l));

  const completeCols = SUMMARY_COLS.filter(col => completeLeads.some(r => hasValue(r[col.key])));
  const incompleteCols = SUMMARY_COLS.filter(col => col.required || incompleteLeads.some(r => hasValue(r[col.key])));

  const openDialog = (leads, label) => { setDialogLeads(leads); setDialogLabel(label); setDialogOpen(true); };

  // ── Save Leads button ──
  const SaveLeadsBtn = ({ leads, label }) => (
    <Box sx={{ display: "flex", justifyContent: "flex-end", mt: 1.5, gap: 1.5 }}>
      {/* Keep SQL viewer as secondary option */}
      {/* <Button variant="outlined" size="small"
        onClick={() => {
          const sql = generateSQL(leads);
          const win = window.open("", "_blank", "width=900,height=640");
          win.document.write(`<html><head><title>SQL</title><style>body{margin:0;background:#0f172a;color:#e2e8f0;font-family:'Courier New',monospace;}.toolbar{background:#1e3a5f;padding:12px 20px;display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid #c9a84c;}.toolbar h2{margin:0;color:#c9a84c;font-size:15px;font-family:sans-serif;}.copy-btn{background:#c9a84c;color:#0f172a;border:none;border-radius:6px;padding:7px 18px;font-weight:700;cursor:pointer;font-size:13px;}pre{margin:0;padding:24px;font-size:13px;line-height:1.65;white-space:pre-wrap;word-break:break-word;}</style></head><body><div class="toolbar"><h2>💾 SQL — ${leads.length} record${leads.length !== 1 ? "s" : ""}</h2><button class="copy-btn" onclick="navigator.clipboard.writeText(document.getElementById('sql').innerText).then(()=>this.textContent='✓ Copied!')">Copy SQL</button></div><pre id="sql">${sql.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</pre></body></html>`);
          win.document.close();
        }}
        sx={{ borderRadius: "10px", borderColor: "rgba(201,168,76,0.4)", color: "#6b7280", textTransform: "none", fontSize: "12px", "&:hover": { borderColor: "#c9a84c", color: "#1e3a5f" } }}>
        View SQL
      </Button> */}

      {/* Primary: push to DB */}
      <Button variant="contained" startIcon={<SaveIcon />} onClick={() => openDialog(leads, label)}
        sx={{ background: "linear-gradient(135deg, #1e3a5f 0%, #2a4f7c 100%)", color: "#c9a84c", fontFamily: "'DM Sans', sans-serif", fontWeight: 700, fontSize: "13px", borderRadius: "10px", px: 3, py: 1, textTransform: "none", border: "1px solid rgba(201,168,76,0.4)", "&:hover": { background: "linear-gradient(135deg, #2a4f7c 0%, #1e3a5f 100%)", borderColor: "#c9a84c", color: "#e8c87a" } }}>
        Save Leads ({leads.length})
      </Button>
    </Box>
  );

  return (
    <Box>
      {/* Incomplete */}
      {incompleteLeads.length > 0 && (
        <Box sx={{ mb: 3 }}>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 1 }}>
            <Typography sx={{ color: "#dc2626", fontWeight: 700, fontSize: "13px", fontFamily: "'Playfair Display', serif" }}>⚠️ Incomplete Records ({incompleteLeads.length})</Typography>
            <Chip label="Missing required fields" size="small" sx={{ background: "#fee2e2", color: "#dc2626", fontSize: "10px", fontWeight: 600 }} />
            <Typography sx={{ fontSize: "10px", color: "#6b7280" }}>* = required · Red = missing</Typography>
          </Box>
          <LeadsTable leads={incompleteLeads} cols={incompleteCols} title="Incomplete Companies" />
          <SaveLeadsBtn leads={incompleteLeads} label="Incomplete Leads" />
        </Box>
      )}

      {/* Complete */}
      {completeLeads.length > 0 && (
        <Box>
          <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 1 }}>
            <Typography sx={{ color: "#065f46", fontWeight: 700, fontSize: "13px", fontFamily: "'Playfair Display', serif" }}>✅ Complete Records ({completeLeads.length})</Typography>
            <Typography sx={{ fontSize: "10px", color: "#6b7280" }}>All required fields present · Only columns with data shown</Typography>
          </Box>
          <LeadsTable leads={completeLeads} cols={completeCols} title="Complete Companies Summary" />
          <SaveLeadsBtn leads={completeLeads} label="Complete Leads" />
        </Box>
      )}

      {/* ── DB Push Dialog ── */}
      <SaveLeadsDialog
        open={dialogOpen}
        onClose={() => setDialogOpen(false)}
        leads={dialogLeads}
        label={dialogLabel}
      />
    </Box>
  );
}

// ─────────────────────────────────────────────
// KV TABLE
// ─────────────────────────────────────────────
function KVTable({ headers, rows }) {
  return (
    <TableContainer component={Paper} sx={{ border: "1px solid rgba(201,168,76,0.15)" }}>
      <Table size="small">
        <TableHead><TableRow>{headers.map((h, i) => <TableCell key={i} sx={{ background: "#1e3a5f", color: "#c9a84c", fontWeight: "bold", fontSize: "11px", py: 0.8 }}>{h}</TableCell>)}</TableRow></TableHead>
        <TableBody>
          {rows.map((row, ri) => (
            <TableRow key={ri} sx={{ background: ri % 2 === 0 ? "#fff" : "#f8fafc" }}>
              {row.map((cell, ci) => {
                if (row[0]?.toLowerCase() === "confidence" && ci === 1) { const c = confidenceColor(cell); return <TableCell key={ci} sx={{ fontSize: "12px", py: 0.7 }}><Chip label={cell || "—"} size="small" sx={{ background: c.bg, color: c.color, fontWeight: 600, fontSize: "10px" }} /></TableCell>; }
                return <TableCell key={ci} sx={{ fontSize: "12px", py: 0.7, fontWeight: ci === 0 ? 600 : 400, color: ci === 0 ? "#1e3a5f" : "#374151" }}>{cell || "—"}</TableCell>;
              })}
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </TableContainer>
  );
}

// ─────────────────────────────────────────────
// CONTACT TABLE
// ─────────────────────────────────────────────
function ContactTable({ headers, rows }) {
  const channelIcon = (ch = "") => { const c = ch.toLowerCase(); if (c.includes("website") || c.includes("url")) return "🌐"; if (c.includes("email")) return "✉️"; if (c.includes("phone")) return "📞"; if (c.includes("twitter") || c.includes("x")) return "𝕏"; if (c.includes("linkedin")) return "in"; return "•"; };
  return (
    <TableContainer component={Paper} sx={{ border: "1px solid rgba(201,168,76,0.15)" }}>
      <Table size="small">
        <TableHead><TableRow>{headers.map((h, i) => <TableCell key={i} sx={{ background: "#1e3a5f", color: "#c9a84c", fontWeight: "bold", fontSize: "11px", py: 0.8 }}>{h}</TableCell>)}</TableRow></TableHead>
        <TableBody>
          {rows.map((row, ri) => (
            <TableRow key={ri} sx={{ background: ri % 2 === 0 ? "#fff" : "#f8fafc" }}>
              {row.map((cell, ci) => {
                const isChannel = ci === 0; const isLink = !isChannel && (cell.startsWith("http") || cell.includes("@"));
                return <TableCell key={ci} sx={{ fontSize: "12px", py: 0.7, fontWeight: isChannel ? 600 : 400, color: isChannel ? "#1e3a5f" : "#374151" }}>{isChannel ? <span>{channelIcon(cell)} {cell}</span> : isLink ? <a href={cell.startsWith("http") ? cell : `mailto:${cell}`} target="_blank" rel="noreferrer" style={{ color: "#c9a84c", wordBreak: "break-all" }}>{cell}</a> : (cell || "N/A")}</TableCell>;
              })}
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </TableContainer>
  );
}

// ─────────────────────────────────────────────
// WIDE EXPANSION TABLE
// ─────────────────────────────────────────────
const EXPANSION_COLS = [
  { key: "date", label: "Date" }, { key: "source", label: "Source" }, { key: "company", label: "Company" },
  { key: "country_of_origin", label: "Country" }, { key: "industry", label: "Industry" },
  { key: "expansion_type", label: "Type" }, { key: "investment_location", label: "Location" },
  { key: "investment_amount", label: "Amount" }, { key: "confidence", label: "Confidence" },
];

function ExpansionTable({ headers, rows }) {
  const normalizeKey = (raw) => { const k = stripBold(raw).toLowerCase().trim(); if (k.includes("date")) return "date"; if (k.includes("source")) return "source"; if (k.includes("company")) return "company"; if (k.includes("country")) return "country_of_origin"; if (k.includes("industry") || k.includes("sector")) return "industry"; if (k.includes("expansion") || k.includes("type")) return "expansion_type"; if (k.includes("location") || k.includes("city")) return "investment_location"; if (k.includes("amount") || k.includes("investment")) return "investment_amount"; if (k.includes("contact")) return "contact_details"; if (k.includes("confidence")) return "confidence"; return k.replace(/\s+/g, "_"); };
  const keys = headers.map(normalizeKey);
  const data = rows.map(row => { const obj = {}; keys.forEach((key, i) => { obj[key] = row[i]?.trim() || "—"; }); return obj; });
  const visibleCols = EXPANSION_COLS.filter(col => keys.includes(col.key));
  return (
    <TableContainer component={Paper} sx={{ maxHeight: 380, overflow: "auto", border: "1px solid rgba(201,168,76,0.15)" }}>
      <Table stickyHeader size="small">
        <TableHead><TableRow>{visibleCols.map(col => <TableCell key={col.key} sx={{ background: "#1e3a5f", color: "#c9a84c", fontWeight: "bold", fontSize: "11px", whiteSpace: "nowrap", py: 0.8 }}>{col.label}</TableCell>)}</TableRow></TableHead>
        <TableBody>
          {data.map((row, ri) => (
            <TableRow key={ri} sx={{ background: ri % 2 === 0 ? "#fff" : "#f8fafc" }}>
              {visibleCols.map(col => { if (col.key === "confidence") { const c = confidenceColor(row[col.key]); return <TableCell key={col.key} sx={{ py: 0.7, whiteSpace: "nowrap" }}><Chip label={row[col.key] || "—"} size="small" sx={{ background: c.bg, color: c.color, fontWeight: 600, fontSize: "10px" }} /></TableCell>; } return <TableCell key={col.key} sx={{ fontSize: "12px", py: 0.7, maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis" }}>{row[col.key] || "—"}</TableCell>; })}
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </TableContainer>
  );
}

// ─────────────────────────────────────────────
// SMART TABLE ROUTER
// ─────────────────────────────────────────────
function SmartTable({ headers, rows }) {
  const title = inferTableTitle(headers); const type = tableType(headers);
  const isKV = headers.length === 2 && (headers[0].toLowerCase().includes("field") || headers[0].toLowerCase().includes("channel"));
  let inner;
  if (type === "contact" || (isKV && headers[0].toLowerCase().includes("channel"))) inner = <ContactTable headers={headers} rows={rows} />;
  else if (isKV) inner = <KVTable headers={headers} rows={rows} />;
  else inner = <ExpansionTable headers={headers} rows={rows} />;
  return <TableWithDownload headers={headers} rows={rows} title={title}>{inner}</TableWithDownload>;
}

// ─────────────────────────────────────────────
// COMPANY CARD
// ─────────────────────────────────────────────
function CompanyCard({ section }) {
  return (
    <Box sx={{ border: "1px solid rgba(201,168,76,0.25)", borderRadius: "12px", overflow: "hidden", mb: 2 }}>
      <Box sx={{ background: "linear-gradient(135deg, #1e3a5f 0%, #2a4f7c 100%)", px: 2, py: 1.2, display: "flex", alignItems: "center", gap: 1, flexWrap: "wrap" }}>
        <Typography sx={{ color: "#e8c87a", fontFamily: "'Playfair Display', serif", fontWeight: 700, fontSize: "14px" }}>{section.name}</Typography>
        {section.country && <Chip label={section.country} size="small" sx={{ background: "rgba(201,168,76,0.15)", color: "#e8c87a", fontSize: "10px", height: 18 }} />}
        {section.industry && <Chip label={section.industry} size="small" sx={{ background: "rgba(255,255,255,0.08)", color: "#94a3b8", fontSize: "10px", height: 18 }} />}
      </Box>
      <Box sx={{ px: 2, py: 1.5, background: "#fff" }} className="markdown-body">
        {section.blocks.map((block, bi) => {
          if (block.type === "table") return <SmartTable key={bi} headers={block.headers} rows={block.rows} />;
          const cleaned = block.content.replace(/^\*\*Summary\*\*\s*/im, "").replace(/^\*\*Expansion Data\*\*\s*/im, "").replace(/^\*\*Contact Details\*\*\s*/im, "").trim();
          if (!cleaned) return null;
          return <Box key={bi} sx={{ mb: 1 }}><ReactMarkdown remarkPlugins={[remarkGfm]}>{cleaned.replace(/•/g, "-")}</ReactMarkdown></Box>;
        })}
      </Box>
    </Box>
  );
}

function ToolBadge({ name }) { return <div className="tool-badge"><span className="tool-icon">⚡</span><span>Searching: {name}</span></div>; }

// ─────────────────────────────────────────────
// FULL RESPONSE DOCX DOWNLOAD
// ─────────────────────────────────────────────
function buildDocxTable(headers, rows, effectiveWidth) {
  const colW = Math.floor(effectiveWidth / headers.length);
  const colWs = headers.map((_, i) => i === headers.length - 1 ? effectiveWidth - colW * (headers.length - 1) : colW);
  const border = { style: BorderStyle.SINGLE, size: 1, color: "D1D5DB" }; const borders = { top: border, bottom: border, left: border, right: border }; const margins = { top: 60, bottom: 60, left: 80, right: 80 };
  const hRow = new DocxTableRow({ tableHeader: true, children: headers.map((h, i) => new DocxTableCell({ borders, margins, width: { size: colWs[i], type: WidthType.DXA }, shading: { fill: "1E3A5F", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "C9A84C", size: 16 })] })] })) });
  const dRows = rows.map((row, ri) => new DocxTableRow({ children: row.map((cell, i) => new DocxTableCell({ borders, margins, width: { size: colWs[i], type: WidthType.DXA }, shading: { fill: ri % 2 === 0 ? "FFFFFF" : "F8FAFC", type: ShadingType.CLEAR }, children: [new Paragraph({ children: [new TextRun({ text: String(cell ?? "—"), size: 14 })] })] })) }));
  return new DocxTable({ width: { size: effectiveWidth, type: WidthType.DXA }, columnWidths: colWs, rows: [hRow, ...dRows] });
}

const downloadFullResponseDocx = async (parsedData) => {
  try {
    const MARGIN = 720; const EFF_W = 15840 - MARGIN * 2;
    const textPara = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: stripMarkdown(text), size: opts.size || 18, color: opts.color || "111827", bold: opts.bold })] });
    const children = [new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { after: 160 }, children: [new TextRun({ text: "USA Market Expansion Intelligence Report", bold: true, size: 28, color: "1E3A5F" })] }), new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: `Generated: ${new Date().toLocaleString()}`, size: 16, color: "6B7280", italics: true })] })];
    const { preamble, sections, analysisBlocks } = parsedData;
    for (const b of preamble) { if (b.type === "text") { for (const line of b.content.split("\n").filter(Boolean)) children.push(textPara(line, { size: 20 })); } else children.push(buildDocxTable(b.headers, b.rows, EFF_W)); children.push(new Paragraph({ spacing: { after: 160 }, children: [] })); }
    for (const sec of sections) { children.push(new Paragraph({ spacing: { before: 300, after: 160 }, children: [new TextRun({ text: `${sec.name}  ·  ${sec.country}  ·  ${sec.industry}`, bold: true, size: 22, color: "1E3A5F" })] })); for (const b of sec.blocks) { if (b.type === "text") { for (const line of b.content.split("\n").filter(Boolean)) { const clean = stripMarkdown(line); if (clean) children.push(textPara(clean, { size: 18 })); } } else children.push(buildDocxTable(b.headers, b.rows, EFF_W)); } children.push(new Paragraph({ spacing: { after: 120 }, children: [] })); }
    if (analysisBlocks.length) { children.push(new Paragraph({ spacing: { before: 300, after: 160 }, children: [new TextRun({ text: "Overall Analysis", bold: true, size: 24, color: "1E3A5F" })] })); for (const b of analysisBlocks) { if (b.type === "text") { for (const line of b.content.split("\n").filter(Boolean)) children.push(textPara(stripMarkdown(line), { size: 18, color: "374151" })); } } }
    children.push(new Paragraph({ spacing: { before: 400 }, children: [new TextRun({ text: "All information sourced from verified public reports. Facts only.", size: 14, color: "9CA3AF", italics: true })] }));
    const doc = new Document({ styles: { default: { document: { run: { font: "Arial", size: 18 } } } }, sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } }, children }] });
    const blob = await Packer.toBlob(doc); saveAs(blob, `USA_Market_Expansion_${new Date().toISOString().slice(0, 10)}.docx`);
  } catch (err) { console.error("DOCX error:", err); alert("Download failed: " + err.message); }
};

// ─────────────────────────────────────────────
// MESSAGE COMPONENT
// ─────────────────────────────────────────────
function Message({ msg, onDownloadFull }) {
  if (msg.role === "user") return <div className="msg msg-user"><div className="msg-bubble msg-bubble-user">{msg.content}</div></div>;
  const blocks = parseResponseBlocks(msg.content || "");
  const parsed = parseCompanySections(blocks);
  const hasCompanies = parsed.sections.length > 0;
  const hasTables = blocks.some(b => b.type === "table");
  return (
    <div className="msg msg-agent">
      <div className="agent-avatar"><span>AI</span></div>
      <div className="msg-body">
        <div className="msg-bubble msg-bubble-agent markdown-body">
          {hasTables && (
            <Box sx={{ display: "flex", justifyContent: "flex-end", alignItems: "center", mb: 1, gap: 0.5 }}>
              <Typography sx={{ fontSize: "10px", color: "#9ca3af", letterSpacing: "0.05em" }}>Download full response</Typography>
              <Tooltip title="Download entire response as .docx" placement="top">
                <IconButton size="small" onClick={() => onDownloadFull(parsed)} sx={{ color: "#c9a84c", p: 0.5, "&:hover": { color: "#e8c87a", background: "rgba(201,168,76,0.1)" } }}>
                  <DownloadIcon sx={{ fontSize: 18 }} />
                </IconButton>
              </Tooltip>
            </Box>
          )}
          {hasCompanies ? (
            <Box sx={{ width: "100%" }}>
              {parsed.preamble.map((b, i) => <Box key={i} sx={{ mb: 2 }}>{b.type === "text" ? <ReactMarkdown remarkPlugins={[remarkGfm]}>{b.content.replace(/•/g, "-")}</ReactMarkdown> : <SmartTable headers={b.headers} rows={b.rows} />}</Box>)}
              {parsed.sections.map((sec, i) => <CompanyCard key={i} section={sec} />)}
              {parsed.analysisBlocks.length > 0 && (
                <Box sx={{ mt: 1, pt: 1.5, borderTop: "2px solid rgba(201,168,76,0.3)" }}>
                  <Typography sx={{ color: "#1e3a5f", fontWeight: 700, fontSize: "13px", mb: 1, fontFamily: "'Playfair Display', serif" }}>Overall Analysis</Typography>
                  {parsed.analysisBlocks.map((b, i) => <Box key={i}>{b.type === "text" ? <ReactMarkdown remarkPlugins={[remarkGfm]}>{b.content.replace(/•/g, "-")}</ReactMarkdown> : <SmartTable headers={b.headers} rows={b.rows} />}</Box>)}
                </Box>
              )}
              {parsed.sections.length > 0 && (
                <Box sx={{ mt: 3.5, pt: 2.5, borderTop: "3px solid rgba(201,168,76,0.35)" }}>
                  <Typography sx={{ color: "#1e3a5f", fontWeight: 700, fontSize: "15px", mb: 1.5, fontFamily: "'Playfair Display', serif", display: "flex", alignItems: "center", gap: 1 }}>📋 All Companies Summary</Typography>
                  <SummaryTable sections={parsed.sections} />
                </Box>
              )}
            </Box>
          ) : (
            <Box sx={{ width: "100%" }}>
              {blocks.map((b, i) => <Box key={i} sx={{ mb: 1.5 }}>{b.type === "text" ? <ReactMarkdown remarkPlugins={[remarkGfm]}>{b.content.replace(/•/g, "-").replace(/\n{2,}/g, "\n\n")}</ReactMarkdown> : <SmartTable headers={b.headers} rows={b.rows} />}</Box>)}
            </Box>
          )}
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────
// APP
// ─────────────────────────────────────────────
export default function App() {
  const [sessionId, setSessionId] = useState(null);
  const [messages, setMessages] = useState([]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [starting, setStarting] = useState(false);
  const [error, setError] = useState(null);
  const [selectedIndustry, setSelectedIndustry] = useState("");
  const [showSamples, setShowSamples] = useState(false);
  const bottomRef = useRef(null);

  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [messages]);

  async function startSession() {
    setStarting(true); setError(null);
    try {
      const res = await fetch(`${API}/sessions`, { method: "POST" });
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      setSessionId(data.session_id);
      setMessages([{ id: Date.now(), role: "agent", content: "Session started. Ask me anything about companies entering or expanding into the US market since 2025.", tools: [] }]);
    } catch (e) { setError("Failed to start session: " + e.message); }
    finally { setStarting(false); }
  }

  async function sendMessage(text) {
    if (!text.trim() || loading || !sessionId) return;
    setInput("");
    setLoading(true);
    setError(null);

    const finalMessage = selectedIndustry ? `[Industry: ${selectedIndustry}] ${text}` : text;
    const userMsg = { id: Date.now(), role: "user", content: text };
    const agentMsg = { id: Date.now() + 1, role: "agent", content: "", tools: [] };
    setMessages(prev => [...prev, userMsg, agentMsg]);

    let fullContent = "";
    let retries = 0;
    const maxRetries = 3;

    while (retries <= maxRetries) {
      try {
        const res = await fetch(`${API}/chat`, {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ message: finalMessage, session_id: sessionId }),
        });

        if (!res.ok) throw new Error(await res.text());

        const reader = res.body.getReader();
        const decoder = new TextDecoder();

        while (true) {
          const { done, value } = await reader.read();
          if (done) break;

          for (const line of decoder.decode(value).split("\n")) {
            if (!line.startsWith("data:")) continue;
            try {
              const event = JSON.parse(line.slice(5).trim());
              if (event.type === "text") {
                fullContent += event.content;
                setMessages(prev => prev.map(m =>
                  m.id === agentMsg.id ? { ...m, content: fullContent } : m
                ));
              } else if (event.type === "tool") {
                setMessages(prev => prev.map(m =>
                  m.id === agentMsg.id ? { ...m, tools: [...(m.tools || []), event.content] } : m
                ));
              } else if (event.type === "done") {
                setMessages(prev => prev.map(m =>
                  m.id === agentMsg.id ? { ...m, content: fullContent } : m
                ));
                setLoading(false);
                return; // Success - exit function
              } else if (event.type === "error") {
                throw new Error(event.error?.message || "Unknown error");
              }
            } catch (_) { }
          }
        }
        break; // Success

      } catch (e) {
        retries++;
        if (retries > maxRetries || !e.message.includes("rate limited")) {
          setError("Error: " + e.message);
          setLoading(false);
          return;
        }
        // Wait before retry (exponential backoff)
        const waitTime = Math.pow(2, retries) * 1000; // 2s, 4s, 8s
        console.log(`Rate limited. Retrying in ${waitTime / 1000}s... (attempt ${retries})`);
        await new Promise(r => setTimeout(r, waitTime));
      }
    }
    setLoading(false);
  }

  const handleKey = (e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendMessage(input); } };
  const handleInputFocus = () => setShowSamples(true);
  const handleInputBlur = () => setTimeout(() => setShowSamples(false), 200);
  const handleSampleClick = (q) => { setInput(q); setShowSamples(false); sendMessage(q); };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:wght@300;400;500&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        :root { --navy: #f2f3f5; --navy-mid: #fff; --navy-dark: #1e3a5f; --gold: #c9a84c; --gold-light: #e8c87a; --text: #111827; --text-dim: #6b7280; --border: rgba(201,168,76,0.2); }
        body { font-family: 'DM Sans', sans-serif; background: var(--navy); color: var(--text); height: 100vh; display: flex; flex-direction: column; overflow: hidden; }
        .header { background: #e9e6e6; border-bottom: 1px solid var(--border); padding: 14px 28px; display: flex; align-items: center; gap: 14px; flex-shrink: 0; }
        .header-title { font-family: 'Playfair Display', serif; font-size: 18px; color: var(--navy-dark); }
        .header-sub { font-size: 11px; color: #2f3e4f; letter-spacing: 0.08em; text-transform: uppercase; margin-top: 1px; }
        .header-status { margin-left: auto; display: flex; align-items: center; gap: 7px; font-size: 12px; color: var(--text-dim); }
        .status-dot { width: 7px; height: 7px; border-radius: 50%; background: #4a4a5a; }
        .status-dot.active { background: #3ddc84; box-shadow: 0 0 6px #3ddc84; }
        .chat-area { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
        .messages { flex: 1; overflow-y: auto; padding: 16px 8px; display: flex; flex-direction: column; gap: 12px; align-items: center; }
        .messages::-webkit-scrollbar { width: 4px; }
        .messages::-webkit-scrollbar-thumb { background: var(--navy-dark); border-radius: 2px; }
        .empty { flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 12px; text-align: center; padding: 20px; }
        .empty-icon { font-size: 52px; }
        .empty-title { font-family: 'Playfair Display', serif; font-size: 22px; color: var(--navy-dark); }
        .empty-sub { font-size: 14px; color: var(--text-dim); max-width: 360px; line-height: 1.6; }
        .start-btn { margin-top: 8px; background: var(--gold); color: #fff; border: none; border-radius: 10px; padding: 13px 32px; font-family: 'DM Sans', sans-serif; font-weight: 600; font-size: 14px; cursor: pointer; transition: all 0.18s; }
        .start-btn:hover:not(:disabled) { background: var(--gold-light); transform: translateY(-1px); }
        .start-btn:disabled { opacity: 0.5; cursor: not-allowed; }
        .msg { display: flex; gap: 12px; width: 100%; }
        .msg-user { justify-content: flex-end; }
        .agent-avatar { width: 32px; height: 32px; border-radius: 50%; background: var(--navy-dark); display: flex; align-items: center; justify-content: center; font-size: 12px; flex-shrink: 0; margin-top: 4px; color: #fff; font-weight: 600; }
        .msg-body { display: flex; flex-direction: column; gap: 5px; max-width: 100%; width: 100%; }
        .msg-bubble { padding: 12px 14px; border-radius: 14px; font-size: 14px; line-height: 1.5; word-break: break-word; }
        .msg-bubble-user { background: var(--gold); color: #fff; border-radius: 14px 14px 4px 14px; max-width: 500px; font-weight: 500;text-align: left;direction: ltr; }
        .msg-bubble-agent { background: var(--navy-mid); border: 1px solid var(--border); border-radius: 4px 14px 14px 14px; width: 100%;text-align: left;direction: ltr; }
        .tool-badge { display: flex; align-items: center; gap: 5px; background: rgba(201,168,76,0.1); border: 1px solid rgba(201,168,76,0.3); border-radius: 20px; padding: 3px 10px; font-size: 11px; color: var(--gold-light); }
        .error-bar { margin: 0 24px 12px; background: rgba(176,58,46,0.1); border: 1px solid rgba(176,58,46,0.3); border-radius: 8px; padding: 10px 14px; font-size: 13px; color: #dc2626; }
        .input-area { border-top: 1px solid var(--border); padding: 18px 28px 22px; display: flex; gap: 0; background: var(--navy-mid); flex-shrink: 0; }
        .markdown-body { font-size: 14px; line-height: 1.75; text-align: left; direction: ltr; }
        .markdown-body h1, .markdown-body h2 { color: var(--navy-dark); margin: 10px 0 6px; font-weight: 700; }
        .markdown-body h3, .markdown-body h4 { color: #374151; margin: 8px 0 4px; font-weight: 600; font-size: 13px; }
        .markdown-body p { margin: 4px 0; }
        .markdown-body li { margin: 3px 0; }
        .markdown-body ul, .markdown-body ol { padding-left: 20px; margin: 4px 0; }
        .markdown-body strong { color: var(--navy-dark); }
        .markdown-body a { color: var(--gold); text-decoration: underline; word-break: break-all; }
        .markdown-body blockquote { border-left: 3px solid var(--gold); margin: 8px 0; padding-left: 12px; color: #6b7280; font-style: italic; }
        .markdown-body hr { border: none; border-top: 1px solid #e5e7eb; margin: 14px 0; }
        .markdown-body table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 12px; }
        .markdown-body th, .markdown-body td { border: 1px solid #e5e7eb; padding: 6px 8px; }
        .markdown-body th { background: #1e3a5f; color: #c9a84c; }
        .markdown-body p,
.markdown-body li,
.markdown-body h1,
.markdown-body h2,
.markdown-body h3,
.markdown-body h4 {
  text-align: left !important;
  direction: ltr !important;
}
      `}</style>

      <div className="header">
        <div>
          <div className="header-title">USA Market Expansion Intelligence</div>
          <div className="header-sub">Powered by AI Agents · Facts Only</div>
        </div>
        <div className="header-status">
          <div className={`status-dot ${sessionId ? "active" : ""}`} />
          {sessionId ? "Session Active" : "No Session"}
        </div>
      </div>

      <div className="chat-area">
        {!sessionId ? (
          <div className="empty">
            <div className="empty-icon">🔍</div>
            <div className="empty-title">Market Intelligence Agent</div>
            <div className="empty-sub">Ask about companies entering the US market since 2025. Verified sources only.</div>
            <button className="start-btn" onClick={startSession} disabled={starting}>{starting ? "Starting…" : "Start Session"}</button>
          </div>
        ) : (
          <>
            <div className="messages">
              {messages.map(msg => <Message key={msg.id} msg={msg} onDownloadFull={downloadFullResponseDocx} />)}
              <div ref={bottomRef} />
            </div>
            {error && <div className="error-bar">{error}</div>}

            <div className="input-area" style={{ position: "relative" }}>
              {/* Floating sample questions */}
              {showSamples && (
                <Box sx={{ position: "absolute", bottom: "100%", left: 0, right: 0, zIndex: 20, mb: 1.5 }}>
                  <Paper elevation={8} sx={{ background: "#ffffff", border: "1px solid rgba(201,168,76,0.2)", borderRadius: "18px", overflow: "hidden", boxShadow: "0 12px 35px rgba(30, 58, 95, 0.12)" }}>
                    <Box sx={{ px: 3, pt: 2.2, pb: 1.4, borderBottom: "1px solid rgba(201,168,76,0.15)", background: "linear-gradient(to bottom, #fafafa, #ffffff)" }}>
                      <Typography sx={{ fontSize: "10.5px", color: "#6b7280", fontWeight: 700, letterSpacing: "1px", textTransform: "uppercase", textAlign: "center" }}>TRY THESE QUESTIONS</Typography>
                    </Box>
                    <Box sx={{ py: 1 }}>
                      {SAMPLE_QUESTIONS.map((q, i) => (
                        <Box key={i} onClick={() => handleSampleClick(q)} sx={{ px: 3, py: 1.35, cursor: "pointer", fontSize: "13.8px", lineHeight: 1.45, color: "#374151", textAlign: "left", border: "1px solid #1e3a5f", borderRadius: "5px", transition: "all 0.12s ease", marginBottom: "2px", "&:hover": { background: "rgba(201,168,76,0.07)", color: "#1e3a5f", pl: 3.5 } }}>{q}</Box>
                      ))}
                    </Box>
                  </Paper>
                </Box>
              )}

              {/* Unified input box */}
              <Box sx={{ display: "flex", alignItems: "flex-end", background: "var(--navy)", border: "2px solid var(--border)", borderRadius: "20px", px: 2.5, py: 1.8, width: "100%", gap: 2, boxShadow: "0 4px 12px rgba(30, 58, 95, 0.08)", transition: "all 0.2s ease", "&:focus-within": { borderColor: "var(--gold)", boxShadow: "0 6px 20px rgba(201, 168, 76, 0.15)" } }}>
                <FormControl size="small" sx={{ minWidth: 175, flexShrink: 0, "& .MuiOutlinedInput-root": { background: "rgba(255,255,255,0.7)", borderRadius: "14px", border: "1px solid rgba(201,168,76,0.25)", "&:hover": { borderColor: "var(--gold)" }, "&.Mui-focused": { borderColor: "var(--gold)" } } }}>
                  <InputLabel id="industry-select-label" sx={{ color: "var(--text-dim)", fontSize: "0.75rem", fontWeight: 500, "&.Mui-focused": { color: "var(--gold)" } }}>Industry Type</InputLabel>
                  <Select labelId="industry-select-label" value={selectedIndustry} label="Industry Type" onChange={(e) => setSelectedIndustry(e.target.value)} disabled={loading} sx={{ fontSize: "13.5px", fontWeight: 500, color: selectedIndustry ? "#1e3a5f" : "var(--text-dim)" }}>
                    {INDUSTRIES.map((ind) => <MenuItem key={ind.value} value={ind.value}>{ind.label}</MenuItem>)}
                  </Select>
                </FormControl>

                <Box sx={{ flex: 1, position: "relative", display: "flex", alignItems: "flex-end" }}>
                  <textarea rows={1} placeholder={selectedIndustry ? `Ask about ${selectedIndustry.toLowerCase()} companies entering the US market…` : "Ask about US market expansions, investments, timelines…"} value={input} onChange={(e) => setInput(e.target.value)} onFocus={handleInputFocus} onBlur={handleInputBlur} onKeyDown={handleKey} disabled={loading}
                    style={{ width: "100%", background: "transparent", border: "none", outline: "none", fontSize: "15px", lineHeight: 1.5, color: "var(--text)", resize: "none", padding: "6px 0", minHeight: "28px", maxHeight: "120px", fontFamily: "'DM Sans', sans-serif" }} />
                </Box>

                <IconButton onClick={() => sendMessage(input)} disabled={loading || !input.trim()}
                  sx={{ width: 48, height: 48, background: "linear-gradient(135deg, #1e3a5f 0%, #2a4f7c 100%)", color: "#c9a84c", flexShrink: 0, borderRadius: "14px", "&:hover": { background: "linear-gradient(135deg, #2a4f7c 0%, #1e3a5f 100%)", color: "#e8c87a", transform: "scale(1.05)" }, "&:disabled": { background: "#e5e7eb", color: "#9ca3af" }, transition: "all 0.2s ease" }}>
                  {loading ? "⏳" : "→"}
                </IconButton>
              </Box>
            </div>
          </>
        )}
      </div>
    </>
  );
}