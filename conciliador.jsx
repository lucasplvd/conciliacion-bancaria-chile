import { useState } from "react";
import * as XLSX from "xlsx";
import { Upload, Download, Plus, Trash2, CheckCircle, AlertTriangle, XCircle, FileSpreadsheet, Play, RotateCcw, ChevronDown, ChevronRight, X, Layers } from "lucide-react";

/* ═══════════════════════════════════════════════
   UTILITIES
   ═══════════════════════════════════════════════ */

const normalizeRut = (rut) => {
  if (!rut) return "";
  let s = String(rut).replace(/[.\s\-]/g, "").toUpperCase().trim();
  if (s.length > 1) s = s.slice(0, -1).replace(/^0+/, "") + s.slice(-1);
  return s;
};

const normalizeName = (name) => {
  if (!name) return "";
  return String(name).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9\s]/g, "").replace(/\s+/g, " ").trim();
};

const STOPWORDS = new Set([
  "SPA", "SA", "LTDA", "LIMITADA", "EIRL", "SRL", "SAC", "CIA", "INC", "CORP",
  "SOCIEDAD", "ANONIMA", "EMPRESA", "COMPANIA", "COMPAÑIA",
  "COMERCIAL", "COMERCIALIZADORA", "SERVICIOS", "SERVICIO", "CONSTRUCTORA",
  "CONSTRUCCIONES", "CONSULTORA", "CONSULTORES", "ASESORIA", "ASESORIAS",
  "INGENIERIA", "INVERSIONES", "IMPORTADORA", "EXPORTADORA", "DISTRIBUIDORA",
  "INDUSTRIAL", "INDUSTRIA", "INDUSTRIAS", "INMOBILIARIA", "TRANSPORTES",
  "TRANSPORTE", "LOGISTICA", "SOLUCIONES", "TECNOLOGIA", "COMUNICACIONES",
  "INTERNACIONAL", "NACIONAL", "GENERAL", "GENERALES", "PROFESIONAL",
  "PROFESIONALES", "INTEGRAL", "INTEGRALES",
  "DE", "DEL", "LA", "LAS", "LOS", "EL", "EN", "Y", "E", "CON", "POR", "PARA",
  "SOC", "COM", "SERV", "ING", "ADM", "ADMINISTRADORA",
  "TRANSF", "TRANSFERENCIA", "PAGO", "ABONO", "CARGO", "DEPOSITO",
  "COMPRA", "VENTA", "CUOTA", "PAC", "ABN", "CRD", "DB", "TRAN",
]);

const extractKeywords = (name) => {
  const n = normalizeName(name);
  if (!n) return [];
  return n.split(" ").filter(w => w.length >= 3 && !STOPWORDS.has(w));
};

const nameSimilarity = (a, b) => {
  const na = normalizeName(a), nb = normalizeName(b);
  if (!na || !nb) return { score: 0, method: "" };
  if (na === nb) return { score: 1.0, method: "nombre_exacto" };
  if (na.includes(nb) || nb.includes(na)) return { score: 0.90, method: "nombre_contenido" };
  const kwA = extractKeywords(a);
  const kwB = extractKeywords(b);
  if (kwA.length > 0 && kwB.length > 0) {
    const setA = new Set(kwA), setB = new Set(kwB);
    const exactHits = [...setA].filter(w => setB.has(w));
    if (exactHits.length > 0) {
      const coverage = exactHits.length / Math.min(setA.size, setB.size);
      const sc = Math.max(0.70, Math.min(0.95, coverage * 0.95));
      return { score: sc, method: `keyword(${exactHits.join(",")})` };
    }
    let partialHits = [];
    for (const wa of setA) {
      for (const wb of setB) {
        if (wa.length >= 4 && wb.length >= 4) {
          if (wa.startsWith(wb) || wb.startsWith(wa)) {
            partialHits.push(wa.length <= wb.length ? wa : wb);
          } else if (wa.length >= 5 && wb.length >= 5 && (wa.includes(wb) || wb.includes(wa))) {
            partialHits.push(wa.length <= wb.length ? wa : wb);
          }
        }
      }
    }
    if (partialHits.length > 0) {
      return { score: 0.60, method: `keyword_parcial(${partialHits[0]})` };
    }
  }
  const tokA = new Set(na.split(" ").filter(t => t.length >= 3));
  const tokB = new Set(nb.split(" ").filter(t => t.length >= 3));
  const rawOverlap = [...tokA].filter(t => tokB.has(t));
  if (rawOverlap.length > 0) {
    const cov = rawOverlap.length / Math.min(tokA.size, tokB.size);
    if (cov >= 0.3) return { score: cov * 0.55, method: `token(${rawOverlap.slice(0, 2).join(",")})` };
  }
  return { score: 0, method: "" };
};

const parseD = (v) => { if (!v) return null; const d = v instanceof Date ? v : new Date(v); return isNaN(d.getTime()) ? null : d; };
const daysDiff = (a, b) => (!a || !b) ? Infinity : Math.abs(Math.round((a - b) / 86400000));
const fmtDate = (d) => { if (!d) return ""; const dt = d instanceof Date ? d : new Date(d); return isNaN(dt) ? "" : dt.toLocaleDateString("es-CL"); };
const fmtNum = (n) => (n == null || isNaN(n)) ? "" : new Intl.NumberFormat("es-CL").format(Math.round(n));
const toNum = (v) => { if (v == null) return NaN; if (typeof v === "number") return v; return parseFloat(String(v).replace(/[^\d.\-]/g, "")); };

/* ═══════════════════════════════════════════════
   COLUMN DETECTION
   ═══════════════════════════════════════════════ */

const KW = {
  rut: ["rut proveedor", "rut cliente", "rut", "r.u.t"],
  name: ["razon social", "razón social", "nombre proveedor", "nombre cliente", "beneficiario", "titular", "nombre"],
  folio: ["folio", "numero de factura", "numero factura", "nro doc", "numero doc", "boleta", "n°", "numero"],
  date: ["fecha doc", "fecha mov", "fecha movimiento", "fecha"],
  amount: ["monto total", "monto bruto", "total", "movimiento", "monto neto", "monto", "liquido", "líquido"],
  description: ["comentario banco", "comentario", "descripcion", "descripción", "glosa", "detalle", "concepto",
                 "texto", "referencia", "observacion", "observación", "informacion", "información",
                 "narracion", "narración", "mensaje", "operacion", "operación"],
  credit: ["abono", "abonos", "haber", "credito", "crédito", "deposito", "depósito", "depositos", "depósitos", "ingreso", "ingresos"],
  debit: ["cargo", "cargos", "debe", "debito", "débito", "giro", "giros", "egreso", "egresos", "retiro", "retiros"],
};

const detectMapping = (headers, type) => {
  const hl = headers.map(h => (h || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim());
  const m = {}, used = new Set();
  const fields = type === "bank" ? ["date", "description", "rut", "name", "amount", "credit", "debit"] : ["rut", "name", "folio", "date", "amount"];
  for (const f of fields) for (const kw of (KW[f] || [])) { const idx = hl.findIndex((h, i) => !used.has(i) && h.includes(kw)); if (idx >= 0) { m[f] = headers[idx]; used.add(idx); break; } }
  return m;
};

/* ═══════════════════════════════════════════════
   FILE PARSING
   ═══════════════════════════════════════════════ */

const parseFile = (file) => new Promise((res, rej) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
      res({ fileName: file.name, sheets: wb.SheetNames, workbook: wb });
    } catch (err) { rej(err); }
  };
  reader.onerror = rej;
  reader.readAsArrayBuffer(file);
});

const extractSheet = (wb, name) => {
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
  let hr = 0, ms = 0;
  for (let i = 0; i < Math.min(raw.length, 15); i++) { const s = (raw[i] || []).filter(c => typeof c === "string" && c.trim().length > 2).length; if (s > ms) { ms = s; hr = i; } }
  const headers = (raw[hr] || []).map((h, i) => h ? String(h).trim() : `Col_${i + 1}`);
  const data = raw.slice(hr + 1).filter(r => r && r.some(c => c != null && String(c).trim()));
  return { headers, data, rowCount: data.length };
};

/* ═══════════════════════════════════════════════
   MATCHING ENGINE v4
   ═══════════════════════════════════════════════ */

const buildRefDB = (sources) => {
  const entries = [];
  for (const [key, label] of [["lc", "LC"], ["lv", "LV"], ["remu", "Remuneraciones"], ["boletas", "Boletas"]]) {
    const s = sources[key]; if (!s?.data || !s?.mapping) continue;
    const { headers, data, mapping: mp } = s;
    const gi = (f) => mp[f] ? headers.indexOf(mp[f]) : -1;
    for (const row of data) {
      const g = (f) => { const i = gi(f); return i >= 0 ? row[i] : null; };
      const amt = toNum(g("amount"));
      const nameVal = String(g("name") || "");
      entries.push({
        rut: String(g("rut") || ""), rutN: normalizeRut(g("rut")),
        name: nameVal, nameN: normalizeName(nameVal), keywords: extractKeywords(nameVal),
        folio: String(g("folio") || ""), date: parseD(g("date")),
        amount: isNaN(amt) ? 0 : amt, absAmount: isNaN(amt) ? 0 : Math.abs(amt),
        source: label, matched: false,
      });
    }
  }
  const rutIndex = new Map(), kwIndex = new Map(), amtIndex = new Map();
  entries.forEach((e, i) => {
    if (e.rutN) { if (!rutIndex.has(e.rutN)) rutIndex.set(e.rutN, []); rutIndex.get(e.rutN).push(i); }
    for (const kw of e.keywords) { if (!kwIndex.has(kw)) kwIndex.set(kw, []); kwIndex.get(kw).push(i); }
    if (e.absAmount > 0) { const k = Math.round(e.absAmount); if (!amtIndex.has(k)) amtIndex.set(k, []); amtIndex.get(k).push(i); }
  });
  return { entries, rutIndex, kwIndex, amtIndex };
};

const scoreAmount = (bAmt, rAmt) => {
  const ba = Math.abs(bAmt), ra = Math.abs(rAmt);
  if (ba === 0 || ra === 0) return { s: 0, l: "" };
  if (Math.round(ba) === Math.round(ra)) return { s: 25, l: "monto_exacto" };
  const p = Math.abs(ba - ra) / Math.max(ba, ra);
  if (p <= 0.01) return { s: 18, l: "monto_~1%" };
  if (p <= 0.05) return { s: 10, l: "monto_~5%" };
  return { s: 0, l: "" };
};

const scoreDate = (bDate, rDate) => {
  const dd = daysDiff(bDate, rDate);
  if (dd === 0) return { s: 15, l: "fecha_exacta" };
  if (dd <= 3) return { s: 12, l: `fecha(±${dd}d)` };
  if (dd <= 7) return { s: 8, l: `fecha(±${dd}d)` };
  if (dd <= 15) return { s: 5, l: `fecha(±${dd}d)` };
  if (dd <= 30) return { s: 2, l: `fecha(±${dd}d)` };
  return { s: 0, l: "" };
};

const findBestMatch = (db, bankAmt, bankDate, bankRut, bankName, bankDesc) => {
  const rn = normalizeRut(bankRut);
  const bd = parseD(bankDate);
  const nameTexts = [bankName, bankDesc].filter(Boolean);

  const scoreCand = (ci, baseScore, baseReasons) => {
    const r = db.entries[ci]; if (r.matched) return null;
    let sc = baseScore, rs = [...baseReasons];
    const am = scoreAmount(bankAmt, r.amount); sc += am.s; if (am.l) rs.push(am.l);
    if (bd && r.date) { const dm = scoreDate(bd, r.date); sc += dm.s; if (dm.l) rs.push(dm.l); }
    return { ...r, score: sc, reasons: rs.join(", ") };
  };

  const bestOf = (cands, baseScore, baseReasons) => {
    let best = null, bs = 0;
    for (const ci of cands) {
      const c = scoreCand(ci, baseScore, baseReasons);
      if (c && c.score > bs) { bs = c.score; best = c; }
    }
    return best;
  };

  // ── PASS 1: RUT ──
  if (rn && rn.length >= 3) {
    const exact = db.rutIndex.get(rn) || [];
    if (exact.length > 0) {
      let best = null, bs = 0;
      for (const ci of exact) {
        const r = db.entries[ci]; if (r.matched) continue;
        let sc = 35, rs = ["rut_exacto"];
        let bestSim = { score: 0, method: "" };
        for (const nt of nameTexts) { const sim = nameSimilarity(nt, r.name); if (sim.score > bestSim.score) bestSim = sim; }
        if (bestSim.score >= 0.4) { sc += Math.round(bestSim.score * 15); rs.push(bestSim.method); }
        const am = scoreAmount(bankAmt, r.amount); sc += am.s; if (am.l) rs.push(am.l);
        if (bd && r.date) { const dm = scoreDate(bd, r.date); sc += dm.s; if (dm.l) rs.push(dm.l); }
        if (sc > bs) { bs = sc; best = { ...r, score: sc, reasons: rs.join(", ") }; }
      }
      if (best && best.score >= 40) return best;
    }
    if (rn.length >= 5) {
      const partials = [];
      for (const [ref, idx] of db.rutIndex) if (ref !== rn && (ref.includes(rn) || rn.includes(ref))) partials.push(...idx);
      if (partials.length) { const r = bestOf(partials, 25, ["rut_parcial"]); if (r && r.score >= 40) return r; }
    }
  }

  // ── PASS 2: Name/Keyword ──
  if (nameTexts.length > 0) {
    const candSet = new Set();
    for (const nt of nameTexts) {
      const kws = extractKeywords(nt);
      for (const kw of kws) (db.kwIndex.get(kw) || []).forEach(i => candSet.add(i));
      if (kws.length > 0) {
        for (const [indexKw, indices] of db.kwIndex) {
          for (const bankKw of kws) {
            if (bankKw.length >= 4 && indexKw.length >= 4 && bankKw !== indexKw) {
              if (bankKw.startsWith(indexKw) || indexKw.startsWith(bankKw) ||
                  (bankKw.length >= 5 && indexKw.length >= 5 && (bankKw.includes(indexKw) || indexKw.includes(bankKw)))) {
                indices.forEach(i => candSet.add(i));
              }
            }
          }
        }
      }
    }

    if (candSet.size > 0) {
      let best = null, bs = 0;
      const nameWeight = rn ? 30 : 35;
      const pass2Min = rn ? 30 : 25;
      for (const ci of candSet) {
        const r = db.entries[ci]; if (r.matched) continue;
        let bestSim = { score: 0, method: "" };
        for (const nt of nameTexts) { const sim = nameSimilarity(nt, r.name); if (sim.score > bestSim.score) bestSim = sim; }
        if (bestSim.score < 0.35) continue;
        let sc = Math.round(bestSim.score * nameWeight), rs = [bestSim.method];
        const am = scoreAmount(bankAmt, r.amount); sc += am.s; if (am.l) rs.push(am.l);
        if (bd && r.date) { const dm = scoreDate(bd, r.date); sc += dm.s; if (dm.l) rs.push(dm.l); }
        if (sc > bs) { bs = sc; best = { ...r, score: sc, reasons: rs.join(", ") }; }
      }
      if (best && best.score >= pass2Min) return best;
    }
  }

  // ── PASS 3: Amount only ──
  const absAmt = Math.abs(bankAmt);
  if (absAmt > 0) {
    const cands = db.amtIndex.get(Math.round(absAmt)) || [];
    let best = null, bs = 0;
    for (const ci of cands) {
      const r = db.entries[ci]; if (r.matched) continue;
      let sc = 20, rs = ["solo_monto"];
      if (bd && r.date) { const dm = scoreDate(bd, r.date); sc += dm.s; if (dm.l) rs.push(dm.l); }
      if (sc > bs) { bs = sc; best = { ...r, score: sc, reasons: rs.join(", ") }; }
    }
    if (best) return best;
  }
  return null;
};

/* ── PASS 4: Split payments ── */

const findSplitPayments = (movements, db) => {
  const unmIdx = []; for (let i = 0; i < movements.length; i++) if (!movements[i].match) unmIdx.push(i);
  if (unmIdx.length < 2) return;
  const unmRefs = []; db.entries.forEach((e, i) => { if (!e.matched && e.absAmount > 0) unmRefs.push(i); });
  if (!unmRefs.length) return;

  const groups = new Map();
  for (const mi of unmIdx) {
    const m = movements[mi], rn = normalizeRut(m.rut), kws = extractKeywords(m.name || m.desc);
    const key = rn.length >= 3 ? `R:${rn}` : kws.length > 0 ? `N:${kws.sort().join("_")}` : null;
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(mi);
  }

  for (const [gk, memberIdx] of groups) {
    if (memberIdx.length < 2) continue;
    const isRut = gk.startsWith("R:");
    const gid = gk.slice(2);

    const candRefs = unmRefs.filter(ri => {
      const r = db.entries[ri]; if (r.matched) return false;
      if (isRut) return r.rutN === gid;
      const refKws = r.keywords;
      const bankKws = gid.split("_");
      return bankKws.some(bk => refKws.some(rk => rk === bk || (bk.length >= 4 && rk.length >= 4 && (bk.startsWith(rk) || rk.startsWith(bk)))));
    });
    if (!candRefs.length) continue;

    const members = memberIdx.map(mi => ({ idx: mi, abs: Math.abs(movements[mi].amount) }));
    const totalSum = members.reduce((s, m) => s + m.abs, 0);

    for (const ri of candRefs) {
      const ref = db.entries[ri]; if (ref.matched) continue;
      const tryApply = (subset, refEntry, refIdx) => {
        const n = subset.length;
        refEntry.matched = true; db.entries[refIdx].matched = true;
        for (let p = 0; p < n; p++) {
          const mi = subset[p].idx, base = isRut ? 45 : 35;
          let rs = [isRut ? "rut_exacto" : "nombre_grupo", `pago_parcial(${p + 1}/${n})`, "suma_montos_ok"];
          const bd = parseD(movements[mi].date);
          if (bd && refEntry.date) { const dm = scoreDate(bd, refEntry.date); if (dm.l) rs.push(dm.l); }
          movements[mi].match = { rut: refEntry.rut, rutN: refEntry.rutN, name: refEntry.name, nameN: refEntry.nameN,
            folio: refEntry.folio, date: refEntry.date, amount: refEntry.amount, absAmount: refEntry.absAmount,
            source: refEntry.source, score: base, reasons: rs.join(", "), splitPayment: true,
            splitFolio: refEntry.folio,
            splitLabel: `Pago parcial ${p + 1}/${n} (total factura: ${fmtNum(refEntry.amount)})` };
        }
      };

      if (Math.abs(totalSum - ref.absAmount) / ref.absAmount <= 0.02) { tryApply(members, ref, ri); break; }
      if (members.length <= 6) {
        let found = false;
        for (let mask = 3; mask < (1 << members.length) && !found; mask++) {
          const sub = []; let sum = 0;
          for (let b = 0; b < members.length; b++) if (mask & (1 << b)) { sub.push(members[b]); sum += members[b].abs; }
          if (sub.length >= 2 && Math.abs(sum - ref.absAmount) / ref.absAmount <= 0.02) { tryApply(sub, ref, ri); found = true; }
        }
      }
    }
  }
};

/* ═══════════════════════════════════════════════
   EXCEL EXPORT
   ═══════════════════════════════════════════════ */

const exportXLSX = (results) => {
  const wb = XLSX.utils.book_new();
  const sum = [["CONCILIACIÓN BANCARIA — RESUMEN"], [],
    ["Banco", "Movimientos", "Con Match", "Match Fuerte (≥55)", "Match Parcial", "Pagos Parciales", "Sin Match", "% Acierto"]];
  for (const r of results) {
    const sp = r.movements.filter(m => m.match?.splitPayment).length;
    sum.push([r.label, r.total, r.matched, r.strong, r.partial, sp, r.noMatch, r.total > 0 ? `${Math.round(r.matched / r.total * 100)}%` : "0%"]);
  }
  sum.push([], ["Motor v4 — Prioridad: RUT → Nombre (keyword) → Monto → Fecha"],
    ["Matching por keyword: ignora palabras genéricas (SPA, LTDA, SOCIEDAD, COMERCIAL, etc.)"],
    ["y compara solo palabras distintivas (ej: 'AMERIX' calza con 'SOCIEDAD COMERCIAL AMERIX SPA')"],
    [], ["Match Fuerte (≥55): Alta confianza"], ["Match Parcial (<55): Revisar"], ["Pago Parcial: Varios movimientos suman una factura"]);
  const ws1 = XLSX.utils.aoa_to_sheet(sum);
  ws1["!cols"] = [{ wch: 70 }, { wch: 14 }, { wch: 12 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Resumen");
  for (const r of results) {
    const d = [[`CONCILIACIÓN — ${r.label.toUpperCase()}`], [],
      ["Fecha Banco", "Descripción", "Monto Banco", "RUT Banco", "Nombre Banco", "",
        "Estado", "RUT RCV", "Razón Social RCV", "Folio", "Fecha RCV", "Monto RCV", "Origen", "Criterios", "Score", "Factura Ref. (Pago Parcial)", "Detalle Pago Parcial"]];
    for (const m of r.movements) {
      const isSp = m.match?.splitPayment;
      const st = !m.match ? "SIN MATCH" : isSp ? "PAGO PARCIAL" : m.match.score >= 55 ? "MATCH FUERTE" : "MATCH PARCIAL";
      d.push([fmtDate(m.date), m.desc || "", m.amount, m.rut || "", m.name || "", "",
        st, m.match?.rut || "", m.match?.name || "", m.match?.folio || "",
        m.match?.date ? fmtDate(m.match.date) : "", m.match?.amount || "",
        m.match?.source || "", m.match?.reasons || "", m.match?.score || "",
        m.match?.splitFolio || "", m.match?.splitLabel || ""]);
    }
    const ws = XLSX.utils.aoa_to_sheet(d);
    ws["!cols"] = [{ wch: 12 }, { wch: 36 }, { wch: 14 }, { wch: 13 }, { wch: 26 }, { wch: 1 }, { wch: 14 }, { wch: 13 }, { wch: 30 }, { wch: 10 }, { wch: 12 }, { wch: 13 }, { wch: 14 }, { wch: 30 }, { wch: 6 }, { wch: 26 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, ws, r.label.substring(0, 31));
  }
  XLSX.writeFile(wb, `Conciliacion_${new Date().toISOString().slice(0, 10)}.xlsx`);
};

/* ═══════════════════════════════════════════════
   CONSTANTS & METADATA
   ═══════════════════════════════════════════════ */

const SRC_META = {
  lc: { label: "Libro de Compras (LC)", desc: "RCV Compras del SII", icon: "📥", fields: ["rut", "name", "folio", "date", "amount"] },
  lv: { label: "Libro de Ventas (LV)", desc: "RCV Ventas del SII", icon: "📤", fields: ["rut", "name", "folio", "date", "amount"] },
  remu: { label: "Remuneraciones", desc: "Nómina de sueldos", icon: "👥", fields: ["rut", "name", "amount", "date"] },
  boletas: { label: "Boletas de Honorarios", desc: "Boletas emitidas/recibidas", icon: "🧾", fields: ["rut", "name", "folio", "date", "amount"] },
};
const FL = { rut: "RUT", name: "Razón Social / Nombre", folio: "Folio", date: "Fecha", amount: "Monto", description: "Descripción / Glosa", credit: "Abono / Haber", debit: "Cargo / Debe" };
const BANK_FIELDS = ["date", "description", "rut", "name", "amount", "credit", "debit"];
const STRONG = 55;

/* ═══════════════════════════════════════════════
   COMPONENT
   ═══════════════════════════════════════════════ */

export default function App() {
  const [sources, setSources] = useState({ lc: null, lv: null, remu: null, boletas: null });
  const [banks, setBanks] = useState([{ id: 1, label: "Banco 1" }]);
  const [results, setResults] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [step, setStep] = useState("upload");
  const [expanded, setExpanded] = useState({});
  const [filter, setFilter] = useState("all");

  const toggle = (k) => setExpanded(p => ({ ...p, [k]: !p[k] }));

  const handleSrcFile = async (key, file) => {
    try {
      const p = await parseFile(file);
      const { headers, data, rowCount } = extractSheet(p.workbook, p.sheets[0]);
      setSources(prev => ({ ...prev, [key]: { fileName: p.fileName, sheets: p.sheets, selectedSheet: p.sheets[0], headers, data, rowCount, mapping: detectMapping(headers, "source"), workbook: p.workbook } }));
    } catch (e) { alert("Error: " + e.message); }
  };
  const handleSrcSheet = (key, sheet) => setSources(prev => { const s = prev[key]; if (!s) return prev; const d = extractSheet(s.workbook, sheet); return { ...prev, [key]: { ...s, selectedSheet: sheet, ...d, mapping: detectMapping(d.headers, "source") } }; });
  const handleSrcMap = (key, field, val) => setSources(prev => ({ ...prev, [key]: { ...prev[key], mapping: { ...prev[key].mapping, [field]: val || undefined } } }));
  const clearSrc = (key) => setSources(prev => ({ ...prev, [key]: null }));

  const addBank = () => setBanks(p => [...p, { id: Date.now(), label: `Banco ${p.length + 1}` }]);
  const removeBank = (id) => setBanks(p => p.filter(b => b.id !== id));
  const updateBankLabel = (id, label) => setBanks(p => p.map(b => b.id === id ? { ...b, label } : b));
  const handleBankFile = async (id, file) => {
    try {
      const p = await parseFile(file);
      const { headers, data, rowCount } = extractSheet(p.workbook, p.sheets[0]);
      setBanks(prev => prev.map(b => b.id === id ? { ...b, fileName: p.fileName, sheets: p.sheets, selectedSheet: p.sheets[0], headers, data, rowCount, mapping: detectMapping(headers, "bank"), workbook: p.workbook } : b));
    } catch (e) { alert("Error: " + e.message); }
  };
  const handleBankSheet = (id, sheet) => setBanks(prev => prev.map(b => { if (b.id !== id || !b.workbook) return b; const d = extractSheet(b.workbook, sheet); return { ...b, selectedSheet: sheet, ...d, mapping: detectMapping(d.headers, "bank") }; }));
  const handleBankMap = (id, field, val) => setBanks(prev => prev.map(b => b.id === id ? { ...b, mapping: { ...b.mapping, [field]: val || undefined } } : b));
  const clearBank = (id) => setBanks(prev => prev.map(b => b.id === id ? { id: b.id, label: b.label } : b));

  const runMatching = () => {
    setProcessing(true); setFilter("all");
    setTimeout(() => {
      try {
        const refDB = buildRefDB(sources);
        const allResults = [];
        for (const bank of banks) {
          if (!bank.data || !bank.mapping) continue;
          const mp = bank.mapping, { headers, data } = bank;
          const gi = (f) => mp[f] ? headers.indexOf(mp[f]) : -1;
          refDB.entries.forEach(e => e.matched = false);

          const rawMov = [];
          for (const row of data) {
            const g = (f) => { const i = gi(f); return i >= 0 ? row[i] : null; };
            let amt;
            if (mp.amount) amt = toNum(g("amount"));
            else { const cr = toNum(g("credit")) || 0, db2 = toNum(g("debit")) || 0; amt = cr > 0 && db2 === 0 ? cr : db2 > 0 && cr === 0 ? -db2 : cr - db2; }
            if (isNaN(amt) || amt === 0) continue;
            const descVal = g("description");
            const nameVal = g("name") ? String(g("name")) : "";
            let fallbackText = "";
            if (!nameVal && !descVal) {
              fallbackText = headers.map((_, i) => row[i])
                .filter(c => c != null && typeof c === "string" && c.trim().length > 2)
                .join(" ");
            }
            rawMov.push({ date: g("date"), desc: descVal || fallbackText || null, amount: amt, rut: g("rut") ? String(g("rut")) : "", name: nameVal });
          }

          const sorted = rawMov.map((m, i) => ({ ...m, oi: i }));
          sorted.sort((a, b) => {
            const ar = normalizeRut(a.rut).length >= 3 ? 1 : 0, br = normalizeRut(b.rut).length >= 3 ? 1 : 0;
            if (ar !== br) return br - ar;
            const an = extractKeywords(a.name || a.desc).length, bn2 = extractKeywords(b.name || b.desc).length;
            return bn2 - an;
          });

          const matchArr = new Array(rawMov.length).fill(null);
          for (const m of sorted) {
            const match = findBestMatch(refDB, m.amount, m.date, m.rut, m.name, m.desc);
            if (match) {
              const ri = refDB.entries.findIndex(e => e.rutN === normalizeRut(match.rut) && e.folio === match.folio && e.source === match.source && !e.matched);
              if (ri >= 0) refDB.entries[ri].matched = true;
            }
            matchArr[m.oi] = match;
          }

          const movements = rawMov.map((m, i) => ({ ...m, match: matchArr[i] }));
          findSplitPayments(movements, refDB);

          let matched = 0, strong = 0, partial = 0, splits = 0;
          for (const m of movements) { if (m.match) { matched++; if (m.match.splitPayment) splits++; else if (m.match.score >= STRONG) strong++; else partial++; } }

          allResults.push({ label: bank.label, total: movements.length, matched, strong, partial, splits,
            noMatch: movements.length - matched, pct: movements.length > 0 ? Math.round(matched / movements.length * 100) : 0, movements });
        }
        setResults(allResults); setStep("results");
      } catch (e) { alert("Error: " + e.message); }
      setProcessing(false);
    }, 50);
  };

  const canRun = banks.some(b => b.data) && Object.values(sources).some(s => s?.data);
  const srcCount = Object.values(sources).filter(Boolean).length;
  const bankCount = banks.filter(b => b.data).length;

  /* ═══════════════════════════════════════════════
     SUB-COMPONENTS
     ═══════════════════════════════════════════════ */

  const MappingRow = ({ field, mapping, headers, onChange }) => (
    <div className="flex items-center gap-2">
      <span className="w-28 text-right text-xs font-medium" style={{ color: "#64748b" }}>{FL[field]}</span>
      <select className="flex-1 text-xs px-2 py-1 rounded border" style={{ borderColor: "#e2e8f0", background: "#f8fafc" }}
        value={mapping[field] || ""} onChange={e => onChange(field, e.target.value)}>
        <option value="">— sin asignar —</option>
        {headers.filter(h => !h.startsWith("Col_")).map(h => <option key={h} value={h}>{h}</option>)}
      </select>
      {mapping[field] ? <CheckCircle size={14} color="#059669" /> : <div style={{ width: 14 }} />}
    </div>
  );

  const FileUploadZone = ({ onFile }) => (
    <label className="flex flex-col items-center justify-center border-2 border-dashed rounded-xl cursor-pointer transition-all hover:border-emerald-400 p-4" style={{ borderColor: "#cbd5e1" }}>
      <Upload size={22} color="#94a3b8" />
      <span className="text-xs mt-1" style={{ color: "#64748b" }}>Subir archivo (.xlsx .xls .csv)</span>
      <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
    </label>
  );

  const FileBadge = ({ name, rows, onClear }) => (
    <div className="flex items-center gap-2 px-3 py-2 rounded-lg" style={{ background: "#f1f5f9" }}>
      <FileSpreadsheet size={14} color="#64748b" />
      <span className="text-xs flex-1 truncate" style={{ color: "#334155" }}>{name}</span>
      <span className="text-xs px-1.5 py-0.5 rounded" style={{ background: "#e2e8f0", color: "#64748b" }}>{rows} filas</span>
      <button onClick={onClear} className="hover:opacity-70"><X size={14} color="#94a3b8" /></button>
    </div>
  );

  /* ═══════════════════════════════════════════════
     UPLOAD VIEW
     ═══════════════════════════════════════════════ */

  const renderUpload = () => (
    <div className="space-y-6">
      <div className="rounded-xl p-4" style={{ background: "#eff6ff", border: "1px solid #bfdbfe" }}>
        <div className="text-xs font-semibold mb-1" style={{ color: "#1d4ed8" }}>Motor v4 — Matching por palabra clave</div>
        <div className="text-xs leading-relaxed" style={{ color: "#1e40af" }}>
          <strong>1° RUT</strong> → <strong>2° Palabra clave</strong> (ignora SPA, LTDA, SOCIEDAD, COMERCIAL, etc.) → <strong>3° Monto</strong> → <strong>4° Fecha</strong>
          &nbsp;·&nbsp; Ejemplo: "AMERIX" calza con "SOCIEDAD COMERCIAL AMERIX SPA"
        </div>
      </div>

      <div>
        <h2 className="text-sm font-semibold mb-3 flex items-center gap-2" style={{ color: "#1e293b" }}>
          <span className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold" style={{ background: "#dbeafe", color: "#2563eb" }}>1</span>
          Fuentes de referencia <span className="text-xs font-normal px-2 py-0.5 rounded-full" style={{ background: "#f1f5f9", color: "#64748b" }}>{srcCount}/4</span>
        </h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          {Object.entries(SRC_META).map(([key, meta]) => {
            const s = sources[key];
            return (
              <div key={key} className="rounded-xl border p-4" style={{ borderColor: s ? "#86efac" : "#e2e8f0", background: s ? "#f0fdf4" : "#fff" }}>
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-2">
                    <span>{meta.icon}</span>
                    <div><div className="text-sm font-semibold" style={{ color: "#1e293b" }}>{meta.label}</div><div className="text-xs" style={{ color: "#94a3b8" }}>{meta.desc}</div></div>
                  </div>
                  {s && <CheckCircle size={18} color="#059669" />}
                </div>
                {!s ? <FileUploadZone onFile={f => handleSrcFile(key, f)} /> : (
                  <div className="space-y-2">
                    <FileBadge name={s.fileName} rows={s.rowCount} onClear={() => clearSrc(key)} />
                    {s.sheets.length > 1 && <select className="w-full text-xs px-2 py-1.5 border rounded-lg" style={{ borderColor: "#e2e8f0" }} value={s.selectedSheet} onChange={e => handleSrcSheet(key, e.target.value)}>{s.sheets.map(sh => <option key={sh}>{sh}</option>)}</select>}
                    <button onClick={() => toggle(key)} className="text-xs flex items-center gap-1 hover:opacity-70" style={{ color: "#3b82f6" }}>{expanded[key] ? <ChevronDown size={12} /> : <ChevronRight size={12} />} Mapeo</button>
                    {expanded[key] && <div className="space-y-1.5 pt-1 px-2 rounded-lg" style={{ background: "#fff" }}>{meta.fields.map(f => <MappingRow key={f} field={f} mapping={s.mapping} headers={s.headers} onChange={(f, v) => handleSrcMap(key, f, v)} />)}</div>}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      <div>
        <h2 className="text-sm font-semibold mb-3 flex items-center gap-2" style={{ color: "#1e293b" }}>
          <span className="w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold" style={{ background: "#dbeafe", color: "#2563eb" }}>2</span>
          Cartolas de Banco <span className="text-xs font-normal px-2 py-0.5 rounded-full" style={{ background: "#f1f5f9", color: "#64748b" }}>{bankCount}</span>
        </h2>
        <div className="space-y-3">
          {banks.map(bank => (
            <div key={bank.id} className="rounded-xl border p-4" style={{ borderColor: bank.data ? "#86efac" : "#e2e8f0", background: bank.data ? "#f0fdf4" : "#fff" }}>
              <div className="flex items-center gap-3 mb-3">
                <span className="text-lg">🏦</span>
                <input className="text-sm font-semibold border-b border-transparent hover:border-slate-300 focus:border-blue-500 outline-none px-1 py-0.5 flex-1 bg-transparent" style={{ color: "#1e293b" }} value={bank.label} onChange={e => updateBankLabel(bank.id, e.target.value)} />
                {bank.data && <CheckCircle size={18} color="#059669" />}
                {banks.length > 1 && <button onClick={() => removeBank(bank.id)} className="hover:opacity-70"><Trash2 size={16} color="#f43f5e" /></button>}
              </div>
              {!bank.data ? <FileUploadZone onFile={f => handleBankFile(bank.id, f)} /> : (
                <div className="space-y-2">
                  <FileBadge name={bank.fileName} rows={bank.rowCount} onClear={() => clearBank(bank.id)} />
                  {bank.sheets?.length > 1 && <select className="w-full text-xs px-2 py-1.5 border rounded-lg" style={{ borderColor: "#e2e8f0" }} value={bank.selectedSheet} onChange={e => handleBankSheet(bank.id, e.target.value)}>{bank.sheets.map(sh => <option key={sh}>{sh}</option>)}</select>}
                  <button onClick={() => toggle(`b${bank.id}`)} className="text-xs flex items-center gap-1 hover:opacity-70" style={{ color: "#3b82f6" }}>{expanded[`b${bank.id}`] ? <ChevronDown size={12} /> : <ChevronRight size={12} />} Mapeo</button>
                  {expanded[`b${bank.id}`] && bank.mapping && (
                    <div className="space-y-1.5 pt-1 px-2 rounded-lg" style={{ background: "#fff" }}>
                      <div className="text-xs px-2 py-1.5 rounded" style={{ background: "#fffbeb", color: "#92400e", borderLeft: "3px solid #f59e0b" }}>Columna única → <strong>Monto</strong> · Separadas → <strong>Abono</strong> + <strong>Cargo</strong></div>
                      {BANK_FIELDS.map(f => <MappingRow key={f} field={f} mapping={bank.mapping} headers={bank.headers} onChange={(f, v) => handleBankMap(bank.id, f, v)} />)}
                    </div>
                  )}
                </div>
              )}
            </div>
          ))}
          <button onClick={addBank} className="w-full py-3 rounded-xl border-2 border-dashed flex items-center justify-center gap-2 text-sm font-medium hover:border-blue-400 hover:bg-blue-50" style={{ borderColor: "#cbd5e1", color: "#64748b" }}><Plus size={16} /> Agregar banco</button>
        </div>
      </div>

      <div className="flex items-center justify-between pt-2 pb-4">
        <div className="text-xs" style={{ color: "#94a3b8" }}>{canRun ? `✓ ${bankCount} banco(s) × ${srcCount} fuente(s)` : "Carga al menos 1 fuente y 1 cartola"}</div>
        <button onClick={runMatching} disabled={!canRun || processing} className="px-6 py-3 rounded-xl text-sm font-bold flex items-center gap-2 shadow-lg hover:shadow-xl disabled:opacity-40"
          style={{ background: canRun ? "linear-gradient(135deg, #059669, #047857)" : "#94a3b8", color: "#fff" }}>
          {processing ? <RotateCcw size={16} className="animate-spin" /> : <Play size={16} />}
          {processing ? "Procesando..." : "Ejecutar Conciliación"}
        </button>
      </div>
    </div>
  );

  /* ═══════════════════════════════════════════════
     RESULTS VIEW
     ═══════════════════════════════════════════════ */

  const renderResults = () => {
    if (!results) return null;
    const tM = results.reduce((a, r) => a + r.total, 0), tMa = results.reduce((a, r) => a + r.matched, 0);
    const tS = results.reduce((a, r) => a + r.strong, 0), tSp = results.reduce((a, r) => a + r.splits, 0);

    return (
      <div className="space-y-5">
        <div className="rounded-xl overflow-hidden" style={{ background: "linear-gradient(135deg, #0f172a, #1e293b)" }}>
          <div className="p-5">
            <div className="text-xs font-bold mb-4 tracking-wide" style={{ color: "#94a3b8", textTransform: "uppercase", letterSpacing: "0.08em" }}>Resumen</div>
            <div className="grid grid-cols-5 gap-3">
              {[{ v: tM, l: "Movimientos", c: "#f1f5f9" }, { v: tMa, l: "Con Match", c: "#34d399" }, { v: tS, l: "Fuerte", c: "#6ee7b7" }, { v: tSp, l: "Pago Parcial", c: "#a78bfa" }, { v: tM > 0 ? `${Math.round(tMa / tM * 100)}%` : "0%", l: "Acierto", c: "#fbbf24" }].map((s, i) => (
                <div key={i} className="text-center"><div className="text-xl font-bold" style={{ color: s.c }}>{s.v}</div><div className="text-xs mt-0.5" style={{ color: "#64748b" }}>{s.l}</div></div>
              ))}
            </div>
          </div>
        </div>

        {results.map((r, ri) => {
          const fm = filter === "all" ? r.movements : filter === "strong" ? r.movements.filter(m => m.match && !m.match.splitPayment && m.match.score >= STRONG) : filter === "partial" ? r.movements.filter(m => m.match && !m.match.splitPayment && m.match.score < STRONG) : filter === "split" ? r.movements.filter(m => m.match?.splitPayment) : r.movements.filter(m => !m.match);

          return (
            <div key={ri} className="rounded-xl border overflow-hidden" style={{ borderColor: "#e2e8f0", background: "#fff" }}>
              <div className="p-4 flex items-center justify-between" style={{ borderBottom: "1px solid #e2e8f0" }}>
                <div className="flex items-center gap-3"><span className="text-lg">🏦</span><div><div className="text-sm font-bold" style={{ color: "#1e293b" }}>{r.label}</div><div className="text-xs" style={{ color: "#64748b" }}>{r.total} movimientos</div></div></div>
                <div className="relative w-12 h-12">
                  <svg viewBox="0 0 36 36" className="w-12 h-12 -rotate-90"><path d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831" fill="none" stroke="#e2e8f0" strokeWidth="3" /><path d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831" fill="none" stroke={r.pct >= 50 ? "#059669" : r.pct >= 25 ? "#d97706" : "#e11d48"} strokeWidth="3" strokeDasharray={`${r.pct}, 100`} strokeLinecap="round" /></svg>
                  <div className="absolute inset-0 flex items-center justify-center"><span className="text-xs font-bold" style={{ color: "#1e293b" }}>{r.pct}%</span></div>
                </div>
              </div>

              <div className="grid grid-cols-5 divide-x" style={{ borderBottom: "1px solid #e2e8f0" }}>
                {[{ l: "Fuerte", v: r.strong, c: "#059669", bg: "#ecfdf5" }, { l: "Parcial", v: r.partial, c: "#d97706", bg: "#fffbeb" }, { l: "Pago Parcial", v: r.splits, c: "#7c3aed", bg: "#f5f3ff" }, { l: "Sin Match", v: r.noMatch, c: "#e11d48", bg: "#fff1f2" }, { l: "% Acierto", v: `${r.pct}%`, c: "#2563eb", bg: "#eff6ff" }].map((s, si) => (
                  <div key={si} className="py-2 px-2 text-center" style={{ background: s.bg }}><div className="text-base font-bold" style={{ color: s.c }}>{s.v}</div><div style={{ color: "#64748b", fontSize: 10 }}>{s.l}</div></div>
                ))}
              </div>

              <div className="px-4 py-2 flex gap-1 flex-wrap" style={{ background: "#f8fafc", borderBottom: "1px solid #e2e8f0" }}>
                {[{ k: "all", l: "Todos", n: r.movements.length }, { k: "strong", l: "Fuertes", n: r.strong }, { k: "partial", l: "Parciales", n: r.partial }, { k: "split", l: "Pagos Parciales", n: r.splits }, { k: "none", l: "Sin match", n: r.noMatch }].map(t => (
                  <button key={t.k} onClick={() => setFilter(t.k)} className="px-3 py-1 rounded-lg text-xs font-medium" style={{ background: filter === t.k ? "#1e293b" : "transparent", color: filter === t.k ? "#fff" : "#64748b" }}>{t.l} ({t.n})</button>
                ))}
              </div>

              <div className="overflow-x-auto" style={{ maxHeight: 450 }}>
                <table className="w-full text-xs">
                  <thead className="sticky top-0 z-10"><tr style={{ background: "#f1f5f9" }}>
                    {["Estado", "Fecha", "Desc. Banco", "Monto Banco", "→", "RUT", "Razón Social", "Folio", "Monto RCV", "Origen", "Criterios", "Score", "Nota"].map(h => (
                      <th key={h} className="px-2 py-2 text-left font-semibold whitespace-nowrap" style={{ color: "#475569" }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {fm.slice(0, 300).map((m, mi) => {
                      const isSp = m.match?.splitPayment, isS = !isSp && m.match?.score >= STRONG, isP = !isSp && m.match && m.match.score < STRONG;
                      const bg = isSp ? "#f5f3ff" : isS ? "#f0fdf4" : isP ? "#fffbeb" : !m.match ? "#fff5f5" : "#fff";
                      return (
                        <tr key={mi} style={{ background: bg, borderBottom: "1px solid #f1f5f9" }}>
                          <td className="px-2 py-1.5">
                            {isSp ? <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded font-medium" style={{ background: "#ede9fe", color: "#5b21b6", fontSize: 10 }}><Layers size={10} />Split</span>
                              : isS ? <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded font-medium" style={{ background: "#dcfce7", color: "#166534", fontSize: 10 }}><CheckCircle size={10} />Fuerte</span>
                              : isP ? <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded font-medium" style={{ background: "#fef3c7", color: "#92400e", fontSize: 10 }}><AlertTriangle size={10} />Parcial</span>
                              : <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded font-medium" style={{ background: "#fecdd3", color: "#9f1239", fontSize: 10 }}><XCircle size={10} />—</span>}
                          </td>
                          <td className="px-2 py-1.5 whitespace-nowrap" style={{ color: "#475569" }}>{fmtDate(m.date)}</td>
                          <td className="px-2 py-1.5 max-w-32 truncate" title={m.desc || m.name} style={{ color: "#334155" }}>{m.desc || m.name || "—"}</td>
                          <td className="px-2 py-1.5 text-right font-mono whitespace-nowrap" style={{ color: m.amount < 0 ? "#e11d48" : "#059669" }}>{fmtNum(m.amount)}</td>
                          <td className="px-1 text-center" style={{ color: "#cbd5e1" }}>→</td>
                          <td className="px-2 py-1.5 font-mono" style={{ color: "#475569" }}>{m.match?.rut || ""}</td>
                          <td className="px-2 py-1.5 max-w-32 truncate" title={m.match?.name} style={{ color: "#334155" }}>{m.match?.name || ""}</td>
                          <td className="px-2 py-1.5 font-mono" style={{ color: "#475569" }}>{m.match?.folio || ""}</td>
                          <td className="px-2 py-1.5 text-right font-mono whitespace-nowrap" style={{ color: "#475569" }}>{m.match ? fmtNum(m.match.amount) : ""}</td>
                          <td className="px-2 py-1.5">{m.match?.source && <span className="px-1.5 py-0.5 rounded" style={{ fontSize: 10, background: m.match.source === "LC" ? "#dbeafe" : m.match.source === "LV" ? "#dcfce7" : m.match.source === "Remuneraciones" ? "#fce7f3" : "#f3e8ff", color: m.match.source === "LC" ? "#1d4ed8" : m.match.source === "LV" ? "#166534" : m.match.source === "Remuneraciones" ? "#9d174d" : "#7c3aed" }}>{m.match.source}</span>}</td>
                          <td className="px-2 py-1.5" style={{ color: "#64748b", fontSize: 10 }}>{m.match?.reasons || ""}</td>
                          <td className="px-2 py-1.5 text-center font-mono font-bold" style={{ color: isSp ? "#7c3aed" : isS ? "#059669" : isP ? "#d97706" : "#cbd5e1" }}>{m.match?.score || ""}</td>
                          <td className="px-2 py-1.5" style={{ fontSize: 10, color: "#7c3aed", fontWeight: isSp ? 600 : 400 }}>{m.match?.splitLabel || ""}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                {fm.length > 300 && <div className="text-center py-2 text-xs" style={{ color: "#94a3b8" }}>300/{fm.length} en pantalla</div>}
                {fm.length === 0 && <div className="text-center py-8 text-xs" style={{ color: "#94a3b8" }}>Sin movimientos en este filtro</div>}
              </div>
            </div>
          );
        })}

        <div className="flex items-center justify-between py-4">
          <button onClick={() => setStep("upload")} className="px-4 py-2.5 rounded-xl text-sm flex items-center gap-2 border hover:bg-slate-50" style={{ borderColor: "#e2e8f0", color: "#475569" }}>← Volver</button>
          <button onClick={() => exportXLSX(results)} className="px-6 py-3 rounded-xl text-sm font-bold flex items-center gap-2 shadow-lg hover:shadow-xl" style={{ background: "linear-gradient(135deg, #1e40af, #3b82f6)", color: "#fff" }}><Download size={16} /> Descargar Excel</button>
        </div>
      </div>
    );
  };

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#f8fafc", minHeight: "100vh" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');*{font-family:'DM Sans','Segoe UI',system-ui,sans-serif;box-sizing:border-box}::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:3px}select{appearance:auto}`}</style>
      <div style={{ background: "linear-gradient(135deg, #0f172a, #1e293b)", borderBottom: "1px solid #334155" }}>
        <div className="max-w-6xl mx-auto px-5 py-4 flex items-center justify-between">
          <div><h1 className="text-lg font-bold tracking-tight" style={{ color: "#f1f5f9" }}>📊 Conciliación Bancaria <span className="text-xs font-normal" style={{ color: "#64748b" }}>v4</span></h1><p className="text-xs mt-0.5" style={{ color: "#64748b" }}>RUT → Keyword → Monto → Fecha · Pagos parciales</p></div>
          <div className="flex gap-2">
            {[{ k: "upload", l: "⚙️ Config" }, { k: "results", l: "📈 Resultados" }].map(s => (
              <button key={s.k} onClick={() => (s.k === "results" && !results) ? null : setStep(s.k)} className="px-4 py-2 rounded-lg text-xs font-semibold"
                style={{ background: step === s.k ? "#fff" : "rgba(255,255,255,0.08)", color: step === s.k ? "#0f172a" : "#94a3b8", opacity: s.k === "results" && !results ? 0.3 : 1 }}>{s.l}</button>
            ))}
          </div>
        </div>
      </div>
      <div className="max-w-6xl mx-auto px-5 py-6">{step === "upload" ? renderUpload() : renderResults()}</div>
    </div>
  );
}
