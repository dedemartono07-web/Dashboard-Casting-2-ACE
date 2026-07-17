import { useState, useEffect, useRef } from "react";
import { LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

const C = {
  bg: "#060C18", card: "#0C1526", border: "#172340",
  green: "#22C55E", blue: "#60A5FA", amber: "#FBBF24",
  red: "#F87171", purple: "#A78BFA", orange: "#FB923C",
  txt: "#F1F5F9", txt2: "#94A3B8", txt3: "#475569",
};
const PC = ["#22C55E", "#60A5FA", "#FBBF24"];
const PL = ["SAFETY", "QUALITY", "PRODUKSI"];
const PSUB = ["PRODUCTIVITY", "PERFORMANCE FOREMAN", "KWH & GENTANI", "MATERIAL"];
const PHOTO_SECS = ["Melting", "Molding", "Finishing"];

const DEF = {
  monthStr: "Mei 2025",
  safety: {
    activities: [
      { name: "KYT", plan: 10, aktual: 7, remain: 2 },
      { name: "HYT", plan: 6, aktual: 2, remain: 0 },
      { name: "Yokoten", plan: 4, aktual: 2, remain: 0 },
    ],
  },
  quality: {
    melting: [
      { name: "Marubo", count: 2, types: [], w5: { What: "Cacat permukaan cor", When: "Mei 2025", Where: "Furnace A", Who: "Operator Melting", Why: "Komposisi tidak sesuai spec" } },
      { name: "Komposisi", count: 3, types: [], w5: { What: "Kadar elemen melebihi batas", When: "Mei 2025", Where: "Furnace B", Who: "Operator Melting", Why: "Penambahan tidak terkontrol" } },
      { name: "Material", count: 0, types: [], w5: null },
    ],
    finishing: [{ name: "Flowout Finishing", count: 3, types: ["Porosity", "Cold Shut", "Shrinkage"], w5: { What: "Cairan keluar cetakan", When: "Mei 2025", Where: "Line Finishing 1", Who: "Operator Finishing", Why: "Tekanan tuang terlalu tinggi" } }],
    core: [{ name: "Flowout Core", count: 0, types: [], w5: null }],
  },
  prodSummary: {
    ace1: { moldTotal: 46799, moldJamAvg: 139.7, prodPctAvg: 97.0, kwhTotal: 19302, gentaniFurnace: 570.3, gentaniCore: 43.9, gentaniFinishing: 69.9, mhMolding: 3.28, mhFinishing: 3.06 },
    ace2: { moldTotal: 52244, moldJamAvg: 134.1, prodPctAvg: 93.1, kwhTotal: 24728, gentaniFurnace: 553.0, gentaniCore: 43.9, gentaniFinishing: 69.9, mhMolding: 3.28, mhFinishing: 3.06 },
    pcs: { Finishing: 753046, EDP: 208994, Core: 267415 },
  },
  foreman: {
    melting: { Sudimin: 479, "Eko Winardi": 371 },
    molding: { Sunarno: 95, Sonang: 135, "Agus M.": 173, Angga: 295 },
  },
  mainMat: {
    ace1: [["FC",766580],["FCD450",868227],["FCD500",313671],["FCD600",44592],["High MN",252506],["Low MN",477517],["Chips Inhouse",162996],["Chips Outhouse",110283],["Kubik High MN",9934],["Kubik Low MN",373251],["Kawul",0],["Garbage",0]],
    ace2: [["FC",1960284],["FCD450",106592],["FCD500",87392],["FCD600",18588],["High MN",739946],["Low MN",158648],["Chips Inhouse",1004632],["Chips Outhouse",249074],["Kubik High MN",25810],["Kubik Low MN",50635],["Kawul",12744],["Garbage",34368]],
  },
  alloyMat: {
    ace1: [["Asbury",51886],["Elmag",0],["Lamet",4561],["Fe-Si",818],["SiC",21023],["Si-Mn",1712],["Sn",155],["S",546],["Cu",1205],["Sb",0]],
    ace2: [["Asbury",19820],["Elmag",101],["Lamet",1362],["Fe-Si",2645],["SiC",19794],["Si-Mn",4633],["Sn",462],["S",1677],["Cu",3528],["Sb",0]],
  },
  trend: {
    mold: [{d:"03",a1:140.2,a2:145.5},{d:"04",a1:143.3,a2:145.5},{d:"05",a1:140.4,a2:152.1},{d:"08",a1:149.4,a2:151.6},{d:"09",a1:149.7,a2:146.5},{d:"10",a1:149.1,a2:144.1},{d:"11",a1:150.6,a2:136.7},{d:"12",a1:135.9,a2:142.5},{d:"13",a1:144.5,a2:154.1},{d:"15",a1:146.5,a2:150.6},{d:"17",a1:138.5,a2:139.9},{d:"18",a1:143.3,a2:152.1},{d:"19",a1:142.6,a2:150.5},{d:"20",a1:151.5,a2:141.1},{d:"22",a1:149.1,a2:149.7}],
    gentani: [{d:"02",a1:565,a2:551},{d:"03",a1:576,a2:558},{d:"04",a1:563,a2:568},{d:"05",a1:544,a2:565},{d:"08",a1:572,a2:586},{d:"09",a1:555,a2:569},{d:"10",a1:580,a2:577},{d:"11",a1:571,a2:533},{d:"12",a1:596,a2:560},{d:"15",a1:552,a2:557},{d:"17",a1:597,a2:563},{d:"18",a1:598,a2:542},{d:"19",a1:590,a2:559},{d:"20",a1:593,a2:556},{d:"22",a1:580,a2:567}],
    forMelt: [{d:"03",s:56,e:12},{d:"04",s:25,e:5},{d:"05",s:61,e:18},{d:"08",s:26,e:26},{d:"09",s:19,e:23},{d:"10",s:15,e:13},{d:"11",s:12,e:12},{d:"12",s:41,e:32},{d:"15",s:30,e:38},{d:"17",s:24,e:80},{d:"18",s:34,e:40},{d:"19",s:63,e:20},{d:"20",s:16,e:13},{d:"22",s:25,e:16}],
    forMold: [{d:"03",sn:3,so:42,ag:10,an:13},{d:"04",sn:23,so:2,ag:11,an:10},{d:"05",sn:8,so:13,ag:8,an:25},{d:"08",sn:2,so:6,ag:9,an:13},{d:"09",sn:5,so:6,ag:16,an:22},{d:"10",sn:2,so:10,ag:5,an:17},{d:"11",sn:2,so:null,ag:12,an:10},{d:"12",sn:7,so:20,ag:7,an:20},{d:"15",sn:3,so:6,ag:3,an:6},{d:"17",sn:2,so:2,ag:5,an:39},{d:"18",sn:13,so:3,ag:5,an:11},{d:"19",sn:5,so:null,ag:11,an:8},{d:"20",sn:10,so:1,ag:6,an:25},{d:"22",sn:1,so:2,ag:6,an:17}],
  },
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const fmt = (n) => n == null ? "-" : Number(n).toLocaleString("id-ID", { maximumFractionDigits: 1 });
const pctColor = (v, t, low = false) => {
  if (!v || !t) return C.txt3;
  const r = v / t;
  if (low) return r <= 1 ? C.green : r <= 1.15 ? C.amber : C.red;
  return r >= 1 ? C.green : r >= 0.9 ? C.amber : C.red;
};

function useBreakpoint() {
  const calc = (w) => w < 600 ? "mobile" : w < 1024 ? "tablet" : "tv";
  const [bp, setBp] = useState(() => calc(typeof window !== "undefined" ? window.innerWidth : 1920));
  useEffect(() => {
    const h = () => setBp(calc(window.innerWidth));
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return bp;
}

// ─── STORAGE ─────────────────────────────────────────────────────────────────
const SK = {
  data: "ace-dashboard-data",
  update: "ace-last-update",
  photo: (s) => "ace-photos-" + s,
};
const sSet = async (k, v) => {
  try { await window.storage.set(k, JSON.stringify(v)); return true; } catch (e) { return false; }
};
const sGet = async (k) => {
  try { const r = await window.storage.get(k); return r ? JSON.parse(r.value) : null; } catch (e) { return null; }
};

// ─── EXCEL PARSER ────────────────────────────────────────────────────────────
function parseExcel(ab) {
  const wb = XLSX.read(ab, { type: "array", cellDates: true });
  const wsR = wb.Sheets["Rekap Summary"];
  if (!wsR) throw new Error("Sheet 'Rekap Summary' tidak ditemukan");

  const rekapRows = XLSX.utils.sheet_to_json(wsR, { header: 1, raw: true, defval: null });
  const sections = {};
  let cur = "root";
  rekapRows.forEach((row) => {
    if (!row || row.length < 2) return;
    const lbl = row[1] != null ? String(row[1]).trim() : null;
    const val = row[2] != null && row[2] !== "" ? row[2] : null;
    if (!lbl) return;
    if (val === null) {
      cur = lbl;
      if (!sections[cur]) sections[cur] = {};
    } else {
      if (!sections[cur]) sections[cur] = {};
      sections[cur][lbl] = val;
    }
  });

  const fS = (...pats) => {
    for (const p of pats) {
      const kws = Array.isArray(p) ? p : [p];
      const key = Object.keys(sections).find((k) => kws.every((kw) => k.toUpperCase().includes(kw.toUpperCase())));
      if (key) return sections[key];
    }
    return {};
  };
  const gS = (sec, key, def = 0) => { const v = sec[key]; return v != null ? Number(v) : def; };

  const saf = fS("SAFETY"), qual = fS("QUALITY"), prod = fS("PRODUKSI");
  const for_ = fS(["PERFORMANCE"], ["FOREMAN"], ["FORMAN"]);
  const kwh1 = fS(["ACE 1", "KWH"]), kwh2 = fS(["ACE 2", "KWH"]);
  const gent = fS("GENTANI");
  const mat1 = fS(["MAIN", "ACE 1"]), mat2 = fS(["MAIN", "ACE 2"]);
  const al1 = fS(["ALLOY", "ACE 1"]), al2 = fS(["ALLOY", "ACE 2"]);

  const wsD = wb.Sheets["Data Harian"] || wb.Sheets[Object.keys(wb.Sheets).find((s) => s.includes("Data Harian")) || ""];
  const MONTHS = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  const fDay = (v) => {
    if (v instanceof Date) return String(v.getDate()).padStart(2, "0");
    if (typeof v === "number") { const d = new Date(Math.round((v - 25569) * 864e5)); return String(d.getUTCDate()).padStart(2, "0"); }
    return String(v).split(/[-/]/)[0].trim().padStart(2, "0");
  };
  const fMonthStr = (v) => {
    if (v instanceof Date) return MONTHS[v.getMonth()] + " " + v.getFullYear();
    if (typeof v === "number") { const d = new Date(Math.round((v - 25569) * 864e5)); return MONTHS[d.getUTCMonth()] + " " + d.getUTCFullYear(); }
    return "";
  };
  const nv = (v) => v != null && !isNaN(+v) && +v !== 0 ? Math.round(+v * 100) / 100 : null;
  const trend = { mold: [], gentani: [], forMelt: [], forMold: [] };
  let monthStr = "";

  if (wsD) {
    const dRows = XLSX.utils.sheet_to_json(wsD, { header: 1, raw: true, defval: null });
    dRows.slice(3).forEach((row) => {
      if (!row || !row[0]) return;
      if (typeof row[0] === "string" && row[0].includes("TOTAL")) return;
      const d = fDay(row[0]);
      if (!monthStr && row[0]) monthStr = fMonthStr(row[0]);
      if (nv(row[17]) !== null || nv(row[18]) !== null)
        trend.mold.push({ d, a1: nv(row[17]) ? Math.round(nv(row[17]) * 10) / 10 : null, a2: nv(row[18]) ? Math.round(nv(row[18]) * 10) / 10 : null });
      if (nv(row[38]) !== null || nv(row[39]) !== null)
        trend.gentani.push({ d, a1: nv(row[38]) ? Math.round(nv(row[38])) : null, a2: nv(row[39]) ? Math.round(nv(row[39])) : null });
      if (nv(row[24]) !== null || nv(row[25]) !== null)
        trend.forMelt.push({ d, s: nv(row[24]), e: nv(row[25]) });
      if (nv(row[26]) !== null || nv(row[27]) !== null)
        trend.forMold.push({ d, sn: nv(row[26]), so: nv(row[27]), ag: nv(row[28]), an: nv(row[29]) });
    });
  }

  const toP = (v) => v < 2 ? Math.round(v * 100 * 10) / 10 : Math.round(v * 10) / 10;
  const getM = (s, k) => { const v = s[k]; return v != null ? Math.round(+v) : 0; };

  return {
    monthStr: monthStr || "Mei 2025",
    safety: {
      activities: [
        { name: "KYT", plan: gS(saf, "KYT Plan (Total)"), aktual: gS(saf, "KYT Actual (Total)"), remain: gS(saf, "KYT Remaining") },
        { name: "HYT", plan: gS(saf, "HYT Plan"), aktual: gS(saf, "HYT Actual"), remain: gS(saf, "HYT Remaining") },
        { name: "Yokoten", plan: gS(saf, "Yokoten Plan"), aktual: gS(saf, "Yokoten Actual"), remain: gS(saf, "Yokoten Remaining") },
      ],
    },
    quality: {
      melting: [
        { name: "Marubo", count: gS(qual, "Abnormal Marubo"), types: [], w5: null },
        { name: "Komposisi", count: gS(qual, "Abnormal Komposisi"), types: [], w5: null },
        { name: "Material", count: gS(qual, "Abnormal Material"), types: [], w5: null },
      ],
      finishing: [{ name: "Flowout Finishing", count: gS(qual, "Flowout Finishing"), types: [], w5: null }],
      core: [{ name: "Flowout Core", count: gS(qual, "Flowout Core"), types: [], w5: null }],
    },
    prodSummary: {
      ace1: { moldTotal: gS(prod, "Mold ACE1"), moldJamAvg: Math.round(gS(prod, "Avg MoldJam ACE1") * 10) / 10, prodPctAvg: toP(gS(prod, "Avg Prod% ACE1")), kwhTotal: Math.round(gS(kwh1, "Total KWh Furnace") || gS(kwh1, "Total KWh Remelting")), gentaniFurnace: Math.round(gS(gent, "Gentani Furnace ACE 1") * 10) / 10, gentaniCore: Math.round(gS(gent, "Core") * 10) / 10, gentaniFinishing: Math.round(gS(gent, "Finishing") * 10) / 10, mhMolding: Math.round(gS(gent, "Man Hour Molding") * 100) / 100, mhFinishing: Math.round(gS(gent, "Man Hour Finishing") * 100) / 100 },
      ace2: { moldTotal: gS(prod, "Mold ACE2"), moldJamAvg: Math.round(gS(prod, "Avg MoldJam ACE2") * 10) / 10, prodPctAvg: toP(gS(prod, "Avg Prod% ACE2")), kwhTotal: Math.round(gS(kwh2, "Total KWh Furnace") || gS(kwh2, "Total KWh Remelting")), gentaniFurnace: Math.round(gS(gent, "Gentani Furnace ACE 2") * 10) / 10, gentaniCore: Math.round(gS(gent, "Core") * 10) / 10, gentaniFinishing: Math.round(gS(gent, "Finishing") * 10) / 10, mhMolding: Math.round(gS(gent, "Man Hour Molding") * 100) / 100, mhFinishing: Math.round(gS(gent, "Man Hour Finishing") * 100) / 100 },
      pcs: { Finishing: gS(prod, "Total PCS Finishing"), EDP: gS(prod, "Total PCS EDP"), Core: gS(prod, "Total PCS Core") },
    },
    foreman: {
      melting: { Sudimin: gS(for_, "Sudimin (min)") || gS(for_, "Avg PourWait Sudimin"), "Eko Winardi": gS(for_, "Eko Winardi (min)") || gS(for_, "Avg PourWait Eko Winardi") },
      molding: { Sunarno: gS(for_, "Sunarno (Mold)") || gS(for_, "Avg LossMold Sunarno"), Sonang: gS(for_, "Sonang (Mold)") || gS(for_, "Avg LossMold Sonang"), "Agus M.": gS(for_, "Agus Maryanta (Mold)") || gS(for_, "Avg LossMold Agus M."), Angga: gS(for_, "Angga (Mold)") || gS(for_, "Avg LossMold Angga") },
    },
    mainMat: {
      ace1: [["FC",getM(mat1,"FC")],["FCD450",getM(mat1,"FCD450")],["FCD500",getM(mat1,"FCD500")],["FCD600",getM(mat1,"FCD600")],["High MN",getM(mat1,"High Mn")],["Low MN",getM(mat1,"Low Mn")],["Chips Inhouse",getM(mat1,"Inhouse")],["Chips Outhouse",getM(mat1,"Outhouse")],["Kubik High MN",getM(mat1,"HighMn")||getM(mat1,"HighMn Scrap")],["Kubik Low MN",getM(mat1,"LowMn")||getM(mat1,"LowMn Scrap")],["Kawul",getM(mat1,"Kawul")],["Garbage",getM(mat1,"Garbage")]],
      ace2: [["FC",getM(mat2,"FC")],["FCD450",getM(mat2,"FCD450")],["FCD500",getM(mat2,"FCD500")],["FCD600",getM(mat2,"FCD600")],["High MN",getM(mat2,"High Mn")],["Low MN",getM(mat2,"Low Mn")],["Chips Inhouse",getM(mat2,"Inhouse")],["Chips Outhouse",getM(mat2,"Outhouse")],["Kubik High MN",getM(mat2,"HighMn")||getM(mat2,"HighMn Scrap")],["Kubik Low MN",getM(mat2,"LowMn")||getM(mat2,"LowMn Scrap")],["Kawul",getM(mat2,"Kawul")],["Garbage",getM(mat2,"Garbage")]],
    },
    alloyMat: {
      ace1: [["Asbury",getM(al1,"Asbury")],["Elmag",getM(al1,"Elmag")],["Lamet",getM(al1,"Lamet")],["Fe-Si",getM(al1,"Fe-Si")],["SiC",getM(al1,"SiC")],["Si-Mn",getM(al1,"Si-Mn")],["Sn",getM(al1,"Sn")],["S",getM(al1,"S")],["Cu",getM(al1,"Cu")],["Sb",getM(al1,"Sb")]],
      ace2: [["Asbury",getM(al2,"Asbury")],["Elmag",getM(al2,"Elmag")],["Lamet",getM(al2,"Lamet")],["Fe-Si",getM(al2,"Fe-Si")],["SiC",getM(al2,"SiC")],["Si-Mn",getM(al2,"Si-Mn")],["Sn",getM(al2,"Sn")],["S",getM(al2,"S")],["Cu",getM(al2,"Cu")],["Sb",getM(al2,"Sb")]],
    },
    trend,
  };
}

// ─── IMAGE RESIZE ────────────────────────────────────────────────────────────
async function resizeImg(file, maxPx = 900, q = 0.78) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onerror = () => rej(new Error("Gagal membaca file"));
    r.onload = (e) => {
      const img = new Image();
      img.onerror = () => rej(new Error("Format gambar tidak valid"));
      img.onload = () => {
        let w = img.width, h = img.height;
        if (w > maxPx || h > maxPx) {
          if (w > h) { h = Math.round(h * maxPx / w); w = maxPx; }
          else { w = Math.round(w * maxPx / h); h = maxPx; }
        }
        const cv = document.createElement("canvas");
        cv.width = w; cv.height = h;
        cv.getContext("2d").drawImage(img, 0, 0, w, h);
        res(cv.toDataURL("image/jpeg", q));
      };
      img.src = e.target.result;
    };
    r.readAsDataURL(file);
  });
}

// ─── SHARED UI ───────────────────────────────────────────────────────────────
function Card({ children, style }) {
  return (
    <div style={{ background: C.card, border: "1px solid " + C.border, borderRadius: 10, padding: 14, ...style }}>
      {children}
    </div>
  );
}

function SecTitle({ children, color }) {
  return (
    <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: 2, color: color || C.txt2, textTransform: "uppercase", marginBottom: 10, paddingBottom: 7, borderBottom: "1px solid " + C.border }}>
      {children}
    </div>
  );
}

function PBar({ value, max, color }) {
  const pct = Math.min((value / max) * 100, 100);
  return (
    <div style={{ background: C.border, borderRadius: 3, height: 6, overflow: "hidden", marginTop: 4 }}>
      <div style={{ width: pct + "%", height: "100%", background: color || C.green, borderRadius: 3, transition: "width .5s" }} />
    </div>
  );
}

function Clock() {
  const [t, setT] = useState(new Date());
  useEffect(() => {
    const id = setInterval(() => setT(new Date()), 1000);
    return () => clearInterval(id);
  }, []);
  const p = (n) => String(n).padStart(2, "0");
  const D = ["Min","Sen","Sel","Rab","Kam","Jum","Sab"];
  const M = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
  return (
    <div style={{ textAlign: "right" }}>
      <div style={{ fontSize: 22, fontWeight: 800, color: C.txt, letterSpacing: 2 }}>
        {p(t.getHours())}:{p(t.getMinutes())}:{p(t.getSeconds())}
      </div>
      <div style={{ fontSize: 10, color: C.txt2 }}>
        {D[t.getDay()]}, {t.getDate()} {M[t.getMonth()]} {t.getFullYear()}
      </div>
    </div>
  );
}

function Toast({ msg, type }) {
  if (!msg) return null;
  const col = type === "success" ? C.green : type === "error" ? C.red : C.blue;
  const icon = type === "success" ? "✅" : type === "error" ? "❌" : "⏳";
  return (
    <div style={{ position: "fixed", bottom: 20, left: "50%", transform: "translateX(-50%)", background: C.card, border: "1px solid " + col, borderRadius: 8, padding: "10px 20px", color: C.txt, fontSize: 13, fontWeight: 600, zIndex: 300, boxShadow: "0 4px 20px rgba(0,0,0,.6)", maxWidth: "90vw", textAlign: "center", whiteSpace: "nowrap" }}>
      {icon} {msg}
    </div>
  );
}

// ─── PHOTO GRID ──────────────────────────────────────────────────────────────
function PhotoGrid({ section, photos, onAdd, onDel, mob }) {
  const ref = useRef();
  const MAX = 6;
  const W = mob ? 80 : 118, H = mob ? 60 : 82;

  const handleFiles = async (e) => {
    const files = Array.from(e.target.files);
    for (const f of files) {
      if (photos.length >= MAX) break;
      try {
        const b = await resizeImg(f);
        onAdd(section, { data: b, name: f.name });
      } catch (err) {
        alert("Gagal upload " + f.name + ": " + err.message);
      }
    }
    e.target.value = "";
  };

  return (
    <div style={{ marginBottom: 16 }}>
      <div style={{ fontSize: 11, fontWeight: 700, color: C.green, letterSpacing: 1, textTransform: "uppercase", marginBottom: 8 }}>{section}</div>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-start" }}>
        {photos.map((ph, i) => (
          <div key={i} style={{ position: "relative", width: W, height: H, borderRadius: 8, overflow: "hidden", border: "2px solid " + C.border, flexShrink: 0 }}>
            <img src={ph.data} alt={ph.name} style={{ width: "100%", height: "100%", objectFit: "cover" }} />
            <button onClick={() => onDel(section, i)} style={{ position: "absolute", top: 3, right: 3, background: "rgba(0,0,0,.8)", border: "none", color: C.red, borderRadius: "50%", width: 20, height: 20, cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center", padding: 0, fontWeight: 800, lineHeight: 1 }}>
              x
            </button>
          </div>
        ))}
        {photos.length < MAX && (
          <>
            <input ref={ref} type="file" accept="image/*" multiple style={{ display: "none" }} onChange={handleFiles} />
            <div onClick={() => ref.current && ref.current.click()} style={{ width: W, height: H, borderRadius: 8, border: "2px dashed " + C.green + "66", background: "#0A1628", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: C.green, fontSize: 10, cursor: "pointer", flexShrink: 0, gap: 4 }}>
              <div style={{ fontSize: 22, fontWeight: 300, lineHeight: 1 }}>+</div>
              <div>Tambah Foto</div>
            </div>
          </>
        )}
        {photos.length === 0 && (
          <div style={{ fontSize: 11, color: C.txt3, alignSelf: "center", fontStyle: "italic" }}>Belum ada foto</div>
        )}
      </div>
    </div>
  );
}

// ─── SAFETY PILLAR ────────────────────────────────────────────────────────────
function SafetyPillar({ bp, data, photos, onAdd, onDel }) {
  const mob = bp === "mobile";
  return (
    <div style={{ display: "flex", flexDirection: mob ? "column" : "row", gap: 14, height: "100%", overflow: "auto" }}>
      <Card style={{ flex: mob ? undefined : 1 }}>
        <SecTitle color={C.green}>Kegiatan Safety — {data.monthStr}</SecTitle>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 280 }}>
            <thead>
              <tr>
                {["Kegiatan", "Planning", "Aktual", "Remaining"].map((h) => (
                  <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontSize: 11, color: C.txt2, fontWeight: 600, borderBottom: "1px solid " + C.border }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.safety.activities.map((r) => (
                <tr key={r.name}>
                  <td style={{ padding: "14px 10px", color: C.txt, fontWeight: 700, fontSize: mob ? 14 : 16 }}>{r.name}</td>
                  <td style={{ padding: "14px 10px", color: C.blue, fontWeight: 800, fontSize: mob ? 20 : 26 }}>{r.plan}</td>
                  <td style={{ padding: "14px 10px" }}>
                    <span style={{ color: C.green, fontWeight: 800, fontSize: mob ? 20 : 26 }}>{r.aktual}</span>
                    <PBar value={r.aktual} max={r.plan} color={C.green} />
                  </td>
                  <td style={{ padding: "14px 10px", color: r.remain > 0 ? C.amber : C.green, fontWeight: 800, fontSize: mob ? 20 : 26 }}>{r.remain}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      <Card style={{ flex: mob ? undefined : 1.6, overflow: "auto" }}>
        <SecTitle color={C.green}>Best Practice</SecTitle>
        {PHOTO_SECS.map((sec) => (
          <PhotoGrid key={sec} section={sec} photos={photos[sec] || []} onAdd={onAdd} onDel={onDel} mob={mob} />
        ))}
      </Card>
    </div>
  );
}

// ─── QUALITY PILLAR ───────────────────────────────────────────────────────────
function AbnCard({ name, count, types, w5 }) {
  const [open, setOpen] = useState(false);
  const col = count === 0 ? C.green : C.red;
  return (
    <div style={{ background: "#080F1F", border: "1px solid " + (count > 0 ? C.red + "55" : C.border), borderRadius: 8, padding: 12, marginBottom: 8 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
        <span style={{ color: C.txt, fontWeight: 600, fontSize: 13, flex: 1 }}>{name}</span>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 26, fontWeight: 900, color: col, minWidth: 30, textAlign: "center" }}>{count}</span>
          <span style={{ fontSize: 10, color: C.txt3 }}>kali</span>
          {w5 && (
            <button onClick={() => setOpen(!open)} style={{ background: open ? C.blue : "transparent", border: "1px solid " + C.blue, color: open ? "#000" : C.blue, borderRadius: 4, padding: "3px 8px", fontSize: 10, cursor: "pointer", flexShrink: 0 }}>
              5W {open ? "▲" : "▼"}
            </button>
          )}
        </div>
      </div>
      {types && types.length > 0 && (
        <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginTop: 6 }}>
          {types.map((t) => (
            <span key={t} style={{ background: "#1A2240", border: "1px solid " + C.orange, color: C.orange, fontSize: 10, padding: "2px 6px", borderRadius: 4 }}>{t}</span>
          ))}
        </div>
      )}
      {open && w5 && (
        <div style={{ marginTop: 10, padding: 10, background: "#0D1929", borderRadius: 6, borderLeft: "3px solid " + C.blue }}>
          {Object.entries(w5).map(([k, v]) => (
            <div key={k} style={{ display: "flex", gap: 8, marginBottom: 4 }}>
              <span style={{ color: C.blue, fontWeight: 700, fontSize: 11, width: 40, flexShrink: 0 }}>{k}</span>
              <span style={{ color: C.txt2, fontSize: 11 }}>{v}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function QualityPillar({ bp, data }) {
  const mob = bp === "mobile";
  return (
    <div style={{ display: "flex", flexDirection: mob ? "column" : "row", gap: 14, height: "100%", overflow: "auto" }}>
      <Card style={{ flex: 1, overflow: "auto" }}>
        <SecTitle color={C.blue}>Melting — Abnormal</SecTitle>
        {data.quality.melting.map((x) => <AbnCard key={x.name} {...x} />)}
      </Card>
      <Card style={{ flex: 1, overflow: "auto" }}>
        <SecTitle color={C.blue}>Finishing — Flowout</SecTitle>
        {data.quality.finishing.map((x) => <AbnCard key={x.name} {...x} />)}
      </Card>
      <Card style={{ flex: 1, overflow: "auto" }}>
        <SecTitle color={C.blue}>Core — Flowout</SecTitle>
        {data.quality.core.map((x) => <AbnCard key={x.name} {...x} />)}
      </Card>
    </div>
  );
}

// ─── PRODUCTION ──────────────────────────────────────────────────────────────
const cStyle = { background: "#0C1526", border: "1px solid #172340", borderRadius: 8, fontSize: 11 };
const cMargin = { top: 5, right: 10, left: -15, bottom: 0 };

function MetricBox({ label, value, target, unit, low }) {
  const col = pctColor(value, target, low);
  return (
    <div style={{ flex: 1, background: "#080F1F", border: "1px solid " + col + "33", borderRadius: 8, padding: "10px 12px", minWidth: 0 }}>
      <div style={{ fontSize: 10, color: C.txt2, textTransform: "uppercase", letterSpacing: 1, marginBottom: 6, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}</div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 4, flexWrap: "wrap" }}>
        <span style={{ fontSize: 24, fontWeight: 900, color: col }}>{fmt(value)}</span>
        <span style={{ fontSize: 10, color: C.txt3 }}>{unit}</span>
      </div>
      <div style={{ fontSize: 10, color: C.txt3, marginTop: 2 }}>Target: {fmt(target)} {unit}</div>
      <PBar value={value} max={target * 1.1} color={col} />
    </div>
  );
}

function ProductivitySub({ bp, data }) {
  const mob = bp === "mobile";
  const s = data.prodSummary;
  const td = data.trend.mold.map((r) => ({ day: r.d, "ACE 1": r.a1, "ACE 2": r.a2, Target: 145 }));
  const totalPcs = Object.values(s.pcs).reduce((a, b) => a + b, 0);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12, height: "100%", overflow: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: mob ? "1fr" : "1fr 1fr", gap: 12 }}>
        {[["ACE 1", s.ace1], ["ACE 2", s.ace2]].map(([lbl, d]) => (
          <Card key={lbl}>
            <SecTitle color={C.amber}>{lbl} — Productivity</SecTitle>
            <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
              <MetricBox label="Mold/Jam (Avg)" value={d.moldJamAvg} target={145} unit="mold" />
              <MetricBox label="Produktivitas" value={d.prodPctAvg} target={97} unit="%" />
            </div>
            <div style={{ background: "#080F1F", borderRadius: 8, padding: "10px 12px", border: "1px solid " + C.border }}>
              <div style={{ fontSize: 10, color: C.txt2, marginBottom: 4, textTransform: "uppercase", letterSpacing: 1 }}>Total Mold Baik</div>
              <span style={{ fontSize: 26, fontWeight: 900, color: C.amber }}>{fmt(d.moldTotal)}</span>
              <span style={{ fontSize: 10, color: C.txt3 }}> mold</span>
            </div>
          </Card>
        ))}
      </div>
      <Card>
        <SecTitle color={C.amber}>Total PCS — {data.monthStr}</SecTitle>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {[["Finishing", s.pcs.Finishing, C.green], ["EDP", s.pcs.EDP, C.blue], ["Core", s.pcs.Core, C.purple], ["TOTAL", totalPcs, C.amber]].map(([k, v, col]) => (
            <div key={k} style={{ flex: 1, minWidth: 60, textAlign: "center", background: "#0A1628", borderRadius: 6, padding: "10px 4px", border: "1px solid " + col + "33" }}>
              <div style={{ fontSize: 9, color: C.txt2, marginBottom: 4 }}>{k}</div>
              <div style={{ fontSize: mob ? 15 : 20, fontWeight: 800, color: col }}>{fmt(v)}</div>
            </div>
          ))}
        </div>
      </Card>
      <Card style={{ flex: 1 }}>
        <SecTitle color={C.amber}>Trend Mold/Jam Harian — Target: 145</SecTitle>
        <ResponsiveContainer width="100%" height={mob ? 150 : 185}>
          <LineChart data={td} margin={cMargin}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="day" tick={{ fill: C.txt3, fontSize: 9 }} />
            <YAxis domain={[100, 175]} tick={{ fill: C.txt3, fontSize: 9 }} />
            <Tooltip contentStyle={cStyle} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            <Line type="monotone" dataKey="ACE 1" stroke={C.green} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
            <Line type="monotone" dataKey="ACE 2" stroke={C.blue} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
            <Line type="monotone" dataKey="Target" stroke={C.amber} strokeWidth={1} strokeDasharray="6 4" dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

function ForemanSub({ bp, data }) {
  const mob = bp === "mobile";
  const fm = data.foreman;
  const md = data.trend.forMelt.map((r) => ({ day: r.d, Sudimin: r.s, "Eko Winardi": r.e }));
  const od = data.trend.forMold.map((r) => ({ day: r.d, Sunarno: r.sn, Sonang: r.so, "Agus M.": r.ag, Angga: r.an }));
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12, height: "100%", overflow: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: mob ? "1fr" : "1fr 1fr", gap: 12 }}>
        <Card>
          <SecTitle color={C.amber}>Total Loss Time Pour Wait — Melting (menit)</SecTitle>
          <div style={{ display: "flex", gap: 8 }}>
            {Object.entries(fm.melting).map(([k, v]) => (
              <div key={k} style={{ flex: 1, background: "#0A1628", borderRadius: 6, padding: 10, border: "1px solid " + C.border, textAlign: "center" }}>
                <div style={{ fontSize: 10, color: C.txt2, marginBottom: 4 }}>{k}</div>
                <div style={{ fontSize: 26, fontWeight: 800, color: C.amber }}>{v}</div>
                <div style={{ fontSize: 9, color: C.txt3 }}>menit total</div>
              </div>
            ))}
          </div>
        </Card>
        <Card>
          <SecTitle color={C.amber}>Total Loss Mold — Molding</SecTitle>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
            {Object.entries(fm.molding).map(([k, v]) => (
              <div key={k} style={{ flex: 1, minWidth: 60, background: "#0A1628", borderRadius: 6, padding: "10px 6px", border: "1px solid " + C.border, textAlign: "center" }}>
                <div style={{ fontSize: 9, color: C.txt2, marginBottom: 4 }}>{k}</div>
                <div style={{ fontSize: 22, fontWeight: 800, color: C.orange }}>{v}</div>
                <div style={{ fontSize: 9, color: C.txt3 }}>mold</div>
              </div>
            ))}
          </div>
        </Card>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: mob ? "1fr" : "1fr 1fr", gap: 12, flex: 1 }}>
        <Card>
          <SecTitle color={C.amber}>Trend Harian — Foreman Melting (menit)</SecTitle>
          <div style={{ fontSize: 10, color: C.txt3, marginBottom: 6 }}>Makin kecil = makin baik</div>
          <ResponsiveContainer width="100%" height={mob ? 145 : 210}>
            <LineChart data={md} margin={cMargin}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
              <XAxis dataKey="day" tick={{ fill: C.txt3, fontSize: 9 }} />
              <YAxis tick={{ fill: C.txt3, fontSize: 9 }} />
              <Tooltip contentStyle={cStyle} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line type="monotone" dataKey="Sudimin" stroke={C.green} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
              <Line type="monotone" dataKey="Eko Winardi" stroke={C.blue} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
            </LineChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SecTitle color={C.amber}>Trend Harian — Foreman Molding (loss mold)</SecTitle>
          <div style={{ fontSize: 10, color: C.txt3, marginBottom: 6 }}>Makin kecil = makin baik</div>
          <ResponsiveContainer width="100%" height={mob ? 145 : 210}>
            <LineChart data={od} margin={cMargin}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
              <XAxis dataKey="day" tick={{ fill: C.txt3, fontSize: 9 }} />
              <YAxis tick={{ fill: C.txt3, fontSize: 9 }} />
              <Tooltip contentStyle={cStyle} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line type="monotone" dataKey="Sunarno" stroke={C.green} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
              <Line type="monotone" dataKey="Sonang" stroke={C.blue} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
              <Line type="monotone" dataKey="Agus M." stroke={C.purple} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
              <Line type="monotone" dataKey="Angga" stroke={C.orange} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
            </LineChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  );
}

function KwhGentaniSub({ bp, data }) {
  const mob = bp === "mobile";
  const s = data.prodSummary;
  const gd = data.trend.gentani.map((r) => ({ day: r.d, "ACE 1": r.a1, "ACE 2": r.a2, Target: 550 }));

  const renderGentani = (d) => {
    const items = [
      { k: "Furnace", v: d.gentaniFurnace, t: 550, u: "KWH/Ton", low: true },
      { k: "Core", v: d.gentaniCore, t: 55, u: "Pcs/Jam", low: false },
      { k: "Finishing", v: d.gentaniFinishing, t: 60, u: "Pcs/Jam", low: false },
      { k: "MH Molding", v: d.mhMolding, t: 8, u: "Jam/Org", low: true },
      { k: "MH Finishing", v: d.mhFinishing, t: 7, u: "Jam/Org", low: true },
    ];
    return items.map((g) => {
      const col = pctColor(g.v, g.t, g.low);
      return (
        <div key={g.k} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 7 }}>
          <span style={{ fontSize: 11, color: C.txt2, width: 90, flexShrink: 0 }}>{g.k}</span>
          <span style={{ fontSize: 16, fontWeight: 800, color: col, width: 50 }}>{fmt(g.v)}</span>
          <span style={{ fontSize: 10, color: C.txt3, width: 72, flexShrink: 0 }}>/{g.t} {g.u}</span>
          <div style={{ flex: 1 }}><PBar value={g.v} max={g.t * 1.2} color={col} /></div>
        </div>
      );
    });
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12, height: "100%", overflow: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: mob ? "1fr" : "1fr 1fr", gap: 12 }}>
        {[["ACE 1", s.ace1], ["ACE 2", s.ace2]].map(([lbl, d]) => (
          <Card key={lbl}>
            <SecTitle color={C.blue}>{lbl} — KWH & Gentani</SecTitle>
            <div style={{ background: "#080F1F", borderRadius: 8, padding: "10px 12px", border: "1px solid " + C.blue + "33", marginBottom: 10 }}>
              <div style={{ fontSize: 10, color: C.txt2, marginBottom: 4, textTransform: "uppercase", letterSpacing: 1 }}>Total KWH Furnace</div>
              <span style={{ fontSize: 24, fontWeight: 800, color: C.blue }}>{fmt(d.kwhTotal)}</span>
              <span style={{ fontSize: 10, color: C.txt3 }}> KWH</span>
            </div>
            {renderGentani(d)}
          </Card>
        ))}
      </div>
      <Card style={{ flex: 1 }}>
        <SecTitle color={C.purple}>Trend Gentani Furnace (KWH/Ton) — Target: 550</SecTitle>
        <ResponsiveContainer width="100%" height={mob ? 145 : 185}>
          <LineChart data={gd} margin={cMargin}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="day" tick={{ fill: C.txt3, fontSize: 9 }} />
            <YAxis domain={[440, 650]} tick={{ fill: C.txt3, fontSize: 9 }} />
            <Tooltip contentStyle={cStyle} />
            <Legend wrapperStyle={{ fontSize: 11 }} />
            <Line type="monotone" dataKey="ACE 1" stroke={C.green} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
            <Line type="monotone" dataKey="ACE 2" stroke={C.blue} strokeWidth={2} dot={{ r: 2 }} connectNulls={false} />
            <Line type="monotone" dataKey="Target" stroke={C.amber} strokeWidth={1} strokeDasharray="6 4" dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

function MatRow({ name, value, color }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 10px", background: "#080F1F", borderRadius: 6 }}>
      <span style={{ fontSize: 11, color: C.txt2 }}>{name}</span>
      <span style={{ fontSize: 13, fontWeight: 800, color: color }}>{value > 0 ? fmt(value) : "—"}</span>
    </div>
  );
}

function MaterialSub({ bp, data }) {
  const mob = bp === "mobile";
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12, height: "100%", overflow: "auto" }}>
      <div style={{ display: "grid", gridTemplateColumns: mob ? "1fr" : "1fr 1fr", gap: 12 }}>
        <Card>
          <SecTitle color={C.amber}>Main Material ACE 1 (Kg)</SecTitle>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
            {data.mainMat.ace1.map(([n, v]) => <MatRow key={n} name={n} value={v} color={C.amber} />)}
          </div>
        </Card>
        <Card>
          <SecTitle color={C.amber}>Main Material ACE 2 (Kg)</SecTitle>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
            {data.mainMat.ace2.map(([n, v]) => <MatRow key={n} name={n} value={v} color={C.amber} />)}
          </div>
        </Card>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: mob ? "1fr" : "1fr 1fr", gap: 12 }}>
        <Card>
          <SecTitle color={C.purple}>Alloy Material ACE 1 (Kg)</SecTitle>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
            {data.alloyMat.ace1.map(([n, v]) => <MatRow key={n} name={n} value={v} color={C.purple} />)}
          </div>
        </Card>
        <Card>
          <SecTitle color={C.purple}>Alloy Material ACE 2 (Kg)</SecTitle>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 4 }}>
            {data.alloyMat.ace2.map(([n, v]) => <MatRow key={n} name={n} value={v} color={C.purple} />)}
          </div>
        </Card>
      </div>
    </div>
  );
}

function ProductionPillar({ bp, data, subDur }) {
  const [sub, setSub] = useState(0);
  const mob = bp === "mobile";
  useEffect(() => {
    const id = setInterval(() => setSub((s) => (s + 1) % PSUB.length), subDur * 1000);
    return () => clearInterval(id);
  }, [subDur]);

  return (
    <div style={{ height: "100%", display: "flex", flexDirection: "column", gap: 10 }}>
      <div style={{ display: "flex", gap: 4, overflowX: "auto", paddingBottom: 2, flexShrink: 0 }}>
        {PSUB.map((n, i) => (
          <button key={i} onClick={() => setSub(i)} style={{ flex: mob ? undefined : 1, whiteSpace: "nowrap", padding: mob ? "6px 8px" : "7px 12px", borderRadius: 6, cursor: "pointer", border: "1px solid " + (i === sub ? C.amber : C.border), background: i === sub ? C.amber + "22" : "transparent", color: i === sub ? C.amber : C.txt2, fontSize: mob ? 9 : 10, fontWeight: 700, letterSpacing: 1 }}>
            {n}
          </button>
        ))}
      </div>
      <div style={{ flex: 1, overflow: "hidden" }}>
        {sub === 0 && <ProductivitySub bp={bp} data={data} />}
        {sub === 1 && <ForemanSub bp={bp} data={data} />}
        {sub === 2 && <KwhGentaniSub bp={bp} data={data} />}
        {sub === 3 && <MaterialSub bp={bp} data={data} />}
      </div>
    </div>
  );
}

// ─── SETTINGS ────────────────────────────────────────────────────────────────
function Settings({ cfg, setCfg, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.75)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: C.card, border: "2px solid " + C.green, borderRadius: 14, padding: 28, minWidth: 300, maxWidth: 420, width: "90%" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: C.txt }}>Settings</div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", color: C.txt2, fontSize: 22, cursor: "pointer" }}>x</button>
        </div>
        {[{ lbl: "Durasi Ganti Pilar (detik)", key: "pillarDur" }, { lbl: "Durasi Sub Produksi (detik)", key: "subDur" }].map(({ lbl, key }) => (
          <div key={key} style={{ marginBottom: 18 }}>
            <label style={{ fontSize: 12, color: C.txt2, display: "block", marginBottom: 8 }}>{lbl}</label>
            <div style={{ display: "flex", gap: 8 }}>
              {[10, 15, 30, 60].map((v) => (
                <button key={v} onClick={() => setCfg((c) => ({ ...c, [key]: v }))} style={{ flex: 1, padding: "8px 0", borderRadius: 6, cursor: "pointer", border: "1px solid " + (cfg[key] === v ? C.green : C.border), background: cfg[key] === v ? C.green + "22" : "transparent", color: cfg[key] === v ? C.green : C.txt2, fontSize: 13, fontWeight: 700 }}>
                  {v}s
                </button>
              ))}
            </div>
          </div>
        ))}
        <div style={{ fontSize: 11, color: C.txt3, marginTop: 8, padding: 10, background: "#080F1F", borderRadius: 6 }}>
          Klik di luar untuk menutup. Perubahan langsung berlaku.
        </div>
      </div>
    </div>
  );
}

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App() {
  const bp = useBreakpoint();
  const mob = bp === "mobile";
  const [pillar, setPillar] = useState(0);
  const [data, setData] = useState(DEF);
  const [photos, setPhotos] = useState({ Melting: [], Molding: [], Finishing: [] });
  const [lastUpdate, setLastUpdate] = useState(null);
  const [cfg, setCfg] = useState({ pillarDur: 15, subDur: 20 });
  const [showCfg, setShowCfg] = useState(false);
  const [toast, setToast] = useState(null);
  const [loading, setLoading] = useState(true);
  const xlsRef = useRef();

  useEffect(() => {
    async function load() {
      try {
        const savedData = await sGet(SK.data);
        if (savedData) setData(savedData);
        const savedUpdate = await sGet(SK.update);
        if (savedUpdate) setLastUpdate(savedUpdate);
        const ph = { Melting: [], Molding: [], Finishing: [] };
        for (const sec of PHOTO_SECS) {
          const p = await sGet(SK.photo(sec));
          if (p && Array.isArray(p)) ph[sec] = p;
        }
        setPhotos(ph);
      } catch (e) {
        console.error("Load error:", e);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, []);

  useEffect(() => {
    const id = setInterval(() => setPillar((p) => (p + 1) % 3), cfg.pillarDur * 1000);
    return () => clearInterval(id);
  }, [cfg.pillarDur]);

  const showToast = (msg, type, dur = 3500) => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), dur);
  };

  const handleExcel = async (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    e.target.value = "";
    showToast("Membaca file Excel...", "info", 10000);
    try {
      const parsed = parseExcel(await file.arrayBuffer());
      setData(parsed);
      const now = new Date().toLocaleString("id-ID");
      setLastUpdate(now);
      await sSet(SK.data, parsed);
      await sSet(SK.update, now);
      showToast("Data berhasil diupdate dari " + file.name, "success", 4000);
    } catch (err) {
      console.error(err);
      showToast("Gagal: " + err.message, "error", 5000);
    }
  };

  const handleAddPhoto = (section, photo) => {
    setPhotos((prev) => {
      const updated = { ...prev, [section]: [...(prev[section] || []), photo] };
      sSet(SK.photo(section), updated[section]);
      return updated;
    });
    showToast("Foto ditambahkan ke Best Practice " + section, "success", 2500);
  };

  const handleDelPhoto = (section, idx) => {
    setPhotos((prev) => {
      const updated = { ...prev, [section]: prev[section].filter((_, i) => i !== idx) };
      sSet(SK.photo(section), updated[section]);
      return updated;
    });
  };

  const pc = PC[pillar];

  if (loading) {
    return (
      <div style={{ background: C.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter',system-ui,sans-serif" }}>
        <div style={{ textAlign: "center", color: C.txt2 }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>⚙️</div>
          <div>Memuat dashboard...</div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ background: C.bg, minHeight: "100vh", fontFamily: "'Inter','Segoe UI',system-ui,sans-serif", display: "flex", flexDirection: "column", padding: mob ? 8 : 12, boxSizing: "border-box", color: C.txt }}>
      {/* HEADER */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: mob ? "10px 12px" : "12px 18px", background: C.card, borderRadius: 10, marginBottom: 10, border: "1px solid " + pc + "44", borderTop: "3px solid " + pc, flexWrap: mob ? "wrap" : "nowrap", gap: 8, transition: "border-color .5s", flexShrink: 0 }}>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: mob ? 13 : 20, fontWeight: 900, color: C.txt, letterSpacing: mob ? 1 : 3 }}>DASHBOARD MONITORING ACE</div>
          <div style={{ fontSize: mob ? 9 : 11, color: C.txt2, marginTop: 2 }}>
            CASTING 2 — LINE ACE  •  {data.monthStr}
            {lastUpdate && <span style={{ marginLeft: 8, color: C.txt3, fontSize: 9 }}>• Update: {lastUpdate}</span>}
          </div>
        </div>
        {!mob && (
          <div style={{ display: "flex", gap: 6 }}>
            {PL.map((n, i) => (
              <button key={i} onClick={() => setPillar(i)} style={{ padding: "7px 18px", borderRadius: 8, cursor: "pointer", border: "2px solid " + (i === pillar ? PC[i] : C.border), background: i === pillar ? PC[i] + "22" : "transparent", color: i === pillar ? PC[i] : C.txt2, fontSize: 11, fontWeight: 700, letterSpacing: 2, transition: "all .3s" }}>
                {n}
              </button>
            ))}
          </div>
        )}
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
          {!mob && <Clock />}
          <input ref={xlsRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleExcel} />
          <button onClick={() => xlsRef.current && xlsRef.current.click()} style={{ background: C.green + "22", border: "1px solid " + C.green, color: C.green, borderRadius: 8, padding: mob ? "8px 10px" : "8px 14px", cursor: "pointer", fontSize: mob ? 11 : 12, fontWeight: 700, whiteSpace: "nowrap" }}>
            {mob ? "Update" : "Update Data Excel"}
          </button>
          <button onClick={() => setShowCfg(true)} style={{ background: "transparent", border: "1px solid " + C.border, color: C.txt2, borderRadius: 8, padding: "8px 10px", cursor: "pointer", fontSize: 16, lineHeight: 1 }}>
            ⚙
          </button>
        </div>
      </div>

      {/* PILLAR LABEL */}
      {!mob && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, flexShrink: 0 }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: pc + "22", border: "1px solid " + pc, borderRadius: 20, padding: "4px 14px", fontSize: 11, fontWeight: 700, color: pc, letterSpacing: 2 }}>
            <span style={{ width: 8, height: 8, borderRadius: "50%", background: pc, display: "inline-block" }} />
            PILAR {pillar + 1} — {PL[pillar]}
          </div>
          <div style={{ flex: 1, height: 1, background: "linear-gradient(to right," + pc + "44,transparent)" }} />
        </div>
      )}

      {/* CONTENT */}
      <div style={{ flex: 1, overflow: "hidden", marginBottom: mob ? 54 : 0 }}>
        {pillar === 0 && <SafetyPillar bp={bp} data={data} photos={photos} onAdd={handleAddPhoto} onDel={handleDelPhoto} />}
        {pillar === 1 && <QualityPillar bp={bp} data={data} />}
        {pillar === 2 && <ProductionPillar bp={bp} data={data} subDur={cfg.subDur} />}
      </div>

      {/* PROGRESS BAR (desktop) */}
      {!mob && (
        <div style={{ display: "flex", gap: 4, marginTop: 8, flexShrink: 0 }}>
          {[0, 1, 2].map((i) => (
            <div key={i} onClick={() => setPillar(i)} style={{ flex: 1, height: 4, borderRadius: 2, background: i === pillar ? PC[i] : C.border, cursor: "pointer", transition: "background .4s" }} />
          ))}
        </div>
      )}

      {/* BOTTOM NAV (mobile) */}
      {mob && (
        <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: C.card, borderTop: "1px solid " + C.border, display: "flex", zIndex: 100 }}>
          {PL.map((n, i) => (
            <button key={i} onClick={() => setPillar(i)} style={{ flex: 1, padding: "10px 4px", border: "none", background: i === pillar ? PC[i] + "22" : "transparent", color: i === pillar ? PC[i] : C.txt2, fontSize: 9, fontWeight: 700, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 3 }}>
              <span style={{ fontSize: 18 }}>{i === 0 ? "S" : i === 1 ? "Q" : "P"}</span>
              {n}
            </button>
          ))}
        </div>
      )}

      {showCfg && <Settings cfg={cfg} setCfg={setCfg} onClose={() => setShowCfg(false)} />}
      <Toast msg={toast && toast.msg} type={toast && toast.type} />
    </div>
  );
}
