let rows = [];
let densityRows = [];
let activeTank = 1;
let currentCalculation = null;
let historyDateFilter = localDateString();
const tanks = [1, 2, 3, 4];
const HISTORY_KEY = "oilSoundingHistoryV1";
const tankState = {
  1: { ullage: 0, t1: 0, t2: 0, t3: 0 },
  2: { ullage: 0, t1: 0, t2: 0, t3: 0 },
  3: { ullage: 0, t1: 0, t2: 0, t3: 0 },
  4: { ullage: 0, t1: 0, t2: 0, t3: 0 }
};

const $ = id => document.getElementById(id);

function fmt(n, d = 3) {
  return Number.isFinite(n)
    ? n.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d })
    : "-";
}

function numeric(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : NaN;
  if (v == null || v === "") return NaN;
  const s = String(v).trim().replace(/\s/g, "");
  if (!s) return NaN;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  return Number(s.replace(/\./g, "").replace(",", "."));
}

function localDateString(date = new Date()) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function localTimeString(date = new Date()) {
  return date.toLocaleTimeString("id-ID", { hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false });
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return "-";
  const [y, m, d] = dateStr.split("-");
  return `${d}-${m}-${y}`;
}

function getHistory() {
  try {
    const data = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
    if (!Array.isArray(data)) return [];
    // Migrasi record lama agar tidak dianggap record baru.
    data.forEach(r => {
      if (!r.signature) r.signature = [r.date, r.tank, r.ullage, r.t1, r.t2, r.t3].join("|");
    });
    return data;
  } catch {
    return [];
  }
}

function setHistory(data) {
  localStorage.setItem(HISTORY_KEY, JSON.stringify(data));
}

function buildTankTabs() {
  const el = $("tankTabs");
  el.innerHTML = tanks.map(n =>
    `<button class="tank-tab ${n === activeTank ? "active" : ""}" data-tank="${n}">ST ${n}</button>`
  ).join("");
  el.querySelectorAll(".tank-tab").forEach(btn => {
    btn.addEventListener("click", () => switchTank(Number(btn.dataset.tank)));
  });
}

function readInputs() {
  return {
    ullage: Number($("ullage").value) || 0,
    t1: Number($("t1").value) || 0,
    t2: Number($("t2").value) || 0,
    t3: Number($("t3").value) || 0
  };
}

function writeInputs(state) {
  $("ullage").value = state.ullage;
  $("t1").value = state.t1;
  $("t2").value = state.t2;
  $("t3").value = state.t3;
}

function saveActiveTankState() {
  tankState[activeTank] = readInputs();
}

async function switchTank(n) {
  saveActiveTankState();
  activeTank = n;
  rows = [];
  buildTankTabs();
  $("tankTitle").textContent = `OST ${n} - Oil Sounding Tank ${n}`;
  writeInputs(tankState[n]);
  clearCalculated();
  $("calibrationBody").innerHTML = "";
  $("excelStatus").textContent = `Membaca data/ST ${n}.xlsx...`;
  await loadExcel(n);
}

async function loadExcel(tankNo) {
  try {
    const res = await fetch(`data/ST ${tankNo}.xlsx`, { cache: "no-store" });
    if (!res.ok) throw new Error(`File ST ${tankNo}.xlsx belum tersedia`);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    if (!wb.SheetNames.length) throw new Error("Workbook tidak memiliki sheet");
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    rows = parseCalibrationSheet(raw);
    if (!rows.length) throw new Error("Kolom Ullage/Volume tidak terbaca dari Excel");
    renderTable();
    $("excelStatus").textContent = `✓ ST ${tankNo}.xlsx berhasil dibaca: ${rows.length} data kalibrasi`;
    calculate();
  } catch (e) {
    rows = [];
    $("excelStatus").textContent = `✕ ${e.message}. Letakkan file tersebut di folder data.`;
    calculate();
  }
}

async function loadDensity() {
  try {
    const res = await fetch("data/density.xlsx", { cache: "no-store" });
    if (!res.ok) throw new Error("File density.xlsx belum tersedia");
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    if (!wb.SheetNames.length) throw new Error("Workbook density tidak memiliki sheet");
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    densityRows = parseDensitySheet(raw);
    if (!densityRows.length) throw new Error("Kolom Temperature/Density tidak terbaca dari Excel");
    renderDensityTable();
    $("densityStatus").textContent = `✓ density.xlsx berhasil dibaca: ${densityRows.length} data density`;
    calculate();
  } catch (e) {
    densityRows = [];
    $("densityBody").innerHTML = "";
    $("densityStatus").textContent = `✕ ${e.message}. Masukkan tabel Temperature & Density ke data/density.xlsx.`;
    calculate();
  }
}

function parseDensitySheet(raw) {
  const found = [];
  const headerRows = Math.min(raw.length, 10);
  for (let r = 0; r < headerRows; r++) {
    for (let c = 0; c < (raw[r] || []).length; c++) {
      const tempHeader = String(raw[r][c] ?? "").trim().toLowerCase();
      const densityHeader = String(raw[r][c + 1] ?? "").trim().toLowerCase();
      const factorHeader = String(raw[r][c + 2] ?? "").trim().toLowerCase();
      if ((tempHeader.includes("temp") || tempHeader.includes("suhu")) && densityHeader.includes("density")) {
        for (let rr = r + 1; rr < raw.length; rr++) {
          const t = numeric(raw[rr]?.[c]);
          const d = numeric(raw[rr]?.[c + 1]);
          const f = numeric(raw[rr]?.[c + 2]);
          if (Number.isFinite(t) && Number.isFinite(d)) {
            found.push({ temperature: t, density: d, factor: Number.isFinite(f) ? f : null });
          }
        }
      }
    }
  }
  const unique = new Map();
  found.forEach(x => unique.set(x.temperature, x));
  return [...unique.values()].sort((a, b) => a.temperature - b.temperature);
}

function renderDensityTable() {
  $("densityBody").innerHTML = densityRows.map(r => `
    <tr><td>${fmt(r.temperature, 1)}</td><td>${fmt(r.density, 4)}</td><td>${Number.isFinite(r.factor) ? fmt(r.factor, 7) : "-"}</td></tr>
  `).join("");
}

function findDensity(temp) {
  if (!densityRows.length || !Number.isFinite(temp)) return null;
  let nearest = densityRows[0];
  let best = Math.abs(temp - nearest.temperature);
  for (const r of densityRows) {
    const d = Math.abs(temp - r.temperature);
    if (d < best) { best = d; nearest = r; }
  }
  return nearest;
}

function parseCalibrationSheet(raw) {
  const parsed = [];
  const headerRowsToCheck = Math.min(raw.length, 5);
  const ullageColumns = [];
  for (let r = 0; r < headerRowsToCheck; r++) {
    for (let c = 0; c < (raw[r] || []).length; c++) {
      if (String(raw[r][c] ?? "").trim().toLowerCase() === "ullage" && !ullageColumns.some(x => x.col === c)) {
        ullageColumns.push({ row: r, col: c });
      }
    }
  }
  for (const block of ullageColumns) {
    const c = block.col;
    for (let r = block.row + 2; r < raw.length; r++) {
      const u = numeric(raw[r]?.[c]);
      const v = numeric(raw[r]?.[c + 1]);
      if (Number.isFinite(u) && Number.isFinite(v)) parsed.push({ ullage: u, volume: v });
    }
  }
  const unique = new Map();
  for (const item of parsed) unique.set(item.ullage, item.volume);
  return [...unique.entries()].map(([ullage, volume]) => ({ ullage: Number(ullage), volume: Number(volume) })).sort((a,b)=>a.ullage-b.ullage);
}

function findVolume(u) {
  if (!rows.length || u < rows[0].ullage || u > rows[rows.length - 1].ullage) return null;
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].ullage === u) {
      const next = rows[i + 1] || null;
      const diff = next ? Math.abs((next.volume - rows[i].volume) / (next.ullage - rows[i].ullage)) : 0;
      return { volume: rows[i].volume, base: rows[i], next, diff };
    }
    if (i < rows.length - 1 && u > rows[i].ullage && u < rows[i + 1].ullage) {
      const a = rows[i], b = rows[i + 1];
      const step = (b.volume - a.volume) / (b.ullage - a.ullage);
      return { volume: a.volume + (u-a.ullage)*step, base:a, next:b, diff:Math.abs(step) };
    }
  }
  return null;
}

function renderTable() {
  $("calibrationBody").innerHTML = rows.map((r, i) => {
    const lcm = i < rows.length - 1 ? Math.abs((rows[i+1].volume-r.volume)/(rows[i+1].ullage-r.ullage)) : 0;
    return `<tr><td>${r.ullage}</td><td>${fmt(r.volume)}</td><td>${fmt(lcm)}</td></tr>`;
  }).join("");
}

function clearCalculated() {
  ["avgTemp","tableTemp","density","actualUllage","baseUllage","difference","baseVolume","litrePerCm","volumeCorrection","oilVolume","massBeforeFactor","vmt","factor"].forEach(id => $(id).textContent = "-");
  currentCalculation = null;
}

function calculate() {
  const u = Number($("ullage").value);
  const temps = [$("t1").value,$("t2").value,$("t3").value].map(Number).filter(Number.isFinite);
  const validTemps = temps.filter(t => t > 0);
  let avg = NaN;
  let densityData = null;

  if (validTemps.length) {
    avg = validTemps.reduce((a,b)=>a+b,0)/validTemps.length;
    $("avgTemp").textContent = fmt(avg,2);
    densityData = findDensity(avg);
    if (densityData) {
      $("tableTemp").textContent = fmt(densityData.temperature,1);
      $("density").textContent = fmt(densityData.density,4);
      $("factor").textContent = Number.isFinite(densityData.factor) ? fmt(densityData.factor,7) : "-";
    } else {
      $("tableTemp").textContent = $("density").textContent = $("factor").textContent = "-";
    }
  } else {
    $("avgTemp").textContent = "-";
    $("tableTemp").textContent = $("density").textContent = $("factor").textContent = "-";
  }

  if (!Number.isFinite(u) || !rows.length || u === 0) {
    ["actualUllage","baseUllage","difference","baseVolume","litrePerCm","volumeCorrection","oilVolume","massBeforeFactor","vmt"].forEach(id => $(id).textContent = "-");
    currentCalculation = null;
    return null;
  }

  const x = findVolume(u);
  if (!x) {
    ["actualUllage","baseUllage","difference","baseVolume","litrePerCm","volumeCorrection","oilVolume","massBeforeFactor","vmt"].forEach(id => $(id).textContent = "-");
    currentCalculation = null;
    return null;
  }

  const correction = (u - x.base.ullage) * x.diff;
  $("actualUllage").textContent = fmt(u,1);
  $("baseUllage").textContent = fmt(x.base.ullage,1);
  $("difference").textContent = fmt(u-x.base.ullage,1);
  $("baseVolume").textContent = fmt(x.base.volume);
  $("litrePerCm").textContent = fmt(x.diff);
  $("volumeCorrection").textContent = fmt(correction);
  $("oilVolume").textContent = fmt(x.volume);

  let mass = NaN, vmt = NaN;
  if (densityData && Number.isFinite(densityData.density)) {
    mass = x.volume * densityData.density / 1000;
    $("massBeforeFactor").textContent = fmt(mass);
    if (Number.isFinite(densityData.factor)) {
      vmt = mass * densityData.factor;
      $("vmt").textContent = fmt(vmt);
    } else $("vmt").textContent = "-";
  } else {
    $("massBeforeFactor").textContent = $("vmt").textContent = "-";
  }

  currentCalculation = {
    tank: activeTank,
    ullage: u,
    t1: Number($("t1").value) || 0,
    t2: Number($("t2").value) || 0,
    t3: Number($("t3").value) || 0,
    avgTemp: Number.isFinite(avg) ? avg : null,
    tableTemp: densityData?.temperature ?? null,
    density: densityData?.density ?? null,
    factor: densityData?.factor ?? null,
    baseUllage: x.base.ullage,
    difference: u - x.base.ullage,
    baseVolume: x.base.volume,
    litrePerCm: x.diff,
    volumeCorrection: correction,
    oilVolume: x.volume,
    massBeforeFactor: Number.isFinite(mass) ? mass : null,
    vmt: Number.isFinite(vmt) ? vmt : null
  };
  return currentCalculation;
}

function historyInputSignature(calc, dateStr) {
  return [dateStr, calc.tank, calc.ullage, calc.t1, calc.t2, calc.t3].join("|");
}

let lastSavedSignature = null;

function saveCurrentHistory() {
  const calc = calculate();
  if (!calc || !Number.isFinite(calc.ullage) || calc.ullage <= 0) {
    alert("Lengkapi Dipp/Ullage dan data sounding terlebih dahulu.");
    return;
  }
  if (!Number.isFinite(calc.vmt)) {
    alert("Perhitungan belum lengkap. Pastikan data temperature, density, dan calibration Excel sudah terbaca.");
    return;
  }

  const now = new Date();
  const dateStr = localDateString(now);
  const signature = historyInputSignature(calc, dateStr);
  const history = getHistory();
  const existing = history.find(r => r.signature === signature);

  if (existing) {
    // Jangan membuat baris baru jika kalkulasi yang sama sudah pernah disimpan.
    Object.assign(existing, calc, { date: dateStr, signature });
    existing.time = existing.time || localTimeString(now);
    setHistory(history);
    lastSavedSignature = signature;
    renderHistory();
    showSaveStatus("✓ Data ini sudah ada di history.", true);
    return;
  }

  history.push({
    id: `${now.getTime()}-${Math.random().toString(36).slice(2,8)}`,
    date: dateStr,
    time: localTimeString(now),
    signature,
    ...calc
  });

  history.sort((a,b)=>`${b.date} ${b.time}`.localeCompare(`${a.date} ${a.time}`));
  setHistory(history);
  lastSavedSignature = signature;
  renderHistory();
  showSaveStatus("✓ Kalkulasi berhasil disimpan ke History.", true);
}

function showSaveStatus(message, ok = false) {
  const el = $("saveHistoryStatus");
  if (!el) return;
  el.textContent = message;
  el.classList.toggle("success", ok);
  clearTimeout(showSaveStatus.timer);
  showSaveStatus.timer = setTimeout(() => { el.textContent = ""; el.classList.remove("success"); }, 3500);
}

function renderHistory() {
  const all = getHistory();
  const filtered = historyDateFilter ? all.filter(r => r.date === historyDateFilter) : all;
  const body = $("historyBody");
  const empty = $("historyEmpty");

  if (!filtered.length) {
    body.innerHTML = "";
    empty.classList.remove("hidden");
    $("historySummary").textContent = historyDateFilter ? `Tidak ada data pada ${formatDateDisplay(historyDateFilter)}.` : "Belum ada history sounding.";
    return;
  }
  empty.classList.add("hidden");

  const dates = [...new Set(filtered.map(r=>r.date))];
  $("historySummary").textContent = `${filtered.length} record${filtered.length > 1 ? "s" : ""} • ${historyDateFilter ? formatDateDisplay(historyDateFilter) : `${dates.length} hari`}`;

  body.innerHTML = filtered.map(r => `
    <tr>
      <td>${formatDateDisplay(r.date)}</td>
      <td>${r.time}</td>
      <td><b>ST ${r.tank}</b></td>
      <td>${fmt(r.ullage,1)}</td>
      <td>${Number.isFinite(r.avgTemp) ? fmt(r.avgTemp,2) : "-"}</td>
      <td>${Number.isFinite(r.tableTemp) ? fmt(r.tableTemp,1) : "-"}</td>
      <td>${Number.isFinite(r.density) ? fmt(r.density,4) : "-"}</td>
      <td>${Number.isFinite(r.oilVolume) ? fmt(r.oilVolume) : "-"}</td>
      <td><b>${Number.isFinite(r.vmt) ? fmt(r.vmt) : "-"}</b></td>
      <td><button class="history-delete" data-id="${r.id}">Hapus</button></td>
    </tr>
  `).join("");

  body.querySelectorAll(".history-delete").forEach(btn => btn.addEventListener("click", () => deleteHistory(btn.dataset.id)));
}

function deleteHistory(id) {
  const next = getHistory().filter(r => r.id !== id);
  setHistory(next);
  renderHistory();
}

function clearAllHistory() {
  const count = getHistory().length;
  if (!count) return;
  if (confirm(`Hapus semua ${count} record history?`)) {
    localStorage.removeItem(HISTORY_KEY);
    renderHistory();
  }
}

function exportHistoryCSV() {
  const all = getHistory();
  const data = historyDateFilter ? all.filter(r=>r.date===historyDateFilter) : all;
  if (!data.length) { alert("Tidak ada history untuk diexport."); return; }
  const headers = ["Tanggal","Jam","Tank","Dipp/Ullage (cm)","Temperature 1 (C)","Temperature 2 (C)","Temperature 3 (C)","Rata-rata Temperature (C)","Temperature Tabel (C)","Density","Ullage Dasar (cm)","Selisih (cm)","Volume Dasar (L)","Volume per cm (L/cm)","Koreksi Volume (L)","Volume Oil (L)","Volume x Density (MT)","Correction Factor","VMT (MT)"];
  const values = data.map(r=>[r.date,r.time,`ST ${r.tank}`,r.ullage,r.t1,r.t2,r.t3,r.avgTemp,r.tableTemp,r.density,r.baseUllage,r.difference,r.baseVolume,r.litrePerCm,r.volumeCorrection,r.oilVolume,r.massBeforeFactor,r.factor,r.vmt]);
  const csv = [headers,...values].map(row=>row.map(v=>`"${String(v ?? "").replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob(["\uFEFF"+csv],{type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `History_Sounding_${historyDateFilter || "Semua"}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function showView(view) {
  $("tankView").classList.toggle("hidden", view === "history");
  $("historyView").classList.toggle("hidden", view !== "history");
  document.querySelectorAll(".top-tab").forEach(btn => btn.classList.toggle("active", btn.dataset.view === view));
  if (view === "history") renderHistory();
}

["ullage","t1","t2","t3"].forEach(id => {
  $(id).addEventListener("input",()=>{
    saveActiveTankState();
    calculate();
    // Input hanya menghitung hasil. History TIDAK disimpan otomatis.
  });
});

$("saveHistoryBtn").addEventListener("click", saveCurrentHistory);

$("historyDate").value = localDateString();
$("historyDate").addEventListener("change", e => { historyDateFilter = e.target.value || null; renderHistory(); });
$("showAllHistoryBtn").addEventListener("click", () => { historyDateFilter = null; $("historyDate").value = ""; renderHistory(); });
$("exportHistoryBtn").addEventListener("click", exportHistoryCSV);
$("clearHistoryBtn").addEventListener("click", clearAllHistory);

document.querySelectorAll(".top-tab").forEach(btn => btn.addEventListener("click", () => showView(btn.dataset.view === "raw" ? "tank" : btn.dataset.view)));

buildTankTabs();
writeInputs(tankState[1]);
renderHistory();
loadExcel(1);
loadDensity();
