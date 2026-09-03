/**
 * Gom CSV thời tiết thành JSON gọn cho dashboard.
 * Node 22+, không cần thư viện ngoài.
 */
const fs = require("fs");
const path = require("path");

const SRC = path.join(__dirname, "..", "df_weather_fixed_utf8.csv");
const OUT = path.join(__dirname, "..", "dashboard", "data", "weather.json");

const REGION_FIX = {
  "Tr [*]ung du và miền núi Bắc Bộ": "Trung du và miền núi Bắc Bộ",
};

function parseCSV(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.length);
  const header = lines[0].split(",");
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split(",");
    if (cols.length < header.length) continue;
    const row = {};
    for (let j = 0; j < header.length; j++) row[header[j]] = cols[j];
    rows.push(row);
  }
  return rows;
}

function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function round(v, d = 2) {
  if (v == null) return null;
  const p = 10 ** d;
  return Math.round(v * p) / p;
}

const raw = fs.readFileSync(SRC, "utf8").replace(/^\uFEFF/, "");
const rows = parseCSV(raw);

const provinceMap = new Map();
const dateSet = new Set();
const regionSet = new Set();
const terrainSet = new Set();
const conditionSet = new Set();

for (const r of rows) {
  const name = r["location.name"];
  if (!provinceMap.has(name)) {
    const region = REGION_FIX[r["location.region"]] || r["location.region"];
    provinceMap.set(name, {
      name,
      region,
      terrain: r["location.terrain"],
      lat: num(r["location.lat"]),
      lon: num(r["location.lon"]),
    });
    regionSet.add(region);
    terrainSet.add(r["location.terrain"]);
  }
  dateSet.add(r["date"]);
  conditionSet.add(r["day.condition.text"]);
}

const provinces = [...provinceMap.values()].sort((a, b) =>
  a.name.localeCompare(b.name, "vi")
);
const dates = [...dateSet].sort();
const regions = [...regionSet].sort((a, b) => a.localeCompare(b, "vi"));
const terrains = [...terrainSet].sort((a, b) => a.localeCompare(b, "vi"));
const conditions = [...conditionSet].sort();

const pIndex = Object.fromEntries(provinces.map((p, i) => [p.name, i]));
const dIndex = Object.fromEntries(dates.map((d, i) => [d, i]));
const cIndex = Object.fromEntries(conditions.map((c, i) => [c, i]));

const N = provinces.length * dates.length;
const tmax = new Array(N).fill(null);
const tmin = new Array(N).fill(null);
const tavg = new Array(N).fill(null);
const precip = new Array(N).fill(null);
const humidity = new Array(N).fill(null);
const wind = new Array(N).fill(null);
const uv = new Array(N).fill(null);
const cond = new Array(N).fill(null);
const rainChance = new Array(N).fill(null);

let missing = 0;
for (const r of rows) {
  const pi = pIndex[r["location.name"]];
  const di = dIndex[r["date"]];
  if (pi == null || di == null) continue;
  const k = pi * dates.length + di;
  tmax[k] = round(num(r["day.maxtemp_c"]), 1);
  tmin[k] = round(num(r["day.mintemp_c"]), 1);
  tavg[k] = round(num(r["day.avgtemp_c"]), 1);
  precip[k] = round(num(r["day.totalprecip_mm"]), 2);
  humidity[k] = round(num(r["day.avghumidity"]), 0);
  wind[k] = round(num(r["day.maxwind_kph"]), 1);
  uv[k] = round(num(r["day.uv"]), 1);
  rainChance[k] = num(r["day.daily_chance_of_rain"]);
  cond[k] = cIndex[r["day.condition.text"]];
}

for (let i = 0; i < N; i++) if (tmax[i] == null) missing++;

// Monthly aggregates by region for heatmap
const monthKeys = [];
const monthIndex = {};
for (const d of dates) {
  const m = d.slice(0, 7);
  if (monthIndex[m] == null) {
    monthIndex[m] = monthKeys.length;
    monthKeys.push(m);
  }
}

const monthly = regions.map(() =>
  monthKeys.map(() => ({ t: 0, p: 0, h: 0, n: 0, rainDays: 0 }))
);
for (const r of rows) {
  const region = REGION_FIX[r["location.region"]] || r["location.region"];
  const ri = regions.indexOf(region);
  const mi = monthIndex[r["date"].slice(0, 7)];
  if (ri < 0 || mi == null) continue;
  const cell = monthly[ri][mi];
  const t = num(r["day.avgtemp_c"]);
  const p = num(r["day.totalprecip_mm"]);
  const h = num(r["day.avghumidity"]);
  if (t != null) {
    cell.t += t;
    cell.n += 1;
  }
  if (p != null) {
    cell.p += p;
    if (p > 0.1) cell.rainDays += 1;
  }
  if (h != null) cell.h += h;
}

const heatmap = {
  months: monthKeys,
  regions,
  avgTemp: monthly.map((row) => row.map((c) => (c.n ? round(c.t / c.n, 2) : null))),
  totalPrecip: monthly.map((row) => row.map((c) => round(c.p, 1))),
  avgHumidity: monthly.map((row) => row.map((c) => (c.n ? round(c.h / c.n, 1) : null))),
};

// Daily regional series (for trend chart)
const dailyRegion = regions.map(() =>
  dates.map(() => ({ t: 0, p: 0, n: 0 }))
);
const unknownRegions = new Set();
for (const r of rows) {
  const region = REGION_FIX[r["location.region"]] || r["location.region"];
  const ri = regions.indexOf(region);
  const di = dIndex[r["date"]];
  if (ri < 0 || di == null) {
    unknownRegions.add(region);
    continue;
  }
  const cell = dailyRegion[ri][di];
  cell.t += num(r["day.avgtemp_c"]) || 0;
  cell.p += num(r["day.totalprecip_mm"]) || 0;
  cell.n += 1;
}
if (unknownRegions.size) {
  console.warn("Unknown regions:", [...unknownRegions]);
}
const regionSeries = {
  dates,
  regions,
  avgTemp: dailyRegion.map((row) => row.map((c) => (c.n ? round(c.t / c.n, 2) : null))),
  avgPrecip: dailyRegion.map((row) => row.map((c) => (c.n ? round(c.p / c.n, 2) : null))),
};

const payload = {
  meta: {
    nRows: rows.length,
    nProvinces: provinces.length,
    nDates: dates.length,
    dateMin: dates[0],
    dateMax: dates[dates.length - 1],
    missingCells: missing,
    generatedAt: new Date().toISOString(),
  },
  provinces,
  dates,
  regions,
  terrains,
  conditions,
  tmax,
  tmin,
  tavg,
  precip,
  humidity,
  wind,
  uv,
  rainChance,
  cond,
  heatmap,
  regionSeries,
};

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, JSON.stringify(payload));
const mb = (fs.statSync(OUT).size / 1024 / 1024).toFixed(2);
console.log(
  JSON.stringify(
    {
      out: OUT,
      sizeMB: mb,
      provinces: provinces.length,
      dates: dates.length,
      regions,
      terrains,
      conditions: conditions.length,
      missingCells: missing,
      dateRange: [dates[0], dates[dates.length - 1]],
    },
    null,
    2
  )
);
