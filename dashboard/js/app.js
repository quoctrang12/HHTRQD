/* Atlas khí tượng VN — linked views for 3 Munzner tasks */
const REGION_COLOR = {
  "Bắc Trung Bộ và Duyên hải miền Trung": "#1d6b6b",
  "Đồng Bằng Sông Cửu Long": "#2f6fad",
  "Đồng Bằng Sông Hồng": "#6b4c9a",
  "Đông Nam Bộ": "#c9a227",
  "Tây Nguyên": "#c24a32",
  "Trung du và miền núi Bắc Bộ": "#3d5a80",
};
const TERRAIN_SYMBOL = { "đồng bằng": "circle", "ven biển": "diamond", "miền núi": "square" };
const ZOOM_BTNS = [
  "zoom2d",
  "zoomIn2d",
  "zoomOut2d",
  "zoomInGeo",
  "zoomOutGeo",
  "panGeo",
  "autoScale2d",
  "hoverClosestGeo",
  "toggleSpikelines",
];
const PLOT_CFG = {
  responsive: true,
  displaylogo: false,
  scrollZoom: false,
  doubleClick: "reset",
  displayModeBar: true,
  modeBarButtonsToRemove: [...ZOOM_BTNS, "lasso2d"],
};
const MAP_CFG = {
  ...PLOT_CFG,
  doubleClick: false,
  modeBarButtonsToRemove: [...ZOOM_BTNS, "lasso2d", "select2d", "pan2d", "resetGeo", "resetScale2d"],
};
const SCATTER_CFG = {
  ...PLOT_CFG,
  modeBarButtonsToRemove: ZOOM_BTNS,
};
const FONT = { family: "Segoe UI, Helvetica Neue, Arial, sans-serif", color: "#1b2a4a" };

let DATA = null;
const state = {
  from: null,
  to: null,
  regions: new Set(),
  terrains: new Set(),
  metric: "tavg",
  lockPi: null,
  brushPi: null,
};

const $ = (id) => document.getElementById(id);
const key = (pi, di) => pi * DATA.dates.length + di;
const metricArr = () => DATA[state.metric];
const metricLabel = () =>
  ({
    tavg: "Nhiệt độ TB (°C)",
    tmax: "Nhiệt độ max (°C)",
    precip: "Lượng mưa (mm)",
    humidity: "Độ ẩm (%)",
    wind: "Gió (km/h)",
  }[state.metric]);

function dateBounds() {
  const i0 = DATA.dates.indexOf(state.from);
  const i1 = DATA.dates.indexOf(state.to);
  return [Math.max(0, i0), i1 < 0 ? DATA.dates.length - 1 : i1];
}

function visiblePi() {
  return DATA.provinces
    .map((p, i) => i)
    .filter((i) => {
      const p = DATA.provinces[i];
      const ok = state.regions.has(p.region) && state.terrains.has(p.terrain);
      if (state.lockPi != null) return ok && i === state.lockPi;
      if (state.brushPi && state.brushPi.size) return ok && state.brushPi.has(i);
      return ok;
    });
}

function mean(vals) {
  let s = 0, n = 0;
  for (const v of vals) if (v != null) { s += v; n++; }
  return n ? s / n : null;
}

function movingAvg(arr, w = 7) {
  const out = new Array(arr.length).fill(null);
  let s = 0;
  for (let i = 0; i < arr.length; i++) {
    s += arr[i] || 0;
    if (i >= w) s -= arr[i - w] || 0;
    const n = Math.min(i + 1, w);
    out[i] = +(s / n).toFixed(2);
  }
  return out;
}

function quantile(sorted, q) {
  if (!sorted.length) return null;
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  if (lo === hi) return sorted[lo];
  return sorted[lo] * (hi - pos) + sorted[hi] * (pos - lo);
}

function layoutBase(extra) {
  return Object.assign(
    {
      font: FONT,
      autosize: true,
      paper_bgcolor: "rgba(0,0,0,0)",
      plot_bgcolor: "#fffdf8",
      margin: { t: 48, r: 24, b: 52, l: 52 },
      hovermode: "closest",
      dragmode: false,
      legend: { orientation: "h", y: -0.18, x: 0, bgcolor: "rgba(255,253,248,0.9)" },
    },
    extra
  );
}

function resizePlots() {
  document.querySelectorAll(".js-plotly-plot").forEach((el) => {
    try {
      Plotly.Plots.resize(el);
    } catch (err) {
      /* plot chưa init */
    }
  });
}

function initFilters() {
  state.from = DATA.dates[0];
  state.to = DATA.dates[DATA.dates.length - 1];
  DATA.regions.forEach((r) => state.regions.add(r));
  DATA.terrains.forEach((t) => state.terrains.add(t));
  $("dateFrom").value = state.from;
  $("dateTo").value = state.to;
  $("dateFrom").min = $("dateTo").min = DATA.dates[0];
  $("dateFrom").max = $("dateTo").max = DATA.dates[DATA.dates.length - 1];

  $("regionBox").innerHTML = DATA.regions
    .map(
      (r) =>
        `<label><input type="checkbox" data-region="${r}" checked>
        <i style="display:inline-block;width:9px;height:9px;background:${REGION_COLOR[r]};border-radius:50%"></i>
        ${r}</label>`
    )
    .join("");
  $("terrainBox").innerHTML = DATA.terrains
    .map((t) => `<label><input type="checkbox" data-terrain="${t}" checked> ${t} (${TERRAIN_SYMBOL[t]})</label>`)
    .join("");
}

function bind() {
  $("dateFrom").onchange = (e) => {
    state.from = e.target.value;
    if (state.from > state.to) {
      state.to = state.from;
      $("dateTo").value = state.to;
    }
    renderAll();
  };
  $("dateTo").onchange = (e) => {
    state.to = e.target.value;
    if (state.to < state.from) {
      state.from = state.to;
      $("dateFrom").value = state.from;
    }
    renderAll();
  };
  $("metric").onchange = (e) => {
    state.metric = e.target.value;
    renderAll();
  };
  $("regionBox").onchange = (e) => {
    const r = e.target.dataset.region;
    if (!r) return;
    e.target.checked ? state.regions.add(r) : state.regions.delete(r);
    renderAll();
  };
  $("terrainBox").onchange = (e) => {
    const t = e.target.dataset.terrain;
    if (!t) return;
    e.target.checked ? state.terrains.add(t) : state.terrains.delete(t);
    renderAll();
  };
  $("btnClearLock").onclick = () => {
    state.lockPi = null;
    $("lockBox").hidden = true;
    renderAll();
  };
  $("btnReset").onclick = () => {
    state.regions = new Set(DATA.regions);
    state.terrains = new Set(DATA.terrains);
    state.lockPi = null;
    state.brushPi = null;
    state.metric = "tavg";
    $("metric").value = "tavg";
    $("lockBox").hidden = true;
    initFilters();
    renderAll();
  };
}

function provinceAgg(pi) {
  const [i0, i1] = dateBounds();
  const t = [], p = [], m = [];
  for (let di = i0; di <= i1; di++) {
    const k = key(pi, di);
    t.push(DATA.tavg[k]);
    p.push(DATA.precip[k]);
    m.push(metricArr()[k]);
  }
  return { t: mean(t), p: mean(p), m: mean(m) };
}

function renderKPIs() {
  const pis = visiblePi();
  const [i0, i1] = dateBounds();
  const temps = [], rains = [], hums = [];
  let hot = { v: -Infinity }, wet = { v: -Infinity };
  for (const pi of pis) {
    for (let di = i0; di <= i1; di++) {
      const k = key(pi, di);
      const t = DATA.tavg[k], p = DATA.precip[k], h = DATA.humidity[k];
      if (t != null) temps.push(t);
      if (p != null) rains.push(p);
      if (h != null) hums.push(h);
      if (t > hot.v) hot = { v: t, p: DATA.provinces[pi].name, d: DATA.dates[di] };
      if (p > wet.v) wet = { v: p, p: DATA.provinces[pi].name, d: DATA.dates[di] };
    }
  }
  const cards = [
    ["Tỉnh đang xem", String(pis.length), "trong bộ lọc"],
    ["Nhiệt độ TB", temps.length ? mean(temps).toFixed(1) + "°C" : "—", `${state.from} → ${state.to}`],
    ["Mưa TB/ngày", rains.length ? mean(rains).toFixed(1) + " mm" : "—", "trung bình tỉnh-ngày"],
    ["Nóng nhất", hot.v > -Infinity ? hot.v.toFixed(1) + "°C" : "—", `${hot.p || ""} ${hot.d || ""}`],
    ["Mưa lớn nhất", wet.v > -Infinity ? wet.v.toFixed(1) + " mm" : "—", `${wet.p || ""} ${wet.d || ""}`],
  ];
  $("kpis").innerHTML = cards
    .map(([k, v, s]) => `<div class="kpi"><span>${k}</span><b>${v}</b><span>${s}</span></div>`)
    .join("");
}

function renderMap() {
  const pis = DATA.provinces
    .map((p, i) => i)
    .filter((i) => state.regions.has(DATA.provinces[i].region) && state.terrains.has(DATA.provinces[i].terrain));
  const traces = DATA.terrains
    .filter((t) => state.terrains.has(t))
    .map((terrain) => {
      const subset = pis.filter((i) => DATA.provinces[i].terrain === terrain);
      const agg = subset.map(provinceAgg);
      return {
        type: "scattergeo",
        lon: subset.map((i) => DATA.provinces[i].lon),
        lat: subset.map((i) => DATA.provinces[i].lat),
        text: subset.map(
          (i, j) =>
            `<b>${DATA.provinces[i].name}</b><br>${DATA.provinces[i].region}<br>` +
            `${metricLabel()}: ${agg[j].m != null ? agg[j].m.toFixed(1) : "—"}<br>` +
            `Mưa TB: ${agg[j].p != null ? agg[j].p.toFixed(1) : "—"} mm`
        ),
        hoverinfo: "text",
        customdata: subset,
        marker: {
          symbol: TERRAIN_SYMBOL[terrain],
          size: agg.map((a) => Math.max(8, Math.min(28, 8 + (a.p || 0) * 1.6))),
          color: agg.map((a) => a.m),
          colorscale: state.metric === "precip" ? "Teal" : "RdYlBu",
          reversescale: state.metric !== "precip",
          cmin: (() => {
            const ms = agg.map((a) => a.m).filter((v) => v != null);
            return ms.length ? Math.min(...ms) : 0;
          })(),
          cmax: (() => {
            const ms = agg.map((a) => a.m).filter((v) => v != null);
            return ms.length ? Math.max(...ms) : 1;
          })(),
          colorbar: {
            title: { text: metricLabel() },
            len: 0.72,
            thickness: 12,
            x: 1.01,
            y: 0.52,
            xpad: 4,
            ypad: 0,
          },
          line: { width: 0.6, color: "#1b2a4a" },
          opacity: 0.92,
          showscale: terrain === DATA.terrains.filter((t) => state.terrains.has(t))[0],
        },
        name: terrain,
      };
    })
    .filter((tr) => tr.lon.length);

  if (!traces.length) return;

  Plotly.react(
    "mapPlot",
    traces,
    layoutBase({
      title: { text: "Bản đồ tỷ lệ — click một tỉnh để khóa view", font: { size: 13 } },
      dragmode: false,
      geo: {
        resolution: 50,
        showframe: false,
        showcoastlines: true,
        coastlinecolor: "#b7b1a4",
        showland: true,
        landcolor: "#efe8d8",
        showocean: true,
        oceancolor: "#d7e6ea",
        showcountries: true,
        countrycolor: "#c1b9a8",
        lonaxis: { range: [102.1, 110.4] },
        lataxis: { range: [8.2, 23.5] },
        center: { lon: 106.3, lat: 15.9 },
        projection: { type: "mercator" },
        bgcolor: "rgba(0,0,0,0)",
        domain: { x: [0, 0.9], y: [0.02, 1] },
      },
      margin: { t: 44, r: 52, b: 36, l: 4 },
      legend: { orientation: "h", y: -0.08, x: 0, font: { size: 11 } },
      uirevision: "map",
    }),
    MAP_CFG
  );

  const node = $("mapPlot");
  if (!node.dataset.clickBound) {
    node.dataset.clickBound = "1";
    node.on("plotly_click", (ev) => {
      if (!ev.points || !ev.points.length) return;
      const pi = ev.points[0].customdata;
      state.lockPi = pi;
      state.brushPi = null;
      $("lockBox").hidden = false;
      $("lockName").textContent = "Đang khóa: " + DATA.provinces[pi].name;
      renderAll();
    });
  }
}

function renderHist() {
  const pis = visiblePi();
  const [i0, i1] = dateBounds();
  const vals = [];
  for (const pi of pis) {
    for (let di = i0; di <= i1; di++) {
      const v = metricArr()[key(pi, di)];
      if (v != null) vals.push(v);
    }
  }
  Plotly.react(
    "histPlot",
    [
      {
        type: "histogram",
        x: vals,
        nbinsx: 36,
        marker: { color: "#1d6b6b", line: { width: 0 } },
        opacity: 0.88,
        name: "Phân bố",
      },
    ],
    layoutBase({
      title: { text: `Phân bố ${metricLabel().toLowerCase()}`, font: { size: 13 } },
      xaxis: { title: metricLabel(), automargin: true },
      yaxis: { title: "Số tỉnh-ngày", automargin: true },
      bargap: 0.05,
      uirevision: "hist",
    }),
    PLOT_CFG
  );
}

function renderTrend() {
  const [i0, i1] = dateBounds();
  const dates = DATA.regionSeries.dates.slice(i0, i1 + 1);
  const traces = DATA.regionSeries.regions
    .filter((r) => state.regions.has(r))
    .map((r, idx) => {
      const ri = DATA.regionSeries.regions.indexOf(r);
      const src = state.metric === "precip" ? DATA.regionSeries.avgPrecip : DATA.regionSeries.avgTemp;
      const y = movingAvg(src[ri].slice(i0, i1 + 1));
      return {
        type: "scatter",
        mode: "lines",
        x: dates,
        y,
        name: r.replace("Bắc Trung Bộ và Duyên hải miền Trung", "DH miền Trung"),
        line: { color: REGION_COLOR[r], width: 2.2 },
        hovertemplate: "%{x}<br>%{y:.1f}<extra>" + r + "</extra>",
      };
    });

  if (state.lockPi != null) {
    const y = [];
    for (let di = i0; di <= i1; di++) {
      const arr = state.metric === "precip" ? DATA.precip : DATA.tavg;
      y.push(arr[key(state.lockPi, di)]);
    }
    traces.push({
      type: "scatter",
      mode: "lines",
      x: dates,
      y: movingAvg(y),
      name: DATA.provinces[state.lockPi].name,
      line: { color: "#1b2a4a", width: 2.6, dash: "dot" },
    });
  }

  Plotly.react(
    "trendPlot",
    traces,
    layoutBase({
      title: {
        text: state.metric === "precip" ? "Xu hướng mưa (TB 7 ngày)" : "Xu hướng nhiệt độ TB (làm mượt 7 ngày)",
        font: { size: 13 },
      },
      xaxis: { title: "Ngày" },
      yaxis: { title: state.metric === "precip" ? "mm/ngày" : "°C" },
      legend: { orientation: "h", y: -0.22, font: { size: 11 } },
      margin: { t: 44, r: 16, b: 84, l: 48 },
      uirevision: "trend",
    }),
    PLOT_CFG
  );
}

function renderHeat() {
  const [i0, i1] = dateBounds();
  const m0 = DATA.dates[i0].slice(0, 7);
  const m1 = DATA.dates[i1].slice(0, 7);
  const months = DATA.heatmap.months.filter((m) => m >= m0 && m <= m1);
  const regions = DATA.heatmap.regions.filter((r) => state.regions.has(r));
  const zSrc =
    state.metric === "precip"
      ? DATA.heatmap.totalPrecip
      : state.metric === "humidity"
        ? DATA.heatmap.avgHumidity
        : DATA.heatmap.avgTemp;
  const z = regions.map((r) => {
    const ri = DATA.heatmap.regions.indexOf(r);
    return months.map((m) => zSrc[ri][DATA.heatmap.months.indexOf(m)]);
  });
  Plotly.react(
    "heatPlot",
    [
      {
        type: "heatmap",
        x: months,
        y: regions.map((r) => r.replace("Bắc Trung Bộ và Duyên hải miền Trung", "DH miền Trung")),
        z,
        colorscale: state.metric === "precip" ? "Teal" : "RdYlBu",
        reversescale: state.metric !== "precip",
        hovertemplate: "%{y}<br>%{x}<br>%{z:.1f}<extra></extra>",
        colorbar: { title: { text: state.metric === "precip" ? "mm" : "°C" }, len: 0.72, thickness: 12, x: 1.02 },
      },
    ],
    layoutBase({
      title: { text: state.metric === "precip" ? "Tổng mưa theo tháng" : "Nhiệt độ TB theo tháng", font: { size: 13 } },
      xaxis: { title: "Tháng", tickangle: -45, automargin: true },
      yaxis: { automargin: true },
      margin: { t: 44, r: 56, b: 72, l: 110 },
      uirevision: "heat",
    }),
    PLOT_CFG
  );
}

function collectPoints() {
  const pis = visiblePi();
  const [i0, i1] = dateBounds();
  const pts = [];
  for (const pi of pis) {
    for (let di = i0; di <= i1; di++) {
      const k = key(pi, di);
      const t = DATA.tavg[k], p = DATA.precip[k], h = DATA.humidity[k];
      if (t == null || p == null) continue;
      pts.push({ pi, di, t, p, h, name: DATA.provinces[pi].name, region: DATA.provinces[pi].region });
    }
  }
  return pts;
}

function tagOutliers(pts) {
  const rains = pts.map((d) => d.p).sort((a, b) => a - b);
  const temps = pts.map((d) => d.t);
  const q1 = quantile(rains, 0.25);
  const q3 = quantile(rains, 0.75);
  const fence = q3 + 1.5 * (q3 - q1);
  const mu = mean(temps);
  const sd = Math.sqrt(mean(temps.map((t) => (t - mu) ** 2)) || 0) || 1;
  for (const d of pts) {
    const reasons = [];
    if (d.p > fence) reasons.push("mưa IQR");
    const z = (d.t - mu) / sd;
    if (Math.abs(z) >= 2.5) reasons.push(z > 0 ? "nóng z≥2.5" : "lạnh z≤-2.5");
    d.outlier = reasons.length > 0;
    d.reason = reasons.join(" + ") || "";
    d.score = (d.p > fence ? (d.p - fence) / (fence || 1) : 0) + Math.max(0, Math.abs(z) - 2.5);
  }
  return { pts, fence };
}

function renderScatter() {
  const { pts } = tagOutliers(collectPoints());
  const normal = pts.filter((d) => !d.outlier);
  const out = pts.filter((d) => d.outlier);
  const mk = (subset, color, name, size) => ({
    type: "scattergl",
    mode: "markers",
    x: subset.map((d) => d.t),
    y: subset.map((d) => d.p),
    name,
    text: subset.map(
      (d) =>
        `${d.name}<br>${DATA.dates[d.di]}<br>T° ${d.t.toFixed(1)} · mưa ${d.p.toFixed(1)} mm` +
        (d.reason ? `<br>${d.reason}` : "")
    ),
    hoverinfo: "text",
    customdata: subset.map((d) => d.pi),
    marker: { color, size, opacity: 0.55, line: { width: 0 } },
  });

  Plotly.react(
    "scatterPlot",
    [mk(normal, "#7a8499", "Bình thường", 5), mk(out, "#c24a32", "Outlier", 8)],
    layoutBase({
      title: { text: "Nhiệt độ TB × lượng mưa — chọn lasso để liên kết bản đồ", font: { size: 13 } },
      xaxis: { title: "Nhiệt độ TB (°C)", automargin: true },
      yaxis: { title: "Lượng mưa (mm/ngày)", automargin: true },
      dragmode: "select",
      legend: { orientation: "h", y: 1.02, x: 0, bgcolor: "rgba(255,253,248,0.92)", font: { size: 11 } },
      margin: { t: 56, r: 20, b: 48, l: 56 },
      uirevision: "scatter",
    }),
    SCATTER_CFG
  );

  const node = $("scatterPlot");
  if (!node.dataset.selectBound) {
    node.dataset.selectBound = "1";
    node.on("plotly_selected", (ev) => {
      if (!ev || !ev.points.length) {
        state.brushPi = null;
        renderMap();
        return;
      }
      state.brushPi = new Set(ev.points.map((p) => p.customdata));
      renderMap();
    });
    node.on("plotly_deselect", () => {
      state.brushPi = null;
      renderMap();
    });
  }
}

function renderBox() {
  const pts = collectPoints();
  const traces = DATA.terrains
    .filter((t) => state.terrains.has(t))
    .map((t) => {
      const subset = pts.filter((d) => DATA.provinces[d.pi].terrain === t);
      return {
        type: "box",
        y: subset.map((d) => d.p),
        name: t,
        boxpoints: "outliers",
        marker: { color: "#c24a32", size: 4 },
        line: { color: "#1b2a4a" },
        fillcolor: "#efe8d8",
      };
    });
  Plotly.react(
    "boxPlot",
    traces,
    layoutBase({
      title: { text: "Boxplot mưa theo địa hình — fences IQR hiện outlier", font: { size: 13 } },
      yaxis: { title: "mm/ngày", automargin: true },
      showlegend: false,
      uirevision: "box",
    }),
    PLOT_CFG
  );
}

function renderTable() {
  const { pts } = tagOutliers(collectPoints());
  const top = pts
    .filter((d) => d.outlier)
    .sort((a, b) => b.score - a.score)
    .slice(0, 15);
  const tb = $("outlierTable").querySelector("tbody");
  tb.innerHTML = top
    .map(
      (d) =>
        `<tr>
          <td>${DATA.dates[d.di]}</td>
          <td>${d.name}</td>
          <td>${d.region.replace("Bắc Trung Bộ và Duyên hải miền Trung", "DH miền Trung")}</td>
          <td>${d.t.toFixed(1)}</td>
          <td>${d.p.toFixed(1)}</td>
          <td>${d.h != null ? d.h.toFixed(0) : "—"}</td>
          <td>${d.reason}</td>
        </tr>`
    )
    .join("");
  if (!top.length) tb.innerHTML = `<tr><td colspan="7">Không có outlier trong bộ lọc hiện tại.</td></tr>`;
}

function renderAll() {
  if (!state.regions.size || !state.terrains.size) {
    $("kpis").innerHTML = `<div class="kpi"><span>Bộ lọc trống</span><b>0</b><span>Chọn lại vùng hoặc địa hình</span></div>`;
    return;
  }
  renderKPIs();
  renderMap();
  renderHist();
  renderTrend();
  renderHeat();
  renderScatter();
  renderBox();
  renderTable();
  requestAnimationFrame(resizePlots);
}

async function main() {
  const res = await fetch("data/weather.json");
  if (!res.ok) {
    document.body.innerHTML =
      "<p style='padding:40px'>Không đọc được data/weather.json. Hãy chạy <code>node serve.js</code> trong thư mục dashboard.</p>";
    return;
  }
  DATA = await res.json();
  initFilters();
  bind();
  renderAll();
  let resizeTimer = 0;
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(resizePlots, 120);
  });
}

main().catch((err) => {
  document.body.insertAdjacentHTML("afterbegin", `<p style="padding:20px;color:#c24a32">${err.message}</p>`);
});
