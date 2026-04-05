/**
 * TIP 指數儀表板：CSV 解析、淨值曲線、風險指標、多期累積報酬直條圖
 * 市場基準（Beta）：發行量加權股價指數，代碼 t00
 */
(function () {
  "use strict";

  const MARKET_CODE = "t00";
  const TRADING_DAYS_PER_YEAR = 252;
  const LINE_COLORS = ["#6ec8ff", "#c9a227", "#b01030", "#3fb950", "#8E44AD"];
  const PLOTLY_DARK = {
    paper_bgcolor: "#1b1e23",
    plot_bgcolor: "#1b1e23",
    font: { color: "#e6edf3", size: 12 },
    xaxis: {
      gridcolor: "#30363d",
      linecolor: "#30363d",
      zerolinecolor: "#30363d",
      tickfont: { color: "#8b949e" },
    },
    yaxis: {
      gridcolor: "#30363d",
      linecolor: "#30363d",
      zerolinecolor: "#30363d",
      tickfont: { color: "#8b949e" },
    },
    legend: {
      orientation: "h",
      yanchor: "top",
      y: -0.18,
      x: 0.5,
      xanchor: "center",
      font: { color: "#e6edf3" },
    },
    margin: { t: 24, r: 24, b: 100, l: 56 },
    hovermode: "x unified",
  };

  /** @type {Map<string, { name: string, rows: { t: number, p: number }[] }>} */
  let byCode = new Map();
  let currentHorizon = 1;

  const el = (id) => document.getElementById(id);

  function setStatus(msg, isErr) {
    const s = el("statusLine");
    s.textContent = msg || "";
    s.classList.toggle("err", !!isErr);
  }

  function parseYmd(s) {
    const m = String(s).trim().replace(/-/g, "/").match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
    if (!m) return null;
    const y = +m[1];
    const mo = +m[2] - 1;
    const d = +m[3];
    const dt = new Date(y, mo, d);
    dt.setHours(12, 0, 0, 0);
    return dt.getTime();
  }

  function ymdFromTs(ts) {
    const d = new Date(ts);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function parsePrice(v) {
    const n = parseFloat(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : NaN;
  }

  function ingestRows(rows) {
    byCode = new Map();
    for (const r of rows) {
      const code = (r["指數代碼"] || r["\u6307\u6578\u4ee3\u78bc"] || "").trim();
      if (!code) continue;
      const name = (r["指數名稱"] || r["\u6307\u6578\u540d\u7a31"] || "").trim();
      const t = parseYmd(r["日期"] || r["\u65e5\u671f"]);
      const p = parsePrice(r["價格指數值"] || r["\u50f9\u683c\u6307\u6578\u503c"]);
      if (t == null || !Number.isFinite(p) || p <= 0) continue;
      if (!byCode.has(code)) byCode.set(code, { name: name || code, rows: [] });
      const entry = byCode.get(code);
      if (name) entry.name = name;
      entry.rows.push({ t, p });
    }
    for (const [, v] of byCode) {
      v.rows.sort((a, b) => a.t - b.t);
      const dedup = [];
      let lastT = null;
      for (const row of v.rows) {
        if (row.t === lastT) {
          dedup[dedup.length - 1] = row;
        } else {
          dedup.push(row);
          lastT = row.t;
        }
      }
      v.rows = dedup;
    }
  }

  function loadCsvText(text) {
    const parsed = Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      transformHeader: (h) => String(h || "").replace(/^\uFEFF/, "").trim(),
    });
    if (parsed.errors && parsed.errors.length) {
      console.warn(parsed.errors);
    }
    ingestRows(parsed.data);
    if (byCode.size === 0) {
      setStatus("無法從 CSV 解析出有效列（請確認欄位：指數代碼、指數名稱、日期、價格指數值）。", true);
      return false;
    }
    setStatus(`已載入 ${byCode.size} 支指數序列。`);
    fillIndexSelect();
    initDateDefaults();
    return true;
  }

  function fillIndexSelect() {
    const sel = el("indexSelect");
    const filter = (el("indexFilter").value || "").trim().toLowerCase();
    sel.innerHTML = "";
    const codes = [...byCode.keys()].sort((a, b) => a.localeCompare(b));
    for (const c of codes) {
      const name = byCode.get(c).name;
      const line = `${c} — ${name}`.toLowerCase();
      if (filter && !line.includes(filter) && !c.toLowerCase().includes(filter)) continue;
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = `${c} — ${name}`;
      sel.appendChild(opt);
    }
  }

  function globalMinMaxDate() {
    let minT = Infinity;
    let maxT = -Infinity;
    for (const [, v] of byCode) {
      if (!v.rows.length) continue;
      minT = Math.min(minT, v.rows[0].t);
      maxT = Math.max(maxT, v.rows[v.rows.length - 1].t);
    }
    return { minT, maxT };
  }

  function initDateDefaults() {
    const { minT, maxT } = globalMinMaxDate();
    if (!Number.isFinite(maxT)) return;
    el("endDate").value = ymdFromTs(maxT);
    const base = new Date(maxT);
    base.setDate(base.getDate() - 90);
    const bTs = base.getTime();
    el("baseDate").value = ymdFromTs(Math.max(minT, bTs));
  }

  function sliceRange(rows, t0, t1) {
    return rows.filter((r) => r.t >= t0 && r.t <= t1);
  }

  function firstIndexOnOrAfter(rows, t0) {
    let i = 0;
    while (i < rows.length && rows[i].t < t0) i++;
    return i;
  }

  function alignedDailyReturns(assetRows, mktRows, t0, t1) {
    const mMap = new Map(mktRows.map((r) => [r.t, r.p]));
    const aIn = sliceRange(assetRows, t0, t1);
    const ra = [];
    const rm = [];
    for (let i = 1; i < aIn.length; i++) {
      const prev = aIn[i - 1];
      const cur = aIn[i];
      const pm0 = mMap.get(prev.t);
      const pm1 = mMap.get(cur.t);
      if (pm0 == null || pm1 == null || pm0 <= 0 || pm1 <= 0) continue;
      if (prev.p <= 0 || cur.p <= 0) continue;
      ra.push(cur.p / prev.p - 1);
      rm.push(pm1 / pm0 - 1);
    }
    return { ra, rm };
  }

  function mean(arr) {
    if (!arr.length) return 0;
    return arr.reduce((s, x) => s + x, 0) / arr.length;
  }

  function variance(arr, mu) {
    if (arr.length < 2) return 0;
    const m = mu !== undefined ? mu : mean(arr);
    return arr.reduce((s, x) => s + (x - m) ** 2, 0) / (arr.length - 1);
  }

  function covariance(a, b) {
    const n = Math.min(a.length, b.length);
    if (n < 2) return 0;
    const ma = mean(a.slice(0, n));
    const mb = mean(b.slice(0, n));
    let s = 0;
    for (let i = 0; i < n; i++) s += (a[i] - ma) * (b[i] - mb);
    return s / (n - 1);
  }

  function betaFromReturns(ra, rm) {
    const v = variance(rm);
    if (v < 1e-16) return null;
    return covariance(ra, rm) / v;
  }

  function sortinoRatio(ra) {
    const negSq = ra.map((r) => (r < 0 ? r * r : 0));
    const dd = Math.sqrt(mean(negSq));
    if (dd < 1e-12) return null;
    return (mean(ra) / dd) * Math.sqrt(TRADING_DAYS_PER_YEAR);
  }

  function historicalVaR95(ra) {
    if (ra.length < 5) return null;
    const s = [...ra].sort((x, y) => x - y);
    const idx = Math.floor(0.05 * (s.length - 1));
    return -s[idx];
  }

  function averageDrawdownFromReturns(ra) {
    let w = 1;
    let peak = 1;
    const dds = [];
    for (const r of ra) {
      w *= 1 + r;
      peak = Math.max(peak, w);
      dds.push((peak - w) / peak);
    }
    return dds.length ? mean(dds) : null;
  }

  function downsideVolAnnualized(ra) {
    const negSq = ra.map((r) => (r < 0 ? r * r : 0));
    return Math.sqrt(mean(negSq)) * Math.sqrt(TRADING_DAYS_PER_YEAR);
  }

  function fmtPct(x, digits) {
    if (x == null || Number.isNaN(x)) return "—";
    return `${(100 * x).toFixed(digits)}%`;
  }

  function fmtNum(x, digits) {
    if (x == null || Number.isNaN(x)) return "—";
    return x.toFixed(digits);
  }

  function getSelectedCodes() {
    const sel = el("indexSelect");
    return [...sel.selectedOptions].map((o) => o.value).filter(Boolean);
  }

  function buildLineTraces(codes, t0, t1) {
    const traces = [];
    codes.forEach((code, idx) => {
      const series = byCode.get(code);
      if (!series) return;
      const rows = series.rows;
      const i0 = firstIndexOnOrAfter(rows, t0);
      if (i0 >= rows.length || rows[i0].t > t1) return;
      const sub = [];
      for (let i = i0; i < rows.length && rows[i].t <= t1; i++) sub.push(rows[i]);
      if (sub.length < 2) return;
      const p0 = sub[0].p;
      const xs = sub.map((r) => ymdFromTs(r.t));
      const ys = sub.map((r) => (100 * r.p) / p0);
      const color = LINE_COLORS[idx % LINE_COLORS.length];
      traces.push({
        type: "scatter",
        mode: "lines",
        name: `${code} ${series.name}`,
        x: xs,
        y: ys,
        line: { color, width: 2 },
        hovertemplate: "%{x}<br>%{y:.2f}<extra></extra>",
      });
    });
    return traces;
  }

  function renderLineChart(traces) {
    const layout = {
      ...PLOTLY_DARK,
      title: { text: "淨值走勢（基準日 = 100）", font: { color: "#e6edf3", size: 14 } },
      xaxis: { ...PLOTLY_DARK.xaxis, title: "日期" },
      yaxis: { ...PLOTLY_DARK.yaxis, title: "淨值" },
    };
    Plotly.react("lineChart", traces, layout, { responsive: true, displaylogo: false });
  }

  function fillStatsTable(codes, t0, t1) {
    const tbody = el("statsTable").querySelector("tbody");
    tbody.innerHTML = "";
    const mkt = byCode.get(MARKET_CODE);
    if (!mkt) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="7">資料中無市場指數 ${MARKET_CODE}（發行量加權股價指數），無法計算 Beta。</td>`;
      tbody.appendChild(tr);
    }
    for (const code of codes) {
      const series = byCode.get(code);
      const tr = document.createElement("tr");
      if (!series) {
        tr.innerHTML = `<td>${code}</td><td colspan="6">無資料</td>`;
        tbody.appendChild(tr);
        continue;
      }
      let betaS = "—";
      let sortinoS = "—";
      let varS = "—";
      let addS = "—";
      let ddsS = "—";
      if (mkt) {
        const { ra, rm } = alignedDailyReturns(series.rows, mkt.rows, t0, t1);
        if (ra.length >= 2 && rm.length >= 2) {
          const b = betaFromReturns(ra, rm);
          betaS = b == null ? "—" : fmtNum(b, 3);
          const so = sortinoRatio(ra);
          sortinoS = so == null ? "—" : fmtNum(so, 3);
          const v95 = historicalVaR95(ra);
          varS = v95 == null ? "—" : fmtPct(v95, 2);
          const add = averageDrawdownFromReturns(ra);
          addS = add == null ? "—" : fmtPct(add, 2);
          const ddv = downsideVolAnnualized(ra);
          ddsS = fmtPct(ddv, 2);
        }
      }
      tr.innerHTML = `
        <td>${code}</td>
        <td>${series.name}</td>
        <td>${betaS}</td>
        <td>${sortinoS}</td>
        <td>${varS}</td>
        <td>${addS}</td>
        <td>${ddsS}</td>`;
      tbody.appendChild(tr);
    }
  }

  /** 截止日 endTs（含）前，最近 horizon 個「交易日」的累積報酬（horizon=1 為當日） */
  function cumulativeOverHorizon(rows, endTs, horizon) {
    const upto = rows.filter((r) => r.t <= endTs);
    if (upto.length < horizon + 1) return null;
    const last = upto[upto.length - 1];
    const start = upto[upto.length - 1 - horizon];
    return last.p / start.p - 1;
  }

  function commonEndDate(codes) {
    let minLast = Infinity;
    for (const code of codes) {
      const rows = byCode.get(code)?.rows;
      if (!rows || !rows.length) return null;
      const last = rows[rows.length - 1].t;
      minLast = Math.min(minLast, last);
    }
    return minLast;
  }

  function renderBarChart(codes, endTs, horizon) {
    const items = [];
    for (const code of codes) {
      const series = byCode.get(code);
      if (!series) continue;
      const cr = cumulativeOverHorizon(series.rows, endTs, horizon);
      if (cr == null) continue;
      items.push({ code, name: series.name, v: cr });
    }
    if (!items.length) {
      Plotly.purge("barChart");
      el("barTable").querySelector("tbody").innerHTML = "";
      return;
    }
    items.sort((a, b) => b.v - a.v);
    const xTick = items.map((x) => x.code);
    const vals = items.map((x) => x.v);
    const yPct = vals.map((v) => 100 * v);
    const maxY = Math.max(...yPct);
    const minY = Math.min(...yPct);
    /** 預留柱頂／柱底文字（textposition outside）空間，避免被裁切 */
    const topPad = Math.max(2.8, Math.abs(maxY) * 0.32 + 2);
    const botPad = Math.max(2.8, Math.abs(minY) * 0.32 + 2);
    const yAxisHi = maxY + topPad;
    const yAxisLo = minY < 0 ? minY - botPad : 0;

    const colors = vals.map((v) => (v >= 0 ? "#e85d5d" : "#3dcda8"));
    const trace = {
      type: "bar",
      x: xTick,
      y: yPct,
      marker: { color: colors, line: { width: 0 } },
      text: vals.map((v) => fmtPct(v, 2)),
      textposition: "outside",
      textfont: { color: "#e6edf3", size: 14 },
      cliponaxis: false,
      hovertemplate: "%{x}<br>%{y:.2f}%<extra></extra>",
    };
    const layout = {
      ...PLOTLY_DARK,
      font: { color: "#e6edf3", size: 14 },
      title: {
        text: `累積報酬（${horizon} 個交易日，截止 ${ymdFromTs(endTs)}）`,
        font: { color: "#e6edf3", size: 18 },
      },
      xaxis: {
        ...PLOTLY_DARK.xaxis,
        tickangle: -90,
        automargin: true,
        title: { text: "指數", font: { color: "#8b949e", size: 15 } },
        tickfont: { size: 13, color: "#8b949e" },
      },
      yaxis: {
        ...PLOTLY_DARK.yaxis,
        title: { text: "累積報酬率（%）", font: { color: "#8b949e", size: 15 } },
        tickfont: { size: 13, color: "#8b949e" },
        ticksuffix: "%",
        range: [yAxisLo, yAxisHi],
        automargin: true,
      },
      margin: { t: 88, r: 48, b: 200, l: 80 },
      bargap: 0.12,
    };
    Plotly.react("barChart", [trace], layout, { responsive: true, displaylogo: false });

    const tbody = el("barTable").querySelector("tbody");
    tbody.innerHTML = "";
    for (const it of items) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${it.code}</td><td>${it.name}</td><td>${fmtPct(it.v, 3)}</td>`;
      tbody.appendChild(tr);
    }
  }

  function applyCharts() {
    const codes = getSelectedCodes();
    if (!codes.length) {
      setStatus("請至少選擇一檔指數。", true);
      return;
    }
    if (!byCode.size) {
      setStatus("請先載入 CSV。", true);
      return;
    }
    const baseStr = el("baseDate").value;
    const endStr = el("endDate").value;
    if (!baseStr) {
      setStatus("請設定基準日。", true);
      return;
    }
    const t0 = new Date(baseStr + "T12:00:00").getTime();
    let t1 = endStr ? new Date(endStr + "T12:00:00").getTime() : globalMinMaxDate().maxT;
    if (t1 < t0) {
      setStatus("結束日不可早於基準日。", true);
      return;
    }

    const traces = buildLineTraces(codes, t0, t1);
    if (!traces.length) {
      setStatus("選定區間內沒有足夠價格資料可畫淨值線。", true);
      return;
    }
    renderLineChart(traces);
    fillStatsTable(codes, t0, t1);

    const endTs = commonEndDate(codes);
    if (endTs == null) {
      el("barAsOf").textContent = "";
      return;
    }
    el("barAsOf").textContent = `比較截止日（所選指數共同最後交易日）：${ymdFromTs(endTs)}`;
    renderBarChart(codes, endTs, currentHorizon);
    setStatus(`已更新：${codes.length} 檔指數。`);
  }

  function onHorizonTab(ev) {
    const btn = ev.target.closest("button[data-horizon]");
    if (!btn) return;
    document.querySelectorAll("#horizonTabs .tab").forEach((b) => b.classList.remove("active"));
    btn.classList.add("active");
    currentHorizon = +btn.getAttribute("data-horizon");
    const codes = getSelectedCodes();
    const endTs = commonEndDate(codes);
    if (endTs != null && codes.length) renderBarChart(codes, endTs, currentHorizon);
  }

  async function fetchDefaultCsv() {
    setStatus("載入中…");
    try {
      const res = await fetch("../output/all_history.csv", { cache: "no-store" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const text = await res.text();
      loadCsvText(text);
    } catch (e) {
      setStatus(`無法載入預設檔案（請確認已用專案根目錄啟動 http.server）：${e.message}`, true);
    }
  }

  el("fileCsv").addEventListener("change", (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      loadCsvText(String(reader.result || ""));
    };
    reader.readAsText(f, "UTF-8");
  });

  el("btnLoadDefault").addEventListener("click", fetchDefaultCsv);
  el("btnApply").addEventListener("click", applyCharts);
  el("indexFilter").addEventListener("input", () => fillIndexSelect());
  el("horizonTabs").addEventListener("click", onHorizonTab);

  window.addEventListener("resize", () => {
    try {
      Plotly.Plots.resize("lineChart");
      Plotly.Plots.resize("barChart");
    } catch (_) {
      /* 圖尚未建立時略過 */
    }
  });
})();
