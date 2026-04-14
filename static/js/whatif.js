// Depends on: main.js (tableData, formatDollar), table.js (BUCKET_COLORS, escapeHtml)

// ── What If mode ─────────────────────────────────────────────────────────
let whatIfActive = false;

function toggleWhatIf() {
  // Cannot enter What If while Trends is active
  if (trendsActive) return;

  whatIfActive = !whatIfActive;
  const btn = document.getElementById("whatif-btn");
  const table = document.querySelector(".preview-table");
  btn.classList.toggle("active", whatIfActive);
  table.classList.toggle("whatif-active", whatIfActive);
  if (!whatIfActive) {
    document.getElementById("whatif-reset-all-btn").classList.add("hidden");
  } else {
    updateWhatIfResetAll();
  }
}

// nMonths scales avg→total so % of Income stays on the same dimensional basis as total income.
function _getNMonths() {
  return (tableData.total_spending > 0)
    ? tableData.total_raw_spending / tableData.total_spending : 1;
}

function _pctOfIncome(avg, income) {
  const nMonths = _getNMonths();
  return income ? ((avg * nMonths / income) * 100).toFixed(1) : "0.0";
}

function _updateBucketTotalRow(bucket, income) {
  // Sum avg amounts for all category rows in this bucket
  let bucketTotal = 0;
  document.querySelectorAll(`.cat-data-row[data-bucket="${bucket}"]`).forEach(r => {
    bucketTotal += parseFloat(r.querySelector("[data-field='amt']").dataset.raw) || 0;
  });

  // Find the TOTAL row for this bucket and update avg and % of income columns
  // Columns: [0]=label [1]=raw_total(unchanged) [2]=avg [3]=pct_income [4]=slider
  const totalLabel = `TOTAL — ${bucket}`;
  document.querySelectorAll(".preview-table tr").forEach(r => {
    const firstTd = r.querySelector("td");
    if (firstTd && firstTd.textContent.trim() === totalLabel) {
      const cells = r.querySelectorAll("td");
      cells[2].textContent = formatDollar(bucketTotal);
      cells[3].textContent = _pctOfIncome(bucketTotal, income) + "%";
    }
  });

  return bucketTotal;
}

function _updateGrandRevealRows(income) {
  // Tally avg spend per bucket across all category rows
  const allBuckets = {};
  document.querySelectorAll(".preview-table .cat-data-row").forEach(r => {
    const b = r.dataset.bucket;
    if (!allBuckets[b]) allBuckets[b] = 0;
    allBuckets[b] += parseFloat(r.querySelector("[data-field='amt']").dataset.raw) || 0;
  });
  const totalSpending = Object.values(allBuckets).reduce((a, b) => a + b, 0);

  // Grand Reveal columns: [0]=bucket [1]=raw_total(unchanged) [2]=avg [3]=%spending [4]=%income
  // TOTAL SPENDING:       [0]=label  [1]=raw_total(unchanged) [2]=avg [3]=100%      [4]=%income
  document.querySelectorAll(".preview-table tr").forEach(r => {
    const firstTd = r.querySelector("td");
    if (!firstTd) return;
    const label = firstTd.textContent.trim();

    Object.keys(allBuckets).forEach(b => {
      if (label === b.toUpperCase()) {
        const cells = r.querySelectorAll("td");
        cells[2].textContent = formatDollar(allBuckets[b]);
        cells[3].textContent = totalSpending
          ? ((allBuckets[b] / totalSpending) * 100).toFixed(1) + "%" : "0.0%";
        cells[4].textContent = _pctOfIncome(allBuckets[b], income) + "%";
      }
    });

    if (label === "TOTAL SPENDING") {
      const cells = r.querySelectorAll("td");
      cells[2].textContent = formatDollar(totalSpending);
      cells[4].textContent = _pctOfIncome(totalSpending, income) + "%";
    }
  });
}

function _updateBubble(slider, pct) {
  const bubble = slider.closest(".slider-wrap")?.querySelector(".slider-bubble");
  if (!bubble) return;
  // Offset formula keeps the bubble tooltip centered over the thumb across the slider range
  const THUMB_HALF_WIDTH = 8;
  const pos = (slider.value - slider.min) / (slider.max - slider.min);
  bubble.style.left = `calc(${pos * 100}% + ${THUMB_HALF_WIDTH - pos * (THUMB_HALF_WIDTH * 2)}px)`;
  bubble.textContent = pct + "%";
}

function attachSliderListeners() {
  document.querySelectorAll(".preview-table .whatif-slider").forEach(slider => {
    slider.addEventListener("input", function () {
      const val    = parseFloat(this.value);
      const row    = this.closest(".cat-data-row");
      const amtCell = row.querySelector("[data-field='amt']");
      const income  = tableData.income;

      amtCell.dataset.raw  = val;
      amtCell.textContent  = formatDollar(val);

      const pct = _pctOfIncome(val, income);
      row.querySelector(".cat-pct").textContent = pct + "%";

      _updateBucketTotalRow(row.dataset.bucket, income);
      _updateGrandRevealRows(income);
      _updateBubble(this, pct);

      const resetBtn = this.closest("td").querySelector(".whatif-reset-btn");
      if (resetBtn) {
        resetBtn.classList.toggle("hidden",
          parseFloat(this.value) === parseFloat(this.dataset.original));
      }
      updateWhatIfResetAll();
    });

    slider.addEventListener("pointerdown", function () {
      const bubble = this.closest(".slider-wrap")?.querySelector(".slider-bubble");
      if (bubble) {
        const income = tableData ? tableData.income : 0;
        const pct = _pctOfIncome(parseFloat(this.value), income);
        _updateBubble(this, pct);
        bubble.classList.remove("hidden");
      }
    });

    slider.addEventListener("pointerup", function () {
      const bubble = this.closest(".slider-wrap")?.querySelector(".slider-bubble");
      if (bubble) bubble.classList.add("hidden");
    });
  });
}

const _resetAllBtn = document.getElementById("whatif-reset-all-btn");

function updateWhatIfResetAll() {
  const sliders    = document.querySelectorAll(".whatif-slider");
  const anyChanged = Array.from(sliders).some(
    s => parseFloat(s.value) !== parseFloat(s.dataset.original)
  );
  _resetAllBtn.classList.toggle("hidden", !(whatIfActive && anyChanged));
}

function resetSlider(resetBtn) {
  const slider = resetBtn.closest("td").querySelector(".whatif-slider");
  slider.value = slider.dataset.original;
  slider.dispatchEvent(new Event("input"));
}

function resetAllSliders() {
  document.querySelectorAll(".whatif-slider").forEach(slider => {
    slider.value = slider.dataset.original;
    slider.dispatchEvent(new Event("input"));
  });
}

// ── Trends ───────────────────────────────────────────────────────────────
let trendsActive       = false;
let trendsChartInstance = null;

// Line colors match the bar chart border colors
const TREND_LINE_COLORS = {
  "Necessities": "#4fa8bc",
  "Luxuries":    "#5fa87c",
  "Future Self": "#d47a56"
};
const TREND_FILL_COLORS = {
  "Necessities": "rgba(79,168,188,0.08)",
  "Luxuries":    "rgba(95,168,124,0.08)",
  "Future Self": "rgba(212,122,86,0.08)"
};

function toggleTrends() {
  // If What If is active, turn it off before entering Trends
  if (whatIfActive) {
    whatIfActive = false;
    const table = document.querySelector(".preview-table");
    document.getElementById("whatif-btn").classList.remove("active");
    document.getElementById("whatif-reset-all-btn").classList.add("hidden");
    if (table) table.classList.remove("whatif-active");
  }

  trendsActive = !trendsActive;
  document.getElementById("trends-btn").classList.toggle("active", trendsActive);
  document.getElementById("table-container").classList.toggle("hidden", trendsActive);
  const tc = document.getElementById("trends-container");
  if (trendsActive) {
    tc.classList.remove("hidden");
    renderTrends();
  } else {
    if (trendsChartInstance) {
      trendsChartInstance.destroy();
      trendsChartInstance = null;
    }
    tc.classList.add("hidden");
  }
}

function renderTrends() {
  const { breakdown, months } = trendData;
  const bucketOrder = ["Necessities", "Luxuries", "Future Self"];

  // Sum all categories within a bucket for each month
  const datasets = bucketOrder
    .filter(bucket =>
      tableData.buckets.find(b => b.name === bucket && b.categories.length > 0)
    )
    .map(bucket => {
      const cats = Object.keys(breakdown).filter(cat =>
        tableData.buckets.find(b => b.name === bucket && b.categories.find(c => c.name === cat))
      );
      const monthlyTotals = months.map(m =>
        parseFloat(cats.reduce((sum, cat) => sum + (breakdown[cat]?.[m] ?? 0), 0).toFixed(2))
      );
      return {
        label:           bucket,
        data:            monthlyTotals,
        borderColor:     TREND_LINE_COLORS[bucket],
        backgroundColor: TREND_FILL_COLORS[bucket],
        fill:            true,
        tension:         0.35,
        pointRadius:     5,
        pointHoverRadius: 7,
        borderWidth:     2.5
      };
    });

  // Replace container contents with a fresh canvas
  const tc = document.getElementById("trends-container");
  tc.innerHTML = '<div style="max-width:760px;margin:16px auto 0;"><canvas id="trends-chart"></canvas></div>';

  if (trendsChartInstance) {
    trendsChartInstance.destroy();
  }

  trendsChartInstance = new Chart(document.getElementById("trends-chart"), {
    type: "line",
    data: { labels: months, datasets },
    options: {
      responsive: true,
      plugins: {
        legend: {
          display: true,
          labels: {
            font: { family: "DM Sans", size: 13 },
            usePointStyle: true,
            padding: 20
          }
        },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${formatDollar(ctx.parsed.y)}`
          }
        },
        datalabels: { display: false }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            font: { family: "DM Sans", size: 12 },
            callback: val => "$" + val.toLocaleString()
          },
          grid: { color: "#ebebeb" }
        },
        x: {
          ticks: { font: { family: "DM Sans", size: 12, weight: "600" } },
          grid: { display: false }
        }
      }
    }
  });
}
