// Depends on: main.js (formatDollar, tableData, trendData, attachEditListeners, attachSliderListeners)
Chart.register(ChartDataLabels);

const BUCKET_COLORS = {
  "Necessities": "#dbeef7",
  "Luxuries":    "#daf2e5",
  "Future Self": "#fef0e0"
};

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function _renderBucketSection(bucket, income, totalSpending) {
  const color = BUCKET_COLORS[bucket.name] || "#f8f8f8";
  const bucketPctSpending = totalSpending
    ? ((bucket.total / totalSpending) * 100).toFixed(1) : "0.0";
  let html = "";

  html += `<tr><td colspan="6" style="background:${color};font-weight:bold;font-size:13px;padding:10px;">▶ ${escapeHtml(bucket.name.toUpperCase())}</td></tr>`;
  html += `<tr style="font-size:11px;font-weight:bold;text-align:center;">
    <td style="background:${color};">Category</td>
    <td style="background:${color};">Total Spending (all months)</td>
    <td style="background:${color};">Avg Monthly Spending</td>
    <td style="background:${color};">% of Income</td>
    <td class="pct-spending-col" style="background:${color};">% of Spending</td>
    <td style="background:${color};"></td>
  </tr>`;

  if (bucket.categories.length === 0) {
    html += `<tr><td colspan="4" style="background:${color};color:#888;">(No categories assigned)</td><td class="pct-spending-col" style="background:${color};"></td><td style="background:${color};"></td></tr>`;
  } else {
    bucket.categories.forEach(cat => {
      const pctIncome   = income        ? ((cat.raw_total / income)        * 100).toFixed(1) : "0.0";
      const pctSpending = totalSpending ? ((cat.amt       / totalSpending) * 100).toFixed(1) : "0.0";
      html += `<tr class="cat-data-row" data-bucket="${escapeHtml(bucket.name)}">
        <td class="editable" contenteditable="true" style="background:${color};" data-field="name">${escapeHtml(cat.name)}</td>
        <td style="background:${color};text-align:right;" data-field="raw_total">${formatDollar(cat.raw_total)}</td>
        <td class="editable" contenteditable="true" style="background:${color};text-align:right;" data-field="amt" data-raw="${cat.amt}">${formatDollar(cat.amt)}</td>
        <td class="cat-pct" style="background:${color};text-align:center;">${pctIncome}%</td>
        <td class="pct-spending-col" style="background:${color};text-align:center;">${pctSpending}%</td>
        <td style="background:${color};"><div class="slider-wrap"><span class="slider-bubble hidden"></span><input type="range" class="whatif-slider" min="0" max="${cat.amt * 2 || 1000}" step="1" value="${cat.amt}" data-original="${cat.amt}"><button class="whatif-reset-btn hidden" onclick="resetSlider(this)" title="Reset to original">↺</button></div></td>
      </tr>`;
    });
  }

  html += `<tr>
    <td style="background:#2c3e50;color:white;font-weight:bold;text-align:center;">TOTAL — ${escapeHtml(bucket.name)}</td>
    <td style="background:#2c3e50;color:white;font-weight:bold;text-align:right;">${formatDollar(bucket.raw_total)}</td>
    <td style="background:#2c3e50;color:white;font-weight:bold;text-align:right;">${formatDollar(bucket.total)}</td>
    <td style="background:#2c3e50;color:white;font-weight:bold;text-align:center;">${bucket.pct_income.toFixed(1)}%</td>
    <td class="pct-spending-col" style="background:#2c3e50;color:white;font-weight:bold;text-align:center;">${bucketPctSpending}%</td>
    <td style="background:#2c3e50;color:white;"></td>
  </tr>`;
  html += `<tr class="spacer-row"><td colspan="6"></td></tr>`;

  return html;
}

function _renderGrandReveal(data) {
  const remaining    = data.remaining_income;
  const remainingPct = data.income > 0 ? ((remaining / data.income) * 100).toFixed(1) : "0.0";
  const remColor     = remaining < 0 ? "#c0392b" : "#145a32";
  const remBg        = remaining < 0 ? "#FDECEA"  : "#D5F5E3";

  let html = "";
  html += `<tr class="spacer-row"><td colspan="6"></td></tr>`;
  html += `<tr><td colspan="6" style="background:#1a5276;color:white;font-weight:bold;font-size:13px;text-align:center;padding:10px;">─── GRAND REVEAL ───</td></tr>`;
  html += `<tr style="background:#1a5276;color:white;font-size:11px;font-weight:bold;text-align:center;">
    <td>Bucket</td><td>Total Spending (all months)</td><td>Avg Monthly Spending</td><td>% of Spending</td><td>% of Income</td><td></td>
  </tr>`;
  data.grand_reveal.forEach(row => {
    html += `<tr style="text-align:center;font-weight:bold;">
      <td>${escapeHtml(row.name.toUpperCase())}</td>
      <td>${formatDollar(row.raw_total)}</td>
      <td>${formatDollar(row.total)}</td>
      <td>${row.pct_spending.toFixed(1)}%</td>
      <td>${row.pct_income.toFixed(1)}%</td>
      <td></td>
    </tr>`;
  });
  html += `<tr style="background:#2c3e50;color:white;font-weight:bold;text-align:center;">
    <td>TOTAL SPENDING</td>
    <td>${formatDollar(data.total_raw_spending)}</td>
    <td>${formatDollar(data.total_spending)}</td>
    <td>100.0%</td>
    <td>${data.total_pct_income.toFixed(1)}%</td>
    <td></td>
  </tr>`;
  html += `<tr style="background:${remBg};color:${remColor};font-weight:bold;text-align:center;">
    <td>💰 Remaining Income</td>
    <td>${formatDollar(remaining)}</td>
    <td>—</td>
    <td>—</td>
    <td>${remainingPct}%</td>
    <td></td>
  </tr>`;

  return html;
}

function renderTable(data) {
  tableData = data;
  const income = data.income;

  let html = `<table class="preview-table">`;

  html += `<tr><td colspan="6" style="background:#2c3e50;color:white;font-weight:bold;font-size:14px;text-align:center;padding:12px;">THE AWARENESS ENGINE — My Financial Reality</td></tr>`;
  html += `<tr class="spacer-row"><td colspan="6"></td></tr>`;
  html += `<tr>
    <td style="font-weight:bold;">Total Income</td>
    <td style="font-weight:bold;text-align:right;">${formatDollar(income)}</td>
    <td></td>
    <td></td>
    <td class="pct-spending-col"></td>
    <td style="font-weight:bold;text-align:center;">100%</td>
  </tr>`;
  html += `<tr class="spacer-row"><td colspan="6"></td></tr>`;

  data.buckets.forEach(bucket => {
    html += _renderBucketSection(bucket, income, data.total_spending);
  });

  html += _renderGrandReveal(data);
  html += `</table>`;

  document.getElementById("table-container").innerHTML = html;
  attachEditListeners();
  attachSliderListeners();
  renderBarChart(data);
}

let barChartInstance = null;

function renderBarChart(data) {
  const totalSpending = data.total_spending || 1;
  const labels = data.buckets.map(b => b.name);
  const values = data.buckets.map(b => b.total);
  const pcts = data.buckets.map(b => parseFloat(((b.total / totalSpending) * 100).toFixed(1)));
  const colors = data.buckets.map(b => BUCKET_COLORS[b.name] || "#ccc");
  const borderColors = ["#4fa8bc", "#5fa87c", "#d47a56"];

  // Top-spending category per bucket (by avg monthly amount)
  const topCategories = data.buckets.map(b =>
    b.categories.length
      ? b.categories.reduce((max, c) => c.amt > max.amt ? c : max)
      : null
  );

  document.getElementById("bar-container").classList.remove("hidden");

  if (barChartInstance) {
    barChartInstance.destroy();
  }

  // Custom HTML tooltip — enables bolded $ amount and top-expense line
  const tooltipHandler = (context) => {
    const { chart, tooltip } = context;
    let el = chart.canvas.parentNode.querySelector('.bar-tooltip');
    if (!el) {
      el = document.createElement('div');
      el.className = 'bar-tooltip';
      el.style.cssText = [
        'position:absolute',
        'background:#0f1f3d',
        'color:white',
        'border-radius:10px',
        'padding:10px 14px',
        'font-family:DM Sans,Arial,sans-serif',
        'font-size:13px',
        'pointer-events:none',
        'transition:opacity 0.15s',
        'z-index:100',
        'min-width:170px',
        'box-shadow:0 4px 16px rgba(0,0,0,0.25)',
        'white-space:nowrap'
      ].join(';');
      chart.canvas.parentNode.style.position = 'relative';
      chart.canvas.parentNode.appendChild(el);
    }

    if (tooltip.opacity === 0) {
      el.style.opacity = '0';
      return;
    }

    const i = tooltip.dataPoints[0].dataIndex;
    const top = topCategories[i];
    el.innerHTML = `
      <div style="margin-bottom:4px;opacity:0.8;font-size:12px;">${pcts[i]}% of spending</div>
      <div style="font-size:18px;font-weight:700;letter-spacing:0.3px;">${formatDollar(values[i])}</div>
      ${top ? `
      <div style="margin-top:8px;padding-top:8px;border-top:1px solid rgba(255,255,255,0.2);font-size:12px;">
        ★ Top expense<br>
        <strong>${escapeHtml(top.name)}</strong><br>
        <span style="opacity:0.75;">${formatDollar(top.amt)}/mo avg</span>
      </div>` : ''}
    `;

    const { offsetLeft: posX, offsetTop: posY } = chart.canvas;
    el.style.opacity = '1';
    el.style.left = (posX + tooltip.caretX + 14) + 'px';
    el.style.top  = (posY + tooltip.caretY - 10) + 'px';
  };

  barChartInstance = new Chart(document.getElementById("bar-chart"), {
    type: "bar",
    data: {
      labels: labels,
      datasets: [{
        data: pcts,
        backgroundColor: colors,
        borderColor: borderColors,
        borderWidth: 2,
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        datalabels: {
          anchor: 'end',
          align: 'top',
          formatter: val => val + '%',
          font: { family: 'DM Sans', size: 13, weight: '600' },
          color: '#333'
        },
        tooltip: {
          enabled: false,
          external: tooltipHandler
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          max: 100,
          ticks: {
            font: { family: "DM Sans", size: 12 },
            callback: val => val + "%"
          },
          grid: { color: "#ebebeb" }
        },
        x: {
          ticks: { font: { family: "DM Sans", size: 13, weight: "600" } },
          grid: { display: false }
        }
      }
    }
  });
}
