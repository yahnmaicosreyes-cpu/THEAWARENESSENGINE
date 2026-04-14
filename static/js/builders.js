// Depends on: main.js (formatDollar)

// ── Single Month Budget Builder ───────────────────────────────────────────
function tmplKey(bucket) { return bucket.replace(/\s+/g, "-").toLowerCase(); }

const TMPL_DEFAULTS = {
  "Necessities": ["Rent / Mortgage", "Groceries", "Utilities", "Transportation", "Insurance", "Phone Bill"],
  "Luxuries":    ["Dining Out", "Entertainment", "Shopping", "Subscriptions", "Personal Care"],
  "Future Self": ["Emergency Fund", "Retirement (401k/IRA)", "Investments", "Education / Skills", "Debt Payoff"]
};

function tmplAddRow(bucket, name = "", amount = "") {
  const container = document.getElementById("tmpl-rows-" + tmplKey(bucket));
  const row = document.createElement("tr");
  row.className = "tmpl-row";
  row.innerHTML = `
    <td><input type="text" placeholder="Category name" value="${name}" oninput="tmplRecalculate()"></td>
    <td style="width:120px;"><input type="number" placeholder="0.00" value="${amount}" min="0" step="0.01" oninput="tmplRecalculate()"></td>
    <td style="width:36px;text-align:center;"><button class="tmpl-remove-btn" onclick="this.closest('.tmpl-row').remove(); tmplRecalculate();" title="Remove">✕</button></td>
  `;
  container.appendChild(row);
  tmplRecalculate();
}

function tmplRecalculate() {
  const income = parseFloat(document.getElementById("tmpl-income").value) || 0;
  const subtotals = {};
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    const key = tmplKey(bucket);
    const rows = document.querySelectorAll("#tmpl-rows-" + key + " .tmpl-row");
    let subtotal = 0;
    rows.forEach(row => {
      const amt = parseFloat(row.querySelector("input[type='number']").value) || 0;
      subtotal += amt;
    });
    subtotals[key] = subtotal;
    document.getElementById("tmpl-sub-" + key).textContent = formatDollar(subtotal);
    const pctEl = document.getElementById("tmpl-pct-" + key);
    pctEl.textContent = income > 0 ? ((subtotal / income) * 100).toFixed(1) + "%" : "0%";
  });

  // Grand Reveal
  const totalExp = Object.values(subtotals).reduce((a, b) => a + b, 0);
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    const key = tmplKey(bucket);
    const amt = subtotals[key];
    document.getElementById("gr-amt-" + key).textContent = formatDollar(amt);
    document.getElementById("gr-inc-" + key).textContent = income > 0 ? ((amt / income) * 100).toFixed(1) + "%" : "0%";
    document.getElementById("gr-spd-" + key).textContent = totalExp > 0 ? ((amt / totalExp) * 100).toFixed(1) + "%" : "0%";
  });
  document.getElementById("gr-total-expenses").textContent = formatDollar(totalExp);
  document.getElementById("gr-total-pct-inc").textContent = income > 0 ? ((totalExp / income) * 100).toFixed(1) + "%" : "0%";
  document.getElementById("gr-total-income").textContent = formatDollar(income);
  const remaining = income - totalExp;
  document.getElementById("gr-remaining").textContent = formatDollar(remaining);
  document.getElementById("gr-remaining-pct").textContent = income > 0 ? ((remaining / income) * 100).toFixed(1) + "%" : "0%";
  document.getElementById("gr-remaining-row").style.color = remaining < 0 ? "#c0392b" : "#145a32";
}

async function tmplDownload() {
  const income = parseFloat(document.getElementById("tmpl-income").value) || 0;
  const buckets = {};
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    const rows = document.querySelectorAll("#tmpl-rows-" + tmplKey(bucket) + " .tmpl-row");
    buckets[bucket] = [];
    rows.forEach(row => {
      const name = row.querySelector("input[type='text']").value.trim() || "Category";
      const amount = parseFloat(row.querySelector("input[type='number']").value) || 0;
      buckets[bucket].push({ name, amount });
    });
  });

  const msgEl = document.getElementById("tmpl-msg");
  msgEl.textContent = "";
  try {
    const resp = await fetch("/template/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ income, buckets })
    });
    if (!resp.ok) throw new Error("Server error: " + resp.status);
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Budget_Template.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch (e) {
    msgEl.textContent = "Download failed: " + e.message;
  }
}

// Populate default rows on load
(function initTemplate() {
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    TMPL_DEFAULTS[bucket].forEach(name => tmplAddRow(bucket, name, ""));
  });
})();

// ── Last 3 Months Budget Builder ─────────────────────────────────────────
// Reuses TMPL_DEFAULTS — same default categories as the single-month builder.

function tmThreeAddRow(bucket, name = "") {
  const container = document.getElementById("tm-rows-" + tmplKey(bucket));
  const row = document.createElement("tr");
  row.className = "tmpl-row";
  row.innerHTML = `
    <td><input type="text" placeholder="Category name" value="${name}" oninput="tmThreeRecalculate()"></td>
    <td style="width:110px;"><input type="number" placeholder="0.00" min="0" step="0.01" oninput="tmThreeRecalculate()"></td>
    <td style="width:110px;"><input type="number" placeholder="0.00" min="0" step="0.01" oninput="tmThreeRecalculate()"></td>
    <td style="width:110px;"><input type="number" placeholder="0.00" min="0" step="0.01" oninput="tmThreeRecalculate()"></td>
    <td style="width:90px;text-align:right;font-weight:600;color:#0f1f3d;padding:6px 10px;" class="tm-row-avg">$0.00</td>
    <td style="width:36px;text-align:center;"><button class="tmpl-remove-btn" onclick="this.closest('.tmpl-row').remove(); tmThreeRecalculate();" title="Remove">✕</button></td>
  `;
  container.appendChild(row);
  tmThreeRecalculate();
}

function tmThreeUpdateHeaders() {
  [1, 2, 3].forEach(i => {
    const name = document.getElementById(`tm-month${i}-name`).value || `Month ${i}`;
    document.querySelectorAll(`.tm-col-m${i}`).forEach(el => el.textContent = name);
  });
}

function tmThreeRecalculate() {
  tmThreeUpdateHeaders();
  const incomes = [1, 2, 3].map(i => parseFloat(document.getElementById(`tm-income-${i}`).value) || 0);
  const avgIncome = (incomes[0] + incomes[1] + incomes[2]) / 3;

  const subtotalAvgs = {};
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    const key = tmplKey(bucket);
    const rows = document.querySelectorAll(`#tm-rows-${key} .tmpl-row`);
    let sub1 = 0, sub2 = 0, sub3 = 0;
    rows.forEach(row => {
      const amts = row.querySelectorAll("input[type='number']");
      const a1 = parseFloat(amts[0].value) || 0;
      const a2 = parseFloat(amts[1].value) || 0;
      const a3 = parseFloat(amts[2].value) || 0;
      sub1 += a1; sub2 += a2; sub3 += a3;
      row.querySelector(".tm-row-avg").textContent = formatDollar((a1 + a2 + a3) / 3);
    });
    const subAvg = (sub1 + sub2 + sub3) / 3;
    subtotalAvgs[key] = subAvg;
    document.getElementById(`tm-sub-${key}`).textContent = formatDollar(subAvg);
    document.getElementById(`tm-pct-${key}`).textContent =
      avgIncome > 0 ? ((subAvg / avgIncome) * 100).toFixed(1) + "%" : "0%";
  });

  const totalExp = Object.values(subtotalAvgs).reduce((a, b) => a + b, 0);
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    const key = tmplKey(bucket);
    const amt = subtotalAvgs[key];
    document.getElementById(`tm-gr-amt-${key}`).textContent = formatDollar(amt);
    document.getElementById(`tm-gr-inc-${key}`).textContent =
      avgIncome > 0 ? ((amt / avgIncome) * 100).toFixed(1) + "%" : "0%";
    document.getElementById(`tm-gr-spd-${key}`).textContent =
      totalExp > 0 ? ((amt / totalExp) * 100).toFixed(1) + "%" : "0%";
  });
  document.getElementById("tm-gr-total-expenses").textContent = formatDollar(totalExp);
  document.getElementById("tm-gr-total-pct-inc").textContent =
    avgIncome > 0 ? ((totalExp / avgIncome) * 100).toFixed(1) + "%" : "0%";
  document.getElementById("tm-gr-total-income").textContent = formatDollar(avgIncome);
  const remaining = avgIncome - totalExp;
  document.getElementById("tm-gr-remaining").textContent = formatDollar(remaining);
  document.getElementById("tm-gr-remaining-pct").textContent =
    avgIncome > 0 ? ((remaining / avgIncome) * 100).toFixed(1) + "%" : "0%";
  document.getElementById("tm-gr-remaining-row").style.color = remaining < 0 ? "#c0392b" : "#145a32";
}

async function tmThreeDownload() {
  const incomes = [1, 2, 3].map(i => parseFloat(document.getElementById(`tm-income-${i}`).value) || 0);
  const monthNames = [1, 2, 3].map(i => document.getElementById(`tm-month${i}-name`).value.trim() || `Month ${i}`);
  const buckets = {};
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    const rows = document.querySelectorAll(`#tm-rows-${tmplKey(bucket)} .tmpl-row`);
    buckets[bucket] = [];
    rows.forEach(row => {
      const name = row.querySelector("input[type='text']").value.trim() || "Category";
      const amts = row.querySelectorAll("input[type='number']");
      buckets[bucket].push({ name, amounts: [0, 1, 2].map(i => parseFloat(amts[i].value) || 0) });
    });
  });
  const msgEl = document.getElementById("tm-msg");
  msgEl.textContent = "";
  try {
    const resp = await fetch("/three-month/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ incomes, month_names: monthNames, buckets })
    });
    if (!resp.ok) throw new Error("Server error: " + resp.status);
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "Last_3_Months_Budget.xlsx";
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
  } catch (e) { msgEl.textContent = "Download failed: " + e.message; }
}

(function initThreeMonth() {
  ["Necessities", "Luxuries", "Future Self"].forEach(bucket => {
    TMPL_DEFAULTS[bucket].forEach(name => tmThreeAddRow(bucket, name));
  });
})();
