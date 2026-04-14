// ── State ─────────────────────────────────────────────────────────────────
let selectedFile      = null;
let categories        = [];           // [{ name, amount }]
let assignments       = {};           // { cat_name: bucket_label }
let tableData         = null;         // current rendered table data
let originalTableData = null;         // snapshot at last save/generate
let trendData         = null;         // { monthly_breakdown, trend_months }
const BUCKETS = ["Necessities", "Luxuries", "Future Self", "Leave Out"];

// ── Helpers ───────────────────────────────────────────────────────────────
function formatDollar(n) {
  const abs = Math.abs(Number(n));
  const s = "$" + abs.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return n < 0 ? "-" + s : s;
}

function resetApp() {
  location.reload();
}

// ── Step 1: File select ──────────────────────────────────────────────────
function handleFileSelect(input) {
  selectedFile = input.files[0];
  document.getElementById("file-name").textContent =
    selectedFile ? `Selected: ${selectedFile.name}` : "";
  document.getElementById("upload-btn").disabled = !selectedFile;
}

async function uploadFile() {
  const btn = document.getElementById("upload-btn");
  const msg = document.getElementById("upload-msg");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Scanning...';
  msg.innerHTML = "";

  const formData = new FormData();
  formData.append("file", selectedFile);

  try {
    const res  = await fetch("/upload", { method: "POST", body: formData });
    const data = await res.json();

    if (data.error) {
      msg.innerHTML = `<div class="msg msg-error">❌ ${escapeHtml(data.error)}</div>`;
      btn.disabled = false;
      btn.textContent = "Scan My Spreadsheet";
      return;
    }

    // Show expenses and income
    const totalExpenses = data.categories.reduce((sum, c) => sum + c.amount, 0);
    document.getElementById("expenses-display").textContent = formatDollar(totalExpenses);
    document.getElementById("income-display").textContent = formatDollar(data.income);

    // Build month checkboxes
    const monthBox = document.getElementById("month-checkboxes");
    monthBox.innerHTML = "";
    data.months_found.forEach(m => {
      const label = document.createElement("label");
      label.style.cssText = "display:flex;align-items:center;gap:4px;cursor:pointer;";
      label.innerHTML = `<input type="checkbox" value="${escapeHtml(m)}" checked> ${escapeHtml(m)}`;
      monthBox.appendChild(label);
    });

    // Build category rows
    categories = data.categories;
    assignments = {};
    buildCategoryRows(categories);

    // Transition to step 2
    document.getElementById("step1-card").style.opacity = "0.5";
    document.getElementById("step2-card").classList.remove("hidden");
    document.getElementById("step2-card").scrollIntoView({ behavior: "smooth" });

    btn.innerHTML = "✅ Done";
    msg.innerHTML =
      `<div class="msg msg-success">✅ Found data from: ${data.months_found.map(escapeHtml).join(", ")}</div>`;
    return;

  } catch (e) {
    msg.innerHTML = `<div class="msg msg-error">❌ Something went wrong: ${e.message}</div>`;
  }

  btn.disabled = false;
  btn.textContent = "Scan My Spreadsheet";
}

// ── Step 2: Category assignment ──────────────────────────────────────────
function buildCategoryRows(cats) {
  const list = document.getElementById("category-list");
  list.innerHTML = "";

  cats.forEach(cat => {
    const row = document.createElement("div");
    row.className = "cat-row";
    row.id = `row-${sanitizeId(cat.name)}`;

    const info = document.createElement("div");
    info.className = "cat-info";
    info.innerHTML = `
      <div class="cat-name">${escapeHtml(cat.name)}</div>
      <div class="cat-avg">Avg monthly: ${formatDollar(cat.amount)}</div>
    `;

    const choices = document.createElement("div");
    choices.className = "bucket-choices";

    BUCKETS.forEach(bucket => {
      const btn = document.createElement("button");
      btn.className = "choice-btn";
      btn.textContent = bucket;
      btn.dataset.bucket = bucket;
      btn.onclick = () => selectBucket(cat.name, bucket);
      choices.appendChild(btn);
    });

    row.appendChild(info);
    row.appendChild(choices);
    list.appendChild(row);
  });

  updateProgress();
}

function sanitizeId(name) {
  return name.replace(/[^a-zA-Z0-9]/g, "_");
}

function selectBucket(catName, bucket) {
  assignments[catName] = bucket;

  const row = document.getElementById(`row-${sanitizeId(catName)}`);
  row.querySelectorAll(".choice-btn").forEach(btn => {
    btn.className = "choice-btn";
    if (btn.dataset.bucket === bucket) {
      const key = bucket.replace(" ", "");
      btn.classList.add(`selected-${key}`);
    }
  });

  updateProgress();
}

function updateProgress() {
  const total   = categories.length;
  const labeled = Object.keys(assignments).length;
  const pct     = total ? Math.round((labeled / total) * 100) : 0;

  document.getElementById("progress-fill").style.width = pct + "%";
  document.getElementById("progress-text").textContent =
    `${labeled} of ${total} labeled`;
  document.getElementById("generate-btn").disabled = (labeled < total);
}

// ── Step 3: Generate ─────────────────────────────────────────────────────
async function generateReport() {
  const btn = document.getElementById("generate-btn");
  const msg = document.getElementById("generate-msg");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Building your report...';
  msg.innerHTML = "";

  try {
    const selected_months = Array.from(
      document.querySelectorAll("#month-checkboxes input:checked")
    ).map(cb => cb.value);

    const res  = await fetch("/generate", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ assignments, selected_months })
    });
    const data = await res.json();

    if (data.error) {
      msg.innerHTML = `<div class="msg msg-error">❌ ${escapeHtml(data.error)}</div>`;
      btn.disabled = false;
      btn.textContent = "⚡ Generate My Awareness Report";
      return;
    }

    renderTable(data.table);
    originalTableData = JSON.parse(JSON.stringify(data.table));
    trendData = { breakdown: data.monthly_breakdown, months: data.trend_months };

    btn.innerHTML = "✅ Done";

    document.getElementById("step2-card").style.opacity = "0.5";
    document.getElementById("step3-card").classList.remove("hidden");
    document.getElementById("step3-card").scrollIntoView({ behavior: "smooth" });

  } catch (e) {
    msg.innerHTML = `<div class="msg msg-error">❌ ${e.message}</div>`;
    btn.disabled = false;
    btn.textContent = "⚡ Generate My Awareness Report";
  }
}

// ── Inline editing ────────────────────────────────────────────────────────
function attachEditListeners() {
  document.querySelectorAll(".preview-table .editable").forEach(cell => {
    cell.addEventListener("focus", function () {
      if (this.dataset.field === "amt") {
        this.textContent = this.dataset.raw;
        const range = document.createRange();
        range.selectNodeContents(this);
        range.collapse(false);
        window.getSelection().removeAllRanges();
        window.getSelection().addRange(range);
      }
    });
    cell.addEventListener("blur", function () {
      if (this.dataset.field === "amt") {
        const val = parseFloat(this.textContent.replace(/[^0-9.-]/g, "")) || 0;
        this.dataset.raw = val;
        this.textContent = formatDollar(val);
      }
      markDirty(this);
    });
  });
}

function markDirty(cell) {
  cell.classList.add("edited");
}

// ── Save overlay ──────────────────────────────────────────────────────────
function showSaveOverlay() {
  document.getElementById("save-overlay").classList.remove("hidden");
}

function confirmDiscard() {
  document.getElementById("save-overlay").classList.add("hidden");
  document.getElementById("save-msg").innerHTML = "";
  renderTable(JSON.parse(JSON.stringify(originalTableData)));
}

async function confirmSave() {
  document.getElementById("save-overlay").classList.add("hidden");
  const msg = document.getElementById("save-msg");
  msg.innerHTML = "";

  const editedBuckets = {};
  document.querySelectorAll(".preview-table .cat-data-row").forEach(row => {
    const bucket = row.dataset.bucket;
    if (!editedBuckets[bucket]) editedBuckets[bucket] = [];
    const nameCell = row.querySelector("[data-field='name']");
    const amtCell  = row.querySelector("[data-field='amt']");
    const amt = parseFloat(String(amtCell.dataset.raw || amtCell.textContent).replace(/[^0-9.-]/g, "")) || 0;
    editedBuckets[bucket].push({ name: nameCell.textContent.trim(), amt });
  });

  try {
    const res  = await fetch("/save", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ income: tableData.income, buckets: editedBuckets })
    });
    const data = await res.json();

    if (data.error) {
      msg.innerHTML = `<div class="msg msg-error">❌ ${escapeHtml(data.error)}</div>`;
      return;
    }

    originalTableData = JSON.parse(JSON.stringify(data.table));
    renderTable(data.table);
    msg.innerHTML = `<div class="msg msg-success">✅ Changes saved and report regenerated.</div>`;
  } catch (e) {
    msg.innerHTML = `<div class="msg msg-error">❌ ${e.message}</div>`;
  }
}

// ── Navigation ────────────────────────────────────────────────────────────
function goHome() {
  document.getElementById("home-screen").classList.remove("hidden");
  document.getElementById("builder-choice-screen").classList.add("hidden");
  document.getElementById("step1-card").classList.add("hidden");
  document.getElementById("template-card").classList.add("hidden");
  document.getElementById("three-month-card").classList.add("hidden");
}

function chooseUpload() {
  document.getElementById("home-screen").classList.add("hidden");
  document.getElementById("step1-card").classList.remove("hidden");
  document.getElementById("step1-card").scrollIntoView({ behavior: "smooth" });
}

function chooseBuild() {
  document.getElementById("home-screen").classList.add("hidden");
  document.getElementById("builder-choice-screen").classList.remove("hidden");
  document.getElementById("builder-choice-screen").scrollIntoView({ behavior: "smooth" });
}

function chooseSingleMonth() {
  document.getElementById("builder-choice-screen").classList.add("hidden");
  document.getElementById("template-card").classList.remove("hidden");
  document.getElementById("template-card").scrollIntoView({ behavior: "smooth" });
}

function chooseThreeMonth() {
  document.getElementById("builder-choice-screen").classList.add("hidden");
  document.getElementById("three-month-card").classList.remove("hidden");
  document.getElementById("three-month-card").scrollIntoView({ behavior: "smooth" });
}
