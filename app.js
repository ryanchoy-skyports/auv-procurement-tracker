const state = {
  rows: [],
  filterStatus: "All",
  search: "",
  live: false,
};

const el = {
  authArea: document.getElementById("auth-area"),
  banner: document.getElementById("banner"),
  totalEstimated: document.getElementById("total-estimated"),
  totalConfirmed: document.getElementById("total-confirmed"),
  totalItems: document.getElementById("total-items"),
  totalDone: document.getElementById("total-done"),
  filterStatus: document.getElementById("filter-status"),
  search: document.getElementById("search"),
  lastSync: document.getElementById("last-sync"),
  tbody: document.getElementById("tbody"),
};

const STATUS_CLASS = {
  Done: "done",
  Pending: "pending",
  "On Hold": "on-hold",
  "Not Started": "not-started",
};

function money(value, currency) {
  if (value === null || value === undefined || value === "") return "—";
  return `${Number(value).toLocaleString(undefined, { maximumFractionDigits: 0 })} ${currency || ""}`.trim();
}

function renderTotals() {
  const visible = state.rows;
  const totalEstimated = visible.reduce((sum, r) => sum + (Number(r.estimatedCost) || 0), 0);
  const totalConfirmed = visible.reduce((sum, r) => sum + (Number(r.confirmedCost) || 0), 0);
  const doneCount = visible.filter((r) => r.status === "Done").length;

  el.totalEstimated.textContent = money(totalEstimated, "AED");
  el.totalConfirmed.textContent = money(totalConfirmed, "AED");
  el.totalItems.textContent = String(visible.length);
  el.totalDone.textContent = `${doneCount} / ${visible.length}`;
}

function matchesFilters(row) {
  if (state.filterStatus !== "All" && row.status !== state.filterStatus) return false;
  if (state.search) {
    const q = state.search.toLowerCase();
    return (
      row.item.toLowerCase().includes(q) ||
      row.owner.toLowerCase().includes(q) ||
      row.update.toLowerCase().includes(q) ||
      row.nextAction.toLowerCase().includes(q)
    );
  }
  return true;
}

function renderRows() {
  el.tbody.innerHTML = "";
  const filtered = state.rows.filter(matchesFilters);

  for (const row of filtered) {
    const tr = document.createElement("tr");
    tr.dataset.id = row.id;

    tr.innerHTML = `
      <td>
        <div class="item-name">${escapeHtml(row.item)}</div>
        <div class="item-owner">${escapeHtml(row.owner)}${row.coreFunction ? " · " + escapeHtml(row.coreFunction) : ""}</div>
      </td>
      <td>
        <select class="status-select ${STATUS_CLASS[row.status] || ""}" data-field="status">
          ${["Not Started", "Pending", "On Hold", "Done"].map((s) => `<option value="${s}" ${s === row.status ? "selected" : ""}>${s}</option>`).join("")}
        </select>
      </td>
      <td class="note-col"><textarea data-field="update" rows="2" placeholder="Status note">${escapeHtml(row.update)}</textarea></td>
      <td class="note-col"><textarea data-field="nextAction" rows="2" placeholder="Next action">${escapeHtml(row.nextAction)}</textarea></td>
      <td class="cost-col"><input type="number" data-field="estimatedCost" value="${row.estimatedCost ?? ""}" placeholder="0"></td>
      <td class="cost-col"><input type="number" data-field="confirmedCost" value="${row.confirmedCost ?? ""}" placeholder="0"></td>
      <td>
        <select data-field="costStatus">
          ${["Estimate", "Quote Received", "PO Issued", "Invoiced", "Paid"].map((s) => `<option value="${s}" ${s === row.costStatus ? "selected" : ""}>${s}</option>`).join("")}
        </select>
      </td>
      <td><span class="saving-flag">saved</span></td>
    `;

    tr.querySelectorAll("[data-field]").forEach((input) => {
      input.addEventListener("change", () => onFieldChange(row, input));
    });

    el.tbody.appendChild(tr);
  }

  renderTotals();
}

function escapeHtml(s) {
  return String(s || "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

async function onFieldChange(row, input) {
  const field = input.dataset.field;
  let value = input.value;
  if (field === "estimatedCost" || field === "confirmedCost") {
    value = value === "" ? null : Number(value);
  }
  row[field] = value;

  const tr = input.closest("tr");
  if (field === "status") {
    input.className = `status-select ${STATUS_CLASS[value] || ""}`;
  }

  tr.classList.add("saving");
  try {
    if (state.live) {
      await window.SkyportsGraph.updateItem(row.id, row);
    }
  } catch (err) {
    console.error(err);
    alert(`Could not save "${row.item}": ${err.message}`);
  } finally {
    tr.classList.remove("saving");
    renderTotals();
    el.lastSync.textContent = `Last saved ${new Date().toLocaleTimeString()}`;
  }
}

function seedRowsToState(seedRows) {
  state.rows = seedRows.map((r, i) => ({ id: r.id ?? `demo-${i}`, ...r }));
}

async function loadDemo() {
  const seed = await fetch("data/seed.json").then((r) => r.json());
  seedRowsToState(seed);
  renderRows();
  el.lastSync.textContent = "Demo mode — changes are not saved anywhere";
}

async function loadLive() {
  const rows = await window.SkyportsGraph.listItems();
  state.rows = rows;
  renderRows();
  el.lastSync.textContent = `Synced with SharePoint at ${new Date().toLocaleTimeString()}`;
}

function renderAuthArea() {
  el.authArea.innerHTML = "";
  if (!window.SkyportsGraph.isConfigured()) {
    const pill = document.createElement("span");
    pill.className = "mode-pill";
    pill.textContent = "Demo mode (not connected to SharePoint yet)";
    el.authArea.appendChild(pill);
    return;
  }

  const btn = document.createElement("button");
  btn.className = "btn";
  btn.textContent = state.live ? "Signed in" : "Sign in with Microsoft";
  btn.disabled = state.live;
  btn.addEventListener("click", () => {
    // loginRedirect navigates away immediately; the return trip is handled in init().
    btn.disabled = true;
    btn.textContent = "Signing in...";
    window.SkyportsGraph.signIn().catch((err) => {
      console.error(err);
      alert(`Sign-in failed: ${err.message}`);
      renderAuthArea();
    });
  });
  el.authArea.appendChild(btn);
}

function wireToolbar() {
  el.filterStatus.addEventListener("change", () => {
    state.filterStatus = el.filterStatus.value;
    renderRows();
  });
  el.search.addEventListener("input", () => {
    state.search = el.search.value;
    renderRows();
  });
}

function showBanner(message, kind) {
  el.banner.textContent = message;
  el.banner.className = kind === "error" ? "banner error" : "banner";
  el.banner.style.display = "block";
}
function hideBanner() {
  el.banner.style.display = "none";
}

async function init() {
  wireToolbar();

  if (!window.SkyportsGraph.isConfigured()) {
    hideBanner();
    renderAuthArea();
    await loadDemo();
    return;
  }

  let account = null;
  try {
    account = await window.SkyportsGraph.initAuth();
  } catch (err) {
    console.error(err);
    showBanner(`Sign-in failed: ${err.message}`, "error");
    renderAuthArea();
    await loadDemo();
    return;
  }

  if (!account) {
    showBanner("Not signed in yet — showing demo data. Click Sign in with Microsoft above to load the live SharePoint list.");
    renderAuthArea();
    await loadDemo();
    return;
  }

  try {
    await loadLive();
    state.live = true;
    hideBanner();
    renderAuthArea();
  } catch (err) {
    console.error(err);
    state.live = false;
    showBanner(`Signed in as ${account.username}, but couldn't load the SharePoint list: ${err.message}`, "error");
    renderAuthArea();
    await loadDemo();
  }
}

init();
