/* SharePoint REST API + Search Tab — Holostik Dashboard
   With localStorage caching for fast loads + incremental refresh */

var D = {so:[],tdd:[],disp:[],soH:[],tddH:[],dispH:[]};
var results = [];
var filterMode = "all";
var sortC = -1, sortDir = 1;

var SP_BASE = "https://holostiktest-my.sharepoint.com/personal/amit_raina_holostik_com";
var SP_SO_URL   = SP_BASE + "/_api/web/lists/getbytitle('SOLog')/items?$top=2000&$orderby=Modified desc";
var SP_TDD_URL  = SP_BASE + "/_api/web/lists/getbytitle('TDDLog')/items?$top=2000&$orderby=Modified desc";
var SP_DISP_URL = SP_BASE + "/_api/web/lists/getbytitle('DispatchLog')/items?$top=2000&$orderby=Modified desc";

var CACHE_KEY = "holostik_dashboard_cache_v1";
var CACHE_TTL_MS = 1000 * 60 * 60 * 6; // 6 hours

/* ── Cache helpers ─────────────────────────────────────── */
function saveCache() {
  try {
    var payload = {
      so: D.so, tdd: D.tdd, disp: D.disp,
      soH: D.soH, tddH: D.tddH, dispH: D.dispH,
      savedAt: Date.now()
    };
    localStorage.setItem(CACHE_KEY, JSON.stringify(payload));
  } catch(e) {
    console.warn("Could not save cache:", e.message);
  }
}

function loadCache() {
  try {
    var raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return null;
    var payload = JSON.parse(raw);
    return payload;
  } catch(e) {
    return null;
  }
}

function cacheAge(payload) {
  if (!payload || !payload.savedAt) return Infinity;
  return Date.now() - payload.savedAt;
}

function formatAge(ms) {
  var mins = Math.floor(ms / 60000);
  if (mins < 1) return "just now";
  if (mins < 60) return mins + " min ago";
  var hrs = Math.floor(mins / 60);
  if (hrs < 24) return hrs + "h " + (mins % 60) + "m ago";
  return Math.floor(hrs / 24) + " day(s) ago";
}

/* ── Clean SharePoint row ──────────────────────────────── */
function cleanSPRow(spRow) {
  var SKIP = ["__metadata","odata.type","odata.id","odata.etag","odata.editLink",
    "FileSystemObjectType","Id","ID","ServerRedirectedEmbedUri","ServerRedirectedEmbedUrl",
    "ContentTypeId","Title","AuthorId","EditorId","Modified","Created",
    "OData__UIVersionString","Attachments","GUID","ComplianceAssetId"];
  var row = {};
  Object.keys(spRow).forEach(function(k) {
    if (SKIP.indexOf(k) === -1 && typeof spRow[k] !== "object" && spRow[k] !== null) {
      var label = k
        .replace(/_x0020_/g," ").replace(/_x002e_/g,".")
        .replace(/_x002f_/g,"/").replace(/_x002d_/g,"-")
        .replace(/_x0028_/g,"(").replace(/_x0029_/g,")");
      row[label] = String(spRow[k]);
    }
  });
  return row;
}

async function fetchSP(url) {
  var resp = await fetch(url, {
    method: "GET",
    headers: {
      "Accept": "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose"
    },
    credentials: "include"
  });
  if (!resp.ok) {
    throw new Error("SharePoint returned " + resp.status + ". Open SharePoint in a new tab, sign in, then click Refresh.");
  }
  var json = await resp.json();
  return json.d ? json.d.results : (json.value || []);
}

/* ── Load — uses cache first, then fetches in background ── */
async function loadFromSharePoint(forceFresh) {
  var statusEl = document.getElementById("srch-status");
  if (!statusEl) return;

  // STEP 1: Show cached data immediately if available
  var cached = loadCache();
  var usedCache = false;
  if (cached && !forceFresh) {
    D.so = cached.so || []; D.tdd = cached.tdd || []; D.disp = cached.disp || [];
    D.soH = cached.soH || []; D.tddH = cached.tddH || []; D.dispH = cached.dispH || [];
    document.getElementById("c-so").textContent   = D.so.length;
    document.getElementById("c-tdd").textContent  = D.tdd.length;
    document.getElementById("c-disp").textContent = D.disp.length;
    statusEl.textContent = "Cached: " + D.so.length + " SO  " + D.tdd.length + " TDD  " + D.disp.length + " Dispatch  (saved " + formatAge(cacheAge(cached)) + ")";
    buildHeaders();
    doSearch();
    usedCache = true;

    // If cache is fresh enough, don't auto-refetch — wait for manual Refresh
    if (cacheAge(cached) < CACHE_TTL_MS) {
      return;
    }
  }

  // STEP 2: Fetch fresh data from SharePoint (background or full load)
  if (!usedCache) {
    document.getElementById("rtbl-body").innerHTML =
      "<tr><td colspan='10' class='empty-msg'><div class='no-icon'>⏳</div>Loading live data from SharePoint...</td></tr>";
  }
  statusEl.textContent = usedCache ? "Refreshing in background..." : "Fetching SO Log...";

  try {
    var soRaw = await fetchSP(SP_SO_URL);
    statusEl.textContent = "Fetching TDD Log...";
    var tddRaw = await fetchSP(SP_TDD_URL);
    statusEl.textContent = "Fetching Dispatch Log...";
    var dispRaw = await fetchSP(SP_DISP_URL);

    D.so   = soRaw.map(function(r)  { return cleanSPRow(r); });
    D.tdd  = tddRaw.map(function(r) { return cleanSPRow(r); });
    D.disp = dispRaw.map(function(r){ return cleanSPRow(r); });
    D.soH   = D.so.length   ? Object.keys(D.so[0])   : ["SO Number"];
    D.tddH  = D.tdd.length  ? Object.keys(D.tdd[0])  : ["SO No."];
    D.dispH = D.disp.length ? Object.keys(D.disp[0]) : ["Sales Order No."];

    document.getElementById("c-so").textContent   = D.so.length;
    document.getElementById("c-tdd").textContent  = D.tdd.length;
    document.getElementById("c-disp").textContent = D.disp.length;

    var t = new Date().toLocaleTimeString("en-IN");
    statusEl.textContent = "Live: " + D.so.length + " SO  " + D.tdd.length + " TDD  " + D.disp.length + " Dispatch  Updated: " + t;

    saveCache();
    buildHeaders();
    doSearch();
  } catch(err) {
    if (usedCache) {
      // Keep showing cached data, just note the refresh failed
      statusEl.textContent = "Showing cached data — refresh failed: " + err.message;
    } else {
      statusEl.textContent = "Error loading data";
      document.getElementById("rtbl-body").innerHTML =
        "<tr><td colspan='10' class='empty-msg' style='color:#C00000'>" +
        "<div class='no-icon'>⚠️</div><strong>Could not load from SharePoint</strong><br><br>" +
        err.message + "<br><br>" +
        "<a href='" + SP_BASE + "' target='_blank' style='color:#1F4E79'>Open SharePoint</a>, sign in, then click Refresh." +
        "</td></tr>";
    }
    console.error(err);
  }
}

/* ── Manual refresh — always forces fresh fetch ────────── */
function manualRefresh() {
  loadFromSharePoint(true);
}

function soKey(row) {
  return (row["SO Number"] || row["SO No."] || row["Sales Order No."] || "").trim();
}

function gv(row) {
  return row ? Object.values(row).join(" ") : "";
}

function buildHeaders() {
  var head = document.getElementById("rtbl-head");
  if (!head) return;
  var html = "<tr>";
  D.soH.forEach(function(h, i){
    html += "<th class='th-so' onclick='sortBy(" + i + ",0)'>" + h + " &#8645;</th>";
  });
  D.tddH.forEach(function(h, i){
    html += "<th class='th-tdd' onclick='sortBy(" + i + ",1)'>" + h + " &#8645;</th>";
  });
  D.dispH.forEach(function(h, i){
    html += "<th class='th-disp' onclick='sortBy(" + i + ",2)'>" + h + " &#8645;</th>";
  });
  html += "<th class='th-stat'>SO</th><th class='th-stat'>TDD</th><th class='th-stat'>Dispatch</th></tr>";
  head.innerHTML = html;
}

function doSearch() {
  var numEl = document.getElementById("s-num");
  var txtEl = document.getElementById("s-txt");
  var numQ = numEl ? numEl.value.trim().toLowerCase() : "";
  var txtQ = txtEl ? txtEl.value.trim().toLowerCase() : "";

  if (!D.so.length && !D.tdd.length && !D.disp.length) {
    document.getElementById("rtbl-body").innerHTML =
      "<tr><td colspan='10' class='empty-msg'>No data yet — click Refresh to load from SharePoint.</td></tr>";
    return;
  }

  var seen = {}, soNums = [];
  D.so.forEach(function(r){   var n = soKey(r); if (n && !seen[n]) { seen[n] = 1; soNums.push(n); } });
  D.tdd.forEach(function(r){  var n = soKey(r); if (n && !seen[n]) { seen[n] = 1; soNums.push(n); } });
  D.disp.forEach(function(r){ var n = soKey(r); if (n && !seen[n]) { seen[n] = 1; soNums.push(n); } });

  results = [];
  soNums.forEach(function(sn) {
    var soR   = D.so.find(function(r){   return soKey(r) === sn; }) || {};
    var tddR  = D.tdd.find(function(r){  return soKey(r) === sn; }) || {};
    var dispR = D.disp.find(function(r){ return soKey(r) === sn; }) || {};
    var hay   = (gv(soR) + " " + gv(tddR) + " " + gv(dispR)).toLowerCase();
    if (numQ && hay.indexOf(numQ) === -1) return;
    if (txtQ && hay.indexOf(txtQ) === -1) return;
    results.push({
      soR: soR, tddR: tddR, dispR: dispR, sn: sn,
      soFound:   Object.keys(soR).length   > 0,
      tddFound:  Object.keys(tddR).length  > 0,
      dispFound: Object.keys(dispR).length > 0
    });
  });

  if (sortC >= 0) {
    results.sort(function(a, b) {
      var src = window._sortSrc || 0;
      var h   = src === 0 ? D.soH : (src === 1 ? D.tddH : D.dispH);
      var obj = src === 0 ? "soR" : (src === 1 ? "tddR" : "dispR");
      var av  = (a[obj][h[sortC]] || "").toString();
      var bv  = (b[obj][h[sortC]] || "").toString();
      return av.localeCompare(bv, undefined, {numeric: true}) * sortDir;
    });
  }

  var resEl = document.getElementById("c-res");
  var cntEl = document.getElementById("tbl-count");
  if (resEl) resEl.textContent = results.length;
  if (cntEl) cntEl.textContent = results.length + " rows";
  renderTable(txtQ || numQ);
}

function hl(text, q) {
  if (!q || !text || text === "—") return text || "—";
  var re = new RegExp("(" + q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")", "gi");
  return text.replace(re, "<mark>$1</mark>");
}

function renderTable(q) {
  var tbody = document.getElementById("rtbl-body");
  if (!tbody) return;
  if (!results.length) {
    tbody.innerHTML = "<tr><td colspan='10' class='empty-msg'>No matching records found</td></tr>";
    return;
  }
  var showSO   = filterMode === "all" || filterMode === "so";
  var showTDD  = filterMode === "all" || filterMode === "tdd";
  var showDisp = filterMode === "all" || filterMode === "disp";

  tbody.innerHTML = results.map(function(row) {
    var cells = "";
    if (filterMode !== "stat") {
      if (showSO)   D.soH.forEach(function(h){   cells += "<td class='td-n'>" + hl(row.soR[h]   || "—", q) + "</td>"; });
      if (showTDD)  D.tddH.forEach(function(h){  cells += "<td class='td-n'>" + hl(row.tddR[h]  || "—", q) + "</td>"; });
      if (showDisp) D.dispH.forEach(function(h){ cells += "<td class='td-n'>" + hl(row.dispR[h] || "—", q) + "</td>"; });
    } else {
      cells = "<td class='td-n' style='font-weight:700;color:#1F4E79'>" + row.sn + "</td>";
    }
    cells += "<td style='text-align:center'><span class='" + (row.soFound   ? "bdn" : "bdp") + "'>" + (row.soFound   ? "✔" : "✘") + "</span></td>";
    cells += "<td style='text-align:center'><span class='" + (row.tddFound  ? "bdn" : "bdp") + "'>" + (row.tddFound  ? "✔" : "✘") + "</span></td>";
    cells += "<td style='text-align:center'><span class='" + (row.dispFound ? "bdn" : "bdp") + "'>" + (row.dispFound ? "✔" : "✘") + "</span></td>";
    return "<tr>" + cells + "</tr>";
  }).join("");
}

function sortBy(idx, src) {
  window._sortSrc = src;
  if (sortC === idx && window._lastSrc === src) sortDir *= -1;
  else { sortC = idx; sortDir = 1; }
  window._lastSrc = src;
  doSearch();
}

function setFilter(el, mode) {
  document.querySelectorAll(".chip").forEach(function(c){ c.classList.remove("on"); });
  el.classList.add("on");
  filterMode = mode;
  doSearch();
}

function clearSearch() {
  var n = document.getElementById("s-num");
  var t = document.getElementById("s-txt");
  if (n) n.value = "";
  if (t) t.value = "";
  doSearch();
}

function doExport() {
  if (!results.length) { alert("No data to export"); return; }
  var headers = D.soH.concat(D.tddH).concat(D.dispH).concat(["SO Released","TDD Set","Dispatched"]);
  var rows = results.map(function(r) {
    var vals = [];
    D.soH.forEach(function(h){   vals.push('"' + (r.soR[h]   || "").replace(/"/g, "") + '"'); });
    D.tddH.forEach(function(h){  vals.push('"' + (r.tddR[h]  || "").replace(/"/g, "") + '"'); });
    D.dispH.forEach(function(h){ vals.push('"' + (r.dispR[h] || "").replace(/"/g, "") + '"'); });
    vals.push(r.soFound ? "YES" : "NO", r.tddFound ? "YES" : "NO", r.dispFound ? "YES" : "NO");
    return vals.join(",");
  });
  var csv  = [headers.map(function(h){ return '"' + h + '"'; }).join(",")].concat(rows).join("\n");
  var blob = new Blob([csv], {type: "text/csv"});
  var a    = document.createElement("a");
  a.href   = URL.createObjectURL(blob);
  a.download = "Holostik_Report_" + new Date().toISOString().slice(0, 10) + ".csv";
  a.click();
}

function clearCacheAndReload() {
  try { localStorage.removeItem(CACHE_KEY); } catch(e) {}
  loadFromSharePoint(true);
}

/* ── Auto-load when search tab opens ──────────────────────
   Uses cache first (instant), refreshes in background if stale */
var _origST = window.switchTab;
window.switchTab = function(el, tab) {
  var ts = document.getElementById("tab-search");
  if (tab === "search") {
    document.querySelectorAll(".tab").forEach(function(t){ t.classList.remove("active"); });
    el.classList.add("active");
    document.querySelector(".card").style.display = "none";
    if (ts) ts.style.display = "block";
    if (!D.so.length && !D.tdd.length && !D.disp.length) {
      loadFromSharePoint(false);  // uses cache if available
    }
  } else {
    if (ts) ts.style.display = "none";
    document.querySelector(".card").style.display = "block";
    if (_origST) _origST(el, tab);
  }
};

window.loadPastedData = function() { manualRefresh(); };
window.clearAllData   = function() { clearCacheAndReload(); };
