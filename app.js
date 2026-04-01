/* ========================================================
 * マイ銘柄ソートリスト — Main Application (app.js)
 * ======================================================== */

/* ===== Firebase Config (REPLACE with your own) ===== */
var firebaseConfig = {
  apiKey: "AIzaSyCXpU-IX55WPJCVlhEhbcOZQzPmVKZt9PU",
  authDomain: "my-stock-list-7ae4d.firebaseapp.com",
  projectId: "my-stock-list-7ae4d",
  storageBucket: "my-stock-list-7ae4d.firebasestorage.app",
  messagingSenderId: "1072316134172",
  appId: "1:1072316134172:web:27abd8b3b8906afa2480d9"
};

firebase.initializeApp(firebaseConfig);
var auth = firebase.auth();
var db = firebase.firestore();
var currentUser = null;
var currentTab = "dashboard";
var editMode = false;
var dashEditMode = false;
var tabData = {};
var settings = { tabOrder: [], tabNames: {} };
var stockMaster = Object.assign({}, STOCK_MASTER);
var autoSaveTimer = null;
var presetData = null;
var MAX_STOCKS = { normal: 500, range: 1000 };

function isRangeTab(n) { return n >= 21 && n <= 30; }
function getMaxStocks(n) { return isRangeTab(n) ? MAX_STOCKS.range : MAX_STOCKS.normal; }
function lookupName(code) { return stockMaster[code] || ""; }

/* ===== Auth ===== */
document.getElementById("loginBtn").addEventListener("click", function() {
  auth.signInWithPopup(new firebase.auth.GoogleAuthProvider()).catch(function(e) {
    alert("ログインエラー: " + e.message);
  });
});
document.getElementById("logoutBtn").addEventListener("click", function() {
  auth.signOut();
});

auth.onAuthStateChanged(function(user) {
  if (user) {
    currentUser = user;
    document.getElementById("loginScreen").style.display = "none";
    document.getElementById("app").style.display = "flex";
    document.getElementById("userName").textContent = user.displayName || user.email;
    loadUserData();
  } else {
    currentUser = null;
    document.getElementById("loginScreen").style.display = "flex";
    document.getElementById("app").style.display = "none";
  }
});

/* ===== Data Loading ===== */
function loadUserData() {
  var userRef = db.collection("users").doc(currentUser.uid);
  userRef.get().then(function(doc) {
    if (doc.exists) {
      var d = doc.data();
      settings = d.settings || { tabOrder: DEFAULT_TABS.order.slice(), tabNames: Object.assign({}, DEFAULT_TABS.names) };
      if (d.customMaster) Object.assign(stockMaster, d.customMaster);
    } else {
      settings = { tabOrder: DEFAULT_TABS.order.slice(), tabNames: Object.assign({}, DEFAULT_TABS.names) };
      initializePresetTabs(userRef);
    }
    renderTabBar();
    showTab("dashboard");
  }).catch(function(e) {
    console.error("Load error:", e);
    settings = { tabOrder: DEFAULT_TABS.order.slice(), tabNames: Object.assign({}, DEFAULT_TABS.names) };
    renderTabBar();
    showTab("dashboard");
  });
}

function initializePresetTabs(userRef) {
  presetData = generatePresetTabs();
  var batch = db.batch();
  batch.set(userRef, { settings: settings });
  for (var t = 18; t <= 30; t++) {
    if (presetData[t]) {
      batch.set(userRef.collection("tabs").doc(String(t)), { stocks: presetData[t] });
    }
  }
  batch.commit().catch(function(e) { console.error("Init error:", e); });
}

function loadTabData(tabNum, callback) {
  if (tabData[tabNum]) { callback(tabData[tabNum]); return; }
  var userRef = db.collection("users").doc(currentUser.uid);
  userRef.collection("tabs").doc(String(tabNum)).get().then(function(doc) {
    tabData[tabNum] = doc.exists ? (doc.data().stocks || []) : [];
    callback(tabData[tabNum]);
  }).catch(function() {
    tabData[tabNum] = [];
    callback([]);
  });
}

function saveSettings() {
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(function() {
    db.collection("users").doc(currentUser.uid).set({ settings: settings }, { merge: true })
      .catch(function(e) { console.error("Save settings error:", e); });
  }, 500);
}

function saveTabData(tabNum) {
  clearTimeout(autoSaveTimer);
  autoSaveTimer = setTimeout(function() {
    db.collection("users").doc(currentUser.uid).collection("tabs").doc(String(tabNum))
      .set({ stocks: tabData[tabNum] || [] })
      .catch(function(e) { console.error("Save tab error:", e); });
  }, 500);
}

/* ===== Tab Bar ===== */
function renderTabBar() {
  var bar = document.getElementById("tabBar");
  bar.innerHTML = '<button class="tab' + (currentTab === "dashboard" ? " active" : "") + '" data-tab="dashboard">📋 DB</button>';
  settings.tabOrder.forEach(function(n) {
    var name = settings.tabNames[n] || "";
    var label = name ? n + "." + (name.length > 6 ? name.substring(0,6) + ".." : name) : String(n);
    bar.innerHTML += '<button class="tab' + (currentTab === String(n) ? " active" : "") + '" data-tab="' + n + '">' + label + '</button>';
  });
  bar.querySelectorAll(".tab").forEach(function(btn) {
    btn.addEventListener("click", function() { showTab(this.dataset.tab); });
  });
}

/* ===== Show Tab ===== */
function showTab(tab) {
  currentTab = tab;
  editMode = false;
  document.getElementById("tabEditToggle").classList.remove("active");
  document.getElementById("tabEditToggle").textContent = "✏️ 編集モード";
  renderTabBar();
  if (tab === "dashboard") {
    document.getElementById("dashboardView").style.display = "";
    document.getElementById("tabPageView").style.display = "none";
    renderDashboard();
  } else {
    document.getElementById("dashboardView").style.display = "none";
    document.getElementById("tabPageView").style.display = "";
    var n = parseInt(tab);
    document.getElementById("tabNameInput").value = settings.tabNames[n] || "";
    document.getElementById("tabNameInput").readOnly = true;
    document.getElementById("bulkArea").style.display = "none";
    loadTabData(n, function() { renderStockGrid(n); });
  }
}

/* ===== Dashboard ===== */
function renderDashboard() {
  var grid = document.getElementById("dashboardGrid");
  grid.innerHTML = "";
  settings.tabOrder.forEach(function(n, i) {
    var name = settings.tabNames[n] || "";
    var div = document.createElement("div");
    div.className = "dash-item";
    div.dataset.tabNum = n;
    div.draggable = dashEditMode;
    if (dashEditMode) {
      div.innerHTML = '<span class="dash-num">' + n + '</span>' +
        '<input class="dash-name-input" value="' + (name||"") + '" placeholder="タブ名..." data-n="' + n + '">' +
        '<div class="dash-move-btns"><button data-dir="up" data-n="' + n + '">▲</button><button data-dir="down" data-n="' + n + '">▼</button></div>';
    } else {
      var count = tabData[n] ? tabData[n].filter(function(s){return s.type!=="section"&&s.code;}).length : "—";
      div.innerHTML = '<span class="dash-num">' + n + '</span><span class="dash-name">' + (name || "(未設定)") + '</span><span class="dash-count">' + count + '</span>';
      div.addEventListener("click", function() { showTab(String(n)); });
    }
    grid.appendChild(div);
  });
  if (dashEditMode) {
    grid.querySelectorAll(".dash-name-input").forEach(function(inp) {
      inp.addEventListener("change", function() {
        settings.tabNames[parseInt(this.dataset.n)] = this.value;
        saveSettings();
        renderTabBar();
      });
    });
    grid.querySelectorAll(".dash-move-btns button").forEach(function(btn) {
      btn.addEventListener("click", function(e) {
        e.stopPropagation();
        var n = parseInt(this.dataset.n);
        var dir = this.dataset.dir;
        var idx = settings.tabOrder.indexOf(n);
        if (dir === "up" && idx > 0) {
          settings.tabOrder.splice(idx, 1);
          settings.tabOrder.splice(idx - 1, 0, n);
        } else if (dir === "down" && idx < settings.tabOrder.length - 1) {
          settings.tabOrder.splice(idx, 1);
          settings.tabOrder.splice(idx + 1, 0, n);
        }
        saveSettings();
        renderTabBar();
        renderDashboard();
      });
    });
    initDashDragDrop(grid);
  }
}

function initDashDragDrop(grid) {
  var items = grid.querySelectorAll(".dash-item");
  var dragItem = null;
  items.forEach(function(item) {
    item.addEventListener("dragstart", function(e) {
      dragItem = this;
      this.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
    });
    item.addEventListener("dragend", function() {
      this.classList.remove("dragging");
      dragItem = null;
      items.forEach(function(i) { i.classList.remove("drag-over"); });
    });
    item.addEventListener("dragover", function(e) {
      e.preventDefault();
      if (dragItem && dragItem !== this) this.classList.add("drag-over");
    });
    item.addEventListener("dragleave", function() { this.classList.remove("drag-over"); });
    item.addEventListener("drop", function(e) {
      e.preventDefault();
      if (!dragItem || dragItem === this) return;
      var fromN = parseInt(dragItem.dataset.tabNum);
      var toN = parseInt(this.dataset.tabNum);
      var fromIdx = settings.tabOrder.indexOf(fromN);
      var toIdx = settings.tabOrder.indexOf(toN);
      settings.tabOrder.splice(fromIdx, 1);
      settings.tabOrder.splice(toIdx, 0, fromN);
      saveSettings();
      renderTabBar();
      renderDashboard();
    });
  });
}

document.getElementById("dashEditToggle").addEventListener("click", function() {
  dashEditMode = !dashEditMode;
  this.classList.toggle("active", dashEditMode);
  this.textContent = dashEditMode ? "✅ 完了" : "✏️ 編集";
  renderDashboard();
});

/* ===== Tab Page Edit Mode ===== */
document.getElementById("tabEditToggle").addEventListener("click", function() {
  editMode = !editMode;
  this.classList.toggle("active", editMode);
  this.textContent = editMode ? "👁️ 表示モード" : "✏️ 編集モード";
  var n = parseInt(currentTab);
  document.getElementById("tabNameInput").readOnly = !editMode;
  document.getElementById("bulkArea").style.display = editMode ? "" : "none";
  renderStockGrid(n);
});

document.getElementById("tabNameInput").addEventListener("change", function() {
  var n = parseInt(currentTab);
  settings.tabNames[n] = this.value;
  saveSettings();
  renderTabBar();
});

/* ===== Stock Grid Rendering ===== */
function renderStockGrid(tabNum) {
  var grid = document.getElementById("stockGrid");
  var stocks = tabData[tabNum] || [];
  var maxRows = isRangeTab(tabNum) ? 200 : 100;
  grid.innerHTML = "";

  // Split into 5 columns
  var cols = [[],[],[],[],[]];
  var colIdx = 0;
  var rowInCol = 0;
  var stockNum = 0;
  stocks.forEach(function(item) {
    if (item.type === "section") {
      cols[colIdx].push({ type:"section", label: item.label });
    } else {
      stockNum++;
      cols[colIdx].push({ type:"stock", num: stockNum, code: item.code, name: item.name || lookupName(item.code), idx: cols[colIdx].length });
    }
    rowInCol++;
    if (rowInCol >= maxRows) {
      colIdx = Math.min(colIdx + 1, 4);
      rowInCol = 0;
    }
  });

  for (var c = 0; c < 5; c++) {
    var colDiv = document.createElement("div");
    colDiv.className = "stock-col";
    cols[c].forEach(function(item, ri) {
      var row = document.createElement("div");
      row.className = "stock-row" + (item.type === "section" ? " section-row" : "");
      if (item.type === "section") {
        row.innerHTML = '<span class="section-label">--- ' + item.label + ' ---</span>';
      } else {
        row.innerHTML = '<span class="row-num">' + item.num + '</span>' +
          '<span class="row-code">' + item.code + '</span>' +
          '<span class="row-name">' + (item.name || "") + '</span>';
        if (editMode) {
          row.innerHTML += '<span class="row-actions">' +
            '<button class="del-btn" title="削除">✕</button></span>';
          row.draggable = true;
        }
      }
      colDiv.appendChild(row);
    });
    grid.appendChild(colDiv);
  }

  // Stock count
  var totalStocks = stocks.filter(function(s) { return s.type !== "section" && s.code; }).length;
  document.getElementById("stockCount").textContent = totalStocks + " / " + getMaxStocks(tabNum) + " 銘柄";

  if (editMode) initStockEditEvents(tabNum);
}

function initStockEditEvents(tabNum) {
  document.querySelectorAll(".stock-row .del-btn").forEach(function(btn, i) {
    btn.addEventListener("click", function(e) {
      e.stopPropagation();
      // Find actual index in flat array
      var stocks = tabData[tabNum];
      var allRows = document.querySelectorAll(".stock-row");
      var rowIdx = Array.from(allRows).indexOf(this.closest(".stock-row"));
      if (rowIdx >= 0 && rowIdx < stocks.length) {
        stocks.splice(rowIdx, 1);
        saveTabData(tabNum);
        renderStockGrid(tabNum);
      }
    });
  });
}

/* ===== Bulk Input ===== */
document.getElementById("bulkAddBtn").addEventListener("click", function() {
  var input = document.getElementById("bulkInput").value.trim();
  if (!input) return;
  var n = parseInt(currentTab);
  var stocks = tabData[n] || [];
  var codes = input.split(/[,\s\u3000]+/).filter(function(c) { return c.length >= 3; });
  var existingCodes = new Set(stocks.filter(function(s){return s.code;}).map(function(s){return s.code;}));
  var added = 0;
  var max = getMaxStocks(n);
  var currentCount = stocks.filter(function(s){return s.type!=="section"&&s.code;}).length;

  codes.forEach(function(code) {
    code = code.trim().toUpperCase();
    if (code && !existingCodes.has(code) && currentCount + added < max) {
      stocks.push({ code: code, name: lookupName(code) });
      existingCodes.add(code);
      added++;
    }
  });

  tabData[n] = stocks;
  saveTabData(n);
  renderStockGrid(n);
  document.getElementById("bulkInput").value = "";
});

document.getElementById("bulkExportBtn").addEventListener("click", function() {
  var n = parseInt(currentTab);
  var stocks = tabData[n] || [];
  var codes = stocks.filter(function(s){return s.code;}).map(function(s){return s.code;});
  document.getElementById("bulkOutput").value = codes.join(",");
});

document.getElementById("bulkCopyBtn").addEventListener("click", function() {
  var output = document.getElementById("bulkOutput");
  output.select();
  document.execCommand("copy");
});

document.getElementById("addSectionBtn").addEventListener("click", function() {
  var label = document.getElementById("sectionInput").value.trim();
  if (!label) return;
  var n = parseInt(currentTab);
  var stocks = tabData[n] || [];
  stocks.push({ type: "section", label: label });
  tabData[n] = stocks;
  saveTabData(n);
  renderStockGrid(n);
  document.getElementById("sectionInput").value = "";
});

document.getElementById("resetBtn").addEventListener("click", function() {
  if (!confirm("このタブの全銘柄を削除しますか？")) return;
  var n = parseInt(currentTab);
  tabData[n] = [];
  saveTabData(n);
  renderStockGrid(n);
});

/* ===== Admin Modal ===== */
document.getElementById("adminBtn").addEventListener("click", function() {
  document.getElementById("masterCount").textContent = Object.keys(stockMaster).length;
  document.getElementById("adminModal").classList.add("show");
});

document.getElementById("masterUploadBtn").addEventListener("click", function() {
  var file = document.getElementById("masterFileInput").files[0];
  if (!file) { setStatus("masterStatus","ファイルを選択してください","error"); return; }
  var reader = new FileReader();
  reader.onload = function(e) {
    var text = e.target.result;
    var lines = text.split(/\r?\n/);
    var newMaster = {};
    var count = 0;
    lines.forEach(function(line) {
      var parts = line.split(",");
      if (parts.length >= 2) {
        var code = parts[0].trim().replace(/"/g,"");
        var name = parts[1].trim().replace(/"/g,"");
        if (code && name && /^[0-9A-Za-z]{3,5}$/.test(code)) {
          newMaster[code] = name;
          count++;
        }
      }
    });
    if (count > 0) {
      stockMaster = newMaster;
      db.collection("users").doc(currentUser.uid).set({ customMaster: stockMaster }, { merge: true });
      document.getElementById("masterCount").textContent = count;
      setStatus("masterStatus", count + "件のマスタを更新しました","success");
    } else {
      setStatus("masterStatus","有効なデータが見つかりません","error");
    }
  };
  reader.readAsText(file, "Shift_JIS");
});

document.getElementById("regenRangeBtn").addEventListener("click", function() {
  if (!confirm("タブ21-30を再生成しますか？既存データは上書きされます。")) return;
  var presets = generatePresetTabs();
  var batch = db.batch();
  var userRef = db.collection("users").doc(currentUser.uid);
  for (var t = 21; t <= 30; t++) {
    tabData[t] = presets[t] || [];
    batch.set(userRef.collection("tabs").doc(String(t)), { stocks: tabData[t] });
  }
  batch.commit().then(function() {
    setStatus("rangeStatus","タブ21-30を再生成しました","success");
    if (parseInt(currentTab) >= 21) renderStockGrid(parseInt(currentTab));
  }).catch(function(e) {
    setStatus("rangeStatus","エラー: " + e.message,"error");
  });
});

function setStatus(id, msg, type) {
  var el = document.getElementById(id);
  el.textContent = msg;
  el.className = "admin-status " + type;
  setTimeout(function() { el.textContent = ""; }, 5000);
}

/* ===== Help Modal ===== */
document.getElementById("helpBtn").addEventListener("click", function() {
  document.getElementById("helpModal").classList.add("show");
});

/* ===== Modal Close on Overlay Click ===== */
document.querySelectorAll(".modal-overlay").forEach(function(overlay) {
  overlay.addEventListener("click", function(e) {
    if (e.target === this) this.classList.remove("show");
  });
});

/* ===== Index Upload ===== */
document.getElementById("indexUploadBtn").addEventListener("click", function() {
  setStatus("indexStatus","処理中...","success");
  var promises = [];
  document.querySelectorAll(".index-file").forEach(function(input) {
    if (input.files[0]) {
      promises.push(new Promise(function(resolve) {
        var reader = new FileReader();
        reader.onload = function(e) {
          var codes = e.target.result.split(/[\r\n,\s]+/).filter(function(c) { return /^[0-9A-Za-z]{3,5}$/.test(c.trim()); });
          resolve({ index: input.dataset.index, codes: codes });
        };
        reader.readAsText(input.files[0]);
      }));
    }
  });
  if (promises.length === 0) { setStatus("indexStatus","ファイルを選択してください","error"); return; }
  Promise.all(promises).then(function(results) {
    var indexData = {};
    results.forEach(function(r) { indexData[r.index] = r.codes; });
    // Rebuild tabs 18-20
    rebuildIndexTabs(indexData);
    setStatus("indexStatus","指数銘柄を更新しました","success");
  });
});

function rebuildIndexTabs(indexData) {
  var userRef = db.collection("users").doc(currentUser.uid);
  var batch = db.batch();

  if (indexData.jpx150 || indexData.growth250) {
    var tab18 = [];
    if (indexData.jpx150) {
      tab18.push({ type:"section", label:"JPX日経150" });
      indexData.jpx150.forEach(function(c) { tab18.push({ code:c, name:lookupName(c) }); });
    }
    if (indexData.growth250) {
      tab18.push({ type:"section", label:"グロース250" });
      indexData.growth250.forEach(function(c) { tab18.push({ code:c, name:lookupName(c) }); });
    }
    tabData[18] = tab18;
    batch.set(userRef.collection("tabs").doc("18"), { stocks: tab18 });
  }

  if (indexData.jpx400) {
    var tab19 = [{ type:"section", label:"JPX日経400" }];
    indexData.jpx400.forEach(function(c) { tab19.push({ code:c, name:lookupName(c) }); });
    tabData[19] = tab19;
    batch.set(userRef.collection("tabs").doc("19"), { stocks: tab19 });
  }

  if (indexData.nikkei225 || indexData.topix100) {
    var tab20 = [];
    if (indexData.nikkei225) {
      tab20.push({ type:"section", label:"日経225" });
      indexData.nikkei225.forEach(function(c) { tab20.push({ code:c, name:lookupName(c) }); });
    }
    if (indexData.topix100) {
      tab20.push({ type:"section", label:"TOPIX100" });
      var nkSet = new Set(indexData.nikkei225 || []);
      indexData.topix100.forEach(function(c) {
        if (!nkSet.has(c)) tab20.push({ code:c, name:lookupName(c) });
      });
    }
    tabData[20] = tab20;
    batch.set(userRef.collection("tabs").doc("20"), { stocks: tab20 });
  }

  batch.commit();
}
