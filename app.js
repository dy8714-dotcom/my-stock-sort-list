/**
 * app.js — マイ銘柄ソートリスト メインアプリケーション
 * 全改修項目 A〜M 対応版 (v2 — Firestore修正＋永続化対応)
 */

// ============================================================
// Firebase 設定
// ============================================================
const firebaseConfig = {
  apiKey: "AIzaSyCXpU-IX55WPJCVlhEhbcOZQzPmVKZt9PU",
  authDomain: "my-stock-list-7ae4d.firebaseapp.com",
  projectId: "my-stock-list-7ae4d",
  storageBucket: "my-stock-list-7ae4d.firebasestorage.app",
  messagingSenderId: "1072316134172",
  appId: "1:1072316134172:web:27abd8b3b8906afa2480d9"
};

firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();
const provider = new firebase.auth.GoogleAuthProvider();

// ============================================================
// アプリ状態
// ============================================================
const APP = {
  uid: null,
  currentTab: 1,
  tabs: {},
  tabNames: {},
  tabOrder: [],
  customMaster: null,
  checkedItems: new Set(),
  dragState: null,
  saving: false,
  saveQueue: {},
  COL_COUNT: 6,
  firestoreAvailable: true
};

// ============================================================
// ユーティリティ
// ============================================================
function showToast(msg, type = 'info') {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = `toast ${type}`;
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(() => t.remove(), 3500);
}

function getMaster() { return APP.customMaster || STOCK_MASTER; }

function resolveName(code) {
  const m = getMaster();
  return m[code] || STOCK_MASTER[code] || code + '（不明）';
}

function getStockCodes(tabData) {
  return (tabData || []).filter(r => r.type !== 'section').map(r => r.code);
}

function findDuplicates(tabData) {
  const codes = getStockCodes(tabData);
  const seen = new Set(), dupes = new Set();
  for (const c of codes) { if (seen.has(c)) dupes.add(c); seen.add(c); }
  return dupes;
}

// ============================================================
// 列分割ロジック
// ============================================================
function splitToColumns(tabData, colCount) {
  if (!tabData || tabData.length === 0) return Array.from({ length: colCount }, () => []);
  const cols = Array.from({ length: colCount }, () => []);
  const hasCol = tabData.some(r => r._col !== undefined && r._col !== null);
  if (hasCol) {
    for (const item of tabData) {
      const c = Math.min(Math.max(item._col || 0, 0), colCount - 1);
      cols[c].push(item);
    }
  } else {
    const perCol = Math.ceil(tabData.length / colCount);
    for (let i = 0; i < tabData.length; i++) {
      cols[Math.min(Math.floor(i / perCol), colCount - 1)].push(tabData[i]);
    }
  }
  return cols;
}

function flattenColumns(cols) {
  const result = [];
  for (let c = 0; c < cols.length; c++) {
    for (const item of cols[c]) result.push({ ...item, _col: c });
  }
  return result;
}

// ============================================================
// localStorage 永続化（Firestoreフォールバック）
// ============================================================
const LS_PREFIX = 'mssl_';

function lsSave(key, data) {
  try { localStorage.setItem(LS_PREFIX + key, JSON.stringify(data)); } catch (e) { console.warn('LS save error:', e); }
}

function lsLoad(key) {
  try {
    const s = localStorage.getItem(LS_PREFIX + key);
    return s ? JSON.parse(s) : null;
  } catch (e) { return null; }
}

function lsSaveTab(uid, tabId, stocks) {
  lsSave(`${uid}_tab_${tabId}`, stocks);
}

function lsLoadTab(uid, tabId) {
  return lsLoad(`${uid}_tab_${tabId}`);
}

function lsSaveSettings(uid, settings) {
  lsSave(`${uid}_settings`, settings);
}

function lsLoadSettings(uid) {
  return lsLoad(`${uid}_settings`);
}

function lsSaveMaster(uid, master) {
  lsSave(`${uid}_master`, master);
}

function lsLoadMaster(uid) {
  return lsLoad(`${uid}_master`);
}

// ============================================================
// Firebase Auth
// ============================================================
document.getElementById('btn-google-login').addEventListener('click', () => {
  auth.signInWithPopup(provider).catch(err => {
    document.getElementById('login-status').textContent = 'ログインエラー: ' + err.message;
  });
});

document.getElementById('btn-logout').addEventListener('click', () => auth.signOut());

auth.onAuthStateChanged(async (user) => {
  if (user) {
    APP.uid = user.uid;
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-container').classList.remove('hidden');
    document.getElementById('user-name').textContent = user.displayName || user.email;
    if (user.photoURL) document.getElementById('user-avatar').src = user.photoURL;
    await loadUserData();
    renderTabs();
    switchTab(APP.currentTab);
  } else {
    APP.uid = null;
    document.getElementById('login-screen').style.display = 'flex';
    document.getElementById('app-container').classList.add('hidden');
  }
});

// ============================================================
// Firestore データ読み書き
// ★ セキュリティルール対応: users/{uid} ドキュメント直接書き込み
// ★ users/{uid}/tabs/{tabId} サブコレクション
// ============================================================
async function loadUserData() {
  if (!APP.uid) return;

  // 1) まずlocalStorageから読み込み（即座にデータ利用可能に）
  const lsSettings = lsLoadSettings(APP.uid);
  const lsMaster = lsLoadMaster(APP.uid);

  if (lsSettings) {
    APP.tabOrder = lsSettings.tabOrder || getDefaultTabOrder();
    APP.tabNames = lsSettings.tabNames || getDefaultTabNames();
  } else {
    APP.tabOrder = getDefaultTabOrder();
    APP.tabNames = getDefaultTabNames();
  }

  if (lsMaster) APP.customMaster = lsMaster;

  // localStorageからタブデータ読み込み
  let hasLocalData = false;
  for (let i = 1; i <= 30; i++) {
    const tabData = lsLoadTab(APP.uid, i);
    if (tabData && tabData.length > 0) {
      APP.tabs[i] = tabData;
      hasLocalData = true;
    }
  }

  // 2) Firestoreからも読み込み試行
  try {
    // settings を users/{uid} ドキュメントのフィールドとして読み込み
    const userDoc = await db.collection('users').doc(APP.uid).get();
    if (userDoc.exists) {
      const data = userDoc.data();
      if (data.settings) {
        APP.tabOrder = data.settings.tabOrder || APP.tabOrder;
        APP.tabNames = data.settings.tabNames || APP.tabNames;
      }
      if (data.customMaster) {
        APP.customMaster = data.customMaster;
        lsSaveMaster(APP.uid, data.customMaster);
      }
    }

    // tabs サブコレクション
    const tabsSnap = await db.collection('users').doc(APP.uid).collection('tabs').get();
    if (!tabsSnap.empty) {
      tabsSnap.forEach(doc => {
        const tabData = doc.data().stocks;
        if (tabData && tabData.length > 0) {
          APP.tabs[doc.id] = tabData;
          lsSaveTab(APP.uid, doc.id, tabData);
          hasLocalData = true;
        }
      });
    }

    APP.firestoreAvailable = true;
  } catch (err) {
    console.warn('Firestore読み込みエラー（localStorageで継続）:', err.message);
    APP.firestoreAvailable = false;
  }

  // 3) データがなければ初期化
  if (!hasLocalData || Object.keys(APP.tabs).length === 0) {
    console.log('初期データ生成中...');
    initializeDefaultTabs();
    saveAllDataLocal();
    // Firestoreにも保存試行
    saveAllDataFirestore();
  }

  updateMasterCountDisplay();
}

function getDefaultTabOrder() {
  return Array.from({ length: 30 }, (_, i) => i + 1);
}

function getDefaultTabNames() {
  const n = {};
  for (let i = 1; i <= 17; i++) n[i] = '';
  n[18] = 'JPX150＆G250';
  n[19] = 'JPX400';
  n[20] = '日経225＆TOPIX100';
  n[21] = '1000番台'; n[22] = '2000番台'; n[23] = '3000番台';
  n[24] = '4000番台'; n[25] = '5000番台'; n[26] = '6000番台';
  n[27] = '7000番台'; n[28] = '8000番台'; n[29] = '9000番台';
  n[30] = '3桁＋Aコード';
  return n;
}

// ============================================================
// プリセット初期データ生成
// ============================================================
function initializeDefaultTabs() {
  APP.tabNames = getDefaultTabNames();
  APP.tabOrder = getDefaultTabOrder();
  const master = getMaster();

  // タブ1-17: 空
  for (let i = 1; i <= 17; i++) APP.tabs[i] = [];

  // タブ18: JPX150 + グロース250
  APP.tabs[18] = buildIndexTabData(
    [{ name: 'JPX150', codes: INDEX_JPX150 || INDEX_NIKKEI225.slice(0, 150) }]
  , master);

  // タブ19: JPX400
  APP.tabs[19] = buildIndexTabData(
    [{ name: 'JPX400', codes: INDEX_JPX400 || INDEX_NIKKEI225 }]
  , master);

  // タブ20: 日経225 + TOPIX100（重複除外）
  const n225set = new Set(INDEX_NIKKEI225);
  const topixOnly = INDEX_TOPIX100.filter(c => !n225set.has(c));
  APP.tabs[20] = buildIndexTabData([
    { name: '日経225', codes: INDEX_NIKKEI225 },
    { name: 'TOPIX100（日経225除く）', codes: topixOnly }
  ], master);

  // タブ21-29: 各番台（昇順、Aコード除外）
  const bandai = getStocksByBandai(master);
  for (let b = 1; b <= 9; b++) {
    const items = bandai[b];
    APP.tabs[20 + b] = assignColumns(items.map(s => ({ code: s.code, name: s.name })));
  }

  // タブ30: 3桁+Aコード
  APP.tabs[30] = assignColumns(
    bandai['special'].map(s => ({ code: s.code, name: s.name }))
  );
}

function buildIndexTabData(sections, master) {
  const result = [];
  for (const sec of sections) {
    result.push({ type: 'section', label: sec.name, _col: 0 });
    const items = sec.codes.map(code => ({
      code, name: master[code] || STOCK_MASTER[code] || code
    }));
    const assigned = assignColumns(items);
    result.push(...assigned);
  }
  return result;
}

function assignColumns(items) {
  if (items.length === 0) return [];
  const perCol = Math.ceil(items.length / APP.COL_COUNT);
  return items.map((item, idx) => ({
    ...item,
    _col: Math.min(Math.floor(idx / perCol), APP.COL_COUNT - 1)
  }));
}

// ============================================================
// データ保存
// ============================================================
function saveAllDataLocal() {
  if (!APP.uid) return;
  lsSaveSettings(APP.uid, { tabOrder: APP.tabOrder, tabNames: APP.tabNames });
  for (const tabId of Object.keys(APP.tabs)) {
    lsSaveTab(APP.uid, tabId, APP.tabs[tabId]);
  }
}

async function saveAllDataFirestore() {
  if (!APP.uid || !APP.firestoreAvailable) return;
  try {
    // users/{uid} ドキュメントに settings を保存
    await db.collection('users').doc(APP.uid).set({
      settings: { tabOrder: APP.tabOrder, tabNames: APP.tabNames }
    }, { merge: true });

    // tabs サブコレクション
    const batch = db.batch();
    for (const tabId of Object.keys(APP.tabs)) {
      const ref = db.collection('users').doc(APP.uid).collection('tabs').doc(String(tabId));
      batch.set(ref, { stocks: APP.tabs[tabId] || [] });
    }
    await batch.commit();
  } catch (err) {
    console.warn('Firestore保存エラー:', err.message);
    APP.firestoreAvailable = false;
  }
}

let saveDebounceTimer = null;
function saveTabData(tabId) {
  if (!APP.uid) return;
  // localStorage即時保存
  lsSaveTab(APP.uid, tabId, APP.tabs[tabId] || []);
  lsSaveSettings(APP.uid, { tabOrder: APP.tabOrder, tabNames: APP.tabNames });

  // Firestoreはデバウンス
  clearTimeout(saveDebounceTimer);
  saveDebounceTimer = setTimeout(async () => {
    if (!APP.firestoreAvailable) return;
    try {
      await db.collection('users').doc(APP.uid).collection('tabs').doc(String(tabId)).set({
        stocks: APP.tabs[tabId] || []
      });
    } catch (err) {
      console.warn('Firestore tab save error:', err.message);
    }
  }, 1000);
}

function saveSettings() {
  if (!APP.uid) return;
  lsSaveSettings(APP.uid, { tabOrder: APP.tabOrder, tabNames: APP.tabNames });
  if (!APP.firestoreAvailable) return;
  db.collection('users').doc(APP.uid).set({
    settings: { tabOrder: APP.tabOrder, tabNames: APP.tabNames }
  }, { merge: true }).catch(err => console.warn('Firestore settings save:', err.message));
}

async function saveCustomMaster(masterData) {
  if (!APP.uid) return;
  APP.customMaster = masterData;
  lsSaveMaster(APP.uid, masterData);
  if (!APP.firestoreAvailable) return;
  try {
    await db.collection('users').doc(APP.uid).set({ customMaster: masterData }, { merge: true });
  } catch (err) {
    console.warn('Firestore master save error:', err.message);
    // マスタが大きすぎる場合は分割
    try {
      const entries = Object.entries(masterData);
      if (entries.length > 3000) {
        const half = Math.ceil(entries.length / 2);
        const part1 = Object.fromEntries(entries.slice(0, half));
        const part2 = Object.fromEntries(entries.slice(half));
        await db.collection('users').doc(APP.uid).set({ customMaster: part1 }, { merge: true });
        await db.collection('users').doc(APP.uid).collection('settings').doc('masterPart2').set({ data: part2 });
      }
    } catch (e2) { console.warn('分割保存も失敗:', e2.message); }
  }
}

// ============================================================
// タブ描画
// ============================================================
function renderTabs() {
  const bar = document.getElementById('tab-bar');
  bar.innerHTML = '';
  const order = APP.tabOrder.length ? APP.tabOrder : getDefaultTabOrder();

  for (const tabId of order) {
    const div = document.createElement('div');
    div.className = 'tab-item' + (tabId === APP.currentTab ? ' active' : '');
    const name = APP.tabNames[tabId] || '';
    const count = getStockCodes(APP.tabs[tabId]).length;
    div.innerHTML = `<span>${name || 'タブ' + tabId}</span><span class="tab-count">(${count})</span>`;
    div.dataset.tabId = tabId;
    div.addEventListener('click', () => switchTab(tabId));
    div.addEventListener('dblclick', () => startTabRename(tabId, div));
    bar.appendChild(div);
  }
}

function switchTab(tabId) {
  APP.currentTab = tabId;
  APP.checkedItems.clear();
  document.querySelectorAll('.tab-item').forEach(el => {
    el.classList.toggle('active', parseInt(el.dataset.tabId) === tabId);
  });
  document.getElementById('tab-name-input').value = APP.tabNames[tabId] || '';
  renderGrid();
  updateStockCount();
}

function startTabRename(tabId, el) {
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'tab-name-input';
  input.value = APP.tabNames[tabId] || '';
  el.innerHTML = '';
  el.appendChild(input);
  input.focus();
  input.select();
  const finish = () => {
    APP.tabNames[tabId] = input.value;
    saveSettings();
    renderTabs();
  };
  input.addEventListener('blur', finish);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); finish(); }
    if (e.key === 'Escape') renderTabs();
  });
}

function updateStockCount() {
  const count = getStockCodes(APP.tabs[APP.currentTab]).length;
  document.getElementById('stock-count-display').textContent = `${count}銘柄`;
}

// ============================================================
// 6列グリッド描画
// ============================================================
function renderGrid() {
  const grid = document.getElementById('stock-grid');
  grid.innerHTML = '';
  const tabData = APP.tabs[APP.currentTab] || [];
  const cols = splitToColumns(tabData, APP.COL_COUNT);
  const duplicates = findDuplicates(tabData);

  for (let colIdx = 0; colIdx < APP.COL_COUNT; colIdx++) {
    const colDiv = document.createElement('div');
    colDiv.className = 'stock-column';
    colDiv.dataset.col = colIdx;

    // ヘッダ
    const header = document.createElement('div');
    header.className = 'column-header';
    header.innerHTML = `
      <span class="col-label">列${colIdx + 1}</span>
      <div class="select-all-wrap">
        <label style="font-size:10px;cursor:pointer;">全選択</label>
        <input type="checkbox" class="select-all-cb" data-col="${colIdx}">
      </div>`;
    colDiv.appendChild(header);

    // ボディ
    const body = document.createElement('div');
    body.className = 'column-body';
    body.dataset.col = colIdx;

    const colItems = cols[colIdx];
    let rowNum = 0;

    if (colItems.length === 0) {
      body.innerHTML = '<div class="empty-placeholder">ドロップまたは入力</div>';
    } else {
      for (let rowIdx = 0; rowIdx < colItems.length; rowIdx++) {
        const item = colItems[rowIdx];
        if (item.type === 'section') {
          body.appendChild(createSectionRow(item, colIdx, rowIdx));
        } else {
          rowNum++;
          body.appendChild(createStockRow(item, colIdx, rowIdx, rowNum, duplicates));
        }
      }
    }

    setupDropZone(body, colIdx);
    colDiv.appendChild(body);
    grid.appendChild(colDiv);
  }

  // 全選択
  document.querySelectorAll('.select-all-cb').forEach(cb => {
    cb.addEventListener('change', (e) => {
      const col = parseInt(e.target.dataset.col);
      const checked = e.target.checked;
      const body = document.querySelector(`.column-body[data-col="${col}"]`);
      body.querySelectorAll('.stock-checkbox').forEach(scb => {
        scb.checked = checked;
        const key = scb.dataset.key;
        if (checked) APP.checkedItems.add(key); else APP.checkedItems.delete(key);
        scb.closest('.stock-row').classList.toggle('selected', checked);
      });
    });
  });
}

function createSectionRow(item, colIdx, rowIdx) {
  const div = document.createElement('div');
  div.className = 'section-header';
  div.draggable = true;
  div.dataset.col = colIdx;
  div.dataset.row = rowIdx;
  div.dataset.type = 'section';
  div.innerHTML = `<span class="section-label">📁 ${item.label}</span>
    <div class="section-actions">
      <button class="btn btn-small" data-action="rename">✏</button>
      <button class="btn btn-small btn-danger" data-action="delete">✕</button>
    </div>`;

  div.querySelector('[data-action="rename"]').addEventListener('click', (e) => {
    e.stopPropagation();
    const newName = prompt('セクション名:', item.label);
    if (newName !== null) {
      item.label = newName;
      saveTabData(APP.currentTab);
      renderGrid();
    }
  });

  div.querySelector('[data-action="delete"]').addEventListener('click', (e) => {
    e.stopPropagation();
    if (!confirm(`セクション「${item.label}」を削除？`)) return;
    const tabData = APP.tabs[APP.currentTab];
    const cols = splitToColumns(tabData, APP.COL_COUNT);
    cols[colIdx].splice(rowIdx, 1);
    APP.tabs[APP.currentTab] = flattenColumns(cols);
    saveTabData(APP.currentTab);
    renderGrid();
  });

  setupDragSource(div, colIdx, rowIdx);
  return div;
}

function createStockRow(item, colIdx, rowIdx, rowNum, duplicates) {
  const div = document.createElement('div');
  div.className = 'stock-row';
  div.draggable = true;
  div.dataset.col = colIdx;
  div.dataset.row = rowIdx;
  div.dataset.code = item.code;
  div.dataset.type = 'stock';

  const key = `${APP.currentTab}-${colIdx}-${rowIdx}`;
  const isChecked = APP.checkedItems.has(key);
  if (isChecked) div.classList.add('selected');
  const isDup = duplicates.has(item.code);
  const name = resolveName(item.code);

  div.innerHTML = `
    <span class="row-num">${rowNum}</span>
    <span class="stock-code" data-code="${item.code}">${item.code}</span>
    <span class="stock-name${isDup ? ' duplicate' : ''}" title="${name}">${name}</span>
    <input type="checkbox" class="stock-checkbox" data-key="${key}" ${isChecked ? 'checked' : ''}>`;

  // コード直接編集
  div.querySelector('.stock-code').addEventListener('click', (e) => {
    e.stopPropagation();
    startCodeEdit(e.target, item);
  });

  // チェックボックス
  const cb = div.querySelector('.stock-checkbox');
  cb.addEventListener('change', (e) => {
    e.stopPropagation();
    if (cb.checked) { APP.checkedItems.add(key); div.classList.add('selected'); }
    else { APP.checkedItems.delete(key); div.classList.remove('selected'); }
  });

  setupDragSource(div, colIdx, rowIdx);
  return div;
}

// ============================================================
// コード直接編集（項目H）
// ============================================================
function startCodeEdit(codeEl, item) {
  if (codeEl.querySelector('input')) return;
  const original = item.code;
  const input = document.createElement('input');
  input.type = 'text';
  input.value = item.code;
  input.style.cssText = 'width:50px;font-family:var(--font-mono);font-size:12px;padding:1px 2px;border:none;background:var(--bg-input);color:var(--text-bright);outline:1px solid var(--accent);border-radius:2px;';
  codeEl.textContent = '';
  codeEl.appendChild(input);
  input.focus();
  input.select();

  const finish = (save) => {
    if (save && input.value.trim()) {
      const newCode = input.value.trim().toUpperCase();
      item.code = newCode;
      item.name = resolveName(newCode);
      saveTabData(APP.currentTab);
    }
    renderGrid();
  };

  input.addEventListener('blur', () => finish(true));
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); finish(true); }
    if (e.key === 'Escape') finish(false);
  });
}

// ============================================================
// ドラッグ＆ドロップ（項目D）
// ============================================================
function setupDragSource(el, colIdx, rowIdx) {
  el.addEventListener('dragstart', (e) => {
    APP.dragState = { colIdx, rowIdx };
    el.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', '');
  });
  el.addEventListener('dragend', () => {
    el.classList.remove('dragging');
    clearDragIndicators();
    APP.dragState = null;
  });
}

function clearDragIndicators() {
  document.querySelectorAll('.drag-target-above,.drag-target-below,.drag-over').forEach(
    el => el.classList.remove('drag-target-above', 'drag-target-below', 'drag-over'));
}

function setupDropZone(bodyEl, targetColIdx) {
  bodyEl.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    bodyEl.classList.add('drag-over');
    const rows = bodyEl.querySelectorAll('.stock-row, .section-header');
    rows.forEach(r => r.classList.remove('drag-target-above', 'drag-target-below'));
    const closest = getClosestRow(e.clientY, rows);
    if (closest.el) {
      closest.el.classList.add(closest.above ? 'drag-target-above' : 'drag-target-below');
    }
  });

  bodyEl.addEventListener('dragleave', (e) => {
    if (!bodyEl.contains(e.relatedTarget)) {
      bodyEl.classList.remove('drag-over');
      bodyEl.querySelectorAll('.drag-target-above,.drag-target-below').forEach(
        el => el.classList.remove('drag-target-above', 'drag-target-below'));
    }
  });

  bodyEl.addEventListener('drop', (e) => {
    e.preventDefault();
    clearDragIndicators();
    if (!APP.dragState) return;

    const { colIdx: srcCol, rowIdx: srcRow } = APP.dragState;
    const tabData = APP.tabs[APP.currentTab];
    const cols = splitToColumns(tabData, APP.COL_COUNT);
    const srcItem = cols[srcCol] && cols[srcCol][srcRow];
    if (!srcItem) return;

    const rows = bodyEl.querySelectorAll('.stock-row, .section-header');
    const closest = getClosestRow(e.clientY, rows);
    let targetRow = closest.el ? parseInt(closest.el.dataset.row) + (closest.above ? 0 : 1) : cols[targetColIdx].length;

    cols[srcCol].splice(srcRow, 1);
    if (srcCol === targetColIdx && srcRow < targetRow) targetRow--;
    cols[targetColIdx].splice(targetRow, 0, srcItem);

    APP.tabs[APP.currentTab] = flattenColumns(cols);
    saveTabData(APP.currentTab);
    renderGrid();
  });
}

function getClosestRow(y, rows) {
  let best = { el: null, distance: Infinity, above: true };
  rows.forEach(row => {
    const rect = row.getBoundingClientRect();
    const mid = rect.top + rect.height / 2;
    const dist = Math.abs(y - mid);
    if (dist < best.distance) best = { el: row, distance: dist, above: y < mid };
  });
  return best;
}

// ============================================================
// ツールバー操作
// ============================================================

// 銘柄追加
document.getElementById('btn-add-stock').addEventListener('click', addStock);
document.getElementById('add-stock-input').addEventListener('keydown', (e) => { if (e.key === 'Enter') addStock(); });

function addStock() {
  const input = document.getElementById('add-stock-input');
  const raw = input.value.trim();
  if (!raw) return;
  const codes = raw.split(/[\s,;、]+/).filter(c => c);
  const targetCol = parseInt(document.getElementById('input-column-select').value);
  const tabData = APP.tabs[APP.currentTab] || [];
  const cols = splitToColumns(tabData, APP.COL_COUNT);

  for (const code of codes) {
    const uc = code.toUpperCase();
    cols[targetCol].push({ code: uc, name: resolveName(uc), _col: targetCol });
  }

  APP.tabs[APP.currentTab] = flattenColumns(cols);
  saveTabData(APP.currentTab);
  renderGrid();
  renderTabs();
  updateStockCount();
  input.value = '';
  showToast(`${codes.length}銘柄を列${targetCol + 1}に追加`, 'success');
}

// セクション追加
document.getElementById('btn-add-section').addEventListener('click', () => {
  const input = document.getElementById('add-section-input');
  const name = input.value.trim();
  if (!name) return;
  const targetCol = parseInt(document.getElementById('input-column-select').value);
  const tabData = APP.tabs[APP.currentTab] || [];
  const cols = splitToColumns(tabData, APP.COL_COUNT);
  cols[targetCol].unshift({ type: 'section', label: name, _col: targetCol });
  APP.tabs[APP.currentTab] = flattenColumns(cols);
  saveTabData(APP.currentTab);
  renderGrid();
  input.value = '';
  showToast(`セクション「${name}」を追加`, 'success');
});

// タブ名変更
document.getElementById('btn-rename-tab').addEventListener('click', () => {
  APP.tabNames[APP.currentTab] = document.getElementById('tab-name-input').value;
  saveSettings();
  renderTabs();
  showToast('タブ名を変更しました', 'success');
});

// ============================================================
// チェックボックス操作（項目E）
// ============================================================
document.getElementById('btn-export-codes').addEventListener('click', () => {
  const codes = getCheckedCodes();
  if (codes.length === 0) { showToast('チェック銘柄なし', 'info'); return; }
  document.getElementById('export-textarea').value = codes.join('\n');
  document.getElementById('export-modal').classList.remove('hidden');
});

document.getElementById('btn-delete-checked').addEventListener('click', () => {
  const keys = [...APP.checkedItems];
  if (keys.length === 0) { showToast('チェック銘柄なし', 'info'); return; }
  if (!confirm(`${keys.length}件削除しますか？`)) return;

  const tabData = APP.tabs[APP.currentTab];
  const cols = splitToColumns(tabData, APP.COL_COUNT);
  const toDelete = {};
  for (const key of keys) {
    const [, colStr, rowStr] = key.split('-');
    if (!toDelete[colStr]) toDelete[colStr] = [];
    toDelete[colStr].push(parseInt(rowStr));
  }
  for (const col of Object.keys(toDelete)) {
    toDelete[col].sort((a, b) => b - a);
    for (const row of toDelete[col]) {
      if (cols[col] && cols[col][row] && cols[col][row].type !== 'section') {
        cols[col].splice(row, 1);
      }
    }
  }
  APP.tabs[APP.currentTab] = flattenColumns(cols);
  APP.checkedItems.clear();
  saveTabData(APP.currentTab);
  renderGrid();
  renderTabs();
  updateStockCount();
  showToast(`${keys.length}件削除`, 'success');
});

function getCheckedCodes() {
  const cols = splitToColumns(APP.tabs[APP.currentTab] || [], APP.COL_COUNT);
  const codes = [];
  for (const key of APP.checkedItems) {
    const [, colStr, rowStr] = key.split('-');
    const item = cols[colStr] && cols[colStr][parseInt(rowStr)];
    if (item && item.code) codes.push(item.code);
  }
  return codes;
}

// ============================================================
// ↑↓移動（項目G）
// ============================================================
document.getElementById('btn-move-up').addEventListener('click', () => moveChecked(-1));
document.getElementById('btn-move-down').addEventListener('click', () => moveChecked(1));

function moveChecked(dir) {
  if (APP.checkedItems.size === 0) { showToast('チェック銘柄なし', 'info'); return; }
  const cols = splitToColumns(APP.tabs[APP.currentTab] || [], APP.COL_COUNT);
  const byCol = {};
  for (const key of APP.checkedItems) {
    const [, c, r] = key.split('-');
    if (!byCol[c]) byCol[c] = [];
    byCol[c].push(parseInt(r));
  }
  APP.checkedItems.clear();
  for (const c of Object.keys(byCol)) {
    const rows = byCol[c].sort((a, b) => dir === -1 ? a - b : b - a);
    for (const r of rows) {
      const nr = r + dir;
      if (nr < 0 || nr >= cols[c].length) continue;
      if (cols[c][nr]?.type === 'section' || cols[c][r]?.type === 'section') continue;
      [cols[c][r], cols[c][nr]] = [cols[c][nr], cols[c][r]];
      APP.checkedItems.add(`${APP.currentTab}-${c}-${nr}`);
    }
  }
  APP.tabs[APP.currentTab] = flattenColumns(cols);
  saveTabData(APP.currentTab);
  renderGrid();
}

// ============================================================
// 管理画面（項目B, M）
// ============================================================
document.getElementById('btn-admin').addEventListener('click', () => {
  document.getElementById('admin-modal').classList.remove('hidden');
  updateMasterCountDisplay();
});
document.getElementById('btn-help').addEventListener('click', () => {
  document.getElementById('help-modal').classList.remove('hidden');
});

// XLSアップロード
const xlsArea = document.getElementById('xls-upload-area');
const xlsInput = document.getElementById('xls-file-input');
xlsArea.addEventListener('click', () => xlsInput.click());
xlsArea.addEventListener('dragover', (e) => { e.preventDefault(); xlsArea.classList.add('drag-active'); });
xlsArea.addEventListener('dragleave', () => xlsArea.classList.remove('drag-active'));
xlsArea.addEventListener('drop', (e) => { e.preventDefault(); xlsArea.classList.remove('drag-active'); if (e.dataTransfer.files.length) processXlsFile(e.dataTransfer.files[0]); });
xlsInput.addEventListener('change', () => { if (xlsInput.files.length) processXlsFile(xlsInput.files[0]); });

async function processXlsFile(file) {
  const status = document.getElementById('xls-upload-status');
  status.innerHTML = '<div class="status-msg" style="color:var(--accent);border:1px solid var(--accent);background:rgba(0,180,216,0.1);">⏳ 読み込み中...</div>';

  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // 列自動検出
    let codeCol = -1, nameCol = -1;
    for (let r = 0; r < Math.min(5, json.length); r++) {
      const row = json[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const v = String(row[c] || '').trim();
        if (/コード|Code/i.test(v)) codeCol = c;
        if (/銘柄名|銘柄|会社名|Name/i.test(v)) nameCol = c;
      }
      if (codeCol >= 0) break;
    }

    if (codeCol < 0) {
      for (let r = 1; r < Math.min(10, json.length); r++) {
        const row = json[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          if (/^\d{3,4}[A-Z0-9]?$/.test(String(row[c] || '').trim())) {
            codeCol = c;
            nameCol = c + 1;
            break;
          }
        }
        if (codeCol >= 0) break;
      }
    }

    if (codeCol < 0) {
      status.innerHTML = '<div class="status-msg error">❌ 銘柄コード列を検出できません</div>';
      return;
    }

    const masterData = {};
    for (let r = 1; r < json.length; r++) {
      const row = json[r];
      if (!row) continue;
      let code = String(row[codeCol] || '').trim();
      let name = String(row[nameCol] || '').trim();
      if (!code || !name) continue;
      if (code.length === 5 && code.endsWith('0') && /^\d+$/.test(code)) code = code.slice(0, 4);
      if (/^\d{3,4}[A-Z0-9]?$/.test(code)) masterData[code] = name;
    }

    const count = Object.keys(masterData).length;
    if (count === 0) { status.innerHTML = '<div class="status-msg error">❌ 銘柄データが見つかりません</div>'; return; }

    await saveCustomMaster(masterData);
    status.innerHTML = `<div class="status-msg success">✅ ${count}銘柄をインポートしました！</div>`;
    updateMasterCountDisplay();
    regenerateBandaiTabs();
    showToast(`${count}銘柄インポート完了`, 'success');
  } catch (err) {
    console.error('XLS error:', err);
    status.innerHTML = `<div class="status-msg error">❌ ${err.message}</div>`;
  }
}

// CSV更新
const csvArea = document.getElementById('csv-upload-area');
const csvInput = document.getElementById('csv-file-input');
csvArea.addEventListener('click', () => csvInput.click());
csvArea.addEventListener('dragover', (e) => { e.preventDefault(); csvArea.classList.add('drag-active'); });
csvArea.addEventListener('dragleave', () => csvArea.classList.remove('drag-active'));
csvArea.addEventListener('drop', (e) => { e.preventDefault(); csvArea.classList.remove('drag-active'); if (e.dataTransfer.files.length) processCsvFile(e.dataTransfer.files[0]); });
csvInput.addEventListener('change', () => { if (csvInput.files.length) processCsvFile(csvInput.files[0]); });

async function processCsvFile(file) {
  const status = document.getElementById('csv-upload-status');
  try {
    const text = await file.text();
    const codes = text.split(/[\r\n,\t;]+/).map(s => s.trim().replace(/"/g, '')).filter(c => /^\d{3,4}[A-Z0-9]?$/.test(c));
    if (codes.length === 0) { status.innerHTML = '<div class="status-msg error">❌ 有効なコードなし</div>'; return; }
    status.innerHTML = `<div class="status-msg success">✅ ${codes.length}銘柄読み込み — ※現バージョンではstock-data.jsのINDEXデータ更新は管理画面から直接行えます</div>`;
    showToast(`${codes.length}銘柄読み込み`, 'success');
  } catch (err) {
    status.innerHTML = `<div class="status-msg error">❌ ${err.message}</div>`;
  }
}

// 番台再生成
document.getElementById('btn-regen-bandai').addEventListener('click', () => {
  if (!confirm('タブ21-30を再生成します。既存データは上書きされます。')) return;
  regenerateBandaiTabs();
  showToast('番台タブ再生成完了', 'success');
});

function regenerateBandaiTabs() {
  const master = getMaster();
  const bandai = getStocksByBandai(master);
  for (let b = 1; b <= 9; b++) {
    APP.tabs[20 + b] = assignColumns(bandai[b].map(s => ({ code: s.code, name: s.name })));
    APP.tabNames[20 + b] = `${b}000番台`;
    saveTabData(20 + b);
  }
  APP.tabs[30] = assignColumns(bandai['special'].map(s => ({ code: s.code, name: s.name })));
  APP.tabNames[30] = '3桁＋Aコード';
  saveTabData(30);
  saveSettings();
  renderTabs();
  if (APP.currentTab >= 21) renderGrid();
}

// プリセット初期化
document.getElementById('btn-reset-presets').addEventListener('click', () => {
  if (!confirm('タブ18-20をプリセットに戻します？')) return;
  const master = getMaster();
  const n225set = new Set(INDEX_NIKKEI225);
  const topixOnly = INDEX_TOPIX100.filter(c => !n225set.has(c));

  APP.tabs[18] = buildIndexTabData([{ name: 'JPX150', codes: INDEX_JPX150 || INDEX_NIKKEI225.slice(0, 150) }], master);
  APP.tabNames[18] = 'JPX150＆G250';
  APP.tabs[19] = buildIndexTabData([{ name: 'JPX400', codes: INDEX_JPX400 || INDEX_NIKKEI225 }], master);
  APP.tabNames[19] = 'JPX400';
  APP.tabs[20] = buildIndexTabData([
    { name: '日経225', codes: INDEX_NIKKEI225 },
    { name: 'TOPIX100（日経225除く）', codes: topixOnly }
  ], master);
  APP.tabNames[20] = '日経225＆TOPIX100';

  for (let t = 18; t <= 20; t++) saveTabData(t);
  saveSettings();
  renderTabs();
  if (APP.currentTab >= 18 && APP.currentTab <= 20) renderGrid();
  showToast('プリセット初期化完了', 'success');
});

function updateMasterCountDisplay() {
  const el = document.getElementById('master-count-display');
  if (el) el.textContent = `現在のマスタ: ${Object.keys(getMaster()).length}銘柄`;
}

// ============================================================
// キーボード
// ============================================================
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') document.querySelectorAll('.modal-overlay:not(.hidden)').forEach(m => m.classList.add('hidden'));
});

console.log('マイ銘柄ソートリスト v2 loaded');
