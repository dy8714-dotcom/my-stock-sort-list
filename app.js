/**
 * app.js — マイ銘柄ソートリスト メインアプリケーション
 * 全改修項目 A〜M 対応版
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
  tabs: {},            // { 1: [{code,name,type?},...], 2: [...], ... }
  tabNames: {},        // { 1:"", 2:"", ..., 20:"日経225＆TOPIX100" }
  tabOrder: [],        // [1,2,...,30]
  customMaster: null,  // Firestoreから読み込んだカスタムマスタ
  checkedItems: new Set(), // "tabId-colIdx-rowIdx" のSet
  dragState: null,
  saving: false,
  saveTimeout: null,
  COL_COUNT: 6,
  MAX_ROWS_TAB_1_20: 50,
  MAX_ROWS_TAB_21_30: 200,
  MAX_SECTIONS: 10,
  // 指数データ（初期値はstock-data.jsから、CSVアップロードで更新可能）
  indices: {
    nikkei225: [...INDEX_NIKKEI225],
    topix100: [...INDEX_TOPIX100],
    jpx150: [...INDEX_JPX150],
    jpx400: [...INDEX_JPX400],
    growth250: [...INDEX_GROWTH250]
  }
};

// ============================================================
// ユーティリティ
// ============================================================
function showToast(msg, type = 'info') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = msg;
  container.appendChild(toast);
  setTimeout(() => { toast.remove(); }, 3500);
}

function getMaster() {
  return APP.customMaster || STOCK_MASTER;
}

function resolveName(code) {
  const m = getMaster();
  return m[code] || STOCK_MASTER[code] || code + '（不明）';
}

/** タブデータからフラット銘柄リストを取得（セクション除外） */
function getStockCodes(tabData) {
  if (!tabData) return [];
  return tabData.filter(r => r.type !== 'section').map(r => r.code);
}

/** 重複コードSet取得 */
function findDuplicates(tabData) {
  const codes = getStockCodes(tabData);
  const seen = new Set();
  const dupes = new Set();
  for (const c of codes) {
    if (seen.has(c)) dupes.add(c);
    seen.add(c);
  }
  return dupes;
}

/** タブデータを6列に分割。セクションは分割ポイント */
function splitToColumns(tabData, colCount) {
  if (!tabData || tabData.length === 0) {
    return Array.from({ length: colCount }, () => []);
  }
  // tabDataはフラットな配列。colCountの列に均等配置
  // セクションはその列の先頭に配置
  const cols = Array.from({ length: colCount }, () => []);
  let colIdx = 0;

  // 各アイテムの_colプロパティがあればそれに従う。なければ順番に配置
  const hasColInfo = tabData.some(r => r._col !== undefined);

  if (hasColInfo) {
    for (const item of tabData) {
      const c = Math.min(item._col || 0, colCount - 1);
      cols[c].push(item);
    }
  } else {
    // 均等配置
    const itemsPerCol = Math.ceil(tabData.length / colCount);
    for (let i = 0; i < tabData.length; i++) {
      const c = Math.min(Math.floor(i / itemsPerCol), colCount - 1);
      cols[c].push(tabData[i]);
    }
  }
  return cols;
}

/** 列データからフラットデータに戻す（_col情報付加） */
function flattenColumns(cols) {
  const result = [];
  for (let c = 0; c < cols.length; c++) {
    for (const item of cols[c]) {
      result.push({ ...item, _col: c });
    }
  }
  return result;
}

// ============================================================
// Firebase Auth
// ============================================================
document.getElementById('btn-google-login').addEventListener('click', () => {
  auth.signInWithPopup(provider).catch(err => {
    document.getElementById('login-status').textContent = 'ログインエラー: ' + err.message;
  });
});

document.getElementById('btn-logout').addEventListener('click', () => {
  auth.signOut();
});

auth.onAuthStateChanged(async (user) => {
  if (user) {
    APP.uid = user.uid;
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('app-container').classList.remove('hidden');
    document.getElementById('user-name').textContent = user.displayName || user.email;
    if (user.photoURL) {
      document.getElementById('user-avatar').src = user.photoURL;
    }
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
// ============================================================
async function loadUserData() {
  if (!APP.uid) return;
  try {
    // 設定読み込み
    const settingsDoc = await db.collection('users').doc(APP.uid).collection('settings').doc('main').get();
    if (settingsDoc.exists) {
      const data = settingsDoc.data();
      APP.tabOrder = data.tabOrder || Array.from({ length: 30 }, (_, i) => i + 1);
      APP.tabNames = data.tabNames || {};
      if (data.indices) {
        // CSVアップロードで保存した指数データ
        for (const key of Object.keys(data.indices)) {
          APP.indices[key] = data.indices[key];
        }
      }
    } else {
      APP.tabOrder = Array.from({ length: 30 }, (_, i) => i + 1);
      APP.tabNames = getDefaultTabNames();
    }

    // カスタムマスタ読み込み
    const masterDoc = await db.collection('users').doc(APP.uid).collection('settings').doc('customMaster').get();
    if (masterDoc.exists) {
      APP.customMaster = masterDoc.data().data || null;
    }

    // タブデータ読み込み
    const tabsSnapshot = await db.collection('users').doc(APP.uid).collection('tabs').get();
    APP.tabs = {};
    tabsSnapshot.forEach(doc => {
      APP.tabs[doc.id] = doc.data().stocks || [];
    });

    // 初回なら初期データ生成
    if (Object.keys(APP.tabs).length === 0) {
      await initializeDefaultTabs();
    }

    updateMasterCountDisplay();
  } catch (err) {
    console.error('データ読み込みエラー:', err);
    showToast('データ読み込みエラー', 'error');
  }
}

function getDefaultTabNames() {
  const names = {};
  for (let i = 1; i <= 17; i++) names[i] = '';
  names[18] = 'JPX150＆G250';
  names[19] = 'JPX400';
  names[20] = '日経225＆TOPIX100';
  names[21] = '1000番台';
  names[22] = '2000番台';
  names[23] = '3000番台';
  names[24] = '4000番台';
  names[25] = '5000番台';
  names[26] = '6000番台';
  names[27] = '7000番台';
  names[28] = '8000番台';
  names[29] = '9000番台';
  names[30] = '3桁＋Aコード';
  return names;
}

async function initializeDefaultTabs() {
  APP.tabNames = getDefaultTabNames();
  const master = getMaster();

  // タブ1-17: 空
  for (let i = 1; i <= 17; i++) APP.tabs[i] = [];

  // タブ18: JPX150 + グロース250
  APP.tabs[18] = buildIndexTab(['JPX150', 'グロース250'],
    [APP.indices.jpx150, APP.indices.growth250], master);

  // タブ19: JPX400
  APP.tabs[19] = buildIndexTab(['JPX400'], [APP.indices.jpx400], master);

  // タブ20: 日経225 + TOPIX100（重複除外）
  APP.tabs[20] = buildNikkeiTopixTab(master);

  // タブ21-29: 番台（Aコード除外）
  const bandaiData = getStocksByBandai(master);
  for (let b = 1; b <= 9; b++) {
    APP.tabs[20 + b] = bandaiData[b].map(s => ({ code: s.code, name: s.name, _col: undefined }));
    // 列配置
    const items = APP.tabs[20 + b];
    const colSize = Math.ceil(items.length / APP.COL_COUNT);
    items.forEach((item, idx) => {
      item._col = Math.min(Math.floor(idx / colSize), APP.COL_COUNT - 1);
    });
  }

  // タブ30: 3桁+Aコード
  APP.tabs[30] = bandaiData['special'].map((s, idx) => {
    const colSize = Math.ceil(bandaiData['special'].length / APP.COL_COUNT);
    return { code: s.code, name: s.name, _col: Math.min(Math.floor(idx / colSize), APP.COL_COUNT - 1) };
  });

  await saveAllData();
}

function buildIndexTab(sectionNames, codeLists, master) {
  const result = [];
  for (let i = 0; i < sectionNames.length; i++) {
    result.push({ type: 'section', label: sectionNames[i], _col: 0 });
    const codes = codeLists[i];
    const colSize = Math.ceil(codes.length / APP.COL_COUNT);
    codes.forEach((code, idx) => {
      result.push({
        code,
        name: master[code] || STOCK_MASTER[code] || code,
        _col: Math.min(Math.floor(idx / colSize), APP.COL_COUNT - 1)
      });
    });
  }
  return result;
}

function buildNikkeiTopixTab(master) {
  const result = [];
  // 日経225セクション
  result.push({ type: 'section', label: '日経225', _col: 0 });
  const n225 = APP.indices.nikkei225;
  const colSize1 = Math.ceil(n225.length / APP.COL_COUNT);
  n225.forEach((code, idx) => {
    result.push({
      code,
      name: master[code] || STOCK_MASTER[code] || code,
      _col: Math.min(Math.floor(idx / colSize1), APP.COL_COUNT - 1)
    });
  });

  // TOPIX100セクション（日経225と重複するものを除外）
  const n225Set = new Set(n225);
  const topixOnly = APP.indices.topix100.filter(c => !n225Set.has(c));
  result.push({ type: 'section', label: 'TOPIX100（日経225除く）', _col: 0 });
  const colSize2 = Math.ceil(topixOnly.length / APP.COL_COUNT);
  topixOnly.forEach((code, idx) => {
    result.push({
      code,
      name: master[code] || STOCK_MASTER[code] || code,
      _col: Math.min(Math.floor(idx / colSize2), APP.COL_COUNT - 1)
    });
  });
  return result;
}

async function saveTabData(tabId) {
  if (!APP.uid || APP.saving) return;
  // デバウンス
  if (APP.saveTimeout) clearTimeout(APP.saveTimeout);
  APP.saveTimeout = setTimeout(async () => {
    APP.saving = true;
    try {
      // _colを保持して保存
      const data = (APP.tabs[tabId] || []).map(item => {
        const obj = { ...item };
        // Firestoreに保存するフィールドのみ
        return obj;
      });
      await db.collection('users').doc(APP.uid).collection('tabs').doc(String(tabId)).set({ stocks: data });
    } catch (err) {
      console.error('保存エラー:', err);
      showToast('保存に失敗しました', 'error');
    } finally {
      APP.saving = false;
    }
  }, 500);
}

async function saveSettings() {
  if (!APP.uid) return;
  try {
    await db.collection('users').doc(APP.uid).collection('settings').doc('main').set({
      tabOrder: APP.tabOrder,
      tabNames: APP.tabNames,
      indices: APP.indices
    });
  } catch (err) {
    console.error('設定保存エラー:', err);
  }
}

async function saveCustomMaster(masterData) {
  if (!APP.uid) return;
  try {
    // Firestoreの1ドキュメントサイズ制限(1MB)を考慮
    // 大きい場合は複数ドキュメントに分割
    const entries = Object.entries(masterData);
    if (entries.length <= 5000) {
      await db.collection('users').doc(APP.uid).collection('settings').doc('customMaster').set({ data: masterData });
    } else {
      // 分割保存
      const chunk1 = {};
      const chunk2 = {};
      entries.forEach(([k, v], i) => {
        if (i < 3000) chunk1[k] = v;
        else chunk2[k] = v;
      });
      await db.collection('users').doc(APP.uid).collection('settings').doc('customMaster').set({ data: chunk1 });
      await db.collection('users').doc(APP.uid).collection('settings').doc('customMaster2').set({ data: chunk2 });
    }
    APP.customMaster = masterData;
  } catch (err) {
    console.error('マスタ保存エラー:', err);
    throw err;
  }
}

async function saveAllData() {
  if (!APP.uid) return;
  try {
    await saveSettings();
    for (const tabId of Object.keys(APP.tabs)) {
      await db.collection('users').doc(APP.uid).collection('tabs').doc(String(tabId)).set({
        stocks: APP.tabs[tabId] || []
      });
    }
    showToast('保存完了', 'success');
  } catch (err) {
    console.error('一括保存エラー:', err);
    showToast('保存に失敗しました', 'error');
  }
}

// ============================================================
// タブ描画
// ============================================================
function renderTabs() {
  const bar = document.getElementById('tab-bar');
  bar.innerHTML = '';
  const order = APP.tabOrder.length ? APP.tabOrder : Array.from({ length: 30 }, (_, i) => i + 1);

  for (const tabId of order) {
    const div = document.createElement('div');
    div.className = 'tab-item' + (tabId === APP.currentTab ? ' active' : '');
    const name = APP.tabNames[tabId] || '';
    const count = (APP.tabs[tabId] || []).filter(r => r.type !== 'section').length;
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

  // タブアクティブ表示更新
  document.querySelectorAll('.tab-item').forEach(el => {
    el.classList.toggle('active', parseInt(el.dataset.tabId) === tabId);
  });

  // タブ名入力欄更新
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
    if (e.key === 'Escape') { renderTabs(); }
  });
}

function updateStockCount() {
  const data = APP.tabs[APP.currentTab] || [];
  const count = data.filter(r => r.type !== 'section').length;
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
        <input type="checkbox" class="select-all-cb" data-col="${colIdx}" title="列${colIdx + 1}の全銘柄を選択">
      </div>
    `;
    colDiv.appendChild(header);

    // ボディ
    const body = document.createElement('div');
    body.className = 'column-body';
    body.dataset.col = colIdx;

    const colItems = cols[colIdx];
    let rowNum = 0;

    if (colItems.length === 0) {
      const placeholder = document.createElement('div');
      placeholder.className = 'empty-placeholder';
      placeholder.textContent = 'ドロップまたは入力';
      body.appendChild(placeholder);
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

    // ドロップゾーン
    setupDropZone(body, colIdx);
    colDiv.appendChild(body);
    grid.appendChild(colDiv);
  }

  // 全選択チェックボックスイベント
  document.querySelectorAll('.select-all-cb').forEach(cb => {
    cb.addEventListener('change', (e) => {
      const colIdx = parseInt(e.target.dataset.col);
      const checked = e.target.checked;
      const body = document.querySelector(`.column-body[data-col="${colIdx}"]`);
      body.querySelectorAll('.stock-checkbox').forEach(scb => {
        scb.checked = checked;
        const key = scb.dataset.key;
        if (checked) APP.checkedItems.add(key);
        else APP.checkedItems.delete(key);
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

  div.innerHTML = `
    <span class="section-label">📁 ${item.label}</span>
    <div class="section-actions">
      <button class="btn btn-small" title="セクション名変更" data-action="rename-section">✏</button>
      <button class="btn btn-small btn-danger" title="セクション削除" data-action="delete-section">✕</button>
    </div>
  `;

  // セクション名変更
  div.querySelector('[data-action="rename-section"]').addEventListener('click', (e) => {
    e.stopPropagation();
    const newName = prompt('セクション名:', item.label);
    if (newName !== null) {
      item.label = newName;
      saveTabData(APP.currentTab);
      renderGrid();
    }
  });

  // セクション削除
  div.querySelector('[data-action="delete-section"]').addEventListener('click', (e) => {
    e.stopPropagation();
    if (confirm(`セクション「${item.label}」を削除しますか？（中の銘柄は残ります）`)) {
      const tabData = APP.tabs[APP.currentTab];
      const flatIdx = findFlatIndex(colIdx, rowIdx);
      if (flatIdx >= 0) {
        tabData.splice(flatIdx, 1);
        saveTabData(APP.currentTab);
        renderGrid();
        updateStockCount();
      }
    }
  });

  // ドラッグ
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

  const isDuplicate = duplicates.has(item.code);
  const name = resolveName(item.code);

  div.innerHTML = `
    <span class="row-num">${rowNum}</span>
    <span class="stock-code" data-code="${item.code}" title="クリックして編集">${item.code}</span>
    <span class="stock-name${isDuplicate ? ' duplicate' : ''}" title="${name}">${name}</span>
    <input type="checkbox" class="stock-checkbox" data-key="${key}" ${isChecked ? 'checked' : ''}>
  `;

  // コード直接編集（項目H）
  const codeEl = div.querySelector('.stock-code');
  codeEl.addEventListener('click', (e) => {
    e.stopPropagation();
    startCodeEdit(codeEl, item, colIdx, rowIdx);
  });

  // チェックボックス
  const cb = div.querySelector('.stock-checkbox');
  cb.addEventListener('change', (e) => {
    e.stopPropagation();
    if (cb.checked) {
      APP.checkedItems.add(key);
      div.classList.add('selected');
    } else {
      APP.checkedItems.delete(key);
      div.classList.remove('selected');
    }
  });

  // ドラッグ
  setupDragSource(div, colIdx, rowIdx);
  return div;
}

// ============================================================
// コード直接編集（項目H）
// ============================================================
function startCodeEdit(codeEl, item, colIdx, rowIdx) {
  if (codeEl.classList.contains('editing')) return;
  codeEl.classList.add('editing');
  const original = item.code;

  const input = document.createElement('input');
  input.type = 'text';
  input.value = item.code;
  input.style.cssText = 'width:50px;font-family:var(--font-mono);font-size:12px;padding:1px 2px;border:none;background:transparent;color:var(--text-bright);outline:none;';
  codeEl.textContent = '';
  codeEl.appendChild(input);
  input.focus();
  input.select();

  const finish = (save) => {
    codeEl.classList.remove('editing');
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
    if (e.key === 'Escape') { finish(false); }
  });
}

// ============================================================
// ドラッグ＆ドロップ（項目D）
// ============================================================
function setupDragSource(el, colIdx, rowIdx) {
  el.addEventListener('dragstart', (e) => {
    APP.dragState = { colIdx, rowIdx, tabId: APP.currentTab };
    el.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', JSON.stringify({ colIdx, rowIdx }));
  });

  el.addEventListener('dragend', () => {
    el.classList.remove('dragging');
    document.querySelectorAll('.drag-target-above,.drag-target-below,.drag-over').forEach(
      el => el.classList.remove('drag-target-above', 'drag-target-below', 'drag-over')
    );
    APP.dragState = null;
  });
}

function setupDropZone(bodyEl, targetColIdx) {
  bodyEl.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    bodyEl.classList.add('drag-over');

    // 挿入位置の視覚フィードバック
    const rows = bodyEl.querySelectorAll('.stock-row, .section-header');
    rows.forEach(r => r.classList.remove('drag-target-above', 'drag-target-below'));

    const closest = getClosestRow(e.clientY, rows);
    if (closest.el) {
      if (closest.above) closest.el.classList.add('drag-target-above');
      else closest.el.classList.add('drag-target-below');
    }
  });

  bodyEl.addEventListener('dragleave', (e) => {
    if (!bodyEl.contains(e.relatedTarget)) {
      bodyEl.classList.remove('drag-over');
      bodyEl.querySelectorAll('.drag-target-above,.drag-target-below').forEach(
        el => el.classList.remove('drag-target-above', 'drag-target-below')
      );
    }
  });

  bodyEl.addEventListener('drop', (e) => {
    e.preventDefault();
    bodyEl.classList.remove('drag-over');
    bodyEl.querySelectorAll('.drag-target-above,.drag-target-below').forEach(
      el => el.classList.remove('drag-target-above', 'drag-target-below')
    );

    if (!APP.dragState) return;

    const { colIdx: srcCol, rowIdx: srcRow } = APP.dragState;
    const tabData = APP.tabs[APP.currentTab];
    const cols = splitToColumns(tabData, APP.COL_COUNT);

    // ソースアイテム取得
    const srcItem = cols[srcCol][srcRow];
    if (!srcItem) return;

    // ドロップ位置計算
    const rows = bodyEl.querySelectorAll('.stock-row, .section-header');
    const closest = getClosestRow(e.clientY, rows);
    let targetRow;
    if (closest.el) {
      targetRow = parseInt(closest.el.dataset.row);
      if (!closest.above) targetRow++;
    } else {
      targetRow = cols[targetColIdx].length;
    }

    // ソースから削除
    cols[srcCol].splice(srcRow, 1);

    // 同じ列で後方から前方への移動の場合、インデックス調整
    if (srcCol === targetColIdx && srcRow < targetRow) {
      targetRow--;
    }

    // ターゲットに挿入
    cols[targetColIdx].splice(targetRow, 0, srcItem);

    // フラット化して保存
    APP.tabs[APP.currentTab] = flattenColumns(cols);
    saveTabData(APP.currentTab);
    renderGrid();
    APP.dragState = null;
  });
}

function getClosestRow(y, rows) {
  let closest = { el: null, distance: Infinity, above: true };
  rows.forEach(row => {
    const rect = row.getBoundingClientRect();
    const mid = rect.top + rect.height / 2;
    const dist = Math.abs(y - mid);
    if (dist < closest.distance) {
      closest = { el: row, distance: dist, above: y < mid };
    }
  });
  return closest;
}

function findFlatIndex(colIdx, rowIdx) {
  const tabData = APP.tabs[APP.currentTab];
  const cols = splitToColumns(tabData, APP.COL_COUNT);
  const item = cols[colIdx][rowIdx];
  if (!item) return -1;
  return tabData.indexOf(item);
}

// ============================================================
// ツールバー操作
// ============================================================

// 銘柄追加
document.getElementById('btn-add-stock').addEventListener('click', addStock);
document.getElementById('add-stock-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') addStock();
});

function addStock() {
  const input = document.getElementById('add-stock-input');
  const rawCodes = input.value.trim();
  if (!rawCodes) return;

  const codes = rawCodes.split(/[\s,;、]+/).filter(c => c);
  const targetCol = parseInt(document.getElementById('input-column-select').value);
  const tabData = APP.tabs[APP.currentTab] || [];
  const cols = splitToColumns(tabData, APP.COL_COUNT);

  for (const code of codes) {
    const upperCode = code.toUpperCase();
    cols[targetCol].push({
      code: upperCode,
      name: resolveName(upperCode),
      _col: targetCol
    });
  }

  APP.tabs[APP.currentTab] = flattenColumns(cols);
  saveTabData(APP.currentTab);
  renderGrid();
  renderTabs(); // カウント更新
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

  // セクション数チェック
  const sectionCount = tabData.filter(r => r.type === 'section').length;
  if (sectionCount >= APP.MAX_SECTIONS * APP.COL_COUNT) {
    showToast('セクション数の上限に達しました', 'error');
    return;
  }

  cols[targetCol].unshift({ type: 'section', label: name, _col: targetCol });
  APP.tabs[APP.currentTab] = flattenColumns(cols);
  saveTabData(APP.currentTab);
  renderGrid();
  input.value = '';
  showToast(`セクション「${name}」を追加`, 'success');
});

// タブ名変更
document.getElementById('btn-rename-tab').addEventListener('click', () => {
  const input = document.getElementById('tab-name-input');
  APP.tabNames[APP.currentTab] = input.value;
  saveSettings();
  renderTabs();
  showToast('タブ名を変更しました', 'success');
});

// ============================================================
// チェックボックス関連操作（項目E）
// ============================================================

// コード出力
document.getElementById('btn-export-codes').addEventListener('click', () => {
  const codes = getCheckedCodes();
  if (codes.length === 0) {
    showToast('チェックされた銘柄がありません', 'info');
    return;
  }
  document.getElementById('export-textarea').value = codes.join('\n');
  document.getElementById('export-modal').classList.remove('hidden');
});

// 選択削除
document.getElementById('btn-delete-checked').addEventListener('click', () => {
  const checkedKeys = [...APP.checkedItems];
  if (checkedKeys.length === 0) {
    showToast('チェックされた銘柄がありません', 'info');
    return;
  }
  if (!confirm(`${checkedKeys.length}件の銘柄を削除しますか？`)) return;

  const tabData = APP.tabs[APP.currentTab];
  const cols = splitToColumns(tabData, APP.COL_COUNT);

  // 削除対象のインデックスを収集（逆順で削除）
  const toDelete = {};
  for (const key of checkedKeys) {
    const [, colStr, rowStr] = key.split('-');
    const col = parseInt(colStr);
    const row = parseInt(rowStr);
    if (!toDelete[col]) toDelete[col] = [];
    toDelete[col].push(row);
  }

  for (const col of Object.keys(toDelete)) {
    const rows = toDelete[col].sort((a, b) => b - a); // 逆順
    for (const row of rows) {
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
  showToast(`${checkedKeys.length}件削除しました`, 'success');
});

function getCheckedCodes() {
  const tabData = APP.tabs[APP.currentTab];
  const cols = splitToColumns(tabData, APP.COL_COUNT);
  const codes = [];

  for (const key of APP.checkedItems) {
    const [, colStr, rowStr] = key.split('-');
    const col = parseInt(colStr);
    const row = parseInt(rowStr);
    if (cols[col] && cols[col][row] && cols[col][row].code) {
      codes.push(cols[col][row].code);
    }
  }
  return codes;
}

// ============================================================
// ↑↓移動ボタン（項目G）
// ============================================================
document.getElementById('btn-move-up').addEventListener('click', () => moveChecked(-1));
document.getElementById('btn-move-down').addEventListener('click', () => moveChecked(1));

function moveChecked(direction) {
  const checkedKeys = [...APP.checkedItems].sort();
  if (checkedKeys.length === 0) {
    showToast('チェックされた銘柄がありません', 'info');
    return;
  }

  const tabData = APP.tabs[APP.currentTab];
  const cols = splitToColumns(tabData, APP.COL_COUNT);

  // 列ごとにグループ化
  const byCol = {};
  for (const key of checkedKeys) {
    const [, colStr, rowStr] = key.split('-');
    const col = parseInt(colStr);
    const row = parseInt(rowStr);
    if (!byCol[col]) byCol[col] = [];
    byCol[col].push(row);
  }

  APP.checkedItems.clear();

  for (const colStr of Object.keys(byCol)) {
    const col = parseInt(colStr);
    const rows = byCol[col].sort((a, b) => direction === -1 ? a - b : b - a);

    for (const row of rows) {
      const newRow = row + direction;
      if (newRow < 0 || newRow >= cols[col].length) continue;
      // セクションはスキップ
      if (cols[col][newRow] && cols[col][newRow].type === 'section') continue;
      if (cols[col][row] && cols[col][row].type === 'section') continue;

      // スワップ
      [cols[col][row], cols[col][newRow]] = [cols[col][newRow], cols[col][row]];

      // チェック状態を新位置に
      APP.checkedItems.add(`${APP.currentTab}-${col}-${newRow}`);
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

// --- XLSアップロード（項目B） ---
const xlsArea = document.getElementById('xls-upload-area');
const xlsInput = document.getElementById('xls-file-input');

xlsArea.addEventListener('click', () => xlsInput.click());
xlsArea.addEventListener('dragover', (e) => { e.preventDefault(); xlsArea.classList.add('drag-active'); });
xlsArea.addEventListener('dragleave', () => xlsArea.classList.remove('drag-active'));
xlsArea.addEventListener('drop', (e) => {
  e.preventDefault();
  xlsArea.classList.remove('drag-active');
  if (e.dataTransfer.files.length) processXlsFile(e.dataTransfer.files[0]);
});
xlsInput.addEventListener('change', () => {
  if (xlsInput.files.length) processXlsFile(xlsInput.files[0]);
});

async function processXlsFile(file) {
  const statusEl = document.getElementById('xls-upload-status');
  statusEl.innerHTML = '<div class="status-msg info" style="color:var(--accent);border:1px solid var(--accent);background:rgba(0,180,216,0.1);">⏳ 読み込み中...</div>';

  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // JPX data_j.xls の形式:
    // 通常はヘッダ行があり、銘柄コードは1列目or2列目、銘柄名は2列目or3列目
    // 列を自動検出
    const masterData = {};
    let codeColIdx = -1;
    let nameColIdx = -1;

    // ヘッダ行を探す
    for (let r = 0; r < Math.min(5, json.length); r++) {
      const row = json[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c] || '').trim();
        if (val === 'コード' || val === '銘柄コード' || val === 'Code' || val === 'code') codeColIdx = c;
        if (val === '銘柄名' || val === '会社名' || val === '銘柄' || val === 'Name' || val === 'name') nameColIdx = c;
      }
      if (codeColIdx >= 0 && nameColIdx >= 0) break;
    }

    // ヘッダ検出失敗の場合はデフォルト（1列目=コード、2列目=銘柄名とか推測）
    if (codeColIdx < 0) {
      // 数字パターンで推測
      for (let r = 1; r < Math.min(10, json.length); r++) {
        const row = json[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || '').trim();
          if (/^\d{4}[A-Z0-9]?$/.test(val) || /^\d{3}$/.test(val) || /^\d{3}[A-Z]$/.test(val)) {
            codeColIdx = c;
            nameColIdx = c + 1;
            break;
          }
        }
        if (codeColIdx >= 0) break;
      }
    }

    if (codeColIdx < 0) {
      statusEl.innerHTML = '<div class="status-msg error">❌ 銘柄コード列を検出できませんでした。JPX公式のdata_j.xlsか確認してください。</div>';
      return;
    }

    // データ取り込み
    const startRow = json.findIndex((row, i) => {
      if (i === 0) return false;
      const val = String(row?.[codeColIdx] || '').trim();
      return /^\d{3,4}[A-Z0-9]?$/.test(val);
    });

    for (let r = Math.max(startRow, 1); r < json.length; r++) {
      const row = json[r];
      if (!row) continue;
      let code = String(row[codeColIdx] || '').trim();
      let name = String(row[nameColIdx] || '').trim();
      if (!code) continue;
      // 4桁数字 or 3桁数字 or 末尾にA-Zが付くコード
      if (/^\d{3,5}[A-Z0-9]?$/.test(code)) {
        // 5桁で末尾0のTDnetコードは4桁に変換
        if (code.length === 5 && code.endsWith('0')) {
          code = code.slice(0, 4);
        }
        masterData[code] = name;
      }
    }

    const count = Object.keys(masterData).length;
    if (count === 0) {
      statusEl.innerHTML = '<div class="status-msg error">❌ 銘柄データが見つかりませんでした。ファイル形式を確認してください。</div>';
      return;
    }

    // 保存
    await saveCustomMaster(masterData);
    statusEl.innerHTML = `<div class="status-msg success">✅ ${count}銘柄をインポートしました！</div>`;
    updateMasterCountDisplay();

    // 番台タブ再生成
    regenerateBandaiTabs();
    showToast(`${count}銘柄をインポートしました`, 'success');

  } catch (err) {
    console.error('XLS処理エラー:', err);
    statusEl.innerHTML = `<div class="status-msg error">❌ エラー: ${err.message}</div>`;
  }
}

// --- CSV更新（指数構成銘柄） ---
const csvArea = document.getElementById('csv-upload-area');
const csvInput = document.getElementById('csv-file-input');

csvArea.addEventListener('click', () => csvInput.click());
csvArea.addEventListener('dragover', (e) => { e.preventDefault(); csvArea.classList.add('drag-active'); });
csvArea.addEventListener('dragleave', () => csvArea.classList.remove('drag-active'));
csvArea.addEventListener('drop', (e) => {
  e.preventDefault();
  csvArea.classList.remove('drag-active');
  if (e.dataTransfer.files.length) processCsvFile(e.dataTransfer.files[0]);
});
csvInput.addEventListener('change', () => {
  if (csvInput.files.length) processCsvFile(csvInput.files[0]);
});

async function processCsvFile(file) {
  const statusEl = document.getElementById('csv-upload-status');
  const target = document.getElementById('index-csv-target').value;

  try {
    const text = await file.text();
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l);
    const codes = [];
    for (const line of lines) {
      const parts = line.split(/[,\t;]/);
      for (const part of parts) {
        const code = part.trim().replace(/"/g, '');
        if (/^\d{3,4}[A-Z0-9]?$/.test(code)) {
          codes.push(code);
        }
      }
    }

    if (codes.length === 0) {
      statusEl.innerHTML = '<div class="status-msg error">❌ 有効な銘柄コードが見つかりませんでした。</div>';
      return;
    }

    APP.indices[target] = codes;
    await saveSettings();

    const targetNames = {
      nikkei225: '日経225', topix100: 'TOPIX100', jpx150: 'JPX150',
      jpx400: 'JPX400', growth250: 'グロース250'
    };
    statusEl.innerHTML = `<div class="status-msg success">✅ ${targetNames[target]}: ${codes.length}銘柄を更新しました</div>`;
    showToast(`${targetNames[target]}を更新しました`, 'success');

  } catch (err) {
    statusEl.innerHTML = `<div class="status-msg error">❌ エラー: ${err.message}</div>`;
  }
}

// --- 番台タブ再生成 ---
document.getElementById('btn-regen-bandai').addEventListener('click', () => {
  if (!confirm('タブ21-30を再生成します。既存データは上書きされます。よろしいですか？')) return;
  regenerateBandaiTabs();
  showToast('番台タブを再生成しました', 'success');
});

function regenerateBandaiTabs() {
  const master = getMaster();
  const bandaiData = getStocksByBandai(master);

  for (let b = 1; b <= 9; b++) {
    const items = bandaiData[b];
    const colSize = Math.ceil(items.length / APP.COL_COUNT);
    APP.tabs[20 + b] = items.map((s, idx) => ({
      code: s.code,
      name: s.name,
      _col: Math.min(Math.floor(idx / colSize), APP.COL_COUNT - 1)
    }));
    saveTabData(20 + b);
  }

  // タブ30: 3桁+Aコード
  const special = bandaiData['special'];
  const colSize = Math.ceil(special.length / APP.COL_COUNT);
  APP.tabs[30] = special.map((s, idx) => ({
    code: s.code,
    name: s.name,
    _col: Math.min(Math.floor(idx / colSize), APP.COL_COUNT - 1)
  }));
  saveTabData(30);

  renderTabs();
  if (APP.currentTab >= 21) renderGrid();
}

// --- プリセット初期化 ---
document.getElementById('btn-reset-presets').addEventListener('click', async () => {
  if (!confirm('タブ18-20をプリセットに戻します。よろしいですか？')) return;
  const master = getMaster();

  // タブ18: JPX150 + グロース250
  APP.tabs[18] = buildIndexTab(['JPX150', 'グロース250'],
    [APP.indices.jpx150, APP.indices.growth250], master);
  APP.tabNames[18] = 'JPX150＆G250';
  await saveTabData(18);

  // タブ19: JPX400
  APP.tabs[19] = buildIndexTab(['JPX400'], [APP.indices.jpx400], master);
  APP.tabNames[19] = 'JPX400';
  await saveTabData(19);

  // タブ20: 日経225 + TOPIX100
  APP.tabs[20] = buildNikkeiTopixTab(master);
  APP.tabNames[20] = '日経225＆TOPIX100';
  await saveTabData(20);

  await saveSettings();
  renderTabs();
  if (APP.currentTab >= 18 && APP.currentTab <= 20) renderGrid();
  showToast('プリセットを初期化しました', 'success');
});

function updateMasterCountDisplay() {
  const master = getMaster();
  const count = Object.keys(master).length;
  const el = document.getElementById('master-count-display');
  if (el) el.textContent = `現在のマスタ: ${count}銘柄`;
}

// ============================================================
// キーボードショートカット
// ============================================================
document.addEventListener('keydown', (e) => {
  // Escでモーダル閉じる
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay:not(.hidden)').forEach(m => m.classList.add('hidden'));
  }
});

// ============================================================
// 初期化完了
// ============================================================
console.log('マイ銘柄ソートリスト — app.js loaded');
