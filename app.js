/**
 * app.js v5 — マイ銘柄ソートリスト
 * タブ1-17: Firestoreから読み込み（ユーザーデータ）
 * タブ18-30: 毎回stock-data.jsから生成（常に完全・最新）→Firestoreにも保存
 */
const firebaseConfig={apiKey:"AIzaSyCXpU-IX55WPJCVlhEhbcOZQzPmVKZt9PU",authDomain:"my-stock-list-7ae4d.firebaseapp.com",projectId:"my-stock-list-7ae4d",storageBucket:"my-stock-list-7ae4d.firebasestorage.app",messagingSenderId:"1072316134172",appId:"1:1072316134172:web:27abd8b3b8906afa2480d9"};
firebase.initializeApp(firebaseConfig);
const auth=firebase.auth(),db=firebase.firestore(),provider=new firebase.auth.GoogleAuthProvider();
const APP={uid:null,currentTab:1,tabs:{},tabNames:{},customMaster:null,checkedItems:new Set(),dragState:null,COL_COUNT:6,stockMemos:{},currentMemoThemes:[],currentMemoCode:null};

function showToast(m,t='info'){const d=document.createElement('div');d.className='toast '+t;d.textContent=m;document.getElementById('toast-container').appendChild(d);setTimeout(()=>d.remove(),3500);}
function getMaster(){return APP.customMaster||STOCK_MASTER;}
function resolveName(c){const m=getMaster();return m[c]||STOCK_MASTER[c]||c+'（不明）';}
function findDuplicates(d){const codes=(d||[]).filter(r=>r.type!=='section').map(r=>r.code);const s=new Set(),dup=new Set();for(const c of codes){if(s.has(c))dup.add(c);s.add(c);}return dup;}
function splitToColumns(data,n){if(!data||!data.length)return Array.from({length:n},()=>[]);const cols=Array.from({length:n},()=>[]);if(data.some(r=>r._col!==undefined&&r._col!==null)){for(const item of data)cols[Math.min(Math.max(item._col||0,0),n-1)].push(item);}else{const per=Math.ceil(data.length/n);data.forEach((item,i)=>cols[Math.min(Math.floor(i/per),n-1)].push(item));}return cols;}
function flattenColumns(cols){const r=[];cols.forEach((col,c)=>col.forEach(item=>r.push({...item,_col:c})));return r;}
function compactForSave(tabData){return(tabData||[]).map(item=>{if(item.type==='section')return{type:'section',label:item.label,_col:item._col||0};return{code:item.code,_col:item._col||0};});}

// Auth
document.getElementById('btn-google-login').addEventListener('click',()=>{auth.signInWithPopup(provider).catch(err=>{document.getElementById('login-status').textContent='エラー: '+err.message;});});
document.getElementById('btn-logout').addEventListener('click',()=>auth.signOut());
auth.onAuthStateChanged(async(user)=>{if(user){APP.uid=user.uid;document.getElementById('login-screen').style.display='none';document.getElementById('app-container').classList.remove('hidden');document.getElementById('user-name').textContent=user.displayName||user.email;if(user.photoURL)document.getElementById('user-avatar').src=user.photoURL;await loadAllData();renderTabs();switchTab(APP.currentTab);}else{APP.uid=null;document.getElementById('login-screen').style.display='flex';document.getElementById('app-container').classList.add('hidden');}});

// === データ読み込み ===
function getDefaultTabNames(){const n={};for(let i=1;i<=17;i++)n[i]='';n[18]='JPX150＆グロース250';n[19]='JPX400';n[20]='日経225＆TOPIX100';for(let b=1;b<=9;b++)n[20+b]=b+'000番台';n[30]='3桁＋Aコード';return n;}

async function loadAllData(){
  if(!APP.uid)return;
  APP.tabNames=getDefaultTabNames();
  APP.stockMemos={};
  for(let i=1;i<=30;i++)APP.tabs[i]=[];

  // ★ タブ18-30: 常にstock-data.jsから生成（古いFirestoreデータは無視）
  const preset=generatePresetTabs(APP.COL_COUNT,APP.customMaster);
  for(let t=18;t<=30;t++){
    APP.tabs[t]=preset.tabs[t]||[];
    APP.tabNames[t]=preset.names[t];
  }

  // タブ1-17: Firestoreから読み込み
  try{
    const userDoc=await db.collection('users').doc(APP.uid).get();
    if(userDoc.exists){
      const d=userDoc.data();
      if(d.tabNames){for(let i=1;i<=17;i++)if(d.tabNames[i])APP.tabNames[i]=d.tabNames[i];}
      if(d.customMaster){APP.customMaster=d.customMaster;
        // カスタムマスタがある場合、プリセットを再生成
        const p2=generatePresetTabs(APP.COL_COUNT,d.customMaster);
        for(let t=18;t<=30;t++){APP.tabs[t]=p2.tabs[t]||[];APP.tabNames[t]=p2.names[t];}
      }
    }
    const snap=await db.collection('users').doc(APP.uid).collection('tabs').get();
    snap.forEach(doc=>{
      const id=parseInt(doc.id);const s=doc.data().stocks;
      // タブ1-17のみFirestoreから読み込み
      if(id>=1&&id<=17&&s&&s.length>0){
        APP.tabs[id]=s.map(item=>{if(item.type==='section')return item;return{...item,name:resolveName(item.code)};});
      }
    });
    // 銘柄メモ（stockMemos）の読込（N機能）
    const memosSnap=await db.collection('users').doc(APP.uid).collection('stockMemos').get();
    memosSnap.forEach(doc=>{APP.stockMemos[doc.id]=doc.data();});
  }catch(err){console.warn('Firestore:',err.message);}

  // タブ18-30をFirestoreにも保存（バックグラウンド、エラーは無視）
  for(let t=18;t<=30;t++){
    db.collection('users').doc(APP.uid).collection('tabs').doc(String(t))
      .set({stocks:compactForSave(APP.tabs[t])}).catch(()=>{});
  }
  db.collection('users').doc(APP.uid).set({tabNames:APP.tabNames},{merge:true}).catch(()=>{});

  updateMasterCountDisplay();
}

let saveTimer=null;
function saveTabData(tabId){
  if(!APP.uid||tabId>=18)return; // タブ18-30は自動生成なので保存不要
  clearTimeout(saveTimer);
  saveTimer=setTimeout(()=>{
    db.collection('users').doc(APP.uid).collection('tabs').doc(String(tabId))
      .set({stocks:compactForSave(APP.tabs[tabId])}).catch(e=>showToast('保存エラー','error'));
  },800);
}
function saveSettings(){
  if(!APP.uid)return;
  db.collection('users').doc(APP.uid).set({tabNames:APP.tabNames},{merge:true}).catch(()=>{});
}

// === タブ描画 ===
function renderTabs(){const bar=document.getElementById('tab-bar');bar.innerHTML='';for(let t=1;t<=30;t++){const div=document.createElement('div');div.className='tab-item'+(t===APP.currentTab?' active':'');const name=APP.tabNames[t]||'';const cnt=(APP.tabs[t]||[]).filter(r=>r.type!=='section').length;div.innerHTML='<span>'+(name||'タブ'+t)+'</span><span class="tab-count">('+cnt+')</span>';div.dataset.tabId=t;div.addEventListener('click',()=>switchTab(t));div.addEventListener('dblclick',()=>startTabRename(t,div));bar.appendChild(div);}}
function switchTab(tabId){APP.currentTab=tabId;APP.checkedItems.clear();document.querySelectorAll('.tab-item').forEach(el=>el.classList.toggle('active',parseInt(el.dataset.tabId)===tabId));document.getElementById('tab-name-input').value=APP.tabNames[tabId]||'';(APP.tabs[tabId]||[]).forEach(item=>{if(item.code&&!item.name)item.name=resolveName(item.code);});renderGrid();updateStockCount();}
function startTabRename(tabId,el){const inp=document.createElement('input');inp.type='text';inp.className='tab-name-input';inp.value=APP.tabNames[tabId]||'';el.innerHTML='';el.appendChild(inp);inp.focus();inp.select();const done=()=>{APP.tabNames[tabId]=inp.value;saveSettings();renderTabs();};inp.addEventListener('blur',done);inp.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();done();}if(e.key==='Escape')renderTabs();});}
function updateStockCount(){document.getElementById('stock-count-display').textContent=(APP.tabs[APP.currentTab]||[]).filter(r=>r.type!=='section').length+'銘柄';}

// === グリッド描画 ===
function renderGrid(){const grid=document.getElementById('stock-grid');grid.innerHTML='';const tabData=APP.tabs[APP.currentTab]||[];const cols=splitToColumns(tabData,APP.COL_COUNT);const dups=findDuplicates(tabData);for(let ci=0;ci<APP.COL_COUNT;ci++){const colDiv=document.createElement('div');colDiv.className='stock-column';const hdr=document.createElement('div');hdr.className='column-header';hdr.innerHTML='<span class="col-label">列'+(ci+1)+'</span><div class="select-all-wrap"><label style="font-size:10px;cursor:pointer">全選択</label><input type="checkbox" class="select-all-cb" data-col="'+ci+'"></div>';colDiv.appendChild(hdr);const body=document.createElement('div');body.className='column-body';body.dataset.col=ci;const items=cols[ci];let rn=0;if(!items.length)body.innerHTML='<div class="empty-placeholder">ドロップまたは入力</div>';else for(let ri=0;ri<items.length;ri++){const item=items[ri];if(item.type==='section')body.appendChild(mkSection(item,ci,ri));else{rn++;body.appendChild(mkStockRow(item,ci,ri,rn,dups));}}setupDropZone(body,ci);colDiv.appendChild(body);grid.appendChild(colDiv);}document.querySelectorAll('.select-all-cb').forEach(cb=>{cb.addEventListener('change',e=>{const ci=parseInt(e.target.dataset.col),ck=e.target.checked;document.querySelector('.column-body[data-col="'+ci+'"]').querySelectorAll('.stock-checkbox').forEach(s=>{s.checked=ck;const k=s.dataset.key;if(ck)APP.checkedItems.add(k);else APP.checkedItems.delete(k);s.closest('.stock-row').classList.toggle('selected',ck);});});});}

function mkSection(item,ci,ri){const d=document.createElement('div');d.className='section-header';d.draggable=true;d.dataset.col=ci;d.dataset.row=ri;d.dataset.type='section';d.innerHTML='<span class="section-label">📁 '+item.label+'</span><div class="section-actions"><button class="btn btn-small" data-a="ren">✏</button><button class="btn btn-small btn-danger" data-a="del">✕</button></div>';d.querySelector('[data-a="ren"]').addEventListener('click',e=>{e.stopPropagation();const n=prompt('セクション名:',item.label);if(n!==null){item.label=n;saveTabData(APP.currentTab);renderGrid();}});d.querySelector('[data-a="del"]').addEventListener('click',e=>{e.stopPropagation();if(!confirm('セクション「'+item.label+'」を削除？'))return;const cols=splitToColumns(APP.tabs[APP.currentTab],APP.COL_COUNT);cols[ci].splice(ri,1);APP.tabs[APP.currentTab]=flattenColumns(cols);saveTabData(APP.currentTab);renderGrid();});setupDrag(d,ci,ri);return d;}
function mkStockRow(item,ci,ri,rn,dups){const d=document.createElement('div');d.className='stock-row';d.draggable=true;d.dataset.col=ci;d.dataset.row=ri;d.dataset.type='stock';const key=APP.currentTab+'-'+ci+'-'+ri;if(APP.checkedItems.has(key))d.classList.add('selected');const nm=item.name||resolveName(item.code);const hasMemo=!!APP.stockMemos[item.code];const memoMarkHtml=hasMemo?'<span class="memo-mark" title="メモあり">●</span>':'';d.innerHTML='<span class="row-num">'+rn+'</span><span class="stock-code">'+item.code+'</span>'+memoMarkHtml+'<span class="stock-name'+(dups.has(item.code)?' duplicate':'')+'" title="'+nm+' — クリックでメモ編集">'+nm+'</span><input type="checkbox" class="stock-checkbox" data-key="'+key+'" '+(APP.checkedItems.has(key)?'checked':'')+'>';d.querySelector('.stock-code').addEventListener('click',e=>{e.stopPropagation();startCodeEdit(e.target,item);});d.querySelector('.stock-name').addEventListener('click',e=>{e.stopPropagation();openMemoModal(item.code,nm);});const cb=d.querySelector('.stock-checkbox');cb.addEventListener('change',e=>{e.stopPropagation();if(cb.checked){APP.checkedItems.add(key);d.classList.add('selected');}else{APP.checkedItems.delete(key);d.classList.remove('selected');}});setupDrag(d,ci,ri);return d;}
function startCodeEdit(el,item){if(el.querySelector('input'))return;const inp=document.createElement('input');inp.type='text';inp.value=item.code;inp.style.cssText='width:50px;font-family:var(--font-mono);font-size:12px;padding:1px;border:none;background:var(--bg-input);color:var(--text-bright);outline:1px solid var(--accent);border-radius:2px;';el.textContent='';el.appendChild(inp);inp.focus();inp.select();const fin=s=>{if(s&&inp.value.trim()){item.code=inp.value.trim().toUpperCase();item.name=resolveName(item.code);saveTabData(APP.currentTab);}renderGrid();};inp.addEventListener('blur',()=>fin(true));inp.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();fin(true);}if(e.key==='Escape')fin(false);});}

// === D&D ===
function setupDrag(el,ci,ri){el.addEventListener('dragstart',e=>{APP.dragState={colIdx:ci,rowIdx:ri};el.classList.add('dragging');e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain','');});el.addEventListener('dragend',()=>{el.classList.remove('dragging');clearDI();APP.dragState=null;});}
function clearDI(){document.querySelectorAll('.drag-target-above,.drag-target-below,.drag-over').forEach(e=>e.classList.remove('drag-target-above','drag-target-below','drag-over'));}
function setupDropZone(body,tci){body.addEventListener('dragover',e=>{e.preventDefault();e.dataTransfer.dropEffect='move';body.classList.add('drag-over');const rows=body.querySelectorAll('.stock-row,.section-header');rows.forEach(r=>r.classList.remove('drag-target-above','drag-target-below'));const cl=closestRow(e.clientY,rows);if(cl.el)cl.el.classList.add(cl.above?'drag-target-above':'drag-target-below');});body.addEventListener('dragleave',e=>{if(!body.contains(e.relatedTarget)){body.classList.remove('drag-over');body.querySelectorAll('.drag-target-above,.drag-target-below').forEach(e=>e.classList.remove('drag-target-above','drag-target-below'));}});body.addEventListener('drop',e=>{e.preventDefault();clearDI();if(!APP.dragState)return;const{colIdx:sc,rowIdx:sr}=APP.dragState;const cols=splitToColumns(APP.tabs[APP.currentTab],APP.COL_COUNT);const src=cols[sc]&&cols[sc][sr];if(!src)return;const rows=body.querySelectorAll('.stock-row,.section-header');const cl=closestRow(e.clientY,rows);let tr=cl.el?parseInt(cl.el.dataset.row)+(cl.above?0:1):cols[tci].length;cols[sc].splice(sr,1);if(sc===tci&&sr<tr)tr--;cols[tci].splice(tr,0,src);APP.tabs[APP.currentTab]=flattenColumns(cols);saveTabData(APP.currentTab);renderGrid();});}
function closestRow(y,rows){let b={el:null,distance:Infinity,above:true};rows.forEach(r=>{const rect=r.getBoundingClientRect();const mid=rect.top+rect.height/2;const d=Math.abs(y-mid);if(d<b.distance)b={el:r,distance:d,above:y<mid};});return b;}

// === ツールバー ===
document.getElementById('btn-add-stock').addEventListener('click',addStock);document.getElementById('add-stock-input').addEventListener('keydown',e=>{if(e.key==='Enter')addStock();});
function addStock(){const inp=document.getElementById('add-stock-input');const raw=inp.value.trim();if(!raw)return;const codes=raw.split(/[\s,;、]+/).filter(c=>c);const tc=parseInt(document.getElementById('input-column-select').value);const cols=splitToColumns(APP.tabs[APP.currentTab]||[],APP.COL_COUNT);for(const c of codes){const uc=c.toUpperCase();cols[tc].push({code:uc,name:resolveName(uc),_col:tc});}APP.tabs[APP.currentTab]=flattenColumns(cols);saveTabData(APP.currentTab);renderGrid();renderTabs();updateStockCount();inp.value='';showToast(codes.length+'銘柄追加','success');}
document.getElementById('btn-add-section').addEventListener('click',()=>{const inp=document.getElementById('add-section-input');const n=inp.value.trim();if(!n)return;const tc=parseInt(document.getElementById('input-column-select').value);const cols=splitToColumns(APP.tabs[APP.currentTab]||[],APP.COL_COUNT);cols[tc].unshift({type:'section',label:n,_col:tc});APP.tabs[APP.currentTab]=flattenColumns(cols);saveTabData(APP.currentTab);renderGrid();inp.value='';showToast('セクション追加','success');});
document.getElementById('btn-rename-tab').addEventListener('click',()=>{APP.tabNames[APP.currentTab]=document.getElementById('tab-name-input').value;saveSettings();renderTabs();showToast('タブ名変更','success');});
document.getElementById('btn-export-codes').addEventListener('click',()=>{const codes=getCheckedCodes();if(!codes.length){showToast('チェックなし','info');return;}document.getElementById('export-textarea').value=codes.join('\n');document.getElementById('export-modal').classList.remove('hidden');});
document.getElementById('btn-delete-checked').addEventListener('click',()=>{const keys=[...APP.checkedItems];if(!keys.length){showToast('チェックなし','info');return;}if(!confirm(keys.length+'件削除？'))return;const cols=splitToColumns(APP.tabs[APP.currentTab],APP.COL_COUNT);const del={};for(const k of keys){const p=k.split('-');if(!del[p[1]])del[p[1]]=[];del[p[1]].push(parseInt(p[2]));}for(const c of Object.keys(del)){del[c].sort((a,b)=>b-a);for(const r of del[c])if(cols[c]&&cols[c][r]&&cols[c][r].type!=='section')cols[c].splice(r,1);}APP.tabs[APP.currentTab]=flattenColumns(cols);APP.checkedItems.clear();saveTabData(APP.currentTab);renderGrid();renderTabs();updateStockCount();showToast(keys.length+'件削除','success');});
function getCheckedCodes(){const cols=splitToColumns(APP.tabs[APP.currentTab]||[],APP.COL_COUNT);return[...APP.checkedItems].map(k=>{const p=k.split('-');return cols[p[1]]&&cols[p[1]][parseInt(p[2])]&&cols[p[1]][parseInt(p[2])].code;}).filter(Boolean);}
document.getElementById('btn-move-up').addEventListener('click',()=>moveChecked(-1));document.getElementById('btn-move-down').addEventListener('click',()=>moveChecked(1));
function moveChecked(dir){if(!APP.checkedItems.size){showToast('チェックなし','info');return;}const cols=splitToColumns(APP.tabs[APP.currentTab]||[],APP.COL_COUNT);const bc={};for(const k of APP.checkedItems){const p=k.split('-');if(!bc[p[1]])bc[p[1]]=[];bc[p[1]].push(parseInt(p[2]));}APP.checkedItems.clear();for(const c of Object.keys(bc)){const rows=bc[c].sort((a,b)=>dir===-1?a-b:b-a);for(const r of rows){const nr=r+dir;if(nr<0||nr>=cols[c].length)continue;if(cols[c][nr]&&cols[c][nr].type==='section')continue;if(cols[c][r]&&cols[c][r].type==='section')continue;const tmp=cols[c][r];cols[c][r]=cols[c][nr];cols[c][nr]=tmp;APP.checkedItems.add(APP.currentTab+'-'+c+'-'+nr);}}APP.tabs[APP.currentTab]=flattenColumns(cols);saveTabData(APP.currentTab);renderGrid();}

// === 管理画面 ===
document.getElementById('btn-admin').addEventListener('click',()=>{document.getElementById('admin-modal').classList.remove('hidden');updateMasterCountDisplay();});
document.getElementById('btn-help').addEventListener('click',()=>{document.getElementById('help-modal').classList.remove('hidden');});
const xlsA=document.getElementById('xls-upload-area'),xlsI=document.getElementById('xls-file-input');xlsA.addEventListener('click',()=>xlsI.click());xlsA.addEventListener('dragover',e=>{e.preventDefault();xlsA.classList.add('drag-active');});xlsA.addEventListener('dragleave',()=>xlsA.classList.remove('drag-active'));xlsA.addEventListener('drop',e=>{e.preventDefault();xlsA.classList.remove('drag-active');if(e.dataTransfer.files.length)processXls(e.dataTransfer.files[0]);});xlsI.addEventListener('change',()=>{if(xlsI.files.length)processXls(xlsI.files[0]);});
async function processXls(file){const st=document.getElementById('xls-upload-status');st.innerHTML='<div class="status-msg" style="color:var(--accent);border:1px solid var(--accent);background:rgba(0,180,216,0.1)">⏳ 読み込み中...</div>';try{const data=await file.arrayBuffer();const wb=XLSX.read(data,{type:'array'});const json=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1});let cc=-1,nc=-1;for(let r=0;r<Math.min(5,json.length);r++){const row=json[r];if(!row)continue;for(let c=0;c<row.length;c++){const v=String(row[c]||'').trim();if(/コード|Code/i.test(v))cc=c;if(/銘柄名|銘柄|会社名|Name/i.test(v))nc=c;}if(cc>=0)break;}if(cc<0){for(let r=1;r<Math.min(10,json.length);r++){const row=json[r];if(!row)continue;for(let c=0;c<row.length;c++){if(String(row[c]||'').trim().length>=3){cc=c;nc=c+1;break;}}if(cc>=0)break;}}if(cc<0){st.innerHTML='<div class="status-msg error">❌ コード列を検出できません</div>';return;}const master={};for(let r=1;r<json.length;r++){const row=json[r];if(!row)continue;let code=String(row[cc]||'').trim(),name=String(row[nc]||'').trim().replace(/\u3000/g,' ');if(code&&name)master[code]=name;}const cnt=Object.keys(master).length;if(!cnt){st.innerHTML='<div class="status-msg error">❌ データなし</div>';return;}APP.customMaster=master;const preset=generatePresetTabs(APP.COL_COUNT,master);for(let t=18;t<=30;t++){APP.tabs[t]=preset.tabs[t]||[];APP.tabNames[t]=preset.names[t];}try{await db.collection('users').doc(APP.uid).set({customMaster:master,tabNames:APP.tabNames},{merge:true});}catch(e){console.warn('マスタ保存:',e.message);}st.innerHTML='<div class="status-msg success">✅ '+cnt+'銘柄インポート完了！</div>';updateMasterCountDisplay();renderTabs();if(APP.currentTab>=18)renderGrid();showToast(cnt+'銘柄インポート','success');}catch(e){st.innerHTML='<div class="status-msg error">❌ '+e.message+'</div>';}}
const csvA=document.getElementById('csv-upload-area'),csvI=document.getElementById('csv-file-input');csvA.addEventListener('click',()=>csvI.click());csvA.addEventListener('dragover',e=>{e.preventDefault();csvA.classList.add('drag-active');});csvA.addEventListener('dragleave',()=>csvA.classList.remove('drag-active'));csvA.addEventListener('drop',e=>{e.preventDefault();csvA.classList.remove('drag-active');if(e.dataTransfer.files.length)processCsv(e.dataTransfer.files[0]);});csvI.addEventListener('change',()=>{if(csvI.files.length)processCsv(csvI.files[0]);});
async function processCsv(file){const st=document.getElementById('csv-upload-status');try{const t=await file.text();const codes=t.split(/[\r\n,\t;]+/).map(s=>s.trim().replace(/"/g,'')).filter(c=>c.length>=3);st.innerHTML=codes.length?'<div class="status-msg success">✅ '+codes.length+'銘柄読み込み</div>':'<div class="status-msg error">❌ 有効なコードなし</div>';}catch(e){st.innerHTML='<div class="status-msg error">❌ '+e.message+'</div>';}}
document.getElementById('btn-regen-bandai').addEventListener('click',()=>{const preset=generatePresetTabs(APP.COL_COUNT,APP.customMaster);for(let t=18;t<=30;t++){APP.tabs[t]=preset.tabs[t]||[];APP.tabNames[t]=preset.names[t];}renderTabs();if(APP.currentTab>=18)renderGrid();showToast('番台タブ再生成完了','success');});
document.getElementById('btn-reset-presets').addEventListener('click',()=>{const preset=generatePresetTabs(APP.COL_COUNT,APP.customMaster);for(let t=18;t<=30;t++){APP.tabs[t]=preset.tabs[t]||[];APP.tabNames[t]=preset.names[t];}renderTabs();if(APP.currentTab>=18)renderGrid();showToast('プリセット初期化完了','success');});
function updateMasterCountDisplay(){const el=document.getElementById('master-count-display');if(el)el.textContent='現在のマスタ: '+Object.keys(getMaster()).length+'銘柄';}
document.addEventListener('keydown',e=>{if(e.key==='Escape')document.querySelectorAll('.modal-overlay:not(.hidden)').forEach(m=>m.classList.add('hidden'));});
console.log('v5 loaded — '+Object.keys(STOCK_MASTER).length+' stocks');

// === ダッシュボード ===
document.getElementById('btn-dashboard').addEventListener('click',()=>{
  const grid=document.getElementById('dashboard-grid');
  grid.innerHTML='';
  for(let t=1;t<=30;t++){
    const btn=document.createElement('div');
    const name=APP.tabNames[t]||'タブ'+t;
    const cnt=(APP.tabs[t]||[]).filter(r=>r.type!=='section').length;
    btn.style.cssText='padding:8px 12px;background:var(--bg-card);border:1px solid var(--border);border-radius:4px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;transition:all 0.15s;';
    btn.innerHTML='<span style="color:var(--text-bright);font-size:13px;">'+name+'</span><span style="color:var(--text-dim);font-size:11px;">'+cnt+'銘柄</span>';
    btn.addEventListener('mouseenter',()=>{btn.style.borderColor='var(--accent)';btn.style.background='var(--bg-hover)';});
    btn.addEventListener('mouseleave',()=>{btn.style.borderColor='var(--border)';btn.style.background='var(--bg-card)';});
    btn.addEventListener('click',()=>{switchTab(t);document.getElementById('dashboard-modal').classList.add('hidden');});
    grid.appendChild(btn);
  }
  document.getElementById('dashboard-modal').classList.remove('hidden');
});

// === CSV取り込み（タブ1-15） ===
document.getElementById('btn-csv-import').addEventListener('click',()=>{
  if(APP.currentTab>17){showToast('CSV取り込みはタブ1〜17のみ対応です','error');return;}
  document.getElementById('csv-tab-upload-status').innerHTML='';
  document.getElementById('csv-import-modal').classList.remove('hidden');
});

const csvTA=document.getElementById('csv-tab-upload-area'),csvTI=document.getElementById('csv-tab-file-input');
csvTA.addEventListener('click',()=>csvTI.click());
csvTA.addEventListener('dragover',e=>{e.preventDefault();csvTA.classList.add('drag-active');});
csvTA.addEventListener('dragleave',()=>csvTA.classList.remove('drag-active'));
csvTA.addEventListener('drop',e=>{e.preventDefault();csvTA.classList.remove('drag-active');if(e.dataTransfer.files.length)processCsvTab(e.dataTransfer.files[0]);});
csvTI.addEventListener('change',()=>{if(csvTI.files.length)processCsvTab(csvTI.files[0]);csvTI.value='';});

function processCsvTab(file){
  const st=document.getElementById('csv-tab-upload-status');
  const tabId=APP.currentTab;
  if(tabId>15){st.innerHTML='<div class="status-msg error">❌ タブ1〜15のみ対応</div>';return;}
  const reader=new FileReader();
  reader.onload=function(e){
    const text=e.target.result;
    const lines=text.split(/[\r\n]+/).map(s=>s.trim()).filter(s=>s&&!s.startsWith('#'));
    const codes=lines.filter(c=>/^[\dA-Za-z]{3,5}$/.test(c)).map(c=>c.toUpperCase());
    if(!codes.length){st.innerHTML='<div class="status-msg error">❌ 有効な銘柄コードが見つかりません</div>';return;}

    // セクション名 = ファイル名（拡張子除く）
    const sectionName=file.name.replace(/\.[^.]+$/,'');
    const tabData=APP.tabs[tabId]||[];
    const existingCount=tabData.filter(r=>r.type!=='section').length;
    const MAX=300;
    const remaining=MAX-existingCount;
    if(remaining<=0){st.innerHTML='<div class="status-msg error">❌ このタブは既に300銘柄に達しています</div>';return;}

    const toAdd=codes.slice(0,remaining);
    const cols=splitToColumns(tabData,APP.COL_COUNT);

    // 空きがある列を探して順に追加
    // まず既存データの末尾位置を把握
    let totalAdded=0;
    const perCol=50; // 1列最大50

    // セクションヘッダを全列に追加
    for(let c=0;c<APP.COL_COUNT;c++){
      cols[c].push({type:'section',label:sectionName,_col:c});
    }

    // 銘柄を列1→2→3→4→5→6の順に追加
    for(const code of toAdd){
      // 各列の銘柄数を数えて一番少ない列に入れる（左から優先）
      let targetCol=0;
      let minCount=Infinity;
      for(let c=0;c<APP.COL_COUNT;c++){
        const stockCount=cols[c].filter(r=>r.type!=='section').length;
        if(stockCount<minCount){minCount=stockCount;targetCol=c;}
      }
      // 1列50超えたら次の列
      if(minCount>=perCol){
        // 全列50超え→追加不可
        break;
      }
      cols[targetCol].push({code:code,name:resolveName(code),_col:targetCol});
      totalAdded++;
    }

    APP.tabs[tabId]=flattenColumns(cols);
    saveTabData(tabId);
    renderGrid();renderTabs();updateStockCount();

    const skipped=codes.length-totalAdded;
    let msg='✅ '+totalAdded+'銘柄を取り込みました（セクション: '+sectionName+'）';
    if(skipped>0)msg+=' ※'+skipped+'銘柄は上限超過のため取り込まれませんでした';
    st.innerHTML='<div class="status-msg success">'+msg+'</div>';
    showToast(totalAdded+'銘柄取り込み完了','success');
  };
  reader.readAsText(file,'UTF-8');
}

// ==================================================================
// === 銘柄詳細メモモーダル（機能N）==================================
// ==================================================================

// 東証33業種分類（プルダウン選択肢）
const INDUSTRY_LIST = [
  "水産・農林業","鉱業","建設業","食料品","繊維製品","パルプ・紙","化学","医薬品",
  "石油・石炭製品","ゴム製品","ガラス・土石製品","鉄鋼","非鉄金属","金属製品","機械",
  "電気機器","輸送用機器","精密機器","その他製品","電気・ガス業","陸運業","海運業",
  "空運業","倉庫・運輸関連業","情報・通信業","卸売業","小売業","銀行業","証券・商品先物取引業",
  "保険業","その他金融業","不動産業","サービス業"
];

// 業種プルダウン初期化
(function initIndustryOptions(){
  const sel=document.getElementById('memo-industry');
  if(!sel)return;
  for(const ind of INDUSTRY_LIST){
    const opt=document.createElement('option');
    opt.value=ind;opt.textContent=ind;sel.appendChild(opt);
  }
})();

// モーダルを開く
function openMemoModal(code,displayName){
  APP.currentMemoCode=code;
  document.getElementById('memo-code').textContent=code;
  document.getElementById('memo-name').textContent=displayName;

  const existing=APP.stockMemos[code]||{};
  document.getElementById('memo-price').value=(existing.price!==undefined&&existing.price!==null)?existing.price:'';
  document.getElementById('memo-industry').value=existing.industry||'';
  APP.currentMemoThemes=Array.isArray(existing.themes)?[...existing.themes]:[];
  renderMemoThemes();
  document.getElementById('memo-reason').value=existing.reason||'';
  document.getElementById('memo-note').value=existing.note||'';
  updateCharCount('memo-reason','memo-reason-count');
  updateCharCount('memo-note','memo-note-count');

  // 既存メモあり → 削除ボタン表示 / なし → 非表示
  document.getElementById('memo-delete').style.display=APP.stockMemos[code]?'inline-flex':'none';

  document.getElementById('memo-theme-input').value='';
  document.getElementById('memo-modal').classList.remove('hidden');
  setTimeout(()=>document.getElementById('memo-price').focus(),80);
}

// テーマタグ描画
function renderMemoThemes(){
  const box=document.getElementById('memo-theme-tags');
  box.innerHTML='';
  APP.currentMemoThemes.forEach((t,i)=>{
    const tag=document.createElement('span');
    tag.className='theme-tag';
    tag.textContent=t;
    const rm=document.createElement('button');
    rm.type='button';rm.textContent='×';rm.title='削除';
    rm.addEventListener('click',()=>{APP.currentMemoThemes.splice(i,1);renderMemoThemes();});
    tag.appendChild(rm);
    box.appendChild(tag);
  });
}

// テーマ追加
function addThemeFromInput(){
  const inp=document.getElementById('memo-theme-input');
  const v=inp.value.trim();
  if(!v)return;
  if(APP.currentMemoThemes.includes(v)){showToast('同じテーマが既に登録されています','info');return;}
  if(APP.currentMemoThemes.length>=20){showToast('テーマは20個までです','error');return;}
  APP.currentMemoThemes.push(v);
  inp.value='';
  renderMemoThemes();
}
document.getElementById('memo-theme-add').addEventListener('click',addThemeFromInput);
document.getElementById('memo-theme-input').addEventListener('keydown',e=>{
  if(e.key==='Enter'){e.preventDefault();addThemeFromInput();}
});

// 文字数カウント
function updateCharCount(textareaId,countElId){
  const el=document.getElementById(textareaId);
  const cnt=document.getElementById(countElId);
  const len=el.value.length;
  cnt.textContent=len+' / 100';
  cnt.classList.toggle('over',len>100);
}
document.getElementById('memo-reason').addEventListener('input',()=>updateCharCount('memo-reason','memo-reason-count'));
document.getElementById('memo-note').addEventListener('input',()=>updateCharCount('memo-note','memo-note-count'));

// 保存
document.getElementById('memo-save').addEventListener('click',async()=>{
  const code=APP.currentMemoCode;
  if(!code){showToast('銘柄コードが不正です','error');return;}
  const reason=document.getElementById('memo-reason').value.trim();
  const note=document.getElementById('memo-note').value.trim();
  if(reason.length>100||note.length>100){showToast('100文字以内で入力してください','error');return;}
  const priceRaw=document.getElementById('memo-price').value.trim();
  const price=priceRaw===''?null:Number(priceRaw);
  if(priceRaw!==''&&Number.isNaN(price)){showToast('株価は数値で入力してください','error');return;}
  const data={
    price:price,
    industry:document.getElementById('memo-industry').value||'',
    themes:[...APP.currentMemoThemes],
    reason:reason,
    note:note,
    updatedAt:firebase.firestore.FieldValue.serverTimestamp()
  };
  try{
    await db.collection('users').doc(APP.uid).collection('stockMemos').doc(code).set(data);
    APP.stockMemos[code]={...data,updatedAt:new Date()};
    showToast('メモを保存しました','success');
    document.getElementById('memo-modal').classList.add('hidden');
    renderGrid(); // メモ済みマーク更新
  }catch(e){
    console.error(e);
    showToast('保存に失敗: '+e.message,'error');
  }
});

// 削除
document.getElementById('memo-delete').addEventListener('click',async()=>{
  const code=APP.currentMemoCode;
  if(!code)return;
  if(!confirm(code+' のメモを削除しますか？'))return;
  try{
    await db.collection('users').doc(APP.uid).collection('stockMemos').doc(code).delete();
    delete APP.stockMemos[code];
    showToast('メモを削除しました','success');
    document.getElementById('memo-modal').classList.add('hidden');
    renderGrid();
  }catch(e){
    console.error(e);
    showToast('削除に失敗: '+e.message,'error');
  }
});

console.log('v6.1 loaded — '+Object.keys(STOCK_MASTER).length+' stocks, memo feature enabled');
