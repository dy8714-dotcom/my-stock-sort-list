/**
 * app.js v5 — マイ銘柄ソートリスト
 * タブ1-17: Firestoreから読み込み（ユーザーデータ）
 * タブ18-30: 毎回stock-data.jsから生成（常に完全・最新）→Firestoreにも保存
 */
const firebaseConfig={apiKey:"AIzaSyCXpU-IX55WPJCVlhEhbcOZQzPmVKZt9PU",authDomain:"my-stock-list-7ae4d.firebaseapp.com",projectId:"my-stock-list-7ae4d",storageBucket:"my-stock-list-7ae4d.firebasestorage.app",messagingSenderId:"1072316134172",appId:"1:1072316134172:web:27abd8b3b8906afa2480d9"};
firebase.initializeApp(firebaseConfig);
const auth=firebase.auth(),db=firebase.firestore(),provider=new firebase.auth.GoogleAuthProvider();
const APP={uid:null,currentTab:1,tabs:{},tabNames:{},customMaster:null,checkedItems:new Set(),dragState:null,COL_COUNT:6};

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
function getDefaultTabNames(){const n={};for(let i=1;i<=17;i++)n[i]='';n[18]='JPX150＆TOPIX100';n[19]='JPX400';n[20]='日経225＆TOPIX100';for(let b=1;b<=9;b++)n[20+b]=b+'000番台';n[30]='3桁＋Aコード';return n;}

async function loadAllData(){
  if(!APP.uid)return;
  APP.tabNames=getDefaultTabNames();
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
function mkStockRow(item,ci,ri,rn,dups){const d=document.createElement('div');d.className='stock-row';d.draggable=true;d.dataset.col=ci;d.dataset.row=ri;d.dataset.type='stock';const key=APP.currentTab+'-'+ci+'-'+ri;if(APP.checkedItems.has(key))d.classList.add('selected');const nm=item.name||resolveName(item.code);d.innerHTML='<span class="row-num">'+rn+'</span><span class="stock-code">'+item.code+'</span><span class="stock-name'+(dups.has(item.code)?' duplicate':'')+'" title="'+nm+'">'+nm+'</span><input type="checkbox" class="stock-checkbox" data-key="'+key+'" '+(APP.checkedItems.has(key)?'checked':'')+'>';d.querySelector('.stock-code').addEventListener('click',e=>{e.stopPropagation();startCodeEdit(e.target,item);});const cb=d.querySelector('.stock-checkbox');cb.addEventListener('change',e=>{e.stopPropagation();if(cb.checked){APP.checkedItems.add(key);d.classList.add('selected');}else{APP.checkedItems.delete(key);d.classList.remove('selected');}});setupDrag(d,ci,ri);return d;}
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
