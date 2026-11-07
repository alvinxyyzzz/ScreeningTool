import {importWorkbook,validateData} from './excel.js';
import {createModel,screenFacility,librarySearch,lines,toCSV} from './model.js';
const $=selector=>document.querySelector(selector);
const ICON=n=>`<svg aria-hidden="true"><use href="#icon-${n}"/></svg>`;
const safe=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const count=v=>Number(v||0).toLocaleString('en-US');
const clip=(s,n=78)=>String(s||'').length>n?String(s).slice(0,n)+'…':String(s||'');
const initial={page:'overview',data:null,previous:null,model:null,pending:null,screenInputs:{name:'',higg:'',country:'',brand:''},screenResults:null,screenBatch:null,library:{query:'',country:'',brand:'',program:'',onlyStrategic:false,onlyRisk:false,includeExternal:false}};
const state=initial;
const titles={overview:'總覽',screen:'工廠篩查',library:'合作資料庫',quality:'資料健康度',updates:'資料更新'};
let toastTimer=null;
function toast(message,stay=false){const node=$('#toast');node.hidden=false;node.textContent=message;clearTimeout(toastTimer);if(!stay)toastTimer=setTimeout(()=>node.hidden=true,4800)}
function sectionHead(title,description='',action=''){
 return `<div class="section-head"><div><h2>${safe(title)}</h2>${description?`<p>${safe(description)}</p>`:''}</div>${action}</div>`;
}
function metric(title,value,foot,iconName,warning=false){return `<div class="metric ${warning?'metric-warning':''}"><div class="metric-top"><span>${safe(title)}</span><div class="metric-icon">${ICON(iconName)}</div></div><div class="metric-value mono">${count(value)}</div><div class="metric-foot">${safe(foot)}</div></div>`}
function badge(label,kind=''){return `<span class="tag ${kind}">${safe(label)}</span>`}
function programBadges(f){return f.programs.slice(0,4).map(p=>badge(p,p==='CTS'?'green':p==='CTM'?'blue':'')).join('')||(f.strategics.length?badge('Strategic only','blue'):f.assessments.length?badge('Assessment only','blue'):f.parents.length?badge('Manufacturer only','gray'):badge('來源紀錄','gray'))}
function pageHead(kicker,title,desc='',right=''){return `<div class="page-head"><div><div class="page-kicker">${safe(kicker)}</div><h1 class="page-title">${title}</h1>${desc?`<p class="page-desc">${desc}</p>`:''}</div>${right}</div>`}
function buttonAction(action,label,kind='outline-btn',iconName=''){return `<button class="${kind}" data-action="${action}">${iconName?ICON(iconName):''}${safe(label)}</button>`}
function startImport(){ $('#file-picker').click(); }
function shortStatus(f){return f.external?badge('Source only','blue'):f.risks.length?badge('需複核','orange'):f.strategics.length?badge('Strategic','green'):badge('已收錄','gray')}
function renderEmptyLanding(){
 return `${pageHead('FACTORY INTELLIGENCE · PRIVATE WORKSPACE','你的工廠資料，<span class="serif-em">一眼看懂。</span>','把分散的歷史合作、品牌、工廠與戰略名單整理成真正可搜尋、可追溯的工作台。')}
 <section class="welcome-hero"><div class="hero-eyebrow">YOUR DATABASE. YOUR DEVICE.</div><h2>From spreadsheets<br/>to <span style="color:#b2f774">clear decisions.</span></h2><p>直接匯入你每月更新的 Excel .xlsm 檔案。不需要先轉 JSON，也不會上傳到任何伺服器。完整搜尋、篩查、來源明細與異常提醒都留在你的裝置。</p><button class="mint-btn" data-action="pick-file">${ICON('upload')} 匯入 Screening Tool Excel ${ICON('arrow')}</button></section>
 <div class="onboarding-grid"><div class="card"><div class="step-badge">01 / IMPORT</div><h3>匯入 Excel</h3><p>支援你現有的 .xlsm 與私人 JSON 備份。不執行任何巨集。</p></div><div class="card"><div class="step-badge">02 / SCREEN</div><h3>查詢工廠關聯</h3><p>比對 Higg ID、工廠名稱及國家，查閱 CTA、CTS、CTM 等合作紀錄。</p></div><div class="card"><div class="step-badge">03 / REVIEW</div><h3>掌握資料風險</h3><p>提示多重 Higg ID、跨國合併及來源資料缺口，避免誤判。</p></div></div>`;
}
function renderOverview(){
 if(!state.model)return renderEmptyLanding();
 const m=state.model,st=m.stats;
 const recent=[...m.facilities].sort((a,b)=>b.engagements.length-a.engagements.length).slice(0,6);
 const source=Object.entries(m.sourceCounts).sort((a,b)=>b[1]-a[1]).slice(0,6),max=Math.max(...source.map(x=>x[1]),1);
 return `${pageHead('YOUR DATA AT A GLANCE','工廠合作情報，<span class="serif-em">盡在掌握。</span>','由你的 Screening Tool 資料庫建立 · 每個結果都能追溯到原始來源。',buttonAction('pick-file','更新 Excel','outline-btn','upload'))}
 <section class="hero"><div><div class="hero-eyebrow">QUICK SCREENING / 01</div><h2>從一間工廠開始，<br/><span>找到完整合作脈絡。</span></h2><p>輸入工廠名稱或 Higg ID，探索過往品牌、服務、評估與戰略名單。</p><form id="hero-form" class="search-wide">${ICON('search')}<input aria-label="搜尋工廠" name="quick-search" placeholder="例如：Envoy Textiles · 42555" autocomplete="off"/><button class="dark-btn" type="submit">開始篩查 ${ICON('arrow')}</button></form></div><div class="hero-art" aria-hidden="true">${ICON('layers')}<span class="orbit"></span></div></section>
 <section class="section grid-4">${metric('MASTER FACTORIES',st.facilities,'已整併的工廠主紀錄','database')}${metric('ENGAGEMENT RECORDS',st.engagements,'含多種服務及執行階段','layers')}${metric('STRATEGIC ENTRIES',st.strategic,'戰略名單來源紀錄','spark')}${metric('REVIEW FLAGS',st.risks,'潛在異常 Master 合併','alert',true)}</section>
 <section class="section split-layout"><div class="card card-pad">${sectionHead('合作紀錄較多的工廠','可點選查看各品牌與專案來源',`<button data-view="library">查看全部 ${ICON('arrow')}</button>`)}<div class="table-scroll"><table class="simple-table"><thead><tr><th>工廠名稱</th><th>國家</th><th>紀錄</th></tr></thead><tbody>${recent.map(f=>`<tr class="row-click" data-action="detail" data-id="${safe(f.id)}"><td><div class="table-name">${safe(clip(f.name,40))}</div><div class="table-country">Higg ${safe(f.higgs.join(', ')||'—')}</div></td><td class="table-country">${safe(lines(f.country)[0]||'—')}</td><td>${badge(f.engagements.length+' entries','green')}</td></tr>`).join('')}</tbody></table></div></div><div class="card card-pad">${sectionHead('資料來源分布','資料來自不同歷史 tracker 與計畫')}${source.map(([name,n])=>`<div class="bar-row"><span class="bar-label" title="${safe(name)}">${safe(name.replace(' database','').replace(' tracker_',' · '))}</span><span class="bar-track"><span class="bar-fill" style="display:block;width:${Math.max(2,n/max*100)}%"></span></span><span class="bar-num">${count(n)}</span></div>`).join('')}<div class="meta-foot">此處是來源列數，並非獨立工廠數或減碳成效。</div></div></section>`;
}
function screenResultItem(r){const f=r.facility;
 return `<article class="card result-card" data-action="detail" data-id="${safe(f.id)}" tabindex="0" role="button"><div><div class="result-top">${badge(r.tier==='exact'?'高度匹配':r.tier==='probable'?'相似匹配':'建議複核',`confidence-${r.tier}`)}<span>${safe(r.reason)}</span>${f.external?badge('非合作 Master','blue'):''}</div><h3 class="result-name">${safe(f.name)}</h3><div class="result-mid"><span>${safe(lines(f.country)[0]||'國家未知')}</span><span>·</span><span>Higg ${safe(f.higgs.join(' / ')||'—')}</span>${f.risks.length?badge('需複核','orange'):''}</div><div class="result-mid" style="margin-top:10px">${programBadges(f)}</div></div><div class="result-meta">${ICON('chevron')}<div class="result-count">${f.engagements.length} 筆合作<br/>${f.strategics.length} 筆戰略</div></div></article>`;
}
function screenResultsHTML(){
 if(state.screenResults===null)return `<div class="card empty-block"><div class="empty-symbol">${ICON('search')}</div><h3>準備好你的第一筆篩查</h3><p>建議優先用 Higg ID，比單靠工廠名稱更可靠；若只輸入名稱，會明確標示匹配可信度。</p></div>`;
 if(!state.screenResults.length)return `<div class="card empty-block"><div class="empty-symbol">${ICON('alert')}</div><h3>這次沒有找到可信的匹配</h3><p>不代表工廠一定沒有合作紀錄。請嘗試英文別名、Higg ID，或放寬國家條件；來源資料可能尚未完整收錄。</p></div>`;
 return `<div class="inline-note" style="margin-bottom:12px">已找到 ${state.screenResults.length} 筆候選。匹配標籤僅表示查詢相似度，並不保證同一法人或同一廠址；請點開來源明細確認。</div><div class="results-list">${state.screenResults.map(screenResultItem).join('')}</div>`;
}
function renderScreen(){
 if(!state.model)return renderNoData('篩查前先匯入資料','工廠篩查需要你本機的 Excel 資料庫。');
 const v=state.screenInputs;
 return `${pageHead('MATCH WITH CONFIDENCE','Factory <span class="serif-em">Screening</span>','支援 Higg ID 精確查詢、工廠名稱與國家組合比對；保留原始合作紀錄及資訊來源。')}
 <form class="card screen-form" id="screen-form"><div class="form-grid"><div class="field"><label for="screen-name">工廠名稱 / Facility</label><input id="screen-name" name="name" value="${safe(v.name)}" placeholder="輸入工廠名稱" autocomplete="off"/></div><div class="field"><label for="screen-higg">Higg ID · 優先</label><input id="screen-higg" name="higg" inputmode="numeric" value="${safe(v.higg)}" placeholder="例如 42555"/></div><div class="field"><label for="screen-country">國家 / Country</label><input id="screen-country" name="country" list="country-options" value="${safe(v.country)}" placeholder="可選填"/></div><div class="field"><label for="screen-brand">品牌 / Brand</label><input id="screen-brand" name="brand" value="${safe(v.brand)}" placeholder="可選填"/></div></div><div class="form-actions"><div class="form-help">只要有 Higg ID 或工廠名稱即可查詢。<br/>名稱匹配不會自動視為身份確認。</div><button class="dark-btn" type="submit">${ICON('search')} 執行篩查</button></div></form>
 <datalist id="country-options">${state.model.countryOptions.map(c=>`<option value="${safe(c)}"></option>`).join('')}</datalist>
 <section class="section">${sectionHead('Screening results',state.screenResults?`符合條件的候選：${state.screenResults.length} 筆`:'等待篩查輸入',state.screenResults?.length?`<button data-action="export-screen">${ICON('download')} 匯出結果</button>`:'')}${screenResultsHTML()}</section>
 <section class="section"><div class="card card-pad"><div class="section-head"><div><h2>Batch screening · 多筆篩查</h2><p>貼上多行工廠資料，快速取得摘要並匯出 CSV。每行格式：工廠名稱, 國家, Higg ID</p></div></div><div class="field"><textarea id="batch-input" placeholder="Envoy Textiles, Bangladesh, 42555&#10;PT. Gistex, Indonesia, &#10;SQ Hues, Bangladesh, 22692"></textarea></div><div class="form-actions"><span class="form-help">逗號分隔；每次最多 200 筆。不會傳到伺服器。</span><button class="soft-btn" data-action="batch">${ICON('layers')} 批次篩查</button></div>${state.screenBatch?`<div class="inline-note" style="margin-top:15px">已分析 ${state.screenBatch.length} 筆輸入，其中 ${state.screenBatch.filter(x=>x.match).length} 筆有候選。<button class="outline-btn" style="margin-top:10px" data-action="export-batch">${ICON('download')} 匯出結果 CSV</button></div>`:''}</div></section>`;
}
function renderNoData(title,desc){return `${pageHead('GET STARTED',title,desc)}<div class="card empty-block"><div class="empty-symbol">${ICON('database')}</div><h3>尚未連結你的資料庫</h3><p>匯入新版 Screening Tool .xlsm 或先前匯出的私有 JSON，即可使用完整功能。</p>${buttonAction('pick-file','匯入 Excel','dark-btn','upload')}</div>`}
function libraryTable(){
 const list=librarySearch(state.model,state.library);return `<div class="table-scroll"><table class="simple-table"><thead><tr><th>FACTORY / HIGG ID</th><th>COUNTRY</th><th>BRANDS</th><th>PROGRAMS</th><th>RECORDS</th><th>FLAG</th></tr></thead><tbody>${list.map(f=>`<tr class="row-click" data-action="detail" data-id="${safe(f.id)}"><td><div class="table-name">${safe(f.name)}</div><div class="table-country">${safe(f.higgs.join(', ')||'No Higg ID')}${f.external?' · Source only':''}</div></td><td class="table-country">${safe(lines(f.country).join(', ')||'—')}</td><td class="table-country">${safe(clip(f.brands.join(', '),41)||'—')}</td><td>${programBadges(f)}</td><td class="mono">${f.engagements.length}</td><td>${f.risks.length?badge('Review','orange'):shortStatus(f)}</td></tr>`).join('')}</tbody></table></div><div class="pagination-note">顯示 ${list.length} 筆／${count(state.model.facilities.length)} 筆 · 最多顯示前 200 筆，請以搜尋與篩選縮小範圍。</div>`;
}
function renderLibrary(){
 if(!state.model)return renderNoData('合作資料庫','匯入 Excel 後可搜尋所有 Master List 工廠與歷史合作紀錄。');
 const l=state.library;
 const opt=(values,current,placeholder)=>`<option value="">${placeholder}</option>`+values.map(x=>`<option value="${safe(x)}" ${current===x?'selected':''}>${safe(x)}</option>`).join('');
 return `${pageHead('THE SOURCE OF TRUTH','Factory <span class="serif-em">Explorer</span>','查工廠、找品牌、篩選服務，並回到原始來源確認每一筆合作歷史。')}
 <div class="filters-bar"><input id="library-query" class="search-field" type="search" placeholder="搜尋工廠、Higg ID、品牌或母公司…" value="${safe(l.query)}" aria-label="搜尋資料庫"/>
 <select id="library-country" aria-label="篩選國家">${opt(state.model.countryOptions,l.country,'全部國家')}</select><select id="library-program" aria-label="篩選服務">${opt(['CTA','CTS','BPAP','CTM','Other','Implementation'],l.program,'全部服務')}</select></div>
 <div class="filters-bar"><label class="checkline"><input type="checkbox" id="include-external" ${l.includeExternal?'checked':''}/>包含尚未連結 Master 的來源工廠 (${count(state.model.stats.external)})</label><label class="checkline"><input type="checkbox" id="only-strategic" ${l.onlyStrategic?'checked':''}/>只看 Strategic List 命中</label><label class="checkline"><input type="checkbox" id="only-risk" ${l.onlyRisk?'checked':''}/>只看需複核工廠</label></div>
 <div id="library-results" class="card list-card">${libraryTable()}</div>`;
}
function renderQuality(){
 if(!state.model)return renderNoData('資料健康度','匯入 Excel 後會自動偵測可能需要重新確認的 Master 合併紀錄。');
 const m=state.model,dup=[...m.byHigg.entries()].filter(([h,s])=>s.size>1);
 return `${pageHead('DATA INTEGRITY / REVIEW QUEUE','資料品質，<span class="serif-em">看得見。</span>','Excel 可以繼續由你維護；App 幫你發現可能影響篩查結論的關聯問題。')}
 <div class="grid-4">${metric('NEEDS REVIEW',m.risks.length,'疑似多 Higg ID 或跨國合併','alert',true)}${metric('DUPLICATE IDs',dup.length,'相同 Higg ID 出現不同 Master','layers',true)}${metric('UNLINKED STRATEGIC',m.orphans.strategics,'未安全連結至 Master','filter')}${metric('UNLINKED ASSESSMENTS',m.orphans.assessments,'未安全連結至 Master','shield')}</div>
 <section class="section"><div class="warn-inline">重要：<strong>提醒不等於確認有誤</strong>。一間製造商可能有多個廠址或 Higg ID，部分合併也可能是你刻意設計。App 會保守保留原始資料，不自動拆分、覆寫或替你決定正確的 Master。</div></section>
 <section class="section card">${sectionHead('Review queue / Master 合併複核','按項目查看姓名別名、國家與 Higg ID',`<span class="tag orange">${m.risks.length} records</span>`).replace('class="section-head"','class="section-head" style="padding:21px 20px 0"')}
 ${m.risks.length?m.risks.slice(0,120).map(f=>`<div class="quality-row" data-action="detail" data-id="${safe(f.id)}"><span class="risk-dot"></span><div style="flex:1;min-width:0"><h3>${safe(f.name)}</h3><p>MASTER ${safe(f.id)} · Higg ${safe(f.higgs.join(' / ')||'—')} · ${safe(lines(f.country).join(' / '))}</p><div class="risk-reasons">${safe(f.risks.join(' · '))}</div></div>${ICON('chevron')}</div>`).join(''):`<div class="empty-block">沒有偵測到指定類型的合併異常。</div>`}</section>
 <div class="meta-foot">檢查規則：單一 Master 多個 Higg ID、單一 Master 多國家、相同 Higg ID 落入不同 Master。輔助來源只在唯一識別條件符合時自動關聯，未關聯數量不代表資料錯誤。</div>`;
}
function stampsHTML(stamps={}){return Object.entries(stamps).map(([k,arr])=>`<div class="source-row"><strong>${safe(k.replace('Input - ',''))}</strong><span>${safe((arr||[]).join(', ')||'未標示日期')}</span></div>`).join('')}
function updatesPreview(){
 const p=state.pending;if(!p)return '';
 const m=state.data,compare=[['工廠 Master','master'],['合作紀錄','engagements'],['戰略名單','strategic'],['Assessments','assessments']];
 return `<section class="section card card-pad"><div class="page-kicker">READY TO APPLY / IMPORT PREVIEW</div><h2 style="margin:0">新資料已讀取，尚未覆寫</h2><p class="page-desc">${safe(p.meta?.name||'私人資料備份')} · 請檢查差異後，再確認更新。</p>
 <div class="diff">${compare.map(([label,key])=>`<div class="diff-cell"><strong>${count(p[key]?.length)}</strong><span>${safe(label)}</span>${m?`<em>${p[key].length-m[key].length>=0?'+':''}${count(p[key].length-m[key].length)} vs current</em>`:''}</div>`).join('')}</div>
 <div class="warn-inline">資料匯入使用儲存於 Excel 的公式計算結果，不會重新執行 Excel 公式或 VBA。若你剛更新過原始輸入，請先在桌面版 Excel 完整重新計算並儲存，再匯入。</div>
 <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:16px"><button class="dark-btn" data-action="apply-data">${ICON('check')} 確認更新本機資料</button><button class="outline-btn" data-action="cancel-data">取消此次匯入</button></div></section>`;
}
function renderUpdates(){
 const m=state.model;
 return `${pageHead('ONE FILE. EVERY MONTH.','更新資料，<span class="serif-em">不再折騰。</span>','你照舊維護 Excel。新版只需直接在 iPhone 或 Mac 匯入更新檔；無須手動轉 JSON、重新上傳網頁。')}
 <div class="update-grid"><div class="card card-pad"><div class="section-head"><div><h2>Excel → Screening Atlas</h2><p>只在本機解析與儲存，不連上傳 API。</p></div></div><button class="import-zone" data-action="pick-file" style="width:100%">${ICON('upload')}<h3>選擇新的 .xlsm / .xlsx</h3><p>從 iPhone「檔案」選 Excel，或從 Mac 選擇本機檔案。也支援匯入 JSON 私人備份。</p><span class="dark-btn">選擇檔案 ${ICON('arrow')}</span></button><div class="meta-foot">支援 Safari 16.4+。VBA 不會執行；匯入前請先確認 Excel 計算結果已更新。請勿把私人資料 JSON 放到公開 GitHub。</div></div>
 <div class="card card-pad">${sectionHead('目前版本 / CURRENT SNAPSHOT','你在這個裝置上使用的資料')}
 ${m?`<div class="source-row"><strong>檔案</strong><span>${safe(clip(state.data.meta?.name,38))}</span></div><div class="source-row"><strong>匯入版本日期</strong><span>${safe(state.data.meta?.created||'—')}</span></div><div class="source-row"><strong>MASTER</strong><span>${count(m.stats.facilities)} 筆</span></div><div class="source-row"><strong>ENGAGEMENTS</strong><span>${count(m.stats.engagements)} 筆</span></div><div class="source-row"><strong>SHA-256</strong><span class="mono">${safe((state.data.meta?.sha256||'').slice(0,15))}…</span></div><div style="display:flex;gap:9px;flex-wrap:wrap;margin-top:20px">${buttonAction('export-backup','下載私人備份 JSON','outline-btn','download')}${state.previous?buttonAction('restore-previous','回復前一版本','soft-btn','clock'):''}</div>`:`<div class="empty-block" style="padding:30px"><p>此裝置尚無已匯入版本。</p></div>`}</div></div>
 ${updatesPreview()}
 ${m?`<section class="section update-grid"><div class="card card-pad">${sectionHead('Source freshness / 各來源更新標記','擷取自 Excel 各 Input 工作表第一列')}${stampsHTML(state.data.meta?.sourceStamps)}</div><div class="card card-pad">${sectionHead('Update log','來源 Excel 的逐次維護紀錄')}<div class="table-scroll"><table class="simple-table"><thead><tr><th>日期</th><th>維護者</th><th>更新項目</th></tr></thead><tbody>${state.data.updates.slice(-10).reverse().map(u=>`<tr><td class="mono">${safe(u.date)}</td><td>${safe(u.owner)}</td><td class="table-country">${safe(Object.entries(u).filter(([k,v])=>k!=='owner'&&k!=='date'&&v).map(([k,v])=>k+(v==='✔'?'':` (${v})`)).join(', '))}</td></tr>`).join('')}</tbody></table></div></div></section>`:''}
 <section class="section"><div class="inline-note"><strong>安全提醒：</strong>「同一個 App 網址」不代表「同一份跨裝置資料」。目前是隱私優先的本機版：每台裝置各自匯入一次，沒有自動跨裝置同步。若之後要多人協作、自動同步與管理權限，再增加私有後端更合適。</div></section>`;
}
function render(){
 document.title=`${titles[state.page]} · Screening Atlas`;
 $('#breadcrumb').textContent=titles[state.page];
 $('#side-database').textContent=state.data?`${count(state.model?.stats.facilities)} FACTORIES · LOCAL`:'尚未匯入資料';
 for(const b of document.querySelectorAll('[data-view]'))b.classList.toggle('selected',b.dataset.view===state.page);
 const pages={overview:renderOverview,screen:renderScreen,library:renderLibrary,quality:renderQuality,updates:renderUpdates};
 $('#main-content').innerHTML=pages[state.page]();
}

function detailHTML(f){
 const main=state.model;
 const record=(head,sub,extra)=>`<div class="detail-record"><div class="detail-record-top"><strong>${safe(head||'—')}</strong><small>${safe(sub||'')}</small></div>${extra?`<p>${safe(extra)}</p>`:''}</div>`;
 const topBrands=[...new Set(f.brands)].slice(0,12),parent=[...new Set(f.parents.map(p=>p.parent))];
 return `<div class="drawer" role="dialog" aria-modal="true" aria-labelledby="drawer-title"><div class="drawer-top"><span class="mini-label">${f.external?'SOURCE-ONLY · NOT LINKED TO MASTER':'FACTORY INTELLIGENCE / MASTER #'+safe(f.id)}</span><button class="round-icon-btn" data-action="close-drawer" aria-label="關閉">${ICON('close')}</button></div>
 <h2 id="drawer-title" class="drawer-title">${safe(f.name)}</h2><p class="drawer-sub">${safe(lines(f.country).join(', ')||'國家未知')} · Higg ID：${safe(f.higgs.join(' / ')||'無紀錄')} · ${safe(lines(f.tier).join(' / ')||'Tier 未知')}</p>
 <div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:14px">${programBadges(f)}${f.risks.length?badge('需人工確認','orange'):''}</div>
 ${f.external?'<div class="inline-note" style="margin-top:17px">這是原始來源名單中的獨立工廠，尚未安全連到 Engagement Master。此處不代表已合作，也不代表已完成評估。</div>':''}<div class="drawer-stats"><div class="drawer-stat"><strong>${count(f.engagements.length)}</strong><span>來源合作紀錄</span></div><div class="drawer-stat"><strong>${count(f.strategics.length)}</strong><span>Strategic entries</span></div><div class="drawer-stat"><strong>${count(f.assessments.length)}</strong><span>Assessments</span></div></div>
 ${f.risks.length?`<div class="warn-inline"><strong>身份與合併提醒</strong><div style="margin-top:5px">${safe(f.risks.join('；'))}。請以原始來源確認多個廠址/識別碼是否真的屬同一工廠。</div></div>`:''}
 <section class="drawer-section"><h3>KNOWN ALIASES · 工廠別名</h3><div class="detail-list">${(f.names.length?f.names:[f.name]).slice(0,18).map(x=>record(x,'Recorded name')).join('')}</div></section>
 ${parent.length?`<section class="drawer-section"><h3>MANUFACTURER / PARENT GROUP</h3><div class="detail-list">${parent.map(x=>record(x,'Manufacturer info')).join('')}</div></section>`:''}
 <section class="drawer-section"><h3>BRANDS · 關聯品牌</h3><div style="display:flex;gap:5px;flex-wrap:wrap">${topBrands.length?topBrands.map(x=>badge(x,'blue')).join(''):badge('尚無','gray')}</div></section>
 <section class="drawer-section"><h3>ENGAGEMENT RECORDS · 合作明細 (${f.engagements.length})</h3><div class="detail-list">${f.engagements.length?f.engagements.slice(0,120).map(e=>record(e.brand||e.source,e.source,`${e.services||'未標服務'} · ${e.start||'—'} – ${e.end||'—'} · ${e.ctaVersion||''} · ${e.name}`)).join(''):record('沒有已連結合作紀錄','','')}</div></section>
 ${f.strategics.length?`<section class="drawer-section"><h3>STRATEGIC LISTS (${f.strategics.length})</h3><div class="detail-list">${f.strategics.slice(0,35).map(e=>record(e.source,'Strategic source',`${e.name} · Higg ${e.higg||'—'}`)).join('')}</div></section>`:''}
 ${f.assessments.length?`<section class="drawer-section"><h3>ASSESSMENT RECORDS (${f.assessments.length})</h3><div class="detail-list">${f.assessments.slice(0,45).map(e=>record(e.brand,e.date,`${e.method} · ${e.engineer||'未標工程師'} · ${e.name}`)).join('')}</div></section>`:''}
 ${f.implementations.length?`<section class="drawer-section"><h3>IMPLEMENTATION SERVICES (${f.implementations.length})</h3><div class="detail-list">${f.implementations.slice(0,35).map(e=>record(e.services,e.status,`${e.brand} · ${e.start}–${e.end} · ${e.owner}`)).join('')}</div></section>`:''}
 <div class="meta-foot">來源：你匯入的 Excel Master list、Consolidated list、Strategic list、Assessment Records、Implementation records、Manufacturer info。顯示結果為歷史資料，不代表現在仍有合約或專案進行中。輔助資料使用保守的唯一識別連結。</div></div>`;
}
function openDetail(id){const f=state.model?.byMaster.get(id);if(!f){toast('找不到這筆工廠資料');return}
 const d=$('#drawer');d.innerHTML=detailHTML(f);d.hidden=false;document.body.style.overflow='hidden';d.querySelector('button')?.focus();
}
function closeDetail(){const d=$('#drawer');d.hidden=true;d.innerHTML='';document.body.style.overflow=''}
const download=(name,content,mime='application/json')=>{
 const blob=new Blob([content],{type:mime});const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);
};
function doScreen(form){
 state.screenInputs=Object.fromEntries(new FormData(form));
 state.screenResults=screenFacility(state.model,state.screenInputs);
 state.screenBatch=null;render();
}
function doBatch(){
 const text=$('#batch-input')?.value.trim();if(!text){toast('請先貼上多行資料');return}
 const inputs=text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean).slice(0,200);
 state.screenBatch=inputs.map(input=>{
  const [name='',country='',higg='']=input.split(/[,，\t]/).map(x=>x.trim());
  const matches=screenFacility(state.model,{name,country,higg}),best=matches[0];
  return {input,name,country,higg,match:best?.facility.name||'',matchedHigg:best?.facility.higgs.join(' / ')||'',score:best?.score||0,confidence:best?.tier||'none',recordCount:best?.facility.engagements.length||0,reviewRequired:best?.facility.risks.length?'YES':'NO'};
 });
 render();toast(`已完成 ${state.screenBatch.length} 筆批次篩查`);
}
function refreshLibrary(){if(state.page!=='library')return;
 const n=$('#library-results');if(n)n.innerHTML=libraryTable();
}
function go(page){if(!titles[page])return;state.page=page;closeDetail();render();window.scrollTo({top:0,behavior:'instant'});}
async function readInput(file){
 if(!file)return;toast('正在本機解析資料，請勿關閉此畫面…',true);
 try{
  let data;
  if(file.name.toLowerCase().endsWith('.json')){data=validateData(JSON.parse(await file.text()));}
  else if(/\.(xlsx|xlsm)$/i.test(file.name)){data=validateData(await importWorkbook(file,text=>toast(text,true)));}
  else throw new Error('目前只支援 .xlsm、.xlsx 或 JSON 私人備份。');
  state.pending=data;go('updates');toast('解析完成！請檢查下方筆數並確認更新。');
  document.querySelector('.diff')?.scrollIntoView({block:'center',behavior:'smooth'});
 }catch(e){console.error(e);toast(`匯入未完成：${e.message}`)}
}
function getDB(){return new Promise((resolve,reject)=>{
 const req=indexedDB.open('screening-atlas-private-v1',1);
 req.onupgradeneeded=()=>req.result.createObjectStore('snapshots');
 req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);
});}
async function store(key,value){const db=await getDB();return new Promise((res,rej)=>{
 const tx=db.transaction('snapshots','readwrite');tx.objectStore('snapshots').put(value,key);tx.oncomplete=()=>{db.close();res()};tx.onerror=()=>{db.close();rej(tx.error)};
});}
async function load(key){const db=await getDB();return new Promise((res,rej)=>{
 const tx=db.transaction('snapshots','readonly'),req=tx.objectStore('snapshots').get(key);
 req.onsuccess=()=>{db.close();res(req.result||null)};req.onerror=()=>{db.close();rej(req.error)};
});}
async function applyData(){
 if(!state.pending)return;
 const next=state.pending;toast('正在更新本機資料庫…',true);
 try{
  if(state.data)await store('previous',state.data);
  await store('current',next);
  state.previous=state.data;
  state.data=next;state.model=createModel(next);state.pending=null;state.screenResults=null;state.screenBatch=null;
  go('overview');toast(`更新成功：${count(state.model.stats.facilities)} 家 Master 工廠已載入`);
 }catch(e){console.error(e);toast('本機儲存失敗：'+e.message+'。請檢查 Safari 儲存空間。')}
}
async function restorePrevious(){if(!state.previous)return;
 if(!confirm('要將目前版本換回上一個版本嗎？目前版本會保存成上一版備份。'))return;
 const prev=state.previous;state.pending=prev;await applyData();
}

document.addEventListener('click',async e=>{
 const nav=e.target.closest('[data-view]');if(nav){go(nav.dataset.view);return}
 const el=e.target.closest('[data-action]');if(!el)return;
 const action=el.dataset.action;
 if(action==='pick-file'){startImport();return}
 if(action==='detail'){openDetail(el.dataset.id);return}
 if(action==='close-drawer'){closeDetail();return}
 if(action==='cancel-data'){state.pending=null;render();return}
 if(action==='apply-data'){await applyData();return}
 if(action==='export-backup'){if(state.data)download(`screening-private-${state.data.meta?.created||'backup'}.json`,JSON.stringify(state.data));return}
 if(action==='restore-previous'){await restorePrevious();return}
 if(action==='batch'){doBatch();return}
 if(action==='export-screen'&&state.screenResults){
  const rows=state.screenResults.map(({facility:f,score,reason})=>({facility:f.name,higg:f.higgs.join(' / '),country:f.country,score,reason,brands:f.brands.join('; '),programs:f.programs.join('; '),engagements:f.engagements.length,strategic:f.strategics.length,review:f.risks.join('; ')}));download('screening-results.csv',toCSV(rows),'text/csv;charset=utf-8');return;
 }
 if(action==='export-batch'&&state.screenBatch){download('screening-batch-results.csv',toCSV(state.screenBatch),'text/csv;charset=utf-8');return}
});
$('#file-picker').addEventListener('change',e=>{const f=e.target.files?.[0];e.target.value='';readInput(f)});
document.addEventListener('submit',e=>{
 if(e.target.id==='hero-form'){
  e.preventDefault();const query=new FormData(e.target).get('quick-search')||'';
  state.screenInputs={name:/^\d{3,8}$/.test(query.trim())?'':query,higg:/^\d{3,8}$/.test(query.trim())?query:'',country:'',brand:''};
  state.screenResults=screenFacility(state.model,state.screenInputs);go('screen');
 } else if(e.target.id==='screen-form'){e.preventDefault();doScreen(e.target)}
});
document.addEventListener('input',e=>{
 if(e.target.id==='library-query'){state.library.query=e.target.value;refreshLibrary()}
});
document.addEventListener('change',e=>{
 if(e.target.id==='library-country'){state.library.country=e.target.value;refreshLibrary()}
 if(e.target.id==='library-program'){state.library.program=e.target.value;refreshLibrary()}
 if(e.target.id==='include-external'){state.library.includeExternal=e.target.checked;refreshLibrary()}
 if(e.target.id==='only-strategic'){state.library.onlyStrategic=e.target.checked;refreshLibrary()}
 if(e.target.id==='only-risk'){state.library.onlyRisk=e.target.checked;refreshLibrary()}
});
$('#drawer').addEventListener('click',e=>{if(e.target.id==='drawer')closeDetail()});
document.addEventListener('keydown',e=>{
 if(e.key==='Escape'){closeDetail();return}
 const typing=['INPUT','TEXTAREA','SELECT'].includes(document.activeElement?.tagName);
 if(!typing&&!e.metaKey&&!e.ctrlKey){if(e.key==='g')go('overview');if(e.key==='s')go('screen')}
 if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();go('screen');$('#screen-name')?.focus()}
 if(e.key==='Enter'&&e.target.closest('.result-card')){e.preventDefault();openDetail(e.target.closest('.result-card').dataset.id)}
});
(async function init(){
 render();
 try{
  const data=await load('current');state.previous=await load('previous');
  if(data){state.data=validateData(data);state.model=createModel(state.data);render()}
 }catch(e){console.error(e);toast('目前無法開啟本機儲存空間；請勿使用私密瀏覽模式。')}
 if('serviceWorker' in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./sw.js').catch(()=>{});
})();
