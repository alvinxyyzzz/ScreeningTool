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
 return `<section class="lookup-state"><div class="lookup-state-icon">${ICON('lock')}</div><h2>這台裝置尚未載入查詢資料</h2><p>請由資料維護者在本機準備最新版 Screening 資料後再查詢。網頁程式本身不包含公司的完整工廠資料。</p><div class="lookup-muted">一般使用者無需查看 Excel Master List 或資料維護頁面。</div></section>`;
}
function renderOverview(){
 const v=state.screenInputs||{};
 const query=v.query||v.higg||v.name||'';
 const advanced=!!(v.country||v.brand);
 return `<section class="lookup-intro">
   <div class="lookup-eyebrow"><span class="lookup-dot"></span>FACTORY ENGAGEMENT LOOKUP</div>
   <h1>這家工廠，<span>我們合作過嗎？</span></h1>
   <p>搜尋工廠名稱或 Higg ID，快速確認歷史合作、服務項目、品牌與評估紀錄。</p>
 </section>
 <section class="lookup-panel">
  <form id="screen-form" class="lookup-form">
   <label for="screen-query">搜尋工廠</label>
   <div class="lookup-input-row">
    <div class="lookup-input-wrap">${ICON('search')}<input type="search" name="query" id="screen-query" autocomplete="off" spellcheck="false" value="${safe(query)}" placeholder="輸入工廠名稱或 Higg ID…" ${!state.model?'disabled':''}/></div>
    <button class="lookup-submit" type="submit" ${!state.model?'disabled':''}>搜尋 ${ICON('arrow')}</button>
   </div>
   <div class="lookup-helpers">
    <span>建議優先使用 Higg ID，或輸入完整的工廠名稱。</span>
    <details class="lookup-advanced" ${advanced?'open':''}><summary>${ICON('filter')} 進階篩選</summary><div class="lookup-filters">
      <div class="lookup-filter"><label for="screen-country">國家（選填）</label><input id="screen-country" name="country" list="country-options" value="${safe(v.country||'')}" placeholder="例如 Vietnam" /></div>
      <div class="lookup-filter"><label for="screen-brand">品牌（選填）</label><input id="screen-brand" name="brand" value="${safe(v.brand||'')}" placeholder="例如 lululemon" /></div>
     </div></details>
    ${state.model?`<datalist id="country-options">${state.model.countryOptions.map(c=>`<option value="${safe(c)}"></option>`).join('')}</datalist>`:''}
   </div>
  </form>
 </section>
 <div id="result-anchor">${screenResultsHTML()}</div>
 ${state.model?`<section class="lookup-batch"><details ${state.screenBatch?'open':''}>
 <summary><span>${ICON('layers')} <strong>需要一次查多家工廠？</strong><small>展開批次查詢</small></span>${ICON('chevron')}</summary>
 <div class="batch-inner"><p>每行一間工廠，格式：<strong>工廠名稱, 國家, Higg ID</strong>（後兩欄可留空），最多 200 筆。</p>
 <textarea id="batch-input" aria-label="批次查詢工廠" placeholder="Envoy Textiles, Bangladesh, 42555&#10;Everest Textile, Taiwan, 5506"></textarea>
 <div class="batch-actions"><button class="lookup-submit" data-action="batch" type="button">執行批次查詢</button>
 ${state.screenBatch?`<span>已完成 ${state.screenBatch.length} 筆，其中 ${state.screenBatch.filter(x=>x.match).length} 筆找到候選</span><button class="batch-export" data-action="export-batch" type="button">${ICON('download')} 下載 CSV 結果</button>`:''}</div></div>
 </details></section>`:''}
 <div class="lookup-footnote">本工具僅供查詢歷史資料；「未找到紀錄」不代表從未合作。名稱近似或多個 Higg ID 的結果需要人工核對。資料保存在此裝置，不會自動同步。</div>`;
}
function screenResultItem(r){
 const f=r.facility;
 const hasCooperation=f.engagements.length>0;
 const status=hasCooperation?'有歷史合作紀錄':f.implementations.length?'有服務執行紀錄':f.assessments.length?'有評估紀錄':f.strategics.length?'僅出現在名單':'其他來源紀錄';
 const style=hasCooperation?'status-found':f.implementations.length||f.assessments.length?'status-assessed':'status-list';
 const services=f.programs.slice(0,4).map(x=>badge(x)).join('');
 const brands=f.brands.slice(0,3).join(' · ');
 return `<article class="lookup-result" data-action="detail" data-id="${safe(f.id)}" tabindex="0" role="button" aria-label="查看 ${safe(f.name)} 的歷史紀錄">
 <div class="lookup-result-main"><div class="lookup-result-top"><span class="record-status ${style}"><span class="status-dot"></span>${status}</span><span class="match-hint">${r.tier==='review'?'身份需再確認':r.reason}</span></div>
 <h3>${safe(f.name)}</h3><div class="lookup-result-place">${safe(lines(f.country).join(' / ')||'國家未記錄')} <span>·</span> Higg ID ${safe(f.higgs.join(' / ')||'未記錄')}</div>
 <div class="lookup-result-tags">${services}${brands?`<span class="tag brand-preview">${safe(brands)}</span>`:''}${f.risks.length?badge('身份待核對','orange'):''}</div>
 </div><div class="lookup-result-side"><strong>${count(f.engagements.length)}</strong><span>筆歷史合作</span><div class="lookup-open">查看紀錄 ${ICON('arrow')}</div></div></article>`;
}
function screenResultsHTML(){
 if(!state.model)return renderEmptyLanding();
 if(state.screenResults===null)return `<div class="lookup-before"><div class="before-heading">查詢後可以看到</div><div class="before-grid">
 <div>${ICON('check')}<strong>合作紀錄</strong><span>過去是否有接觸或合作</span></div>
 <div>${ICON('layers')}<strong>服務與品牌</strong><span>CTA、CTS、CTM 等相關項目</span></div>
 <div>${ICON('clock')}<strong>時間與來源</strong><span>方便核對歷史紀錄</span></div>
 </div></div>`;
 if(!state.screenResults.length)return `<section class="lookup-state lookup-no-results"><div class="lookup-state-icon">${ICON('search')}</div><h2>目前沒有找到匹配紀錄</h2><p>可試試完整英文名稱、其他廠名或 Higg ID，並移除進階篩選。這不代表一定沒有合作歷史。</p></section>`;
 return `<section class="lookup-results"><div class="lookup-results-head"><div><div class="lookup-eyebrow">SEARCH RESULTS</div><h2>找到 ${state.screenResults.length} 筆可能匹配的工廠</h2></div><span>點開查看歷史紀錄</span></div>
 ${state.screenResults.map(screenResultItem).join('')}
 <p class="lookup-results-note">${ICON('alert')} 結果依名稱及識別碼匹配程度排序；同名工廠或資料合併異常需再核對。</p></section>`;
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
 const fileName=state.data?.meta?.name||'尚未匯入資料';
 return `<div class="manage-back"><button data-view="overview" class="outline-btn">${ICON('arrow')} 返回工廠搜尋</button></div>
 <div class="manage-head"><div class="lookup-eyebrow">MAINTAINER WORKSPACE · LOCAL ONLY</div><h1>資料更新</h1><p>僅供資料維護者操作。你的 Excel 仍是主資料來源；本頁不會上傳檔案到 GitHub。</p></div>
 <div class="manage-grid"><section class="card card-pad"><h2>匯入新版 Excel</h2><p>支援 .xlsm、.xlsx 和私人 JSON 備份；僅在此裝置解析，不會執行 VBA。</p><button class="import-zone" data-action="pick-file" type="button">${ICON('upload')}<h3>選擇檔案</h3><p>匯入後會先顯示差異，確認才會取代這台裝置的資料。</p></button></section>
 <section class="card card-pad"><h2>目前本機版本</h2><div class="source-row"><strong>檔案</strong><span>${safe(clip(fileName,32))}</span></div><div class="source-row"><strong>匯入日期</strong><span>${safe(state.data?.meta?.created||'—')}</span></div>
 ${m?`<div class="manage-actions">${buttonAction('export-backup','下載私人 JSON 備份','outline-btn','download')}${state.previous?buttonAction('restore-previous','回復上一版','soft-btn','clock'):''}</div>`:'<p class="manage-muted">目前沒有已匯入版本。</p>'}
 </section></div>
 ${updatesPreview()}
 <p class="lookup-detail-disclaimer">請先在 Excel 重新計算並儲存，再匯入。每個裝置的資料彼此獨立；此頁不具備登入與權限驗證功能，網址隱藏並不等於資料保護。</p>`;
}
function render(){
 if(!state.routeReady){state.routeReady=true;if(new URLSearchParams(location.search).get('manage')==='1')state.page='updates';}
 if(state.page!=='updates')state.page='overview';
 document.title=(state.page==='updates'?'資料更新':'Screening')+' · Factory History';
 const crumb=$('#breadcrumb');if(crumb)crumb.textContent=state.page==='updates'?'維護資料':'搜尋工廠';
 const db=$('#side-database');if(db)db.textContent=state.model?'查詢資料已就緒':'尚未載入查詢資料';
 $('#main-content').innerHTML=state.page==='updates'?renderUpdates():renderOverview();
}
function detailHTML(f){
 const companies=[...new Set(f.brands)].filter(Boolean);
 const parents=[...new Set(f.parents.map(x=>x.parent).filter(Boolean))];
 const gotHistory=f.engagements.length>0;
 const status=gotHistory?'有歷史合作紀錄':f.implementations.length?'有服務執行紀錄':f.assessments.length?'有評估紀錄':f.strategics.length?'僅出現在名單':'其他來源紀錄';
 const item=(h,meta,sub)=>`<div class="history-item"><div class="history-line"></div><div class="history-copy"><strong>${safe(h||'未記錄服務')}</strong><div class="history-meta">${safe(meta||'未記錄日期')}</div>${sub?`<div class="history-sub">${safe(sub)}</div>`:''}</div></div>`;
 const engagements=f.engagements.map(e=>{
  const dates=[e.start,e.end].filter(Boolean).join(' – ')||'時間未記錄';
  return item(e.services||e.source||'合作紀錄',[e.brand,dates].filter(Boolean).join(' · '),`來源：${e.source||'未註明'}${e.name?' · 登記工廠：'+e.name:''}`);
 }).join('');
 const assessments=f.assessments.map(e=>item('工廠評估',[e.brand,e.date,e.method].filter(Boolean).join(' · '),e.engineer?'負責人員：'+e.engineer:'')).join('');
 const implementations=f.implementations.map(e=>item(e.services||'Implementation',[e.brand,e.status,[e.start,e.end].filter(Boolean).join(' – ')].filter(Boolean).join(' · '),e.owner?'負責人員：'+e.owner:'')).join('');
 const strategicNames=[...new Set(f.strategics.map(x=>x.source).filter(Boolean))];
 return `<div class="drawer lookup-drawer" role="dialog" aria-modal="true" aria-labelledby="drawer-title">
  <div class="drawer-top"><span class="mini-label">FACTORY HISTORY</span><button class="round-icon-btn" data-action="close-drawer" aria-label="關閉明細">${ICON('close')}</button></div>
  <div class="lookup-detail-heading"><span class="record-status ${gotHistory?'status-found':f.assessments.length||f.implementations.length?'status-assessed':'status-list'}">${status}</span><h2 id="drawer-title" class="drawer-title">${safe(f.name)}</h2>
  <p class="drawer-sub">${safe(lines(f.country).join(' / ')||'國家未記錄')} · Higg ID：${safe(f.higgs.join(' / ')||'未記錄')}${f.tier?' · '+safe(lines(f.tier).join(' / ')):''}</p></div>
  ${f.risks.length?`<div class="warn-inline"><strong>請先核對工廠身份：</strong>${safe(f.risks.join('；'))}。系統不會把不同廠址自動視為同一家。</div>`:''}
  <div class="lookup-detail-stats"><div><strong>${f.engagements.length}</strong><span>合作紀錄</span></div><div><strong>${f.assessments.length}</strong><span>評估紀錄</span></div><div><strong>${f.strategics.length}</strong><span>戰略名單</span></div></div>
  <section class="lookup-detail-section"><h3>歷史合作紀錄 <span>${f.engagements.length}</span></h3>
  ${engagements||'<p class="lookup-detail-empty">目前沒有已確認連結的歷史合作紀錄。</p>'}
  </section>
  ${implementations?`<section class="lookup-detail-section"><h3>執行服務紀錄 <span>${f.implementations.length}</span></h3>${implementations}</section>`:''}
  ${assessments?`<section class="lookup-detail-section"><h3>評估紀錄 <span>${f.assessments.length}</span></h3>${assessments}</section>`:''}
  ${strategicNames.length?`<section class="lookup-detail-section"><h3>出現在戰略／品牌名單</h3><p class="lookup-detail-desc">名單命中不等於已經合作。</p><div class="lookup-mini-tags">${strategicNames.slice(0,24).map(n=>badge(n,'blue')).join('')}</div></section>`:''}
  ${companies.length?`<section class="lookup-detail-section"><h3>關聯品牌</h3><div class="lookup-mini-tags">${companies.slice(0,30).map(n=>badge(n,'blue')).join('')}</div></section>`:''}
  ${parents.length?`<section class="lookup-detail-section"><h3>關聯製造商</h3><div class="lookup-mini-tags">${parents.slice(0,12).map(n=>badge(n)).join('')}</div></section>`:''}
  <details class="lookup-evidence"><summary>查看工廠別名與識別資訊</summary><div>${(f.aliases||[]).slice(0,18).map(x=>`<div>${safe(x)}</div>`).join('')||'無其他名稱記錄'}</div></details>
  <p class="lookup-detail-disclaimer">以上來自維護者提供的歷史資料，僅供內部查詢，不表示目前仍有合約或服務進行中。來源未連結不等於沒有合作；若用於正式決策，請與資料維護者確認。</p></div>`;
}
function openDetail(id){const f=state.model?.byMaster.get(id);if(!f){toast('找不到這筆工廠資料');return}
 const d=$('#drawer');d.innerHTML=detailHTML(f);d.hidden=false;document.body.style.overflow='hidden';d.querySelector('button')?.focus();
}
function closeDetail(){const d=$('#drawer');d.hidden=true;d.innerHTML='';document.body.style.overflow=''}
const download=(name,content,mime='application/json')=>{
 const blob=new Blob([content],{type:mime});const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);
};
function doScreen(form){
 const entry=Object.fromEntries(new FormData(form));
 const query=String(entry.query||'').trim();
 if(!query){toast('請輸入工廠名稱或 Higg ID');return}
 if(!state.model){toast('此裝置尚未載入查詢資料');return}
 const higg=query.match(/^(?:higg(?:\s*id)?[\s:#：-]*)?(\d{3,8})$/i)?.[1]||'';
 state.screenInputs={query,name:higg?'':query,higg,country:String(entry.country||'').trim(),brand:String(entry.brand||'').trim()};
 state.screenResults=screenFacility(state.model,state.screenInputs);
 state.screenBatch=null;
 render();
 document.querySelector('#result-anchor')?.scrollIntoView({behavior:'smooth',block:'start'});
}
function doBatch(){
 const text=$('#batch-input')?.value.trim();if(!text){toast('請先貼上多行資料');return}
 const inputs=text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean).slice(0,200);
 state.screenBatch=inputs.map(input=>{
  const [name='',country='',higg='']=input.split(/[,，\t]/).map(x=>x.trim());
  const higgOnly=/^\d{3,8}$/.test(name)&&!country&&!higg;
  const matches=screenFacility(state.model,{name:higgOnly?'':name,country,higg:higgOnly?name:higg}),best=matches[0];
  return {input,name,country,higg,match:best?.facility.name||'',matchedHigg:best?.facility.higgs.join(' / ')||'',score:best?.score||0,confidence:best?.tier||'none',recordCount:best?.facility.engagements.length||0,reviewRequired:best?.facility.risks.length?'YES':'NO'};
 });
 render();toast(`已完成 ${state.screenBatch.length} 筆批次篩查`);
}
function refreshLibrary(){if(state.page!=='library')return;
 const n=$('#library-results');if(n)n.innerHTML=libraryTable();
}
function go(page){
 state.page=page==='updates'?'updates':'overview';
 closeDetail();render();window.scrollTo({top:0,behavior:'instant'});
}
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
 if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();go('screen');$('#screen-query')?.focus()}
 if(e.key==='Enter'&&e.target.closest('.lookup-result')){e.preventDefault();openDetail(e.target.closest('.lookup-result').dataset.id)}
});
(async function init(){
 render();
 try{
  const data=await load('current');state.previous=await load('previous');
  if(data){state.data=validateData(data);state.model=createModel(state.data);render()}
 }catch(e){console.error(e);toast('目前無法開啟本機儲存空間；請勿使用私密瀏覽模式。')}
 if('serviceWorker' in navigator&&location.protocol==='https:')navigator.serviceWorker.register('./sw.js').catch(()=>{});
})();
