const clean = v => String(v??'').normalize('NFKC').replace(/\u00a0/g,' ').trim();
export const lines = v => [...new Set(clean(v).split(/\r?\n/).map(s=>s.trim()).filter(Boolean))];
export const normalized = v => clean(v).toLowerCase().replace(/&/g,'and').replace(/(?:股份有限公司|有限公司|有限責任公司|有限责任公司|集團|集团|分公司|工廠|工厂|company|corporation|limited|ltd|co\b|pvt|private|inc\b|sarl|llc|gmbh)/gi,'').replace(/[^\p{L}\p{N}]+/gu,'');
export const countries = v => [...new Set(lines(v).map(normalized).map(c=>({'usa':'unitedstates','southkorea':'korea','turkiye':'turkey','viet nam':'vietnam','taiwanprovinceofchina':'taiwan'}[c]||c)))];
const ids = v => [...new Set((clean(v).match(/(?<!\d)\d{3,8}(?!\d)/g)||[]))];
const uniq = a => [...new Set(a.filter(Boolean))];
const fields = (o,keys)=>keys.flatMap(k=>lines(o[k]));
const tokens = s=>new Set((clean(s).toLowerCase().match(/[\p{L}\p{N}]{2,}/gu)||[]));
const dateYear = v=>{let m=String(v||'').match(/(?:20\d\d|\b\d\d\b)$/);return m?m[0]:''};

export function createModel(data){
  const facilities=data.master.filter(m=>m.id&&m.name).map(m=>{
    const names=fields(m,['name']);const higgs=ids(m.higg), co=countries(m.country), brand=lines(m.brand), sv=lines(m.services);
    return {...m,names,name:names[0]||m.name,higgs,countries:co,brands:brand,serviceNames:sv,
      aliases:uniq([...names,...lines(m.alias1),...lines(m.alias2)]),programs:[],engagements:[],strategics:[],assessments:[],implementations:[],parents:[],risks:[],searchIndex:''};
  });
  const byMaster=new Map(facilities.map(f=>[f.id,f]));
  const byHigg=new Map(),byAlias=new Map();
  for(const f of facilities){
    for(const id of f.higgs){let set=byHigg.get(id)||new Set();set.add(f);byHigg.set(id,set)}
    for(const name of f.aliases){
      const norm=normalized(name);if(norm.length<4)continue;
      for(const c of f.countries){const key=norm+'|'+c;let set=byAlias.get(key)||new Set();set.add(f);byAlias.set(key,set)}
    }
    if(f.higgs.length>1)f.risks.push('同一 Master 合併多個 Higg ID');
    if(f.countries.length>1)f.risks.push('同一 Master 合併多個國家');
  }
  const orphans={engagements:0,strategics:0,manufacturers:0,assessments:0,implementations:0};
  const unlinked={strategics:[],assessments:[],implementations:[],manufacturers:[]};
  const attach=(type,list,proc)=>{
    for(const item of list){
      const f=proc(item);
      if(f)f[type].push(item);else {orphans[type]++;if(unlinked[type])unlinked[type].push(item)}
    }
  };
  attach('engagements',data.engagements,rec=>byMaster.get(rec.masterId));
  for(const [id,set] of byHigg)if(set.size>1)for(const f of set)f.risks.push(`Higg ID ${id} 同時存在於其他 Master`);
  // High-precision auxiliary matching only: never auto-join conflicting IDs or countries.
  function resolveExact(rec){
    const h=ids(rec.higg);const c=countries(rec.country);
    if(h.length){
      const matches=uniq(h.flatMap(x=>[...(byHigg.get(x)||[])]));
      if(matches.length!==1)return null;
      const f=matches[0];
      if(c.length && !c.some(v=>f.countries.includes(v)))return null;
      return f;
    }
    const norm=normalized(rec.name);
    if(norm.length<5||!c.length)return null;
    const matches=uniq(c.flatMap(country=>[...(byAlias.get(norm+'|'+country)||[])]));
    return matches.length===1?matches[0]:null;
  }
  attach('strategics',data.strategic,resolveExact);
  attach('assessments',data.assessments,resolveExact);
  attach('implementations',data.implementations,resolveExact);
  // Manufacturer data uses a different facility name key and same safe resolution.
  for(const rec of data.manufacturers){
    const f=resolveExact(rec);if(f)f.parents.push(rec);else {orphans.manufacturers++;unlinked.manufacturers.push(rec)}
  }
  for(const f of facilities){
    const p=new Set();
    for(const e of f.engagements){
      const s=(e.services+' '+e.source).toUpperCase();
      if(/BPAP/.test(s))p.add('BPAP');
      if(/CTM/.test(s))p.add('CTM');
      if(/CTS|CLP/.test(s))p.add('CTS');
      if(/CTA|CWTA|ASSESS/.test(s))p.add('CTA');
      if(/OTHER|CPO|CONSULT/.test(s))p.add('Other');
    }
    if(f.implementations.length)p.add('Implementation');
    f.programs=[...p];
    f.brands=uniq([...f.brands,...f.engagements.flatMap(e=>lines(e.brand)),...f.assessments.flatMap(e=>lines(e.brand))]);
    f.searchIndex=normalized([f.name,...f.names,...f.aliases,...f.brands,...f.higgs,...f.countries,...f.parents.map(p=>p.parent)].join(' '));
  }
  // Keep unmatched prospective/assessment/parent records searchable WITHOUT
  // falsely claiming they are the same facility as an engagement Master.
  const externalGroups=new Map(), externalByName=new Map();
  for(const type of ['strategics','assessments','implementations','manufacturers']){
    for(const rec of unlinked[type]){
      const country=countries(rec.country), allIds=ids(rec.higg), norm=normalized(rec.name);
      if(!norm||!country.length)continue;
      const nameKey=norm+'|'+country[0];
      let key=allIds.length===1?`h:${allIds[0]}|${country[0]}`:`n:${nameKey}`;
      // A nameless-ID-free row can join an ID group only if the name-country
      // identifies ONE external group; never merge two conflicting IDs.
      if(!allIds.length && externalByName.has(nameKey) && externalByName.get(nameKey).size===1)
        key=[...externalByName.get(nameKey)][0];
      let f=externalGroups.get(key);
      if(!f){
        f={id:'external:'+key,index:'',name:rec.name,names:[rec.name],aliases:[rec.name],higg:rec.higg||'',higgs:allIds,country:rec.country,countries:country,brand:'',brands:[],tier:'',serviceNames:[],programs:[],engagements:[],strategics:[],assessments:[],implementations:[],parents:[],risks:[],external:true,searchIndex:''};
        externalGroups.set(key,f);
      }
      f[type==='manufacturers'?'parents':type].push(rec);
      f.names=uniq([...f.names,rec.name]);f.aliases=uniq([...f.aliases,rec.name,rec.alias].filter(Boolean));
      f.higgs=uniq([...f.higgs,...allIds]);
      if(rec.brand)f.brands=uniq([...f.brands,...lines(rec.brand)]);
      let set=externalByName.get(nameKey)||new Set();set.add(key);externalByName.set(nameKey,set);
    }
  }
  const externalFacilities=[...externalGroups.values()];
  for(const f of externalFacilities){
    if(f.implementations.length)f.programs.push('Implementation');
    f.searchIndex=normalized([...f.names,...f.aliases,...f.higgs,...f.countries,...f.brands,...f.parents.map(p=>p.parent),...f.strategics.map(x=>x.source)].join(' '));
    byMaster.set(f.id,f);
  }
  const allFacilities=[...facilities,...externalFacilities];
  const countryOptions=[...new Set(allFacilities.flatMap(f=>lines(f.country)))].sort((a,b)=>a.localeCompare(b));
  const brandOptions=[...new Set(facilities.flatMap(f=>f.brands))].sort((a,b)=>a.localeCompare(b));
  const sourceCounts={};for(const x of data.engagements)sourceCounts[x.source]=(sourceCounts[x.source]||0)+1;
  const risks=facilities.filter(f=>f.risks.length);
  return {data,facilities,externalFacilities,allFacilities,byMaster,byHigg,byAlias,orphans,sourceCounts,risks,countryOptions,brandOptions,resolveExact,
    stats:{facilities:facilities.length,external:externalFacilities.length,engagements:data.engagements.length,strategic:data.strategic.length,assessments:data.assessments.length,implementations:data.implementations.length,risks:risks.length}};
}

function likeness(a,b){
  if(!a||!b)return 0;
  if(a===b)return 1;
  if(a.length<5||b.length<5)return 0;
  if(a.includes(b)||b.includes(a))return Math.min(a.length,b.length)/Math.max(a.length,b.length)>.68?.86:.65;
  const A=tokens(a),B=tokens(b);const common=[...A].filter(x=>B.has(x)).length;
  if(!A.size||!B.size)return 0;
  return common/(A.size+B.size-common);
}
export function screenFacility(model,{name='',higg='',country='',brand=''}={}){
  const queryIds=ids(higg),qname=normalized(name),qc=normalized(country),qb=normalized(brand);
  if(!queryIds.length && qname.length<3)return [];
  const results=[];
  for(const f of model.allFacilities){
    if(qc && !f.countries.includes(qc))continue;
    const idMatch=queryIds.some(id=>f.higgs.includes(id));
    const nearNames=f.aliases.map(normalized);
    let score=0,reason='';
    if(queryIds.length){
      if(!idMatch){
        const best=qname?Math.max(0,...nearNames.map(n=>likeness(qname,n))):0;
        if(best<.8)continue;
        score=f.higgs.length?48:53;
        reason=f.higgs.length?'名稱相似，但 Higg ID 不一致（請人工確認）':'名稱相似，但來源沒有 Higg ID 可核對';
      } else {
        score=f.external?92:96;reason='Higg ID 完全一致';
        if(qname && nearNames.includes(qname)){score=f.external?96:100;reason='Higg ID 與工廠名稱一致';}
        else if(qname && !nearNames.some(n=>n.includes(qname)||qname.includes(n)))score=90;
      }
    } else {
      const best=Math.max(0,...nearNames.map(n=>likeness(qname,n)));
      if(best===1){score=qc?(f.external?85:90):78;reason='工廠名稱完全一致'+(qc?'，且國家一致':'');}
      else if(best>=.8){score=qc?79:69;reason='工廠名稱高度相似';}
      else if(best>=.65){score=qc?68:57;reason='工廠名稱部分相符';}
      else continue;
    }
    if(qb && f.brands.some(b=>normalized(b)===qb))score=Math.min(100,score+3);
    results.push({facility:f,score,reason,tier:score>=90?'exact':score>=76?'probable':'review'});
  }
  return results.sort((a,b)=>b.score-a.score||b.facility.engagements.length-a.facility.engagements.length).slice(0,40);
}
export function librarySearch(model,{query='',country='',brand='',program='',onlyStrategic=false,onlyRisk=false,includeExternal=false,limit=200}={}){
  const q=normalized(query),ct=normalized(country),br=normalized(brand);
  return (includeExternal?model.allFacilities:model.facilities).filter(f=>(!q||f.searchIndex.includes(q))&&(!ct||f.countries.includes(ct))&&
    (!br||f.brands.some(b=>normalized(b)===br))&&(!program||f.programs.includes(program))&&
    (!onlyStrategic||f.strategics.length)&&(!onlyRisk||f.risks.length)).slice(0,limit);
}
export const toCSV=(rows)=>{
  if(!rows.length)return '';
  const keys=Object.keys(rows[0]);let csv='\uFEFF'+keys.map(esc).join(',')+'\r\n';
  for(const row of rows)csv+=keys.map(k=>esc(row[k])).join(',')+'\r\n';
  return csv;
};
function esc(x){const s=String(x??'').replace(/^[=+@-]/,'\'$&');return /[",\n\r]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s}
