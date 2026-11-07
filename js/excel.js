/* Pure-browser, dependency-free parser for the small, specific set of XML tables
   in the Screening Tool .xlsx/.xlsm format. Does not execute VBA/macros.
   Requires Safari 16.4+ (DecompressionStream('deflate-raw')). */
const encoder = new TextDecoder('utf-8');
const at = (r, k) => String(r[k] ?? '').trim();

class ZipReader {
  constructor(buffer) {
    this.buffer = buffer; this.dv = new DataView(buffer); this.files = new Map();
    const d = this.dv, length = buffer.byteLength;
    let eocd = -1;
    for (let i=length-22; i>=Math.max(0,length-65558); i--) {
      if (d.getUint32(i,true) === 0x06054b50) { eocd=i; break; }
    }
    if (eocd<0) throw new Error('不是有效的 Excel ZIP 檔案。');
    const count=d.getUint16(eocd+10,true), off=d.getUint32(eocd+16,true);
    if (count>5000) throw new Error('檔案包含過多項目。');
    let cursor=off;
    for(let i=0;i<count;i++) {
      if(d.getUint32(cursor,true)!==0x02014b50) throw new Error('ZIP 索引異常。');
      const flags=d.getUint16(cursor+8,true), method=d.getUint16(cursor+10,true);
      const csize=d.getUint32(cursor+20,true), usize=d.getUint32(cursor+24,true);
      const nlen=d.getUint16(cursor+28,true), xlen=d.getUint16(cursor+30,true), clen=d.getUint16(cursor+32,true);
      const local=d.getUint32(cursor+42,true);
      if (nlen>400 || usize>50*1024*1024) throw new Error('檔案項目過大，已停止匯入。');
      const name=encoder.decode(new Uint8Array(buffer,cursor+46,nlen));
      this.files.set(name,{flags,method,csize,usize,local});
      cursor+=46+nlen+xlen+clen;
    }
  }
  async get(name){
    const m=this.files.get(name);if(!m)return null;
    if(m.flags&1)throw new Error('暫不支援有 ZIP 密碼加密的 Excel。');
    if(![0,8].includes(m.method))throw new Error('不支援此壓縮格式。');
    const d=this.dv,p=m.local;
    if(d.getUint32(p,true)!==0x04034b50)throw new Error('ZIP 內部索引異常。');
    const start=p+30+d.getUint16(p+26,true)+d.getUint16(p+28,true);
    const chunk=new Uint8Array(this.buffer,start,m.csize);
    if(m.method===0) return encoder.decode(chunk);
    if(typeof DecompressionStream==='undefined')throw new Error('瀏覽器版本太舊，請更新 iOS / Safari。');
    const out=await new Response(new Blob([chunk]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
    if(out.byteLength>50*1024*1024)throw new Error('解壓縮內容過大。');
    return encoder.decode(out);
  }
}
const parseXml = s => {
  const doc=new DOMParser().parseFromString(s,'application/xml');
  if(doc.getElementsByTagName('parsererror').length)throw new Error('Excel XML 解析失敗。');
  return doc;
};
const localName = n => n.localName || n.nodeName.split(':').pop();
const elements = (n, name) => Array.from(n.getElementsByTagNameNS('*',name));
function columnNumber(ref){
  const match=(ref||'').match(/^[A-Z]+/);if(!match)return -1;
  return [...match[0]].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0)-1;
}
function getRows(xml, strings, start=1){
  const root=parseXml(xml), rows=elements(root,'row'), result=[];
  for(const el of rows){
    const index=Number(el.getAttribute('r'));
    if(index<start)continue;
    const row={};
    for(const cell of Array.from(el.children)) {
      if(localName(cell)!=='c')continue;
      const col=columnNumber(cell.getAttribute('r'));
      if(col<0)continue;
      const type=cell.getAttribute('t');
      let v=null;
      for(const part of Array.from(cell.children)) {
        if(localName(part)==='v') {v=part.textContent; break;}
        if(localName(part)==='is')v=part.textContent;
      }
      if(type==='s' && v!==null)v=strings[Number(v)]??'';
      if(type==='b' && v!==null)v=v==='1'?'TRUE':'FALSE';
      if(v!==null && String(v).trim()!=='')row[col]=String(v).trim();
    }
    if(Object.keys(row).length)result.push({index,row});
  }
  return result;
}
function mapping(rows,namecol,cols){
  return rows.filter(({row})=>at(row,namecol))
    .map(({row})=>Object.fromEntries(Object.entries(cols).map(([k,i])=>[k,at(row,i)])));
}
const convertExcelDate = v => {
  if(!/^\d{5}(\.\d+)?$/.test(v||''))return v||'';
  return new Date((Number(v)-25569)*86400000).toISOString().slice(0,10);
};

export async function importWorkbook(file, progress=()=>{}){
  if(file.size>40*1024*1024)throw new Error('檔案超過 40 MB。請匯入完整但已壓縮的新版 Excel。');
  progress('正在解讀 Excel 結構…');
  const zip=new ZipReader(await file.arrayBuffer());
  const workbook=parseXml(await zip.get('xl/workbook.xml')||'');
  const rel=parseXml(await zip.get('xl/_rels/workbook.xml.rels')||'');
  const targets=new Map(elements(rel,'Relationship').map(e=>[e.getAttribute('Id'),e.getAttribute('Target')]));
  const sheets=new Map();
  for(const e of elements(workbook,'sheet')) {
    const rid=e.getAttribute('r:id')||e.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    const target=targets.get(rid);if(!target)continue;
    const path=target.startsWith('/')?target.slice(1):target.startsWith('xl/')?target:'xl/'+target;
    sheets.set(e.getAttribute('name'),path);
  }
  for(const required of ['Master list','Consolidated list','Input - Strategic list']) {
    if(!sheets.has(required))throw new Error(`這份 Excel 缺少必要工作表：${required}。請確認是 Screening Tool 完整版本。`);
  }
  progress('正在讀取共用文字資料…');
  const strings=[];
  const sxml=await zip.get('xl/sharedStrings.xml');
  if(sxml){
    const doc=parseXml(sxml);
    for(const si of elements(doc,'si')) strings.push(si.textContent||'');
  }
  const rowCache=new Map();
  async function rowsFor(name,from){
    const path=sheets.get(name);if(!path)return [];
    if(!rowCache.has(name)){
      progress(`正在處理：${name}`);
      const xml=await zip.get(path);
      rowCache.set(name,xml?getRows(xml,strings,1):[]);
    }
    return rowCache.get(name).filter(({index})=>index>=from);
  }
  const master=mapping(await rowsFor('Master list',2),3,{id:16,index:0,brand:1,group:2,name:3,tier:4,country:5,higg:7,sources:8,type:9,start:10,end:11,services:12,ctaVersion:13,alias1:14,alias2:15});
  const engagements=mapping(await rowsFor('Consolidated list',2),3,{index:0,brand:1,group:2,name:3,tier:6,country:7,higg:9,source:10,type:11,start:12,end:13,services:14,ctaVersion:15,alias1:16,alias2:17,masterId:18});
  const manufacturers=mapping(await rowsFor('Input - Manufacturer info',3),2,{parent:1,name:2,country:3,higg:4,alias:5});
  const strategic=mapping(await rowsFor('Input - Strategic list',3),1,{name:1,country:2,higg:3,source:4,alias:5});
  const assessments=mapping(await rowsFor('Input - Assessment Records',3),3,{brand:1,group:2,name:3,type:4,country:5,date:7,method:8,engineer:9,comment:17,alias:15});
  const implementations=mapping(await rowsFor('Input - Implementation records',3),3,{brand:2,name:3,tier:4,country:5,higg:6,services:7,status:8,start:9,end:10,owner:11,alias:12});
  const updates=(await rowsFor('Update Log',4)).filter(x=>at(x.row,1)).map(({row})=>({owner:at(row,1),date:convertExcelDate(at(row,2)),
    'CLP':at(row,3),'CTA':at(row,4),'Aii tracker':at(row,5),'NonAii tracker':at(row,6),
    'BPAP':at(row,7),'CTM':at(row,8),'Manufacturer info':at(row,9),'Strategic list':at(row,10),
    'Implementation records':at(row,11),'Assessment records':at(row,12)}));
  const sourceStamps={};
  for(const name of ['Input - CLP','Input - CTA','Input - Aii tracker','Input - NonAii tracker','Input - BPAP','Input - CTM','Input - Manufacturer info','Input - Strategic list','Input - Implementation records','Input - Assessment Records']){
    const r=await rowsFor(name,1),one=r.find(x=>x.index===1);
    sourceStamps[name]=(one?Object.values(one.row).filter(s=>s.includes('Last update:')):[]);
  }
  if(master.length<100 || engagements.length<100)throw new Error('資料筆數異常，已中止更新。請先在 Excel 重新計算並儲存。');
  const hash=globalThis.crypto?.subtle?await crypto.subtle.digest('SHA-256',await file.arrayBuffer()):null;
  const sha256=hash?[...new Uint8Array(hash)].map(n=>n.toString(16).padStart(2,'0')).join(''):'';
  progress('完成資料驗證');
  return {schema:2,meta:{name:file.name,created:new Date().toISOString().slice(0,10),sha256,sourceStamps},master,engagements,manufacturers,strategic,assessments,implementations,updates};
}
export function validateData(data){
  if(!data||data.schema!==2||!Array.isArray(data.master)||!Array.isArray(data.engagements))throw new Error('不是此 App 支援的 Screening 私人備份 JSON。');
  if(data.master.length<100||data.engagements.length<100)throw new Error('資料缺漏；更新已取消。');
  for(const k of ['manufacturers','strategic','assessments','implementations','updates'])if(!Array.isArray(data[k]))data[k]=[];
  return data;
}
