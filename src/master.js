// ================================================================ STR / EIT converters + lookup tables
// ---------------------------------------------------------------- lookup data (embedded, replaceable from Excel)
const LK_KINDS={
  wg:{label:'Config WorkGroup',file:'plan_config_work_group',cols:['DISCIPLINECODE','DISCIPLINENAME','PROJECTTYPE','FACILITYTYPE','WORKGROUPCODE','WORKGROUPNAME','ISACTIVE'],need:['DISCIPLINECODE','WORKGROUPCODE','WORKGROUPNAME']},
  proj:{label:'Project',file:'sys_project_list',cols:['CODE','NAME','PROJECTTYPE'],need:['CODE','PROJECTTYPE']},
  fac:{label:'Facility',file:'sys_facility_list',cols:['PROJECTCODE','CODE','GROUPFACILITYCODE','FACILITYTYPE'],need:['PROJECTCODE','CODE','FACILITYTYPE']}
};
const LK={};
function lkInit(){for(const k in LK_KINDS)LK[k]=ls('lk.'+k)||{rows:LOOKUP_EMBED[k],src:k==='wg'?`bản nhúng sẵn (config ${LOOKUP_EMBED.date})`:`bản nhúng sẵn (xuất ${LOOKUP_EMBED.pfdate})`}}
const blankv=v=>{const s=String(v??'').trim();return s===''||s.toUpperCase()==='NULL'};
let IDX=null;
function buildIdx(){
  const pt=new Map(), ft=new Map(), wg=new Map();
  for(const [c,,t] of LK.proj.rows) if(c&&!blankv(t)) pt.set(norm(c),norm(t));
  for(const [p,c,,t] of LK.fac.rows) if(c&&!blankv(t)) ft.set(norm(p)+'|'+norm(c),norm(t));
  for(const r of LK.wg.rows){
    const [d,,p,f,code,name]=r; if(!d||blankv(name)||num(code)==null) continue;
    const k=[norm(d),norm(p),norm(f),norm(name)].join('|'); let g=wg.get(k); if(!g){g=new Set();wg.set(k,g)} g.add(Math.round(num(code)));
  }
  IDX={pt,ft,wg};
}
const MSG={
  loc:'Không đổi được tên đầy đủ thành mã', wg:'Không tìm thấy WorkGroupName trong config', dup:'Tên trùng nhiều mã trong config',
  fac:'Không tìm thấy FacilityType của Facility trong sys_facility_list', proj:'Không tìm thấy ProjectType của Project trong sys_project_list',
  disc:'Discipline chưa có trong config WorkGroup', rule:'Không xác định được WorkGroup từ loại vật tư', volt:'Không đọc được Voltage để phân LV / HV-MV', noname:'WorkGroupName trống'
};
// returns {code} or {why}
function wgLookup(disc,name,pc,fc,pad){
  const p=IDX.pt.get(norm(pc)); if(!p) return {why:`${MSG.proj} (${pc||'trống'})`};
  const f=IDX.ft.get(norm(pc)+'|'+norm(fc)); if(!f) return {why:`${MSG.fac} (${pc} / ${fc||'trống'})`};
  const g=IDX.wg.get([disc,p,f,norm(name)].join('|'));
  const fmtc=c=>pad?String(c).padStart(3,'0'):String(c);
  if(!g) return {why:`${MSG.wg} (${name}, ${disc} ${p}/${f})`};
  if(g.size>1) return {why:`${MSG.dup} (${name}, ${p}/${f}: ${[...g].map(fmtc).join(' / ')})`};
  return {code:fmtc([...g][0])};
}

// ---------------------------------------------------------------- location
const LOC_MAP=[['Cellar Deck','CED','đoán'],['Mezzanine Deck','MZD','đoán'],['Main Deck','MAD','đoán'],['Sub-Cellar Deck','SCD','đoán'],['Sub Cellar Deck','SCD','đoán'],['Subcellar Deck','SCD','đoán'],
  ['Upper Deck','UPD','đoán (EIT)'],['Lower Deck','LED','đoán (EIT)'],['Weather Deck','WED','đoán (EIT)'],['Middle Deck','MID','đoán (EIT)'],['Roof Deck','ROD','đoán (EIT)'],['Utility Deck','UTD','đoán (EIT)'],['All','ALL','']];
const LOC_CODES=new Set(['CED','MZD','MAD','SCD','UPD','LED','WED','MID','ROD','UTD','CAD','ALL']);
const LOC_IDX=new Map(LOC_MAP.map(([n,c])=>[norm(n),c]));
function locCode(v){const k=norm(v);if(!k)return{code:null};if(LOC_CODES.has(k))return{code:k};if(LOC_IDX.has(k))return{code:LOC_IDX.get(k)};return{code:null,why:MSG.loc}}

// ---------------------------------------------------------------- EIT types
const EIT_TYPES={
  SUPPORT:{label:'Support',dt:'SUPPORT',tpl:'SUPPORT',rule:'ELE/INS 1 Prefab Support; Type có UNIVERSAL → 3 Prefab Universal Support'},
  TRAY:{label:'Tray / Ladder',dt:'TRAY LADDER',tpl:'TRAY',rule:'ELE/INS 5 Tray/Ladder'},
  EQUIPMENT:{label:'Equipment',dt:'EQUIPMENT',tpl:'EQUIPMENT',rule:'ELE 8 Electrical Equipment · INS 8 Instrument Equipment'},
  TUBING:{label:'Tubing',dt:'TUBING FITTING',tpl:'TUBING',rule:'INS 7 Tubing'},
  CABLE:{label:'Cable schedule',dt:'CABLE',tpl:'CABLE',rule:'ELE: Voltage ≤ 1 kV → 7 LV Cable, > 1 kV → 6 HV/MV Cable · INS 6 Cable'}
};
const EIT_RULES=[
  ['Support','Type không có UNIVERSAL','Prefab Electrical Support (Carbon Steel)','Prefab Instrument Support (Carbon Steel)'],
  ['Support','Type có UNIVERSAL','Prefab Universal Support','Prefab Universal Support'],
  ['Tray / Ladder','','Tray/Ladder','Tray/Ladder'],
  ['Equipment','','Electrical Equipment','Instrument Equipment'],
  ['Tubing','','—','Tubing'],
  ['Cable','Voltage lớn nhất ≤ 1 kV','LV Cable','Cable'],
  ['Cable','Voltage lớn nhất > 1 kV','HV/MV Cable','Cable'],
  ['(mọi loại)','DisciplineCode TELECOM, F&G…','chưa có trong config → trống + remark','']
];
const DISC_CODE={ELECTRICAL:'ELE',INSTRUMENT:'INS'};
function eitWgName(type,dc,row,kv){
  if(type==='SUPPORT') return /UNIVERSAL/i.test(String(row.Type??''))?'Prefab Universal Support':(dc==='ELE'?'Prefab Electrical Support (Carbon Steel)':'Prefab Instrument Support (Carbon Steel)');
  if(type==='TRAY') return 'Tray/Ladder';
  if(type==='EQUIPMENT') return dc==='ELE'?'Electrical Equipment':'Instrument Equipment';
  if(type==='TUBING') return dc==='INS'?'Tubing':null;
  if(type==='CABLE'){if(dc==='INS')return'Cable';if(kv==null)return null;return kv<=1?'LV Cable':'HV/MV Cable'}
  return null;
}
function kvOf(v){const m=String(v??'').match(/\d+(\.\d+)?/g);if(!m)return null;let x=Math.max(...m.map(Number));return x>=100?x/1000:x}
const UNIT_N={M:'m',MM:'mm',M2:'m2',EA:'Ea',PCS:'Pcs',SET:'Set'};

// ---------------------------------------------------------------- output specs
const STR_HEADS=['OrderBy','ProjectCode','FacilityCode','MemberSize','ProfileType','WorkReleaseNo','MTOType','PieceMarkNo','AFCDrawingNo','AFCRev','AFCDrawingsTitle','Source','ItemCode','MaterialCode','MaterialType','Description','MaterialGrade','Unit','LengthArea_m_m2','Quantity','NetDesign_m_m2','WeightUnit_kg_m_m2','TotalGrossDesign_MT','SurfacePaiting_m2','Location','WorkGroup','WorkGroupName','ConstructionCode','Remark'];
const EIT_HEADS=['RowIndex','ProjectCode','IsSubcontractor','GroupFacilityCode','FacilityCode','DisciplineCode','DisciplineType','Type','TagName','ItemCode','MaterialCode','MaterialDescription','Unit','NetQty','Contingency','ConsAllowance','PTSCMCStockQty','FinalQty','UnitWeight_kg_m3','TotalWeight_MT','Location','Remark','FabricationType','Scope','SystemNo','SystemDescription','SubSystemNo','SubSystemDescription','WorkGroup','WorkGroupName','ConstructionCode','CablePulling_Mhrs','CableGlandingFrom_Mhrs','CableGlandingTo_Mhrs','CableTerminatingFrom_Mhrs','CableTerminatingTo_Mhrs'];
const HELPER_HEADS=['Location (tên đầy đủ)','Remark convert'];
// column rule: {t, src (input column or null), note, kind: copy|blank|const|calc}
function specOf(kind,type){
  const C=(t,src,note,k='copy',v)=>({t,src,note,k,v});
  if(kind==='str'){
    const special={
      OrderBy:C('OrderBy',null,'Để trống','blank'), Quantity:C('Quantity',null,'Để trống (theo thống nhất)','blank'),
      SurfacePaiting_m2:C('SurfacePaiting_m2',null,'Luôn = 0','const',0), ConstructionCode:C('ConstructionCode',null,'Để trống','blank'),
      Location:C('Location','Location','Tên đầy đủ → mã (bảng Location)','calc'),
      WorkGroup:C('WorkGroup','WorkGroupName','Tra config: ProjectType + FacilityType + STR + WorkGroupName → mã 3 số (001)','calc'),
      WorkGroupName:C('WorkGroupName','WorkGroupName','Giữ nguyên (cột thêm vào template)')
    };
    return {heads:STR_HEADS,sheet:'DATA',cols:STR_HEADS.map(t=>special[t]||C(t,t,'Cùng tên')),
      keys:['PROJECTCODE','MEMBERSIZE','ITEMCODE','MATERIALCODE','WORKGROUPNAME'],req:['ProjectCode','FacilityCode','WorkGroupName']};
  }
  const cab=type==='CABLE', dt=EIT_TYPES[type].dt;
  const blank=t=>C(t,null,'Để trống (input không có)','blank');
  const m={
    RowIndex:C('RowIndex',null,'Để trống (hệ thống tự sinh)','blank'),
    IsSubcontractor:C('IsSubcontractor','IsSubContractor','Cùng tên'),
    DisciplineCode:cab?C('DisciplineCode','DisciplineCable','Bỏ chữ "Cable", viết hoa (Electrical Cable → ELECTRICAL)','calc'):C('DisciplineCode','DisciplineCode','Viết hoa','calc'),
    DisciplineType:C('DisciplineType',null,`Tự điền "${dt}" cho mọi dòng (để lọc giống web)`,'const',dt),
    Type:cab?C('Type','Voltage_V','Electrical: ≤ 1 kV → LV Cable, > 1 kV → HV/MV Cable; Instrument / khác → Cable','calc'):C('Type','Type','Cùng tên'),
    TagName:cab?C('TagName','CableName','Lấy CableName'):C('TagName','TagName','Cùng tên'),
    MaterialDescription:cab?C('MaterialDescription','Description','Lấy Description'):C('MaterialDescription','MaterialDescription','Cùng tên'),
    Unit:cab?C('Unit',null,'Luôn = m','const','m'):C('Unit','Unit','Chuẩn hoa/thường: M→m, EA→Ea, MM→mm, M2→m2 (không đổi đơn vị)','calc'),
    NetQty:cab?C('NetQty','Length_Mtrs','Lấy Length_Mtrs'):C('NetQty','NetQty','Cùng tên'),
    Contingency:blank('Contingency'),ConsAllowance:blank('ConsAllowance'),PTSCMCStockQty:blank('PTSCMCStockQty'),
    FinalQty:C('FinalQty',cab?'Length_Mtrs':'NetQty','= NetQty','calc'),
    UnitWeight_kg_m3:cab?blank('UnitWeight_kg_m3'):C('UnitWeight_kg_m3','UnitWeight_kg_m3','Cùng tên'),
    TotalWeight_MT:cab?blank('TotalWeight_MT'):C('TotalWeight_MT','TotalWeight_Kg','= TotalWeight_Kg / 1000','calc'),
    Location:cab?C('Location',null,'Để trống (cable chỉ có From/To Location là tên phòng)','blank'):C('Location','Location','Tên đầy đủ → mã (bảng Location)','calc'),
    FabricationType:blank('FabricationType'),
    Scope:C('Scope','ScopeOfSupply','Lấy ScopeOfSupply'),
    SubSystemNo:blank('SubSystemNo'),SubSystemDescription:blank('SubSystemDescription'),
    WorkGroup:C('WorkGroup',null,'Tra config: ProjectType + FacilityType + ELE/INS + WorkGroupName → mã (số trơn)','calc'),
    WorkGroupName:C('WorkGroupName',null,'Suy theo loại: '+EIT_TYPES[type].rule,'calc'),
    ConstructionCode:blank('ConstructionCode')
  };
  for(const t of ['CablePulling_Mhrs','CableGlandingFrom_Mhrs','CableGlandingTo_Mhrs','CableTerminatingFrom_Mhrs','CableTerminatingTo_Mhrs']) m[t]=C(t,null,'Để trống (hệ thống tính)','blank');
  return {heads:EIT_HEADS,sheet:'Sheet1',cols:EIT_HEADS.map(t=>m[t]||C(t,t,'Cùng tên')),
    keys:['PROJECTCODE','TAGNAME','CABLENAME','ITEMCODE','MATERIALCODE'],
    req:cab?['ProjectCode','FacilityCode','DisciplineCable','CableName','Length_Mtrs']:['ProjectCode','FacilityCode','DisciplineCode','NetQty']};
}

// ---------------------------------------------------------------- state & UI scaffold
const MS={str:{kind:'str'},eit:{kind:'eit',type:ls('eitType')||'SUPPORT'}};
function dropHtml(k,title){return `
  <label class="drop" id="${k}-drop" for="${k}-file">
    <svg class="pipe" viewBox="0 0 48 48" aria-hidden="true"><rect x="6" y="8" width="36" height="6" fill="var(--accent)"/><rect x="21" y="14" width="6" height="24" fill="var(--muted)"/><rect x="6" y="38" width="36" height="6" fill="var(--accent)"/></svg>
    <div style="min-width:0;display:grid;gap:2px"><span class="fname" id="${k}-fname">${title}</span><span class="muted" id="${k}-fmeta">.xlsx · .xlsm · .xls · .csv</span></div>
    <span class="btn primary">Chọn file…</span>
  </label><input type="file" id="${k}-file" accept=".xlsx,.xlsm,.xls,.csv" hidden>`}
function mapPanel(k,n){return `
  <div id="${k}-err" class="banner bad" hidden></div>
  <div class="panel"><div class="ph"><h2><span class="step">${n}</span>Cách map cột</h2><div class="chips" id="${k}-chips"></div></div>
    <div class="pb tscroll"><table class="maptable"><thead><tr><th>Cột output</th><th>Lấy từ cột input</th><th>Xử lý</th></tr></thead><tbody id="${k}-map"></tbody></table></div></div>
  <div class="panel"><div class="ph"><h2><span class="step">${String(+n+1).padStart(2,'0')}</span>Convert &amp; tải file</h2></div>
    <div class="pb" style="display:grid;gap:16px">
      <div class="runbar"><button class="btn primary" id="${k}-run" disabled>Convert</button><span class="muted mono" id="${k}-msg">Chọn file để bắt đầu</span></div>
      <div id="${k}-res" hidden style="display:grid;gap:16px">
        <div class="tiles" id="${k}-tiles"></div>
        <div class="btns"><button class="btn primary" id="${k}-save">Tải file kết quả (.xlsx)</button><span class="muted" id="${k}-hint" style="align-self:center"></span></div>
        <div style="display:grid;gap:8px"><h3 id="${k}-isst"></h3><div class="restable"><table><thead><tr><th>Vấn đề (cột Remark convert)</th><th>Số dòng</th></tr></thead><tbody id="${k}-iss"></tbody></table></div></div>
        <div style="display:grid;gap:8px"><h3>Xem trước (tối đa 50 dòng)</h3><div class="restable" id="${k}-prev"></div></div>
      </div>
    </div></div>`}
function buildMasterUI(){
  $('#sec-str').innerHTML=`
    <div class="panel"><div class="ph"><h2><span class="step">01</span>Chọn file Structural MTO Master</h2><button class="btn small" data-tpl="STR">Tải template input STR</button></div>
      <div class="pb" style="display:grid;gap:10px">${dropHtml('str','Kéo-thả file Structural MTO Master vào đây hoặc bấm để chọn')}
        <p class="muted" style="font-size:13px">Kết quả: sheet <span class="mono">DATA</span> theo template StructuralMTOMaster (29 cột, thêm WorkGroupName ngay sau WorkGroup) + 2 cột phụ màu xám cần xóa trước khi nạp. Sheet <span class="mono">Map</span> ghi lại cách map và tóm tắt.</p></div></div>
    ${mapPanel('str','02')}`;
  $('#sec-eit').innerHTML=`
    <div class="panel"><div class="ph"><h2><span class="step">01</span>Chọn loại EIT</h2><span class="muted" id="eit-typenote"></span></div>
      <div class="pb" style="display:grid;gap:12px">
        <div class="seg" role="radiogroup" aria-label="Loại EIT" id="eit-types">${Object.entries(EIT_TYPES).map(([k,t])=>`<button role="radio" data-type="${k}" aria-checked="false"><b>${t.label}</b><small>DisciplineType = ${t.dt}</small></button>`).join('')}</div>
        <div class="typecard" id="eit-typeinfo"></div>
      </div></div>
    <div class="panel"><div class="ph"><h2><span class="step">02</span>Chọn file EIT MTO Master</h2></div>
      <div class="pb" style="display:grid;gap:10px">${dropHtml('eit','Kéo-thả file EIT MTO Master vào đây hoặc bấm để chọn')}
        <p class="muted" style="font-size:13px">Chọn file xong, trang tự nhận loại (Cable / Support / Tray / Equipment / Tubing) theo cột và giá trị DisciplineType; bạn vẫn đổi được ở bước 01. Kết quả: layout mtoc_eit_mto_master (36 cột, thêm WorkGroupName sau WorkGroup) + 2 cột phụ màu xám.</p></div></div>
    ${mapPanel('eit','03')}`;
  for(const k of ['str','eit']){
    const drop=$(`#${k}-drop`);
    ['dragenter','dragover'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.add('over')}));
    ['dragleave','drop'].forEach(e=>drop.addEventListener(e,ev=>{ev.preventDefault();drop.classList.remove('over')}));
    drop.addEventListener('drop',ev=>{const f=ev.dataTransfer.files[0];if(f)loadMaster(k,f)});
    $(`#${k}-file`).addEventListener('change',e=>{const f=e.target.files[0];if(f)loadMaster(k,f);e.target.value=''});
    $(`#${k}-run`).addEventListener('click',()=>runMaster(k));
    $(`#${k}-save`).addEventListener('click',()=>{const r=MS[k].result;if(r)saveFile(r.name,r.bytes)});
  }
  $('#eit-types').addEventListener('click',e=>{const b=e.target.closest('button[data-type]');if(!b)return;setEitType(b.dataset.type,'chọn tay')});
  document.addEventListener('click',e=>{const b=e.target.closest('[data-tpl]');if(b)downloadTemplate(b.dataset.tpl)});
  setEitType(MS.eit.type,null); renderMasterMap('str');
}
function setEitType(t,why){
  const st=MS.eit; st.type=t; ls('eitType',t);
  for(const b of document.querySelectorAll('#eit-types button'))b.setAttribute('aria-checked',b.dataset.type===t);
  const T=EIT_TYPES[t];
  $('#eit-typeinfo').innerHTML=`<div><div class="k">DisciplineType tự điền</div><div class="v">${esc(T.dt)}</div></div><div><div class="k">WorkGroup</div><div class="v">${esc(T.rule)}</div></div><div style="align-self:center"><button class="btn small" data-tpl="${T.tpl}">Tải template input ${esc(T.label)}</button></div>`;
  $('#eit-typenote').textContent=why?`Loại hiện tại: ${T.label} (${why})`:'';
  st.result=null; $('#eit-res').hidden=true;
  renderMasterMap('eit');
}

// ---------------------------------------------------------------- reading input
function findHeader(aoa){
  for(let r=0;r<Math.min(10,aoa.length);r++){const row=aoa[r]||[];if(row.some(v=>norm(v)==='PROJECTCODE'))return r}
  return -1;
}
async function loadMaster(k,file){
  const st=MS[k];
  $(`#${k}-fname`).textContent=file.name; $(`#${k}-fmeta`).innerHTML='<span class="spin"></span> Đang đọc file…';
  $(`#${k}-res`).hidden=true; st.result=null; await tick();
  try{
    const wb=await readWorkbook(file); let pick=null;
    for(const name of wb.SheetNames){const ws=wb.Sheets[name];if(!ws||!ws['!ref'])continue;const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,blankrows:true});const hr=findHeader(aoa);if(hr>=0){pick={name,aoa,hr};break}}
    if(!pick){st.file={name:file.name};st.aoa=null;$(`#${k}-fmeta`).textContent='Không tìm thấy dòng tiêu đề (cột ProjectCode) trong 10 dòng đầu của các sheet.';renderMasterMap(k);return}
    const ih=new Map(); (pick.aoa[pick.hr]||[]).forEach((v,c)=>{const n=norm(v);if(n&&!ih.has(n))ih.set(n,c)});
    Object.assign(st,{file:{name:file.name,size:file.size},sheet:pick.name,aoa:pick.aoa,hr:pick.hr,ih});
    const nData=pick.aoa.length-pick.hr-1;
    $(`#${k}-fmeta`).textContent=`${(file.size/1048576).toFixed(2)} MB · sheet ${pick.name} · tiêu đề dòng ${pick.hr+1} · ${fmt(Math.max(0,nData))} dòng dưới tiêu đề`;
    if(k==='eit'){const g=guessEitType(st,file.name);if(g)setEitType(g.type,'tự nhận theo '+g.why);else renderMasterMap(k)}
    else renderMasterMap(k);
  }catch(e){$(`#${k}-fmeta`).textContent='Không đọc được file: '+e.message;toast('Không đọc được file. Kiểm tra file có đang được bảo vệ mật khẩu hoặc bị hỏng không.')}
}
function guessEitType(st,fname){
  if(st.ih.has('CABLENAME')||st.ih.has('DISCIPLINECABLE'))return{type:'CABLE',why:'cột CableName'};
  const c=st.ih.get('DISCIPLINETYPE');
  if(c!=null){const cnt={};for(let r=st.hr+1;r<st.aoa.length;r++){const v=norm((st.aoa[r]||[])[c]);if(!v)continue;const t=/SUPPORT/.test(v)?'SUPPORT':/TRAY|LADDER/.test(v)?'TRAY':/EQUIP/.test(v)?'EQUIPMENT':/TUB/.test(v)?'TUBING':/CABLE/.test(v)?'CABLE':null;if(t)cnt[t]=(cnt[t]||0)+1}
    const best=Object.entries(cnt).sort((a,b)=>b[1]-a[1])[0]; if(best)return{type:best[0],why:'giá trị DisciplineType'}}
  const f=fname.toUpperCase();
  for(const [t,re] of [['CABLE',/CABLE/],['SUPPORT',/SUPPORT/],['TRAY',/TRAY|LADDER/],['EQUIPMENT',/EQUIP/],['TUBING',/TUB/]]) if(re.test(f))return{type:t,why:'tên file'};
  return null;
}

// ---------------------------------------------------------------- mapping view
function renderMasterMap(k){
  const st=MS[k], sp=specOf(k,st.type), has=st.ih||new Map(), loaded=!!st.aoa;
  let h='';
  for(const c of sp.cols){
    const found=c.src&&has.has(norm(c.src)), srcTxt=c.src?(loaded?(found?`${c.src} · cột ${colLetter(has.get(norm(c.src))+1)}`:`${c.src} (không có trong file)`):c.src):'—';
    let lvl=c.k==='blank'?'none':c.k==='copy'?(loaded&&!found?'warn':'ok'):'auto';
    if(loaded&&c.src&&!found&&sp.req.includes(c.src))lvl='req';
    const note=c.k==='copy'&&loaded&&!found?'Không có trong file → để trống':c.note;
    h+=`<tr><td class="t">${esc(c.t)}</td><td class="mono" style="font-size:12.5px">${esc(srcTxt)}</td><td><span class="pill ${lvl}">${esc(note)}</span></td></tr>`;
  }
  h+=`<tr class="helper"><td class="t">(phụ) Location (tên đầy đủ)</td><td class="mono" style="font-size:12.5px">Location</td><td><span class="pill none">Giữ tên gốc để đối chiếu · xóa trước khi nạp</span></td></tr><tr class="helper"><td class="t">(phụ) Remark convert</td><td>—</td><td><span class="pill none">Lý do ô bị tô hồng · xóa trước khi nạp</span></td></tr>`;
  $(`#${k}-map`).innerHTML=h;
  const err=$(`#${k}-err`), run=$(`#${k}-run`);
  if(!loaded){$(`#${k}-chips`).innerHTML='<span class="chip">chưa chọn file</span>';err.hidden=true;run.disabled=true;$(`#${k}-msg`).textContent='Chọn file để bắt đầu';return}
  const miss=sp.req.filter(c=>!has.has(norm(c)));
  const used=new Set(sp.cols.map(c=>c.src&&norm(c.src)).filter(Boolean)); used.add('LOCATION');
  const unused=[...has.keys()].filter(x=>!used.has(x));
  $(`#${k}-chips`).innerHTML=`<span class="chip">sheet <b>${esc(st.sheet)}</b></span><span class="chip">tiêu đề dòng <b>${st.hr+1}</b></span><span class="chip">cột input không dùng: <b>${unused.length}</b></span>`;
  if(miss.length){
    err.innerHTML=`<h3>Không map được file này${k==='eit'?` theo loại ${esc(EIT_TYPES[st.type].label)}`:''}</h3><p>Thiếu cột bắt buộc: ${miss.map(c=>`<code>${esc(c)}</code>`).join(' ')}</p><p>Các cột tìm thấy ở dòng ${st.hr+1}: <span class="mono" style="font-size:12.5px">${esc([...has.keys()].join(', '))}</span></p><p>${k==='eit'?'Kiểm tra lại loại EIT ở bước 01, hoặc ':''}tải template input để xem đúng tên cột.</p>`;
    err.hidden=false; run.disabled=true; $(`#${k}-msg`).textContent='Thiếu cột bắt buộc';
  } else {err.hidden=true;run.disabled=false;$(`#${k}-msg`).textContent='Sẵn sàng'+(unused.length?` · cột input không dùng: ${unused.join(', ')}`:'')}
}

// ---------------------------------------------------------------- convert
function cellv(v){if(v==null)return null;if(typeof v==='number')return isFinite(v)?v:null;const s=String(v).trim();return(s===''||s.toUpperCase()==='NULL')?null:s}
const numOr=v=>{const n=num(v);return n!=null?n:v};
function convertMaster(k){
  const st=MS[k], type=st.type, sp=specOf(k,type), ih=st.ih, aoa=st.aoa, H=sp.heads, nH=H.length;
  const ix=Object.fromEntries(H.map((t,i)=>[t,i]));
  const recs=[], issues=new Map(); const stat={rows:0,locOk:0,locMiss:0,wgOk:0,wgMiss:0,dtChanged:0};
  const addIss=w=>issues.set(w,(issues.get(w)||0)+1);
  for(let r=st.hr+1;r<aoa.length;r++){
    const src=aoa[r]||[]; const g=name=>{const c=ih.get(norm(name));return c==null?null:cellv(src[c])};
    if(!sp.keys.some(x=>ih.has(x)&&cellv(src[ih.get(x)])!=null)) continue;
    const o=new Array(nH+2).fill(null), pink=new Set(), flags=[];
    for(const c of sp.cols){ if(c.k==='copy'&&c.src)o[ix[c.t]]=g(c.src); else if(c.k==='const')o[ix[c.t]]=c.v }
    // location
    let locFull=null;
    if(sp.cols[ix.Location].k!=='blank'){
      locFull=g('Location'); const L=locCode(locFull);
      if(locFull!=null){if(L.code){o[ix.Location]=L.code;stat.locOk++}else{pink.add(ix.Location);flags.push(L.why);stat.locMiss++}}
    }
    let why=null;
    if(k==='str'){
      const name=g('WorkGroupName'); o[ix.WorkGroupName]=name;
      if(name==null){why=MSG.noname}
      else{const w=wgLookup('STR',name,g('ProjectCode'),g('FacilityCode'),true);if(w.code)o[ix.WorkGroup]=w.code;else why=w.why}
    } else {
      const cab=type==='CABLE';
      const disc=cab?norm(String(g('DisciplineCable')??'').replace(/\s*cable\s*$/i,'')):norm(g('DisciplineCode'));
      o[ix.DisciplineCode]=disc||null;
      const inDT=g('DisciplineType'); if(inDT!=null&&norm(inDT)!==EIT_TYPES[type].dt)stat.dtChanged++;
      let kv=null;
      if(cab){
        kv=kvOf(g('Voltage_V'));
        o[ix.Type]=disc==='ELECTRICAL'?(kv==null?null:(kv<=1?'LV Cable':'HV/MV Cable')):'Cable';
        o[ix.NetQty]=numOr(g('Length_Mtrs'));
        if(disc==='ELECTRICAL'&&kv==null)flags.push(`${MSG.volt} (${g('Voltage_V')??'trống'})`);
      } else {
        const u=norm(g('Unit')); o[ix.Unit]=UNIT_N[u]||g('Unit');
        o[ix.NetQty]=numOr(g('NetQty'));
        const tw=num(g('TotalWeight_Kg')); o[ix.TotalWeight_MT]=tw!=null?Math.round(tw/1000*1e9)/1e9:null;
        for(const t of ['UnitWeight_kg_m3'])o[ix[t]]=numOr(o[ix[t]]);
      }
      o[ix.FinalQty]=o[ix.NetQty];
      const dc=DISC_CODE[disc];
      const name=dc?eitWgName(type,dc,{Type:o[ix.Type]},kv):null; o[ix.WorkGroupName]=name;
      if(!dc)why=`${MSG.disc} (${disc||'trống'})`;
      else if(!name){if(!(type==='CABLE'&&kv==null))why=`${MSG.rule} (${EIT_TYPES[type].label}, ${disc})`}
      else{const w=wgLookup(dc,name,g('ProjectCode'),g('FacilityCode'),false);if(w.code)o[ix.WorkGroup]=w.code;else why=w.why}
    }
    if(o[ix.WorkGroup]!=null)stat.wgOk++;else{stat.wgMiss++;pink.add(ix.WorkGroup)}
    if(why)flags.push(why);
    flags.forEach(addIss);
    o[nH]=locFull; o[nH+1]=flags.join('; ')||null;
    recs.push({o,pink}); stat.rows++;
  }
  return {recs,issues:[...issues.entries()].sort((a,b)=>b[1]-a[1]),stat,sp};
}
function masterXlsx(k,res){
  const st=MS[k], sp=res.sp, H=[...sp.heads,...HELPER_HEADS], nH=sp.heads.length, n=res.recs.length;
  const textCols=new Set(['WorkGroup','ItemCode','MaterialCode','TagName','SystemNo'].map(t=>sp.heads.indexOf(t)).filter(i=>i>=0));
  const cellAt=(r,c)=>{if(r===0)return H[c];const v=res.recs[r-1].o[c];return(v!=null&&textCols.has(c))?String(v):v};
  const styleAt=(r,c)=>{if(r===0)return c>=nH?xf(5,1):xf(0,1);const o=res.recs[r-1];if(o.pink.has(c))return xf(4,0);if(c===nH+1&&o.o[c])return xf(0,2);return 0};
  const widths=H.map((h,c)=>{let w=String(h).length;for(let i=0;i<Math.min(n,300);i++){const v=res.recs[i].o[c];if(v!=null)w=Math.max(w,String(v).length)}return Math.min(Math.max(w+2,8),c===nH+1?70:45)});
  const sst=newSst();
  const main=sheetXml(n+1,H.length,cellAt,styleAt,widths,true,sst);
  const s=res.stat;
  const aoa=[['Cột output','Lấy từ cột input','Xử lý'],...sp.cols.map(c=>[c.t,c.src?(st.ih.has(norm(c.src))?c.src:`${c.src} (không có)`):'—',c.k==='copy'&&c.src&&!st.ih.has(norm(c.src))?'Không có trong file → để trống':c.note]),
    ['(phụ) Location (tên đầy đủ)','Location','Xóa trước khi nạp'],['(phụ) Remark convert','—','Xóa trước khi nạp'],[],
    ['TÓM TẮT'],['File nguồn',st.file.name],['Sheet / dòng tiêu đề',`${st.sheet} / dòng ${st.hr+1}`],
    ...(k==='eit'?[['Loại EIT',`${EIT_TYPES[st.type].label} → DisciplineType = ${EIT_TYPES[st.type].dt}`],['Dòng có DisciplineType khác trong file gốc',s.dtChanged]]:[]),
    ['Số dòng xuất',s.rows],['Location đổi được / không đổi được',`${s.locOk} / ${s.locMiss}`],['WorkGroup map được / trống',`${s.wgOk} / ${s.wgMiss}`],
    ['Config WorkGroup',LK.wg.src],['Project list',LK.proj.src],['Facility list',LK.fac.src],[],
    ['VẤN ĐỀ','Số dòng'],...res.issues.map(([w,c])=>[w,c])];
  const nMap=sp.cols.length+2;
  const mapXml=sheetXml(aoa.length,3,(r,c)=>(aoa[r]||[])[c],(r,c)=>r===0?xf(5,1):(r===nMap+2||r===aoa.length-res.issues.length-1)?xf(0,1):0,[34,40,80],true,sst);
  return buildXlsx([{name:sp.sheet,xml:main},{name:'Map',xml:mapXml}],sst);
}
async function runMaster(k){
  const st=MS[k]; if(!st.aoa)return;
  const btn=$(`#${k}-run`); btn.disabled=true; $(`#${k}-msg`).textContent='Đang convert…'; await tick();
  try{
    const t0=performance.now(); const res=convertMaster(k);
    if(!res.recs.length)throw new Error('Không có dòng dữ liệu nào để xuất.');
    res.bytes=masterXlsx(k,res);
    const base=st.file.name.replace(/\.[^.]+$/,''); res.name=`${base} - Output.xlsx`;
    st.result=res; renderMasterResult(k,res);
    $(`#${k}-msg`).textContent=`Xong · ${fmt(res.recs.length)} dòng · ${((performance.now()-t0)/1000).toFixed(1)} s`;
  }catch(e){$(`#${k}-msg`).textContent='Lỗi: '+e.message;toast('Lỗi khi convert: '+e.message)}
  finally{btn.disabled=false}
}
function renderMasterResult(k,res){
  const s=res.stat, st=MS[k];
  const t=[[fmt(s.rows),'Số dòng xuất',''],[fmt(s.wgOk),'WorkGroup map được',''],[fmt(s.wgMiss),'WorkGroup trống (tô hồng)','p'],[fmt(s.locOk),'Location đổi được',''],[fmt(s.locMiss),'Location không đổi được (tô hồng)','p']];
  if(k==='eit')t.push([fmt(s.dtChanged),`Dòng có DisciplineType khác "${EIT_TYPES[st.type].dt}" trong file gốc (đã ghi đè)`,'']);
  $(`#${k}-tiles`).innerHTML=t.map(([n,l,sw])=>`<div class="tile"><span class="n">${n}</span><span class="k">${sw?`<span class="sw ${sw}"></span>`:''}${esc(l)}</span></div>`).join('');
  $(`#${k}-hint`).textContent=`${res.name} · sheet ${res.sp.sheet} + Map · xóa 2 cột phụ màu xám trước khi nạp hệ thống`;
  $(`#${k}-isst`).textContent=res.issues.length?`Vấn đề cần xem (${res.issues.length} loại)`:'Không có vấn đề: mọi dòng đều map được Location và WorkGroup';
  $(`#${k}-iss`).innerHTML=res.issues.map(([w,c])=>`<tr><td class="wrap">${esc(w)}</td><td class="num">${fmt(c)}</td></tr>`).join('');
  const H=[...res.sp.heads,...HELPER_HEADS], nH=res.sp.heads.length;
  const rows=res.recs.slice(0,50);
  $(`#${k}-prev`).innerHTML=`<table><thead><tr>${H.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(o=>`<tr>${H.map((h,c)=>`<td class="${o.pink.has(c)?'pk':''}${c===nH+1&&o.o[c]?' rk':''}">${esc(o.o[c]??'')}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  $(`#${k}-res`).hidden=false;
}

// ---------------------------------------------------------------- input templates
function downloadTemplate(key){
  const T=INPUT_TPL[key]; if(!T)return;
  const name=key==='STR'?'Structural MTO Master - template.xlsx':`EIT MTO Master (${EIT_TYPES[key]?.label||key}) - template.xlsx`;
  const sst=newSst(), nc=T.heads.length, w=T.heads.map(h=>Math.max(12,String(h).length+3));
  const tplXml=sheetXml(2,nc,(r,c)=>r===0?T.flags[c]:T.heads[c],(r,c)=>r===0?(T.flags[c]?xf(2,0):0):xf(5,1),w,false,sst);
  const ex=[T.flags,T.heads,T.example];
  const exXml=sheetXml(3,nc,(r,c)=>ex[r][c],(r,c)=>r===0?(T.flags[c]?xf(2,0):0):r===1?xf(5,1):0,w,false,sst);
  const notes=[['Ghi chú'],['Dòng 1: cờ "Must meet" = cột bắt buộc nhập. Dòng 2: tên cột, không đổi tên.'],['Nhập dữ liệu từ dòng 3. Sheet "Ví dụ" có 1 dòng mẫu.'],
    ...(key==='STR'?[['WorkGroupName: nhập đúng tên như bảng config (tab Bảng config & tra cứu, Discipline STR).'],['Location: nhập tên deck (Cellar deck, Main deck…) hoặc mã (CED, MAD…).']]
      :[[`DisciplineType: tool tự điền "${EIT_TYPES[key].dt}" khi convert loại ${EIT_TYPES[key].label}.`],['DisciplineCode: ELECTRICAL / INSTRUMENT (TELECOM, F&G chưa có trong config WorkGroup).'],...(key==='CABLE'?[['Voltage_V: ghi dạng 0.6/1 (kV) để tool phân LV / HV-MV.']]:[['Location: nhập tên deck (Main Deck…) hoặc mã (MAD…).']])])];
  const noteXml=sheetXml(notes.length,1,(r,c)=>notes[r][0],(r)=>r===0?xf(5,1):0,[100],false,sst);
  saveFile(name,buildXlsx([{name:'Sheet1',xml:tplXml},{name:'Ví dụ',xml:exXml},{name:'Ghi chú',xml:noteXml}],sst));
}

// ---------------------------------------------------------------- lookup tab
function renderLookup(){
  const t=[[fmt(LK.wg.rows.length),'dòng config WorkGroup',LK.wg.src],[fmt(LK.proj.rows.length),'project',LK.proj.src],[fmt(LK.fac.rows.length),'facility',LK.fac.src]];
  $('#lkTiles').innerHTML=t.map(([n,k,s])=>`<div class="tile"><span class="n">${n}</span><span class="k">${esc(k)}</span><span class="k mono" style="font-size:11.5px">${esc(s)}</span></div>`).join('');
  $('#lkRule').innerHTML=EIT_RULES.map(r=>`<tr>${r.map(x=>`<td>${esc(x)}</td>`).join('')}</tr>`).join('');
  $('#locBody').innerHTML=LOC_MAP.map(([n,c,s])=>`<tr><td>${esc(n)}</td><td class="mono">${c}</td><td class="muted">${esc(s)}</td></tr>`).join('');
  const opt=(id,vals,first)=>{const el=$(id),cur=el.value;el.innerHTML=`<option value="">${first}</option>`+vals.map(v=>`<option>${esc(v)}</option>`).join('');if(vals.includes(cur))el.value=cur};
  const uniq=i=>[...new Set(LK.wg.rows.map(r=>norm(r[i])).filter(Boolean))].sort();
  opt('#wgDisc',uniq(0),'Tất cả'); if(!$('#wgDisc').dataset.init){$('#wgDisc').value='STR';$('#wgDisc').dataset.init=1}
  opt('#wgPT',uniq(2),'Tất cả'); opt('#wgFT',uniq(3),'Tất cả');
  renderWg(); renderPj(); renderFc();
}
function renderWg(){
  const d=$('#wgDisc').value,p=$('#wgPT').value,f=$('#wgFT').value,q=norm($('#wgQ').value),dupOnly=$('#wgDup').checked;
  const rows=LK.wg.rows.filter(r=>(!d||norm(r[0])===d)&&(!p||norm(r[2])===p)&&(!f||norm(r[3])===f)&&(!q||norm(r[5]).includes(q)||String(r[4]??'')===q));
  const out=[];const seen=new Set();
  for(const r of rows){const k=[norm(r[0]),norm(r[2]),norm(r[3]),norm(r[5])].join('|');const g=IDX.wg.get(k);const dup=g&&g.size>1;if(dupOnly&&!dup)continue;const sig=k+'|'+r[4];const rep=seen.has(sig);seen.add(sig);out.push({r,dup,rep})}
  $('#wgCount').textContent=`${fmt(out.length)} dòng${out.length>600?' (hiện 600 dòng đầu)':''}`;
  $('#wgBody').innerHTML=out.slice(0,600).map(({r,dup,rep})=>`<tr><td class="mono">${esc(r[0])}</td><td class="mono">${esc(r[2])}</td><td class="mono">${esc(r[3])}</td><td class="mono">${esc(r[4])}</td><td>${esc(r[5])}</td><td class="mono">${esc(r[6])}</td><td>${dup?'<span class="pill req">trùng tên, khác mã</span>':rep?'<span class="pill none">dòng lặp</span>':''}</td></tr>`).join('');
}
function renderPj(){const q=norm($('#pjQ').value);$('#pjBody').innerHTML=LK.proj.rows.filter(r=>!q||norm(r[0]).includes(q)||norm(r[1]).includes(q)).map(r=>`<tr><td class="mono">${esc(r[0])}</td><td>${esc(r[1])}</td><td class="mono">${esc(r[2]??'')||'<span class="pill warn">trống</span>'}</td></tr>`).join('')}
function renderFc(){const q=norm($('#fcQ').value),nl=$('#fcNull').checked;$('#fcBody').innerHTML=LK.fac.rows.filter(r=>(!q||norm(r[0]).includes(q)||norm(r[1]).includes(q))&&(!nl||blankv(r[3]))).map(r=>`<tr><td class="mono">${esc(r[0])}</td><td class="mono">${esc(r[1])}</td><td class="mono">${esc(r[2]??'')}</td><td class="mono">${blankv(r[3])?'<span class="pill warn">NULL</span>':esc(r[3])}</td></tr>`).join('')}
async function importLookup(files){
  const done=[];
  for(const file of files){
    try{
      const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
      const aoa=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:null});
      const h=new Map();(aoa[0]||[]).forEach((v,c)=>{const n=norm(v);if(n&&!h.has(n))h.set(n,c)});
      const kind=h.has('WORKGROUPCODE')?'wg':h.has('PROJECTCODE')&&h.has('FACILITYTYPE')?'fac':h.has('PROJECTTYPE')&&h.has('CODE')?'proj':null;
      if(!kind){toast(`${file.name}: không nhận ra loại file (cần plan_config_work_group, sys_project_list hoặc sys_facility_list)`);continue}
      const K=LK_KINDS[kind];const miss=K.need.filter(c=>!h.has(c));if(miss.length){toast(`${file.name}: thiếu cột ${miss.join(', ')}`);continue}
      const rows=aoa.slice(1).map(r=>K.cols.map(c=>h.has(c)?cellv(r[h.get(c)]):null)).filter(r=>r.some(v=>v!=null));
      LK[kind]={rows,src:`${file.name} (nạp ${new Date().toLocaleDateString('vi-VN')})`};ls('lk.'+kind,LK[kind]);done.push(`${K.label}: ${rows.length} dòng`);
    }catch(e){toast(`${file.name}: không đọc được (${e.message})`)}
  }
  if(done.length){buildIdx();renderLookup();toast('Đã cập nhật '+done.join(' · '))}
}

// ---------------------------------------------------------------- boot
lkInit(); buildIdx(); buildMasterUI(); renderLookup();
$('#lkLoad').addEventListener('click',()=>$('#lkFile').click());
$('#lkFile').addEventListener('change',e=>{const f=[...e.target.files];if(f.length)importLookup(f);e.target.value=''});
$('#lkReset').addEventListener('click',()=>{for(const k in LK_KINDS)lsDel('lk.'+k);lkInit();buildIdx();renderLookup();toast('Đang dùng dữ liệu nhúng sẵn')});
for(const id of ['#wgDisc','#wgPT','#wgFT','#wgDup'])$(id).addEventListener('change',renderWg);
$('#wgQ').addEventListener('input',renderWg); $('#pjQ').addEventListener('input',renderPj); $('#fcQ').addEventListener('input',renderFc); $('#fcNull').addEventListener('change',renderFc);
if(typeof window.__selftest==='function') window.__selftest();
