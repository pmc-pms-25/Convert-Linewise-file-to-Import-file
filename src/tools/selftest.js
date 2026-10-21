// Self-test hook injected into test pages by build.ps1 (not part of the shipped page)
window.__selftest=async function(){
  const p=new URLSearchParams(location.search); const f=p.get('f'); if(!f) return;
  const out=document.createElement('pre'); out.id='selftest'; document.body.appendChild(out);
  try{
    const buf=await new Promise((ok,bad)=>{const x=new XMLHttpRequest();x.open('GET',f);x.responseType='arraybuffer';x.onload=()=>ok(x.response);x.onerror=()=>bad(new Error('xhr'));x.send()});
    const md=p.get('mode'); if(md){ await loadMaster(md,new File([buf],decodeURIComponent(f.split('/').pop()))); if(p.get('type'))setEitType(p.get('type'),'test'); await runMaster(md); const r=MS[md].result; let b64='';
      if(r&&p.get('b64')){let s='';const u=r.bytes;for(let i=0;i<u.length;i+=32768)s+=String.fromCharCode.apply(null,u.subarray(i,i+32768));b64=btoa(s)}
      out.textContent='RESULT '+JSON.stringify(r?{type:MS[md].type,rows:r.recs.length,stat:r.stat,issues:r.issues,msg:document.getElementById(md+'-msg').textContent}:{err:document.getElementById(md+'-msg').textContent,box:document.getElementById(md+'-err').innerText})+'\nB64 '+b64; document.title='DONE'; return }
    await loadFile(new File([buf],decodeURIComponent(f.split('/').pop())));
    if(p.get('sheet')){S.cur=S.sheets.findIndex(s=>s.name===p.get('sheet'));const sh=S.sheets[S.cur];if(!sh.sizeMode)guessModes(sh);renderMap()}
    const sh=S.sheets[S.cur];
    if(p.get('sys'))$('#inSys').value=p.get('sys');
    if(p.get('pc'))$('#const-ProjectCode').value=p.get('pc');
    const miss=missReq(sh);
    if(miss.length){out.textContent='RESULT '+JSON.stringify({unmappable:miss,err:$('#errBox').innerText.slice(0,400)});document.title='DONE';return}
    const t0=performance.now(); await run(); const r=S.result;
    const o={sheet:sh.name,size:sh.sizeMode,pipe:sh.pipeMode,rows:r.recs.length,st:r.st,skipped:r.skipped,pink:r.pinkRows,combos:r.combos.length,bytes:r.bytes.length,sec:((performance.now()-t0)/1000).toFixed(1),msg:$('#runMsg').textContent};
    let b64='';if(p.get('b64')){let s='';const u=r.bytes;for(let i=0;i<u.length;i+=32768)s+=String.fromCharCode.apply(null,u.subarray(i,i+32768));b64=btoa(s)}
    out.textContent='RESULT '+JSON.stringify(o)+'\nB64 '+b64;
  }catch(e){out.textContent='RESULT '+JSON.stringify({error:String(e&&e.stack||e)})}
  document.title='DONE';
};
