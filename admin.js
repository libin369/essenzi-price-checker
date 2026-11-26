'use strict';
(()=>{
const el=id=>document.getElementById(id), dialog=el('adminDialog');
let token='', pending=null, sha='', baseline=null, locked=false;
const owner=location.hostname.endsWith('.github.io')?location.hostname.split('.')[0]:'', repo=location.pathname.split('/').filter(Boolean)[0]||'';
const endpoint=`https://api.github.com/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/prices.json`;
const msg=s=>el('adminMessage').textContent=s;
const text=v=>v==null?'':String(v).trim();
function convert(mapping,rows,baseline){
const norm=v=>text(v).toUpperCase().replace(/[^A-Z0-9]/g,'');
const headers=rows[0].map(norm), col=(...names)=>headers.findIndex(h=>names.includes(h));
const c={code:col('ITEMCODE','ITEMNO'),barcode:col('BARCODE'),name:col('DESCRIPTION','DESCRIPITON'),brand:col('BRAND'),was:col('WASPRICE','WAS','SYSRSP'),now:col('NOWPRICE','NOW','ROUND')};
if(['code','barcode','name','was','now'].some(k=>c[k]<0))throw Error('Required headers: ItemCode, Barcode, Description, WAS Price, NOW Price. Brand optional.');
const products=new Map(),aliases=new Map(),old=new Map(baseline.products.map(p=>[p.code,p])), invalid=[];
for(const [i,r] of rows.slice(1).entries()){
 if(!r.some(v=>text(v)))continue;
 const k=text(r[c.code]).toUpperCase(),b=text(r[c.barcode]),name=text(r[c.name]),was=Number(r[c.was]),now=Number(r[c.now]);
 if(!k||!b||!name||!text(r[c.was])||!text(r[c.now])||!Number.isFinite(was)||!Number.isFinite(now)||Math.min(was,now)<0){invalid.push(i+2);continue;}
 if(!aliases.has(k))aliases.set(k,new Set([...(Object.hasOwn(mapping,k)?mapping[k]:[]),...(old.get(k)?.barcodes||[])]));aliases.get(k).add(b);
 if(!products.has(k)||now>products.get(k).now)products.set(k,{code:k,name,brand:text(r[c.brand])||old.get(k)?.brand||'',was,now});
}
if(invalid.length)throw Error('Invalid rows: '+invalid.slice(0,30).join(', ')+'. Nothing published.');if(!products.size)throw Error('No products found.');
return {updated:new Date().toISOString(),products:[...products.values()].map(p=>({...p,barcodes:[...aliases.get(p.code)].sort()}))};}
async function api(method='GET',body){const r=await fetch(endpoint,{method,headers:{Authorization:'Bearer '+token,Accept:'application/vnd.github+json','Content-Type':'application/json','X-GitHub-Api-Version':'2022-11-28'},...(body?{body:JSON.stringify(body)}:{})});if(!r.ok)throw Error(`GitHub request failed (${r.status}). Check token, repository permissions and branch.`);return r.json();}
function reset(){token='';pending=null;sha='';baseline=null;el('adminToken').value='';el('excelFile').value='';el('uploadPanel').hidden=true;el('loginPanel').hidden=false;el('preview').hidden=true;el('publish').disabled=true;msg('');}
function readExcel(buffer){return new Promise((resolve,reject)=>{const worker=new Worker('excel-worker.js');const timer=setTimeout(()=>{worker.terminate();reject(Error('Excel reader timed out. Check internet.'));},45000);const finish=()=>{clearTimeout(timer);worker.terminate();};worker.onmessage=e=>{finish();e.data.error?reject(Error(e.data.error)):resolve(e.data);};worker.onerror=()=>{finish();reject(Error('Excel reader failed. Check internet.'));};worker.postMessage(buffer,[buffer]);});}
el('adminOpen').onclick=()=>{reset();dialog.showModal();};el('adminClose').onclick=()=>{if(!locked)dialog.close();};dialog.addEventListener('cancel',e=>{if(locked)e.preventDefault();});dialog.addEventListener('close',reset);
async function run(button,fn){if(locked)return;locked=true;button.disabled=true;el('adminClose').disabled=true;el('excelFile').disabled=true;try{await fn();}catch(e){msg(e.message);}finally{locked=false;button.disabled=false;el('adminClose').disabled=false;el('excelFile').disabled=false;if(button===el('publish'))button.disabled=!pending;}}
el('unlock').onclick=()=>run(el('unlock'),async()=>{if(!owner||!repo)throw Error('Open the deployed GitHub Pages project URL first.');token=el('adminToken').value.trim();el('adminToken').value='';if(!/^(github_pat_|ghp_)[A-Za-z0-9_]+$/.test(token)){token='';throw Error('Enter a valid GitHub access token.');}try{const a=await api();sha=a.sha;baseline=JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(a.content.replace(/\s/g,'')),c=>c.charCodeAt(0))));}catch(e){token='';throw e;}el('loginPanel').hidden=true;el('uploadPanel').hidden=false;msg('Credential accepted for reading. GitHub verifies write permission when publishing.');});
el('excelFile').onchange=()=>{pending=null;el('publish').disabled=true;el('preview').hidden=true;};
el('validate').onclick=()=>run(el('validate'),async()=>{pending=null;el('publish').disabled=true;el('preview').hidden=true;const f=el('excelFile').files[0];if(!f||!f.name.toLowerCase().endsWith('.xlsx'))throw Error('Select an .xlsx workbook.');const parsed=await readExcel(await f.arrayBuffer());const mappingResponse=await fetch('barcode_mapping.json',{cache:'no-store'});if(!mappingResponse.ok)throw Error('Saved barcode mapping unavailable.');const mapping=await mappingResponse.json();pending=convert(mapping,parsed.rows,baseline);const old=new Map(baseline.products.map(p=>[p.code,p]));const now=new Map(pending.products.map(p=>[p.code,p]));const changed=pending.products.filter(p=>old.has(p.code)&&(old.get(p.code).now!==p.now||old.get(p.code).was!==p.was));const added=pending.products.filter(p=>!old.has(p.code)).length,removed=baseline.products.filter(p=>!now.has(p.code)).length;el('preview').textContent=`Products: ${now.size}\nNew: ${added} · Removed: ${removed}\nPrice changes: ${changed.length}\n`+changed.slice(0,12).map(p=>`${p.code}: ${old.get(p.code).now} → ${p.now}`).join('\n');el('preview').hidden=false;el('publish').disabled=false;msg('Validated. Review changes before publishing.');});
el('publish').onclick=()=>run(el('publish'),async()=>{if(!pending)return;if(!confirm('Replace the entire live price list with '+pending.products.length+' products?'))return;const current=await api();if(current.sha!==sha){pending=null;throw Error('Live prices changed. Close, unlock and validate again.');}pending.updated=new Date().toISOString();const bytes=new TextEncoder().encode(JSON.stringify(pending));let binary='';for(let i=0;i<bytes.length;i+=8192)binary+=String.fromCharCode(...bytes.subarray(i,i+8192));await api('PUT',{message:'Update exhibition prices',content:btoa(binary),sha});pending=null;token='';el('uploadPanel').hidden=true;el('loginPanel').hidden=false;msg('Published to GitHub. Wait for Pages deployment, then use Sync prices. Administrator locked.');});
})();
