"use strict";
const KEY_ID = "InteractiveBRD.DocumentId.v1";
const KEY_AUTO = "Office.AutoShowTaskpaneWithDocument";
const $ = id => document.getElementById(id);
let catalog = [], officeReady = false;
function message(s, isError=false){ $("status").textContent=s; $("status").classList.toggle("error",isError); }
function selectedBRD(){return catalog.find(d=>d.id===$("brd").value);}
function fillIllustrations(){
  const doc=selectedBRD(); const el=$("illustration"); el.replaceChildren(new Option("Select illustration", ""));
  (doc?.illustrations || []).forEach(i=>el.add(new Option(i.title,i.id)));
  if(doc?.illustrations?.length){el.value=doc.illustrations[0].id;}
  showIllustration();
}
function showIllustration(){
  const doc=selectedBRD(), item=doc?.illustrations?.find(x=>x.id===$("illustration").value);
  const f=$("illustration-frame"); $("diagram-title").textContent=item?.title||"No illustration selected";
  $("version").textContent=doc?`${doc.id} • v${doc.version}`:"";
  if(!item){f.hidden=true;f.removeAttribute("src");$("empty").hidden=false;return;}
  // Catalog paths are not arbitrary URLs supplied by the user.
  const u=new URL(item.path,location.origin);
  if(u.origin!==location.origin || !u.pathname.startsWith("/illustrations/") || !u.pathname.endsWith(".html")){
    message("Rejected unsafe illustration URL in catalog.",true);f.hidden=true;return;
  }
  f.src=u.pathname;f.hidden=false;$("empty").hidden=true;
}
function saveSettings(){
  return new Promise((resolve,reject)=>Office.context.document.settings.saveAsync(result=>{
    if(result.status===Office.AsyncResultStatus.Succeeded) resolve();
    else reject(new Error(result.error?.message || "Could not save document settings"));
  }));
}
async function bindDocument(){
  if(!officeReady){ message("Open this viewer in Microsoft Word to associate the document.",true);return; }
  const doc=selectedBRD();if(!doc){message("Select a BRD first.",true);return;}
  const s=Office.context.document.settings;
  s.set(KEY_ID,doc.id);
  s.set(KEY_AUTO,$("autoopen").checked);
  try{await saveSettings();message(`Associated ${doc.id}. Save the Word document (Ctrl+S) to retain settings.`);}
  catch(e){message(`Association failed: ${e.message}`,true);}
}
async function load(){
  try{
    const response=await fetch("/catalog.json",{cache:"no-store"});
    if(!response.ok)throw new Error(`Catalog HTTP ${response.status}`);
    const json=await response.json();
    if(json.schemaVersion!==1 || !Array.isArray(json.documents))throw new Error("Invalid catalog schema");
    catalog=json.documents;
    const menu=$("brd");menu.replaceChildren(new Option("Select a BRD", ""));
    catalog.forEach(d=>menu.add(new Option(`${d.id} — ${d.title}`,d.id)));
    let id="";
    if(officeReady){
      id=Office.context.document.settings.get(KEY_ID)||"";
      $("autoopen").checked=Office.context.document.settings.get(KEY_AUTO)===true;
    }
    if(id&&catalog.some(d=>d.id===id)){menu.value=id;message(`Document mapping found: ${id}.`);}
    else if(id){message(`The document refers to ${id}, which is not in the current catalog.`,true);}
    else{message(officeReady?"Select the BRD, then associate it with this Word document.":"Browser preview mode. In Word, association and auto-open become available.");}
    fillIllustrations();
  }catch(e){message(`Could not load BRD catalog: ${e.message}`,true);}
}
window.addEventListener("DOMContentLoaded",()=>{
  $("brd").addEventListener("change",fillIllustrations);
  $("illustration").addEventListener("change",showIllustration);
  $("bind").addEventListener("click",bindDocument);
  $("reload").addEventListener("click",()=>{const f=$("illustration-frame");if(f.src)f.contentWindow?.location.reload();else load();});
  if(typeof Office!=="undefined" && Office.onReady){
    Office.onReady(info=>{officeReady=info.host===Office.HostType.Word;
      $("office-state").textContent=officeReady?"Word connected":"Browser preview";
      $("bind").disabled=!officeReady;$("autoopen").disabled=!officeReady;
      load();
    }).catch(()=>{ $("office-state").textContent="Browser preview";$("bind").disabled=true;$("autoopen").disabled=true;load(); });
  }else{$("office-state").textContent="Browser preview";$("bind").disabled=true;$("autoopen").disabled=true;load();}
});
