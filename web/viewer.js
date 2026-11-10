"use strict";
const byId=id=>document.getElementById(id);
let documents=[];
function status(message,error=false){byId("status").textContent=message;byId("status").classList.toggle("error",error);}
function currentDocument(){return documents.find(d=>d.id===byId("brd").value);}
function safePath(value){
  if(typeof value!=="string"||!/^\/illustrations\/[a-zA-Z0-9._/-]+\.html$/.test(value)||value.includes("..")||value.includes("//"))return null;
  const u=new URL(value,location.origin);
  return u.origin===location.origin?u.pathname:null;
}
function chooseIllustration(){
  const d=currentDocument(),selector=byId("illustration"),previous=selector.value;
  selector.replaceChildren(new Option("Select an illustration",""));
  for(const a of d?.illustrations||[])selector.add(new Option(a.title,a.id));
  const requested=new URLSearchParams(location.search).get("illustration");
  selector.value=(d?.illustrations||[]).some(x=>x.id===requested)?requested:(d?.illustrations||[]).some(x=>x.id===previous)?previous:(d?.illustrations?.[0]?.id||"");
  show();
}
function show(){
  const d=currentDocument(),a=d?.illustrations?.find(x=>x.id===byId("illustration").value);
  const path=a?safePath(a.path):null;
  const frame=byId("frame"),open=byId("open"),copy=byId("copy");
  byId("project").textContent=d?.id||"BRD illustration";
  byId("title").textContent=a?.title||"Choose an illustration";
  byId("description").textContent=a?.description||"";
  byId("version").textContent=d?"Version "+d.version:"";
  if(!path){frame.hidden=true;frame.removeAttribute("src");byId("empty").hidden=false;open.href="#";open.setAttribute("aria-disabled","true");copy.disabled=true;if(a)status("Invalid illustration path in catalog.",true);return;}
  frame.hidden=false;byId("empty").hidden=true;
  if(frame.getAttribute("src")!==path)frame.src=path;
  open.href=path;open.setAttribute("aria-disabled","false");copy.disabled=false;
  const params=new URLSearchParams({brd:d.id,illustration:a.id});
  history.replaceState(null,"",location.pathname+"?"+params.toString());
  status("Ready. Copy the shareable link for your Word document, or open the HTML full-page.");
}
async function copyLink(){
  const url=location.href;
  try{await navigator.clipboard.writeText(url);status("Shareable link copied. Paste it into your .docx.");}
  catch{status("Copy unavailable. Select the browser address bar and copy the URL.",true);}
}
function fillDocuments(){
  const query=byId("search").value.trim().toLowerCase(),prev=byId("brd").value;
  const filtered=documents.filter(d=>(d.id+" "+d.title+" "+d.version).toLowerCase().includes(query));
  const el=byId("brd");
  el.replaceChildren(new Option("Select a BRD",""));
  for(const d of filtered)el.add(new Option(d.id+" — "+d.title,d.id));
  const requested=new URLSearchParams(location.search).get("brd");
  el.value=filtered.some(d=>d.id===requested)?requested:filtered.some(d=>d.id===prev)?prev:(filtered.length===1?filtered[0].id:"");
  chooseIllustration();
  if(!filtered.length)status("No matching BRDs.");
}
async function init(){
  byId("search").addEventListener("input",fillDocuments);
  byId("brd").addEventListener("change",()=>{history.replaceState(null,"",location.pathname);chooseIllustration();});
  byId("illustration").addEventListener("change",show);
  byId("copy").addEventListener("click",copyLink);
  try{
    const response=await fetch("/catalog.json",{cache:"no-store"});
    if(!response.ok)throw Error("Catalog HTTP "+response.status);
    const json=await response.json();
    if(json.schemaVersion!==1||!Array.isArray(json.documents))throw Error("Unexpected catalog format");
    documents=json.documents;
    fillDocuments();
    if(!documents.length)status("No BRDs registered yet.");
  }catch(e){status("Cannot load catalog: "+e.message,true);}
}
document.addEventListener("DOMContentLoaded",init);
