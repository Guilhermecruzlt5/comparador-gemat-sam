/* Leitura integrada de arquivos do Manual Interativo SEOM/SEFISC */
(()=>{
"use strict";
const style=document.createElement("style");
style.textContent=".seom-read-actions{display:flex;gap:9px;flex-wrap:wrap;margin-top:12px}.seom-doc-modal{padding:0;border:0;border-radius:15px;width:min(1160px,96vw);max-width:none;height:min(92vh,960px);max-height:96vh;background:white;overflow:hidden;box-shadow:0 25px 85px #001d32aa}.seom-doc-modal::backdrop{background:#001a2ecb}.seom-doc-shell{height:100%;display:flex;flex-direction:column}.seom-doc-header{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:13px 15px;background:#eff7f8;border-bottom:1px solid #cbdde2}.seom-doc-header strong{flex:1;min-width:200px;color:#0b374a;font-size:15px}.seom-doc-tools{display:flex;gap:6px;flex-wrap:wrap}.seom-doc-content{background:#eaf1f3;flex:1;min-height:0;display:flex;flex-direction:column}.seom-doc-content iframe{width:100%;height:100%;border:0;background:white}.seom-doc-info{padding:9px 15px;color:#466172;font-size:12px;border-top:1px solid #d8e5e8}.seom-doc-status{padding:30px;line-height:1.6;color:#18465a;text-align:center}@media(max-width:690px){.seom-doc-modal{width:100vw;height:100dvh;max-height:100dvh;border-radius:0}.seom-doc-header{padding:9px}.seom-doc-header strong{font-size:13px}}";
document.head.appendChild(style);
const dlg=document.createElement("dialog");dlg.className="seom-doc-modal";dlg.setAttribute("aria-label","Leitor de documentos SEOM");
dlg.innerHTML='<div class="seom-doc-shell"><div class="seom-doc-header"><strong id="seomDocTitle">Ler documento</strong><div class="seom-doc-tools"><button type="button" class="btn aqua" id="seomDocTab">↗ Nova guia</button><button type="button" class="btn primary" id="seomDocDownload">↓ Baixar</button><button type="button" class="btn" id="seomDocClose">✕ Fechar</button></div></div><div class="seom-doc-content" id="seomDocContent"></div><div class="seom-doc-info" id="seomDocInfo">O arquivo original está disponível para download.</div></div>';
document.body.appendChild(dlg);
const content=document.getElementById("seomDocContent"),heading=document.getElementById("seomDocTitle"),info=document.getElementById("seomDocInfo");
let active=null,activeUrl=null;
function esc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
function inject(){
document.querySelectorAll(".resource").forEach(card=>{const b=card.querySelector("[data-download]");if(!b||card.querySelector("[data-seom-read]"))return;const wrap=document.createElement("div");wrap.className="seom-read-actions";const view=document.createElement("button");view.className="btn aqua";view.type="button";view.dataset.seomRead=b.dataset.download;view.textContent="👁 Ler documento";b.before(wrap);wrap.append(view,b);});
document.querySelectorAll(".links-box [data-download]").forEach(b=>{if(b.parentNode.querySelector('[data-seom-read="'+b.dataset.download+'"]'))return;const view=document.createElement("button");view.className="btn aqua";view.type="button";view.dataset.seomRead=b.dataset.download;view.textContent="👁 Ler documento";b.before(view);});
}
const main=document.getElementById("main");if(main){inject();new MutationObserver(inject).observe(main,{childList:true,subtree:true});}
function decode(text){const bin=atob(text),out=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)out[i]=bin.charCodeAt(i);return out;}
function urlOf(data,type){return URL.createObjectURL(new Blob([data],{type}));}
function frame(url,isolated){content.replaceChildren();const f=document.createElement("iframe");f.title="Visualização de documento";if(isolated!==null)f.setAttribute("sandbox",isolated);f.src=url;content.append(f);}
function html(inner){return '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font:14px/1.65 Arial,Helvetica,sans-serif;color:#173b4d;max-width:970px;padding:24px;margin:auto}h1{font-size:23px;color:#083751;border-bottom:3px solid #4bc7c1;padding-bottom:14px}h2{font-size:18px;margin-top:23px}p{margin:9px 0}table{width:100%;border-collapse:collapse;margin:14px 0}td{padding:8px;border:1px solid #c8d7dc;vertical-align:top}section{padding:15px;border:1px solid #d5e5e9;margin:18px 0;border-radius:12px}.note{border-left:4px solid #dfb23e;background:#fff9e6;padding:12px;font-size:13px}</style></head><body>'+inner+'</body></html>';}
async function unzip(data,pick){
const dv=new DataView(data.buffer,data.byteOffset,data.byteLength),decoder=new TextDecoder("utf-8");let eocd=-1;
for(let i=data.length-22;i>=Math.max(0,data.length-65557);i--){if(dv.getUint32(i,true)===0x06054b50){eocd=i;break;}}
if(eocd<0)throw Error("Arquivo do Office não reconhecido.");
let pos=dv.getUint32(eocd+16,true),count=dv.getUint16(eocd+10,true),files={};
for(let i=0;i<count;i++){
if(dv.getUint32(pos,true)!==0x02014b50)throw Error("Estrutura do arquivo Office inválida.");
const method=dv.getUint16(pos+10,true),size=dv.getUint32(pos+20,true),nl=dv.getUint16(pos+28,true),extra=dv.getUint16(pos+30,true),comment=dv.getUint16(pos+32,true),local=dv.getUint32(pos+42,true);
const name=decoder.decode(data.subarray(pos+46,pos+46+nl));pos+=46+nl+extra+comment;if(!pick(name))continue;
const start=local+30+dv.getUint16(local+26,true)+dv.getUint16(local+28,true);
const packed=data.subarray(start,start+size);let raw;
if(method===0)raw=packed;
else if(method===8){if(typeof DecompressionStream==="undefined")throw Error("O navegador não permite abrir este arquivo em leitura. Faça o download.");const stream=new Blob([packed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));raw=new Uint8Array(await new Response(stream).arrayBuffer());}
else throw Error("Método de compactação não suportado.");
files[name]=decoder.decode(raw);
}
return files;
}
const W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
function wText(n){return Array.from(n.getElementsByTagNameNS(W,"t")).map(t=>t.textContent).join("").trim();}
async function readWord(bytes,name){
const files=await unzip(bytes,n=>n==="word/document.xml"),raw=files["word/document.xml"];if(!raw)throw Error("Conteúdo do Word não encontrado.");
const x=new DOMParser().parseFromString(raw,"application/xml"),body=x.getElementsByTagNameNS(W,"body")[0];if(!body)throw Error("Documento Word sem texto disponível.");
let chunks=[];
for(const item of body.children){
if(item.localName==="p"){const t=wText(item);if(!t)continue;const st=item.getElementsByTagNameNS(W,"pStyle")[0],v=st?.getAttributeNS(W,"val")||"";const tag=/heading1|titulo1/i.test(v)?"h1":/heading|titulo/i.test(v)?"h2":"p";chunks.push("<"+tag+">"+esc(t)+"</"+tag+">");}
else if(item.localName==="tbl"){let rows=Array.from(item.children).filter(n=>n.localName==="tr");chunks.push("<table>"+rows.map(r=>"<tr>"+Array.from(r.children).filter(c=>c.localName==="tc").map(c=>"<td>"+esc(wText(c))+"</td>").join("")+"</tr>").join("")+"</table>");}
}
return html("<h1>"+esc(name)+"</h1><div class=note>Prévia para leitura do texto e das tabelas. Elementos gráficos e detalhes de formatação podem não ser exibidos; o original permanece disponível em “Baixar”.</div>"+(chunks.join("")||"<p>Sem texto extraível.</p>"));
}
async function readPowerPoint(bytes,name){
const files=await unzip(bytes,n=>/^ppt\/slides\/slide\d+\.xml$/.test(n)),names=Object.keys(files).sort((a,b)=>Number(a.match(/slide(\d+)/)[1])-Number(b.match(/slide(\d+)/)[1]));
let result="<h1>"+esc(name)+"</h1><div class=note>Prévia textual dos "+names.length+" slides. Para ver imagens, gráficos e diagramação, baixe o PowerPoint original.</div>";
for(let i=0;i<names.length;i++){const xml=new DOMParser().parseFromString(files[names[i]],"application/xml"),ns="http://schemas.openxmlformats.org/drawingml/2006/main";const paragraphs=Array.from(xml.getElementsByTagNameNS(ns,"p")).map(p=>Array.from(p.getElementsByTagNameNS(ns,"t")).map(t=>t.textContent).join("")).filter(Boolean);result+="<section><h2>Slide "+(i+1)+"</h2>"+(paragraphs.map(t=>"<p>"+esc(t)+"</p>").join("")||"<p>Sem texto extraível. Consulte o arquivo original.</p>")+"</section>";}
return html(result);
}
async function show(key,source){
if(typeof DOWNLOADS==="undefined"||!DOWNLOADS[key])return;
const file=DOWNLOADS[key];active=key;
heading.textContent=source.closest(".resource")?.querySelector("strong")?.textContent?.replace(/^📎\s*/,"")||file.name;
content.innerHTML='<div class="seom-doc-status">Preparando a leitura…</div>';
info.textContent="O arquivo original continua disponível para baixar.";
if(!dlg.open)dlg.showModal();
if(activeUrl){URL.revokeObjectURL(activeUrl);activeUrl=null;}
try{
const bytes=decode(file.b64),mime=file.mime||"",filename=file.name.toLowerCase();
if(active!==key)return;
if(mime==="application/pdf"){activeUrl=urlOf(bytes,"application/pdf");frame(activeUrl,null);info.textContent="PDF original • Use os controles de páginas do navegador. Se necessário, abra em nova guia.";}
else if(mime==="text/html"||filename.endsWith(".html")){activeUrl=urlOf(bytes,"text/html");frame(activeUrl,"allow-scripts allow-forms allow-downloads");info.textContent="Aplicativo HTML aberto em uma área isolada.";}
else if(filename.endsWith(".docx")){const preview=await readWord(bytes,file.name);if(active!==key)return;activeUrl=urlOf(new TextEncoder().encode(preview),"text/html");frame(activeUrl,"");}
else if(filename.endsWith(".pptx")){const preview=await readPowerPoint(bytes,file.name);if(active!==key)return;activeUrl=urlOf(new TextEncoder().encode(preview),"text/html");frame(activeUrl,"");}
else content.innerHTML='<div class="seom-doc-status">Não há prévia para este formato. Utilize “Baixar”.</div>';
}catch(err){console.error(err);content.textContent="Não foi possível mostrar a leitura: "+(err.message||"erro de arquivo")+". Você pode baixar o original.";}
}
document.addEventListener("click",e=>{const button=e.target.closest("[data-seom-read]");if(!button)return;e.preventDefault();e.stopPropagation();show(button.dataset.seomRead,button);});
document.getElementById("seomDocClose").onclick=()=>dlg.close();
document.getElementById("seomDocTab").onclick=()=>{if(activeUrl)window.open(activeUrl,"_blank","noopener,noreferrer");};
document.getElementById("seomDocDownload").onclick=()=>{if(!active)return;const b=document.createElement("button");b.dataset.download=active;b.hidden=true;document.body.append(b);b.click();b.remove();};
dlg.addEventListener("close",()=>{active=null;content.replaceChildren();if(activeUrl){URL.revokeObjectURL(activeUrl);activeUrl=null;}});
})();