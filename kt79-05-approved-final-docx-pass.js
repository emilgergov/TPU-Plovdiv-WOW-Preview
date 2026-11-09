/* KT79-05 consolidated last-mile DOCX corrections. Preview only.
   Preserve pinned release V27.28. Operates after all original DOCX/ZIP guards.
   No new legal conclusion is synthesized; approved wording and typography only. */
(function(root){'use strict';
const NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const enc=new TextEncoder(), dec=new TextDecoder();
const FILES=['01_Proekt_na_stanovishte_Klasicheski.docx','02_Proekt_na_stanovishte_Eksperten.docx'];
const APPROVED='Инвестиционното предложение не е свързано с промяна предназначението и/или начина на трайно ползване на имота.';
const OLD='За евентуална промяна на предназначението и/или начина на трайно ползване на имота няма потвърдени данни в представената документация.';
const safe=(x)=>String(x||'');
function txt(p){return [...p.getElementsByTagNameNS(NS,'t')].map(x=>x.textContent||'').join('');}
function paragraphs(doc){const body=doc.getElementsByTagNameNS(NS,'body')[0];return [...(body?.childNodes||[])].filter(n=>n.nodeType===1&&n.localName==='p'&&n.namespaceURI===NS);}
function element(doc,name,value){const el=doc.createElementNS(NS,'w:'+name);if(value!==undefined)el.textContent=value;return el;}
function fixParagraph(doc,p,segments){
 const originals=[...p.getElementsByTagNameNS(NS,'r')];
 if(!originals.length)throw Error('KT79: cannot reconstruct a Word paragraph without runs');
 // Preserve Word paragraph properties, numbering, spacing and indentation.
 const base=originals[0], prototypePr=[...base.childNodes].find(n=>n.nodeType===1&&n.localName==='rPr');
 for(const node of [...p.childNodes])if(!(node.nodeType===1&&node.localName==='pPr'))p.removeChild(node);
 for(const part of segments){
   const run=element(doc,'r');
   let pr=prototypePr?prototypePr.cloneNode(true):element(doc,'rPr');
   for(const name of ['b','i']){
     const el=[...pr.childNodes].find(n=>n.nodeType===1&&n.localName===name);
     if(name==='b'&&part.bold){if(!el)pr.appendChild(element(doc,'b'));}
     if(name==='i'&&part.italic){if(!el)pr.appendChild(element(doc,'i'));}
     if(name==='b'&&!part.bold&&el)pr.removeChild(el);
     if(name==='i'&&!part.italic&&el)pr.removeChild(el);
   }
   run.appendChild(pr);const t=element(doc,'t',part.text);t.setAttribute('xml:space','preserve');run.appendChild(t);p.appendChild(run);
 }
}
function replaceAcrossTextRuns(p,pattern,replacement){
 const runs=[...p.getElementsByTagNameNS(NS,'r')];
 const chunks=runs.map(r=>({r,nodes:[...r.getElementsByTagNameNS(NS,'t')]})).filter(x=>x.nodes.length===1);
 const all=chunks.map(x=>x.nodes[0].textContent||'').join('');
 if(!pattern.test(all))return false;
 const result=all.replace(pattern,replacement);
 // Only typography, punctuation or an explicitly approved phrase: keep
 // existing individual run styles where a one-run substitution is possible.
 if(chunks.length===1){chunks[0].nodes[0].textContent=result;return true;}
 // For multi-run replacements fall back to a single run, without altering pPr.
 const doc=p.ownerDocument;fixParagraph(doc,p,[{text:result}]);return true;
}
function boldRanges(p,patterns){
 let matches=[];const full=txt(p);
 for(const re of patterns){re.lastIndex=0;for(const m of full.matchAll(re))if(m[0])matches.push([m.index,m.index+m[0].length]);}
 if(!matches.length)return 0;
 const runs=[...p.getElementsByTagNameNS(NS,'r')];
 let pos=0,edited=0;
 for(const run of runs){
  const nodes=[...run.getElementsByTagNameNS(NS,'t')];if(nodes.length!==1)continue;
  const t=nodes[0],value=t.textContent||'',start=pos;pos+=value.length;
  const edges=new Set([0,value.length]);
  for(const [a,b] of matches){if(a<pos&&b>start){edges.add(Math.max(0,a-start));edges.add(Math.min(value.length,b-start));}}
  const pieces=[...edges].sort((a,b)=>a-b);
  if(pieces.length===2&&!matches.some(([a,b])=>a<pos&&b>start))continue;
  const parent=run.parentNode;
  for(let i=0;i<pieces.length-1;i++){
   const x=pieces[i],y=pieces[i+1];if(y<=x)continue;
   const clone=run.cloneNode(true),tt=[...clone.getElementsByTagNameNS(NS,'t')][0];
   tt.textContent=value.slice(x,y);tt.setAttribute('xml:space','preserve');
   if(matches.some(([a,b])=>a<start+y&&b>start+x)){
    let rp=[...clone.childNodes].find(n=>n.nodeType===1&&n.localName==='rPr');
    if(!rp){rp=element(p.ownerDocument,'rPr');clone.insertBefore(rp,clone.firstChild);}
    for(const b of [...rp.childNodes].filter(n=>n.nodeType===1&&n.localName==='b'))rp.removeChild(b);
    rp.appendChild(element(p.ownerDocument,'b'));edited++;
   }
   parent.insertBefore(clone,run);
  }
  parent.removeChild(run);
 }
 return edited;
}
const BOLD_PATTERNS=[
 /BG0000194\s*„Река Чая“/gu,/BG0001031\s*„Родопи\s*[–-]\s*Средни“/gu,
 /BG0000578\s*„Река Марица“/gu,
 /(?:природна забележителност\s*)?„Свети Дух“/gu
];
function processFile(bytes,{classic,zt}){
 const zip=root.fflate.unzipSync(bytes),orig=zip['word/document.xml'];if(!orig)throw Error('KT79: DOCX XML missing');
 const doc=new DOMParser().parseFromString(dec.decode(orig),'application/xml');
 if(doc.getElementsByTagName('parsererror').length)throw Error('KT79: invalid Word XML');
 const ps=paragraphs(doc), report={paragraphs:ps.length,edits:[],boldSegments:0};
 const hasThird=ps.some(p=>txt(p).includes('46989.205.218'));
 const hasA=ps.some(p=>txt(p).includes('56784.13.102'));
 const hasB=ps.some(p=>txt(p).includes('36676.129.457'))||hasThird;
 if(classic&&hasA){
  const p=ps[10];
  if(p&&/РОГОШКО ШОСЕ.СЕВЕР/iu.test(txt(p))){
   fixParagraph(doc,p,[{text:'Относно: Уведомление за инвестиционно предложение (ИП): '},{text:'„Складове за промишлени стоки и офиси“',bold:true},{text:' в поземлени имоти с идентификатори 56784.13.102, 56784.13.105 и 56784.13.106 по КККР на гр. Пловдив, местност „Рогошко шосе - север“, община Пловдив, област Пловдив.',italic:true}]);
   report.edits.push('Classic A §11 locality');
  }
 }
 if(hasA&&!classic){for(const p of ps)if(/РОГОШКО ШОСЕ.СЕВЕР/iu.test(txt(p))){
  if(replaceAcrossTextRuns(p,/РОГОШКО ШОСЕ\s*[-–]\s*СЕВЕР/giu,'Рогошко шосе - север'))report.edits.push('Expert A locality');
 }}
 if(classic&&hasThird){
  const p=ps[10];
  if(!p||!txt(p).includes('46989.205.218'))throw Error('KT79: third-case subject slot absent');
  fixParagraph(doc,p,[{text:'Относно: Уведомление за инвестиционно предложение (ИП): '},{text:'„Разполагане на малък посетителски информационен пункт“',bold:true},{text:' в поземлен имот с идентификатор 46989.205.218 по КККР на с. Манастир, местност „Общия балкан“, община Лъки, област Пловдив.',italic:true}]);
  report.edits.push('Classic B+ZT §11 unique subject');
  // Preserve the existing non-final procedural statement; remove only
  // the mechanically repeated subject + parcel/location substring.
  const conclusion=ps.find(x=>txt(x).startsWith('В тази връзка и на основание чл. 2, ал. 2'));
  if(conclusion&&txt(conclusion).includes('изисква проверка')){
   const current=txt(conclusion),suffix=current.slice(current.indexOf('изисква проверка'));
   if(!suffix.includes('46989.205.218')){
     fixParagraph(doc,conclusion,[
       {text:'В тази връзка и на основание чл. 2, ал. 2 от Наредбата за ОС Ви уведомявам, че преценката на компетентния орган за така заявеното инвестиционно предложение: '},
       {text:'„Разполагане на малък посетителски информационен пункт“',bold:true},
       {text:' в поземлен имот с идентификатор 46989.205.218 по КККР на с. Манастир, община Лъки, област Пловдив '},
       {text:suffix}
     ]);
     report.edits.push('Classic B+ZT procedural subject de-duplicated; no conclusion invented');
   }
  }

  const desc=ps.find(x=>txt(x).startsWith('С инвестиционното предложение се предвижда'));
  if(desc&&/посетителски информационен пункт/iu.test(txt(desc))){
   const keepArea=/до\s+20\s*кв/iu.test(txt(desc));
   fixParagraph(doc,desc,[{text:'С инвестиционното предложение се предвижда разполагане на малък посетителски информационен пункт в поземлен имот с идентификатор 46989.205.218 по КККР на с. Манастир, местност „Общия балкан“, община Лъки, област Пловдив.'+(keepArea?' В представеното уведомление е посочена приблизителна застроена площ до 20 кв. м.':'')}]);
   report.edits.push('Classic B+ZT description no duplicated title');
  }
 }
 if(classic&&hasB){
  const p=ps.find(x=>txt(x)===OLD);
  if(p){fixParagraph(doc,p,[{text:APPROVED,bold:true}]);report.edits.push('Classic B §23 approved wording');}
  else if(!ps.some(x=>txt(x)===APPROVED))throw Error('KT79: approved Classic B §23 absent');
 }
 if(hasThird){
  for(const p of ps){
    if(!/Свети Дух/u.test(txt(p)))continue;
    const before=txt(p);
    if(before.includes('2018 г..')){
       replaceAcrossTextRuns(p,/2018 г\.\./gu,'2018 г.');
       report.edits.push('ZT double full stop');
    }
  }
  if(!classic){
   for(const p of ps){
    if(!/46989\.205\.218/.test(txt(p)))continue;
    if(/Х{6,}/.test(txt(p))&&/Манастир/u.test(txt(p))){
      // Never infer missing administrative attributes here: only replace
      // if confirmed in Classic B; flag for separate factual QA instead.
      report.edits.push('Expert admin placeholders require evidence QA');
    }
   }
  }
 }
 for(const p of ps){
  const t=txt(p);
  const patterns=[...BOLD_PATTERNS];
  if(/границите на|пространствената проверка|защитена зона/u.test(t)&&/попада/u.test(t))patterns.push(/\bпопада\b/gu);
  if(/пространствената проверка|границите на/u.test(t)&&/не попадат/u.test(t))patterns.push(/не попадат/gu);
  report.boldSegments+=boldRanges(p,patterns);
 }
 zip['word/document.xml']=enc.encode(new XMLSerializer().serializeToString(doc));
 return {data:root.fflate.zipSync(zip,{level:6}),report};
}
async function fix(blob){
 if(!root.fflate?.unzipSync||!root.fflate?.zipSync)throw Error('KT79: ZIP engine unavailable');
 const files=root.fflate.unzipSync(new Uint8Array(await blob.arrayBuffer()));
 const reports={};
 for(const [i,name] of FILES.entries()){
  if(!files[name])throw Error('KT79: missing Word '+name);
  const r=processFile(files[name],{classic:i===0});
  files[name]=r.data;reports[name]=r.report;
 }
 // All PDFs and non-DOCX files remain byte-identical after roundtrip.
 const output=new Blob([root.fflate.zipSync(files,{level:6})],{type:'application/zip'});
 return {blob:output,reports};
}
root.KT79FinalApproved=Object.freeze({fix});
})(window);
