/* KT79-06..10: template-preserving, property-driven final Word adapter.
 * No parcel IDs, protection-zone IDs or project titles are used as branches.
 * Only input-confirmed GIS/admin values can populate slots.
 * V27.28 pinned runtime and PDF1/PDF2 are read-only.
 */
(function(root){'use strict';
const NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main',enc=new TextEncoder(),dec=new TextDecoder();
const CLASSIC='01_Proekt_na_stanovishte_Klasicheski.docx',EXPERT='02_Proekt_na_stanovishte_Eksperten.docx';
const APPROVED23_TEXT='Инвестиционното предложение не е свързано с промяна предназначението и/или начина на трайно ползване на имота.';
const APPROVED_B_END='е, че не е необходимо провеждане на процедура по реда на Глава втора от Наредбата за ОС.';
function children(n,name){return [...(n?.childNodes||[])].filter(x=>x.nodeType===1&&x.namespaceURI===NS&&x.localName===name);}
function nodes(n,name){return Array.from(n.getElementsByTagNameNS(NS,name));}
function paragraphs(doc){return children(nodes(doc,'body')[0],'p');}
function text(p){return nodes(p,'t').map(x=>x.textContent||'').join('');}
function w(doc,name,value){const e=doc.createElementNS(NS,'w:'+name);if(value!==undefined)e.textContent=value;return e;}
function styleRun(run,prop,on){
 let rp=children(run,'rPr')[0];if(!rp){rp=w(run.ownerDocument,'rPr');run.insertBefore(rp,run.firstChild);}
 for(const old of children(rp,prop))rp.removeChild(old);
 const attr=w(run.ownerDocument,prop);
 if(!on)attr.setAttributeNS(NS,'w:val','0');
 rp.appendChild(attr);
}
function alter(p,a,b,replacement,style=null){
 const full=text(p);if(a<0||b>full.length||a>b)throw Error('KT79: invalid Word span');
 const original=nodes(p,'r').filter(r=>children(r,'t').length===1),parts=[];
 let at=0;
 for(const run of original){const t=children(run,'t')[0],value=t.textContent||'',lo=at,hi=lo+value.length;at=hi;
  if(!value.length||hi<=a||lo>=b){parts.push({run:run.cloneNode(true),lo,hi});continue;}
  const first=Math.max(lo,a),last=Math.min(hi,b);
  const left=value.slice(0,first-lo),right=value.slice(last-lo);
  let before=null,after=null;
  if(left){before=run.cloneNode(true);children(before,'t')[0].textContent=left;parts.push({run:before,lo,hi:first});}
  if(!parts._used && replacement!==null && first===a){
    const middle=run.cloneNode(true);children(middle,'t')[0].textContent=replacement;
    children(middle,'t')[0].setAttribute('xml:space','preserve');
    if(style){if(style.bold!==undefined)styleRun(middle,'b',style.bold);if(style.italic!==undefined)styleRun(middle,'i',style.italic);}
    parts.push({run:middle,lo:a,hi:a+replacement.length});parts._used=true;
  }
  if(right){after=run.cloneNode(true);children(after,'t')[0].textContent=right;parts.push({run:after,lo:last,hi});}
 }
 if(replacement!==null&&!parts._used)throw Error('KT79: Word span crosses unsupported run structure');
 // Insert the retained/cloned runs in order at the first affected run position.
 const first=original.find(r=>{const n=children(r,'t')[0];return n&&n.textContent!==undefined;});
 if(!first)throw Error('KT79: unsupported Word run structure');
 const parent=first.parentNode;
 for(const part of parts)parent.insertBefore(part.run,first);
 for(const run of original)if(run.parentNode)run.parentNode.removeChild(run);
 if(replacement!==null && text(p)!==full.slice(0,a)+replacement+full.slice(b))throw Error('KT79: Word text replacement invariant failed');
}
function replaceOnce(p,re,repl){
 const full=text(p);re.lastIndex=0;const m=re.exec(full);if(!m)return false;
 alter(p,m.index,m.index+m[0].length,typeof repl==='function'?repl(m):repl);
 return true;
}
function setSpanStyle(p,a,b,style){
 const full=text(p);if(a<0||b>full.length||a>=b)throw Error('KT79: invalid bold/italic coordinates');
 let offset=0;
 for(const run of [...nodes(p,'r')]){
  const t=children(run,'t');if(t.length!==1)continue;
  const v=t[0].textContent||'',start=offset,end=start+v.length;offset=end;
  if(end<=a||start>=b)continue;
  const cuts=[0,Math.max(0,a-start),Math.min(v.length,b-start),v.length].filter((x,i,ar)=>x>=0&&x<=v.length);
  const unique=[...new Set(cuts)].sort((x,y)=>x-y),parent=run.parentNode;
  for(let i=0;i<unique.length-1;i++){
   const x=unique[i],y=unique[i+1];if(x===y)continue;
   const clone=run.cloneNode(true);const tt=children(clone,'t')[0];tt.textContent=v.slice(x,y);tt.setAttribute('xml:space','preserve');
   if(start+x>=a&&start+y<=b){
    if(style.bold!==undefined)styleRun(clone,'b',style.bold);
    if(style.italic!==undefined)styleRun(clone,'i',style.italic);
   }
   parent.insertBefore(clone,run);
  }
  parent.removeChild(run);
 }
 if(text(p)!==full)throw Error('KT79: Word style must not change text');
}
function setMatches(p,re,style){
 const t=text(p),ranges=[...t.matchAll(re)].map(m=>[m.index,m.index+m[0].length]);for(const [a,b] of ranges)setSpanStyle(p,a,b,style);
 return ranges.length;
}
function removeDuplicateParcels(p,report){
 const original=text(p);
 if(!/Относно: Уведомление|С инвестиционното предложение се предвижда|В тази връзка и на основание/u.test(original))return;
 const re=/в поземлен(?:и имоти| имот)\s+с\s+идентификатор(?:и)?\s+(\d{5}\.\d+\.\d+)/gu;
 const m=[...original.matchAll(re)];
 if(m.length<2||m[0][1]!==m[1][1])return;
 // Keep the first real cadastral/location block and preserve a distinct area sentence.
 // Only remove the repeated parcel/location sequence; no new data are invented.
 const second=m[1].index;
 const endSentence=original.indexOf(' В представеното уведомление',second);
 const end=endSentence>=0?endSentence:original.length;
 const replacementStart=original.slice(m[0].index,second).replace(/[\s.“]+$/u,'').trimEnd();
 const prefix=original.slice(0,m[0].index);
 const suffix=original.slice(end);
 const final=prefix+replacementStart+(suffix||(/[.]\s*$/u.test(original)?'.':''));
 if(!final.includes(m[0][1]))throw Error('KT79: parcel lost during duplicate cleanup');
 alter(p,0,original.length,final);
 report.edits.push('duplicate parcel occurrence removed without title substitution');
}
function heading(p,report){
 const t=text(p);if(!t.startsWith('Относно: Уведомление'))return;
 const subject=t.match(/„[^“]+“/u);
 if(!subject)throw Error('KT79: unparseable approved subject heading');
 setSpanStyle(p,0,'Относно'.length,{bold:true,italic:true});
 setSpanStyle(p,'Относно'.length,subject.index,{italic:true});
 setSpanStyle(p,subject.index,subject.index+subject[0].length,{bold:true,italic:true});
 setSpanStyle(p,subject.index+subject[0].length,t.length,{bold:false,italic:true});
 report.edits.push('subject heading exact approved bold/italic');
}
function safeAdminFromClassic(ps){
 const line=ps.find(p=>text(p).startsWith('Относно:'));if(!line)return null;
 const t=text(line),m=t.match(/община\s+([^,]+),\s*област\s+([^,.]+)/u);
 if(!m||/Х{4,}/u.test(m[1]+m[2]))return null;
 const loc=(t.match(/местност\s+„([^“]+)“/u)||[])[1];
 const paired=[...ps.flatMap(p=>[...text(p).matchAll(/BG\d{7}\s*„[^“]+“/gu)].map(m=>m[0]))];
 return {municipality:m[1].trim(),region:m[2].trim(),locality:loc||null,zones:[...new Set(paired)]};
}
function normalizeFacts(p,facts,report){
 if(!facts)return;
 if(/община\s+Х{4,}|област\s+Х{4,}/u.test(text(p))){
  if(replaceOnce(p,/община\s+Х{4,}/u,'община '+facts.municipality))report.edits.push('confirmed municipality');
  if(replaceOnce(p,/област\s+Х{4,}/u,'област '+facts.region))report.edits.push('confirmed region');
 }
 if(facts.locality&&/местност\s+[А-Я\s–-]{5,}/u.test(text(p))&&!text(p).includes('местност „')){
  replaceOnce(p,/местност\s+[А-Я\s–-]{5,}/u,'местност „'+facts.locality+'“');
  report.edits.push('confirmed locality');
 }
 for(const paired of facts.zones){
  const code=paired.slice(0,9);
  const re=new RegExp(code+'(?!\\s*„)','gu');
  for(let guard=0;guard<10&&replaceOnce(p,re,paired);guard++)report.edits.push('confirmed full ZZ code and name');
 }
}
function semanticSpans(p,report){
 const t=text(p),gis=/пространствен|границите на|защитен[аи] зони?|защитени територии|ГИС анализ/u.test(t);
 // Remove accidental emphasis inside the complete, different word "попадане".
 if(gis)report.fixedWords+=setMatches(p,/(?<![\p{L}\p{N}])попадане(?![\p{L}\p{N}])/gu,{bold:false});
 if(gis){
  report.bold+=setMatches(p,/(?<![\p{L}\p{N}])не\s+попада(?:т)?(?![\p{L}\p{N}])/gu,{bold:true});
  // A standalone positive "попада" is never a substring of "не попада".
  const s=text(p);
  for(const m of s.matchAll(/(?<![\p{L}\p{N}])попада(?![\p{L}\p{N}])/gu)){
   const before=s.slice(Math.max(0,m.index-5),m.index);
   if(/не\s+$/u.test(before))continue;
   setSpanStyle(p,m.index,m.index+m[0].length,{bold:true});report.bold++;
  }
 }
 report.bold+=setMatches(p,/BG\d{7}\s*„[^“]+“/gu,{bold:true});
 report.bold+=setMatches(p,/(?:природна забележителност|защитена местност|резерват|поддържан резерват|национален парк|природен парк)\s*„[^“]+“/gu,{bold:true});
}
function standardizeClassicB(ps,report){
 const headingText=text(ps.find(p=>text(p).startsWith('Относно:'))||ps[10]);
 const subject=headingText.match(/„([^“]+)“/u);
 if(!subject)throw Error('KT79: Classic B subject missing');
 let para=ps.find(p=>text(p).startsWith('В тази връзка и на основание чл. 2, ал. 2'));
 if(!para)throw Error('KT79: Classic B approved closing paragraph absent');
 const old=text(para);
 if(old.includes('изисква проверка')&&old.includes(subject[1])){
  // The already accepted Scenario B conclusion is a locked, scenario-specific
  // template phrase. ZZ/ZT must never select a different conclusion text.
  const begin=old.indexOf('изисква проверка');
  alter(para,begin,old.length,APPROVED_B_END);
  report.edits.push('Classic B template conclusion aligned, including with ZT');
 }
 if(!text(para).includes(APPROVED_B_END))throw Error('KT79: Classic B conclusion differs from locked template');
 setMatches(para,/„[^“]+“/gu,{bold:true});
 setMatches(para,/не е необходимо провеждане на процедура по реда на Глава втора от Наредбата за ОС\./gu,{bold:true});
 let p23=ps.find(p=>text(p)===APPROVED23_TEXT);
 if(!p23){
   p23=ps.find(p=>/За евентуална промяна на предназначението/u.test(text(p)));
   if(!p23)throw Error('KT79: mandatory approved Classic B land-use paragraph missing');
   alter(p23,0,text(p23).length,APPROVED23_TEXT);
   report.edits.push('mandatory approved Classic B paragraph 23 restored generically');
 }
 setSpanStyle(p23,0,text(p23).length,{bold:true});
}
function fixOne(raw,{classic,facts}){
 const z=root.fflate.unzipSync(raw),old=z['word/document.xml'];
 if(!old)throw Error('KT79: missing document.xml');
 const doc=new DOMParser().parseFromString(dec.decode(old),'application/xml');
 if(doc.getElementsByTagName('parsererror').length)throw Error('KT79: corrupt Word XML');
 const ps=paragraphs(doc),report={paragraphs:ps.length,edits:[],bold:0,fixedWords:0};
 const isB=classic&&ps.some(p=>/^Имотът, предмет на инвестиционното предложение/u.test(text(p))&&/попада/u.test(text(p)));
 for(const p of ps){
  removeDuplicateParcels(p,report);
  if(/Относно: Уведомление/u.test(text(p)))heading(p,report);
  if(replaceOnce(p,/чл\.\s*10,\s*ал\.\s*1\.\s+от/gu,'чл. 10, ал. 1 от'))report.edits.push('legal citation punctuation');
  if(replaceOnce(p,/\d{5}\.\d+\.\d+(?=не\s+попада)/gu,m=>m[0]+' '))report.edits.push('GIS parcel spacing');
  if(replaceOnce(p,/2018 г\.\./gu,'2018 г.'))report.edits.push('ZT punctuation');
  if(!classic)normalizeFacts(p,facts,report);
  if(!classic&&/пространствен|защитена зона/u.test(text(p))&&/попада изцяло или частично/u.test(text(p))){
   replaceOnce(p,/попада изцяло или частично/u,'попада');
   report.edits.push('GIS intersection: no unsupported full/partial classification');
  }
  semanticSpans(p,report);
 }
 if(isB)standardizeClassicB(ps,report);
 // A final negative spatial statement must not have a bolded positive fragment.
 for(const p of ps){
  const t=text(p);
  if(/\d{5}\.\d+\.\d+не\s+попада/u.test(t))throw Error('KT79: missing separator after parcel');
  if(/попада изцяло или частично/u.test(t))throw Error('KT79: unsupported spatial class remains');
 }
 z['word/document.xml']=enc.encode(new XMLSerializer().serializeToString(doc));
 const binary=root.fflate.zipSync(z,{level:6});
 if(text(ps[10]).length===0)throw Error('KT79: lost heading');
 return {binary,report,ps};
}
async function fix(blob){
 if(!root.fflate?.unzipSync||!root.fflate?.zipSync)throw Error('KT79: ZIP engine missing');
 const outer=root.fflate.unzipSync(new Uint8Array(await blob.arrayBuffer()));
 if(!outer[CLASSIC]||!outer[EXPERT])throw Error('KT79: missing original Word documents');
 const originalKeys=Object.keys(outer).sort();
 const c=fixOne(outer[CLASSIC],{classic:true,facts:null});
 const facts=safeAdminFromClassic(c.ps);
 const e=fixOne(outer[EXPERT],{classic:false,facts});
 outer[CLASSIC]=c.binary;outer[EXPERT]=e.binary;
 if(JSON.stringify(Object.keys(outer).sort())!==JSON.stringify(originalKeys))throw Error('KT79: ZIP files changed');
 const bytes=root.fflate.zipSync(outer,{level:6});
 const after=root.fflate.unzipSync(bytes);
 for(const k of originalKeys)if(k!==CLASSIC&&k!==EXPERT){
  const a=outer[k],b=after[k];if(a.length!==b.length||a.some((v,i)=>v!==b[i]))throw Error('KT79: non-Word resource mutated '+k);
 }
 return {blob:new Blob([bytes],{type:'application/zip'}),reports:{Classic:c.report,Expert:e.report},acceptedTemplate:true};
}
root.KT79FinalApproved=Object.freeze({fix});
})(window);
