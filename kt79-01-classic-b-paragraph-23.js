/* KT79-01 — explicitly approved Classic B paragraph 23, all B dossiers.
 * Applies only after existing locked-Classic and cadastral guards have passed.
 * Does not modify Expert Word, hard paragraphs elsewhere, GIS or PDF.
 * Preserves paragraph 23 run/style and confirms exactly one run.
 */
(function(root){'use strict';
 const previous=root.KT7710LockedClassic;
 if(!previous?.fix||!root.fflate?.unzipSync){console.error('KT79-01: required Classic B generator unavailable');return;}
 const FILE='01_Proekt_na_stanovishte_Klasicheski.docx';
 const NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
 const APPROVED='Инвестиционното предложение не е свързано с промяна предназначението и/или начина на трайно ползване на имота.';
 const LEGACY='За евентуална промяна на предназначението и/или начина на трайно ползване на имота няма потвърдени данни в представената документация.';
 const decode=new TextDecoder(),encode=new TextEncoder();
 const paras=doc=>Array.from(doc.getElementsByTagNameNS(NS,'body')[0]?.childNodes||[]).filter(x=>x.nodeType===1&&x.localName==='p'&&x.namespaceURI===NS);
 const read=p=>Array.from(p.getElementsByTagNameNS(NS,'t')).map(x=>x.textContent||'').join('');
 async function fix(blob,options={}){
  const prior=await previous.fix(blob,options);
  const scenario=options.scenario==='B'?'Б':options.scenario;
  if(scenario!=='Б')return prior;
  const entries=root.fflate.unzipSync(new Uint8Array(await prior.blob.arrayBuffer()));
  if(!entries[FILE])throw Error('КТ79-01: липсва Classic Б');
  const doc=root.fflate.unzipSync(entries[FILE]);
  const xml=new DOMParser().parseFromString(decode.decode(doc['word/document.xml']),'application/xml');
  if(xml.getElementsByTagName('parsererror').length)throw Error('КТ79-01: невалиден DOCX XML');
  const ps=paras(xml);
  if(ps.length!==43)throw Error('КТ79-01: неочакван Classic Б формат');
  const p=ps[22],before=read(p);
  if(before===APPROVED)return {...prior,kt7901:'approved-already'};
  if(before!==LEGACY)throw Error('КТ79-01: непознат текст на абзац №23 — необходима експертна проверка');
  const texts=Array.from(p.getElementsByTagNameNS(NS,'t'));
  if(texts.length!==1)throw Error('КТ79-01: оформлението на №23 е променено');
  texts[0].textContent=APPROVED;
  if(read(p)!==APPROVED)throw Error('КТ79-01: неуспешно уеднаквяване');
  // Mutate only the DOCX document.xml, preserving run properties and all
  // additional Word resources exactly. Never touch the Expert file.
  doc['word/document.xml']=encode.encode(new XMLSerializer().serializeToString(xml));
  entries[FILE]=root.fflate.zipSync(doc,{level:6});
  const zip=new Blob([root.fflate.zipSync(entries,{level:6})],{type:'application/zip'});
  return {...prior,blob:zip,kt7901:'approved-paragraph-23'};
 }
 root.KT7710LockedClassic=Object.freeze({...previous,fix});
})(window);
