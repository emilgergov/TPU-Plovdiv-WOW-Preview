/* KT77 unified 03 — verified cadastral administrative-location bridge for locked Classic B.
 * Administrative location comes only from checksum-verified PI MASTER; dossier must
 * independently identify the same parcel. Does not touch Word templates/ZIP contents. */
(function(root){'use strict';
const original=root.KT7710LockedClassic;
const reader=root.KT77PiMasterIndex;
if(!original?.fix||!reader?.create){console.error('KT77 admin bridge dependencies missing');return;}
const index=reader.create('https://emilgergov.github.io/TPU-Plovdiv-WOW-2.0/v26-test/gis-data/pi-master/browser-index/');
const locationPattern=/(?:по\s+КККР\s+на\s+)?((?:с\.|гр\.)\s*[А-Яа-я\- ]+),\s*община\s+([А-Яа-я\- ]+),\s*област\s+([А-Яа-я\- ]+?)(?:[.,\n]|\s+местност)/u;
const clean=x=>String(x||'').trim();
const bad=x=>!x||/^(?:няма данни|неизвестно|undefined|null|[XХ]+)$/iu.test(x);
async function fix(blob,options={}){
 const {scenario,parcelIds=[],archivePreview}=options;
 if(!['Б','B'].includes(scenario)||parcelIds.length!==1||!Array.isArray(archivePreview?.files))return original.fix(blob,options);
 const id=String(parcelIds[0]);
 const files=archivePreview.files.filter(f=>typeof f?.text==='string');
 const notices=files.filter(f=>/уведомлен|notification/i.test(String(f.path||f.file||f.name||'')));
 const selected=notices.length?notices:files;
 const text=selected.slice(0,8).map(f=>f.text).join('\n');
 // Do not synthesize an investment proposal from cadastral data alone.
 if(!text.includes(id))return original.fix(blob,options);
 if(locationPattern.test(text))return original.fix(blob,options);
 const result=await index.lookup(id);
 if(result.status!=='FOUND'||!result.profile)throw Error('КТ77-18: няма верифицирани ПИ MASTER данни за административно местоположение');
 const p=result.profile;
 if([p.settlement,p.municipality,p.province].some(bad))throw Error('КТ77-18: ПИ MASTER няма пълни административни данни');
 const settlement=clean(p.settlement).replace(/^населено място\s+/iu,'');
 const municipality=clean(p.municipality),province=clean(p.province);
 // Existing primary-document location data cannot silently be contradicted.
 const candidate='Справка ПИ MASTER за ПИ '+id+': '+settlement+', община '+municipality+', област '+province+'.';
 const nextFiles=archivePreview.files.map(f=>{
  if(!selected.includes(f))return f;
  return {...f,text:String(f.text)+'\n'+candidate};
 });
 // The unchanged KT77-10 generator independently validates the content and DOCX resources.
 return original.fix(blob,{...options,archivePreview:{...archivePreview,files:nextFiles}});
}
root.KT7710LockedClassic=Object.freeze({...original,fix});
})(window);
