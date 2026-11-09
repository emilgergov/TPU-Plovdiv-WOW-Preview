/* Offline structural regression for KT79-10 universal Word adapter.
   Uses SYNTHETIC dossiers with invented IDs solely to prove no parcel-ID branching.
   Never treats fictional dossiers as authoritative legal/GIS evidence. */
'use strict';
const assert=require('node:assert/strict'), fs=require('node:fs'),vm=require('node:vm');
const {DOMParser,XMLSerializer}=require('@xmldom/xmldom');
const zip=require('fflate');
const enc=new TextEncoder(),dec=new TextDecoder();
const script=fs.readFileSync('kt79-10-universal-word-gate.js','utf8');
const window={fflate:zip};
vm.runInNewContext(script,{window,DOMParser,XMLSerializer,TextEncoder,TextDecoder,Blob,console});
const gate=window.KT79FinalApproved;
assert.equal(typeof gate.fix,'function');
const N='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function xmlParagraph(s,runs=null){
 const chunks=runs||[{text:s}];
 return '<w:p>'+chunks.map(r=>'<w:r><w:rPr>'+(r.bold?'<w:b/>':'')+(r.italic?'<w:i/>':'')+'</w:rPr><w:t xml:space="preserve">'+r.text.replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</w:t></w:r>').join('')+'</w:p>';
}
function docx(paras){
 const xml='<?xml version="1.0" encoding="UTF-8"?>'+
 '<w:document xmlns:w="'+N+'"><w:body>'+paras.map(p=>typeof p==='string'?xmlParagraph(p):xmlParagraph('',p)).join('')+'</w:body></w:document>';
 return zip.zipSync({'word/document.xml':enc.encode(xml)});
}
function content(data){
 const doc=new DOMParser().parseFromString(dec.decode(zip.unzipSync(data)['word/document.xml']),'application/xml');
 const ps=Array.from(doc.getElementsByTagNameNS(N,'body')[0].childNodes).filter(e=>e.nodeType===1&&e.localName==='p');
 const runs=p=>Array.from(p.getElementsByTagNameNS(N,'r')).map(r=>({
   text:Array.from(r.getElementsByTagNameNS(N,'t')).map(t=>t.textContent||'').join(''),
   bold:Array.from(r.getElementsByTagNameNS(N,'b')).some(b=>b.getAttributeNS(N,'val')!=='0'),
   italic:Array.from(r.getElementsByTagNameNS(N,'i')).some(b=>b.getAttributeNS(N,'val')!=='0')
 }));
 return {ps,runs,txt:i=>runs(ps[i-1]).map(x=>x.text).join('')};
}
function phraseRuns(o,i,phrase){
 let pos=0;return o.runs(o.ps[i-1]).filter(r=>{const a=pos;pos+=r.text.length;const full=o.txt(i),idx=full.indexOf(phrase);return idx>=0&&a<idx+phrase.length&&pos>idx;});
}
function synthetic({mode='Б',id='12345.678.901',zt=false,duplicate=false,multi=false,wrongSource=false}={}){
 const name=mode==='Б'?'Обслужваща площадка':'Складови помещения';
 const zone='BG0001031 „Родопи – Средни“';
 const h='Относно: Уведомление за инвестиционно предложение (ИП): „'+name+'“ в поземлен имот с идентификатор '+id+', населено място с. Примерно, община Тестова, област Примерна.';
 const p=Array.from({length:mode==='Б'?44:42},()=>''),e=Array.from({length:50},()=> '');
 p[10]=mode==='А'?
  'Относно: Уведомление за инвестиционно предложение (ИП): „'+name+'“ в поземлени имоти с идентификатори '+id+', 12345.678.902, 12345.678.903, населено място гр. Примерно, община Тестова, област Примерна, местност РОГОШКО ШОСЕ-СЕВЕР.' :
  (duplicate?h+' в поземлен имот с идентификатор '+id+', населено място с. Примерно, община Тестова, област Примерна.':h);
 e[10]='Относно: Уведомление за инвестиционно предложение (ИП): „'+name+'“ в поземлен имот с идентификатор '+id+', населено място с. Примерно, община ХХХХХХХХ, област ХХХХХХХХ'+(zt?', местност ОБЩИЯ БАЛКАН.':'.');
 p[17]='Инвестиционното предложение попада в обхвата на чл. 2, ал. 2, по приложение № 1 към чл. 10, ал. 1. от същата.';
 e[18]=p[17];
 if(mode==='Б'){
  p[18]='С инвестиционното предложение се предвижда '+name+' в поземлен имот с идентификатор '+id+'.';
  p[22]='За евентуална промяна на предназначението и/или начина на трайно ползване на имота няма потвърдени данни в представената документация.';
  p[23]=multi?('Имотите, предмет на инвестиционното предложение, с идентификатори '+id+', 12345.678.999 попадат в границите на защитена зона '+zone+'.'):('Имотът, предмет на инвестиционното предложение, с идентификатор '+id+' попада в границите на защитена зона '+zone+'.');
  p[24]='Защитена зона '+zone+' за опазване на природните местообитания е обявена със заповед.';
  p[25]=zt?
   'При извършената пространствена проверка е установено, че поземленият имот с идентификатор '+id+' попада в границите на природна забележителност „Свети Дух“.' :
   [{text:'При извършената пространствена проверка е установено, че поземленият имот с идентификатор '+id+' не '},{text:'попада',bold:true},{text:' в границите на защитени територии.'}];
  p[34]='В тази връзка и на основание чл. 2, ал. 2 от Наредбата за ОС Ви уведомявам, че преценката на компетентния орган за така заявеното инвестиционно предложение: „'+name+'“ в поземлен имот с идентификатор '+id+' изисква проверка на заявените параметри, приложимите режими на '+zone+' и резултатите от пространствената проверка, преди определяне на окончателния процедурен извод.';
  e[23]='Поземленият имот с идентификатор '+id+' попада изцяло или частично в границите на защитена зона BG0001031.';
  e[24]='Приложими са режимите на BG0001031.';
  e[25]=p[25];
  e[32]=[{text:'Изводът за пространствено '},{text:'попада',bold:true},{text:'не на имота се прави с оглед на BG0001031.'}];
  e[38]='4. ПРЕЦЕНКА НА КОМПЕТЕНТНИЯ ОРГАН';
  e[39]='Настоящият проектен статус не позволява категорично определяне на процедурното ниво или необходимостта от екологична оценка.';
  e[40]='Предварителното мнение е, че е възможно да се приложи опростен административен режим след допълнителна експертна проверка.';
 }else{
  p[22]='Поземленият имот с идентификатор '+id+' не попада в границите на защитени зони.';
  p[24]='Поземленият имот с идентификатор '+id+'не попада в границите на защитени територии.';
  e[23]='При пространствен анализ не е установено попадане на имота в зона.';
  p[31]='Въздействие на предложените дейности причинени поотделно или във взаимодействие с други инвестиционни предложения/планове, програми, проекти (ИП/ППП).';
  p[34]='При всички случаи на промяна в параметрите ИП или на някои от обстоятелствата възложителят е длъжен да уведоми РИОСВ.';
 }
 const files={
 '01_Proekt_na_stanovishte_Klasicheski.docx':docx(p),
 '02_Proekt_na_stanovishte_Eksperten.docx':docx(e),
 '03_GIS_analiz.pdf':enc.encode('%PDF-FIXTURE-A'),
 '04_Normativni_i_dokumentalni_referencii.pdf':enc.encode('%PDF-FIXTURE-B'),
 'README_BG.txt':enc.encode('FIXTURE')
 };
 return {files,blob:new Blob([zip.zipSync(files)],{type:'application/zip'}),id};
}
async function run(){
 let passes=0;
 for(const [mode,id,zt,multi] of [['А','12345.678.901',false,false],['Б','12345.678.901',false,false],['Б','98765.432.109',true,false],['Б','13579.246.810',false,true]]){
  const sample=synthetic({mode,id,zt,duplicate:zt,multi});
  const result=await gate.fix(sample.blob);
  const out=zip.unzipSync(new Uint8Array(await result.blob.arrayBuffer()));
  assert.deepEqual(Object.keys(out).sort(),Object.keys(sample.files).sort());
  for(const name of Object.keys(sample.files).filter(x=>!x.endsWith('.docx')))assert.deepEqual(out[name],sample.files[name]);
  const c=content(out['01_Proekt_na_stanovishte_Klasicheski.docx']);
  const e=content(out['02_Proekt_na_stanovishte_Eksperten.docx']);
  assert.ok(c.txt(11).includes(id),'subject must retain exact arbitrary cadastral identifier');
  assert.ok(c.txt(11).includes('Тестова'),'subject must retain input municipality');
  if(zt)assert.equal((c.txt(11).match(new RegExp(id.replaceAll('.','\\.'),'g'))||[]).length,1,'duplicated parcel in heading forbidden');
  assert.equal(phraseRuns(c,11,'Относно').every(r=>r.bold&&r.italic),true,'heading bold and italic');
  for(const obj of [c,e]){
   assert.equal(phraseRuns(obj,11,'Уведомление за инвестиционно предложение (ИП)').every(r=>!r.bold&&r.italic),true,'heading descriptor italic and NOT bold');
   assert.equal(phraseRuns(obj,11,'Относно:').every(r=>r.bold&&r.italic),true,'whole label bold italic');
  }
  assert.equal(phraseRuns(c,11,'„'+(mode==='Б'?'Обслужваща площадка':'Складови помещения')+'“').every(r=>r.bold&&r.italic),true,'subject title bold+italic');
  assert.ok(!c.txt(18).includes('ал. 1. от'));
  assert.ok(!e.txt(19).includes('ал. 1. от'));
  assert.ok(e.txt(11).includes('община Тестова, област Примерна'));
  if(mode==='Б'){
   if(multi)assert.equal(phraseRuns(c,24,'попадат').every(r=>r.bold),true,'plural positive relation must be entirely bold');
   assert.ok(c.txt(23).includes('не е свързано с промяна предназначението'));
   assert.ok(c.txt(35).endsWith('е, че не е необходимо провеждане на процедура по реда на Глава втора от Наредбата за ОС.'));
   assert.ok(!c.txt(35).includes('изисква проверка'));
   assert.ok(e.txt(24).includes('BG0001031 „Родопи – Средни“'));
   assert.ok(e.txt(25).includes('BG0001031 „Родопи – Средни“'));
   assert.ok(e.txt(33).includes('BG0001031 „Родопи – Средни“'));
   assert.ok(!e.txt(24).includes('попада изцяло или частично'));
   assert.ok(e.txt(41).includes('компетентният орган приема, че не е необходимо провеждане на процедура'),'Expert B approved final authority voice');
   assert.ok(!e.txt(41).includes('Предварителното мнение'),'Expert B no provisional AI opinion');
   assert.ok(!e.txt(40).includes('проектен статус'),'Expert B no provisional legal assessment');
   const incomplete=phraseRuns(e,33,'попадане');assert.equal(incomplete.some(r=>r.bold),false,'inner word попадане must not be bold');
   if(zt)assert.ok(c.txt(11).includes('местност „Общия балкан“'),'confirmed B+ZT locality must be retained');
   if(!zt){
    const neg=phraseRuns(c,26,'не попада');assert.equal(neg.every(r=>r.bold),true,'entire negative relation must be bold');
   }
  }else{
   assert.ok(c.txt(11).includes('местност „Рогошко шосе - север“'),'Classic A locality approved');
   assert.ok(c.txt(11).includes('12345.678.902 и 12345.678.903'),'Classic A correct final conjunction in parcel list');
   assert.ok(c.txt(32).startsWith('Въздействия, произтичащи от предложените дейности'),'Classic A cumulative impacts approved');
   assert.ok(c.txt(35).startsWith('При промяна в параметрите на ИП'),'Classic A amendment notification approved');
   assert.ok(c.txt(25).includes(id+' не попадат')||c.txt(25).includes(id+' не попада'));
   assert.equal(phraseRuns(c,25,'не попада').every(r=>r.bold),true);
  }
  passes++;console.log('PASS',mode,zt?'B+ZT':'',id);
 }
 console.log('PASS KT79 universal synthetic regression',passes+'/4; arbitrary cadastral IDs including multi-parcel B; all non-DOCX files byte-identical');
}
run().catch(e=>{console.error(e);process.exitCode=1});
