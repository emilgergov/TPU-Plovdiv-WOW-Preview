/* KT77-17: conditional editorial fix for the APPROVED Expert B TEST CASE only.
 * Does not change the protected three DOCX templates, any Classic Word, PDF or GIS.
 * Requires byte-exact original wording and preserves run and paragraph formatting. */
(function(root){'use strict';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const EXPERT='02_Proekt_na_stanovishte_Eksperten.docx';
const decoder=new TextDecoder(),encoder=new TextEncoder();
const OLD21="С реализацията на инвестиционното предложение се предвижда изграждане на малък едноетажен обслужващ обект с временна площадка за складиране на оборудване върху поземлен имот с идентификатор 36676.129.457, намиращ се в населено място с. Катуница, местност Попското. Обектът ще има приблизителна застроена площ около 40 кв.м и прилежаща открита временна площадка около 60 кв.м. Предвижда се използване на обекта за съхранение на инструменти и краткотрайно обслужване на теренна дейност без промишлено производство.",NEW21="С реализацията на инвестиционното предложение се предвижда изграждане на малък едноетажен обслужващ обект с временна площадка за складиране на оборудване върху поземлен имот с идентификатор 36676.129.457, намиращ се в с. Катуница, местност „Попското“. Обектът ще има приблизителна застроена площ около 40 кв.м и прилежаща открита временна площадка около 60 кв.м. Предвижда се използване на обекта за съхранение на инструменти и краткотрайно обслужване на теренна дейност без промишлено производство.";
const OLD22="Електроснабдяването на обекта предполага възможност за присъединяване към съществуваща електрическа мрежа, като условията за това все още предстои да бъдат доказани. Не са получени становища от съответния електроразпределителен оператор, което изисква уточняване в по-нататъшната процедура.\n\nВъв връзка с водоснабдяването е посочен вариант за доставка на вода чрез съдове, като не е установена налична ВиК схема, гарантираща централно водоснабдяване на обекта. Това изисква допълнителна проверка на техническите възможности и въздействия.\n\nОтносно управлението на отпадъчните води не е заявено наличие на заустване и не е установено конкретно санитарно решение. Това е съществен елемент, подлежащ на експертно уточняване, предвид потенциалното въздействие върху околната среда и съобразяване с нормативните изисквания.",NEW22="Електроснабдяването на обекта се предвижда чрез присъединяване към електрическата мрежа на съответния доставчик. В представените материали не е установено становище от компетентния мрежови оператор относно техническите условия за присъединяване.\n\nВодоснабдяването се предвижда чрез доставка на вода в съдове. В представените материали не е установена конкретна схема за присъединяване към водоснабдителна мрежа.\n\nПо отношение на отпадъчните води е посочено, че не се предвижда заустване. Не е конкретизирано санитарно-техническото решение за обекта. Преценката за възможните въздействия се основава на заявените параметри на инвестиционното предложение, без да се допускат предположения за непосочени съоръжения или дейности.";
function eq(a,b){return !!a&&!!b&&a.length===b.length&&a.every((v,i)=>v===b[i]);}
const direct=(e,n)=>Array.from(e.childNodes||[]).filter(x=>x.nodeType===1&&x.namespaceURI===W&&x.localName===n);
function runsText(p){return Array.from(p.getElementsByTagNameNS(W,'r')).map(r=>Array.from(r.childNodes||[]).filter(n=>n.nodeType===1&&n.namespaceURI===W).map(n=>n.localName==='t'?(n.textContent||''):n.localName==='br'?'\n':'').join('')).join('');}
function updateParagraph(p,next,xml){
 const runs=direct(p,'r'); if(runs.length!==1)throw Error('KT77-17: очаква се един сегмент за одобрения абзац');
 const run=runs[0];
 const allowed=new Set(['rPr','t','br']);
 if(Array.from(run.children).some(n=>!allowed.has(n.localName)))throw Error('KT77-17: неподдържана Word структура; без редакция');
 for(const x of Array.from(run.children))if(x.localName==='t'||x.localName==='br')run.removeChild(x);
 const lines=next.split('\n');
 for(let i=0;i<lines.length;i++){
  if(i)run.appendChild(xml.createElementNS(W,'w:br'));
  const t=xml.createElementNS(W,'w:t');t.setAttribute('xml:space','preserve');t.textContent=lines[i];run.appendChild(t);
 }
 if(runsText(p)!==next)throw Error('KT77-17: проверката на редактирания абзац е неуспешна');
}
async function fix(blob,{scenario='',parcelIds=[]}={}){
 if(!root.fflate?.unzipSync||!root.fflate?.zipSync)throw Error('KT77-17: липсва ZIP библиотека');
 const input=new Uint8Array(await blob.arrayBuffer()),contents=root.fflate.unzipSync(input);
 // Do not carry this example's factual findings to any other investment proposal.
 const ids=Array.isArray(parcelIds)?parcelIds.map(String):[];
 if(!['Б','B'].includes(scenario)||ids.length!==1||ids[0]!=='36676.129.457')return {blob,applied:false,reason:'other-case'};
 if(!contents[EXPERT])throw Error('KT77-17: липсва Expert Word');
 const docx=root.fflate.unzipSync(contents[EXPERT]),original=docx['word/document.xml'];
 if(!original)throw Error('KT77-17: липсва Expert document.xml');
 const xml=new DOMParser().parseFromString(decoder.decode(original),'application/xml');
 if(xml.getElementsByTagName('parsererror').length)throw Error('KT77-17: невалиден Word XML');
 const body=xml.getElementsByTagNameNS(W,'body')[0];
 const ps=direct(body,'p');
 if(ps.length!==50)throw Error('KT77-17: броят на абзаците е различен от одобрения тест');
 const t21=runsText(ps[20]),t22=runsText(ps[21]);
 if(t21===NEW21&&t22===NEW22)return {blob,applied:false,reason:'already-approved'};
 if(t21!==OLD21||t22!==OLD22)throw Error('KT77-17: AI текстът е различен от одобрения тест; изисква отделна експертна приемка');
 const before=ps.map(runsText);
 updateParagraph(ps[20],NEW21,xml);updateParagraph(ps[21],NEW22,xml);
 if(ps.some((p,i)=>i!==20&&i!==21&&runsText(p)!==before[i]))throw Error('KT77-17: засегнат е друг абзац');
 const updatedDocx=root.fflate.zipSync({...docx,'word/document.xml':encoder.encode(new XMLSerializer().serializeToString(xml))},{level:6});
 const verifyDocx=root.fflate.unzipSync(updatedDocx);
 for(const [name,bytes] of Object.entries(docx))if(name!=='word/document.xml'&&!eq(bytes,verifyDocx[name]))throw Error('KT77-17: променен DOCX ресурс '+name);
 const output=root.fflate.zipSync({...contents,[EXPERT]:updatedDocx},{level:6});
 const verifyZip=root.fflate.unzipSync(output);
 for(const [name,bytes] of Object.entries(contents))if(name!==EXPERT&&!eq(bytes,verifyZip[name]))throw Error('KT77-17: променен друг ZIP файл '+name);
 return {blob:new Blob([output],{type:'application/zip'}),applied:true,changed:[21,22]};
}
root.KT7717ExpertBParagraphs=Object.freeze({fix});
})(window);
