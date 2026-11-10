/* KT79-24: test the actual inline pre-Word joint quality gate,
 * using synthetic AI responses. Never bypass legal or KT42 validation. */
'use strict';
const fs=require('node:fs'),assert=require('node:assert/strict');
const html=fs.readFileSync('releases/v27.34/index.html','utf8');
const first=html.indexOf('// KT79-24: a SINGLE combined');
const last=html.indexOf('if(!window.KT6722CaseNormalize?.normalize)',first);
assert.ok(first>0&&last>first,'joint quality gate must be present');
const section=html.slice(first,last);
const AsyncFunction=Object.getPrototypeOf(async function(){}).constructor;
const run=new AsyncFunction('input','setStatus',
 'let kt67Plan=input.first;let kt60={scenario:input.scenario||"B"};let v27AiDraft="";'+
 'const found=[{input:"12345.678.901"}],workingDraft={},kt60QualityEvidence={facts:[]};'+
 'const kt7605RetryAuthor=input.retry;const window={V26NormativeRegister:null,'+
 'V27VerifiedContext:{legal:()=>({sources:[]})},'+
 'KT603Evidence:{build:({aiDraft})=>({scenario:aiDraft.scenario||"B"})},'+
 'KT67AIContent:{plan:draft=>draft},'+
 'KT6722LegalQuality:{inspect:p=>({ok:!p.issues.length,issues:p.issues,warnings:[]})}};'+
 section+';return {plan:kt67Plan,draft:v27AiDraft,validated:kt6722Legal.ok,missing:kt77MissingImpacts};');
const full=()=>Array(5).fill('Въздействията са оценени въз основа на заявените в преписката факти.');
const plan=(issues=[],impacts=full(),scenario='B')=>({readyForIntegration:true,classic:{impactsProject:impacts},issues,scenario});
const issue=code=>({code});
(async()=>{
 let events=[];
 let a=await run({first:plan(),retry:null},s=>events.push(s));
 assert.equal(a.validated,true);
 let attempts=0;
 let candidates=[plan([issue('UNSUPPORTED_CATEGORICAL_OS')]),plan([],full())];
 let b=await run({first:plan([issue('UNSUPPORTED_CATEGORICAL_OS')],['НЯМА ДАННИ',...full().slice(1)]),retry:async()=>{attempts++;return candidates.shift();}},s=>events.push(s));
 assert.equal(attempts,2,'do not stop at a draft that fails legal or impact validation');
 assert.equal(b.validated,true);assert.equal(b.missing.length,0);
 attempts=0;
 candidates=[plan([],['НЯМА ДАННИ',...full().slice(1)]),plan([],full())];
 b=await run({first:plan([],['НЯМА ДАННИ',...full().slice(1)]),retry:async()=>{attempts++;return candidates.shift();}},s=>events.push(s));
 assert.equal(attempts,2);assert.equal(b.missing.length,0);
 let rejected=false;
 try{await run({first:plan([issue('UNSUPPORTED_CATEGORICAL_OS')]),retry:async()=>plan([issue('UNSUPPORTED_CATEGORICAL_OS')])},()=>{});}
 catch(e){rejected=/Не е издаден ZIP/.test(String(e.message));}
 assert.ok(rejected,'unsupported legal assertions never produce ZIP');
 let forbiddenAttempts=0;
 rejected=false;
 try{await run({first:plan([issue('UNVERIFIED_ORDER')]),retry:async()=>{forbiddenAttempts++;return plan();}},()=>{});}
 catch(e){rejected=/неприет правен контрол/.test(String(e.message));}
 assert.ok(rejected);assert.equal(forbiddenAttempts,0,'non-retryable legal defects must fail closed');
 rejected=false;
 try{await run({first:plan([],full()),retry:null,scenario:'A'},()=>{});}
 catch(e){rejected=true;}
 assert.equal(rejected,false);
 console.log('PASS KT79-24 joint gate 5/5: legal + 5 impacts, bounded retry, fail closed; no generator changes');
})().catch(e=>{console.error(e);process.exitCode=1});
