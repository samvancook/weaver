import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
import {INVENTORY_ID,prepareEventIntake,footageEventAsSet,currentEventReceipt} from '../footage_event_intake.mjs';
const server=readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
const getFunction=name=>server.match(new RegExp(`async function ${name}\\([\\s\\S]*?\\n}\\n`))[0];
const id='video_individual_123456';
const manifest={sourceSpreadsheetId:INVENTORY_ID,eventName:'Pilot',omFolderId:'event_folder_123456789',closed:false,
 titles:[{sourceRow:6000,performer:'Inventory poet',title:'Inventory title',rights:'Have video release',contentLink:`https://drive.google.com/file/d/${id}/view`}]};
test('receiver saves before acknowledgement, reloads queue after restart, and curation uses inventory titles',async()=>{
 const ledger=new Map(),sets=new Map();let reply;
 const context={prepareEventIntake,footageEventAsSet,currentEventReceipt,PRIORITY_VIDEO_SETS:sets,
  process:{env:{WEAVER_FOOTAGE_BRIDGE_ENABLED:'true'}},decodeURIComponent,
  verifyFootageHelperCaller:async()=>({email:'owner@example.com'}),
  getDriveFileMetadata:async fileId=>({id:fileId,mimeType:'video/mp4',name:'Clip 1.mp4'}),
  readRequestBody:async()=>JSON.stringify(manifest),
  sendJson:(_res,status,body)=>{reply={status,body};},
  cleanSheetWhitespace:x=>String(x||'').trim(),isPriorityVideoFileIncluded:()=>true,
  getPriorityVideoReleaseStatus:()=> 'inventory_curation_only',reconcileVideoReviews:()=>[],
  parsePriorityVideoFileName:()=>{throw Error('Must not infer imported titles from filename');},
  syncWeaverRuntimeDb:async(action,payload)=>{
   if(action==='save_footage_event'){ledger.set(payload.record.eventId,payload.record);return {ok:true,record:payload.record};}
   if(action==='get_footage_events')return {ok:true,records:payload.eventId?[ledger.get(payload.eventId)].filter(Boolean):[...ledger.values()]};
   if(action==='get_curation_reviews')return {ok:true,records:[]};
   throw Error('Unexpected action '+action);
  }
 };
 const route=server.slice(server.indexOf('  if (url.pathname === "/api/footage/events"'),server.indexOf('  if ((url.pathname.startsWith("/api/intake/") &&'));
 const api=runInNewContext(`(async (req,res,url)=>{${route}})`,context);
 await api({method:'POST'},{},{pathname:'/api/footage/events'});
 assert.equal(reply.status,200);assert.equal(reply.body.receipt.acceptedCount,1);assert.equal(ledger.size,1);
 const eventId=reply.body.receipt.eventId;
 // Simulate process memory loss. Use the production rehydration and playlist functions.
 sets.clear();
 const load=runInNewContext(`${getFunction('listVideoSetFiles')}\n${getFunction('getVideoSetMediaType')}\n${getFunction('refreshFootageEventSets')}\n${getFunction('loadPriorityVideoSet')}\n(async id=>{await refreshFootageEventSets();return loadPriorityVideoSet(id);})`,context);
 const playlist=await load(eventId);
 assert.equal(playlist.count,1);assert.equal(playlist.items[0].author,'Inventory poet');assert.equal(playlist.items[0].poemTitle,'Inventory title');
 assert.equal(playlist.items[0].publicationRestricted,true);assert.equal(playlist.items[0].sourceMediaType,'unknown');
 await api({method:'POST'},{},{pathname:'/api/footage/events'});assert.equal(ledger.size,1);
 await api({method:'GET'},{},{pathname:'/api/footage/events/'+eventId});assert.equal(reply.body.receipt.readyForClosureReview,true);
 context.getDriveFileMetadata=async()=>{throw Error('403');};
 await api({method:'GET'},{},{pathname:'/api/footage/events/'+eventId});assert.equal(reply.body.receipt.readyForClosureReview,false);
});
test('Helper rebuilds current manifest server-side and stale receipts never confirm inventory edits',()=>{
 const source=readFileSync(new URL('../integrations/footage-helper/WeaverBridge.gs',import.meta.url),'utf8');
 const pointers=new Map();let report={eventName:'Pilot',event:{omFolderLink:`https://drive.google.com/drive/folders/${manifest.omFolderId}`},closed:false,
  titles:manifest.titles.map(t=>({...t,rowNumber:t.sourceRow}))};
 let receivedPayload;
 const receipt={...manifest,eventId:'inventory-example',manifest,acceptedCount:1,submittedCount:1,blockedCount:0,readyForClosureReview:true,blocked:[]};
 const api=runInNewContext(source+`;({sendEventToWeaver,checkWeaverEvent})`,{
  RETITLE_CONFIG:{spreadsheetId:INVENTORY_ID},getIntakeReport:()=>structuredClone(report),intakeKey_:x=>String(x).trim().toLowerCase(),
  Utilities:{DigestAlgorithm:{SHA_256:'SHA256'},computeDigest:()=>[1,2,3]},
  ScriptApp:{getIdentityToken:()=> 'identity-token'},
  PropertiesService:{getScriptProperties:()=>({getProperty:key=>key==='WEAVER_BRIDGE_ENABLED'?'true':key==='WEAVER_BRIDGE_URL'?'https://weaver.example.com':pointers.get(key),setProperty:(k,v)=>pointers.set(k,v)})},
  UrlFetchApp:{fetch:(url,options)=>{if(options.payload)receivedPayload=JSON.parse(options.payload);return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({ok:true,receipt})};}}
 });
 assert.equal(api.sendEventToWeaver('Pilot').acceptedCount,1);assert.equal(receivedPayload.titles[0].sourceRow,6000);
 assert.equal(api.checkWeaverEvent('Pilot').readyForClosureReview,true);
 report.titles[0].title='Changed in inventory';assert.equal(api.checkWeaverEvent('Pilot').status,'inventory_changed');
 report.closed=true;assert.throws(()=>api.sendEventToWeaver('Pilot'),/already closed/);
});
test('only the deploying owner can configure the fixed live-test target or read connection identity',()=>{
 const source=readFileSync(new URL('../integrations/footage-helper/ConnectionSetup.gs',import.meta.url),'utf8');
 let active='owner@example.com',saved;
 const api=runInNewContext(source+';({configureWeaverLiveTest,getWeaverConnectionInfo})',{
  Session:{getActiveUser:()=>({getEmail:()=>active}),getEffectiveUser:()=>({getEmail:()=> 'owner@example.com'})},
  PropertiesService:{getScriptProperties:()=>({setProperties:v=>{saved=v;}})},
  weaverConnectionInfo_:()=>({audience:'exact-helper-audience',effectiveEmail:'owner@example.com',emailVerified:true})
 });
 assert.equal(api.configureWeaverLiveTest().audience,'exact-helper-audience');
 assert.equal(saved.WEAVER_BRIDGE_ENABLED,'true');assert.match(saved.WEAVER_BRIDGE_URL,/^https:\/\/footage-bridge---/);
 active='other@example.com';assert.throws(()=>api.configureWeaverLiveTest(),/deploying owner/);assert.throws(()=>api.getWeaverConnectionInfo(),/deploying owner/);
 active='';assert.throws(()=>api.getWeaverConnectionInfo(),/deploying owner/);
});
