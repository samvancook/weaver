import {createHash} from 'node:crypto';
const clean=v=>String(v??'').trim(),norm=v=>clean(v).normalize('NFKC').replace(/\s+/gu,' ').toLowerCase();
export function videoExcerptDeliveryFingerprint(r){const p=r.payload||{},file=clean(r.sourceVideoFileId||p.sourceVideoFileId)||clean(r.sourceVideoUrl||p.sourceVideoUrl).match(/\/file\/d\/([\w-]+)/)?.[1]||'';
 return createHash('sha256').update(JSON.stringify([clean(r.sourceRecordId),norm(r.author),norm(r.poemTitle),norm(r.excerpt),file,clean(r.sourceEvent||p.sourceEvent),clean(r.bookTitle),clean(r.bookShortener),clean(r.releaseCatalog)])).digest('hex');}

export async function runVideoExcerptDeliveries({db,loadApproved,canSend,send,inspect,recordIds=[],budgetMs=45000,max=2,now=Date.now}){
 const started=now(),result=await db('get_video_excerpt_delivery_queue',{});if(!result.ok)throw Error('Durable excerpt queue unavailable');
 const ids=new Set(recordIds),work=(result.records||[]).filter(r=>(!ids.size||ids.has(r.recordId))&&['queued','waiting_video','sending','needs_reconciliation','acknowledged'].includes(r.delivery?.state))
  .sort((a,b)=>String(a.delivery.checkedAt||a.delivery.queuedAt).localeCompare(String(b.delivery.checkedAt||b.delivery.queuedAt))||a.recordId.localeCompare(b.recordId));
 const summary={ok:true,checked:0,sends:0,received:0,held:0};if(!work.length)return summary;
 const byId=new Map((await loadApproved()).map(r=>[r.recordId,r]));
 const touch=(record,reason,error)=>db('review_video_excerpt_delivery',{expectedRecord:record,reason,error});
 for(let record of work){if(summary.checked>=max||now()-started>budgetMs-25000)break;
  if(record.delivery.state==='sending'&&record.delivery.leaseExpiresAt*1000>now())continue;
  summary.checked++;
  const approved=byId.get(record.recordId);
  if(!approved||videoExcerptDeliveryFingerprint(approved)!==record.delivery.sourceFingerprint){
   if(record.delivery.state==='sending'){const expired=await touch(record,'expired','Prior attempt expired while the approved source changed');record=expired.record||record;}
   await touch(record,approved?'source_changed':'approval_withdrawn','Current approval or source version changed; review before delivery');summary.held++;continue;
  }
  if(record.delivery.state==='sending'){const expired=await touch(record,'expired','Prior worker stopped; check receiver before any further send');if(!expired.saved)continue;record=expired.record;}
  if(['acknowledged','needs_reconciliation'].includes(record.delivery.state)){
   try{await inspect(record);summary.received++;}catch(e){await touch(record,'inspection_failed',e.message);summary.held++;}continue;
  }
  let parent;try{parent=await canSend(approved);}catch(e){await touch(record,'waiting_video',e.message);summary.held++;continue;}
  if(!parent){await touch(record,'waiting_video','Waiting for a verified parent video; no excerpt sent');continue;}
  const claim=await db('claim_video_excerpt_delivery',{expectedRecord:record,sourceFingerprint:videoExcerptDeliveryFingerprint(approved)});if(!claim.claimed)continue;
  let response;
  try{
   const fresh=(await loadApproved()).find(r=>r.recordId===record.recordId);
   if(!fresh||videoExcerptDeliveryFingerprint(fresh)!==record.delivery.sourceFingerprint)throw Error('Approval changed after the delivery claim; reconcile before sending');
   summary.sends++;response=await send(fresh,parent);
  }catch(e){response={ok:false,error:e.message};}
  const saved=await db('finish_video_excerpt_delivery',{recordId:record.recordId,token:claim.token,acknowledged:response?.ok===true&&Boolean(response.poetryPleaseItemId),canonicalExcerptId:response?.poetryPleaseItemId||'',error:response?.error||'Receiver outcome uncertain; no automatic resend'});
  if(!saved.saved){summary.held++;continue;}
  if(saved.record.delivery.state==='acknowledged'&&now()-started<budgetMs-25000){try{await inspect(saved.record);summary.received++;}catch(e){await touch(saved.record,'inspection_failed',e.message);summary.held++;}}
 }
 return summary;
}
