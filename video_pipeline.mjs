import {createHash} from 'node:crypto';
export const VIDEO_AUTO_POLICY=Object.freeze({version:'2026-11-13',minScore:8.5,minReviewers:5,minExcerpts:1});
const clean=v=>String(v??'').trim();
export function videoQualitySummary(reviews=[]){
  const current=new Map();
  for(const review of reviews){const email=clean(review.reviewerEmail).toLowerCase();if(!email)continue;
    const prior=current.get(email);if(!prior||clean(review.updatedAt)>=clean(prior.updatedAt))current.set(email,review);}
  const valid=[...current.values()].map(review=>{const legacy=clean(review.legacyScore);const score=review.ratingSource==='legacy_import'&&legacy!==''?Number(legacy):({dislike:2.5,meh:5,like:7.5,moved_me:10}[clean(review.rating).toLowerCase()]);return{review,score};}).filter(r=>Number.isFinite(r.score)&&r.score>=0&&r.score<=10);
  const excerpts=new Set(valid.flatMap(r=>r.review.excerptRecordIds||[]).map(clean).filter(Boolean));
  const score=valid.length?Number((valid.reduce((sum,r)=>sum+r.score,0)/valid.length).toFixed(2)):0;
  return{reviewerCount:valid.length,excerptCount:excerpts.size,candidateScore:score,
    reviewComplete:valid.length>=VIDEO_AUTO_POLICY.minReviewers&&excerpts.size>=VIDEO_AUTO_POLICY.minExcerpts};
}
export const videoPipelineKey=c=>createHash('sha256').update(`${c.prioritySetId}|${c.sourceFileId}`).digest('hex').slice(0,32);
const messages={awaiting_reviews:'Waiting for five independent reviews',below_threshold:'Below the 8.5 automatic-delivery threshold',needs_excerpt:'Waiting for one collected excerpt',held:'Held for an explicit decision or clearance',needs_editing:'Original media: waiting for a finished edit',needs_media_classification:'Check the source media role: OM or FV is not verified',needs_editing_decision:'Finished source conflicts with a saved editing decision; review that decision',needs_identity:'Waiting for a canonical identity check',ready:'Qualified for automatic Poetry Please delivery',receiver_check:'Checking the existing receiver record',needs_reconciliation:'Receiver outcome needs reconciliation; no blind resend',received:'Poetry Please receipt confirmed',media_pending:'Waiting for verified media',media_failed:'Receiver media verification needs attention',awaiting_pp_readiness:'Media verified; waiting for approved-caption readiness from Poetry Please',calendar_ready:'Ready for the Calendar consumer',needs_attention:'Pipeline check needs attention'};
export function localVideoPipelineState(candidate,prior={},now=Date.now()){
  const c=candidate,g=c.gate||{},handoff=g.poetryPleaseHandoff||{};
  let stage;
  if(c.publicationRestricted||c.releaseStatus||['hold','reject'].includes(g.decision))stage='held';
  else if(c.existingCanonicalVideoId||handoff.status)stage=prior.sourceRecordId===c.videoRecordId&&prior.receiverCheckedAt?prior.stage:'receiver_check';
  else if(c.reviewCount<VIDEO_AUTO_POLICY.minReviewers)stage='awaiting_reviews';
  else if(c.baseScore<VIDEO_AUTO_POLICY.minScore)stage='below_threshold';
  else if(!(c.excerptRecordIds||[]).length)stage='needs_excerpt';
  else if(!['OM','FV'].includes(c.sourceMediaType))stage='needs_media_classification';
  else if(g.decision==='send_to_editing')stage=c.sourceMediaType==='OM'?'needs_editing':'needs_editing_decision';
  else if(c.sourceMediaType==='OM'&&!g.publishableAssetUrl)stage='needs_editing';
  else if(c.requiresManualIdentityReview)stage='needs_identity';
  else stage='ready';
  return{...prior,schemaVersion:'1',pipelineId:videoPipelineKey(c),prioritySetId:c.prioritySetId,sourceFileId:c.sourceFileId,
    sourceRecordId:c.videoRecordId,sourceEvent:c.sourceEvent,sourceMediaType:c.sourceMediaType,sourceMediaTypeReason:c.sourceMediaTypeReason||'',canonicalVideoId:c.existingCanonicalVideoId||handoff.canonicalVideoId||'',
    author:c.author,title:c.poemTitle,policy:VIDEO_AUTO_POLICY,score:c.baseScore,reviewerCount:c.reviewCount,
    excerptCount:(c.excerptRecordIds||[]).length,stage,message:messages[stage]||messages.receiver_check,calendarReady:Boolean(stage==='calendar_ready'&&prior.calendarReady&&prior.sourceRecordId===c.videoRecordId),checkedAt:new Date(now).toISOString()};
}
export function applyVideoPipelineReceiver(state,candidate,inspection,now=Date.now()){
  const receiver=inspection.receiver;
  if(!receiver||receiver.sourceRecordId!==candidate.videoRecordId)throw Error('Receiver source does not match pipeline source');
  if(!receiver.received)return{...state,stage:'needs_reconciliation',message:messages.needs_reconciliation,calendarReady:false,receiverCheckedAt:new Date(now).toISOString()};
  if(candidate.existingCanonicalVideoId&&receiver.canonicalVideoId!==candidate.existingCanonicalVideoId)throw Error('Receiver canonical identity changed');
  const media=receiver.media||{},readiness=inspection.readiness;
  const contract=readiness?.schemaVersion==='1'&&readiness.canonicalVideoId===receiver.canonicalVideoId&&readiness.sourceRecordId===receiver.sourceRecordId;
  const mediaEvidence=contract&&readiness.media?.status==='verified'&&clean(readiness.media.assetId)&&clean(readiness.media.outputFileId)
    &&['durable','drive_link'].includes(readiness.media.mode)&&Number.isFinite(Date.parse(readiness.media.verifiedAt));
  const durable=media.durableCopyStatus==='verified'&&media.storageMode==='durable'&&Number(media.durableByteCount)>0&&clean(media.durableChecksum);
  const mediaReady=Boolean(durable||mediaEvidence);
  const captionReady=Boolean(contract&&readiness.caption?.status==='approved'&&clean(readiness.caption.canonicalExcerptId)&&clean(readiness.caption.fingerprint)&&Number.isFinite(Date.parse(readiness.caption.approvedAt)));
  // The optional, authenticated PP contract binds the chosen output and caption.
  // Current receipts without that contract cannot claim Calendar readiness.
  const calendarReady=Boolean(mediaEvidence&&captionReady&&Number.isFinite(Date.parse(readiness.checkedAt))&&now-Date.parse(readiness.checkedAt)>=0&&now-Date.parse(readiness.checkedAt)<=15*60000);
  let stage=calendarReady?'calendar_ready':mediaReady?'awaiting_pp_readiness':media.durableCopyStatus==='failed'?'media_failed':'media_pending';
  if(candidate.publicationRestricted||candidate.releaseStatus||['hold','reject','send_to_editing'].includes(candidate.gate?.decision)){stage='held';}
  const out={...state,stage,message:messages[stage]||messages.receiver_check,canonicalVideoId:receiver.canonicalVideoId,receiverSourceRecordId:receiver.sourceRecordId,
    receiverCheckedAt:new Date(now).toISOString(),mediaReady,captionReady,calendarReady:stage==='calendar_ready',
    receivedReviewCount:Number(receiver.review?.reviewCount||0),receivedExcerptLinkCount:Number(receiver.review?.selectedExcerptIdCount||0),
    mediaStatus:clean(media.durableCopyStatus||media.mediaStatus||'unknown')};
  if(calendarReady){out.calendarEvidence={schemaVersion:'1',canonicalVideoId:receiver.canonicalVideoId,sourceRecordId:receiver.sourceRecordId,
    mediaAssetId:readiness.media.assetId,outputFileId:readiness.media.outputFileId,captionExcerptId:readiness.caption.canonicalExcerptId,
    captionFingerprint:readiness.caption.fingerprint,checkedAt:readiness.checkedAt};}
  else delete out.calendarEvidence;
  return out;
}
export function publicVideoPipelineState(state,now=Date.now()){
  if(!state)return null;const fresh=Number.isFinite(Date.parse(state.receiverCheckedAt))&&now-Date.parse(state.receiverCheckedAt)>=0&&now-Date.parse(state.receiverCheckedAt)<=15*60000;
  return{schemaVersion:'1',pipelineId:state.pipelineId,prioritySetId:state.prioritySetId,sourceFileId:state.sourceFileId,sourceRecordId:state.sourceRecordId,
    canonicalVideoId:state.canonicalVideoId||'',sourceMediaType:state.sourceMediaType||'unknown',sourceMediaTypeReason:state.sourceMediaTypeReason||'',stage:state.stage,message:state.message,score:state.score,reviewerCount:state.reviewerCount,
    excerptCount:state.excerptCount,policy:VIDEO_AUTO_POLICY,checkedAt:state.checkedAt,receiverCheckedAt:state.receiverCheckedAt||'',
    mediaReady:Boolean(state.mediaReady&&fresh),captionReady:Boolean(state.captionReady&&fresh),calendarReady:Boolean(state.calendarReady&&fresh),
    readinessFresh:fresh,...(state.calendarReady&&fresh?{calendarEvidence:state.calendarEvidence}:{})};
}
export async function advanceVideoPipeline({candidates,previous=[],inspect,deliver,now=Date.now,maxChecks=6,budgetMs=210000}){
  const started=now(),byId=new Map(previous.map(s=>[s.pipelineId,s]));
  const states=candidates.map(c=>localVideoPipelineState(c,byId.get(videoPipelineKey(c)),now()));
  const work=candidates.map((candidate,i)=>({candidate,state:states[i]})).filter(x=>x.state.stage==='ready'||x.candidate.existingCanonicalVideoId||x.candidate.gate?.poetryPleaseHandoff?.status)
    .sort((a,b)=>(a.state.stage==='ready'?0:1)-(b.state.stage==='ready'?0:1)||Date.parse(a.state.receiverCheckedAt||'1970-01-01')-Date.parse(b.state.receiverCheckedAt||'1970-01-01'));
  let checked=0,sends=0;
  for(const {candidate,state}of work){if(checked>=maxChecks||now()-started>budgetMs-30000)break;
    try{
      if(state.stage==='ready'){
        if(sends||now()-started>budgetMs-145000)continue;
        // Existing durable claims and receiver preflight own the actual send.
        sends++;const result=await deliver(candidate);if(!result){Object.assign(state,{stage:'needs_attention',message:'Qualification changed; waiting for the next check'});continue;}if(result?.requiresManualIdentityReview){Object.assign(state,{stage:'needs_identity',message:messages.needs_identity});continue;}
        if(result?.status==='needs_media_classification'){Object.assign(state,{stage:'needs_media_classification',message:messages.needs_media_classification});continue;}
        if(result?.status==='send_to_editing'){Object.assign(state,{stage:'needs_editing',message:messages.needs_editing});continue;}
        if(result?.ok===false&&!result?.canonicalVideoId){Object.assign(state,{stage:'needs_reconciliation',message:messages.needs_reconciliation});continue;}
      }
      checked++;const inspection=await inspect(candidate);Object.assign(state,applyVideoPipelineReceiver(state,candidate,inspection,now()));
    }catch(error){Object.assign(state,{stage:'needs_attention',message:messages.needs_attention,error:clean(error.message).slice(0,500),calendarReady:false,receiverCheckedAt:new Date(now()).toISOString()});delete state.calendarEvidence;}
  }
  return{ok:true,policy:VIDEO_AUTO_POLICY,checked,sends,states};
}
