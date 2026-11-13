import test from 'node:test';import assert from 'node:assert/strict';
import {VIDEO_AUTO_POLICY,videoQualitySummary,localVideoPipelineState,applyVideoPipelineReceiver,publicVideoPipelineState,advanceVideoPipeline,videoPipelineKey} from '../video_pipeline.mjs';
import {verifyVideoPipelineCaller,VIDEO_PIPELINE_AUDIENCE,VIDEO_PIPELINE_SERVICE_ACCOUNT} from '../video_pipeline_auth.mjs';
const now=Date.parse('2026-11-13T16:00:00Z');
const c={prioritySetId:'event',sourceFileId:'fv',videoRecordId:'weaver:video:root',sourceEvent:'Event',author:'Poet',poemTitle:'Poem',baseScore:8.5,reviewCount:5,excerptRecordIds:['caption-source'],sourceMediaType:'FV',requiresManualIdentityReview:false};
const received={receiver:{received:true,sourceRecordId:c.videoRecordId,canonicalVideoId:'canonical',media:{durableCopyStatus:'verified',storageMode:'durable',durableByteCount:100,durableChecksum:'hash'},review:{reviewCount:5,selectedExcerptIdCount:1}}};
const ready={schemaVersion:'1',canonicalVideoId:'canonical',sourceRecordId:c.videoRecordId,checkedAt:new Date(now).toISOString(),media:{status:'verified',assetId:'asset',outputFileId:'output',mode:'durable',verifiedAt:new Date(now).toISOString()},caption:{status:'approved',canonicalExcerptId:'EXC',fingerprint:'text-fingerprint',approvedAt:new Date(now).toISOString()}};
test('five independent valid current reviews and 8.5 qualify; excerpts add no score',()=>{
 const reviews=Array.from({length:5},(_,i)=>({reviewerEmail:`r${i}@x`,rating:i<2?'moved_me':'like',updatedAt:'2026-11-13',excerptRecordIds:i===0?['one']:[]}));
 assert.deepEqual(videoQualitySummary(reviews),{reviewerCount:5,excerptCount:1,candidateScore:8.5,reviewComplete:true});
 assert.equal(videoQualitySummary([...reviews,{...reviews[0],reviewerEmail:'R0@X'}]).reviewerCount,5);
 assert.equal(videoQualitySummary([...reviews,{...reviews[0],rating:'dislike',updatedAt:'2026-11-14'}]).candidateScore,7);
 assert.equal(videoQualitySummary(reviews.map(r=>({...r,rating:'invalid'}))).reviewerCount,0);
 assert.equal(VIDEO_AUTO_POLICY.minScore,8.5);
});
test('qualification respects reviewer minimum, score, excerpt, editing, identity, restrictions and holds',()=>{
 assert.equal(localVideoPipelineState(c).stage,'ready');
 for(const [patch,stage]of [[{reviewCount:4},'awaiting_reviews'],[{baseScore:8.49},'below_threshold'],[{excerptRecordIds:[]},'needs_excerpt'],[{sourceMediaType:'OM'},'needs_editing'],[{requiresManualIdentityReview:true},'needs_identity'],[{publicationRestricted:true},'held'],[{gate:{decision:'hold'}},'held'],[{gate:{decision:'send_to_editing'}},'needs_editing']])assert.equal(localVideoPipelineState({...c,...patch}).stage,stage);
});
test('receiver acceptance, media completion and approved caption are separate positive gates',()=>{
 const state=localVideoPipelineState(c,{},now);
 assert.equal(applyVideoPipelineReceiver(state,c,{receiver:{...received.receiver,media:{durableCopyStatus:'queued'}}},now).stage,'media_pending');
 assert.equal(applyVideoPipelineReceiver(state,c,received,now).stage,'awaiting_pp_readiness');
 const final=applyVideoPipelineReceiver(state,c,{...received,readiness:ready},now);assert.equal(final.calendarReady,true);assert.equal(final.calendarEvidence.captionExcerptId,'EXC');
 for(const patch of [{schemaVersion:'2'},{canonicalVideoId:'wrong'},{sourceRecordId:'wrong'},{caption:{...ready.caption,status:'collected'}},{media:{...ready.media,status:'queued'}},{checkedAt:'2026-11-12'}])assert.equal(applyVideoPipelineReceiver(state,c,{...received,readiness:{...ready,...patch}},now).calendarReady,false);
 assert.equal(applyVideoPipelineReceiver(state,{...c,gate:{decision:'hold'}},{...received,readiness:ready},now).calendarReady,false);
 assert.equal(publicVideoPipelineState(final,now+16*60000).calendarReady,false);
 assert.throws(()=>applyVideoPipelineReceiver(state,c,{receiver:{...received.receiver,sourceRecordId:'wrong'}},now),/source/);
});
test('full progression sends once, then reconciles an existing canonical without replacing its media',async()=>{
 let sends=0;
 const first=await advanceVideoPipeline({candidates:[c],inspect:async()=>({receiver:{...received.receiver,media:{durableCopyStatus:'queued'}}}),deliver:async()=>{sends++;return{ok:true};},now:()=>now});
 assert.equal(first.states[0].stage,'media_pending');assert.equal(sends,1);
 const sent={...c,existingCanonicalVideoId:'canonical',gate:{poetryPleaseHandoff:{status:'sent_to_poetry_please'}}};
 const next=await advanceVideoPipeline({candidates:[sent],previous:first.states,inspect:async()=>({...received,readiness:ready}),deliver:async()=>{throw Error('Never replace an existing canonical');},now:()=>now});
 assert.equal(next.states[0].calendarReady,true);assert.equal(next.sends,0);
 const stale=localVideoPipelineState({...sent,gate:{decision:'hold',poetryPleaseHandoff:{status:'sent_to_poetry_please'}}},next.states[0],now);assert.equal(stale.calendarReady,false);
});
test('unknown send outcomes are rechecked, never blindly replayed; checks are bounded and rotate',async()=>{
 let sends=0;const first=await advanceVideoPipeline({candidates:[c],inspect:async()=>received,deliver:async()=>{sends++;throw Error('timeout');},now:()=>now});assert.equal(first.states[0].stage,'needs_attention');
 const uncertain={...c,gate:{poetryPleaseHandoff:{status:'needs_reconciliation'}}};
 const second=await advanceVideoPipeline({candidates:[uncertain],previous:first.states,inspect:async()=>received,deliver:async()=>{sends++;},now:()=>now});assert.equal(second.states[0].stage,'awaiting_pp_readiness');assert.equal(sends,1);
 const noReceipt=await advanceVideoPipeline({candidates:[uncertain],inspect:async()=>({receiver:{sourceRecordId:c.videoRecordId,received:false}}),deliver:async()=>{throw Error('No replay');},now:()=>now});assert.equal(noReceipt.states[0].stage,'needs_reconciliation');
 const variants=[0,1].map(i=>({...uncertain,sourceFileId:'variant'+i})),prior=variants.map((v,i)=>({...localVideoPipelineState(v,{},now),pipelineId:videoPipelineKey(v),receiverCheckedAt:new Date(now-i*60000).toISOString()}));let checked=[];
 await advanceVideoPipeline({candidates:variants,previous:prior,maxChecks:1,inspect:async v=>{checked.push(v.sourceFileId);return received;},deliver:async()=>{},now:()=>now});assert.deepEqual(checked,['variant1']);
});
test('scheduled worker accepts only the existing runtime identity with the fixed audience',async()=>{
 let config;const verify={verifyIdToken:async options=>{config=options;return{getPayload:()=>({email:VIDEO_PIPELINE_SERVICE_ACCOUNT,email_verified:true})};}};
 await verifyVideoPipelineCaller({headers:{authorization:'Bearer signed-token'}},verify);assert.equal(config.audience,VIDEO_PIPELINE_AUDIENCE);
 await assert.rejects(verifyVideoPipelineCaller({headers:{authorization:'Bearer signed-token'}},{verifyIdToken:async()=>({getPayload:()=>({email:'other@example.com',email_verified:true})})}),/denied/);
 await assert.rejects(verifyVideoPipelineCaller({headers:{}},verify),/required/);
});
