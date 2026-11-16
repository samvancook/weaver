import test from 'node:test';import assert from 'node:assert/strict';
import {groupVideoVersionFiles,videoPerformanceContext,VIDEO_VERSION_FAMILIES,candidateMatchesVideoSource} from '../video_version_families.mjs';
import {reconcileVideoReviews,videoReviewProgress,buildVideoReviewEvidence,buildVideoContextUpdate} from '../video_curation_handoff.mjs';
const f=VIDEO_VERSION_FAMILIES[0],files=f.fileIds.slice(0,3).map((id,i)=>({id,name:['Click.mov','Click1080.mp4','ClickVV.mp4'][i]}));
const records=[{reviewId:'old',sourceFileId:f.reviewSourceFileId,sourceFileName:'Click.mov',reviewerEmail:'a@bp.com',rating:'like',excerptRecordIds:['old-excerpt'],updatedAt:'2026-09-01'},
 {reviewId:'variant',sourceFileId:files[1].id,reviewerEmail:'A@bp.com',rating:'like',excerptRecordIds:['new-excerpt'],updatedAt:'2026-11-03'},
 {reviewId:'other',sourceFileId:files[2].id,reviewerEmail:'b@bp.com',rating:'moved_me',excerptRecordIds:['another-excerpt'],updatedAt:'2026-11-03'}];
test('format variants become one performance, retain every physical file and deduplicate people without losing excerpts',()=>{
 const grouped=groupVideoVersionFiles(files,f.prioritySetId);assert.equal(grouped.length,1);assert.equal(grouped[0].videoVariants.length,3);
 const reviews=reconcileVideoReviews(grouped,records);assert.equal(reviews.length,2);
 assert.deepEqual(videoReviewProgress(reviews),{reviewerCount:2,excerptCount:3,candidateScore:8.75,reviewComplete:false});
 const context=videoPerformanceContext(grouped[0],reviews.flatMap(r=>r.excerptRecordIds));
 assert.equal(context.assets.length,3);assert.equal(context.excerptRecordIds.length,3);
 assert.equal(buildVideoReviewEvidence(grouped,records,files[0].id).length,3);
 assert.deepEqual(groupVideoVersionFiles(files,'another-event'),files);
});
test('shared context targets existing receiver parent and preserves every exact review source',()=>{
 const grouped=groupVideoVersionFiles(files,f.prioritySetId),reviews=reconcileVideoReviews(grouped,records);
 const c={prioritySetId:f.prioritySetId,sourceFileId:f.sourceFileId,sourceReviewFileId:f.reviewSourceFileId,versionSourceFileIds:f.fileIds,
 videoRecordId:'weaver:video:'+f.reviewSourceFileId,existingCanonicalVideoId:'existing',sourceEvent:'BPL',
 performanceContext:videoPerformanceContext(grouped[0],reviews.flatMap(r=>r.excerptRecordIds)),
 reviewEvidence:buildVideoReviewEvidence(grouped,records,files[0].id),ratings:reviews,
 excerptRecordIds:['old-excerpt','new-excerpt','another-excerpt'],gate:{decision:'ready_for_poetry_please',selectedExcerptRecordIds:['old-excerpt']}};
 const payload=buildVideoContextUpdate(c,{ok:true,received:true,sourceRecordId:c.videoRecordId,canonicalVideoId:'existing',event:{sourceEvent:'BPL'}});
 assert.equal(payload.reviews.length,3);assert.equal(payload.selectedExcerptRecordIds.length,3);
 assert.equal(payload.performanceContext.assets.length,3);
 assert.equal(payload.reviews.find(r=>r.reviewId==='variant').sourceRecordId,'weaver:video:'+files[1].id);
 assert.equal(candidateMatchesVideoSource(c,files[2].id),true);assert.equal(candidateMatchesVideoSource(c,'unrelated'),false);
 assert.ok(!('finalAssetUrl' in payload));assert.ok(!('decision' in payload));
});
test('no filename-only merge, no arbitrary primary replacement, conflicting explicit families fail',()=>{
 assert.equal(groupVideoVersionFiles([{id:'unrelated',name:'Click.mov'},...files],f.prioritySetId).length,2);
 assert.throws(()=>groupVideoVersionFiles(files.slice(1),f.prioritySetId),/representative/);
 assert.throws(()=>groupVideoVersionFiles(files,f.prioritySetId,[f,{...f,reviewSourceFileId:'different'}]),/Conflicting/);
});
test('the grouping contract supports horizontal, vertical, high and low resolution without format-based identity rules',()=>{
 const ids=['horizontal4k','horizontal720','vertical4k','vertical720'];
 const versions=ids.map(id=>({id,name:id+'.mp4'}));const family={...f,sourceFileId:ids[0],fileIds:ids,reviewSourceFileId:ids[0]};
 const grouped=groupVideoVersionFiles(versions,'bpl-charm-city-2026',[family]);
 assert.equal(grouped.length,1);const context=videoPerformanceContext(grouped[0],['shared']);
 assert.equal(context.assets.length,4);assert.deepEqual(context.excerptRecordIds,['shared']);
 assert.ok(context.assets.every(a=>a.relationship==='same_performance'&&a.timing==='source_specific'));
});
