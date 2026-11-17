import test from 'node:test';import assert from 'node:assert/strict';
import {isVerifiedPlatformRendition,preferredVerifiedPlatformRendition,selectedPlatformRenditionVerified} from '../video_rendition_selection.mjs';
import {readVideoRenditions} from '../video_curation_handoff.mjs';
const c={sourceFileId:'master',sourceMediaType:'FV',videoRecordId:'weaver:video:root',versionSourceFileIds:['master','related-mp4']};
const asset={id:'registered',contentCollection:'videos',contentDocId:'canonical',sourceRecordId:c.videoRecordId,sourceVersionFileId:'related-mp4',profileId:'platform-720p-h264-aac',profileVersion:'1',status:'active',storageMode:'drive_reference',validationState:'verified_bytes',outputFileId:'related-mp4',byteCount:123,verifiedByteCount:123,sha256:'a'.repeat(64),verifiedSha256:'a'.repeat(64)};
const list={ok:true,schemaVersion:'1',canonicalVideoId:'canonical',sourceRecordId:c.videoRecordId,renditions:[asset],selected:null};
test('720p on an explicitly related file is verified without substituting an unrelated recording',()=>{
 assert.equal(isVerifiedPlatformRendition(c,'canonical',c.videoRecordId,asset),true);
 for(const patch of [{sourceVersionFileId:'unrelated'},{sourceRecordId:'other'},{contentDocId:'other'},{validationState:'source_reported'},{verifiedSha256:'b'.repeat(64)},{profileVersion:'2'}])assert.equal(isVerifiedPlatformRendition(c,'canonical',c.videoRecordId,{...asset,...patch}),false);
 assert.equal(selectedPlatformRenditionVerified(c,{...list,selectedRendition:asset}),true);
 assert.equal(selectedPlatformRenditionVerified(c,{...list,selectedRendition:{...asset,outputFileId:'wrong'}}),false);
});
test('missing 1080p selector uses the exact registered 720p tuple after source/canonical verification',async()=>{
 const paths=[];
 const r=await readVideoRenditions(c,'canonical',async p=>{paths.push(p);if(paths.length===1)throw Object.assign(Error('missing'),{status:404,code:'video_rendition_not_found'});return paths.length===2?list:{...list,selected:asset};});
 assert.equal(paths.length,3);const q=new URLSearchParams(paths[2].split('?')[1]);assert.equal(q.get('profileId'),'platform-720p-h264-aac');assert.equal(q.get('sourceVersionFileId'),'related-mp4');assert.equal(r.selected.id,'registered');
});
test('ambiguous equally preferred outputs stay unresolved; preferred verified profiles are deterministic',()=>{
 assert.equal(preferredVerifiedPlatformRendition(c,'canonical',c.videoRecordId,[asset,{...asset,id:'second',outputFileId:'second'}]).status,'ambiguous_verified_outputs');
 const high={...asset,id:'high',profileId:'platform-1080p-h264-aac'};assert.equal(preferredVerifiedPlatformRendition(c,'canonical',c.videoRecordId,[asset,high]).asset.id,'high');
 assert.equal(preferredVerifiedPlatformRendition(c,'canonical',c.videoRecordId,[{...asset,validationState:'source_reported'}]).asset,null);
});
test('crossed canonical fallback cannot be selected even if rendition metadata looks verified',async()=>{
 let n=0;await assert.rejects(readVideoRenditions(c,'canonical',async()=>{if(!n++)throw Object.assign(Error('missing'),{status:404,code:'video_rendition_not_found'});return{...list,canonicalVideoId:'wrong'};}),/different video or source/);
});
