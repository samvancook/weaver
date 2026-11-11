import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
const source=readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
const helper=source.match(/function registeredExcerptVideoSource\([\s\S]*?\n}\n/)[0];
const resolve=runInNewContext(`(${helper})`,{cleanSheetWhitespace:v=>String(v??'').trim()});
const record={sourceRecordId:'approved-excerpt',sourceVideoFileId:'alternate-file',sourceEvent:'Event'};
const candidate={sourceFileId:'alternate-file',sourceReviewFileId:'reviewed-root',existingCanonicalVideoId:'canonical',videoRecordId:'weaver:video:reviewed-root',sourceEvent:'Event',gate:{decision:'ready_for_poetry_please'}};
const asset={id:'registered-asset',assetType:'video_rendition',contentCollection:'videos',contentDocId:'canonical',sourceRecordId:candidate.videoRecordId,reviewedSourceFileId:'reviewed-root',outputFileId:'alternate-file',status:'active',role:'master',validationState:'source_reported'};
const receipt={received:true,canonicalVideoId:'canonical',sourceRecordId:candidate.videoRecordId,event:{sourceEvent:'Event'},renditions:[asset]};
test('an explicitly registered alternate master resolves to the existing received canonical only',()=>{
 const result=resolve(record,candidate,receipt);assert.equal(result.canonicalVideoId,'canonical');assert.equal(result.evidence.renditionAssetId,'registered-asset');assert.equal(result.evidence.reviewedSourceFileId,'reviewed-root');assert.equal(record.sourceVideoFileId,'alternate-file');
});
test('a matching title, wrong root, event, canonical, file or unapproved candidate cannot resolve',()=>{
 for(const patch of [{sourceReviewFileId:'wrong'},{sourceFileId:'wrong'},{existingCanonicalVideoId:'wrong'},{sourceEvent:'wrong'},{videoRecordId:'wrong'},{gate:{decision:'needs_work'}}])assert.equal(resolve(record,{...candidate,...patch},receipt),null);
 for(const patch of [{received:false},{sourceRecordId:'wrong'},{canonicalVideoId:'wrong'},{event:{sourceEvent:'wrong'}}])assert.equal(resolve(record,candidate,{...receipt,...patch}),null);
});
test('registry identity must match every link and be unique and active',()=>{
 for(const patch of [{id:''},{assetType:'library_video'},{contentCollection:'other'},{contentDocId:'wrong'},{sourceRecordId:'wrong'},{reviewedSourceFileId:'wrong'},{outputFileId:'wrong'},{status:'archived'},{validationState:'failed'}])assert.equal(resolve(record,candidate,{...receipt,renditions:[{...asset,...patch}]}),null);
 assert.equal(resolve(record,candidate,{...receipt,renditions:[asset,{...asset,id:'ambiguous'}]}),null);
 assert.equal(resolve(record,candidate,{...receipt,renditions:[]}),null);
});
test('platform versions need verified bytes; source-reported evidence is restricted to registered masters',()=>{
 assert.equal(resolve(record,candidate,{...receipt,renditions:[{...asset,role:'platform'}]}),null);
 assert.equal(resolve(record,candidate,{...receipt,renditions:[{...asset,role:'platform',validationState:'verified_bytes'}]}).canonicalVideoId,'canonical');
});
