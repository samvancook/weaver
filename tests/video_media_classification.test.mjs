import test from 'node:test';import assert from 'node:assert/strict';
import {traceVideoFolderMediaType,videoFileMediaType} from '../video_media_classification.mjs';
import {localVideoPipelineState} from '../video_pipeline.mjs';
const folders={camera:{id:'camera',name:'Camera Y',parents:['event']},event:{id:'event',name:'Slam',parents:['om']},om:{id:'om',name:'Button Footage - Original Media (OM)',parents:['root']}};
test('an inaccessible ancestor does not erase the explicit role already verified',async()=>{
 const read=async id=>{if(!folders[id])throw Error('not readable');return folders[id];};
 const result=await traceVideoFolderMediaType('camera',read);
 assert.equal(result.sourceMediaType,'OM');assert.equal(result.sourceMediaTypeFolderId,'om');assert.equal(result.sourceMediaTypeTraceStatus,'ancestor_unavailable');assert.equal(result.sourceMediaTypeTraceStoppedAt,'root');
 assert.equal((await traceVideoFolderMediaType('root',read)).sourceMediaTypeReason,'ancestor_unavailable');
});
test('missing labels, conflicting roles, multiple parents and cycles remain blocked with distinct reasons',async()=>{
 for(const [nodes,reason]of [[{a:{id:'a',name:'Camera Y'}},'no_explicit_folder_role'],[{a:{id:'a',name:'OM',parents:['b']},b:{id:'b',name:'FV'}},'conflicting_folder_roles'],[{a:{id:'a',name:'OM',parents:['b','c']}},'multiple_parents'],[{a:{id:'a',name:'OM',parents:['a']}},'folder_cycle']]){
  const r=await traceVideoFolderMediaType('a',async id=>nodes[id]);assert.equal(r.sourceMediaType,'unknown');assert.equal(r.sourceMediaTypeReason,reason);
 }
});
test('mixed inventory events classify exact files independently; unknown never inherits a sibling role',()=>{
 const set={inventoryEvent:true,sourceMediaType:'FV',inventoryTitles:{a:{assets:[{driveFileId:'a',mediaType:'OM'}]},b:{assets:[{driveFileId:'b',mediaType:'FV'}]},c:{assets:[{driveFileId:'other',mediaType:'FV'}]}}};
 assert.deepEqual(['a','b','c'].map(id=>videoFileMediaType(set,{id},{}).sourceMediaType),['OM','FV','unknown']);
 assert.equal(videoFileMediaType({}, {id:'a'},{sourceMediaType:'OM'}).sourceMediaType,'OM');
});
test('an old automatic editing gate cannot turn an unclassified source into an editing instruction',()=>{
 const c={prioritySetId:'e',sourceFileId:'f',reviewCount:5,baseScore:9,excerptRecordIds:['x'],sourceMediaType:'unknown',gate:{decision:'send_to_editing',decidedBy:'weaver:auto'}};
 assert.equal(localVideoPipelineState(c).stage,'needs_media_classification');
 assert.equal(localVideoPipelineState({...c,sourceMediaType:'OM'}).stage,'needs_editing');
 assert.equal(localVideoPipelineState({...c,sourceMediaType:'FV'}).stage,'needs_editing_decision');
 assert.equal(localVideoPipelineState({...c,gate:{decision:'hold'}}).stage,'held');
});

test('verified administrative ancestry resolves hidden parents only for the exact still-matching folder',async()=>{
 const {recordedFolderMediaType}=await import('../video_media_classification.mjs');
 const observed=await traceVideoFolderMediaType('camera',async()=>({id:'camera',name:'Camera Y'}));
 const record={schemaVersion:'1',status:'confirmed',verificationSource:'drive_metadata',sourceFolderId:'camera',mediaType:'OM',verifiedBy:'admin@example.com',verifiedAt:new Date().toISOString(),folderPath:[{id:'camera',name:'Camera Y',parentId:'om'},{id:'om',name:'Original Media (OM)'}]};
 assert.equal(recordedFolderMediaType(observed,record,'camera').sourceMediaType,'OM');
 for(const patch of [{status:'revoked'},{sourceFolderId:'other'},{mediaType:'FV'},{verifiedBy:''},{folderPath:[{id:'camera',name:'Changed',parentId:'om'},record.folderPath[1]]},{folderPath:[{...record.folderPath[0],parentId:'wrong'},record.folderPath[1]]}])assert.equal(recordedFolderMediaType(observed,{...record,...patch},'camera').sourceMediaType,'unknown');
 assert.equal(recordedFolderMediaType({...observed,sourceMediaTypeReason:'multiple_parents'},record,'camera').sourceMediaType,'unknown');
 assert.equal(recordedFolderMediaType({...observed,sourceMediaType:'FV'},record,'camera').sourceMediaType,'FV');
});
