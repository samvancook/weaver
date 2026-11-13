import {makeVideoAssetLinks,planVideoMediaResolution,withMediaResolution,normalizeVideoIdentity,videoAssetIdentity,assertResolvedVideoMedia,assertSelectedAssetVersion} from './video_resolution.mjs';
const clean=v=>String(v??'').trim();
const fileId=url=>clean(url).match(/\/file\/d\/([\w-]+)/)?.[1]||'';
export function createVideoResolutionService(deps){
 const {db,getFile,fileRole,folderRole,listEntries,getSet,getCandidates,enqueue}=deps;
 const read=async()=>{const r=await db('get_video_media_resolutions',{});if(!r?.ok)throw Error('Durable media resolution ledger unavailable');return r.records||[];};
 const recordFor=(rows,c)=>rows.find(r=>r.type==='recording'&&r.prioritySetId===c.prioritySetId&&r.sourceFileId===c.sourceFileId)||rows.find(r=>r.type==='recording'&&r.prioritySetId===c.prioritySetId&&r.pairing?.status==='verified'&&(c.sourceReviewFileId&&r.reviewedSourceFileId===c.sourceReviewFileId||(r.assets||[]).some(a=>a.fileId===c.sourceFileId&&a.pairingVerified)));
 const scopeFor=(rows,c)=>rows.find(r=>r.type==='event_scope'&&r.prioritySetId===c.prioritySetId);
 const asset=(f,role,evidence)=>{if(f.trashed||!String(f.mimeType).startsWith('video/')||Number(f.size)<=0)throw Error('Linked media is not an available video');return {fileId:f.id,name:f.name,url:`https://drive.google.com/file/d/${f.id}/view`,mediaType:role,mediaTypeEvidence:evidence,
  width:Number(f.videoMediaMetadata?.width||0),height:Number(f.videoMediaMetadata?.height||0),durationMs:Number(f.videoMediaMetadata?.durationMillis||0),byteCount:Number(f.size),checksum:f.md5Checksum||''};};
 const liveAsset=async(id,role)=>{const f=await getFile(id);const type=role||await fileRole(`https://drive.google.com/file/d/${id}/view`);return asset(f,type,'current_drive_metadata');};
 const persist=async(record,prior,actor,event)=>{const r=await db('save_video_media_resolution',{record,expectedRevision:prior?.revision||0,actor,event});if(!r?.ok)throw Error(r.reason||'Resolution changed; reload before saving');return r.record;};
 async function attach(candidates){const rows=await read();return candidates.map(c=>{const r=recordFor(rows,c),root=r?.reviewedSourceFileId||c.sourceReviewFileId;
  const received=candidates.find(v=>v.prioritySetId===c.prioritySetId&&(v.sourceReviewFileId||v.sourceFileId)===root&&v.gate?.poetryPleaseHandoff?.status==='sent_to_poetry_please'&&v.gate.poetryPleaseHandoff.sourceRecordId==='weaver:video:'+root&&v.existingCanonicalVideoId);
  const linked=root&&!c.gate?.poetryPleaseHandoff?.status?{...c,sourceReviewFileId:root,videoRecordId:'weaver:video:'+root,existingCanonicalVideoId:received?.existingCanonicalVideoId||c.existingCanonicalVideoId}:c;
  return withMediaResolution({...linked,mediaResolutionRequired:true,sourceFolderId:getSet(c.prioritySetId)?.folderId||c.sourceFolderId},r);});}
 async function saveLinks(payload,actor){
  const result=await getCandidates(payload.prioritySetId),c=result.candidates.find(c=>c.sourceFileId===payload.sourceFileId);if(!c)throw Error('Exact reviewed candidate unavailable');
  const rows=await read(),prior=recordFor(rows,c);if(prior&&prior.sourceFileId!==c.sourceFileId)throw Error('Save links on the original reviewed recording, not its alternate review file');if((prior?.revision||0)!==Number(payload.expectedRevision||0))throw Error('Resolution changed; reload before saving');
  const urls=[...new Set((payload.urls||[]).map(clean).filter(Boolean))];if(!urls.length||urls.length>12)throw Error('Link between one and twelve exact Drive videos');
  const ids=urls.map(fileId);if(payload.primaryFileId&&!ids.includes(payload.primaryFileId))throw Error('Include the selected primary among the linked URLs');if(ids.some(id=>!id))throw Error('Use exact Drive file links');
  const assets=await Promise.all(ids.map(id=>liveAsset(id)));
  for(const asset of assets)if(asset.fileId===payload.primaryFileId&&asset.mediaType==='unknown'&&payload.confirmPairing===true&&clean(payload.evidence).length>=10){asset.mediaType='FV';asset.mediaTypeEvidence='administrator_confirmed_finished_output';}
  const record=makeVideoAssetLinks(c,prior,payload,actor,assets);
  const saved=await persist(record,prior,actor,'asset_links_saved_without_release_change');await enqueue(c.prioritySetId,c.sourceFileId,'media_links_saved');return {ok:true,record:saved,releaseDecisionChanged:false};
 }
 async function saveScope(payload,actor){
  const set=getSet(payload.prioritySetId);if(!set)throw Error('Choose one existing event');
  if(payload.confirmedCoverage!==true)throw Error('Confirm the event’s complete finished-edit search scope');
  const ids=[...new Set(payload.finishedFolderIds||[])];if(ids.length>8||ids.some(id=>!/^[-\w]{15,}$/.test(id)))throw Error('Use up to eight exact Drive folder IDs');
  for(const id of ids){const [f,r]=await Promise.all([getFile(id),folderRole(id)]);if(f.mimeType!=='application/vnd.google-apps.folder'||r.sourceMediaType!=='FV')throw Error('Every search folder must be verified FV');}
  const rows=await read(),prior=scopeFor(rows,{prioritySetId:set.id});if((prior?.revision||0)!==Number(payload.expectedRevision||0))throw Error('Search scope changed; reload');
  const record={type:'event_scope',schemaVersion:'1',prioritySetId:set.id,sourceFolderId:set.folderId,finishedFolderIds:ids,confirmedCoverage:true,verifiedBy:actor,verifiedAt:new Date().toISOString(),evidence:clean(payload.evidence)};
  if(record.evidence.length<10)throw Error('Record how the event search scope was confirmed');
  const saved=await persist(record,prior,actor,'event_finished_scope_confirmed');return {ok:true,record:saved};
 }
 async function lookup(scope,deadline=Date.now()+40000){
  const queue=[...(scope.finishedFolderIds||[])],seen=new Set(),files=[];
  while(queue.length){if(Date.now()>deadline)throw Error('Finished lookup time budget exhausted; no editing request created');if(seen.size>=32||files.length>1000)throw Error('Event lookup exceeded bounds; split the search scope');const id=queue.shift();if(seen.has(id))throw Error('Repeated or cyclic finished-folder scope');seen.add(id);
   const role=await folderRole(id);if(role.sourceMediaType!=='FV')throw Error('A finished search folder is unclassified or conflicts with FV');
   for(const f of await listEntries(id)){if(f.mimeType==='application/vnd.google-apps.folder')queue.push(f.id);else if(String(f.mimeType).startsWith('video/'))files.push(asset(f,'FV','verified_event_finished_folder'));}
  }
  return files;
 }
 async function resolve(c,allCandidates,rows,deadline){
  const prior=recordFor(rows,c),scope=scopeFor(rows,c);let selected=null,source=null,files=[],complete=false,error='';
  try{
   if(prior?.primaryFileId&&prior.pairing?.status==='verified'){selected=await liveAsset(prior.primaryFileId);const saved=prior.assets?.find(a=>a.fileId===prior.primaryFileId);selected={...saved,...selected,pairingVerified:saved?.pairingVerified,classificationEvidence:saved?.classificationEvidence};if(selected.mediaType==='unknown'&&['administrator_confirmed_finished_output','authorized_drive_ancestry','verified_event_finished_folder','current_drive_metadata','verified_review_source'].includes(saved?.mediaTypeEvidence)){selected.mediaType='FV';selected.mediaTypeEvidence=saved.mediaTypeEvidence;}}
   source=await liveAsset(c.sourceFileId);
   if(source.mediaType==='unknown'&&['OM','FV'].includes(c.sourceMediaType))source={...source,mediaType:c.sourceMediaType,mediaTypeEvidence:'verified_review_source'};
   if(source.mediaType==='unknown'&&selected?.fileId===source.fileId&&selected.mediaType==='FV')source={...source,mediaType:'FV'};
   if(source.mediaType==='unknown'&&prior?.sourceRoleReview?.method==='administrator_confirmation')source={...source,mediaType:prior.sourceRole};
   if(!selected&&c.sourceMediaType==='OM'&&scope?.confirmedCoverage){files=await lookup(scope,deadline);complete=true;}
  }catch(e){error=String(e.message).slice(0,400);}
  const identity=v=>normalizeVideoIdentity(v.author)+'|'+normalizeVideoIdentity(v.poemTitle);
  const sourceCount=new Set(allCandidates.filter(v=>v.prioritySetId===c.prioritySetId&&identity(v)===identity(c)).map(v=>v.sourceReviewFileId||v.sourceFileId)).size;
  const planned=planVideoMediaResolution(c,prior,scope,{files,complete,sourceCount,sourceAsset:source,selectedAsset:selected,error});
  const saved=await persist(planned,prior,'weaver:scheduled-resolution','media_resolution_checked');
  const index=rows.findIndex(r=>r.type==='recording'&&r.prioritySetId===c.prioritySetId&&r.sourceFileId===c.sourceFileId);if(index>=0)rows[index]=saved;else rows.push(saved);
  let updated=withMediaResolution(c,saved);
  if(saved.status==='resolved'){
   const prepared=await db('prepare_resolved_video_gate',{candidate:updated,expectedGate:c.gate||null});
   if(prepared.ok)updated={...updated,gate:prepared.record,sourceReviewFileId:prepared.record.sourceReviewFileId||c.sourceReviewFileId,videoRecordId:prepared.record.videoRecordId,requiresManualIdentityReview:!prepared.record.sourceReviewFileId&&c.requiresManualIdentityReview};
   else updated.mediaResolution={...updated.mediaResolution,status:'needs_release_check',reason:prepared.reason};
  }
  return updated;
 }
 async function prepare(candidates,{budgetMs=55000,max=3}={}){
  const started=Date.now(),rows=await read();let checked=0;
  const work=candidates.filter(c=>!c.existingCanonicalVideoId&&!c.gate?.poetryPleaseHandoff?.status&&!c.publicationRestricted&&!c.releaseStatus&&!['hold','reject'].includes(c.gate?.decision)&&c.reviewCount>=5&&c.baseScore>=8.5&&c.excerptRecordIds?.length)
   .sort((a,b)=>Date.parse(recordFor(rows,a)?.checkedAt||'1970-01-01')-Date.parse(recordFor(rows,b)?.checkedAt||'1970-01-01')||a.sourceFileId.localeCompare(b.sourceFileId));
  const out=[...candidates],visitedRoots=new Set();
  work.sort((a,b)=>Number(recordFor(rows,b)?.sourceFileId===b.sourceFileId)-Number(recordFor(rows,a)?.sourceFileId===a.sourceFileId));
  for(const c of work){const root=c.prioritySetId+'|'+(c.sourceReviewFileId||c.sourceFileId);if(visitedRoots.has(root))continue;visitedRoots.add(root);if(checked>=max||Date.now()-started>budgetMs-20000)break;const prior=recordFor(rows,c);if(prior&&Date.now()-Date.parse(prior.checkedAt)<5*60000&&c.gate?.decision==='ready_for_poetry_please'&&prior.status==='resolved')continue;
   const updated=await resolve(c,candidates,rows,started+budgetMs-5000);out[out.findIndex(v=>v.prioritySetId===c.prioritySetId&&v.sourceFileId===c.sourceFileId)]=updated;checked++;
  }
  return {candidates:out,checked};
 }
 async function verifySelected(candidate){const asset=assertResolvedVideoMedia(candidate);return assertSelectedAssetVersion(asset,await getFile(asset.fileId));}
 async function previewScopes(){const scopes=(await read()).filter(r=>r.type==='event_scope').sort((a,b)=>a.prioritySetId.localeCompare(b.prioritySetId));const deadline=Date.now()+45000,checks=[];for(const scope of scopes.slice(0,3)){try{checks.push({prioritySetId:scope.prioritySetId,ok:true,fileCount:(await lookup(scope,deadline)).length});}catch(e){checks.push({prioritySetId:scope.prioritySetId,ok:false,error:e.message});}}return {checks,uncheckedScopeCount:Math.max(0,scopes.length-3)};}
 return {read,attach,saveLinks,saveScope,prepare,lookup,verifySelected,previewScopes};
}
