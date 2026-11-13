import {createHash} from 'node:crypto';
const clean=v=>String(v??'').trim();
export const RESOLUTION_POLICY='exact-event-author-title-v1';
export const resolutionKey=c=>createHash('sha256').update(`${c.prioritySetId}|${c.sourceFileId}`).digest('hex').slice(0,32);
export const normalizeVideoIdentity=v=>clean(v).normalize('NFKC').replace(/[’‘]/g,"'").replace(/\s+/g,' ').toLowerCase();
export function videoAssetIdentity(name){
 const stem=clean(name).replace(/^\[Vertical Version(?: [AB])?\]\s*/i,'').replace(/^MVI_\d+\.MP4\s*/i,'').replace(/\.(mov|mp4|m4v|mkv)$/i,'');
 const split=stem.indexOf(' - ');return split<1?null:{author:normalizeVideoIdentity(stem.slice(0,split)),title:normalizeVideoIdentity(stem.slice(split+3))};
}
export function mediaResolutionPublic(record,administrative=false){
 if(!record)return null;return {revision:record.revision||0,status:record.status,reason:record.reason||'',sourceRole:record.sourceRole,
  primaryFileId:record.primaryFileId||'',primaryUrl:record.assets?.find(a=>a.fileId===record.primaryFileId)?.url||'',
  pairing:record.pairing?(administrative?record.pairing:{status:record.pairing.status,method:record.pairing.method,verifiedAt:record.pairing.verifiedAt,policy:record.pairing.policy}):null,assets:record.assets||[],lookup:record.lookup||null,editingRequest:record.editingRequest||null,checkedAt:record.checkedAt||'',updatedAt:record.updatedAt||''};
}
export function withMediaResolution(candidate,record){
 if(!record||record.prioritySetId!==candidate.prioritySetId||(record.sourceFileId!==candidate.sourceFileId&&record.reviewedSourceFileId!==candidate.sourceReviewFileId))return candidate;
 const primary=record.assets?.find(a=>a.fileId===record.primaryFileId);
 return {...candidate,mediaResolution:mediaResolutionPublic(record,true),observedSourceRole:candidate.sourceMediaType,sourceMediaType:candidate.sourceMediaType!=='unknown'?candidate.sourceMediaType:record.sourceRole||'unknown',
  knownAssets:[...(candidate.knownAssets||[]),...(record.assets||[]).filter(a=>!(candidate.knownAssets||[]).some(k=>k.fileId===a.fileId)).map(a=>({...a,roles:[a.mediaType==='FV'?'finished_version':'review_source']}))],
  ...(record.status==='resolved'&&primary?{resolvedAssetUrl:primary.url}: {})};
}
export function makeVideoAssetLinks(candidate,prior,payload,actor,assets,now=new Date().toISOString()){
 if(!actor)throw Error('Verified administrator is required');
 if(candidate.gate?.poetryPleaseHandoff?.status||candidate.existingCanonicalVideoId)throw Error('Existing receiver media requires rendition registration, not a primary replacement');
 const merged=new Map((prior?.assets||[]).map(a=>[a.fileId,{...a}]));for(const asset of assets)merged.set(asset.fileId,{...asset});
 const primary=clean(payload.primaryFileId);if(primary&&!merged.has(primary))throw Error('Primary must be one of the linked exact files');
 if(primary&&merged.get(primary).mediaType!=='FV')throw Error('Primary media role must be verified FV');
 for(const asset of merged.values())if(asset.fileId===primary&&payload.confirmPairing===true)asset.pairingVerified=true;
 const role=clean(payload.sourceRole)||prior?.sourceRole||candidate.sourceMediaType;
 if(!['OM','FV','unknown'].includes(role))throw Error('Invalid source role');
 if(candidate.sourceMediaType!=='unknown'&&role!==candidate.sourceMediaType)throw Error('Source role conflicts with current Drive evidence');
 if(candidate.sourceMediaType==='unknown'&&role!=='unknown'&&clean(payload.evidence).length<10)throw Error('Record evidence for confirming the source role');
 const fingerprint=a=>JSON.stringify([a?.fileId,a?.checksum,a?.byteCount,a?.durationMs,a?.width,a?.height]);
 const unchangedPrimary=prior?.primaryFileId===primary&&fingerprint(prior?.assets?.find(a=>a.fileId===primary))===fingerprint(merged.get(primary));
 const confirmed=payload.confirmPairing===true;
 if(confirmed&&(!primary||clean(payload.evidence).length<10))throw Error('Confirm the selected file and record pairing evidence');
 return {...prior,type:'recording',schemaVersion:'1',prioritySetId:candidate.prioritySetId,sourceFileId:candidate.sourceFileId,
  reviewedSourceFileId:candidate.sourceReviewFileId||candidate.sourceFileId,author:candidate.author,title:candidate.poemTitle,sourceEvent:candidate.sourceEvent,
  sourceRole:role,assets:[...merged.values()].sort((a,b)=>a.fileId.localeCompare(b.fileId)),primaryFileId:primary,
  sourceRoleReview:role!==candidate.sourceMediaType?{method:'administrator_confirmation',verifiedBy:actor,verifiedAt:now,evidence:clean(payload.evidence)}:prior?.sourceRoleReview,
  pairing:confirmed?{status:'verified',method:'administrator_confirmation',verifiedBy:actor,verifiedAt:now,evidence:clean(payload.evidence)}:unchangedPrimary&&prior?.pairing?.status==='verified'?prior.pairing:{status:'unconfirmed'},
  status:(confirmed||unchangedPrimary&&prior?.pairing?.status==='verified')&&['OM','FV'].includes(role)?'resolved':'needs_pairing',reason:confirmed?'confirmed_file_correspondence':'File links saved; pairing is separate',checkedAt:now};
}
export function planVideoMediaResolution(candidate,prior,scope,{files=[],complete=false,sourceCount=1,sourceAsset=null,selectedAsset=null,error=''}={},now=new Date().toISOString()){
 const base={...prior,type:'recording',schemaVersion:'1',prioritySetId:candidate.prioritySetId,sourceFileId:candidate.sourceFileId,
  reviewedSourceFileId:candidate.sourceReviewFileId||candidate.sourceFileId,author:candidate.author,title:candidate.poemTitle,sourceEvent:candidate.sourceEvent,
  sourceRole:prior?.sourceRole&&prior.sourceRole!=='unknown'?prior.sourceRole:candidate.sourceMediaType,assets:prior?.assets||[],checkedAt:now};
 const stop=(status,reason)=>({...base,status,reason,...(['needs_pairing','needs_classification'].includes(status)&&base.pairing?{pairing:{...base.pairing,status:'unconfirmed'}}:{})});
 if(['OM','FV'].includes(candidate.observedSourceRole)&&prior?.sourceRole&&candidate.observedSourceRole!==prior.sourceRole)return stop('needs_classification','Recorded source role conflicts with current Drive evidence');
 if(!['OM','FV'].includes(base.sourceRole))return stop('needs_classification','Confirm this reviewed file as FV or OM');
 if(['hold','reject'].includes(candidate.gate?.decision))return stop('held','Explicit release hold retained');
 if(sourceAsset&&['OM','FV'].includes(sourceAsset.mediaType)&&sourceAsset.mediaType!==base.sourceRole)return stop('needs_classification','Current source file conflicts with its confirmed media role');
 const sourceVersion=sourceAsset?{fileId:sourceAsset.fileId,checksum:sourceAsset.checksum,byteCount:sourceAsset.byteCount,durationMs:sourceAsset.durationMs}:null;
 if(prior?.reviewedSourceVersion&&sourceVersion&&['fileId','checksum','byteCount','durationMs'].some(k=>prior.reviewedSourceVersion[k]!==sourceVersion[k]))return stop('needs_pairing','Reviewed source media changed; confirm the recording version');
 if(sourceVersion)base.reviewedSourceVersion=sourceVersion;
 const oldPrimary=prior?.assets?.find(a=>a.fileId===prior.primaryFileId);
 if(oldPrimary&&selectedAsset&&JSON.stringify([oldPrimary.checksum,oldPrimary.byteCount,oldPrimary.durationMs,oldPrimary.width,oldPrimary.height])!==JSON.stringify([selectedAsset.checksum,selectedAsset.byteCount,selectedAsset.durationMs,selectedAsset.width,selectedAsset.height]))return stop('needs_pairing','The linked file changed; confirm the new media version');
 const rework=candidate.gate?.decision==='send_to_editing'&&clean(candidate.gate.editingInstructions)&&candidate.gate.decidedBy!=='weaver:auto';
 if(rework&&!(prior?.pairing?.method==='administrator_confirmation'&&Date.parse(prior.pairing.verifiedAt)>Date.parse(candidate.gate.decidedAt||'1970-01-01')))return stop('needs_pairing','Explicit editing instructions remain open; confirm the completed rework before release');
 if(prior?.pairing?.status==='verified'&&selectedAsset?.mediaType==='FV'&&selectedAsset.fileId===prior.primaryFileId){
  return {...base,assets:base.assets.map(a=>a.fileId===selectedAsset.fileId?selectedAsset:a),status:'resolved',reason:'Verified selected finished file',editingRequest:base.editingRequest?{...base.editingRequest,status:'resolved',resolvedAt:now}:undefined};
 }
 if(error)return stop('lookup_failed',error);
 if(sourceCount!==1)return stop('needs_pairing','More than one reviewed recording has this event, performer and title');
 if(base.sourceRole==='FV'){
  const direct=(candidate.reviewEvidence||[]).length>0&&(candidate.reviewEvidence||[]).every(r=>r.mapping==='direct_file_id'&&r.sourceFileId===candidate.sourceFileId);
  if(!sourceAsset||sourceAsset.mediaType!=='FV')return stop('needs_classification','Current review file is not verified FV');
  if(!candidate.sourceReviewFileId&&!direct)return stop('needs_identity','Review-source lineage requires confirmation');
  return {...base,status:'resolved',reason:'Verified FV review source',primaryFileId:sourceAsset.fileId,assets:[sourceAsset],pairing:{status:'verified',method:'exact_review_file_id',verifiedAt:now,policy:RESOLUTION_POLICY}};
 }
 if(!scope?.confirmedCoverage||scope.prioritySetId!==candidate.prioritySetId||scope.sourceFolderId!==candidate.sourceFolderId)return stop('needs_scope','Confirm the event’s finished-edit search folders');
 if(!complete)return stop('lookup_failed','Finished-edit lookup was not complete; no editing request created');
 const identity={author:normalizeVideoIdentity(candidate.author),title:normalizeVideoIdentity(candidate.poemTitle)};
 const matches=files.filter(f=>{const value=videoAssetIdentity(f.name);return f.mediaType==='FV'&&value&&value.author===identity.author&&value.title===identity.title;});
 base.lookup={status:'complete',checkedAt:now,scopeRevision:scope.revision||0,folderIds:scope.finishedFolderIds||[],matchCount:matches.length,policy:RESOLUTION_POLICY};
 if(!matches.length){
  if(base.assets.some(a=>a.mediaType==='FV'))return stop('needs_pairing','Previously linked FV was not verified; resolve it before requesting another edit');
  return {...base,status:'editing_requested',reason:'Confirmed OM; no matching FV in the complete approved event scope',editingRequest:base.editingRequest||{requestId:'edit-'+resolutionKey({...candidate,sourceFileId:candidate.sourceReviewFileId||candidate.sourceFileId}),status:'open',createdAt:now,sourceFileId:candidate.sourceFileId,sourceVideoUrl:candidate.sourceVideoUrl||`https://drive.google.com/file/d/${candidate.sourceFileId}/view`,sourceEvent:candidate.sourceEvent,author:candidate.author,title:candidate.poemTitle,destinationFolderIds:scope.finishedFolderIds||[],instructions:candidate.gate?.editingInstructions||'Create a finished export for this reviewed recording. Keep formats as separate files; upload to the event finished folders or link the completed output in Weaver.'}};
 }
 const unique=new Map([...base.assets,...matches.map(a=>({...a,pairingVerified:true}))].map(a=>[a.fileId,a]));base.assets=[...unique.values()].sort((a,b)=>a.fileId.localeCompare(b.fileId));
 const preferred=matches.filter(a=>a.height>a.width);const pool=preferred.length?preferred:matches;
 const sorted=[...pool].sort((a,b)=>(b.width*b.height)-(a.width*a.height)||a.fileId.localeCompare(b.fileId));
 if(!sorted[0]?.width||!sorted[0]?.height)return stop('needs_pairing','Export dimensions are not verified');
 if(sorted[1]&&sorted[0].width*sorted[0].height===sorted[1].width*sorted[1].height)return stop('needs_pairing','Multiple equally preferred exports; choose the intended version');
 return {...base,status:'resolved',reason:'Unique exact event/performer/title correspondence',primaryFileId:sorted[0].fileId,
  pairing:{status:'verified',method:'exact_event_author_title_unique_reviewed_recording',policy:RESOLUTION_POLICY,scopeRevision:scope.revision||0,verifiedAt:now},
  ...(base.editingRequest?{editingRequest:{...base.editingRequest,status:'resolved',resolvedAt:now}}:{})};
}
export function assertResolvedVideoMedia(candidate,now=Date.now()){
 const r=candidate.mediaResolution,primary=r?.assets?.find(a=>a.fileId===r.primaryFileId);
 if(r?.status!=='resolved'||r.pairing?.status!=='verified'||primary?.mediaType!=='FV'||!primary.url)throw Error('Resolve and verify the recording’s finished media before delivery');
 if(!Number.isFinite(Date.parse(r.checkedAt))||now-Date.parse(r.checkedAt)>30*60000||Date.parse(r.checkedAt)>now+300000)throw Error('Finished-media verification needs a fresh scheduled check');
 let released;try{const url=new URL(candidate.gate?.publishableAssetUrl);if(url.protocol==='https:'&&url.hostname==='drive.google.com')released=url.pathname.match(/^\/file\/d\/([\w-]+)/)?.[1];}catch{}
 if(released!==primary.fileId)throw Error('Release asset differs from the verified selected file');
 return primary;
}

export function assertSelectedAssetVersion(asset,file){
 if(!file||file.id!==asset.fileId||file.trashed||!String(file.mimeType).startsWith('video/')||Number(file.size)!==asset.byteCount)throw Error('Selected media is unavailable or its bytes changed');
 if(asset.checksum&&file.md5Checksum!==asset.checksum)throw Error('Selected media checksum changed; resolve the new version');
 for(const [field,meta]of [['width','width'],['height','height'],['durationMs','durationMillis']])if(asset[field]&&Number(file.videoMediaMetadata?.[meta]||0)!==asset[field])throw Error('Selected media dimensions or duration changed');
 return asset;
}
