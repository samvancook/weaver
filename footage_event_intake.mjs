import { createHash } from 'node:crypto';

export const INVENTORY_ID = '1npv-dpP22MV-cRs-DRjOimjbWIxf-xvbun5SdP6dl7U';
const text = value => String(value ?? '').trim();
const key = value => text(value).replace(/\s+/g, ' ').toLowerCase();
const hash = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');
const AUTHOR_CONTRACT = 'all button authors, all features, all workshop instructors - contract is not stored with video releases';
export function fileId(url) {
  try {
    const u = new URL(url);
    if (u.protocol !== 'https:' || u.hostname !== 'drive.google.com' || /\/folders\//.test(u.pathname)) return '';
    return (u.pathname.match(/\/file\/d\/([-\w]+)/)?.[1] || u.searchParams.get('id') || '').match(/^[-\w]{15,}$/)?.[0] || '';
  } catch { return ''; }
}
export function validateEventManifest(input) {
  if (input?.sourceSpreadsheetId !== INVENTORY_ID) throw new Error('Unsupported source inventory.');
  if (!text(input.eventName) || !text(input.omFolderId).match(/^[-\w]{15,}$/)) throw new Error('Event name and OM event folder ID are required.');
  if (input.closed !== false) throw new Error('This MVP accepts open events only.');
  if (!Array.isArray(input.titles) || !input.titles.length || input.titles.length > 500) throw new Error('Send between 1 and 500 title records.');
  const titles = input.titles.map(t => {
    if (!Number.isSafeInteger(t.sourceRow) || t.sourceRow < 2) throw new Error('Every title needs its source response row.');
    return { sourceRow: t.sourceRow, performer: text(t.performer), title: text(t.title), rights: text(t.rights), contentLink: text(t.contentLink) };
  }).sort((a,b) => a.sourceRow-b.sourceRow);
  if (new Set(titles.map(t=>t.sourceRow)).size !== titles.length) throw new Error('Duplicate source response row.');
  const manifest = { sourceSpreadsheetId: INVENTORY_ID, eventName: text(input.eventName).replace(/\s+/g,' '), omFolderId: text(input.omFolderId), closed: false, titles };
  return { manifest, eventId: 'inventory-' + hash([INVENTORY_ID,manifest.omFolderId,key(manifest.eventName)]).slice(0,32), digest: hash(manifest) };
}
function validVideo(file, id) {
  return file?.id === id && !file.trashed && String(file.mimeType || '').startsWith('video/');
}
export function describeVideoAsset(file, url, role, mediaType = 'unknown', mediaTypeEvidence = '') {
  const width = Number(file.videoMediaMetadata?.width) || null;
  const height = Number(file.videoMediaMetadata?.height) || null;
  return {assetId:'drive:'+file.id,driveFileId:file.id,fileName:text(file.name),url,
    mimeType:file.mimeType,role,mediaType:['OM','FV'].includes(mediaType)?mediaType:'unknown',mediaTypeEvidence,
    format:width && height ? (width>height?'HV':height>width?'VV':'square'):'unknown',
    width,height,durationMillis:file.videoMediaMetadata?.durationMillis || null};
}
function resultStatus(record, complete) {
  return complete ? (record.mode === 'existing_event' ? 'linked_existing_event' : 'received_for_curation') : 'needs_attention';
}

// A receipt acknowledges queue registration/association, never event completeness.
export async function prepareEventIntake(input, getFile, target = {mode:'new_event'}) {
  const {manifest,eventId,digest} = validateEventManifest(input);
  const accepted=[],blocked=[],seen=new Set();
  for (const title of manifest.titles) {
    const id=fileId(title.contentLink);
    const targetId=target.mode==='existing_event' ? target.pairs?.get(id) : id;
    let reason = !title.performer || !title.title ? 'missing_performer_or_title'
      : !['have video release',AUTHOR_CONTRACT].includes(key(title.rights)) ? 'rights_need_review'
      : !id ? 'individual_drive_video_required'
      : seen.has(id) ? 'duplicate_video_in_event'
      : !targetId ? 'existing_event_asset_connection_missing'
      : target.mode==='existing_event' && !target.available?.has(targetId) ? 'existing_curation_video_not_in_queue'
      : '';
    if (!reason) {
      try {
        const source=await getFile(id);
        const curation=targetId===id ? source : await getFile(targetId);
        if (!validVideo(source,id) || !validVideo(curation,targetId)) reason='not_an_accessible_video';
        else {
          seen.add(id);
          const sourceAsset=describeVideoAsset(source,title.contentLink,'inventory_source',target.inventoryMediaType,target.mediaTypeEvidence);
          const curationAsset=targetId===id ? sourceAsset : describeVideoAsset(curation,
            text(curation.webViewLink)||`https://drive.google.com/file/d/${targetId}/view`,'existing_curation',target.curationMediaType,target.mediaTypeEvidence);
          const finishedVideoFileId=curationAsset.mediaType==='FV'?targetId:sourceAsset.mediaType==='FV'?id:null;
          accepted.push({...title,sourceFileId:targetId,sourceFileName:text(curation.name),mimeType:curation.mimeType,
            inventorySourceFileId:id,sourceVideoUrl:curationAsset.url,
            videoReleaseStatus:'inventory_curation_only',publicationRestricted:true,
            familyId:'footage-family-'+hash([target.prioritySetId||eventId,targetId]).slice(0,32),
            textReference:{author:title.performer,title:title.title,canonicalTextId:null,status:'not_linked'},
            assets:targetId===id?[sourceAsset]:[sourceAsset,curationAsset],
            workflow:{originalMediaFileId:sourceAsset.mediaType==='OM'?id:null,finishedVideoFileId,
              editingStatus:finishedVideoFileId?'finished_video_linked':sourceAsset.mediaType==='OM'?'needs_editing':'media_type_needs_review',
              publicationAssetFileId:finishedVideoFileId,publicationStatus:finishedVideoFileId?'awaiting_curation_approval':'awaiting_finished_video'},
            relationship:targetId===id?'same_asset':'corresponding_event_versions',
            derivationEvidence:target.evidence||'Inventory-linked individual video.',
            exactDerivationVerified:false});
        }
      } catch { reason='weaver_cannot_read_video'; }
    }
    if(reason) blocked.push({...title,reason});
  }
  const mode=target.mode||'new_event';
  const prioritySetId=target.prioritySetId||eventId;
  const complete=accepted.length>0 && blocked.length===0;
  return {schemaVersion:2,eventId,prioritySetId,digest,manifest,eventName:manifest.eventName,
    mode,targetEventName:target.label||manifest.eventName,
    routingDigest:hash([mode,prioritySetId,manifest.titles.map(t=>[fileId(t.contentLink),target.mode==='existing_event'?target.pairs?.get(fileId(t.contentLink))||'':fileId(t.contentLink)])]),
    sourceFolderId:manifest.omFolderId,sourceMediaType:target.inventoryMediaType||'unknown',accepted,blocked,
    submittedCount:manifest.titles.length,acceptedCount:accepted.length,blockedCount:blocked.length,
    linkedExistingCount:mode==='existing_event'?accepted.length:0,
    inventoryAssetCount:new Set(accepted.map(t=>t.inventorySourceFileId)).size,
    curationAssetCount:new Set(accepted.map(t=>t.sourceFileId)).size,
    familyCount:new Set(accepted.map(t=>t.familyId)).size,
    transferComplete:complete,readyForClosureReview:false,eventCompleteness:'not_assessed',
    status:resultStatus({mode},complete),accessCheck:'accepted_asset_metadata_checked',
    canonicalTextLinkStatus:'not_linked',checkedAt:new Date().toISOString(),playbackVerified:false,closedAutomatically:false};
}

export function footageEventAsSet(record) {
  // Linking never shadows a live queue or changes its review source identities.
  if(record.mode==='existing_event') return null;
  const types=new Set((record.accepted||[]).map(t=>t.assets?.find(a=>a.driveFileId===t.sourceFileId)?.mediaType || 'unknown'));
  const mediaType=types.size===1 && ['OM','FV'].includes([...types][0])?[...types][0]:'unknown';
  return {id:record.eventId,label:record.eventName,eventName:record.eventName,folderId:record.sourceFolderId,
    inventoryEvent:true,sourceMediaType:mediaType,sourceMediaLabel:'Inventory-linked video',
    inventoryFiles:(record.accepted||[]).map(t=>({id:t.sourceFileId,name:t.sourceFileName,mimeType:t.mimeType,webViewLink:t.sourceVideoUrl})),
    inventoryTitles:Object.fromEntries((record.accepted||[]).map(t=>[t.sourceFileId,t]))};
}

export async function currentEventReceipt(record,getFile,verifyQueue=async()=>true,readProgress=async()=>({status:'not_checked'})) {
  const ids=[...new Set((record.accepted||[]).flatMap(t=>t.assets?.map(a=>a.driveFileId)||[t.sourceFileId]))];
  const unavailable=[];
  for(const id of ids) {
    try {if(!validVideo(await getFile(id),id)) unavailable.push(id);}
    catch {unavailable.push(id);}
  }
  const targetQueueAvailable=record.mode!=='existing_event'||await verifyQueue(record);
  const complete=record.acceptedCount>0 && record.blockedCount===0 && !unavailable.length && targetQueueAvailable;
  let progress;
  try { progress=await readProgress(record); }
  catch { progress={status:'unavailable'}; }
  return {...record,accepted:(record.accepted||[]).map(t=>({...t,progress:progress.byFileId?.[t.sourceFileId] || null})),
    progressCheck:{status:progress.status,checkedAt:progress.checkedAt || null},
    currentUnavailableFileIds:unavailable,targetQueueAvailable,
    transferComplete:complete,readyForClosureReview:false,eventCompleteness:'not_assessed',
    status:resultStatus(record,complete),checkedAt:new Date().toISOString(),
    playbackVerified:false,closedAutomatically:false};
}
