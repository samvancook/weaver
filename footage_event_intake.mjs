import { createHash } from 'node:crypto';

export const INVENTORY_ID = '1npv-dpP22MV-cRs-DRjOimjbWIxf-xvbun5SdP6dl7U';
const text = value => String(value ?? '').trim();
const key = value => text(value).replace(/\s+/g, ' ').toLowerCase();
const hash = value => createHash('sha256').update(JSON.stringify(value)).digest('hex');
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

// Drive metadata is checked by Weaver's runtime identity, not asserted by Helper.
// This acknowledges a curation queue; it never claims publication or playback approval.
export async function prepareEventIntake(input, getFile) {
  const {manifest,eventId,digest} = validateEventManifest(input);
  const accepted=[],blocked=[],seen=new Set();
  for (const title of manifest.titles) {
    const id=fileId(title.contentLink);
    let reason = !title.performer || !title.title ? 'missing_performer_or_title'
      : key(title.rights)!=='have video release' ? 'rights_need_review'
      : !id ? 'individual_drive_video_required'
      : seen.has(id) ? 'duplicate_video_in_event' : '';
    if (!reason) {
      try {
        const file=await getFile(id);
        if (file?.id!==id || file.trashed || !String(file.mimeType||'').startsWith('video/')) reason='not_an_accessible_video';
        else {
          seen.add(id);
          accepted.push({...title,sourceFileId:id,sourceFileName:text(file.name),mimeType:file.mimeType,
            sourceVideoUrl:title.contentLink,videoReleaseStatus:'inventory_curation_only',publicationRestricted:true});
        }
      } catch { reason='weaver_cannot_read_video'; }
    }
    if(reason) blocked.push({...title,reason});
  }
  return {schemaVersion:1,eventId,prioritySetId:eventId,digest,manifest,eventName:manifest.eventName,
    sourceFolderId:manifest.omFolderId,sourceMediaType:'unknown',accepted,blocked,
    submittedCount:manifest.titles.length,acceptedCount:accepted.length,blockedCount:blocked.length,
    readyForClosureReview:accepted.length>0 && blocked.length===0,
    status:accepted.length>0 && !blocked.length?'ready_for_closure_review':'needs_attention',
    accessCheck:'drive_metadata_verified',playbackVerified:false,closedAutomatically:false};
}

export function footageEventAsSet(record) {
  return {id:record.eventId,label:record.eventName,eventName:record.eventName,folderId:record.sourceFolderId,
    inventoryEvent:true,sourceMediaType:'unknown',sourceMediaLabel:'Inventory-linked video',
    inventoryFiles:(record.accepted||[]).map(t=>({id:t.sourceFileId,name:t.sourceFileName,mimeType:t.mimeType,webViewLink:t.sourceVideoUrl})),
    inventoryTitles:Object.fromEntries((record.accepted||[]).map(t=>[t.sourceFileId,t]))};
}

export async function currentEventReceipt(record,getFile) {
  const unavailable=[];
  for(const t of record.accepted||[]) {
    try { const f=await getFile(t.sourceFileId); if(f?.id!==t.sourceFileId||f.trashed||!String(f.mimeType||'').startsWith('video/')) unavailable.push(t.sourceFileId); }
    catch { unavailable.push(t.sourceFileId); }
  }
  const ready=record.acceptedCount>0 && record.blockedCount===0 && unavailable.length===0;
  return {...record,currentUnavailableFileIds:unavailable,readyForClosureReview:ready,
    status:ready?'ready_for_closure_review':'needs_attention',checkedAt:new Date().toISOString(),
    playbackVerified:false,closedAutomatically:false};
}
