// Explicit folder roles and exact inventory assets are evidence; filenames are not.
const clean=v=>String(v??'').trim();
export function classifyDriveFolderPath(folders) {
  const matches=folders.flatMap(folder=>['FV','OM'].filter(type=>(type==='FV'
    ? /(?:^|[\s([/-])(?:FV|Finished Video)(?:$|[\s)\]/:-])/i
    : /(?:^|[\s([/-])(?:OM|Original Media)(?:$|[\s)\]/:-])/i).test(clean(folder.name)))
    .map(type=>({id:clean(folder.id),name:clean(folder.name),type})));
  const types=new Set(matches.map(m=>m.type)),known=types.size===1;
  return {sourceMediaType:known?[...types][0]:'unknown',sourceMediaTypeFolderId:known?matches[0].id:'',
    sourceMediaTypeReason:known?'explicit_folder_role':types.size?'conflicting_folder_roles':'no_explicit_folder_role',
    sourceMediaTypeEvidence:matches};
}
export async function traceVideoFolderMediaType(folderId,getFolder) {
  const folders=[],visited=new Set();let currentId=folderId,traceStatus='complete';
  while(currentId){
    if(visited.has(currentId)){traceStatus='folder_cycle';break;}
    if(folders.length>=20){traceStatus='folder_depth_limit';break;}
    visited.add(currentId);
    let folder;
    try{folder=await getFolder(currentId);if(!folder||folder.trashed||!clean(folder.name))throw Error('Unavailable');}
    catch{traceStatus='ancestor_unavailable';break;}
    folders.push(folder);
    if((folder.parents||[]).length>1){traceStatus='multiple_parents';break;}
    currentId=folder.parents?.[0]||'';
  }
  const result=classifyDriveFolderPath(folders);
  // An unreadable shared-drive root must not erase an explicit role already read.
  // Structural ambiguity remains blocked; observed conflicting roles always win.
  if(['folder_cycle','folder_depth_limit','multiple_parents'].includes(traceStatus)){
    result.sourceMediaType='unknown';result.sourceMediaTypeFolderId='';result.sourceMediaTypeReason=traceStatus;
  }else if(result.sourceMediaType==='unknown'&&result.sourceMediaTypeReason!=='conflicting_folder_roles'&&traceStatus!=='complete')result.sourceMediaTypeReason=traceStatus;
  return {...result,sourceMediaTypeTraceStatus:traceStatus,sourceMediaTypeObservedFolders:folders.map(f=>({id:f.id,name:f.name})),
    ...(traceStatus!=='complete'?{sourceMediaTypeTraceStoppedAt:currentId}:{})};
}
export function videoFileMediaType(set,file,fallback) {
  if(!set.inventoryEvent)return fallback;
  const assets=(set.inventoryTitles?.[file.id]?.assets||[]).filter(a=>a.driveFileId===file.id);
  const types=new Set(assets.map(a=>a.mediaType||'unknown'));
  const type=assets.length&&types.size===1&&['OM','FV'].includes(assets[0].mediaType)?assets[0].mediaType:'unknown';
  return {sourceMediaType:type,sourceMediaTypeFolderId:'',sourceMediaLabel:'Inventory-linked video',
    sourceMediaTypeReason:type==='unknown'?(types.size>1?'conflicting_inventory_roles':'inventory_role_unverified'):'exact_inventory_asset',
    sourceMediaTypeEvidence:assets.map(a=>({driveFileId:a.driveFileId,mediaType:a.mediaType,mediaTypeEvidence:a.mediaTypeEvidence||null}))};
}

// Administrative evidence is private durable ledger data, bound to the exact
// currently listed source folder. It never changes approval or recording identity.
export function recordedFolderMediaType(observed,record,folderId,now=Date.now()) {
  if(observed.sourceMediaType!=='unknown'||!['no_explicit_folder_role','ancestor_unavailable'].includes(observed.sourceMediaTypeReason))return observed;
  const path=record?.folderPath||[],first=observed.sourceMediaTypeObservedFolders?.[0];
  const connected=path.length>1&&path.every((f,i)=>i===path.length-1||f.parentId===path[i+1].id);
  const proof=classifyDriveFolderPath(path);
  if(record?.schemaVersion!=='1'||record.status!=='confirmed'||record.verificationSource!=='drive_metadata'
    ||record.sourceFolderId!==folderId||path[0]?.id!==folderId||first?.id!==folderId||path[0]?.name!==first.name
    ||!connected||!['OM','FV'].includes(record.mediaType)||proof.sourceMediaType!==record.mediaType
    ||!clean(record.verifiedBy).includes('@')||!Number.isFinite(Date.parse(record.verifiedAt))||Date.parse(record.verifiedAt)>now+300000)return observed;
  return {...observed,sourceMediaType:record.mediaType,sourceMediaTypeFolderId:proof.sourceMediaTypeFolderId,
    sourceMediaTypeReason:'recorded_drive_ancestry',sourceMediaTypeEvidence:proof.sourceMediaTypeEvidence,
    sourceMediaTypeVerification:{sourceFolderId:folderId,verifiedAt:record.verifiedAt,verificationSource:record.verificationSource}};
}
