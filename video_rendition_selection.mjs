const clean=v=>String(v??'').trim();
const driveId=v=>clean(v).match(/\/file\/d\/([\w-]+)/)?.[1]||'';
export const VERIFIED_PLATFORM_PROFILES=Object.freeze(['platform-1080p-h264-aac','platform-720p-h264-aac']);
function sourceVersions(candidate){
 return new Set([candidate.sourceMediaType==='FV'?clean(candidate.sourceFileId):driveId(candidate.gate?.publishableAssetUrl),
  ...(candidate.versionSourceFileIds||[])].map(clean).filter(Boolean));
}
export function isVerifiedPlatformRendition(candidate,canonicalVideoId,sourceRecordId,asset){
 return Boolean(asset&&clean(asset.id||asset.assetId)&&asset.contentCollection==='videos'
  &&asset.contentDocId===canonicalVideoId&&asset.sourceRecordId===sourceRecordId
  &&sourceVersions(candidate).has(asset.sourceVersionFileId)
  &&VERIFIED_PLATFORM_PROFILES.includes(asset.profileId)&&String(asset.profileVersion)==='1'
  &&asset.status==='active'&&asset.storageMode==='drive_reference'&&asset.validationState==='verified_bytes'
  &&clean(asset.outputFileId)&&Number.isSafeInteger(Number(asset.verifiedByteCount))&&Number(asset.verifiedByteCount)>0
  &&Number(asset.verifiedByteCount)===Number(asset.byteCount)
  &&/^[a-f0-9]{64}$/i.test(clean(asset.verifiedSha256))&&asset.verifiedSha256===asset.sha256);
}
export function preferredVerifiedPlatformRendition(candidate,canonicalVideoId,sourceRecordId,assets){
 const valid=assets.filter(r=>isVerifiedPlatformRendition(candidate,canonicalVideoId,sourceRecordId,r));
 for(const profile of VERIFIED_PLATFORM_PROFILES){
  const matches=valid.filter(r=>r.profileId===profile);
  if(!matches.length)continue;
  const current=matches.filter(r=>r.sourceVersionFileId===candidate.sourceFileId);
  const preferred=current.length?current:matches;
  // Equally preferred outputs require an explicit choice, not a filename sort.
  const unique=new Map(preferred.map(r=>[r.id||r.assetId,r]));
  return unique.size===1?{asset:[...unique.values()][0],status:'verified_profile'}:{asset:null,status:'ambiguous_verified_outputs'};
 }
 return {asset:null,status:'no_verified_platform_output'};
}
export function selectedPlatformRenditionVerified(candidate,receiver){
 const asset=receiver.selectedRendition;
 if(!isVerifiedPlatformRendition(candidate,receiver.canonicalVideoId,receiver.sourceRecordId,asset))return false;
 const fields=['id','assetId','contentCollection','contentDocId','sourceRecordId','sourceVersionFileId','outputFileId','profileId','profileVersion','sha256','verifiedSha256','byteCount','verifiedByteCount','status','storageMode','validationState'];
 return (receiver.renditions||[]).some(r=>fields.every(k=>String(r[k]??'')===String(asset[k]??'')));
}
