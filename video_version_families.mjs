// Explicit recording relationships. File names alone never establish identity.
// Original file/review/receiver IDs remain stable; grouping is a read model.
export const VIDEO_VERSION_FAMILIES = Object.freeze([
  {prioritySetId:'bpl-charm-city-2026',sourceFileId:'1U_asBkgfiGzHjc67v4FVrLd7SmCKOOsZ',
    reviewSourceFileId:'1L9ighgQVutO_omi3MF8kQUUwSSizI48i',author:'Black Chakra',title:'Click',
    fileIds:['1U_asBkgfiGzHjc67v4FVrLd7SmCKOOsZ','14hFbu2jE-UYzsutS7cJcaJn72-c-aLDZ','1Cdqr19wjk0mFIzYFsKw9wHG8dMpIIbPK','12INXPROJ_74JsW3g9kLL84ZSR7-78oqa','1L9ighgQVutO_omi3MF8kQUUwSSizI48i']},
  {prioritySetId:'bpl-charm-city-2026',sourceFileId:'1HWgoqACLa4iqfKYs8ErKVgrEyYmeHTa_',
    reviewSourceFileId:'1NDT7CNYldKupqwsezFUy5kuEETr3xCSx',author:'Usman Hameedi',title:'Waterfall',
    fileIds:['1HWgoqACLa4iqfKYs8ErKVgrEyYmeHTa_','1oY_vLVMtivNNu5lwBn6mf2FQN1Szpj4I','1cc8qZbqLdCe8EKVQNEiFseDsDYwyngfP','1NDT7CNYldKupqwsezFUy5kuEETr3xCSx']}
]);
const clean=v=>String(v??'').trim();
export function videoFamilyForFile(id,setId,families=VIDEO_VERSION_FAMILIES){
  const matches=families.filter(f=>(!setId||f.prioritySetId===setId)&&f.fileIds.includes(id));
  if(matches.length>1)throw Error('Conflicting video performance relationships');
  return matches[0]||null;
}
export function groupVideoVersionFiles(files,setId,families=VIDEO_VERSION_FAMILIES){
  const byId=new Map(files.map(f=>[f.id,f])),seen=new Set(),result=[];
  for(const file of files){
    const family=videoFamilyForFile(file.id,setId,families);
    if(!family){result.push(file);continue;}
    if(seen.has(family.reviewSourceFileId))continue;
    seen.add(family.reviewSourceFileId);
    // Do not substitute an arbitrary new primary if the configured representative disappears.
    const representative=byId.get(family.sourceFileId);
    if(!representative)throw Error('Linked performance representative is unavailable');
    result.push({...representative,videoVersionFamily:family,
      reviewSourceFileIds:[...family.fileIds],
      videoVariants:family.fileIds.filter(id=>byId.has(id)).map(id=>byId.get(id))});
  }
  return result;
}
export function videoPerformanceContext(file,excerpts=[]){
  const f=file?.videoVersionFamily;if(!f)return null;
  return{schemaVersion:1,performanceId:'weaver:performance:'+f.reviewSourceFileId,
    sourceRecordId:'weaver:video:'+f.reviewSourceFileId,
    poem:{author:f.author,title:f.title},
    assets:(file.videoVariants||[]).map(v=>({fileId:v.id,fileName:v.name,
      url:v.webViewLink||`https://drive.google.com/file/d/${v.id}/view`,
      relationship:'same_performance',timing:'source_specific'})),
    excerptRecordIds:[...new Set(excerpts.map(clean).filter(Boolean))].sort()};
}
export function candidateMatchesVideoSource(candidate,id){
  const family=videoFamilyForFile(id,candidate.prioritySetId);
  return candidate.sourceFileId===id||candidate.sourceReviewFileId===id
    ||Boolean(family&&candidate.sourceReviewFileId===family.reviewSourceFileId);
}
