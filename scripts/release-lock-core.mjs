import {randomUUID} from 'node:crypto';

const iso=now=>new Date(now).toISOString();
const required=(value,label)=>{if(!value)throw Error(label);};
export function planAcquireRelease(current,input,now=Date.now()){
  required(!current||current.schemaVersion==='1','Unknown release lock schema; reconcile before another release');
  required(!current||['released','aborted'].includes(current.status),
    'Weaver release lock is held or unresolved; reconcile its owner before another release (including expired leases)');
  required(/^[a-f0-9]{40}$/.test(input.commit||''),'Full approved source commit required');
  required(/^[^\s@]+@[^\s@]+$/.test(input.owner||''),'Named release owner email required');
  required(/^weaver-/.test(input.baseRevision||''),'Current Weaver base revision required');
  required(/@sha256:[a-f0-9]{64}$/.test(input.baseImage||''),'Immutable current base image required');
  required(/^[a-f0-9]{64}$/.test(input.manifestSha256||''),'Source manifest hash required');
  for(const tag of ['footage-bridge','video-queues'])required(input.baseAliases?.[tag],'Current operational alias targets required');
  return {schemaVersion:'1',status:'held',claimId:randomUUID(),owner:input.owner.toLowerCase(),
    commit:input.commit,baseRevision:input.baseRevision,baseImage:input.baseImage,
    manifestSha256:input.manifestSha256,baseAliases:Object.fromEntries(['footage-bridge','video-queues'].map(tag=>[tag,input.baseAliases[tag]])),acquiredAt:iso(now),updatedAt:iso(now),
    expiresAt:iso(now+60*60000),previousOutcome:current?{
      status:current.status,owner:current.owner,commit:current.commit,updatedAt:current.updatedAt}:null};
}
export function assertOwnedRelease(current,receipt,now=Date.now(),requireFresh=true){
  required(current?.status==='held','No active Weaver release claim');
  for(const key of ['claimId','owner','commit','baseRevision','baseImage','manifestSha256']){
    required(current[key]&&current[key]===receipt[key],'Release claim changed or belongs to another owner');
  }
  for(const tag of ['footage-bridge','video-queues'])required(current.baseAliases?.[tag]&&current.baseAliases[tag]===receipt.baseAliases?.[tag],'Release alias snapshot changed');
  if(requireFresh)required(Date.parse(current.expiresAt)>now,'Release lease expired; reconcile and abort before a fresh acquisition');
  return current;
}
export function assertReleaseBase(current,service){
  required(service.revision===current.baseRevision&&service.image===current.baseImage,
    'Serving revision/image changed; preserve the newer release before continuing');
  for(const tag of ['footage-bridge','video-queues'])required(service.aliases?.[tag]===current.baseAliases[tag],'Operational alias changed; reconcile before continuing');
}
export function planRenewRelease(current,receipt,service,now=Date.now()){
  assertOwnedRelease(current,receipt,now);assertReleaseBase(current,service);
  return {...current,updatedAt:iso(now),expiresAt:iso(now+60*60000)};
}
export function planCloseRelease(current,receipt,{outcome,service,verification,activeBuilds=0},now=Date.now()){
  assertOwnedRelease(current,receipt,now,false);
  required(activeBuilds===0,'A Weaver build is still active; keep the release lock');
  if(outcome==='aborted')assertReleaseBase(current,service);
  else if(outcome==='released'){
    required(verification?.ok===true&&verification.revision===service.revision,'Exact verified release receipt required');
    required(verification.sourceCommit===current.commit&&verification.sourceManifestSha256===current.manifestSha256,'Verification must bind the approved source commit and manifest');
    required(service.revision!==current.baseRevision,'No new serving revision to finish');
    required(service.image.endsWith('@'+verification.imageDigest),'Verified image differs from serving image');
    for(const tag of ['footage-bridge','video-queues'])required(service.aliases?.[tag]===service.revision,'Main and operational aliases must agree');
    for(const flag of ['sourceImageVerified','authenticatedCanaryVerified','protectedRecordsVerified']){
      required(verification[flag]===true,'Missing release verification: '+flag);
    }
  }else throw Error('Unknown release outcome');
  return {...current,status:outcome,updatedAt:iso(now),finishedAt:iso(now),
    finalRevision:service.revision,finalImage:service.image};
}

// The store's write must use the observed updateTime (or exists=false on create).
// Losing an acquisition race is an error, never an automatic retry/takeover.
export async function acquireRelease(store,input,now=Date.now(),onPlanned=()=>{}){
  const before=await store.read();const claim=planAcquireRelease(before?.value,input,now);
  // Save the exact intended claim before the request: a lost acknowledgement
  // can then be reconciled without creating a second claim or stealing a lock.
  await onPlanned(claim);
  await store.write(claim,before?.version||null);return claim;
}
