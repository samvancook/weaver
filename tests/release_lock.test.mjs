import test from 'node:test';import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {fileURLToPath} from 'node:url';
import {planAcquireRelease,assertOwnedRelease,assertReleaseBase,planRenewRelease,planCloseRelease,acquireRelease} from '../scripts/release-lock-core.mjs';
const now=Date.parse('2026-11-23T20:00:00Z'),input={owner:'operator@example.com',commit:'a'.repeat(40),baseRevision:'weaver-base',baseImage:'registry/image@sha256:'+'b'.repeat(64),manifestSha256:'c'.repeat(64),baseAliases:{'footage-bridge':'weaver-base','video-queues':'weaver-base'}};
const service={revision:input.baseRevision,image:input.baseImage,aliases:input.baseAliases};
test('the maintained deploy entry point stops before build when no shared claim is provided',()=>{
 const root=fileURLToPath(new URL('../',import.meta.url));
 const result=spawnSync('bash',['deploy-cloud-run.sh'],{cwd:root,encoding:'utf8',env:{...process.env,WEAVER_RELEASE_RECEIPT:''}});
 assert.notEqual(result.status,0);assert.match(result.stderr,/Acquire the shared release lock first/);assert.equal(result.stdout,'');
});
test('one durable conditional acquisition wins across independent callers',async()=>{
 let row=null,version=null;
 const store={read:async()=>row?{value:structuredClone(row),version}:null,write:async(v,expected)=>{if(expected!==version)throw Error('CAS conflict');row=v;version=String(Number(version||0)+1);}};
 const results=await Promise.allSettled([acquireRelease(store,input,now),acquireRelease(store,{...input,owner:'other@example.com'},now)]);
 assert.equal(results.filter(x=>x.status==='fulfilled').length,1);assert.equal(results.filter(x=>x.status==='rejected').length,1);assert.equal(row.status,'held');
});
test('expired, unknown and wrong-owner claims cannot be taken over or used',()=>{
 const claim=planAcquireRelease(null,input,now);
 assert.throws(()=>planAcquireRelease(claim,input,now+2*60*60000),/held or unresolved/);
 assert.throws(()=>planAcquireRelease({schemaVersion:'1',status:'unknown'},input,now),/unresolved/);
 assert.throws(()=>planAcquireRelease({schemaVersion:'future',status:'released'},input,now),/schema/);
 assert.throws(()=>assertOwnedRelease(claim,{...claim,claimId:'other'},now),/another owner/);
 assert.throws(()=>assertOwnedRelease(claim,claim,now+2*60*60000),/expired/);
 assert.throws(()=>assertReleaseBase(claim,{...service,revision:'weaver-newer'}),/changed/);
 assert.throws(()=>planRenewRelease(claim,claim,{...service,image:'other'},now),/changed/);
 assert.throws(()=>assertReleaseBase(claim,{...service,aliases:{...service.aliases,'video-queues':'weaver-other'}}),/alias changed/);
});
test('claim intent survives a lost write acknowledgement and is never silently retried',async()=>{
 let receipt,stored,writes=0;const store={read:async()=>null,write:async claim=>{stored=claim;writes++;throw Error('lost response');}};
 await assert.rejects(acquireRelease(store,input,now,claim=>{receipt=structuredClone(claim);}),/lost response/);
 assert.equal(writes,1);assert.equal(receipt.claimId,stored.claimId);assertOwnedRelease(stored,receipt,now);
});
test('abort requires unchanged production, exact ownership and no in-flight build',()=>{
 const claim=planAcquireRelease(null,input,now);
 for(const patch of [{activeBuilds:1},{service:{...service,revision:'weaver-newer'}}])assert.throws(()=>planCloseRelease(claim,claim,{outcome:'aborted',service,activeBuilds:0,...patch},now));
 const closed=planCloseRelease(claim,claim,{outcome:'aborted',service},now);
 assert.equal(closed.status,'aborted');assert.notEqual(planAcquireRelease(closed,input,now).claimId,claim.claimId);
 assert.throws(()=>assertOwnedRelease(closed,claim,now),/No active/);
});
test('finishing requires exact image, main/alias agreement and actual verification evidence',()=>{
 const claim=planAcquireRelease(null,input,now),imageDigest='sha256:'+'d'.repeat(64),revision='weaver-new';
 const live={revision,image:'registry/image@'+imageDigest,aliases:{'footage-bridge':revision,'video-queues':revision}};
 const proof={ok:true,revision,imageDigest,sourceCommit:input.commit,sourceManifestSha256:input.manifestSha256,sourceImageVerified:true,authenticatedCanaryVerified:true,protectedRecordsVerified:true};
 assert.equal(planCloseRelease(claim,claim,{outcome:'released',service:live,verification:proof},now).status,'released');
 for(const verification of [{...proof,ok:false},{...proof,imageDigest:'wrong'},{...proof,sourceCommit:'wrong'},{...proof,sourceManifestSha256:'wrong'},{...proof,protectedRecordsVerified:false}])assert.throws(()=>planCloseRelease(claim,claim,{outcome:'released',service:live,verification},now));
 assert.throws(()=>planCloseRelease(claim,claim,{outcome:'released',service:{...live,aliases:{}},verification:proof},now),/aliases/);
});
