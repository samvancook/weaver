import test from 'node:test';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {hostedVideoObject,inspectExistingHostedVideo} from '../video_hosted_media.mjs';
import {applyVideoPipelineReceiver,localVideoPipelineState} from '../video_pipeline.mjs';

const bytes=Buffer.from('existing hosted video bytes');
const md5=createHash('md5').update(bytes).digest('base64');
const c={prioritySetId:'event',sourceFileId:'finished',videoRecordId:'weaver:video:review',
  existingCanonicalVideoId:'VIDEO',sourceMediaType:'FV',reviewCount:5,baseScore:9,excerptRecordIds:['one']};
const receipt={ok:true,received:true,canonicalVideoId:'VIDEO',sourceRecordId:c.videoRecordId,
  media:{videoUrl:'https://firebasestorage.googleapis.com/v0/b/poetry-please.firebasestorage.app/o/content-library%2Fvideos%2Fvideo%2Fexisting.mov?alt=media&token=private'}};
const headers={'content-type':'video/quicktime','content-length':String(bytes.length),
  'x-goog-generation':'123','x-goog-hash':`crc32c=AAAAAA==, md5=${md5}`,etag:'original'};
function fixture(options={}) {
  let record=options.record||null;const calls=[],writes=[];
  const db=async(action,payload)=>{calls.push(action);if(action.startsWith('get_'))return{ok:true,record};
    writes.push(payload);record=payload.record;return{ok:true,record};};
  const fetchImpl=async(url,init)=>{calls.push(init.method);assert.equal(init.redirect,'error');
    return new Response(init.method==='HEAD'?null:(options.bytes||bytes),{status:200,headers:{...headers,...(init.method==='GET'?options.getHeaders:options.headHeaders)}});};
  return{db,fetchImpl,calls,writes,get record(){return record;}};
}
test('verifies exact existing bytes once, then revalidates generation and checksum without another download',async()=>{
  const f=fixture(),one=await inspectExistingHostedVideo(c,receipt,f);
  assert.equal(one.status,'verified');assert.equal(f.writes.length,1);assert.equal(one.byteCount,bytes.length);
  assert.equal(one.sha256,createHash('sha256').update(bytes).digest('hex'));
  assert.equal(JSON.stringify(f.record).includes('private'),false);
  assert.equal(JSON.stringify(f.record).includes('calendarReady'),false);
  await inspectExistingHostedVideo(c,receipt,f);assert.equal(f.calls.filter(x=>x==='GET').length,1);
  assert.equal(f.calls.filter(x=>x==='HEAD').length,2);
});
test('does not fetch crossed parents, source identities, foreign hosts, objects or buckets',async()=>{
  const changed=[{...receipt,canonicalVideoId:'OTHER'},{...receipt,sourceRecordId:'other'},
    {...receipt,ambiguousVideoMatch:true},...['https://example.com/private.mov',
    receipt.media.videoUrl.replace('poetry-please.firebasestorage.app','foreign-bucket'),
    receipt.media.videoUrl.replace('%2Fvideo%2F','%2Fother%2F')].map(videoUrl=>({...receipt,media:{videoUrl}}))];
  for(const r of changed){const f=fixture();assert.equal(hostedVideoObject(c,r),null);
    assert.equal((await inspectExistingHostedVideo(c,r,f)).status,'not_applicable');assert.equal(f.calls.length,0);}
});
test('changed generation between HEAD and download, corrupt bytes and oversized objects never persist evidence',async()=>{
  for(const options of [{getHeaders:{'x-goog-generation':'124'}},{bytes:Buffer.alloc(bytes.length)},
    {headHeaders:{'content-length':String(601*1024*1024)}}]){
    const f=fixture(options);assert.notEqual((await inspectExistingHostedVideo(c,receipt,f)).status,'verified');assert.equal(f.writes.length,0);
  }
});
test('a cache conflict holds verification; the download budget never forces another transfer',async()=>{
  const first=fixture();await inspectExistingHostedVideo(c,receipt,first);
  const f=fixture({record:{...first.record,md5Hash:'A'.repeat(22)+'=='}});
  assert.equal((await inspectExistingHostedVideo(c,receipt,f)).status,'hosted_verification_conflict');assert.equal(f.writes.length,0);
  const pending=fixture();assert.equal((await inspectExistingHostedVideo(c,receipt,{...pending,claimDownload:()=>false})).status,'awaiting_hosted_byte_verification');
  assert.equal(pending.calls.includes('GET'),false);
});
test('existing hosted-media verification advances only media and preserves separate caption/calendar gates',async()=>{
  const f=fixture(),hostedMedia=await inspectExistingHostedVideo(c,receipt,f);
  const receiver={...receipt,media:{},review:{reviewCount:5,selectedExcerptIdCount:1}};
  const out=applyVideoPipelineReceiver(localVideoPipelineState(c),c,{receiver,receipt,hostedMedia});
  assert.equal(out.mediaReady,true);assert.equal(out.mediaReadinessReason,'existing_hosted_bytes_verified');
  assert.equal(out.stage,'awaiting_pp_readiness');assert.equal(out.captionReady,false);assert.equal(out.calendarReady,false);
  const crossed=applyVideoPipelineReceiver(localVideoPipelineState(c),c,{receiver,receipt,hostedMedia:{...hostedMedia,canonicalVideoId:'OTHER'}});
  assert.equal(crossed.mediaReady,false);
});
