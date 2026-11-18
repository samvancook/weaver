import {createHash} from 'node:crypto';

const BUCKET = 'poetry-please.firebasestorage.app';
const MAX_BYTES = 600 * 1024 * 1024;
const clean = value => String(value ?? '').trim();

export function hostedVideoObject(candidate, receipt) {
  if (receipt?.ok !== true || receipt.received !== true || receipt.ambiguousVideoMatch
    || receipt.sourceRecordId !== candidate.videoRecordId
    || !candidate.existingCanonicalVideoId || receipt.canonicalVideoId !== candidate.existingCanonicalVideoId) return null;
  try {
    const url = new URL(receipt.media?.videoUrl);
    if (url.protocol !== 'https:' || url.hostname !== 'firebasestorage.googleapis.com'
      || url.port || url.username || url.password) return null;
    const match = url.pathname.match(/^\/v0\/b\/([^/]+)\/o\/(.+)$/);
    if (!match || match[1] !== BUCKET) return null;
    const objectName = decodeURIComponent(match[2]);
    if (!objectName.startsWith(`content-library/videos/${receipt.canonicalVideoId.toLowerCase()}/`)
      || objectName.includes('/../')) return null;
    return {schemaVersion:'1', sourceRecordId:receipt.sourceRecordId,
      canonicalVideoId:receipt.canonicalVideoId, bucket:BUCKET, objectName};
  } catch { return null; }
}

function observedObject(response) {
  const byteCount = Number(response.headers.get('content-length'));
  const generation = clean(response.headers.get('x-goog-generation'));
  const md5Hash = clean(response.headers.get('x-goog-hash')).split(',').map(s => s.trim())
    .find(s => s.startsWith('md5='))?.slice(4);
  if (response.status !== 200 || !/^video\/(mp4|quicktime)$/.test(clean(response.headers.get('content-type')).split(';')[0])
    || !Number.isSafeInteger(byteCount) || byteCount <= 0 || byteCount > MAX_BYTES
    || !/^\d+$/.test(generation) || !/^[A-Za-z0-9+/]{22}==$/.test(md5Hash || '')) return null;
  return {byteCount,generation,md5Hash};
}

export function hostedVideoEvidenceMatches(candidate, receipt, evidence) {
  const object = hostedVideoObject(candidate,receipt);
  return Boolean(object && evidence?.status === 'verified' && evidence.validation === 'full_bytes_md5'
    && Object.entries(object).every(([key,value]) => evidence[key] === value)
    && /^\d+$/.test(evidence.generation || '') && /^[a-f0-9]{64}$/.test(evidence.sha256 || '')
    && /^[A-Za-z0-9+/]{22}==$/.test(evidence.md5Hash || '')
    && Number.isSafeInteger(evidence.byteCount) && evidence.byteCount > 0 && evidence.byteCount <= MAX_BYTES
    && Number.isFinite(Date.parse(evidence.verifiedAt)) && Number.isFinite(Date.parse(evidence.checkedAt)));
}

// Read an existing parent only. Never import, replace, classify or approve media.
// A generation-keyed byte receipt survives refreshes; HEAD revalidates it every time.
export async function inspectExistingHostedVideo(candidate,receipt,
  {db,fetchImpl=fetch,claimDownload=()=>true,now=Date.now}={}) {
  const object = hostedVideoObject(candidate,receipt);
  if (!object) return {status:'not_applicable'};
  const url = receipt.media.videoUrl;
  const head = await fetchImpl(url,{method:'HEAD',redirect:'error',cache:'no-store',signal:AbortSignal.timeout(10000)});
  const observed = observedObject(head);
  if (!observed) return {status:'unverified_hosted_object'};
  const identity = {...object,generation:observed.generation};
  const cached = await db('get_video_hosted_media_verification',identity);
  if (!cached?.ok) throw Error('Existing hosted-media verification ledger unavailable');
  const checkedAt = new Date(now()).toISOString();
  if (cached.record) {
    const evidence = {...cached.record,status:'verified',checkedAt};
    if (!hostedVideoEvidenceMatches(candidate,receipt,evidence)
      || evidence.byteCount !== observed.byteCount || evidence.md5Hash !== observed.md5Hash) {
      return {status:'hosted_verification_conflict'};
    }
    return evidence;
  }
  if (!claimDownload()) return {status:'awaiting_hosted_byte_verification'};
  const response = await fetchImpl(url,{method:'GET',redirect:'error',cache:'no-store',
    headers:head.headers.get('etag') ? {'if-match':head.headers.get('etag')} : {},signal:AbortSignal.timeout(60000)});
  const current = observedObject(response);
  if (!current || Object.entries(observed).some(([key,value]) => current[key] !== value)) {
    await response.body?.cancel();
    return {status:'hosted_object_changed'};
  }
  const md5 = createHash('md5'), sha = createHash('sha256');
  let byteCount = 0;
  for await (const chunk of response.body) {
    byteCount += chunk.length;
    if (byteCount > observed.byteCount) throw Error('Existing hosted object exceeded its declared size');
    md5.update(chunk); sha.update(chunk);
  }
  if (byteCount !== observed.byteCount || md5.digest('base64') !== observed.md5Hash) {
    return {status:'hosted_byte_verification_failed'};
  }
  const record = {...identity,...observed,validation:'full_bytes_md5',sha256:sha.digest('hex'),verifiedAt:new Date(now()).toISOString()};
  const saved = await db('save_video_hosted_media_verification',{record});
  if (!saved?.ok || !saved.record) throw Error('Existing hosted-media byte receipt could not be saved');
  const evidence = {...saved.record,status:'verified',checkedAt:new Date(now()).toISOString()};
  if (!hostedVideoEvidenceMatches(candidate,receipt,evidence)
    || Object.keys(record).filter(key => key !== 'verifiedAt').some(key => evidence[key] !== record[key])) throw Error('Stored hosted-media evidence changed');
  return evidence;
}
