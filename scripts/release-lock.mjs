#!/usr/bin/env node
// Coordinates the existing Weaver release routes. No build, deploy or IAM calls.
import fs from 'node:fs';import path from 'node:path';import os from 'node:os';
import {fileURLToPath} from 'node:url';
import {checkReleaseSource} from './check-release-source.mjs';
import {acquireRelease,assertOwnedRelease,assertReleaseBase,planRenewRelease,planCloseRelease} from './release-lock-core.mjs';

const doc='https://firestore.googleapis.com/v1/projects/button-weaver-internal/databases/weaverledger/documents/releaseControl/weaver';
const run='https://us-central1-run.googleapis.com/apis/serving.knative.dev/v1/namespaces/button-weaver-internal';
const encode=v=>v===null?{nullValue:null}:typeof v==='object'?{mapValue:{fields:Object.fromEntries(Object.entries(v).map(([k,x])=>[k,encode(x)]))}}:{stringValue:String(v)};
const decode=v=>'mapValue'in v?Object.fromEntries(Object.entries(v.mapValue.fields||{}).map(([k,x])=>[k,decode(x)])):'nullValue'in v?null:v.stringValue;
async function main(){
  const [mode,...args]=process.argv.slice(2),options={};
  for(let i=0;i<args.length;i+=2){if(!args[i].startsWith('--')||!args[i+1])throw Error('Use --option value pairs');options[args[i].slice(2)]=args[i+1];}
  if(!['status','acquire','check','renew','finish','abort'].includes(mode))throw Error('Use status, acquire, check, renew, finish or abort');
  if(mode!=='status'&&(!options.receipt||!path.isAbsolute(options.receipt)))throw Error('Absolute private --receipt path required');
  // Same existing Sam/clasp Google connection used by the established release tools.
  // Reading this file never launches login, changes scopes or prints tokens.
  const {OAuth2Client}=await import('google-auth-library');
  const credentials=JSON.parse(fs.readFileSync(path.join(os.homedir(),'.clasprc.json'))).tokens?.default;
  if(!credentials?.client_id||!credentials?.refresh_token)throw Error('Existing Google/clasp connection unavailable; no new login attempted');
  const auth=new OAuth2Client(credentials.client_id,credentials.client_secret);auth.setCredentials(credentials);
  const request=options=>auth.request({...options,headers:{"x-goog-user-project":"button-weaver-internal",...(options.headers||{})}});
  const store={read:async()=>{try{const{data:d}=await request({url:doc});return{version:d.updateTime,value:decode({mapValue:{fields:d.fields}})};}catch(e){if(e.response?.status===404)return null;throw e;}},
    write:async(value,version)=>{const params=version?{'currentDocument.updateTime':version}:{'currentDocument.exists':false};try{await request({url:doc,method:'PATCH',params,data:{fields:encode(value).mapValue.fields}});}catch(e){if([409,412].includes(e.response?.status))throw Error('Release claim changed concurrently; no lock acquired or cleared');throw e;}}};
  if(mode==='status'){console.log(JSON.stringify({ok:true,lock:(await store.read())?.value||null}));return;}
  const service=async()=>{const{data:s}=await request({url:run+'/services/weaver'});const revision=s.status.traffic.find(t=>t.percent===100)?.revisionName;if(!revision)throw Error('Expected one 100% serving Weaver revision');const{data:r}=await request({url:run+'/revisions/'+revision});if(r.status.conditions.find(c=>c.type==='Ready')?.status!=='True')throw Error('Serving revision is not Ready');return{revision,image:r.spec.containers[0].image,aliases:Object.fromEntries(s.status.traffic.filter(t=>t.tag).map(t=>[t.tag,t.revisionName]))};};
  const sourceRoot=path.resolve(options['source-root']||fileURLToPath(new URL('..',import.meta.url)));
  if(mode==='acquire'){
    if(fs.existsSync(options.receipt))throw Error('Receipt path already exists; preserve the previous claim');
    // Reserve the local receipt before any cloud write so a local I/O failure cannot lose a claim.
    fs.writeFileSync(options.receipt,'{}\n',{flag:'wx',mode:0o600});
    const source=checkReleaseSource(sourceRoot,options.commit),base=await service();
    const claim=await acquireRelease(store,{commit:source.commit,owner:options.owner,manifestSha256:source.manifestSha256,baseRevision:base.revision,baseImage:base.image,baseAliases:base.aliases},Date.now(),
      planned=>fs.writeFileSync(options.receipt,JSON.stringify(planned,null,2)+'\n',{mode:0o600}));
    fs.writeFileSync(options.receipt,JSON.stringify(claim,null,2)+'\n',{mode:0o600});console.log(JSON.stringify({ok:true,status:'held',owner:claim.owner,commit:claim.commit,baseRevision:claim.baseRevision,receipt:options.receipt}));return;
  }
  const receipt=JSON.parse(fs.readFileSync(options.receipt));const observed=await store.read();if(!observed)throw Error('Release lock document is missing');const current=observed.value,serving=await service();
  if(mode==='check'||mode==='renew'){
    assertOwnedRelease(current,receipt);assertReleaseBase(current,serving);
    const source=checkReleaseSource(sourceRoot,current.commit);if(source.manifestSha256!==current.manifestSha256)throw Error('Release source changed');
    if(mode==='renew')await store.write(planRenewRelease(current,receipt,serving),observed.version);
    console.log(JSON.stringify({ok:true,status:'held',commit:current.commit,baseRevision:serving.revision}));return;
  }
  let activeBuilds=0;for(const location of ['us-central1','global']){const{data:b}=await request({url:'https://cloudbuild.googleapis.com/v1/projects/button-weaver-internal/locations/'+location+'/builds',params:{filter:'status!=SUCCESS AND status!=FAILURE AND status!=INTERNAL_ERROR AND status!=CANCELLED AND status!=TIMEOUT AND status!=EXPIRED',pageSize:100}});activeBuilds+=(b.builds||[]).length;if(b.nextPageToken)throw Error('Active-build scan incomplete; keep the lock');}
  if(mode==='finish'&&(!options.verification||!path.isAbsolute(options.verification)))throw Error('Absolute private --verification path required');
  const verification=mode==='finish'?JSON.parse(fs.readFileSync(options.verification)):null;
  const next=planCloseRelease(current,receipt,{outcome:mode==='finish'?'released':'aborted',service:serving,verification,activeBuilds});
  await store.write(next,observed.version);console.log(JSON.stringify({ok:true,status:next.status,revision:serving.revision}));
}
main().catch(e=>{console.error(JSON.stringify({ok:false,error:e.response?.data?.error?.message||e.message}));process.exitCode=1;});
