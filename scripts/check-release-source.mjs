#!/usr/bin/env node
// Read-only guard for the existing exact-commit release route. Never deploys.
import {execFileSync} from 'node:child_process';
import {createHash} from 'node:crypto';
import {readFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
import path from 'node:path';

export function assertReleaseSource({commit,head,main,dirty,containsMain}){
  if(!/^[a-f0-9]{40}$/.test(commit||''))throw Error('Use the full 40-character approved commit');
  if(commit!==head)throw Error('Checkout differs from the approved commit');
  if(dirty)throw Error('Checkout has uncommitted or untracked files');
  if(!main||!containsMain)throw Error('Approved source does not contain current main; preserve concurrent work first');
}

export function checkReleaseSource(root,commit){
  const git=(...args)=>execFileSync('git',args,{cwd:root,encoding:'utf8',timeout:30000}).trim();
  const head=git('rev-parse','HEAD');
  // Refresh only the remote-tracking ref. No checkout or source changes.
  git('fetch','origin','main');
  const main=git('rev-parse','origin/main');
  let containsMain=false;
  try{git('merge-base','--is-ancestor',main,head);containsMain=true;}catch{}
  assertReleaseSource({commit,head,main,dirty:Boolean(git('status','--porcelain','--untracked-files=all')),containsMain});
  execFileSync(process.execPath,['--test','tests/runtime_image_modules.test.mjs'],{cwd:root,timeout:30000,stdio:'pipe'});
  // Tracked executable/runtime sources only; no credential, data or excerpt corpus.
  const files=git('ls-files','-z').split('\0').filter(Boolean)
    .filter(f=>/\.(mjs|py|sh)$/.test(f)||f.startsWith('public/')||['Dockerfile','.gcloudignore','.dockerignore','package.json','package-lock.json'].includes(f))
    .filter(f=>!f.startsWith('tests/')&&!f.startsWith('scripts/')&&!f.startsWith('apply-'));
  const manifest=Object.fromEntries(files.sort().map(f=>[f,createHash('sha256').update(readFileSync(path.join(root,f))).digest('hex')]));
  return {ok:true,kind:'source_preflight_only',checkedAt:new Date().toISOString(),commit,currentMain:main,files:manifest,
    manifestSha256:createHash('sha256').update(JSON.stringify(manifest)).digest('hex'),
    limits:'Not a deployment, shared release lock, image readback, canary or main/alias acceptance. Recheck before promotion if main or serving revision changes.'};
}

if(process.argv[1]&&path.resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  try{console.log(JSON.stringify(checkReleaseSource(path.resolve(fileURLToPath(new URL('..',import.meta.url))),process.argv[2]),null,2));}
  catch(error){console.error(JSON.stringify({ok:false,error:error.message}));process.exitCode=1;}
}
