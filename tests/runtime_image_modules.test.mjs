import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';

test('every relative runtime JavaScript dependency is copied into the image',()=>{
 const root=new URL('../',import.meta.url);
 const docker=readFileSync(new URL('Dockerfile',root),'utf8');
 const copied=new Set([...docker.matchAll(/^COPY (.+)$/gm)].flatMap(m=>m[1].trim().split(/\s+/).slice(0,-1)));
 const visited=new Set();
 function inspect(path) {
  if(visited.has(path))return;visited.add(path);
  assert.ok(copied.has(path),`Runtime dependency ${path} is missing from Docker COPY`);
  const source=readFileSync(new URL(path,root),'utf8');
  for(const m of source.matchAll(/(?:from\s+|import\s*)["']\.\/([^"']+\.mjs)["']/g))inspect(m[1]);
 }
 inspect('server.mjs');
});
