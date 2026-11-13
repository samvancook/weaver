import fs from 'node:fs';import assert from 'node:assert/strict';import path from 'node:path';
const file=path.join(process.argv[2]||'/app','public/index.html');
const before='Opening this link saves no decisions. Sign in to submit approvals.';
const after='Choose your decisions below, then click Submit Decisions to save them. You must be signed in.';
const html=fs.readFileSync(file,'utf8');assert.equal(html.split(before).length,2,'Changed assignment instruction target; no write made');fs.writeFileSync(file,html.replace(before,after));
console.log('Clarified how to save excerpt review decisions.');
