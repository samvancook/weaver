import test from 'node:test';
import assert from 'node:assert/strict';
import {assertReleaseSource} from '../scripts/check-release-source.mjs';
const good={commit:'a'.repeat(40),head:'a'.repeat(40),main:'b'.repeat(40),dirty:false,containsMain:true};
test('release source rejects stale, dirty or unpinned checkouts before release',()=>{
  assert.doesNotThrow(()=>assertReleaseSource(good));
  for(const patch of [{commit:'main'},{head:'c'.repeat(40)},{dirty:true},{main:''},{containsMain:false}]){
    assert.throws(()=>assertReleaseSource({...good,...patch}));
  }
});
