import test from 'node:test';
import assert from 'node:assert/strict';
import {verifyFootageHelperCaller} from '../footage_helper_auth.mjs';
const config={audience:'exact-helper-client',emails:'owner@buttonpoetry.com'};
const req={headers:{authorization:'Bearer signed.token.value'}};
const verifier=payload=>({async verifyIdToken({audience}){assert.equal(audience,config.audience);return {getPayload:()=>payload};}});
test('Helper trust fails closed when disabled, anonymous, wrong audience/signature or wrong owner',async()=>{
 await assert.rejects(verifyFootageHelperCaller(req,{}),e=>e.statusCode===503);
 await assert.rejects(verifyFootageHelperCaller({headers:{}},config),e=>e.statusCode===401);
 await assert.rejects(verifyFootageHelperCaller(req,config,{async verifyIdToken(){throw Error('audience mismatch');}}),e=>e.statusCode===401);
 for(const payload of [{email:'other@buttonpoetry.com',email_verified:true},{email:config.emails,email_verified:false}]){
  await assert.rejects(verifyFootageHelperCaller(req,config,verifier(payload)),e=>e.statusCode===403);
 }
 assert.equal((await verifyFootageHelperCaller(req,config,verifier({email:config.emails,email_verified:true}))).email,config.emails);
});
