import { OAuth2Client } from 'google-auth-library';
const verifier=new OAuth2Client();
export async function verifyFootageHelperCaller(req, config={}, tokenVerifier=verifier) {
  const audience=String(config.audience||'').trim();
  const allowed=new Set(String(config.emails||'').split(',').map(x=>x.trim().toLowerCase()).filter(Boolean));
  const fail=(message,statusCode)=>{throw Object.assign(new Error(message),{statusCode});};
  if(!audience||!allowed.size)fail('footage_helper_not_configured',503);
  const token=String(req.headers.authorization||'').match(/^Bearer\s+([^\s]+)$/i)?.[1];
  if(!token)fail('helper_authentication_required',401);
  let payload;
  try {payload=(await tokenVerifier.verifyIdToken({idToken:token,audience})).getPayload();}
  catch {fail('helper_token_invalid',401);}
  if(payload?.email_verified!==true||!allowed.has(String(payload.email||'').toLowerCase()))fail('helper_access_denied',403);
  return {email:payload.email};
}
