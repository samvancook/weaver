import {verifyFootageHelperCaller} from './footage_helper_auth.mjs';
export const VIDEO_PIPELINE_AUDIENCE='https://weaver.buttonpoetry.com';
export const VIDEO_PIPELINE_SERVICE_ACCOUNT='weaver-runtime@button-weaver-internal.iam.gserviceaccount.com';
export async function verifyVideoPipelineCaller(req,verifier){
  return verifyFootageHelperCaller(req,{audience:VIDEO_PIPELINE_AUDIENCE,emails:VIDEO_PIPELINE_SERVICE_ACCOUNT},verifier);
}
