#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="button-weaver-internal"
REGION="us-central1"
SERVICE="weaver"
DEPLOY_ACCOUNT="weaver-deployer@button-weaver-internal.iam.gserviceaccount.com"
: "${WEAVER_RELEASE_RECEIPT:?Acquire the shared release lock first; set WEAVER_RELEASE_RECEIPT to its private receipt path}"
node scripts/release-lock.mjs check --receipt "${WEAVER_RELEASE_RECEIPT}"
RELEASE_SUFFIX="$(node -e 'const r=JSON.parse(require("fs").readFileSync(process.argv[1]));process.stdout.write("r"+r.commit.slice(0,10)+"-"+r.claimId.slice(0,8))' "${WEAVER_RELEASE_RECEIPT}")"

echo "Deploying ${SERVICE} to project ${PROJECT_ID} in ${REGION}..."
gcloud run deploy "${SERVICE}" \
  --source . \
  --no-traffic \
  --revision-suffix "${RELEASE_SUFFIX}" \
  --memory 1Gi \
  --concurrency 2 \
  --region "${REGION}" \
  --project "${PROJECT_ID}" \
  --account "${DEPLOY_ACCOUNT}"

node scripts/release-lock.mjs check --receipt "${WEAVER_RELEASE_RECEIPT}"
echo "Built canary ${SERVICE}-${RELEASE_SUFFIX}; main traffic is unchanged."
echo "Keep the shared lock through verification and exact-revision promotion; never route to an unspecified latest revision."
