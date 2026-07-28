#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="button-weaver-internal"
REGION="us-central1"
SERVICE="weaver"
DEPLOY_ACCOUNT="weaver-deployer@button-weaver-internal.iam.gserviceaccount.com"

echo "Deploying ${SERVICE} to project ${PROJECT_ID} in ${REGION}..."
gcloud run deploy "${SERVICE}" \
  --source . \
  --region "${REGION}" \
  --project "${PROJECT_ID}" \
  --account "${DEPLOY_ACCOUNT}"
