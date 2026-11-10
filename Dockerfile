FROM us-central1-docker.pkg.dev/button-weaver-internal/cloud-run-source-deploy/weaver@sha256:3ae5a7b265070c45ce95d2c90c9bedcf62eba4589a028114f6eaf4bbf7cd187f

WORKDIR /app

# Preserve the verified unified-review runtime and data. Only the guarded
# legacy identity recovery and its operator controls change in this release.
COPY video_curation_handoff.mjs weaver_runtime_sync.py ./
COPY public/app.js public/index.html ./public/
