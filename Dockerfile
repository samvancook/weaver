FROM node:22-slim

WORKDIR /app

COPY package.json ./
RUN apt-get update \
  && apt-get install -y --no-install-recommends ca-certificates python3 \
  && rm -rf /var/lib/apt/lists/*
RUN npm install --omit=dev

COPY public ./public
COPY server.mjs ./server.mjs
COPY footage_event_intake.mjs footage_event_links.mjs footage_event_progress.mjs footage_helper_auth.mjs ./
COPY book_intake_claim.mjs ./book_intake_claim.mjs
COPY book_intake_readiness.mjs ./book_intake_readiness.mjs
COPY weaver_progress_auth.mjs ./weaver_progress_auth.mjs
COPY video_curation_handoff.mjs ./video_curation_handoff.mjs
COPY footage_provenance.mjs ./footage_provenance.mjs
COPY video_excerpt_folding.mjs ./video_excerpt_folding.mjs
COPY poem_key_seed.mjs ./poem_key_seed.mjs
COPY catalog_validate.py ./catalog_validate.py
COPY catalog_live.py ./catalog_live.py
COPY catalog_poem_text.py ./catalog_poem_text.py
COPY excerpt_library.py ./excerpt_library.py
COPY excerpt_library_text.py ./excerpt_library_text.py
COPY intake_catalog_lookup.py ./intake_catalog_lookup.py
COPY import_main_catalog_book.py ./import_main_catalog_book.py
COPY prepare_catalog_overlay.py ./prepare_catalog_overlay.py
COPY start-weaver.sh ./start-weaver.sh
COPY graphics_qi_lookup.py ./graphics_qi_lookup.py
COPY weaver_runtime_db.py ./weaver_runtime_db.py
COPY weaver_runtime_sync.py ./weaver_runtime_sync.py
COPY db ./db
COPY data ./data

ENV PORT=8080

EXPOSE 8080

CMD ["sh", "./start-weaver.sh"]
