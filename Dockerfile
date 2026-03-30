FROM node:22-slim

WORKDIR /app

COPY package.json ./
RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 \
  && rm -rf /var/lib/apt/lists/*
RUN npm install --omit=dev

COPY public ./public
COPY server.mjs ./server.mjs
COPY catalog_validate.py ./catalog_validate.py
COPY catalog_poem_text.py ./catalog_poem_text.py
COPY excerpt_library.py ./excerpt_library.py
COPY excerpt_library_text.py ./excerpt_library_text.py
COPY data ./data

ENV PORT=8080

EXPOSE 8080

CMD ["npm", "start"]
