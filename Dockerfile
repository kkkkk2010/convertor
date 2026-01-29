FROM node:18-bookworm-slim

ARG LIBREOFFICE_VERSION=1:7.4.7-1+deb12u6
ARG POPPLER_UTILS_VERSION=22.12.0-2+deb12u1

RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    libreoffice=${LIBREOFFICE_VERSION} \
    poppler-utils=${POPPLER_UTILS_VERSION} \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .
RUN npm run build

CMD ["node", "dist/importer.js"]
