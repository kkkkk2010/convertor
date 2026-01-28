FROM node:20-bookworm

ARG LIBREOFFICE_VERSION=1:7.4.7-1+deb12u5
ARG POPPLER_UTILS_VERSION=22.12.0-2+deb12u2

RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    libreoffice=${LIBREOFFICE_VERSION} \
    poppler-utils=${POPPLER_UTILS_VERSION} \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install

COPY . .
RUN npm run build

ENTRYPOINT ["node", "dist/importer.js"]
