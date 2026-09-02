FROM node:18-bookworm-slim

ARG LIBREOFFICE_VERSION=1:7.4.7-1+deb12u6
ARG POPPLER_UTILS_VERSION=22.12.0-2+deb12u1

RUN set -eux; \
  apt-get update; \
  apt-cache policy libreoffice poppler-utils; \
  install_with_optional_pin() { \
    package_name="$1"; \
    package_version="$2"; \
    if [ -n "$package_version" ] && apt-cache madison "$package_name" | awk '{print $3}' | grep -Fxq "$package_version"; then \
      echo "Installing ${package_name}=${package_version}"; \
      apt-get install -y --no-install-recommends "${package_name}=${package_version}"; \
    else \
      if [ -n "$package_version" ]; then \
        echo "Pinned version ${package_name}=${package_version} is unavailable in current apt repo. Falling back to latest available."; \
      fi; \
      apt-get install -y --no-install-recommends "$package_name"; \
    fi; \
  }; \
  install_with_optional_pin libreoffice "$LIBREOFFICE_VERSION"; \
  install_with_optional_pin poppler-utils "$POPPLER_UTILS_VERSION"; \
  rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci

COPY . .
RUN npm run build

CMD ["node", "dist/server.js"]
