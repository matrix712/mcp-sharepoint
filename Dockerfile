FROM node:22-bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y \
    python3 \
    make \
    g++ \
    libcairo2-dev \
    libpango1.0-dev \
    libjpeg-dev \
    libgif-dev \
    librsvg2-dev \
    pkg-config \
    ca-certificates \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY . .

RUN rm -f package-lock.json

RUN npm install

RUN npm install pdf-parse@2.3.12 --save-exact

RUN npm install -g mcp-proxy

EXPOSE 8000

CMD ["mcp-proxy","--port","8000","--sseEndpoint","/sse","--","node","wrapper.cjs"]
