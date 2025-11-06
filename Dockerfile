# Dockerfile
FROM node:18-slim

ENV TZ=America/Sao_Paulo \
    PUPPETEER_SKIP_DOWNLOAD=true

# Dependências necessárias para o Chromium usado pelo whatsapp-web.js
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    chromium-sandbox \
    fonts-liberation \
    libappindicator3-1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libgconf-2-4 \
    libgdk-pixbuf2.0-0 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --only=production

# PM2 em execução em foreground (pm2-runtime já faz isso)
RUN npm install -g pm2

COPY . .

CMD ["pm2-runtime", "start", "relatorio.js", "--name", "relatorios-operacoes"]