FROM node:20-slim

# Install LibreOffice, poppler-utils (pdftotext), and fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-writer \
      poppler-utils \
      fonts-liberation \
      fonts-dejavu \
      fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

RUN fc-cache -f -v

WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev

COPY server.js ./

ENV HOME=/tmp

EXPOSE 3001
CMD ["node", "server.js"]
