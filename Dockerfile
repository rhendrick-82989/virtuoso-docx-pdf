FROM node:20-slim

# Install LibreOffice headless + minimal font support
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-writer \
      fonts-liberation \
      fonts-dejavu \
      fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Rebuild font cache
RUN fc-cache -f -v

WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev

COPY server.js ./

# LibreOffice needs a writable HOME; /tmp is always writable
ENV HOME=/tmp

EXPOSE 3001
CMD ["node", "server.js"]
