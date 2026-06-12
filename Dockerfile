FROM node:20-slim

# Enable contrib repo (Debian bookworm, deb822 format) for ttf-mscorefonts-installer,
# then install LibreOffice, poppler-utils, fonts.
# Microsoft core fonts (real Times New Roman + Arial) make the PDF rendering
# visually identical to Word. Install is best-effort (|| true): if the font
# download ever fails, the build still succeeds and LibreOffice falls back to
# Liberation Serif/Sans, which are metrically identical (same pagination).
RUN sed -i 's/Components: main/Components: main contrib/' /etc/apt/sources.list.d/debian.sources && \
    apt-get update && \
    echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-writer \
      poppler-utils \
      fonts-liberation \
      fonts-dejavu \
      fontconfig \
      ca-certificates \
      wget \
    && (apt-get install -y --no-install-recommends ttf-mscorefonts-installer || \
        echo "WARN: mscorefonts install failed — falling back to Liberation fonts") \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

RUN fc-cache -f -v && \
    (fc-list | grep -i "times new roman" && echo "Times New Roman: INSTALLED") || \
    echo "Times New Roman: not found, Liberation Serif will substitute"

WORKDIR /app
COPY package.json ./
RUN npm install --omit=dev

COPY server.js ./

ENV HOME=/tmp

EXPOSE 3001
CMD ["node", "server.js"]
