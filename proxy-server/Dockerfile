FROM node:20-alpine
RUN apk add --no-cache poppler-utils tesseract-ocr curl \
  && curl -sL -o /usr/share/tessdata/spa.traineddata \
     https://github.com/tesseract-ocr/tessdata/raw/main/spa.traineddata
WORKDIR /app
COPY package.json ./
RUN npm install
COPY server.js ./
EXPOSE 3456
CMD ["node", "server.js"]
