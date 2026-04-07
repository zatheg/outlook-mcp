FROM node:20-slim

WORKDIR /app

# Copy package files and install production deps only
COPY package*.json ./
RUN npm ci --omit=dev

# Copy application code
COPY . .

# Cloud Run writes to /tmp, set HOME so token path resolves correctly
ENV HOME=/tmp
ENV NODE_ENV=production
ENV PORT=8080

EXPOSE 8080

CMD ["node", "server-http.js"]
