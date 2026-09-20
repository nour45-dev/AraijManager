FROM node:20-alpine

WORKDIR /app

# Copy dependency files
COPY package.json ./

# Install dependencies cleanly
RUN npm install --production

# Copy all application source code
COPY . .

# Expose server port
EXPOSE 3000

# Set environment variables
ENV NODE_ENV=production
ENV PORT=3000

# Start server
CMD ["node", "server.js"]