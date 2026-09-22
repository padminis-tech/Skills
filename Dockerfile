FROM node:20-alpine
WORKDIR /app
COPY server/package.json ./
RUN npm install --production
COPY server/index.js ./
COPY SKILL.md ./SKILL.md
EXPOSE 3000
CMD ["node", "index.js"]
