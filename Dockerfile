FROM node:22-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install --production
COPY app.js ./
ENV PORT=7860
EXPOSE 7860
CMD ["node", "app.js"]