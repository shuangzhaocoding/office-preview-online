# 阶段 1：构建前端静态资源
FROM node:latest AS builder

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm config set registry https://mirrors.huaweicloud.com/repository/npm/ \
 && npm config set fetch-retries 5 \
 && npm config set fetch-retry-mintimeout 20000 \
 && npm config set fetch-retry-maxtimeout 120000 \
 && npm config set fetch-timeout 300000 \
 && npm ci

COPY . .
RUN npm run build

# 阶段 2：Nginx 托管静态文件
FROM swr.cn-north-4.myhuaweicloud.com/ddn-k8s/docker.io/nginx:1.27-alpine

COPY docker/nginx/default.conf /etc/nginx/conf.d/default.conf
COPY --from=builder /app/dist /usr/share/nginx/html

EXPOSE 80

CMD ["nginx", "-g", "daemon off;"]
