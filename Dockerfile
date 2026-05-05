# 使用 Node.js 22 作為基礎鏡像
FROM node:22-slim

# 設定工作空間
WORKDIR /app

# 複製 package.json 並安裝依賴
COPY package*.json ./
RUN npm install

# 複製所有原始碼
COPY . .

# 編譯前端靜態檔案 (Vite)
RUN npm run build

# 開放連接埠 3000
EXPOSE 3000

# 啟動應用程式
# 注意：Node.js 22 可以直接執行 server.ts (如果不含複雜 Enum，否則建議用 tsx)
CMD [ "npm", "start" ]
