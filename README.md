# 医疗图片信息录入工具

基于 Bun + 原生前端的医疗信息录入系统：上传医疗图片后，调用智谱视觉模型自动提取患者信息，人工校对后保存，并支持按日期导出 Excel。

## 功能特性

- 📷 支持拍照/上传图片（最多 10 张）
- 🤖 调用智谱 `glm-4.6v` 进行多图信息提取
- 📝 自动回填医疗表单，支持人工补充和修正
- 🧮 自动计算 BMI、GERDQ 总分、吸氧浓度、氧合指数、TCM 总分
- 💾 记录保存到本地文件（`data/records.json`）
- 📤 支持按日期导出 Excel
- 📱 移动端友好，支持深色模式

## 快速开始

### 1) 安装依赖

```bash
bun install
```

### 2) 启动服务

```bash
bun run dev
```

默认访问地址：`http://localhost:3030`

> 端口可通过环境变量 `PORT` 覆盖，例如：`PORT=8080 bun run dev`

### 3) 配置 API Key

首次进入页面后，点击右上角设置按钮，填写智谱 API Key（保存在浏览器本地存储）。

API Key 获取：[https://open.bigmodel.cn](https://open.bigmodel.cn)

### 4) 使用流程

1. 上传或拍摄医疗图片
2. 点击“开始分析”
3. 检查并编辑自动提取结果
4. 点击“保存”写入本地记录
5. 点击右上角下载按钮按日期导出 Excel

## 数据与存储

- 记录持久化位置：`data/records.json`
- API Key 存储位置：浏览器 `localStorage`（键名：`zhipu_api_key`）
- 导出由前端生成 `.xlsx` 文件并下载

## 主要接口（本地）

- `POST /api/save`：保存或更新记录
- `GET /api/records`：获取全部记录
- `POST /api/records`：按日期筛选记录
- `POST /api/export`：按日期导出数据（返回 JSON，由前端转 Excel）
- `GET /api/dates`：获取可用日期列表

## 项目结构

```text
img-analyze/
├── index.html                    # 主页面（表单与交互结构）
├── src/
│   ├── app.js                    # 前端业务逻辑（上传、识别、计算、保存、导出）
│   └── style.css                 # 样式（移动端 + 深色模式）
├── server.js                     # Bun 服务与本地 API
├── data/records.json             # 本地数据文件
├── package.json
├── Dockerfile
└── README.md
```

## 技术栈

- **运行时**：Bun
- **后端**：Bun HTTP Server（`server.js`）
- **前端**：原生 HTML/CSS/JavaScript
- **AI 模型**：智谱 `glm-4.6v`
- **Excel**：SheetJS (`xlsx`)

## Docker（可选）

```bash
docker build -t img-analyze .
docker run --rm -p 3030:3030 img-analyze
```
