# 图片信息提取工具

纯前端方案，直接调用 GLM-4.5V API，支持拍照上传和 Excel 导出。

## 功能特性

- 📷 拍照或选择图片上传
- 🤖 基于 GLM-4.5V 智能提取结构化信息
- 📝 可编辑的表单界面
- 💾 一键保存到 Excel 文件
- 📱 移动端友好设计
- 🌓 支持深色模式

## 快速开始

### 1. 安装依赖

```bash
bun install
```

### 2. 配置 API Key

在浏览器中首次使用时，会提示输入智谱 GLM-4.5V 的 API Key。

API Key 获取地址：https://open.bigmodel.cn

### 3. 启动服务

```bash
bun run dev
```

或者使用 Bun 作为静态文件服务器：

```bash
bun --hot src/app.js
```

### 4. 使用

1. 打开浏览器访问 `http://localhost:3000`
2. 点击拍照或选择图片
3. 等待 GLM-4.5V 分析图片
4. 编辑提取的信息
5. 点击保存，Excel 文件会自动下载

## 项目结构

```
img-analyze/
├── index.html          # 主页面
├── package.json        # 项目配置
├── bunfig.toml         # Bun 配置
├── src/
│   ├── app.js         # 主逻辑
│   └── style.css      # 样式
└── README.md
```

## 技术栈

- **运行时**: Bun
- **AI 模型**: GLM-4.5V (智谱)
- **Excel 处理**: SheetJS (xlsx)
- **样式**: 原生 CSS (移动端优化)

## 注意事项

- API Key 保存在浏览器本地存储中
- Excel 数据保存在浏览器本地存储中（会话间持久化）
- 建议定期备份重要的 Excel 数据
