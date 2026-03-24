# Excel 处理工具 - GitHub + Netlify 部署指南

## 项目概述

这是一个纯前端的Excel数据处理工具，支持：
- ✅ 上传Excel文件（本地处理，不上传服务器）
- ✅ 按数量或比例随机保留/删除数据行（首行表头不计入）
- ✅ 删除F列及之后所有列（保留A-E共5列）
- ✅ 导出处理后的文件

## 技术栈

- **HTML5 + CSS3 + JavaScript** (原生开发)
- **SheetJS/xlsx** (Excel处理库)
- **零后端依赖** (完全客户端处理)

---

## 🚀 快速开始 - 部署到 GitHub + Netlify

### 第一步：推送到 GitHub

1. **初始化 Git 仓库**（如果还没有）
   ```bash
   cd d:/software/codebuddy-int/Codex
   git init
   git add .
   git commit -m "Initial commit: Excel processing tool"
   ```

2. **创建 GitHub 仓库**
   - 访问 https://github.com/new
   - 输入仓库名称（如：`excel-processor`）
   - 选择 **Public** 或 **Private**
   - 点击 **Create repository**

3. **推送到 GitHub**
   ```bash
   git remote add origin https://github.com/YOUR_USERNAME/excel-processor.git
   git branch -M main
   git push -u origin main
   ```
   > 📝 将 `YOUR_USERNAME` 替换为你的 GitHub 用户名

---

### 第二步：部署到 Netlify

#### 方法一：Netlify 自动部署（推荐）

1. **登录 Netlify**
   - 访问 https://app.netlify.com
   - 使用 GitHub 账号登录

2. **导入 GitHub 仓库**
   - 点击 "Add new site" → "Import an existing project"
   - 选择 `excel-processor` 仓库
   - 点击 "Import site"

3. **配置构建设置**
   - **Build command**: （留空）
   - **Publish directory**: （留空，默认根目录）
   - 点击 "Deploy site"

4. **完成部署**
   - Netlify 会自动构建并部署
   - 部署成功后，你会获得一个类似 `https://random-name.netlify.app` 的 URL
   - 可以在 "Domain settings" 中修改为自定义域名

---

#### 方法二：手动拖拽部署（快速测试）

1. **准备文件夹**
   - 确保文件夹中包含 `excel-processor.html`

2. **上传到 Netlify**
   - 登录 Netlify
   - 拖拽整个文件夹到 Netlify 部署区域
   - 等待部署完成（几秒钟）

---

## 📁 项目文件结构

```
d:/software/codebuddy-int/Codex/
├── excel-processor.html       # 主应用文件
├── excel-processor.html.backup # 备份文件
├── process_ratecard_v3.py      # Python 脚本（可选）
├── README.md                  # 项目说明（本文件）
├── .gitignore                 # Git 忽略配置
└── .workbuddy/                # 工作区配置（不需要提交）
```

---

## ⚙️ 高级配置

### 自定义域名

1. **在 Netlify 中**
   - 进入 "Domain settings"
   - 点击 "Add custom domain"
   - 输入你的域名（如：`excel-tool.yourdomain.com`）

2. **DNS 配置**
   - 按照提示添加 DNS 记录
   - 等待 DNS 生效（通常 5-30 分钟）

### 环境变量（如果需要）

- 如果后续需要添加 API 密钥或配置
- 在 Netlify 的 "Environment variables" 中添加
- 在代码中使用 `process.env.VARIABLE_NAME` 访问

---

## 🎨 本地开发

如果需要修改功能，可以在本地测试后再部署：

```bash
# 方法 1: Python HTTP 服务器
cd d:/software/codebuddy-int/Codex
python -m http.server 8000

# 方法 2: Node.js http-server（需要安装）
npx http-server -p 8000

# 方法 3: VS Code Live Server 插件
# 安装后右键 "Open with Live Server"
```

然后在浏览器访问 `http://localhost:8000/excel-processor.html`

---

## 🔧 常见问题

### Q1: 部署后无法访问？
- 检查 Netlify 部署日志
- 确保文件名正确（`excel-processor.html`）
- 等待 1-2 分钟让 DNS 生效

### Q2: 如何更新网站？
- 修改本地文件
- 推送到 GitHub
- Netlify 会自动重新部署

### Q3: 需要构建步骤吗？
- 不需要！这是一个纯静态 HTML 文件
- 所有功能都在浏览器端运行

### Q4: 如何删除部署的网站？
- 登录 Netlify → 选择网站 → Site settings → Delete site

---

## 📊 性能优化建议

1. **压缩资源**
   - 使用图片压缩工具
   - 压缩 JavaScript 和 CSS

2. **启用 CDN**
   - Netlify 自动提供全球 CDN
   - 访问速度已经很快

3. **缓存策略**
   - Netlify 默认启用缓存
   - 可以在 `netlify.toml` 中自定义

---

## 📄 许可证

MIT License - 可自由使用和修改

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

## 📞 联系方式

如有问题，请：
1. 提交 GitHub Issue
2. 或联系开发者

---

**部署日期**：2026-03-24
**最后更新**：2026-03-24
**版本**：v1.0.0

---

*本文档由 WorkBuddy AI 助手自动生成*
