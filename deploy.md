# 部署说明

## 项目说明
这是一个纯前端的Excel数据处理工具，支持：
- 上传Excel文件（本地处理，不上传服务器）
- 按数量或比例随机保留/删除数据行（首行表头不计入）
- 删除F列及之后所有列（保留A-E共5列）
- 导出处理后的文件

## 技术栈
- HTML + CSS + JavaScript
- SheetJS/xlsx (https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js)

## 部署步骤

### 1. GitHub 部署

```bash
# 初始化Git仓库
git init

# 添加所有文件
git add .

# 提交代码
git commit -m "初始提交：Excel数据处理工具"

# 添加GitHub远程仓库（请替换为你的仓库地址）
git remote add origin https://github.com/YOUR_USERNAME/your-repo-name.git

# 推送到GitHub
git branch -M main
git push -u origin main
```

### 2. Netlify 部署（自动）

1. 登录 [Netlify](https://app.netlify.com/)
2. 点击 "Add new site" → "Import an existing project"
3. 选择 "GitHub" 并授权
4. 选择刚才创建的仓库
5. 配置构建设置：
   - Build command: （留空）
   - Publish directory: （或填写 `.`）
6. 点击 "Deploy site"

### 3. 注意事项

- 这是一个纯静态HTML应用，无需构建步骤
- Netlify会自动从根目录部署
- 部署后会获得一个Netlify域名，可以绑定自定义域名

## 本地预览

在项目目录下运行以下命令启动本地服务器：

```bash
# Python 3
python -m http.server 8000

# Node.js (如果安装了http-server)
npx http-server -p 8000

# 然后在浏览器中访问
# http://localhost:8000
```

## 使用说明

1. 上传Excel文件（.xlsx/.xls）
2. 选择操作类型：
   - 保留行：按数量或比例保留指定行数
   - 删除行：按数量或比例删除指定行数
3. （可选）点击"删除F列及之后列"按钮
4. 点击"下载处理后的文件"导出结果

## 功能特性

- ✅ 完全本地处理，数据不上传服务器
- ✅ 首行表头不计入数量计算
- ✅ 保留/删除操作互斥
- ✅ 必须执行行操作才能下载
- ✅ 精美的暗色主题界面
