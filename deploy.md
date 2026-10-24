# 部署说明

## 项目说明
这是一个纯前端的Excel数据处理工具，支持：
- 多文件上传与批量处理（一次选择/拖入多个文件）
- 操作范围切换：当前文件 / 全部文件（批量）
- Data100ID 匹配过滤（保留匹配 / 删除匹配，支持切换匹配列 A–E）
- 结算比例计算与按目标比例反推交付量
- 各文件明细表与合计汇总
- 按数量或比例随机保留/删除数据行（首行表头不计入）
- 删除F列及之后所有列（保留A-E共5列）
- 导出：单文件下载 / 多文件打包 ZIP

## 技术栈
- HTML + CSS + JavaScript
- SheetJS/xlsx (https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js)
- JSZip (https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js)

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

1. 上传Excel文件（.xlsx/.xls，可一次选择或拖入多个文件）
2. 选择操作范围：「🎯 当前文件」或「📚 全部文件（批量）」
3. Data100ID 匹配过滤：粘贴ID列表 → 选择保留/删除模式 → 选匹配列 → 执行过滤
4. 结算比例面板会显示总合格/总完成/结算比例，批量模式下另有各文件明细表
5. （可选）输入目标结算比例并「按目标比例过滤」，可反复调整或一键还原
6. （可选）「保留行」/「删除行」/「从F列开始删除」，均支持批量
7. 导出：「⬇️ 导出当前文件」或「📦 批量导出（ZIP）」

## 功能特性

- ✅ 完全本地处理，数据不上传服务器
- ✅ 多文件批量上传与处理
- ✅ 首行表头不计入数量计算
- ✅ 保留/删除操作互斥（新操作替换旧操作，不叠加）
- ✅ 各文件独立计算结算比例，明细表含合计行
- ✅ 多文件打包 ZIP 导出，规避浏览器多文件下载拦截
- ✅ 精美的暗色主题界面
