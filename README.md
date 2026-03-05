# PDF 转 Word Chrome 插件

一款简单易用的 Chrome 浏览器插件，点击图标后在页面右侧展示面板，支持将 PDF 文件转换为 Word (.docx) 格式。

## 功能特点

- 点击插件图标，右侧面板滑出显示
- 支持点击上传或拖拽上传 PDF 文件
- 纯前端处理，无需后端服务器
- 保留基本的文本格式和布局
- 转换进度实时显示

## 安装方法

### 开发者模式安装

1. 打开 Chrome 浏览器，访问 `chrome://extensions/`
2. 开启右上角的「开发者模式」开关
3. 点击「加载已解压的扩展程序」按钮
4. 选择本插件所在的文件夹 `pdf转word`
5. 插件图标会显示在浏览器工具栏中

## 使用方法

1. 访问任意网页
2. 点击浏览器工具栏中的插件图标
3. 右侧面板会从右侧滑出
4. 点击上传区域选择 PDF 文件，或将文件拖拽到上传区域
5. 点击「开始转换」按钮
6. 等待转换完成，Word 文件会自动下载

## 文件结构

```
pdf转word/
├── manifest.json          # 插件配置文件
├── background.js          # 后台脚本，处理图标点击事件
├── content.js             # 内容脚本，渲染右侧面板
├── pdf-processor.js       # PDF 解析和 Word 生成核心模块
├── icons/                 # 图标文件
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
├── pdf.min.js             # PDF.js 库（需自行下载）
├── pdf.worker.min.js      # PDF.js Worker（需自行下载）
├── docx.min.js            # docx.js 库（需自行下载）
├── tesseract.min.js       # Tesseract.js 库（需自行下载）
├── tesseract.worker.min.js # Tesseract.js Worker（需自行下载）
└── README.md              # 说明文档
```

## 依赖库下载

由于库文件较大，请自行下载以下文件并放入项目根目录：

| 文件 | 下载地址 |
|------|----------|
| `pdf.min.js` | https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js |
| `pdf.worker.min.js` | https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js |
| `docx.min.js` | https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.min.js （重命名为 docx.min.js）|
| `tesseract.min.js` | https://cdn.jsdelivr.net/npm/tesseract.js@5.0.3/dist/tesseract.min.js |
| `tesseract.worker.min.js` | https://cdn.jsdelivr.net/npm/tesseract.js@5.0.3/dist/worker.min.js （重命名为 tesseract.worker.min.js）|

或使用 npm 安装：
```bash
npm install pdfjs-dist docx tesseract.js
```

## 技术栈

- **pdf.js** - PDF 文件解析
- **docx.js** - Word 文档生成
- **Chrome Extension API** - 浏览器扩展功能
- **Shadow DOM** - 样式隔离

## 注意事项

- 插件需要访问网页内容的权限才能显示面板
- 首次使用需要联网加载 pdf.js 和 docx.js 库
- 转换效果取决于 PDF 文件的复杂度，复杂格式可能需要手动调整

## 更新日志

### v1.0.0
- 初始版本发布
- 支持 PDF 转 Word 基本功能
