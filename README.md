# Office Preview

基于 Vue 3 + Vite 的浏览器端 Office 文件预览页，支持通过 URL 参数打开远程 Word / Excel / PDF / PPT 文件。

## 功能

- 预览 **DOCX**、**Excel (xlsx/xls)**、**PDF**、**PPTX**
- 通过查询参数传入文件地址与类型
- 内置 `/file-proxy` 代理，规避跨域下载问题
- 下载与渲染进度提示
- PDF / PPTX 分页、缩略图侧栏与缩放控制

## 技术栈

- Vue 3
- Vite 5
- [@vue-office](https://github.com/501351981/vue-office)（docx / excel / pdf / pptx）

## 环境要求

- Node.js `>= 18`（推荐使用仓库内 `.nvmrc`）

## 快速开始

```bash
npm install
npm run dev
```

开发服务默认：`http://localhost:5173`

其他命令：

```bash
npm run build    # 构建生产包
npm run preview  # 预览构建产物（同样挂载 file-proxy）
```

## 使用方式

在地址栏传入 `url` 与 `resource_type`：

```
http://localhost:5173/?url=<文件地址>&resource_type=<类型>
```

示例：

```
http://localhost:5173/?url=https%3A%2F%2Fexample.com%2Ffile.docx&resource_type=docx
```

参数也可写在 hash 中，例如：

```
http://localhost:5173/#/?url=...&resource_type=pdf
```

### 参数说明

| 参数 | 必填 | 说明 |
| --- | --- | --- |
| `url` | 是 | 文件下载地址（需可被服务端代理访问），或同源相对路径 |
| `resource_type` | 建议填写 | 文件类型；未填时会尝试从 `url` 后缀推断 |

### 支持的 `resource_type`

| 值 | 实际预览类型 |
| --- | --- |
| `docx` / `doc` / `word` | Word |
| `xlsx` / `xls` / `excel` / `spreadsheet` | Excel |
| `pdf` | PDF |
| `pptx` / `ppt` / `powerpoint` | PPT |

## 文件代理

远程 `http(s)` 地址会走本地代理：

```
/file-proxy?url=<encodeURIComponent(原始地址)>
```

同源相对路径、`blob:`、`data:` 地址不会走代理。

代理仅用于开发 / `vite preview` 场景，生产环境若部署到静态托管，需要自行提供等价的反向代理能力。

## 项目结构

```
office-preview/
├── src/
│   ├── App.vue          # 预览主界面与加载逻辑
│   ├── main.js
│   └── style.css
├── vite.config.js       # Vite 配置与插件挂载
├── vite.file-proxy.js   # 远程文件代理中间件
├── vite.pdf-scale.js    # PDF 缩放相关补丁
├── vite.pptx-lazy.js    # PPTX 懒渲染相关补丁
└── package.json
```

## 说明

- 预览能力依赖 `@vue-office` 各格式组件，复杂排版或特殊 Office 特性可能与桌面软件存在差异。
- 请确保目标文件地址可访问；若下载失败或类型不匹配，页面会显示错误提示。
