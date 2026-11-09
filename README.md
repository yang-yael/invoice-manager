# 发票管理系统 / Invoice Manager

把一堆杂乱的 PDF、图片发票，自动识别票面信息、按报销人和购买方公司分组，生成规范的月度报销表。

本地拖拽上传即可使用，也可以用 Docker 部署到服务器，多人共用。

## 功能

- **批量识别**：PDF 直接提取文字层；扫描件和图片走 OCR（poppler + tesseract 中文）。
- **自动分组**：按报销人分组（文件夹名 / ZIP 文件名），再按票面购买方公司二次分组。
- **汇总表模式**：给一张含发票号码的汇总表 Excel + 一堆以发票号命名的 PDF，系统自动配对。
- **生成报销表**：套用 `发票模板.xlsx` 输出月度发票管理表，发票原件同步归档到输出目录。
- **零重型依赖**：Web 服务基于 Python 标准库 `http.server`，运行时只需 `openpyxl`。

## 技术栈

| 层 | 选型 |
|---|---|
| Web 服务 | Python 3.12 标准库 `ThreadingHTTPServer` |
| Excel 读写 | openpyxl |
| PDF 文字层 | poppler-utils (`pdftotext`) |
| OCR | tesseract-ocr + `chi_sim` 中文语言包 |
| 部署 | Docker / docker compose + nginx 反向代理 |

## 目录结构

```text
.
├── invoice_system/
│   ├── web_app.py          # Web 服务：上传、识别、下载
│   ├── invoice_manager.py  # 核心逻辑：解析、分组、写 Excel
│   └── README.md           # 本地使用与格式说明
├── windows_package/        # Windows 免安装分发包
├── 发票模板.xlsx            # 空白报销表模板
├── Dockerfile
├── docker-compose.yml
├── render.yaml             # Render.com 一键部署
└── 阿里云宝塔部署说明.md     # 自有服务器部署步骤
```

## 本地运行

```bash
pip install -r requirements.txt
python3 invoice_system/web_app.py
```

打开 http://127.0.0.1:8765 即可拖拽上传。

想要 OCR 支持扫描件和图片，macOS 上额外装：

```bash
brew install poppler tesseract tesseract-lang
```

## Docker 部署

```bash
cp .env.example .env   # 改掉 INVOICE_PASSWORD
docker compose up -d --build
```

服务监听 `8765`，云端模式下启用 HTTP Basic 认证。生产环境请用 nginx 反向代理 + HTTPS，不要直接暴露端口。

完整服务器部署步骤见 [阿里云宝塔部署说明.md](阿里云宝塔部署说明.md)。

## 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `HOST` | `127.0.0.1` | 监听地址，容器内为 `0.0.0.0` |
| `PORT` | `8765` | 监听端口 |
| `INVOICE_CLOUD` | 空 | 设为 `1` 开启云端模式（强制登录、不自动开浏览器） |
| `INVOICE_USER` | `admin` | 登录用户名 |
| `INVOICE_PASSWORD` | 空 | 登录密码，云端模式必填 |
| `INVOICE_DATA_DIR` | 项目目录 | 上传与输出的持久化目录 |
| `INVOICE_TEMPLATE` | `发票模板.xlsx` | 报销表模板路径 |

## 隐私说明

`uploads/`、`outputs/`、`data/`、`.env` 以及所有 PDF / 图片 / xlsx 报销表都在 `.gitignore` 中，**真实发票数据不会进入这个仓库**。仓库里唯一的 Excel 是空白模板 `发票模板.xlsx`。

## License

MIT
