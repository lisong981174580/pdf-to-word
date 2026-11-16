# 纸间 · PDF 转 Word

可自托管的 PDF → DOCX 网页工具，中文响应式界面，拖拽上传、转换、下载完整闭环。

## 功能

- 使用 pdf2docx 在 Python 后端实际转换，尽可能保留文字、图片和表格。
- 默认最大 20 MB、100 页；文件格式、密码保护和扫描件检测。
- 每次转换独立进程，120 秒超时，最多同时转换 2 个文件（每个服务进程）。
- 临时文件在响应发送完成或出错后清理；不上传到第三方转换平台。
- 文件路径使用服务端随机目录，不使用用户传入的文件名。

## Docker 启动

```bash
docker compose up --build -d
```

打开 http://localhost:8000 。默认端口仅绑定本机。

## Python 本地启动（Python 3.12）

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Windows 激活命令：`.venv\Scripts\activate`。

## 测试

```bash
pip install -r requirements-dev.txt
python -m pytest -q
```

测试会实际生成 PDF、调用转换接口并检查 DOCX 中的可编辑文字，同时覆盖临时文件清理、错误文件、扫描件、加密文件、大小/页数限制和超时。

## API

`POST /api/convert`，使用 multipart/form-data，字段名 `file`。成功返回 DOCX 二进制；失败返回 JSON `detail`。

```bash
curl -f http://localhost:8000/api/convert -F 'file=@example.pdf' -o example.docx
```

`GET /health`：健康检查。`GET /api/config`：上传大小和页数限制。

## 配置

| 环境变量 | 默认值 | 含义 |
| --- | --- | --- |
| MAX_UPLOAD_MB | 20 | 文件大小上限（MB） |
| MAX_PAGES | 100 | PDF 页数上限 |
| CONVERSION_TIMEOUT | 120 | 转换子进程超时（秒） |

## 使用边界

扫描 PDF 暂不提供 OCR；混合文字与扫描页的文件，仅有文字层的内容保证可编辑。复杂多栏、公式和特殊字体可能需要人工调整，不保证像素级还原。

当前版本适合个人使用和受信任的内网环境。面向公网提供服务前，需要增加身份认证、反向代理请求体限制、请求速率限制、任务队列和容器隔离。应用层限制在 multipart 解析后执行，公网入口仍需限制请求体，避免大文件提前占用磁盘。进程异常退出时的临时目录由宿主系统按策略清理。

依赖组件包括 PyMuPDF（AGPL/商业双许可）；如用于闭源商业订阅服务，应先评估依赖许可证并选择合适授权方案。本仓库未替用户选择开源许可证。
