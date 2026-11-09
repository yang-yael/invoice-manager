#!/usr/bin/env python3
"""Local web uploader for the invoice manager."""

from __future__ import annotations

import html
import io
import json
import mimetypes
import os
import shutil
import subprocess
import sys
import threading
import webbrowser
import zipfile
from base64 import b64decode
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

from invoice_manager import (
    DEFAULT_TEMPLATE,
    ROOT,
    SUMMARY_EXTS,
    SUPPORTED_EXTS,
    export_buyer_results,
    scan_folder,
    unique_path,
)


HOST = os.environ.get("HOST", "127.0.0.1")
PORT = int(os.environ.get("PORT", "8765"))
CLOUD_MODE = os.environ.get("INVOICE_CLOUD", "").strip().lower() in {"1", "true", "yes"}
AUTH_USER = os.environ.get("INVOICE_USER", "admin")
AUTH_PASSWORD = os.environ.get("INVOICE_PASSWORD", "")
DATA_DIR = os.environ.get("INVOICE_DATA_DIR", "").strip()
if DATA_DIR:
    DATA_ROOT = Path(DATA_DIR).expanduser().resolve()
    UPLOAD_ROOT = DATA_ROOT / "uploads"
    OUTPUT_ROOT = DATA_ROOT / "outputs"
else:
    DATA_ROOT = ROOT
    UPLOAD_ROOT = ROOT / "uploads"
    OUTPUT_ROOT = ROOT / "outputs"
AUTO_OPEN = os.environ.get("INVOICE_AUTO_OPEN", "").strip().lower()
if not AUTO_OPEN:
    AUTO_OPEN_BROWSER = not CLOUD_MODE and HOST in {"127.0.0.1", "localhost"}
else:
    AUTO_OPEN_BROWSER = AUTO_OPEN in {"1", "true", "yes"}


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>发票自动整理</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f9;
      --panel: #ffffff;
      --line: #d9dee7;
      --text: #1f2937;
      --muted: #667085;
      --brand: #0f766e;
      --brand-dark: #115e59;
      --danger: #b42318;
      --secondary: #475467;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: var(--bg);
      color: var(--text);
    }
    header {
      height: 64px;
      border-bottom: 1px solid var(--line);
      background: var(--panel);
      display: flex;
      align-items: center;
      padding: 0 28px;
      gap: 14px;
    }
    header h1 { font-size: 20px; margin: 0; font-weight: 650; }
    main {
      max-width: 980px;
      margin: 28px auto;
      padding: 0 20px 40px;
    }
    .toolbar {
      display: grid;
      grid-template-columns: 1fr 1fr auto;
      gap: 12px;
      margin-bottom: 16px;
      align-items: end;
    }
    label { display: block; font-size: 13px; color: var(--muted); margin-bottom: 6px; }
    input[type="text"] {
      width: 100%;
      height: 40px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: white;
      padding: 0 12px;
      font-size: 15px;
    }
    button {
      height: 40px;
      border: 0;
      border-radius: 6px;
      background: var(--brand);
      color: white;
      padding: 0 18px;
      font-size: 15px;
      cursor: pointer;
    }
    button:hover { background: var(--brand-dark); }
    button:disabled { opacity: .5; cursor: not-allowed; }
    button.secondary { background: var(--secondary); }
    button.secondary:hover { background: #344054; }
    .actions {
      display: flex;
      gap: 10px;
      justify-content: center;
      flex-wrap: wrap;
      margin-top: 16px;
    }
    .dropzone {
      min-height: 260px;
      border: 2px dashed #a6b0c0;
      border-radius: 8px;
      background: var(--panel);
      display: flex;
      align-items: center;
      justify-content: center;
      text-align: center;
      padding: 28px;
    }
    .dropzone.dragover {
      border-color: var(--brand);
      background: #eefaf8;
    }
    .dropzone strong { display: block; font-size: 22px; margin-bottom: 10px; }
    .dropzone span { color: var(--muted); }
    .file-input { display: none; }
    .content {
      margin-top: 18px;
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 16px;
      min-height: 120px;
    }
    section h2 { margin: 0 0 12px; font-size: 16px; }
    ul { margin: 0; padding: 0; list-style: none; }
    li {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      padding: 8px 0;
      border-bottom: 1px solid #edf0f5;
      font-size: 14px;
    }
    li:last-child { border-bottom: 0; }
    .muted { color: var(--muted); }
    .error { color: var(--danger); }
    .links a {
      display: inline-flex;
      align-items: center;
      height: 34px;
      margin: 4px 8px 4px 0;
      padding: 0 12px;
      border: 1px solid var(--line);
      border-radius: 6px;
      text-decoration: none;
      color: var(--brand-dark);
      background: white;
    }
    @media (max-width: 760px) {
      .toolbar, .content { grid-template-columns: 1fr; }
      header { padding: 0 18px; }
      main { margin-top: 18px; }
    }
  </style>
</head>
<body>
  <header><h1>发票自动整理</h1><span class="muted">本地上传，本地生成 Excel</span></header>
  <main>
    <div class="toolbar">
      <div>
        <label for="month">总报销批次名称</label>
        <input id="month" type="text" placeholder="必填，例如：4月产5月报销">
      </div>
      <div>
        <label for="reimburser">报销人</label>
        <input id="reimburser" type="text" placeholder="可不填">
      </div>
      <button id="submit" disabled>开始识别</button>
    </div>
    <div id="dropzone" class="dropzone">
      <div>
        <strong>拖入 PDF、图片或汇总表</strong>
        <span>也可以选择总文件夹：总表 Excel 会按发票号码匹配 PDF，再识别购买方公司</span>
        <div class="actions">
          <button type="button" id="chooseFiles">选择文件/ZIP</button>
          <button type="button" id="chooseFolder" class="secondary">选择报销人文件夹</button>
        </div>
      </div>
    </div>
    <input id="files" class="file-input" type="file" multiple accept=".pdf,.png,.jpg,.jpeg,.webp,.bmp,.tif,.tiff,.xlsx,.xlsm,.zip">
    <input id="folders" class="file-input" type="file" multiple webkitdirectory directory>
    <div class="content">
      <section>
        <h2>待处理文件</h2>
        <ul id="fileList"><li class="muted">还没有选择文件</li></ul>
      </section>
      <section>
        <h2>处理结果</h2>
        <div id="status" class="muted">等待上传</div>
        <div id="links" class="links"></div>
      </section>
    </div>
  </main>
  <script>
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('files');
    const folderInput = document.getElementById('folders');
    const chooseFiles = document.getElementById('chooseFiles');
    const chooseFolder = document.getElementById('chooseFolder');
    const fileList = document.getElementById('fileList');
    const submit = document.getElementById('submit');
    const statusBox = document.getElementById('status');
    const links = document.getElementById('links');
    let selectedFiles = [];

    function renderFiles() {
      fileList.innerHTML = '';
      const hasBatchName = !!document.getElementById('month').value.trim();
      if (!selectedFiles.length) {
        fileList.innerHTML = '<li class="muted">还没有选择文件</li>';
        submit.disabled = true;
        return;
      }
      selectedFiles.forEach(file => {
        const li = document.createElement('li');
        const label = file.webkitRelativePath || file.name;
        li.innerHTML = `<span>${label}</span><span class="muted">${Math.ceil(file.size / 1024)} KB</span>`;
        fileList.appendChild(li);
      });
      submit.disabled = !hasBatchName;
    }

    function addFiles(files) {
      const seen = new Set(selectedFiles.map(file => (file.webkitRelativePath || file.name) + ':' + file.size));
      [...files].forEach(file => {
        const key = (file.webkitRelativePath || file.name) + ':' + file.size;
        if (!seen.has(key)) {
          selectedFiles.push(file);
          seen.add(key);
        }
      });
      renderFiles();
    }

    function clearCurrentBatch() {
      selectedFiles = [];
      fileInput.value = '';
      folderInput.value = '';
      document.getElementById('reimburser').value = '';
      renderFiles();
    }

    dropzone.addEventListener('click', event => {
      if (event.target.tagName === 'BUTTON') return;
      fileInput.click();
    });
    chooseFiles.addEventListener('click', event => {
      event.stopPropagation();
      fileInput.click();
    });
    chooseFolder.addEventListener('click', event => {
      event.stopPropagation();
      folderInput.click();
    });
    fileInput.addEventListener('change', event => addFiles(event.target.files));
    folderInput.addEventListener('change', event => addFiles(event.target.files));
    document.getElementById('month').addEventListener('input', renderFiles);
    dropzone.addEventListener('dragover', event => {
      event.preventDefault();
      dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', event => {
      event.preventDefault();
      dropzone.classList.remove('dragover');
      addFiles(event.dataTransfer.files);
    });

    submit.addEventListener('click', async () => {
      if (!selectedFiles.length) return;
      if (!document.getElementById('month').value.trim()) {
        statusBox.textContent = '请先填写总报销批次名称。同一个批次名下会持续汇总所有报销人的发票。';
        statusBox.className = 'error';
        return;
      }
      submit.disabled = true;
      statusBox.textContent = '正在上传和识别...';
      statusBox.className = 'muted';
      links.innerHTML = '';
      const form = new FormData();
      selectedFiles.forEach(file => form.append('files', file, file.webkitRelativePath || file.name));
      form.append('month', document.getElementById('month').value);
      form.append('reimburser', document.getElementById('reimburser').value);
      try {
        const response = await fetch('/api/process', { method: 'POST', body: form });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || '处理失败');
        statusBox.textContent = `本次新增 ${data.saved_count} 个文件；已按 ${data.company_count} 个购买方公司汇总 ${data.count} 个文件，导出 ${data.pdf_count} 个 PDF，${data.review_count} 个需要核对。`;
        const uploadLink = data.cloud_mode ? data.upload_url : data.upload_open_url;
        const outputLink = data.cloud_mode ? data.output_url : data.output_open_url;
        const uploadLabel = data.cloud_mode ? '查看上传目录' : '打开上传目录';
        const outputLabel = data.cloud_mode ? '查看结果目录' : '打开结果目录';
        links.innerHTML = `
          <a href="${uploadLink}" target="_blank" rel="noopener">${uploadLabel}</a>
          <a href="${outputLink}" target="_blank" rel="noopener">${outputLabel}</a>
        `;
        clearCurrentBatch();
      } catch (err) {
        statusBox.textContent = err.message;
        statusBox.className = 'error';
      } finally {
        submit.disabled = !selectedFiles.length;
      }
    });
  </script>
</body>
</html>
"""


def safe_name(name: str, fallback: str) -> str:
    cleaned = "".join(ch for ch in name.strip() if ch not in '/\\:*?"<>|')
    return cleaned or fallback


def safe_relative_path(name: str, fallback: str) -> Path:
    pieces = []
    for raw_part in name.replace("\\", "/").split("/"):
        if raw_part in {"", ".", ".."}:
            continue
        pieces.append(safe_name(raw_part, ""))
    if not pieces:
        pieces = [fallback]
    return Path(*pieces)


def relative_url(path: Path, prefix: str) -> str:
    rel_path = path.relative_to(ROOT).as_posix()
    return f"{prefix}/{quote(rel_path, safe='/')}"


def write_json(handler: BaseHTTPRequestHandler, status: int, payload: dict) -> None:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def parse_content_type(value: str) -> tuple[str, dict[str, str]]:
    pieces = [piece.strip() for piece in value.split(";")]
    params = {}
    for piece in pieces[1:]:
        if "=" not in piece:
            continue
        key, raw = piece.split("=", 1)
        params[key.strip().lower()] = raw.strip().strip('"')
    return pieces[0].lower(), params


def parse_disposition(value: str) -> dict[str, str]:
    _, params = parse_content_type(value)
    return params


def decode_header(raw: bytes) -> str:
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("latin-1", errors="ignore")


def parse_multipart(body: bytes, boundary: str) -> tuple[dict[str, str], list[tuple[str, bytes]]]:
    fields: dict[str, str] = {}
    files: list[tuple[str, bytes]] = []
    marker = b"--" + boundary.encode("utf-8")
    for part in body.split(marker):
        part = part.strip(b"\r\n")
        if not part or part == b"--":
            continue
        if part.endswith(b"--"):
            part = part[:-2].strip(b"\r\n")
        if b"\r\n\r\n" not in part:
            continue
        raw_headers, content = part.split(b"\r\n\r\n", 1)
        headers = {}
        for line in raw_headers.split(b"\r\n"):
            if b":" not in line:
                continue
            key, value = line.split(b":", 1)
            headers[key.decode("ascii", errors="ignore").lower()] = decode_header(value.strip())
        disposition = parse_disposition(headers.get("content-disposition", ""))
        name = disposition.get("name", "")
        filename = disposition.get("filename", "")
        if filename:
            files.append((filename.replace("\\", "/"), content))
        elif name:
            fields[name] = content.decode("utf-8", errors="ignore")
    return fields, files


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: object) -> None:
        print(fmt % args)

    def _authorized(self) -> bool:
        if not AUTH_PASSWORD:
            return True
        header = self.headers.get("Authorization", "")
        if not header.startswith("Basic "):
            return False
        try:
            decoded = b64decode(header.removeprefix("Basic ").strip()).decode("utf-8")
        except Exception:
            return False
        username, sep, password = decoded.partition(":")
        return bool(sep) and username == AUTH_USER and password == AUTH_PASSWORD

    def _require_auth(self) -> bool:
        if self._authorized():
            return True
        self.send_response(401)
        self.send_header("WWW-Authenticate", 'Basic realm="Invoice Manager"')
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        body = "需要登录".encode("utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
        return False

    def do_GET(self) -> None:
        if not self._require_auth():
            return
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_bytes(INDEX_HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/download/"):
            self._serve_download(unquote(parsed.path.removeprefix("/download/")))
            return
        if parsed.path.startswith("/open/"):
            self._open_local(unquote(parsed.path.removeprefix("/open/")))
            return
        write_json(self, 404, {"error": "Not found"})

    def do_POST(self) -> None:
        if not self._require_auth():
            return
        if urlparse(self.path).path != "/api/process":
            write_json(self, 404, {"error": "Not found"})
            return
        ctype, params = parse_content_type(self.headers.get("Content-Type", ""))
        if ctype != "multipart/form-data":
            write_json(self, 400, {"error": "请上传文件"})
            return
        boundary = params.get("boundary", "")
        if not boundary:
            write_json(self, 400, {"error": "上传格式缺少 boundary"})
            return
        length = int(self.headers.get("Content-Length", "0"))
        fields, files = parse_multipart(self.rfile.read(length), boundary)
        try:
            result = self._process_upload(fields, files)
        except Exception as exc:
            write_json(self, 500, {"error": str(exc)})
            return
        write_json(self, 200, result)

    def _send_bytes(self, data: bytes, content_type: str) -> None:
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _resolve_root_path(self, rel_path: str) -> Path | None:
        path = (ROOT / rel_path).resolve()
        if ROOT.resolve() not in path.parents and path != ROOT.resolve():
            return None
        return path

    def _serve_download(self, rel_path: str) -> None:
        path = self._resolve_root_path(rel_path)
        if path is None:
            write_json(self, 403, {"error": "Forbidden"})
            return
        if path.is_dir():
            rows = []
            for p in sorted(path.iterdir()):
                url = relative_url(p, "/download")
                label = html.escape(p.name + ("/" if p.is_dir() else ""))
                rows.append(f"<li><a href='{url}'>{label}</a></li>")
            body = f"<html><meta charset='utf-8'><body><h1>{html.escape(path.name)}</h1><ul>{''.join(rows)}</ul></body></html>"
            self._send_bytes(body.encode("utf-8"), "text/html; charset=utf-8")
            return
        if not path.exists():
            write_json(self, 404, {"error": "文件不存在"})
            return
        ctype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        quoted_name = quote(path.name)
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Disposition", f"attachment; filename=\"download{path.suffix}\"; filename*=UTF-8''{quoted_name}")
        self.send_header("Content-Length", str(path.stat().st_size))
        self.end_headers()
        with path.open("rb") as f:
            shutil.copyfileobj(f, self.wfile)

    def _open_local(self, rel_path: str) -> None:
        path = self._resolve_root_path(rel_path)
        if path is None:
            write_json(self, 403, {"error": "Forbidden"})
            return
        if not path.exists():
            write_json(self, 404, {"error": "文件不存在"})
            return
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                proc = subprocess.run(["open", str(path)], capture_output=True, text=True, check=False)
                if proc.returncode != 0:
                    raise RuntimeError(proc.stderr.strip() or "打开失败")
            else:
                proc = subprocess.run(["xdg-open", str(path)], capture_output=True, text=True, check=False)
                if proc.returncode != 0:
                    raise RuntimeError(proc.stderr.strip() or "打开失败")
        except Exception as exc:
            write_json(self, 500, {"error": str(exc) or "打开失败"})
            return
        body = (
            "<html><meta charset='utf-8'>"
            "<body style='font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",sans-serif;padding:24px;'>"
            f"<h1>已在本机打开</h1><p>{html.escape(str(path))}</p>"
            "<p>这个页面可以关闭，回到发票整理页面继续操作。</p>"
            "</body></html>"
        )
        self._send_bytes(body.encode("utf-8"), "text/html; charset=utf-8")

    def _extract_zip_upload(
        self,
        original_name: str,
        content: bytes,
        upload_dir: Path,
        reimburser_name: str,
    ) -> list[Path]:
        owner = safe_name(reimburser_name, "") or safe_name(
            Path(original_name).stem,
            f"zip_{datetime.now().strftime('%H%M%S')}",
        )
        extract_dir = upload_dir / owner
        extract_dir.mkdir(parents=True, exist_ok=True)

        saved = []
        with zipfile.ZipFile(io.BytesIO(content)) as archive:
            for info in archive.infolist():
                if info.is_dir():
                    continue
                raw_name = info.filename.replace("\\", "/")
                if raw_name.startswith("/") or any(part == ".." for part in raw_name.split("/")):
                    continue
                rel_path = safe_relative_path(info.filename, "")
                if not rel_path.parts or rel_path.parts[0] == "__MACOSX":
                    continue
                if rel_path.suffix.lower() not in SUPPORTED_EXTS | SUMMARY_EXTS:
                    continue
                target = unique_path(extract_dir / rel_path)
                target.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(info) as src, target.open("wb") as dst:
                    shutil.copyfileobj(src, dst)
                saved.append(target)
        return saved

    def _save_upload_file(
        self,
        original_name: str,
        content: bytes,
        upload_dir: Path,
        reimburser_name: str,
        fallback: str,
    ) -> list[Path]:
        rel_path = safe_relative_path(original_name, fallback)
        if rel_path.suffix.lower() == ".zip":
            return self._extract_zip_upload(rel_path.name, content, upload_dir, reimburser_name)

        if rel_path.suffix.lower() not in SUPPORTED_EXTS | SUMMARY_EXTS:
            return []

        if reimburser_name:
            rel_path = Path(safe_name(reimburser_name, "报销人待核对")) / rel_path
        elif len(rel_path.parts) == 1:
            rel_path = Path("报销人待核对") / rel_path

        target = unique_path(upload_dir / rel_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("wb") as f:
            f.write(content)
        return [target]

    def _process_upload(self, fields: dict[str, str], files: list[tuple[str, bytes]]) -> dict:
        raw_month = fields.get("month", "").strip()
        if not raw_month:
            raise ValueError("请填写总报销批次名称")
        month = safe_name(raw_month, datetime.now().strftime("%Y-%m-%d_%H%M%S"))
        reimburser = safe_name(fields.get("reimburser", ""), "")
        upload_dir = UPLOAD_ROOT / month
        output_dir = OUTPUT_ROOT / month
        upload_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        saved = []
        for original_name, content in files:
            if not original_name:
                continue
            saved.extend(
                self._save_upload_file(
                    original_name,
                    content,
                    upload_dir,
                    reimburser,
                    f"invoice_{len(saved) + 1}",
                )
            )

        if not saved:
            raise ValueError("没有收到有效文件")

        records = scan_folder(upload_dir, "")
        export_result = export_buyer_results(DEFAULT_TEMPLATE, output_dir, month, records)
        review_count = sum(1 for record in records if record.needs_review)

        return {
            "count": len(records),
            "saved_count": len(saved),
            "company_count": export_result["company_count"],
            "review_count": review_count,
            "pdf_count": export_result["pdf_count"],
            "upload_url": relative_url(upload_dir, "/download"),
            "output_url": relative_url(output_dir, "/download"),
            "upload_open_url": relative_url(upload_dir, "/open"),
            "output_open_url": relative_url(output_dir, "/open"),
            "cloud_mode": CLOUD_MODE,
        }


def main() -> int:
    UPLOAD_ROOT.mkdir(parents=True, exist_ok=True)
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}"
    print(f"发票自动整理页面已启动：{url}")
    if AUTO_OPEN_BROWSER:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
