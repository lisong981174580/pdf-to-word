"""Self-hosted PDF to DOCX converter. No third-party document uploads."""
import asyncio
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import quote

import fitz
from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

ROOT = Path(__file__).parent
MAX_BYTES = int(os.getenv('MAX_UPLOAD_MB', '20')) * 1024 * 1024
MAX_PAGES = int(os.getenv('MAX_PAGES', '100'))
TIMEOUT = int(os.getenv('CONVERSION_TIMEOUT', '120'))
slots = asyncio.Semaphore(2)
app = FastAPI(title='PDF 转 Word', docs_url=None, redoc_url=None)
app.mount('/static', StaticFiles(directory=ROOT / 'static'), name='static')


@app.get('/')
async def index():
    return FileResponse(ROOT / 'static/index.html')


@app.get('/health')
async def health():
    return {'status': 'ok'}


@app.get('/api/config')
async def config():
    return {'max_upload_bytes': MAX_BYTES, 'max_pages': MAX_PAGES}


def validate_pdf(path):
    try:
        with fitz.open(path) as document:
            if document.needs_pass:
                raise HTTPException(422, '文件有密码保护，请先解密。')
            if not 0 < len(document) <= MAX_PAGES:
                raise HTTPException(422, f'请上传 1–{MAX_PAGES} 页的 PDF。')
            if not any(page.get_text().strip() for page in document):
                raise HTTPException(422, '该 PDF 未检测到文字，可能是扫描件；请先进行 OCR 识别。')
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(422, 'PDF 已损坏或格式不受支持。') from exc


def convert(source, target):
    try:
        result = subprocess.run(
            [sys.executable, '-m', 'app.worker', str(source), str(target)],
            timeout=TIMEOUT, capture_output=True, check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise HTTPException(504, '转换超时，请减少页数后重试。') from exc
    if result.returncode or not target.exists() or target.stat().st_size == 0:
        raise HTTPException(422, '转换失败，请尝试其他 PDF 文件。')


@app.post('/api/convert')
async def convert_pdf(file: UploadFile):
    work = tempfile.TemporaryDirectory(prefix='pdf-word-')
    keep = False
    try:
        if not file.filename or not file.filename.lower().endswith('.pdf'):
            raise HTTPException(415, '请选择 PDF 文件。')
        folder = Path(work.name)
        source, target = folder / 'input.pdf', folder / 'output.docx'
        size = 0
        with source.open('wb') as output:
            while chunk := await file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_BYTES:
                    raise HTTPException(413, f'文件不能超过 {MAX_BYTES // 1024 // 1024} MB。')
                output.write(chunk)
        if size == 0 or not source.read_bytes()[:1024].lstrip().startswith(b'%PDF-'):
            raise HTTPException(422, '文件内容不是有效的 PDF。')
        async with slots:
            await asyncio.to_thread(validate_pdf, source)
            await asyncio.to_thread(convert, source, target)
        # User-controlled filenames never become filesystem paths.
        name = Path(file.filename.replace('\\', '/')).stem[:120] + '.docx'
        response = FileResponse(
            target,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers={'Content-Disposition': "attachment; filename=converted.docx; filename*=UTF-8''" + quote(name, safe=''),
                     'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff'},
            background=BackgroundTask(work.cleanup),
        )
        keep = True
        return response
    finally:
        await file.close()
        if not keep:
            work.cleanup()
