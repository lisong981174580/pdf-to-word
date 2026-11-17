"""Self-hosted document toolbox. Uploaded files never leave this service."""
import asyncio
import io
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import quote

import fitz
from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches
from starlette.background import BackgroundTask

ROOT = Path(__file__).parent
MAX_BYTES = int(os.getenv('MAX_UPLOAD_MB', '20')) * 1024 * 1024
MAX_PAGES = int(os.getenv('MAX_PAGES', '100'))
TIMEOUT = int(os.getenv('CONVERSION_TIMEOUT', '120'))
slots = asyncio.Semaphore(2)
app = FastAPI(title='纸间文档工具箱', docs_url=None, redoc_url=None)
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


def validate_pdf(path, require_text=False):
    try:
        with fitz.open(path) as document:
            if document.needs_pass:
                raise HTTPException(422, '文件有密码保护，请先解密。')
            if not 0 < len(document) <= MAX_PAGES:
                raise HTTPException(422, f'请上传 1–{MAX_PAGES} 页的 PDF。')
            if require_text and not any(page.get_text().strip() for page in document):
                raise HTTPException(422, '该 PDF 未检测到文字，可能是扫描件；请先进行 OCR 识别。')
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(422, 'PDF 已损坏或格式不受支持。') from exc


async def save_upload(upload, target, allowed):
    suffix = Path((upload.filename or '').replace('\\', '/')).suffix.lower()
    if suffix not in allowed:
        raise HTTPException(415, f'文件格式不支持，请上传 {" / ".join(sorted(allowed))} 文件。')
    size = 0
    try:
        with target.open('wb') as output:
            while chunk := await upload.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_BYTES:
                    raise HTTPException(413, f'单个文件不能超过 {MAX_BYTES // 1024 // 1024} MB。')
                output.write(chunk)
    finally:
        await upload.close()
    if not size:
        raise HTTPException(422, '文件不能为空。')
    return suffix


def pdf_to_word(source, target):
    result = subprocess.run([sys.executable, '-m', 'app.worker', str(source), str(target)], timeout=TIMEOUT, capture_output=True)
    if result.returncode:
        raise HTTPException(422, '转换失败，请尝试其他 PDF 文件。')


def pdf_to_images(source, target):
    with fitz.open(source) as document, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as bundle:
        for number, page in enumerate(document, 1):
            bundle.writestr(f'page-{number:03d}.png', page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False).tobytes('png'))


def pdf_to_ppt(source, target):
    deck = Presentation()
    deck.slide_width, deck.slide_height = Inches(13.333), Inches(7.5)
    blank = deck.slide_layouts[6]
    with fitz.open(source) as document:
        for page in document:
            image = io.BytesIO(page.get_pixmap(matrix=fitz.Matrix(1.7, 1.7), alpha=False).tobytes('png'))
            slide = deck.slides.add_slide(blank)
            slide.shapes.add_picture(image, 0, 0, width=deck.slide_width, height=deck.slide_height)
    deck.save(target)


def pdf_to_excel(source, target):
    book = Workbook()
    book.remove(book.active)
    with fitz.open(source) as document:
        for number, page in enumerate(document, 1):
            sheet = book.create_sheet(f'第{number}页')
            words = page.get_text('words', sort=True)
            if not words:
                sheet['A1'] = '本页未检测到可提取文字'
                continue
            current_y, row, parts = None, 1, []
            for x0, y0, _x1, _y1, text, *_ in words:
                if current_y is None or abs(y0 - current_y) < 4:
                    parts.append((x0, text)); current_y = y0 if current_y is None else current_y
                else:
                    sheet.cell(row, 1, ' '.join(value for _, value in sorted(parts))); row += 1
                    parts, current_y = [(x0, text)], y0
            if parts:
                sheet.cell(row, 1, ' '.join(value for _, value in sorted(parts)))
            sheet.column_dimensions['A'].width = 100
    book.save(target)


def images_to_pdf(sources, target):
    output = fitz.open()
    try:
        for source in sources:
            image = fitz.open(source)
            pdf = fitz.open('pdf', image.convert_to_pdf())
            output.insert_pdf(pdf)
            pdf.close(); image.close()
        output.save(target, deflate=True)
    finally:
        output.close()


def merge_pdfs(sources, target):
    output = fitz.open()
    try:
        for source in sources:
            with fitz.open(source) as document:
                output.insert_pdf(document)
        output.save(target, garbage=4, deflate=True)
    finally:
        output.close()


def split_pdf(source, target):
    with fitz.open(source) as document, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as bundle:
        for number in range(len(document)):
            part = fitz.open(); part.insert_pdf(document, from_page=number, to_page=number)
            bundle.writestr(f'page-{number + 1:03d}.pdf', part.tobytes(garbage=4, deflate=True)); part.close()


def transform_pdf(source, target, tool, rotation):
    with fitz.open(source) as document:
        if tool == 'pdf-rotate':
            for page in document:
                page.set_rotation((page.rotation + rotation) % 360)
        document.save(target, garbage=4, deflate=True, clean=True)


def office_to_pdf(source, target, workdir):
    executable = shutil.which('libreoffice') or shutil.which('soffice')
    if not executable:
        raise HTTPException(503, 'Office 转 PDF 需要使用 Docker 版服务（内置 LibreOffice）。')
    # LibreOffice's bundled macOS build does not discover system CJK fonts by
    # default. An explicit fontconfig file also gives Linux predictable aliases
    # for common Windows Chinese fonts such as SimSun and Microsoft YaHei.
    font_config = ROOT.parent / 'fonts.conf'
    cache = workdir / 'font-cache'
    cache.mkdir()
    environment = os.environ.copy()
    if font_config.exists():
        environment['FONTCONFIG_FILE'] = str(font_config)
        environment['XDG_CACHE_HOME'] = str(cache)
    result = subprocess.run(
        [executable, '--headless', '--convert-to', 'pdf', '--outdir', str(workdir), str(source)],
        timeout=TIMEOUT, capture_output=True, env=environment,
    )
    produced = workdir / f'{source.stem}.pdf'
    if result.returncode or not produced.exists():
        raise HTTPException(422, 'Office 文档转换失败，请检查文件是否损坏。')
    produced.replace(target)


TOOL_RULES = {
    'pdf-word': ({'.pdf'}, '.docx', 'Word文档'), 'pdf-ppt': ({'.pdf'}, '.pptx', '演示文稿'),
    'pdf-excel': ({'.pdf'}, '.xlsx', '表格'), 'pdf-images': ({'.pdf'}, '.zip', '图片'),
    'images-pdf': ({'.jpg', '.jpeg', '.png', '.webp'}, '.pdf', '图片合集'),
    'pdf-merge': ({'.pdf'}, '.pdf', '合并文档'), 'pdf-split': ({'.pdf'}, '.zip', '拆分文档'),
    'pdf-compress': ({'.pdf'}, '.pdf', '压缩文档'), 'pdf-rotate': ({'.pdf'}, '.pdf', '旋转文档'),
    'word-pdf': ({'.doc', '.docx'}, '.pdf', 'PDF文档'), 'ppt-pdf': ({'.ppt', '.pptx'}, '.pdf', 'PDF文档'),
    'excel-pdf': ({'.xls', '.xlsx'}, '.pdf', 'PDF文档'),
}


@app.post('/api/tools/{tool}')
async def run_tool(tool: str, files: list[UploadFile], rotation: int = 90):
    if tool not in TOOL_RULES:
        raise HTTPException(404, '未知转换工具。')
    allowed, output_suffix, output_label = TOOL_RULES[tool]
    if not files or (tool == 'pdf-merge' and len(files) < 2):
        raise HTTPException(422, 'PDF 合并至少需要两个文件。' if tool == 'pdf-merge' else '请选择文件。')
    if tool not in {'pdf-merge', 'images-pdf'} and len(files) != 1:
        raise HTTPException(422, '此工具每次只能处理一个文件。')
    work = tempfile.TemporaryDirectory(prefix='paper-tool-'); keep = False
    try:
        folder = Path(work.name); sources = []
        for number, upload in enumerate(files):
            suffix = Path(upload.filename or '').suffix.lower()
            source = folder / f'input-{number}{suffix}'
            await save_upload(upload, source, allowed); sources.append(source)
            if suffix == '.pdf': validate_pdf(source, require_text=tool == 'pdf-word')
        target = folder / f'output{output_suffix}'
        try:
            async with slots:
                if tool == 'pdf-word': await asyncio.to_thread(pdf_to_word, sources[0], target)
                elif tool == 'pdf-ppt': await asyncio.to_thread(pdf_to_ppt, sources[0], target)
                elif tool == 'pdf-excel': await asyncio.to_thread(pdf_to_excel, sources[0], target)
                elif tool == 'pdf-images': await asyncio.to_thread(pdf_to_images, sources[0], target)
                elif tool == 'images-pdf': await asyncio.to_thread(images_to_pdf, sources, target)
                elif tool == 'pdf-merge': await asyncio.to_thread(merge_pdfs, sources, target)
                elif tool == 'pdf-split': await asyncio.to_thread(split_pdf, sources[0], target)
                elif tool in {'pdf-compress', 'pdf-rotate'}: await asyncio.to_thread(transform_pdf, sources[0], target, tool, rotation)
                else: await asyncio.to_thread(office_to_pdf, sources[0], target, folder)
        except subprocess.TimeoutExpired as exc:
            raise HTTPException(504, '转换超时，请减少文件大小后重试。') from exc
        if not target.exists() or not target.stat().st_size: raise HTTPException(422, '未能生成有效的输出文件。')
        original = Path(files[0].filename or 'document').stem[:100]
        filename = f'{original}-{output_label}{output_suffix}'
        response = FileResponse(target, media_type='application/octet-stream', headers={
            'Content-Disposition': 'attachment; filename=result' + output_suffix + "; filename*=UTF-8''" + quote(filename, safe=''),
            'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff'}, background=BackgroundTask(work.cleanup))
        keep = True
        return response
    finally:
        if not keep: work.cleanup()


@app.post('/api/convert')
async def convert_pdf(file: UploadFile):
    """Backward-compatible PDF-to-Word endpoint."""
    return await run_tool('pdf-word', [file])
