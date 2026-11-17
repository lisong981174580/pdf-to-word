import io
import subprocess
import tempfile
import zipfile
from pathlib import Path
from unittest.mock import patch

import fitz
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from PIL import Image
from pptx import Presentation
from app.main import app

client = TestClient(app)


def pdf(text='Hello PDF to Word'):
    doc = fitz.open()
    page = doc.new_page()
    if text:
        page.insert_text((72, 72), text)
    data = doc.tobytes()
    doc.close()
    return data


def test_conversion_returns_editable_text_and_cleans_files():
    with tempfile.TemporaryDirectory() as directory:
        with patch('tempfile.tempdir', directory):
            response = client.post('/api/convert', files={'file': ('测试.pdf', pdf(), 'application/pdf')})
            assert response.status_code == 200, response.text
            document = Document(io.BytesIO(response.content))
            assert 'Hello PDF to Word' in '\n'.join(p.text for p in document.paragraphs)
            assert "filename*=UTF-8''" in response.headers['content-disposition']
            assert response.headers['cache-control'] == 'no-store'
            assert list(Path(directory).iterdir()) == []


def test_invalid_and_scanned_files():
    for name, data, expected in [('x.txt', b'abc', 415), ('x.pdf', b'abc', 422), ('empty.pdf', b'', 422), ('scan.pdf', pdf(''), 422), ('bad.pdf', b'%PDF-broken', 422)]:
        assert client.post('/api/convert', files={'file': (name, data)}).status_code == expected


def test_encrypted_pdf():
    doc = fitz.open(stream=pdf(), filetype='pdf')
    data = doc.tobytes(encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw='owner', user_pw='secret')
    doc.close()
    assert client.post('/api/convert', files={'file': ('locked.pdf', data)}).status_code == 422


def test_limits():
    with patch('app.main.MAX_BYTES', 10):
        assert client.post('/api/convert', files={'file': ('large.pdf', pdf())}).status_code == 413
    with patch('app.main.MAX_PAGES', 0):
        assert client.post('/api/convert', files={'file': ('pages.pdf', pdf())}).status_code == 422


def test_timeout():
    with patch('app.main.subprocess.run', side_effect=subprocess.TimeoutExpired('converter', 1)):
        assert client.post('/api/convert', files={'file': ('slow.pdf', pdf())}).status_code == 504


def test_home_and_health():
    assert client.get('/').status_code == 200
    assert client.get('/static/app.js').status_code == 200
    assert client.get('/health').json() == {'status': 'ok'}


def post_tool(tool, items, query=''):
    return client.post(f'/api/tools/{tool}{query}', files=[('files', item) for item in items])


def test_pdf_export_tools():
    source = ('sample.pdf', pdf('Toolbox content'), 'application/pdf')
    images = post_tool('pdf-images', [source])
    assert images.status_code == 200
    assert zipfile.ZipFile(io.BytesIO(images.content)).namelist() == ['page-001.png']

    slides = post_tool('pdf-ppt', [source])
    assert slides.status_code == 200
    assert len(Presentation(io.BytesIO(slides.content)).slides) == 1

    sheet = post_tool('pdf-excel', [source])
    assert sheet.status_code == 200
    assert 'Toolbox content' in load_workbook(io.BytesIO(sheet.content)).active['A1'].value


def test_pdf_processing_and_image_to_pdf():
    merged = post_tool('pdf-merge', [('a.pdf', pdf('A'), 'application/pdf'), ('b.pdf', pdf('B'), 'application/pdf')])
    assert merged.status_code == 200
    with fitz.open(stream=merged.content, filetype='pdf') as document:
        assert len(document) == 2

    split = post_tool('pdf-split', [('two.pdf', merged.content, 'application/pdf')])
    assert split.status_code == 200
    assert len(zipfile.ZipFile(io.BytesIO(split.content)).namelist()) == 2

    rotated = post_tool('pdf-rotate', [('a.pdf', pdf('A'), 'application/pdf')], '?rotation=90')
    with fitz.open(stream=rotated.content, filetype='pdf') as document:
        assert document[0].rotation == 90

    image = io.BytesIO(); Image.new('RGB', (40, 30), '#2b8a68').save(image, 'PNG')
    result = post_tool('images-pdf', [('image.png', image.getvalue(), 'image/png')])
    with fitz.open(stream=result.content, filetype='pdf') as document:
        assert len(document) == 1


def test_tool_validation_and_office_dependency():
    assert post_tool('pdf-merge', [('a.pdf', pdf(), 'application/pdf')]).status_code == 422
    assert post_tool('pdf-images', [('x.txt', b'no', 'text/plain')]).status_code == 415
    with patch('app.main.shutil.which', return_value=None):
        response = post_tool('word-pdf', [('a.docx', b'fake', 'application/octet-stream')])
        assert response.status_code == 503

