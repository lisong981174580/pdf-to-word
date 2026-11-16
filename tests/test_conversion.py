import io
import subprocess
import tempfile
from pathlib import Path
from unittest.mock import patch

import fitz
from docx import Document
from fastapi.testclient import TestClient
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
