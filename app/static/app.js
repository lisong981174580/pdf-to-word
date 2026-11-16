const input = document.querySelector('#file');
const drop = document.querySelector('#drop');
const button = document.querySelector('#convert');
const status = document.querySelector('#status');
const download = document.querySelector('#download');
let file, url, busy = false, maxBytes = 20 * 1024 * 1024;
function clearDownload() { if (url) URL.revokeObjectURL(url); url = null; download.hidden = true; }
function select(candidate) {
  if (busy) return;
  clearDownload(); file = null; button.disabled = true;
  document.querySelector('#filename').textContent = '拖入 PDF 文件';
  document.querySelector('#filemeta').textContent = '或点击此处选择文件';
  if (!candidate) return;
  if (!/\.pdf$/i.test(candidate.name)) { status.textContent = '请选择 PDF 格式的文件'; return; }
  if (!candidate.size || candidate.size > maxBytes) { status.textContent = '文件为空或超过大小限制'; return; }
  file = candidate;
  document.querySelector('#filename').textContent = file.name;
  document.querySelector('#filemeta').textContent = (file.size / 1024 / 1024).toFixed(2) + ' MB · PDF 文档';
  status.textContent = '文件已就绪，点击开始转换'; button.disabled = false;
}
input.addEventListener('change', () => select(input.files[0]));
drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); if (!busy) input.click(); } });
for (const event of ['dragover', 'dragleave', 'drop']) drop.addEventListener(event, e => {
  e.preventDefault(); drop.classList.toggle('over', event === 'dragover');
  if (event === 'drop') select(e.dataTransfer.files[0]);
});
button.addEventListener('click', async () => {
  if (!file || busy) return;
  busy = true; button.disabled = true; input.disabled = true; clearDownload();
  status.textContent = '正在转换，请稍候。大文件可能需要约 2 分钟…';
  button.firstChild.textContent = '正在转换…';
  const data = new FormData(); data.append('file', file);
  try {
    const response = await fetch('/api/convert', { method: 'POST', body: data });
    if (!response.ok) {
      let message = '转换失败，请稍后重试';
      try { const body = await response.json(); if (typeof body.detail === 'string') message = body.detail; } catch (_) {}
      throw new Error(message);
    }
    url = URL.createObjectURL(await response.blob()); download.href = url;
    download.download = file.name.replace(/\.pdf$/i, '') + '.docx'; download.hidden = false;
    status.textContent = '转换完成，点击下方按钮保存 Word 文档';
  } catch (error) { status.textContent = error.message || '网络连接失败，请重试'; }
  finally { busy = false; button.disabled = false; input.disabled = false; button.firstChild.textContent = '重新转换 '; }
});
fetch('/api/config').then(r => r.json()).then(config => {
  maxBytes = config.max_upload_bytes;
  document.querySelector('#limits').textContent = `支持文字版 PDF · 最大 ${maxBytes / 1024 / 1024} MB / ${config.max_pages} 页`;
}).catch(() => {});
window.addEventListener('beforeunload', clearDownload);
