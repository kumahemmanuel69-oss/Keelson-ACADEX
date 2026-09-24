#!/usr/bin/env python3
"""
A tiny file-drop server so the user can send a PDF through their browser
when the chat attachment path fails.

GET  /                 -> upload page
POST /upload?name=...  -> raw request body is written to uploads/
GET  /health           -> ok
GET  /list             -> JSON of what has arrived
"""
import os
import json
import re
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PORT = int(os.environ.get("PORT", "8000"))
DEST = "/home/user/Keelson-ACADEX/uploads"

PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Send your file</title>
<style>
  * { box-sizing: border-box; }
  body {
    margin: 0; min-height: 100vh; display: flex; align-items: center;
    justify-content: center; padding: 24px;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: #FDF8F0; color: #14262E;
  }
  .card {
    width: 100%; max-width: 620px; background: #fff; border-radius: 16px;
    box-shadow: 0 2px 24px rgba(7,74,53,.10); overflow: hidden;
  }
  .bar { height: 8px; background: #F4B400; }
  .head { padding: 26px 30px 6px; }
  h1 { margin: 0 0 6px; font-size: 22px; color: #074A35; letter-spacing: -.2px; }
  p.sub { margin: 0; color: #5A6B72; font-size: 13.5px; line-height: 1.55; }
  .body { padding: 18px 30px 30px; }
  .drop {
    border: 2.5px dashed #9CC3B4; border-radius: 12px; padding: 34px 20px;
    text-align: center; cursor: pointer; transition: .18s; background: #F7FCFA;
  }
  .drop:hover, .drop.over { border-color: #0B6E4F; background: #E9F5F0; }
  .drop strong { display: block; color: #0B6E4F; font-size: 15px; margin-bottom: 5px; }
  .drop span { color: #5A6B72; font-size: 12.5px; }
  input[type=file] { display: none; }
  .picked { margin-top: 14px; font-size: 13px; color: #0B6E4F; font-weight: 600; display: none; }
  button {
    width: 100%; margin-top: 16px; padding: 14px; border: 0; border-radius: 10px;
    background: #0B6E4F; color: #fff; font-size: 15.5px; font-weight: 700;
    cursor: pointer; transition: .18s;
  }
  button:hover:not(:disabled) { background: #074A35; }
  button:disabled { background: #C7D6D0; cursor: not-allowed; }
  #msg { margin-top: 16px; font-size: 13.5px; line-height: 1.55; display: none;
         padding: 15px 17px; border-radius: 10px; }
  .ok { background: #E3F1EA; color: #074A35; border-left: 4px solid #0B6E4F; }
  .err { background: #FBE9E0; color: #8A2E08; border-left: 4px solid #C1440E; }
  .info { background: #FDF3D8; color: #6B4E00; border-left: 4px solid #F4B400; }
  .note { margin-top: 18px; font-size: 12px; color: #7A8A90; line-height: 1.6; }
  code { background: #F1F5F3; padding: 1.5px 5px; border-radius: 4px;
         font-size: 11.5px; color: #0B6E4F; }
</style>
</head>
<body>
<div class="card">
  <div class="bar"></div>
  <div class="head">
    <h1>Send your file directly</h1>
    <p class="sub">The chat attachment hasn't been reaching the workspace, so use this
    page instead. Pick your PDF &mdash; it goes straight into the project folder.</p>
  </div>
  <div class="body">
    <label class="drop" id="drop" for="file">
      <strong>Click to choose your PDF</strong>
      <span>or drag and drop it here</span>
    </label>
    <input type="file" id="file" accept=".pdf,.png,.jpg,.jpeg,.txt,.md,.docx,.pptx">
    <div class="picked" id="picked"></div>
    <button id="send" disabled>Send file</button>
    <div id="msg"></div>
    <p class="note">Works with any file. Nothing is shared &mdash; it is written straight into
    the sandbox at <code>Keelson-ACADEX/uploads/</code>.</p>
  </div>
</div>
<script>
const drop = document.getElementById('drop');
const input = document.getElementById('file');
const picked = document.getElementById('picked');
const btn = document.getElementById('send');
const msg = document.getElementById('msg');
let chosen = null;

function show(text, cls) {
  msg.className = ''; msg.style.display = 'block';
  msg.classList.add(cls); msg.innerHTML = text;
}

function setFile(f) {
  if (!f) return;
  chosen = f;
  const mb = (f.size / 1048576).toFixed(2);
  picked.textContent = '\\u2713 ' + f.name + '  \\u00b7  ' + mb + ' MB';
  picked.style.display = 'block';
  btn.disabled = false;
  msg.style.display = 'none';
}

input.addEventListener('change', e => setFile(e.target.files[0]));
['dragenter','dragover'].forEach(ev =>
  drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(ev =>
  drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => {
  const f = e.dataTransfer.files[0];
  if (f) { input.files = e.dataTransfer.files; setFile(f); }
});

btn.addEventListener('click', async () => {
  if (!chosen) return;
  btn.disabled = true; btn.textContent = 'Sending\\u2026';
  show('Uploading ' + chosen.name + '\\u2026', 'info');
  try {
    const r = await fetch('/upload?name=' + encodeURIComponent(chosen.name), {
      method: 'POST',
      headers: { 'Content-Type': chosen.type || 'application/octet-stream' },
      body: chosen
    });
    const j = await r.json();
    if (r.ok && j.ok) {
      show('<strong>Received.</strong><br>' + j.name + '  \\u00b7  ' + j.mb +
           ' MB<br><br>You can now go back to the chat and say <em>\\"I\\u2019ve sent it\\"</em>.', 'ok');
      btn.textContent = 'Sent \\u2713';
      btn.style.background = '#0B6E4F';
    } else {
      throw new Error(j.error || ('HTTP ' + r.status));
    }
  } catch (err) {
    show('<strong>Could not send.</strong><br>' + err.message +
         '<br><br>Try again, or paste the text into the chat instead.', 'err');
    btn.disabled = false; btn.textContent = 'Send file';
  }
});
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    server_version = "FileDrop/1.0"

    def log_message(self, fmt, *a):
        print("  %s - %s" % (self.address_string(), fmt % a), flush=True)

    def _send(self, code, body, ctype="text/html; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        # deliberately permissive so the live preview can embed this page
        self.send_header("Access-Control-Allow-Origin", "*")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        try:
            self.wfile.write(body)
        except BrokenPipeError:
            pass

    def do_OPTIONS(self):
        self._send(204, b"")

    def do_HEAD(self):
        # some preview proxies probe with HEAD before serving
        self._send(200, b"", "text/html; charset=utf-8")

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path in ("/", "/index.html"):
            self._send(200, PAGE)
        elif path == "/health":
            self._send(200, "ok", "text/plain")
        elif path == "/list":
            files = []
            if os.path.isdir(DEST):
                for n in sorted(os.listdir(DEST)):
                    p = os.path.join(DEST, n)
                    if os.path.isfile(p):
                        files.append({"name": n, "bytes": os.path.getsize(p)})
            self._send(200, json.dumps({"files": files}), "application/json")
        else:
            self._send(404, "not found", "text/plain")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/upload":
            self._send(404, "not found", "text/plain")
            return
        qs = urllib.parse.parse_qs(parsed.query)
        raw_name = (qs.get("name") or ["upload.pdf"])[0]
        # keep it to a safe basename
        name = os.path.basename(raw_name).replace("\x00", "")
        name = re.sub(r"[^A-Za-z0-9._ ()\-]+", "_", name).strip() or "upload.pdf"
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = 0
        if length <= 0:
            self._send(400, json.dumps({"ok": False, "error": "empty body"}),
                       "application/json")
            return
        os.makedirs(DEST, exist_ok=True)
        dest = os.path.join(DEST, name)
        stem, ext = os.path.splitext(dest)
        n = 2
        while os.path.exists(dest):
            dest = "%s (%d)%s" % (stem, n, ext)
            n += 1
        remaining = length
        try:
            with open(dest, "wb") as fh:
                while remaining > 0:
                    chunk = self.rfile.read(min(262144, remaining))
                    if not chunk:
                        break
                    fh.write(chunk)
                    remaining -= len(chunk)
        except Exception as exc:
            self._send(500, json.dumps({"ok": False, "error": str(exc)}),
                       "application/json")
            return
        written = os.path.getsize(dest)
        print("  RECEIVED %s  (%d bytes)" % (os.path.basename(dest), written), flush=True)
        self._send(200, json.dumps({
            "ok": True,
            "name": os.path.basename(dest),
            "bytes": written,
            "mb": round(written / 1048576.0, 2),
        }), "application/json")


if __name__ == "__main__":
    os.makedirs(DEST, exist_ok=True)
    srv = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print("File drop listening on 0.0.0.0:%d" % PORT, flush=True)
    print("Saving into %s" % DEST, flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
