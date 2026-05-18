"""
RBI PDF Watermark Remover — Flask Web App  (v3)
================================================
Local:   python app.py  →  http://localhost:5000
Deploy:  Render.com  (see render.yaml)

Handles all known RBI watermark variants:
  Variant A - Scanned PDFs (Canon): full-page /Im2 image with soft-mask
  Variant B - Word-generated PDFs : full-page /Im1 RGB image overlay
  Variant C - Text-based          : diagonal repeated text in content stream
"""

import os, re, io, hashlib, time, tempfile
import fitz          # PyMuPDF  — text-watermark removal
import pikepdf       # pikepdf  — image-watermark removal
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024   # 50 MB
app.secret_key = os.environ.get('SECRET_KEY', 'rbi-wm-remover-2025')

TEMP_DIR = tempfile.gettempdir()

# ═══════════════════════════════════════════════════════════════════════════════
#  VARIANT C — Text-based watermark helpers  (PyMuPDF / fitz)
# ═══════════════════════════════════════════════════════════════════════════════

COMMON_WATERMARKS = [
    "DRAFT", "CONFIDENTIAL", "SAMPLE", "WATERMARK", "COPY",
    "DO NOT COPY", "FOR REVIEW", "INTERNAL", "RESTRICTED",
    "TOP SECRET", "CLASSIFIED", "NOT FOR DISTRIBUTION",
    "PREVIEW", "SPECIMEN", "VOID", "CANCELLED", "PAID",
    "RECEIVED", "APPROVED", "REJECTED", "PENDING", "DUPLICATE",
]

COMPANION_PATTERNS = [
    r'^\w{3}-\d{2}-\d{4}',
    r'^\d{1,2}-\d{2}-\d{4}',
    r'^\d{1,2}:\d{2}(:\d{2})?',
    r'^\d{4,6}$',
    r'^[A-Z]+-\d{2}-\d{4}\s+\d{2}',
]

def _is_companion_text(text):
    return any(re.match(p, text.strip(), re.IGNORECASE) for p in COMPANION_PATTERNS)

def _stream_contains_watermark(stream_bytes, terms):
    try:
        text = stream_bytes.decode('latin-1', errors='replace').upper()
    except Exception:
        return False
    return any(t in text for t in terms)

def _stream_is_only_watermark(stream_bytes, terms):
    try:
        text = stream_bytes.decode('latin-1', errors='replace')
    except Exception:
        return False
    bt_blocks = re.findall(r'BT\b(.*?)\bET', text, re.DOTALL)
    if not bt_blocks:
        return False
    terms_upper = [t.upper() for t in terms]
    for block in bt_blocks:
        literals = re.findall(r'\(((?:[^()\\]|\\.)*)\)', block)
        for lit in literals:
            lit_clean = re.sub(r'\\(.)', r'\1', lit)
            if any(t in lit_clean.upper() for t in terms_upper):
                continue
            if lit_clean.strip():
                return False
    return True


# ═══════════════════════════════════════════════════════════════════════════════
#  VARIANT A & B — Full-page image watermark helpers  (pikepdf)
# ═══════════════════════════════════════════════════════════════════════════════

def _is_fullpage_image(img_obj, page):
    """
    True if an XObject image covers ≥85% of the page in both dimensions.
    Real content images are never full-page; only watermark overlays are.
    """
    try:
        w = float(img_obj.get("/Width", 0))
        h = float(img_obj.get("/Height", 0))
        mb = page.mediabox
        pw = abs(float(mb[2]) - float(mb[0]))
        ph = abs(float(mb[3]) - float(mb[1]))
        if pw == 0 or ph == 0:
            return False
        return (w / pw >= 0.85) and (h / ph >= 0.85)
    except Exception:
        return False

def _appears_on_most_pages(pdf, name):
    """
    True if this XObject name appears on at least half the pages.
    A real content image appears on 1-2 pages; a watermark appears everywhere.
    """
    count = sum(
        1 for page in pdf.pages
        if name in page.get("/Resources", {}).get("/XObject", {})
    )
    return count >= max(2, len(pdf.pages) * 0.5)

def _find_watermark_image_names(pdf):
    """
    Auto-detect which XObject names are full-page watermark overlays.
    Returns a set of names like {'/Im1'} or {'/Im2'}.
    """
    candidates = set()
    for page in pdf.pages:
        xobjects = page.get("/Resources", {}).get("/XObject", {})
        for name, obj in xobjects.items():
            try:
                if str(obj.get("/Subtype", "")) != "/Image":
                    continue
                if _is_fullpage_image(obj, page):
                    candidates.add(name)
            except Exception:
                pass
    return {n for n in candidates if _appears_on_most_pages(pdf, n)}

def _remove_image_watermarks(pdf_bytes):
    """
    Remove full-page image watermarks (Variant A & B).
    Returns (cleaned_bytes, pages_cleaned, wm_names_found).
    """
    pdf = pikepdf.open(io.BytesIO(pdf_bytes))
    wm_names = _find_watermark_image_names(pdf)
    pages_cleaned = 0

    if not wm_names:
        return pdf_bytes, 0, set()

    for page in pdf.pages:
        page_modified = False

        # 1. Remove from XObject resource dict
        xobjects = page.get("/Resources", {}).get("/XObject", {})
        for name in wm_names:
            if name in xobjects:
                del xobjects[name]
                page_modified = True

        # 2. Remove draw instruction from content streams
        contents = page.get("/Contents")
        if contents is None:
            if page_modified:
                pages_cleaned += 1
            continue

        streams = list(contents) if isinstance(contents, pikepdf.Array) else [contents]
        for stream in streams:
            try:
                data = stream.read_bytes().decode('latin-1', errors='replace')
                original = data
                for name in wm_names:
                    bare = name.lstrip('/')
                    # Remove: q [transforms] /ImN Do Q
                    data = re.sub(
                        rf'q\b[^Q]*?/{re.escape(bare)}\s+Do[^Q]*?Q\b',
                        '',
                        data,
                        flags=re.DOTALL
                    )
                if data != original:
                    stream.write(data.encode('latin-1'))
                    page_modified = True
            except Exception:
                pass

        if page_modified:
            pages_cleaned += 1

    out = io.BytesIO()
    pdf.save(out)
    return out.getvalue(), pages_cleaned, wm_names


# ═══════════════════════════════════════════════════════════════════════════════
#  VARIANT C — Text-based watermark removal  (PyMuPDF / fitz)
# ═══════════════════════════════════════════════════════════════════════════════

def _remove_text_watermarks(pdf_bytes, custom_term=None):
    """
    Remove text-layer watermarks (diagonal repeated text, Variant C).
    Returns (cleaned_bytes, pages_cleaned, log_list).
    """
    terms = COMMON_WATERMARKS[:]
    if custom_term:
        terms.insert(0, custom_term.upper().strip())

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    log = []
    pages_cleaned = 0

    for page_num in range(len(doc)):
        page = doc[page_num]
        page_removed = 0

        # ── Method 1: redact text spans that match watermark terms ────────────
        blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        wm_rects = []
        companion_rects = []
        wm_terms_found = []

        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    txt = span["text"].strip()
                    if not txt:
                        continue
                    txt_up = txt.upper()
                    is_wm = any(t in txt_up for t in terms)
                    is_comp = _is_companion_text(txt)
                    if is_wm:
                        wm_rects.append(fitz.Rect(span["bbox"]))
                        wm_terms_found.append(txt)
                    elif is_comp and wm_rects:
                        companion_rects.append(fitz.Rect(span["bbox"]))

        all_rects = wm_rects + companion_rects
        if all_rects:
            for r in all_rects:
                page.add_redact_annot(r, fill=(1, 1, 1))
            page.apply_redactions()
            page_removed += len(all_rects)
            log.append({
                "page": page_num + 1,
                "type": "ok",
                "msg": f"Redacted {len(all_rects)} text span(s): {', '.join(set(wm_terms_found))}"
            })

        # ── Method 2: remove XObject overlays (transparent Form XObjects) ─────
        try:
            xrefs = [
                page.get_contents()[i]
                for i in range(len(page.get_contents()))
            ]
            for xref in doc.get_page_xobjects(page_num):
                xobj_xref = xref[0]
                try:
                    xobj_str = doc.xref_stream(xobj_xref)
                    if xobj_str is None:
                        continue
                    xobj_text = xobj_str.decode('latin-1', errors='replace')
                    xobj_dict = doc.xref_object(xobj_xref)

                    is_form = "/Form" in xobj_dict
                    has_opacity = bool(re.search(r'/ca\s+[0-9.]+\s', xobj_dict, re.IGNORECASE))
                    has_multiply = "/Multiply" in xobj_dict
                    wm_in_stream = _stream_contains_watermark(xobj_str, terms)
                    only_wm = _stream_is_only_watermark(xobj_str, terms)

                    if is_form and (has_opacity or has_multiply) and wm_in_stream and only_wm:
                        doc.update_stream(xobj_xref, b"")
                        page_removed += 1
                        log.append({
                            "page": page_num + 1,
                            "type": "ok",
                            "msg": "Removed transparent XObject overlay"
                        })
                except Exception:
                    pass
        except Exception:
            pass

        if page_removed == 0:
            log.append({"page": page_num + 1, "type": "skip", "msg": "No text watermark found"})
        else:
            pages_cleaned += 1

    out = io.BytesIO()
    doc.save(out, garbage=4, deflate=True)
    doc.close()
    return out.getvalue(), pages_cleaned, log


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN REMOVAL ORCHESTRATOR
# ═══════════════════════════════════════════════════════════════════════════════

def remove_watermark(pdf_bytes, custom_term=None):
    """
    Master function — tries image watermark removal first (Variants A & B),
    then text watermark removal (Variant C).

    Returns:
        (cleaned_bytes, total_removed_count, log_list, method_str)
    """
    log = []
    total_removed = 0

    # Pass 1 — image-based watermark (pikepdf)
    cleaned_bytes, img_pages, wm_names = _remove_image_watermarks(pdf_bytes)
    if img_pages > 0:
        total_removed += img_pages
        log.append({
            "page": 0,
            "type": "ok",
            "msg": f"Removed full-page image overlay ({', '.join(wm_names)}) from {img_pages} page(s)"
        })
        method = f"image-overlay ({', '.join(wm_names)})"
    else:
        cleaned_bytes = pdf_bytes
        method = "none"

    # Pass 2 — text-based watermark (fitz/PyMuPDF)
    cleaned_bytes2, txt_pages, txt_log = _remove_text_watermarks(cleaned_bytes, custom_term)
    log.extend(txt_log)
    if txt_pages > 0:
        total_removed += txt_pages
        method = "text-layer" if img_pages == 0 else method + " + text-layer"
        cleaned_bytes = cleaned_bytes2
    elif img_pages == 0:
        # No removal at all — return original
        cleaned_bytes = pdf_bytes

    return cleaned_bytes, total_removed, log, method


# ═══════════════════════════════════════════════════════════════════════════════
#  FLASK ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RBI PDF Watermark Remover</title>
<style>
  :root {
    --bg: #0f1117; --surface: #1a1d27; --border: #2e3148;
    --accent: #4f8ef7; --accent2: #7c5df9;
    --text: #e8eaf6; --muted: #8b90b8; --ok: #4caf7d; --warn: #f7a84f;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--text); font-family: 'Segoe UI', sans-serif;
         min-height: 100vh; display: flex; flex-direction: column; align-items: center;
         padding: 24px 16px; }
  h1  { font-size: 1.5rem; font-weight: 700; margin-bottom: 4px;
        background: linear-gradient(90deg, var(--accent), var(--accent2));
        -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
  .sub { color: var(--muted); font-size: .85rem; margin-bottom: 28px; text-align: center; }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 14px;
          padding: 28px 24px; width: 100%; max-width: 520px; margin-bottom: 20px; }
  .drop-zone { border: 2px dashed var(--border); border-radius: 10px; padding: 36px 20px;
               text-align: center; cursor: pointer; transition: border-color .2s, background .2s; }
  .drop-zone:hover, .drop-zone.over { border-color: var(--accent); background: rgba(79,142,247,.06); }
  .drop-zone input[type=file] { display: none; }
  .drop-zone svg { width: 44px; height: 44px; stroke: var(--accent); margin-bottom: 12px; }
  .drop-zone p { color: var(--muted); font-size: .9rem; }
  .drop-zone p strong { color: var(--text); }
  #file-name { margin-top: 10px; font-size: .85rem; color: var(--ok); min-height: 18px; }
  .opt-row { display: flex; gap: 10px; margin-top: 16px; }
  .opt-row input { flex: 1; background: var(--bg); border: 1px solid var(--border);
                   border-radius: 8px; padding: 9px 12px; color: var(--text); font-size: .9rem; }
  .opt-row input::placeholder { color: var(--muted); }
  .btn { width: 100%; margin-top: 18px; padding: 13px; border: none; border-radius: 10px;
         background: linear-gradient(90deg, var(--accent), var(--accent2));
         color: #fff; font-size: 1rem; font-weight: 600; cursor: pointer; transition: opacity .2s; }
  .btn:disabled { opacity: .45; cursor: default; }
  #progress { display: none; margin-top: 14px; text-align: center; color: var(--muted); font-size:.9rem; }
  #progress .spinner { display: inline-block; width: 18px; height: 18px; border: 3px solid var(--border);
                       border-top-color: var(--accent); border-radius: 50%;
                       animation: spin .8s linear infinite; margin-right: 8px; vertical-align: middle; }
  @keyframes spin { to { transform: rotate(360deg); } }

  .result-card { display: none; }
  .result-card.show { display: block; }
  .result-header { display: flex; align-items: center; gap: 12px; margin-bottom: 16px; }
  .badge { width: 40px; height: 40px; border-radius: 50%; display: flex; align-items: center;
           justify-content: center; font-size: 1.3rem; flex-shrink: 0; }
  .badge.ok   { background: rgba(76,175,125,.15); }
  .badge.warn { background: rgba(247,168,79,.15); }
  .result-title { font-weight: 700; font-size: 1rem; }
  .result-sub   { color: var(--muted); font-size: .8rem; margin-top: 2px; }
  .log-list { list-style: none; max-height: 200px; overflow-y: auto; margin-bottom: 16px; }
  .log-item { display: flex; align-items: flex-start; gap: 8px; padding: 5px 0;
              border-bottom: 1px solid var(--border); font-size: .82rem; color: var(--muted); }
  .log-dot  { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; margin-top: 4px; }
  .log-dot.ok   { background: var(--ok); }
  .log-dot.skip { background: var(--border); }
  .log-dot.warn { background: var(--warn); }
  .log-pg { color: var(--text); font-weight: 600; min-width: 34px; }
  .dl-btn { display: flex; align-items: center; justify-content: center; gap: 8px;
            width: 100%; padding: 12px; border-radius: 10px; text-decoration: none;
            background: var(--ok); color: #fff; font-weight: 600; font-size: .95rem;
            transition: opacity .2s; }
  .dl-btn.dim { background: var(--border); color: var(--muted); pointer-events: none; }
  .dl-btn svg { width: 18px; height: 18px; stroke: currentColor; }
  .method-tag { display: inline-block; background: rgba(79,142,247,.12); color: var(--accent);
                border-radius: 6px; padding: 2px 8px; font-size: .75rem; margin-top: 4px; }
</style>
</head>
<body>

<h1>RBI PDF Watermark Remover</h1>
<p class="sub">Removes per-user tracking watermarks from RBI documents</p>

<div class="card">
  <div class="drop-zone" id="dropZone">
    <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/><line x1="12" y1="18" x2="12" y2="12"/>
      <line x1="9" y1="15" x2="15" y2="15"/>
    </svg>
    <p><strong>Tap to choose PDF</strong> or drag &amp; drop</p>
    <p style="font-size:.78rem;margin-top:4px">Max 50 MB</p>
    <input type="file" id="fileInput" accept=".pdf">
    <div id="file-name"></div>
  </div>
  <div class="opt-row">
    <input type="text" id="customTerm" placeholder="Custom watermark word (optional)">
  </div>
  <button class="btn" id="submitBtn" disabled>Remove Watermark</button>
  <div id="progress"><span class="spinner"></span>Processing, please wait…</div>
</div>

<div class="card result-card" id="resultCard">
  <div class="result-header">
    <div class="badge" id="resultBadge">✅</div>
    <div>
      <div class="result-title" id="resultTitle"></div>
      <div class="result-sub"  id="resultSub"></div>
      <div id="methodTag"></div>
    </div>
  </div>
  <ul class="log-list" id="logList"></ul>
  <a class="dl-btn" id="downloadBtn" href="#" download>
    <svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
    </svg>
    Download Cleaned PDF
  </a>
</div>

<script>
const dropZone  = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileLabel = document.getElementById('file-name');
const submitBtn = document.getElementById('submitBtn');
const progress  = document.getElementById('progress');
const resultCard = document.getElementById('resultCard');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('over');
  if (e.dataTransfer.files[0]) { fileInput.files = e.dataTransfer.files; onFileChosen(); }
});
fileInput.addEventListener('change', onFileChosen);

function onFileChosen() {
  const f = fileInput.files[0];
  if (!f) return;
  fileLabel.textContent = f.name + '  (' + (f.size/1024/1024).toFixed(1) + ' MB)';
  submitBtn.disabled = false;
}

submitBtn.addEventListener('click', async () => {
  const f = fileInput.files[0];
  if (!f) return;

  submitBtn.disabled = true;
  progress.style.display = 'block';
  resultCard.classList.remove('show');

  const fd = new FormData();
  fd.append('pdf', f);
  const ct = document.getElementById('customTerm').value.trim();
  if (ct) fd.append('custom_term', ct);

  try {
    const resp = await fetch('/remove', { method: 'POST', body: fd });
    const data = await resp.json();
    progress.style.display = 'none';
    submitBtn.disabled = false;
    showResult(data, f.name);
  } catch(err) {
    progress.style.display = 'none';
    submitBtn.disabled = false;
    alert('Upload failed: ' + err.message);
  }
});

function showResult(data, origName) {
  resultCard.classList.add('show');
  const ok = data.count > 0;
  const badge = document.getElementById('resultBadge');
  badge.textContent = ok ? '✅' : '⚠️';
  badge.className = 'badge ' + (ok ? 'ok' : 'warn');

  document.getElementById('resultTitle').textContent =
    ok ? data.count + ' page(s) cleaned' : 'No watermark detected';
  document.getElementById('resultSub').textContent = data.filename || '';

  const mt = document.getElementById('methodTag');
  if (data.method && data.method !== 'none') {
    mt.innerHTML = '<span class="method-tag">Method: ' + esc(data.method) + '</span>';
  } else { mt.innerHTML = ''; }

  const ul = document.getElementById('logList');
  ul.innerHTML = '';
  (data.log || []).filter(i => i.page > 0).forEach(item => {
    const li = document.createElement('li');
    li.className = 'log-item';
    li.innerHTML =
      '<div class="log-dot ' + item.type + '"></div>' +
      '<span class="log-pg">Pg ' + item.page + '</span>' +
      '<span>' + esc(item.msg) + '</span>';
    ul.appendChild(li);
  });

  const dlBtn = document.getElementById('downloadBtn');
  dlBtn.href = '/download/' + data.token + '?name=' + encodeURIComponent(data.filename || 'clean.pdf');
  dlBtn.className = 'dl-btn' + (ok ? '' : ' dim');
  resultCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>
"""

# ── Token store (temp-file based, survives gunicorn multi-worker) ─────────────

def _make_token(data: bytes) -> str:
    h = hashlib.sha256(data).hexdigest()[:16]
    token = f"wm_{h}_{int(time.time())}"
    path = os.path.join(TEMP_DIR, f"{token}.pdf")
    with open(path, "wb") as f:
        f.write(data)
    return token

def _get_token_path(token: str):
    if not re.match(r'^wm_[0-9a-f]{16}_\d+$', token):
        return None
    path = os.path.join(TEMP_DIR, f"{token}.pdf")
    return path if os.path.exists(path) else None


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return HTML_PAGE

@app.route("/remove", methods=["POST"])
def remove():
    if "pdf" not in request.files or not request.files["pdf"].filename:
        return jsonify({"error": "No PDF uploaded"}), 400

    file = request.files["pdf"]
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files accepted"}), 400

    try:
        pdf_bytes = file.read()
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 400

    custom_term = request.form.get("custom_term", "").strip() or None

    try:
        cleaned_bytes, count, log, method = remove_watermark(pdf_bytes, custom_term)
    except Exception as e:
        return jsonify({"error": f"Processing error: {e}"}), 500

    token = _make_token(cleaned_bytes)
    base = os.path.splitext(file.filename)[0]
    filename = f"{base}_clean.pdf"

    return jsonify({
        "count":    count,
        "log":      log,
        "method":   method,
        "token":    token,
        "filename": filename,
    })

@app.route("/download/<token>")
def download(token):
    path = _get_token_path(token)
    if not path:
        return "File not found or expired. Please process again.", 404

    filename = request.args.get("name", "cleaned.pdf")

    response = send_file(
        path,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename,
    )

    @response.call_on_close
    def _cleanup():
        try:
            os.remove(path)
        except Exception:
            pass

    return response


if __name__ == "__main__":
    print("\n🚀  RBI PDF Watermark Remover is running!")
    print("   Open: http://localhost:5000\n")
    app.run(debug=False, host="0.0.0.0", port=5000)
