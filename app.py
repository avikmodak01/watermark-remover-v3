"""
RBI PDF Watermark Remover — Flask Web App  (v4 - pymupdf only)
==============================================================
Local:   python app.py  →  http://localhost:5000
Deploy:  Render.com  (see render.yaml)

Dependencies: flask, pymupdf  — NO pikepdf, NO C build tools needed.
pymupdf ships as a pre-compiled wheel so it installs on any platform
including Render.com free tier without any system packages.

Handles all known RBI watermark variants:

  Variant A/B — Full-page image overlay  (e.g. /Im1, /Im2)
    Word-generated and scanned PDFs where a full-page image is drawn
    on top of every page via a content stream instruction /ImN Do.
    Detected by: image dimensions ≥85% of page size, present on ≥50% of pages.
    Removed by:  clearing the draw instruction from the content stream
                 and deleting the entry from the page's XObject resource dict.

  Variant C — Text-layer watermark  (diagonal repeated text)
    Watermark text is in the PDF text layer (can be selected/copied).
    Removed by:  redacting matching text spans + clearing transparent
                 Form XObject overlays that contain the watermark text.
"""

import os, re, io, hashlib, time, tempfile
import fitz          # PyMuPDF — handles everything, no pikepdf needed
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024   # 50 MB
app.secret_key = os.environ.get('SECRET_KEY', 'rbi-wm-remover-2025')

TEMP_DIR = tempfile.gettempdir()

# ─────────────────────────────────────────────────────────────────────────────
#  VARIANT A/B — Full-page image watermark  (using fitz only)
# ─────────────────────────────────────────────────────────────────────────────

def _find_fullpage_image_names(doc):
    """
    Scan all pages for XObject images that are ≥85% of the page size
    AND appear on ≥50% of pages. That combination uniquely identifies
    a watermark overlay — real content images are never full-page and
    never repeat on every single page.

    Returns a set of bare image names e.g. {'Im1', 'Im2'}.

    VBA analogy: like scanning every worksheet for a Picture whose
    Width and Height match the sheet's PrintArea exactly.
    """
    if len(doc) == 0:
        return set()

    # Use page 0 dimensions as reference
    ref_w = doc[0].rect.width
    ref_h = doc[0].rect.height
    if ref_w == 0 or ref_h == 0:
        return set()

    name_count = {}   # name -> how many pages it appears on
    name_xref  = {}   # name -> xref (for dimension lookup)

    for pnum in range(len(doc)):
        page = doc[pnum]
        try:
            res_val = doc.xref_get_key(page.xref, "Resources")
            res_text = res_val[1] if res_val else ""
        except Exception:
            continue

        # Find all /XObject<< /Name NNN 0 R ... >> entries
        xobj_m = re.search(r'/XObject\s*<<([^>]*)>>', res_text)
        if not xobj_m:
            continue

        for name, xref_str in re.findall(r'/(\w+)\s+(\d+)\s+0\s+R',
                                          xobj_m.group(1)):
            xref = int(xref_str)
            try:
                obj_dict = doc.xref_object(xref)
            except Exception:
                continue

            # Must be an Image subtype
            if '/Image' not in obj_dict:
                continue

            w_m = re.search(r'/Width\s+(\d+)', obj_dict)
            h_m = re.search(r'/Height\s+(\d+)', obj_dict)
            if not (w_m and h_m):
                continue

            w, h = int(w_m.group(1)), int(h_m.group(1))
            if (w / ref_w >= 0.85) and (h / ref_h >= 0.85):
                name_count[name] = name_count.get(name, 0) + 1
                name_xref[name] = xref

    threshold = max(2, len(doc) * 0.5)
    return {n for n, cnt in name_count.items() if cnt >= threshold}


def _remove_image_watermarks(doc, wm_names):
    """
    Given a set of image names (e.g. {'Im1'}), remove them from every page:
      1. Erase the draw instruction  q ... /ImN Do ... Q  in content streams.
      2. Delete the /ImN entry from the page's Resources/XObject dict.

    Returns count of pages modified.
    """
    if not wm_names:
        return 0

    pages_cleaned = 0
    for pnum in range(len(doc)):
        page = doc[pnum]
        modified = False

        # ── Step 1: strip draw instruction from content streams ────────────
        for xref in page.get_contents():
            try:
                data = doc.xref_stream(xref)
                if not data:
                    continue
                text = data.decode('latin-1', errors='replace')
                original = text
                for name in wm_names:
                    # Matches: q [any transforms] /ImN Do [anything] Q
                    text = re.sub(
                        rf'q\b[^Q]*?/{re.escape(name)}\s+Do[^Q]*?Q\b',
                        '',
                        text,
                        flags=re.DOTALL
                    )
                if text != original:
                    doc.update_stream(xref, text.encode('latin-1'))
                    modified = True
            except Exception:
                pass

        # ── Step 2: remove from Resources/XObject dict ────────────────────
        try:
            res_val = doc.xref_get_key(page.xref, "Resources")
            if not res_val:
                continue
            res_text = res_val[1]
            new_res = res_text
            for name in wm_names:
                new_res = re.sub(
                    rf'/{re.escape(name)}\s+\d+\s+0\s+R',
                    '',
                    new_res
                )
            if new_res != res_text:
                doc.xref_set_key(page.xref, "Resources", new_res)
                modified = True
        except Exception:
            pass

        if modified:
            pages_cleaned += 1

    return pages_cleaned


# ─────────────────────────────────────────────────────────────────────────────
#  VARIANT C — Text-layer watermark helpers
# ─────────────────────────────────────────────────────────────────────────────

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

def _is_companion(text):
    return any(re.match(p, text.strip(), re.IGNORECASE) for p in COMPANION_PATTERNS)

def _stream_has_watermark(stream_bytes, terms):
    try:
        text = stream_bytes.decode('latin-1', errors='replace').upper()
        return any(t in text for t in terms)
    except Exception:
        return False

def _stream_only_watermark(stream_bytes, terms):
    """True if every text literal in the stream is a watermark term."""
    try:
        text = stream_bytes.decode('latin-1', errors='replace')
    except Exception:
        return False
    bt_blocks = re.findall(r'BT\b(.*?)\bET', text, re.DOTALL)
    if not bt_blocks:
        return False
    terms_upper = [t.upper() for t in terms]
    for block in bt_blocks:
        for lit in re.findall(r'\(((?:[^()\\]|\\.)*)\)', block):
            lit_clean = re.sub(r'\\(.)', r'\1', lit).strip()
            if not lit_clean:
                continue
            if not any(t in lit_clean.upper() for t in terms_upper):
                return False
    return True


def _remove_text_watermarks(doc, terms, log):
    """
    Remove text-layer watermarks from every page of doc (in place).
    Appends entries to log. Returns count of pages modified.
    """
    pages_cleaned = 0

    for pnum in range(len(doc)):
        page = doc[pnum]
        page_removed = 0

        # ── Method 1: redact matching text spans ──────────────────────────
        blocks = page.get_text("dict",
                               flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        wm_rects    = []
        comp_rects  = []
        wm_found    = []

        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    txt = span["text"].strip()
                    if not txt:
                        continue
                    txt_up = txt.upper()
                    if any(t in txt_up for t in terms):
                        wm_rects.append(fitz.Rect(span["bbox"]))
                        wm_found.append(txt)
                    elif _is_companion(txt) and wm_rects:
                        comp_rects.append(fitz.Rect(span["bbox"]))

        all_rects = wm_rects + comp_rects
        if all_rects:
            for r in all_rects:
                page.add_redact_annot(r, fill=(1, 1, 1))
            page.apply_redactions()
            page_removed += len(all_rects)
            log.append({
                "page": pnum + 1, "type": "ok",
                "msg": f"Redacted {len(all_rects)} span(s): "
                       f"{', '.join(set(wm_found))}"
            })

        # ── Method 2: clear transparent Form XObject overlays ─────────────
        for item in doc.get_page_xobjects(pnum):
            xref = item[0]
            try:
                xobj_dict = doc.xref_object(xref)
                xobj_stream = doc.xref_stream(xref)
                if xobj_stream is None:
                    continue
                is_form      = "/Form" in xobj_dict
                has_opacity  = bool(re.search(r'/ca\s+[0-9.]+', xobj_dict,
                                              re.IGNORECASE))
                has_multiply = "/Multiply" in xobj_dict
                has_wm       = _stream_has_watermark(xobj_stream, terms)
                only_wm      = _stream_only_watermark(xobj_stream, terms)

                if is_form and (has_opacity or has_multiply) and has_wm and only_wm:
                    doc.update_stream(xref, b"")
                    page_removed += 1
                    log.append({
                        "page": pnum + 1, "type": "ok",
                        "msg": "Cleared transparent XObject overlay"
                    })
            except Exception:
                pass

        if page_removed == 0:
            log.append({
                "page": pnum + 1, "type": "skip",
                "msg": "No text watermark found"
            })
        else:
            pages_cleaned += 1

    return pages_cleaned


# ─────────────────────────────────────────────────────────────────────────────
#  MASTER ORCHESTRATOR
# ─────────────────────────────────────────────────────────────────────────────

def remove_watermark(pdf_bytes, custom_term=None):
    """
    Remove all watermarks from a PDF (bytes in, bytes out).

    Runs two passes:
      Pass 1 — full-page image overlay detection & removal  (Variant A/B)
      Pass 2 — text-layer watermark redaction               (Variant C)

    Returns:
        cleaned_bytes  : bytes
        total_removed  : int  (pages affected)
        log            : list of dicts  {page, type, msg}
        method         : str describing what was found
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    log = []
    total_removed = 0
    methods = []

    # ── Pass 1: image overlay ─────────────────────────────────────────────
    wm_names = _find_fullpage_image_names(doc)
    if wm_names:
        img_cleaned = _remove_image_watermarks(doc, wm_names)
        if img_cleaned > 0:
            total_removed += img_cleaned
            methods.append(f"image-overlay ({', '.join(sorted(wm_names))})")
            log.insert(0, {
                "page": 0, "type": "ok",
                "msg": (f"Removed full-page image overlay "
                        f"({', '.join(sorted(wm_names))}) "
                        f"from {img_cleaned} page(s)")
            })

    # ── Pass 2: text watermark ────────────────────────────────────────────
    terms = COMMON_WATERMARKS[:]
    if custom_term:
        terms.insert(0, custom_term.upper().strip())

    txt_cleaned = _remove_text_watermarks(doc, terms, log)
    if txt_cleaned > 0:
        total_removed += txt_cleaned
        methods.append("text-layer")

    # ── Save ──────────────────────────────────────────────────────────────
    out = io.BytesIO()
    doc.save(out, garbage=4, deflate=True)
    doc.close()

    method = " + ".join(methods) if methods else "none"
    return out.getvalue(), total_removed, log, method


# ─────────────────────────────────────────────────────────────────────────────
#  TOKEN STORE  (temp-file based — survives gunicorn multi-worker restarts)
# ─────────────────────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────────────────────
#  HTML  (single-file — no templates folder needed)
# ─────────────────────────────────────────────────────────────────────────────

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>RBI PDF Watermark Remover</title>
<style>
  :root {
    --bg:#0f1117;--surface:#1a1d27;--border:#2e3148;
    --accent:#4f8ef7;--accent2:#7c5df9;
    --text:#e8eaf6;--muted:#8b90b8;--ok:#4caf7d;--warn:#f7a84f;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--text);font-family:'Segoe UI',sans-serif;
       min-height:100vh;display:flex;flex-direction:column;align-items:center;
       padding:24px 16px}
  h1{font-size:1.5rem;font-weight:700;margin-bottom:4px;
     background:linear-gradient(90deg,var(--accent),var(--accent2));
     -webkit-background-clip:text;-webkit-text-fill-color:transparent}
  .sub{color:var(--muted);font-size:.85rem;margin-bottom:28px;text-align:center}
  .card{background:var(--surface);border:1px solid var(--border);border-radius:14px;
        padding:28px 24px;width:100%;max-width:520px;margin-bottom:20px}
  .drop-zone{border:2px dashed var(--border);border-radius:10px;padding:36px 20px;
             text-align:center;cursor:pointer;transition:border-color .2s,background .2s}
  .drop-zone:hover,.drop-zone.over{border-color:var(--accent);background:rgba(79,142,247,.06)}
  .drop-zone input[type=file]{display:none}
  .drop-zone svg{width:44px;height:44px;stroke:var(--accent);margin-bottom:12px}
  .drop-zone p{color:var(--muted);font-size:.9rem}
  .drop-zone p strong{color:var(--text)}
  #file-name{margin-top:10px;font-size:.85rem;color:var(--ok);min-height:18px}
  .opt-row{display:flex;gap:10px;margin-top:16px}
  .opt-row input{flex:1;background:var(--bg);border:1px solid var(--border);
                 border-radius:8px;padding:9px 12px;color:var(--text);font-size:.9rem}
  .opt-row input::placeholder{color:var(--muted)}
  .btn{width:100%;margin-top:18px;padding:13px;border:none;border-radius:10px;
       background:linear-gradient(90deg,var(--accent),var(--accent2));
       color:#fff;font-size:1rem;font-weight:600;cursor:pointer;transition:opacity .2s}
  .btn:disabled{opacity:.45;cursor:default}
  #progress{display:none;margin-top:14px;text-align:center;color:var(--muted);font-size:.9rem}
  #progress .spinner{display:inline-block;width:18px;height:18px;
                     border:3px solid var(--border);border-top-color:var(--accent);
                     border-radius:50%;animation:spin .8s linear infinite;
                     margin-right:8px;vertical-align:middle}
  @keyframes spin{to{transform:rotate(360deg)}}
  .result-card{display:none}
  .result-card.show{display:block}
  .result-header{display:flex;align-items:center;gap:12px;margin-bottom:16px}
  .badge{width:40px;height:40px;border-radius:50%;display:flex;align-items:center;
         justify-content:center;font-size:1.3rem;flex-shrink:0}
  .badge.ok{background:rgba(76,175,125,.15)}
  .badge.warn{background:rgba(247,168,79,.15)}
  .result-title{font-weight:700;font-size:1rem}
  .result-sub{color:var(--muted);font-size:.8rem;margin-top:2px}
  .log-list{list-style:none;max-height:200px;overflow-y:auto;margin-bottom:16px}
  .log-item{display:flex;align-items:flex-start;gap:8px;padding:5px 0;
            border-bottom:1px solid var(--border);font-size:.82rem;color:var(--muted)}
  .log-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;margin-top:4px}
  .log-dot.ok{background:var(--ok)}
  .log-dot.skip{background:var(--border)}
  .log-dot.warn{background:var(--warn)}
  .log-pg{color:var(--text);font-weight:600;min-width:34px}
  .dl-btn{display:flex;align-items:center;justify-content:center;gap:8px;
          width:100%;padding:12px;border-radius:10px;text-decoration:none;
          background:var(--ok);color:#fff;font-weight:600;font-size:.95rem;
          transition:opacity .2s}
  .dl-btn.dim{background:var(--border);color:var(--muted);pointer-events:none}
  .dl-btn svg{width:18px;height:18px;stroke:currentColor}
  .method-tag{display:inline-block;background:rgba(79,142,247,.12);color:var(--accent);
              border-radius:6px;padding:2px 8px;font-size:.75rem;margin-top:4px}
</style>
</head>
<body>

<h1>RBI PDF Watermark Remover</h1>
<p class="sub">Removes per-user tracking watermarks from RBI documents</p>

<div class="card">
  <div class="drop-zone" id="dropZone">
    <svg viewBox="0 0 24 24" fill="none" stroke-width="1.5"
         stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
      <polyline points="14 2 14 8 20 8"/>
      <line x1="12" y1="18" x2="12" y2="12"/>
      <line x1="9"  y1="15" x2="15" y2="15"/>
    </svg>
    <p><strong>Tap to choose PDF</strong> or drag &amp; drop</p>
    <p style="font-size:.78rem;margin-top:4px">Max 50 MB</p>
    <input type="file" id="fileInput" accept=".pdf">
    <div id="file-name"></div>
  </div>
  <div class="opt-row">
    <input type="text" id="customTerm"
           placeholder="Custom watermark word (optional)">
  </div>
  <button class="btn" id="submitBtn" disabled>Remove Watermark</button>
  <div id="progress">
    <span class="spinner"></span>Processing, please wait…
  </div>
</div>

<div class="card result-card" id="resultCard">
  <div class="result-header">
    <div class="badge" id="resultBadge">✅</div>
    <div>
      <div class="result-title" id="resultTitle"></div>
      <div class="result-sub"   id="resultSub"></div>
      <div id="methodTag"></div>
    </div>
  </div>
  <ul class="log-list" id="logList"></ul>
  <a class="dl-btn" id="downloadBtn" href="#" download>
    <svg viewBox="0 0 24 24" fill="none" stroke-width="2"
         stroke-linecap="round" stroke-linejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="7 10 12 15 17 10"/>
      <line x1="12" y1="15" x2="12" y2="3"/>
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
dropZone.addEventListener('dragover', e => {
  e.preventDefault(); dropZone.classList.add('over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.classList.remove('over');
  if (e.dataTransfer.files[0]) {
    fileInput.files = e.dataTransfer.files; onFileChosen();
  }
});
fileInput.addEventListener('change', onFileChosen);

function onFileChosen() {
  const f = fileInput.files[0];
  if (!f) return;
  fileLabel.textContent =
    f.name + '  (' + (f.size/1024/1024).toFixed(1) + ' MB)';
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
  badge.textContent  = ok ? '✅' : '⚠️';
  badge.className    = 'badge ' + (ok ? 'ok' : 'warn');
  document.getElementById('resultTitle').textContent =
    ok ? data.count + ' page(s) cleaned' : 'No watermark detected';
  document.getElementById('resultSub').textContent = data.filename || '';

  const mt = document.getElementById('methodTag');
  mt.innerHTML = (data.method && data.method !== 'none')
    ? '<span class="method-tag">Method: ' + esc(data.method) + '</span>'
    : '';

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
  dlBtn.href = '/download/' + data.token +
               '?name=' + encodeURIComponent(data.filename || 'clean.pdf');
  dlBtn.className = 'dl-btn' + (ok ? '' : ' dim');
  resultCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function esc(s) {
  return String(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
#  FLASK ROUTES
# ─────────────────────────────────────────────────────────────────────────────

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
        cleaned_bytes, count, log, method = remove_watermark(pdf_bytes,
                                                              custom_term)
    except Exception as e:
        return jsonify({"error": f"Processing error: {e}"}), 500

    token = _make_token(cleaned_bytes)
    base  = os.path.splitext(file.filename)[0]

    return jsonify({
        "count":    count,
        "log":      log,
        "method":   method,
        "token":    token,
        "filename": f"{base}_clean.pdf",
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
