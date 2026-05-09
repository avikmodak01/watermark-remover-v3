"""
RBI Watermark Remover - Flask Web App  (v2 - auto-detect)
=========================================================

Removes per-user tracking watermarks from RBI PDFs that are issued
through EKP / DMS. Handles two structures we have seen so far:

    Variant A:  Scanned PDFs (Canon SC1011)
                Watermark = a separate image overlay (e.g. /Im2)
                with a soft-mask containing the text.

    Variant B:  Word-generated PDFs (Microsoft Word)
                Watermark = a full-page-sized image (e.g. /Im1, /Im4)
                drawn at the bottom of every page's content stream.

Instead of hardcoding the image name, this version *detects* the
watermark by looking for any image drawn at the page's full size.
Real document content almost never covers the entire page exactly,
so a full-page image draw is a reliable watermark signature.

VBA analogy:
    Each PDF page is like a worksheet with several Pictures on it.
    We scan every page, find any Picture whose width and height
    match the page exactly, and remove it. The other Pictures
    (logos, signatures, real photos) are smaller, so they stay.
"""

from flask import Flask, request, send_file, flash, redirect, url_for, render_template_string
import pikepdf
import re
import io
import os
from collections import defaultdict
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-key")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

# Tolerance (in PDF points) for matching "full-page" image dimensions.
# 5 points = ~1.7 mm, more than enough to absorb rounding errors.
SIZE_TOLERANCE = 5.0

# Regex: matches a "draw image at <w> x <h> position 0,0" instruction.
# In PDF content streams an image is drawn by:
#     <width> 0 0 <height> <x> <y> cm  /<name> Do
# We only care about ones drawn at origin (0,0) at full page size.
FULLPAGE_DRAW = re.compile(
    r'([\d.]+)\s+0\s+0\s+([\d.]+)\s+0\s+0\s+cm\s*/(\w+)\s+Do'
)


def get_page_size(page) -> tuple[float, float]:
    """Return (width, height) of a page in PDF points."""
    box = page.MediaBox
    return float(box[2]) - float(box[0]), float(box[3]) - float(box[1])


def find_watermark_images(pdf: pikepdf.Pdf) -> set[str]:
    """
    Scan every page. Return the set of image names that are drawn
    at full page size on at least one page. Those are the watermarks.

    VBA analogy: loop through every Sheet's Shapes, note the names of
    Pictures whose Width = Sheet width and Height = Sheet height.
    Those are the watermarks.
    """
    appearances: dict[str, int] = defaultdict(int)

    for page in pdf.pages:
        try:
            page_w, page_h = get_page_size(page)
        except Exception:
            continue

        # A page's content can be one stream or an array of streams.
        contents = page.get("/Contents")
        if contents is None:
            continue
        streams = list(contents) if isinstance(contents, pikepdf.Array) else [contents]

        seen_on_this_page: set[str] = set()
        for stream in streams:
            try:
                data = stream.read_bytes().decode("latin-1", errors="ignore")
            except Exception:
                continue
            for w_str, h_str, name in FULLPAGE_DRAW.findall(data):
                w, h = float(w_str), float(h_str)
                if (abs(w - page_w) < SIZE_TOLERANCE
                        and abs(h - page_h) < SIZE_TOLERANCE):
                    seen_on_this_page.add(name)

        for name in seen_on_this_page:
            appearances[name] += 1

    # Any image drawn full-page on at least one page is a watermark.
    # Different PDFs use different watermark images on cover vs. body
    # pages, so we don't require a high page-count threshold.
    return set(appearances.keys())


def remove_watermark(pdf_bytes: bytes) -> tuple[bytes, dict]:
    """
    Open the PDF, find watermark images, strip them from every page they
    appear on, and return the cleaned bytes plus a report.
    """
    pdf = pikepdf.open(io.BytesIO(pdf_bytes))
    total_pages = len(pdf.pages)

    # Step 1: detect which image names are watermarks.
    watermark_names = find_watermark_images(pdf)

    if not watermark_names:
        return pdf_bytes, {
            "total_pages": total_pages,
            "pages_cleaned": 0,
            "watermark_names": [],
            "success": False,
        }

    # Step 2: remove the draw operations and resource references.
    pages_cleaned = 0

    for page in pdf.pages:
        page_changed = False

        # Remove the draw operations from the content stream(s).
        contents = page.get("/Contents")
        if contents is not None:
            streams = list(contents) if isinstance(contents, pikepdf.Array) else [contents]
            for stream in streams:
                try:
                    data = stream.read_bytes().decode("latin-1", errors="ignore")
                except Exception:
                    continue

                original = data
                for name in watermark_names:
                    name_re = re.escape(name)
                    # Two patterns: with the optional /Artifact BMC...EMC wrapper
                    # that Word emits, and without (older Canon scans).
                    patterns = [
                        # Wrapped form: /Artifact BMC q ... /Name Do Q EMC
                        rf'/Artifact\s*BMC\s*q\s+[\d.]+\s+0\s+0\s+[\d.]+\s+0\s+0\s+cm\s*/{name_re}\s+Do\s*Q\s*EMC',
                        # Plain form: q ... /Name Do Q
                        rf'q\s+[^q]*?/{name_re}\s+Do\s*Q',
                    ]
                    for pat in patterns:
                        data = re.sub(pat, '', data, flags=re.DOTALL)

                if data != original:
                    stream.write(data.encode("latin-1"))
                    page_changed = True

        # Remove the watermark images from the resource dictionary too.
        resources = page.get("/Resources")
        if resources is not None and "/XObject" in resources:
            xobjects = resources["/XObject"]
            for name in watermark_names:
                key = f"/{name}"
                if key in xobjects:
                    del xobjects[key]
                    page_changed = True

        if page_changed:
            pages_cleaned += 1

    output = io.BytesIO()
    pdf.save(output)
    output.seek(0)

    return output.read(), {
        "total_pages": total_pages,
        "pages_cleaned": pages_cleaned,
        "watermark_names": sorted(watermark_names),
        "success": pages_cleaned > 0,
    }


# =====================================================================
# Web UI
# =====================================================================

HOME_PAGE = """
<!DOCTYPE html>
<html>
<head>
  <title>RBI Watermark Remover</title>
  <style>
    body { font-family: -apple-system, Arial, sans-serif; max-width: 640px;
           margin: 40px auto; padding: 0 20px; color: #222; }
    h1 { color: #003366; border-bottom: 2px solid #003366; padding-bottom: 8px; }
    .upload-box { border: 2px dashed #003366; border-radius: 8px;
                  padding: 30px; text-align: center; background: #f4f7fb; }
    button { background: #003366; color: white; border: 0; padding: 10px 24px;
             border-radius: 4px; font-size: 15px; cursor: pointer; }
    button:hover { background: #00509e; }
    .flash { background: #fff3cd; border: 1px solid #ffc107; padding: 10px;
             border-radius: 4px; margin: 12px 0; }
    .note { font-size: 13px; color: #666; margin-top: 24px;
            background: #fafafa; padding: 12px; border-left: 3px solid #999; }
  </style>
</head>
<body>
  <h1>RBI Watermark Remover</h1>
  <p>Upload an RBI PDF that has the diagonal user-ID watermark.
     The tool detects the watermark layer automatically and removes it,
     whether the PDF is a Canon scan or a Word-generated document.</p>

  {% with messages = get_flashed_messages() %}
    {% if messages %}
      {% for msg in messages %}<div class="flash">{{ msg }}</div>{% endfor %}
    {% endif %}
  {% endwith %}

  <div class="upload-box">
    <form method="POST" action="/clean" enctype="multipart/form-data">
      <input type="file" name="pdf" accept=".pdf" required><br>
      <button type="submit">Remove Watermark</button>
    </form>
  </div>

  <div class="note">
    <b>Note:</b> The watermark contains a user ID that traces the document
    to a specific person. Use this only on PDFs you have a legitimate
    reason to clean, and do not redistribute cleaned copies of documents
    that were downloaded under another person's ID.
  </div>
</body>
</html>
"""


@app.route("/", methods=["GET"])
def home():
    return render_template_string(HOME_PAGE)


@app.route("/clean", methods=["POST"])
def clean():
    if "pdf" not in request.files or request.files["pdf"].filename == "":
        flash("Please choose a PDF file.")
        return redirect(url_for("home"))

    file = request.files["pdf"]
    if not file.filename.lower().endswith(".pdf"):
        flash("Only PDF files are accepted.")
        return redirect(url_for("home"))

    try:
        pdf_bytes = file.read()
        cleaned_bytes, report = remove_watermark(pdf_bytes)
    except pikepdf.PasswordError:
        flash("This PDF requires a password. Decrypt it first using qpdf.")
        return redirect(url_for("home"))
    except Exception as e:
        flash(f"Could not process the PDF. Error: {e}")
        return redirect(url_for("home"))

    if not report["success"]:
        flash(
            "No watermark layer was detected in this PDF. The download below "
            "is the original file unchanged. If you believe a watermark is "
            "present, send the file to the maintainer for analysis."
        )

    base = os.path.splitext(secure_filename(file.filename))[0]
    return send_file(
        io.BytesIO(cleaned_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"{base}_clean.pdf",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
