"""
RBI Watermark Remover - Flask Web App
======================================

Removes per-user tracking watermarks from RBI scanned PDFs that follow
the structure produced by the EKP/DMS download system:

    Page Layer 1 (bottom): Clean scan of the document (Canon SC1011 JPEG)
    Page Layer 2 (top):    Watermark image "Im2" with a soft mask containing
                           the user ID and download timestamp text

This script removes Layer 2 from every page, leaving Layer 1 untouched.

VBA analogy:
    Each PDF page is like an Excel worksheet with two stacked Pictures.
    We loop through every "sheet", find the Picture named "Im2",
    and delete it. The original scan picture stays untouched.
"""

from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import pikepdf
import re
import io
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-key")

# Limit uploads to 50 MB - RBI circulars are usually 1-10 MB
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# The name of the watermark image in the PDF resources.
# This is what we identified by inspecting the PDF structure.
WATERMARK_IMAGE_NAME = "/Im2"


def remove_watermark(pdf_bytes: bytes) -> tuple[bytes, dict]:
    """
    Take raw PDF bytes, remove the watermark, return cleaned PDF bytes
    plus a small report dict (how many pages, whether watermark was found).

    VBA analogy: think of this as a Function that takes a Workbook,
    deletes a specific Picture from every Sheet, and returns the modified
    Workbook plus a status message.
    """
    # Open the PDF from memory (not from disk - safer for a web app)
    pdf = pikepdf.open(io.BytesIO(pdf_bytes))

    pages_cleaned = 0
    total_pages = len(pdf.pages)

    for page in pdf.pages:
        watermark_found = False

        # Step 1: Remove the watermark image from the page's resource list.
        # In VBA terms: delete the Shape from the worksheet's Shapes collection.
        try:
            resources = page.get("/Resources", {})
            if "/XObject" in resources:
                xobjects = resources["/XObject"]
                if WATERMARK_IMAGE_NAME in xobjects:
                    del xobjects[WATERMARK_IMAGE_NAME]
                    watermark_found = True
        except Exception:
            # If a page has unusual structure, skip it rather than crash.
            pass

        # Step 2: Remove the drawing instruction from the page's content stream.
        # The content stream is a script the PDF reader runs to draw the page.
        # The watermark is drawn by a block like:  q ... /Im2 Do ... Q
        # (q = save state, /Im2 Do = draw the image, Q = restore state)
        try:
            contents = page["/Contents"]
            # Some PDFs store contents as a single stream, others as an array.
            # Handle both.
            if isinstance(contents, pikepdf.Array):
                streams = list(contents)
            else:
                streams = [contents]

            for stream in streams:
                data = stream.read_bytes().decode('latin-1', errors='ignore')
                # Regex meaning: find "q" then anything (non-greedy) then
                # "/Im2 Do" then anything then "Q" - and remove that whole block.
                cleaned = re.sub(
                    r'q\s+[^q]*?' + re.escape(WATERMARK_IMAGE_NAME) + r'\s+Do\s+Q',
                    '',
                    data,
                    flags=re.DOTALL
                )
                if cleaned != data:
                    stream.write(cleaned.encode('latin-1'))
                    watermark_found = True
        except Exception:
            pass

        if watermark_found:
            pages_cleaned += 1

    # Save the modified PDF to memory (not disk).
    output = io.BytesIO()
    pdf.save(output)
    output.seek(0)

    report = {
        "total_pages": total_pages,
        "pages_cleaned": pages_cleaned,
        "success": pages_cleaned > 0,
    }
    return output.read(), report


# =====================================================================
# Flask routes - the URLs that the browser will hit
# =====================================================================

# HTML template kept inline so this is a single-file app.
# (Like keeping a UserForm and its code in the same VBA module.)
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
        input[type=file] { margin: 12px 0; }
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
    <p>Upload an RBI PDF that has the diagonal user-ID + timestamp watermark.
       The tool removes the watermark layer and returns the cleaned scan.</p>

    {% with messages = get_flashed_messages() %}
      {% if messages %}
        {% for msg in messages %}<div class="flash">{{ msg }}</div>{% endfor %}
      {% endif %}
    {% endwith %}

    <div class="upload-box">
        <form method="POST" action="/clean" enctype="multipart/form-data">
            <input type="file" name="pdf" accept=".pdf" required>
            <br>
            <button type="submit">Remove Watermark</button>
        </form>
    </div>

    <div class="note">
        <b>Note:</b> The watermark contains a user ID that traces the document
        to a specific person. Use this only on PDFs you have a legitimate
        reason to clean (e.g. your own downloads where the watermark is
        obstructing readability). Do not redistribute cleaned copies of
        documents downloaded under another person's ID.
    </div>
</body>
</html>
"""


@app.route("/", methods=["GET"])
def home():
    """Show the upload form."""
    return render_template_string_inline(HOME_PAGE)


@app.route("/clean", methods=["POST"])
def clean():
    """Receive the uploaded PDF, clean it, send back the result."""
    # Step 1: Validate the upload.
    if "pdf" not in request.files:
        flash("No file uploaded.")
        return redirect(url_for("home"))

    file = request.files["pdf"]
    if file.filename == "":
        flash("No file selected.")
        return redirect(url_for("home"))

    if not file.filename.lower().endswith(".pdf"):
        flash("Only PDF files are accepted.")
        return redirect(url_for("home"))

    # Step 2: Read the file into memory and run the cleaner.
    try:
        pdf_bytes = file.read()
        cleaned_bytes, report = remove_watermark(pdf_bytes)
    except pikepdf.PasswordError:
        flash("This PDF requires a password. Decrypt it first using qpdf.")
        return redirect(url_for("home"))
    except Exception as e:
        flash(f"Could not process the PDF. Error: {e}")
        return redirect(url_for("home"))

    # Step 3: If we did not find a watermark, warn the user.
    if not report["success"]:
        flash(
            f"No '{WATERMARK_IMAGE_NAME}' watermark layer was found in this PDF. "
            f"The file may use a different watermark format. The download below "
            f"is the original PDF unchanged."
        )

    # Step 4: Build a sensible output filename.
    original_name = secure_filename(file.filename)
    base = os.path.splitext(original_name)[0]
    output_name = f"{base}_clean.pdf"

    # Step 5: Send the cleaned PDF back to the browser as a download.
    return send_file(
        io.BytesIO(cleaned_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=output_name,
    )


def render_template_string_inline(template):
    """Tiny wrapper so we don't need a templates/ folder for one page."""
    from flask import render_template_string
    return render_template_string(template)


if __name__ == "__main__":
    # debug=True gives you the auto-reload + error page while developing.
    # Turn it off in production.
    app.run(host="0.0.0.0", port=5000, debug=True)
