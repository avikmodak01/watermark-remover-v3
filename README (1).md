# RBI Watermark Remover

A small Flask web app that removes the diagonal user-ID + timestamp
watermark from RBI scanned PDFs (the kind produced by the EKP/DMS
download system, e.g. circulars, notifications).

## How it works (in plain language)

Each page of the watermarked PDF is built from two stacked image layers:

1. **Bottom layer** = the clean scan of the document (a JPEG from the
   Canon scanner).
2. **Top layer** = a transparent image named `Im2` whose visible pixels
   form the watermark text (user ID + download time).

Because the two layers are separate, removing the top layer reveals the
clean scan underneath. This script does exactly that — no OCR, no
image inpainting, no quality loss.

## VBA analogy

If you opened the PDF as an Excel workbook:

- Every `Sheet` (page) has two `Pictures` on it.
- We loop through every sheet.
- We delete the picture named `"Im2"`.
- The other picture (the scan) stays in place.

That is the entire trick.

## Files

```
rbi_watermark_remover/
├── app.py              # The Flask app (single file, fully commented)
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

## Setup (local development)

```bash
# 1. Create a folder and drop the files in.
cd rbi_watermark_remover

# 2. Create a virtual environment (optional but recommended).
python -m venv venv
# Windows:    venv\Scripts\activate
# Mac/Linux:  source venv/bin/activate

# 3. Install dependencies.
pip install -r requirements.txt

# 4. Run the app.
python app.py
```

Open `http://localhost:5000` in your browser. Upload a PDF, get the
cleaned version back as a download.

## Deployment to Render.com

This app is small and stateless, so it runs comfortably on Render's
free tier — same setup as your earlier PDF tools.

1. Push this folder to a GitHub repo.
2. On Render, create a new **Web Service** pointing at the repo.
3. **Build command:** `pip install -r requirements.txt`
4. **Start command:** `gunicorn app:app`
5. Add `gunicorn>=21.0.0` to `requirements.txt` before deploying.
6. Set environment variable `FLASK_SECRET_KEY` to any random string.

In `app.py`, replace the hardcoded `app.secret_key = "change-me-..."`
line with:

```python
import os
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-only-key")
```

## How to extend it

The watermark name `/Im2` is hardcoded in `WATERMARK_IMAGE_NAME` at the
top of `app.py`. If you encounter another RBI document where the
watermark image has a different name, you have two options:

### Option A — Inspect the new file and update the constant

Run this Python snippet on the new PDF to find the watermark image name:

```python
import pikepdf
pdf = pikepdf.open("new_file.pdf")
for i, page in enumerate(pdf.pages):
    xobjects = page.get("/Resources", {}).get("/XObject", {})
    print(f"Page {i+1}: {list(xobjects.keys())}")
```

The watermark is usually the image that appears with the **same name on
every page** (because the same overlay is reused). Change
`WATERMARK_IMAGE_NAME` in `app.py` to that name.

### Option B — Auto-detect

A more robust version would look for any image whose object reference
is shared across all pages (that is the signature of a reused
watermark overlay). Pseudocode:

```python
# Build a map: image name -> set of pages it appears on
# The watermark is the image whose set == all pages
# AND whose underlying smask contains text-like high-contrast pixels
```

This would make the tool work on more variants automatically. Worth
adding once you have a few sample files to compare.

## Limitations

- **Only works on this overlay-style watermark.** If a future version
  of EKP starts flattening the watermark into the scan itself
  (i.e. burning the pixels into the JPEG), this approach won't work and
  you'd need OCR + inpainting instead.
- **Does not handle password-protected PDFs.** If the PDF requires a
  user password (not just a permissions / "change" restriction), the
  app will tell you to decrypt it first with `qpdf --decrypt`.
- **50 MB upload limit.** Adjust `MAX_CONTENT_LENGTH` in `app.py` if you
  need bigger.

## A note on responsible use

The watermark exists for a reason — it lets RBI trace which copy was
shared by whom. Use this tool only on PDFs you downloaded yourself,
where the watermark is genuinely getting in the way of reading or
archiving. Don't use it to launder documents that were issued under
someone else's ID.
