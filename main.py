"""
DocMerge API - Convert DOCX to PDF and merge into a single PDF.

Single endpoint: POST /merge-pdf/
"""
import os
import shutil
import subprocess
from tempfile import NamedTemporaryFile, mkdtemp
from typing import List
import logging

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pypdf import PdfMerger

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DocMerge API",
    description="""
    Convert DOCX files to PDF and merge into a single PDF document.

    * `POST /merge-pdf/` - Upload DOCX files, get merged PDF (LibreOffice + pypdf)
    """,
    version="1.0.0",
    servers=[
        {"url": "https://docmerge-production.up.railway.app", "description": "Production"},
        {"url": "http://localhost:8080", "description": "Local"},
    ],
)

_cors = os.getenv("CORS_ORIGINS", "*")
cors_origins = ["*"] if _cors == "*" else [o.strip() for o in _cors.split(",")]
app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


def validate_docx_file(file: UploadFile):
    """Validate that the uploaded file is a DOCX file."""
    filename_lower = (file.filename or "").lower()
    if not filename_lower.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only DOCX files are allowed")


# LibreOffice headless env (Railway/Docker: no display, gen plugin, writable runtime)
_LIBREOFFICE_ENV = {
    **os.environ,
    "SAL_USE_VCLPLUGIN": "gen",
    "HOME": os.environ.get("HOME", "/tmp"),
    "XDG_RUNTIME_DIR": os.environ.get("XDG_RUNTIME_DIR", "/tmp"),
}


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> str:
    """
    Convert DOCX to PDF using LibreOffice headless.
    Uses docx directory as --outdir for reliability (some LibreOffice builds
    ignore --outdir when it differs from the source directory).
    """
    base_name = os.path.basename(docx_path).replace(".docx", ".pdf")
    # Use output_dir; fallback to docx dir if LibreOffice wrote next to source
    expected_path = os.path.join(output_dir, base_name)
    docx_dir = os.path.dirname(os.path.abspath(docx_path))

    try:
        result = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir],
            capture_output=True,
            text=True,
            timeout=120,
            env=_LIBREOFFICE_ENV,
        )
        if result.returncode != 0:
            stderr_preview = (result.stderr or "").strip()[:500]
            logger.warning(f"LibreOffice exit {result.returncode} for {base_name}: {stderr_preview}")

        # Check expected location first, then fallback to docx directory
        for candidate in (expected_path, os.path.join(docx_dir, base_name)):
            if os.path.exists(candidate):
                return candidate

        # Diagnostic: log what LibreOffice produced
        for d in (output_dir, docx_dir):
            if os.path.isdir(d):
                contents = os.listdir(d)
                pdfs = [f for f in contents if f.lower().endswith(".pdf")]
                if pdfs:
                    logger.warning(f"LibreOffice created {pdfs} in {d}, expected {base_name}")
                break
        raise Exception(f"PDF was not generated for {os.path.basename(docx_path)}")
    except subprocess.TimeoutExpired:
        raise Exception(f"Conversion timed out: {os.path.basename(docx_path)}")


def merge_pdfs(pdf_paths: List[str], output_path: str):
    """Merge PDF files using pypdf (streaming, low memory)."""
    merger = PdfMerger()
    try:
        for p in pdf_paths:
            if not os.path.exists(p):
                raise Exception(f"PDF not found: {os.path.basename(p)}")
            merger.append(p)
        with open(output_path, "wb") as f:
            merger.write(f)
    finally:
        merger.close()


@app.post("/merge-pdf/",
          summary="Convert DOCX to PDF and merge",
          description="Upload DOCX files. Each is converted to PDF via LibreOffice, then merged into one PDF.")
async def merge_files_as_pdf(files: List[UploadFile] = File(..., description="DOCX files")):
    """Convert DOCX files to PDF and merge into a single PDF."""
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="At least 2 files required")
    if len(files) > 40:
        raise HTTPException(status_code=400, detail="Maximum 40 files allowed")

    for f in files:
        validate_docx_file(f)

    temp_dir = None
    temp_files = []
    output_file = None

    try:
        temp_dir = mkdtemp()

        for file in files:
            content = await file.read()
            if (file.filename or "").lower().split(".")[-1] != "docx":
                raise HTTPException(status_code=400, detail="Only DOCX supported")
            path = os.path.join(temp_dir, f"doc_{len(temp_files)}.docx")
            with open(path, "wb") as f:
                f.write(content)
            temp_files.append(path)

        output_file = NamedTemporaryFile(delete=False, suffix=".pdf", dir=temp_dir)
        output_file.close()
        temp_files.append(output_file.name)

        from merge_as_pdf import merge_docx_to_pdf
        logger.info(f"Converting {len(temp_files)-1} DOCX to PDF and merging...")
        merge_docx_to_pdf(temp_files[:-1], output_file.name)
        logger.info("Merge complete")

        with open(output_file.name, "rb") as f:
            content = f.read()

        for p in temp_files:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except OSError:
                pass
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)

        def iterfile():
            yield content

        return StreamingResponse(
            iterfile(),
            media_type="application/pdf",
            headers={"Content-Disposition": "attachment; filename=merged_document.pdf"},
        )

    except HTTPException:
        raise
    except Exception as e:
        error_msg = str(e)
        logger.error(f"PDF merge error: {error_msg}")
        for p in temp_files:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except OSError:
                pass
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"PDF merge failed: {error_msg[:200]}")


@app.get("/")
async def health_check():
    """Health check for Railway / load balancers."""
    return {"status": "healthy"}


@app.get("/info")
async def api_info():
    """API information."""
    return {
        "title": "DocMerge API",
        "version": "1.0.0",
        "endpoint": "POST /merge-pdf/ - Upload DOCX files, get merged PDF",
        "requirements": "LibreOffice (included in Docker image)",
    }
