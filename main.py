"""
DocMerge API - Convert DOCX to PDF and merge into a single PDF.

Single endpoint: POST /merge-pdf/
"""
import os
import shutil
import subprocess
import time
from tempfile import NamedTemporaryFile, mkdtemp
from typing import List
import logging

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pypdf import PdfMerger

_log_level = getattr(logging, os.getenv("LOG_LEVEL", "INFO").upper(), logging.INFO)
logging.basicConfig(level=_log_level, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
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


# LibreOffice headless env (no X11 display, gen plugin for software rendering)
def _libreoffice_env(profile_dir: str) -> dict:
    return {
        **os.environ,
        "SAL_USE_VCLPLUGIN": "gen",
        "HOME": profile_dir,
        "XDG_RUNTIME_DIR": profile_dir,
        "TMPDIR": profile_dir,
    }


def _resolve_soffice() -> str:
    """Resolve soffice/libreoffice/lowriter executable. Prefer soffice, fallback to lowriter for DOCX."""
    for cmd in ("soffice", "libreoffice", "lowriter"):
        try:
            path = shutil.which(cmd)
            if path:
                return path
        except Exception:
            pass
    return "soffice"  # fallback; will raise FileNotFoundError if missing


def _resolve_xvfb_run() -> str | None:
    """Resolve xvfb-run for virtual display (fixes 'Can't open display' in containers)."""
    return shutil.which("xvfb-run")


_SOFFICE_CMD: str | None = None
_XVFB_RUN: str | None = None


def convert_docx_to_pdf(docx_path: str, output_dir: str, profile_dir: str | None = None) -> str:
    """
    Convert DOCX to PDF using LibreOffice headless.
    Uses a unique profile dir per conversion to avoid lock conflicts in batch processing.
    """
    global _SOFFICE_CMD, _XVFB_RUN
    if _SOFFICE_CMD is None:
        _SOFFICE_CMD = _resolve_soffice()
        _XVFB_RUN = _resolve_xvfb_run()
        logger.info(f"Using LibreOffice: {_SOFFICE_CMD}, xvfb: {_XVFB_RUN or 'no'}")

    base_name = os.path.basename(docx_path).replace(".docx", ".pdf")
    expected_path = os.path.join(output_dir, base_name)
    docx_path_abs = os.path.abspath(docx_path)
    docx_dir = os.path.dirname(docx_path_abs)

    # Unique profile per conversion: avoids profile locks when processing many files
    prof = profile_dir or mkdtemp(prefix="lo_profile_")
    env = _libreoffice_env(prof)
    logger.debug(f"Converting {docx_path_abs} -> {expected_path}, profile={prof}")

    # LibreOffice often ignores --outdir when it differs from source dir; use docx_dir for reliable output
    outdir_for_lo = docx_dir

    soffice_args = [
        _SOFFICE_CMD,
        "--headless",
        "--norestore",
        f"-env:UserInstallation=file://{prof}",
        "--convert-to",
        "pdf",
        "--outdir",
        outdir_for_lo,
        docx_path_abs,
    ]
    cmd = ([_XVFB_RUN, "-a"] + soffice_args) if _XVFB_RUN else soffice_args

    try:
        t0 = time.perf_counter()
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            env=env,
            cwd=docx_dir,
        )
        elapsed = time.perf_counter() - t0
        logger.info(f"LibreOffice conversion {base_name}: returncode={result.returncode}, elapsed={elapsed:.2f}s")

        # Allow filesystem sync (LibreOffice may flush after process exits)
        if result.returncode == 0:
            time.sleep(0.3)

        if result.returncode != 0:
            stderr_full = (result.stderr or "").strip()
            stdout_full = (result.stdout or "").strip()
            logger.warning(
                f"LibreOffice exit {result.returncode} for {base_name}. stderr: {stderr_full} | stdout: {stdout_full}"
            )

        candidates = [expected_path, os.path.join(docx_dir, base_name), os.path.join(prof, base_name)]
        for candidate in candidates:
            if os.path.exists(candidate):
                logger.debug(f"PDF found at {candidate}")
                return candidate

        # LibreOffice sometimes writes tmpXXX.pdf and overwrites same file; copy to unique path
        if result.returncode == 0:
            search_dirs = (output_dir, docx_dir, prof)
            newest_pdf: str | None = None
            newest_mtime = 0.0
            for d in search_dirs:
                if not os.path.isdir(d):
                    continue
                for f in os.listdir(d):
                    if f.lower().endswith(".pdf"):
                        p = os.path.join(d, f)
                        mtime = os.path.getmtime(p)
                        if mtime > newest_mtime:
                            newest_mtime = mtime
                            newest_pdf = p
            if newest_pdf:
                # Copy to output_dir with expected name so each conversion has unique path
                unique_dest = os.path.join(output_dir, base_name)
                if newest_pdf != unique_dest:
                    shutil.copy2(newest_pdf, unique_dest)
                    logger.info(f"Copied {os.path.basename(newest_pdf)} -> {base_name}")
                return unique_dest

        # Log stdout/stderr even on success - may explain missing output
        if result.returncode == 0:
            lo_stdout = (result.stdout or "").strip()
            lo_stderr = (result.stderr or "").strip()
            if lo_stdout or lo_stderr:
                logger.warning(f"LibreOffice stdout: {lo_stdout[:500]} | stderr: {lo_stderr[:500]}")
        logger.error(f"PDF not found. Checked: {candidates}")
        for d in (output_dir, docx_dir, prof):
            if os.path.isdir(d):
                contents = os.listdir(d)
                logger.error(f"Contents of {d}: {contents}")
                pdfs = [f for f in contents if f.lower().endswith(".pdf")]
                if pdfs:
                    logger.warning(f"LibreOffice created {pdfs} in {d}, expected {base_name}")
        raise Exception(f"PDF was not generated for {os.path.basename(docx_path)}")
    except subprocess.TimeoutExpired:
        raise Exception(f"Conversion timed out: {os.path.basename(docx_path)}")
    finally:
        if profile_dir is None and os.path.isdir(prof):
            shutil.rmtree(prof, ignore_errors=True)


def merge_pdfs(pdf_paths: List[str], output_path: str):
    """Merge PDF files using pypdf (streaming, low memory)."""
    logger.info(f"Merging {len(pdf_paths)} PDFs -> {output_path}")
    merger = PdfMerger()
    try:
        for p in pdf_paths:
            if not os.path.exists(p):
                raise Exception(f"PDF not found: {os.path.basename(p)}")
            merger.append(p)
        with open(output_path, "wb") as f:
            merger.write(f)
        size = os.path.getsize(output_path)
        logger.info(f"Merged PDF size: {size} bytes")
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
    output_dir = None
    temp_files = []
    output_file = None

    try:
        temp_dir = mkdtemp()
        total_bytes = 0

        for file in files:
            content = await file.read()
            total_bytes += len(content)
            if (file.filename or "").lower().split(".")[-1] != "docx":
                raise HTTPException(status_code=400, detail="Only DOCX supported")
            path = os.path.join(temp_dir, f"doc_{len(temp_files)}.docx")
            with open(path, "wb") as f:
                f.write(content)
            temp_files.append(path)

        logger.info(f"Received {len(files)} DOCX files, {total_bytes} bytes total. temp_dir={temp_dir}")

        # Output in separate dir to avoid collision with LibreOffice temp PDFs in docx dir
        output_dir = mkdtemp(prefix="docmerge_out_")
        output_file = NamedTemporaryFile(delete=False, suffix=".pdf", dir=output_dir)
        output_file.close()
        temp_files.append(output_file.name)

        from merge_as_pdf import merge_docx_to_pdf

        t0 = time.perf_counter()
        logger.info(f"Converting {len(temp_files)-1} DOCX to PDF and merging...")
        merge_docx_to_pdf(temp_files[:-1], output_file.name)
        elapsed = time.perf_counter() - t0
        logger.info(f"Merge complete in {elapsed:.2f}s")

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
        if output_dir and os.path.isdir(output_dir):
            shutil.rmtree(output_dir, ignore_errors=True)

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
        logger.error(f"PDF merge error: {error_msg}", exc_info=True)
        for p in temp_files:
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except OSError:
                pass
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)
        if output_dir and os.path.isdir(output_dir):
            shutil.rmtree(output_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"PDF merge failed: {error_msg[:200]}")


@app.on_event("startup")
async def startup_validation():
    """Verify LibreOffice is available before accepting requests."""
    try:
        soffice = _resolve_soffice()
        xvfb = _resolve_xvfb_run()
        logger.info(f"Resolved LibreOffice: {soffice}, xvfb: {xvfb or 'no'}")
        cmd = ([xvfb, "-a", soffice] if xvfb else [soffice]) + ["--headless", "--version"]
        r = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=10,
            env={**os.environ, "SAL_USE_VCLPLUGIN": "gen", "HOME": "/tmp", "XDG_RUNTIME_DIR": "/tmp"},
        )
        if r.returncode == 0:
            version_out = (r.stdout or r.stderr or "").strip()
            logger.info(f"LibreOffice OK: {soffice} | {version_out}")
        else:
            logger.warning(
                f"LibreOffice version check returned {r.returncode}. stdout: {r.stdout}. stderr: {r.stderr}"
            )
    except FileNotFoundError:
        logger.error("LibreOffice/soffice not found in PATH")
    except Exception as e:
        logger.warning(f"LibreOffice startup check: {e}", exc_info=True)


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
