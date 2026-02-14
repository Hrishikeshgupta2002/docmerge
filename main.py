import os
import subprocess
import shutil
from typing import List
import logging
from tempfile import NamedTemporaryFile, mkdtemp
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from docx import Document
from docxcompose.composer import Composer

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DocMerge API",
    description="""
    A production-ready API for document processing operations.
    
    ## Features
    
    * **Merge Documents**: Combine multiple DOCX files into a single document with enterprise-grade merging
    * **Convert to PDF**: Convert DOCX files to PDF format using LibreOffice (100% free)
    
    ## Technology Stack
    
    * FastAPI for high-performance async API
    * LibreOffice for document conversion (open-source, no licensing fees)
    * Docker-based deployment for consistent environments
    * Railway-ready with optimized configuration
    
    ## Endpoints
    
    * `POST /merge-docx/` - Merge multiple DOCX files
    * `POST /docx-to-pdf/` - Convert DOCX file to PDF
    * `GET /` - Health check endpoint
    * `GET /info` - API information and capabilities
    """,
    version="1.0.0",
    contact={
        "name": "DocMerge API",
        "url": "https://docmerge-production.up.railway.app",
    },
    license_info={
        "name": "MIT",
    },
    servers=[
        {
            "url": "https://docmerge-production.up.railway.app",
            "description": "Production server"
        },
        {
            "url": "http://localhost:8080",
            "description": "Local development server"
        }
    ],
    tags_metadata=[
        {
            "name": "Documents",
            "description": "Document processing operations including merging and conversion.",
        },
        {
            "name": "Health",
            "description": "Health check and API information endpoints.",
        },
    ],
)

# Configure CORS for production deployment
# Default allowed origins include Railway domain
default_origins = [
    "https://docmerge-production.up.railway.app",
    "http://docmerge-production.up.railway.app",
]

# Allow custom CORS origins via environment variable (comma-separated)
# If CORS_ORIGINS is set to "*", allow all origins
cors_origins_env = os.getenv("CORS_ORIGINS", "")
if cors_origins_env == "*":
    cors_origins = ["*"]
elif cors_origins_env:
    cors_origins = [origin.strip() for origin in cors_origins_env.split(",")]
else:
    cors_origins = default_origins

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["*"],
)

# Validate file type
def validate_docx_file(file: UploadFile):
    """Validate that the uploaded file is a DOCX file"""
    filename_lower = file.filename.lower() if file.filename else ""

    # Check file extension
    if not filename_lower.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only DOCX files are allowed for merging")

    # Check content type
    allowed_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if file.content_type and file.content_type != allowed_type:
        logger.warning(f"File {file.filename} has unexpected content type: {file.content_type}")


def get_libreoffice_command():
    """
    Get the LibreOffice command path based on the system
    Returns the command to run LibreOffice in headless mode
    """
    # Check common LibreOffice paths
    possible_paths = [
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "libreoffice",  # Fallback to PATH
        "soffice",  # Alternative command name
    ]
    
    for path in possible_paths:
        if shutil.which(path):
            return path
    
    # If not found, return default (will fail with clear error)
    return "libreoffice"


def convert_docx_to_pdf(docx_path: str, output_dir: str, timeout: int = None) -> str:
    """
    Convert DOCX file to PDF using LibreOffice in headless mode
    
    Args:
        docx_path: Path to the input DOCX file
        output_dir: Directory where the PDF should be saved
        timeout: Conversion timeout in seconds (default: 60, complex docs may need 120+)
        
    Returns:
        Path to the generated PDF file
        
    Raises:
        HTTPException: If conversion fails
    """
    try:
        libreoffice_cmd = get_libreoffice_command()
        
        # Use configurable timeout (default 60s, can be increased for complex documents)
        conversion_timeout = timeout or int(os.getenv("LIBREOFFICE_TIMEOUT", "60"))
        
        # LibreOffice command: --headless --convert-to pdf --outdir <dir> <input>
        # Critical: Always specify --outdir to control output location
        # Additional flags for better performance and reliability:
        # --nodefault: Don't load default document templates
        # --nolockcheck: Skip file locking checks (important for temp files)
        # --nologo: Don't show splash screen
        # --norestore: Don't restore previous session
        # --invisible: Run invisibly (no UI)
        # --safe-mode: Run in safe mode (prevents macros, etc.)
        command = [
            libreoffice_cmd,
            "--headless",
            "--nodefault",
            "--nolockcheck",
            "--nologo",
            "--norestore",
            "--invisible",
            "--safe-mode",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            docx_path
        ]
        
        logger.info(f"Converting DOCX to PDF: {os.path.basename(docx_path)} (timeout: {conversion_timeout}s)")
        logger.debug(f"LibreOffice command: {' '.join(command)}")
        
        # Run LibreOffice conversion
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=conversion_timeout,
            check=False
        )
        
        if result.returncode != 0:
            error_msg = result.stderr or result.stdout or "Unknown error"
            logger.error(f"LibreOffice conversion failed: {error_msg}")
            raise HTTPException(
                status_code=500,
                detail=f"PDF conversion failed: {error_msg[:200]}"
            )
        
        # LibreOffice outputs PDF with same name but .pdf extension
        input_filename = Path(docx_path).stem
        expected_pdf_path = os.path.join(output_dir, f"{input_filename}.pdf")
        
        # Check if PDF was actually created
        if not os.path.exists(expected_pdf_path):
            logger.error(f"PDF file not found at expected path: {expected_pdf_path}")
            raise HTTPException(
                status_code=500,
                detail="PDF conversion completed but output file not found"
            )
        
        logger.info(f"Successfully converted DOCX to PDF: {expected_pdf_path}")
        return expected_pdf_path
        
    except subprocess.TimeoutExpired:
        logger.error("LibreOffice conversion timed out")
        raise HTTPException(
            status_code=500,
            detail="PDF conversion timed out. The document may be too large or complex."
        )
    except FileNotFoundError:
        logger.error("LibreOffice not found in system")
        raise HTTPException(
            status_code=500,
            detail="LibreOffice is not installed or not in PATH"
        )
    except Exception as e:
        logger.error(f"Unexpected error during DOCX to PDF conversion: {str(e)}")
        raise HTTPException(
            status_code=500,
            detail=f"PDF conversion failed: {str(e)}"
        )




@app.post("/merge-docx/",
          summary="Merge multiple DOCX files",
          description="Upload multiple DOCX files to merge them into a single document using enterprise-grade merging",
          tags=["Documents"])
async def merge_files(files: List[UploadFile] = File(..., description="List of DOCX files to merge")):
    """
    Merge multiple DOCX files into a single document using docxcompose.Composer
    
    This preserves all formatting including headers, footers, section breaks, and complex layouts.

    Args:
        files: List of DOCX files to merge (minimum 2 files required)

    Returns:
        StreamingResponse: The merged DOCX file
    """
    # Validation: Check if at least 2 files are provided
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="At least 2 files are required for merging")

    # Validation: Check file count limit
    if len(files) > 40:  # Limit to 40 files at once
        raise HTTPException(status_code=400, detail="Maximum 40 files allowed at once")

    # Validate each file - only DOCX allowed
    for file in files:
        validate_docx_file(file)

    docx_files = []
    temp_files = []  # Track all temporary files for cleanup
    output_file = None

    try:
        # Save all DOCX files temporarily for Composer
        for file in files:
            content = await file.read()
            file_extension = file.filename.lower().split('.')[-1]

            if file_extension != 'docx':
                raise HTTPException(status_code=400, detail=f"Only DOCX files are supported. Received: {file_extension}")

            # Save DOCX file temporarily for Composer
            temp_file = NamedTemporaryFile(delete=False, suffix=".docx")
            temp_file.write(content)
            temp_file.close()
            temp_files.append(temp_file.name)
            docx_files.append(temp_file.name)
            logger.debug(f"Added DOCX file: {file.filename}")

        # Merge DOCX files using Composer (preserves formatting, headers, footers, etc.)
        # Use the first DOCX as the master document
        master_doc = Document(docx_files[0])
        composer = Composer(master_doc)

        # Append remaining DOCX files
        for docx_path in docx_files[1:]:
            try:
                doc_to_append = Document(docx_path)
                composer.append(doc_to_append)
                logger.debug(f"Merged DOCX: {os.path.basename(docx_path)}")
            except Exception as e:
                logger.error(f"Error merging DOCX file {docx_path}: {str(e)}")
                raise HTTPException(
                    status_code=400,
                    detail=f"Failed to merge DOCX file: {os.path.basename(docx_path)}"
                )

        # Create output file
        output_file = NamedTemporaryFile(delete=False, suffix=".docx")
        output_file.close()

        # Save the merged document
        master_doc.save(output_file.name)
        logger.info(f"Merged document saved to: {output_file.name} (merged {len(docx_files)} DOCX files)")

        # Read the merged file content before cleanup
        with open(output_file.name, 'rb') as f:
            content = f.read()

        # Clean up temporary files early to free resources
        for temp_path in temp_files:
            try:
                os.unlink(temp_path)
            except OSError:
                pass  # Ignore errors when deleting temp files

        if output_file and os.path.exists(output_file.name):
            try:
                os.unlink(output_file.name)
            except OSError:
                pass  # Ignore errors when deleting temp files

        # Return the merged file content as a streaming response
        def iterfile():
            yield content

        return StreamingResponse(
            iterfile(),
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers={
                "Content-Disposition": "attachment; filename=merged_document.docx"
            }
        )

    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(f"Unexpected error during merging: {str(e)}")

        # Ensure cleanup even if there's an exception
        for temp_path in temp_files:
            try:
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
            except OSError:
                pass  # Ignore errors when deleting temp files

        if output_file and os.path.exists(output_file.name):
            try:
                os.unlink(output_file.name)
            except OSError:
                pass  # Ignore errors when deleting temp files

        raise HTTPException(status_code=500, detail=f"An error occurred during document merging: {str(e)}")


@app.get("/",
         summary="API Health Check",
         description="Check if the API is running properly",
         tags=["Health"])
async def health_check():
    """Health check endpoint to verify API is running"""
    return {"status": "healthy", "message": "DocMerge API is running"}


@app.post("/docx-to-pdf/",
          summary="Convert DOCX to PDF",
          description="Convert a single DOCX file to PDF using LibreOffice",
          tags=["Documents"])
async def docx_to_pdf(file: UploadFile = File(..., description="DOCX file to convert to PDF")):
    """
    Convert a DOCX file to PDF format using LibreOffice
    
    Args:
        file: DOCX file to convert
        
    Returns:
        StreamingResponse: The converted PDF file
    """
    # Validate file type
    filename_lower = file.filename.lower() if file.filename else ""
    if not filename_lower.endswith('.docx'):
        raise HTTPException(
            status_code=400,
            detail="Only DOCX files are supported for PDF conversion"
        )
    
    # Check content type
    if file.content_type and file.content_type != "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        logger.warning(f"File {file.filename} has unexpected content type: {file.content_type}")
    
    temp_docx = None
    temp_output_dir = None
    
    try:
        # Read file content
        content = await file.read()
        
        # Create temporary directory for output (LibreOffice needs a directory)
        temp_output_dir = mkdtemp(prefix="docx_to_pdf_")
        
        # Create temporary DOCX file
        temp_docx = NamedTemporaryFile(delete=False, suffix=".docx", dir=temp_output_dir)
        temp_docx.write(content)
        temp_docx.close()
        
        logger.info(f"Processing DOCX file: {file.filename}")
        
        # Convert DOCX to PDF
        pdf_path = convert_docx_to_pdf(temp_docx.name, temp_output_dir)
        
        # Read the generated PDF
        with open(pdf_path, 'rb') as f:
            pdf_content = f.read()
        
        # Generate output filename
        input_filename = Path(file.filename or "document").stem
        output_filename = f"{input_filename}.pdf"
        
        # Clean up temporary files
        try:
            if temp_docx and os.path.exists(temp_docx.name):
                os.unlink(temp_docx.name)
            if temp_output_dir and os.path.exists(temp_output_dir):
                # Use shutil.rmtree for reliable directory removal
                shutil.rmtree(temp_output_dir, ignore_errors=True)
        except Exception as cleanup_error:
            logger.warning(f"Error during cleanup: {cleanup_error}")
        
        # Return PDF as streaming response
        def iterfile():
            yield pdf_content
        
        return StreamingResponse(
            iterfile(),
            media_type='application/pdf',
            headers={
                "Content-Disposition": f'attachment; filename="{output_filename}"'
            }
        )
        
    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(f"Unexpected error during DOCX to PDF conversion: {str(e)}")
        
        # Cleanup on error
        try:
            if temp_docx and os.path.exists(temp_docx.name):
                os.unlink(temp_docx.name)
            if temp_output_dir and os.path.exists(temp_output_dir):
                shutil.rmtree(temp_output_dir, ignore_errors=True)
        except Exception:
            pass
        
        raise HTTPException(
            status_code=500,
            detail=f"An error occurred during PDF conversion: {str(e)}"
        )


@app.get("/info",
         summary="API Information",
         description="Get information about the API capabilities",
         tags=["Health"])
async def api_info():
    """Endpoint to get API information"""
    return {
        "title": "DocMerge API",
        "version": "1.0.0",
        "description": "API for merging DOCX files and converting DOCX to PDF",
        "features": [
            "Merge multiple DOCX files into a single document",
            "Enterprise-grade DOCX merging using docxcompose.Composer (preserves headers, footers, section breaks)",
            "Convert DOCX files to PDF using LibreOffice headless (100% free, production-ready)",
            "Handles large documents efficiently with configurable timeouts",
            "Secure temporary file handling with automatic cleanup",
            "Docker-based deployment with consistent environment",
            "Optimized LibreOffice conversion with safe-mode and performance flags"
        ],
        "endpoints": {
            "merge": {
                "path": "/merge-docx/",
                "method": "POST",
                "description": "Merge multiple DOCX files",
                "params": "Multiple DOCX files as form data",
                "requirements": "At least 2 DOCX files, max 40 files"
            },
            "convert": {
                "path": "/docx-to-pdf/",
                "method": "POST",
                "description": "Convert DOCX file to PDF",
                "params": "Single DOCX file as form data",
                "requirements": "One valid DOCX file"
            }
        }
    }