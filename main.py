import os
from typing import List
import logging
from tempfile import NamedTemporaryFile

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
    A production-ready API for merging DOCX documents.
    
    ## Features
    
    * **Merge Documents**: Combine multiple DOCX files into a single document with enterprise-grade merging
    * **Preserves Formatting**: Headers, footers, section breaks, and complex layouts are maintained
    
    ## Technology Stack
    
    * FastAPI for high-performance async API
    * docxcompose.Composer for enterprise-grade document merging
    * Docker-based deployment for consistent environments
    * Railway-ready with optimized configuration
    
    ## Endpoints
    
    * `POST /merge-docx/` - Merge multiple DOCX files
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
            "description": "Document merging operations.",
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
        for idx, docx_path in enumerate(docx_files[1:], start=2):
            try:
                doc_to_append = Document(docx_path)
                composer.append(doc_to_append)
                logger.debug(f"Merged DOCX: {os.path.basename(docx_path)}")
            except Exception as e:
                error_msg = str(e)
                logger.error(f"Error merging DOCX file {docx_path}: {error_msg}")
                
                # Handle specific style relationship conflicts
                if "multiple relationships" in error_msg and "styles" in error_msg.lower():
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            f"Style conflict detected when merging document #{idx} ({os.path.basename(docx_path)}). "
                            "This occurs when documents have conflicting style definitions. "
                            "Try merging documents that were created with the same template or normalize styles before merging."
                        )
                    )
                elif "relationship" in error_msg.lower():
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            f"Document structure conflict when merging document #{idx} ({os.path.basename(docx_path)}). "
                            "The documents may have incompatible internal structures. "
                            "Please ensure all documents are valid DOCX files created with compatible versions of Word."
                        )
                    )
                else:
                    raise HTTPException(
                        status_code=400,
                        detail=f"Failed to merge document #{idx} ({os.path.basename(docx_path)}): {error_msg[:200]}"
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
        # Re-raise HTTP exceptions (these already have user-friendly messages)
        raise
    except Exception as e:
        error_msg = str(e)
        logger.error(f"Unexpected error during merging: {error_msg}")

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

        # Provide more context for common errors
        if "relationship" in error_msg.lower():
            raise HTTPException(
                status_code=500,
                detail=(
                    "Document structure error during merging. "
                    "This may occur when documents have incompatible internal structures or style definitions. "
                    "Please ensure all documents are valid DOCX files."
                )
            )
        else:
            raise HTTPException(
                status_code=500,
                detail=f"An unexpected error occurred during document merging: {error_msg[:300]}"
            )


@app.get("/",
         summary="API Health Check",
         description="Check if the API is running properly",
         tags=["Health"])
async def health_check():
    """Health check endpoint to verify API is running"""
    return {"status": "healthy", "message": "DocMerge API is running"}


@app.get("/info",
         summary="API Information",
         description="Get information about the API capabilities",
         tags=["Health"])
async def api_info():
    """Endpoint to get API information"""
    return {
        "title": "DocMerge API",
        "version": "1.0.0",
        "description": "API for merging DOCX files",
        "features": [
            "Merge multiple DOCX files into a single document",
            "Enterprise-grade DOCX merging using docxcompose.Composer (preserves headers, footers, section breaks)",
            "Handles large documents efficiently",
            "Secure temporary file handling with automatic cleanup",
            "Docker-based deployment with consistent environment"
        ],
        "endpoints": {
            "merge": {
                "path": "/merge-docx/",
                "method": "POST",
                "description": "Merge multiple DOCX files",
                "params": "Multiple DOCX files as form data",
                "requirements": "At least 2 DOCX files, max 40 files"
            }
        }
    }