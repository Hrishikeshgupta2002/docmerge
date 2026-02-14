import os
import uuid
from typing import List
import logging
from tempfile import NamedTemporaryFile

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from docx import Document
from docxcompose.composer import Composer

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DocMerge API",
    description="API for merging DOCX files using docxcompose",
    version="1.0.0"
)

# Validate file type
def validate_docx_file(file: UploadFile):
    """Validate that the uploaded file is a DOCX file"""
    # Check file extension
    if not file.filename.lower().endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only DOCX files are allowed")
    
    # Check content type
    if file.content_type != "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        logger.warning(f"File {file.filename} has unexpected content type: {file.content_type}")

@app.post("/merge-docx/",
          summary="Merge multiple DOCX files",
          description="Upload multiple DOCX files to merge them into a single document")
async def merge_docx(files: List[UploadFile] = File(..., description="List of DOCX files to merge")):
    """
    Merge multiple DOCX files into a single document

    Args:
        files: List of DOCX files to merge (minimum 2 files required)

    Returns:
        FileResponse: The merged DOCX file
    """
    # Validation: Check if at least 2 files are provided
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="At least 2 DOCX files are required for merging")

    # Validation: Check file count limit
    if len(files) > 10:  # Limit to 10 files at once
        raise HTTPException(status_code=400, detail="Maximum 10 files allowed at once")

    # Validate each file
    for file in files:
        validate_docx_file(file)

    temp_files = []
    output_file = None

    try:
        # Save uploaded files temporarily
        for file in files:
            temp_file = NamedTemporaryFile(delete=False, suffix=".docx")
            content = await file.read()
            temp_file.write(content)
            temp_file.close()
            temp_files.append(temp_file.name)
            logger.info(f"Saved temporary file: {temp_file.name}")

        # Merge documents
        logger.info(f"Merging {len(temp_files)} documents...")

        # Load the first document as the master
        master = Document(temp_files[0])
        composer = Composer(master)

        # Append remaining documents
        for file_path in temp_files[1:]:
            try:
                doc = Document(file_path)
                composer.append(doc)
                logger.info(f"Appended document: {file_path}")
            except Exception as e:
                logger.error(f"Error appending document {file_path}: {str(e)}")
                raise HTTPException(status_code=400, detail=f"Invalid DOCX file detected: {file_path}")

        # Create output file
        output_file = NamedTemporaryFile(delete=False, suffix=".docx")
        output_file.close()

        # Save the merged document
        composer.save(output_file.name)
        logger.info(f"Merged document saved to: {output_file.name}")

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
         description="Check if the API is running properly")
async def health_check():
    """Health check endpoint to verify API is running"""
    return {"status": "healthy", "message": "DocMerge API is running"}


@app.get("/info",
         summary="API Information",
         description="Get information about the API capabilities")
async def api_info():
    """Endpoint to get API information"""
    return {
        "title": "DocMerge API",
        "version": "1.0.0",
        "description": "API for merging DOCX files using docxcompose",
        "features": [
            "Merge multiple DOCX files into a single document",
            "Pure Python implementation - no external dependencies",
            "Handles large documents efficiently",
            "Secure temporary file handling"
        ],
        "usage": {
            "endpoint": "/merge-docx/",
            "method": "POST",
            "params": "Multiple DOCX files as form data",
            "requirements": "At least 2 DOCX files, max 10 files"
        }
    }