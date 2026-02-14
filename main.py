from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from docx import Document
from docxcompose.composer import Composer
import tempfile
import os
import uuid
from typing import List
import logging

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
    
    # Create temporary directory for processing
    temp_dir = "/tmp/docmerge"
    os.makedirs(temp_dir, exist_ok=True)
    
    temp_files = []
    output_path = None
    
    try:
        # Save uploaded files temporarily
        for file in files:
            temp_file_path = os.path.join(temp_dir, f"{uuid.uuid4()}.docx")
            with open(temp_file_path, "wb") as temp_file:
                content = await file.read()
                temp_file.write(content)
            temp_files.append(temp_file_path)
            logger.info(f"Saved temporary file: {temp_file_path}")
        
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
        
        # Generate unique output filename
        output_filename = f"merged_{uuid.uuid4()}.docx"
        output_path = os.path.join(temp_dir, output_filename)
        
        # Save the merged document
        composer.save(output_path)
        logger.info(f"Merged document saved to: {output_path}")
        
        # Return the merged file
        return FileResponse(
            path=output_path,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            filename="merged_document.docx",
            headers={"Content-Disposition": "attachment; filename=merged_document.docx"}
        )
        
    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        logger.error(f"Unexpected error during merging: {str(e)}")
        raise HTTPException(status_code=500, detail=f"An error occurred during document merging: {str(e)}")
    finally:
        # Clean up temporary files
        cleanup_temp_files(temp_files, output_path)


def cleanup_temp_files(input_files: List[str], output_file: str = None):
    """Clean up temporary files after processing"""
    for file_path in input_files:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.info(f"Removed temporary file: {file_path}")
        except Exception as e:
            logger.error(f"Error removing temporary file {file_path}: {str(e)}")
    
    # Remove output file if it exists
    if output_file and os.path.exists(output_file):
        try:
            os.remove(output_file)
            logger.info(f"Removed temporary output file: {output_file}")
        except Exception as e:
            logger.error(f"Error removing temporary output file {output_file}: {str(e)}")


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