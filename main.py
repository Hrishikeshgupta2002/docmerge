import os
import uuid
from typing import List
import logging
from tempfile import NamedTemporaryFile
from io import BytesIO

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from docx import Document
from docx.shared import Inches
from docxcompose.composer import Composer

# For PDF processing
from pdf2image import convert_from_bytes
from PIL import Image

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DocMerge API",
    description="API for merging DOCX and PDF files",
    version="1.0.0"
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
def validate_supported_file(file: UploadFile):
    """Validate that the uploaded file is either a DOCX or PDF file"""
    filename_lower = file.filename.lower()

    # Check file extension
    if not (filename_lower.endswith('.docx') or filename_lower.endswith('.pdf')):
        raise HTTPException(status_code=400, detail="Only DOCX and PDF files are allowed")

    # Check content type
    allowed_types = [
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",  # DOCX
        "application/pdf"  # PDF
    ]
    if file.content_type not in allowed_types:
        logger.warning(f"File {file.filename} has unexpected content type: {file.content_type}")

def convert_pdf_to_images(pdf_bytes: bytes) -> List[Image.Image]:
    """
    Convert PDF bytes to a list of PIL Image objects
    """
    try:
        # Convert PDF to list of images
        images = convert_from_bytes(pdf_bytes)
        return images
    except Exception as e:
        logger.error(f"Error converting PDF to images: {str(e)}")
        raise HTTPException(status_code=400, detail=f"Could not process PDF file: {str(e)}")


def add_pdf_as_images_to_doc(doc: Document, pdf_bytes: bytes):
    """
    Convert PDF pages to images and add them to the word document
    """
    try:
        # Convert PDF to images
        images = convert_pdf_to_images(pdf_bytes)

        # Add each image to the document
        for i, image in enumerate(images):
            # Add a paragraph to hold the image
            paragraph = doc.add_paragraph()

            # Add the image to the paragraph
            # Resize image to fit page width while preserving aspect ratio
            img_width = Inches(6.5)  # Standard page width with margins

            # Calculate proportional height
            aspect_ratio = image.height / image.width
            img_height = img_width * aspect_ratio

            # Create a BytesIO object to store the image in memory
            img_io = BytesIO()
            image.save(img_io, 'JPEG')
            img_io.seek(0)

            run = paragraph.add_run()
            run.add_picture(img_io, width=img_width, height=img_height)

            # Add page break between pages (except for the last page)
            if i < len(images) - 1:
                doc.add_page_break()

    except Exception as e:
        logger.error(f"Error adding PDF images to document: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Could not add PDF content to document: {str(e)}")


@app.post("/merge-docx/",
          summary="Merge multiple DOCX or PDF files",
          description="Upload multiple DOCX or PDF files to merge them into a single document")
async def merge_files(files: List[UploadFile] = File(..., description="List of DOCX or PDF files to merge")):
    """
    Merge multiple DOCX or PDF files into a single document

    Args:
        files: List of DOCX or PDF files to merge (minimum 2 files required)

    Returns:
        StreamingResponse: The merged DOCX file
    """
    # Validation: Check if at least 2 files are provided
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="At least 2 files are required for merging")

    # Validation: Check file count limit
    if len(files) > 40:  # Limit to 40 files at once
        raise HTTPException(status_code=400, detail="Maximum 40 files allowed at once")

    # Validate each file
    for file in files:
        validate_supported_file(file)

    temp_files = []  # Track DOCX files only
    output_file = None

    try:
        # Create the initial document
        doc = Document()

        # Process each file
        for file in files:
            content = await file.read()
            file_extension = file.filename.lower().split('.')[-1]

            if file_extension == 'docx':
                # Handle DOCX files by adding their content to the current document
                try:
                    # Create a temporary file for the uploaded docx
                    temp_file = NamedTemporaryFile(delete=False, suffix=".docx")
                    temp_file.write(content)
                    temp_file.close()
                    temp_files.append(temp_file.name)

                    # Load the docx and append its paragraphs to our document
                    temp_doc = Document(temp_file.name)
                    for element in temp_doc.element.body:
                        doc.element.body.append(element)

                except Exception as e:
                    logger.error(f"Error processing DOCX file {file.filename}: {str(e)}")
                    raise HTTPException(status_code=400, detail=f"Invalid DOCX file: {file.filename}")

            elif file_extension == 'pdf':
                # Handle PDF files by converting to images and adding to document
                try:
                    add_pdf_as_images_to_doc(doc, content)
                except Exception as e:
                    logger.error(f"Error processing PDF file {file.filename}: {str(e)}")
                    raise HTTPException(status_code=400, detail=f"Invalid PDF file: {file.filename}")
            else:
                raise HTTPException(status_code=400, detail=f"Unsupported file type: {file_extension}")

        # Create output file
        output_file = NamedTemporaryFile(delete=False, suffix=".docx")
        output_file.close()

        # Save the merged document
        doc.save(output_file.name)
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
        "description": "API for merging DOCX and PDF files",
        "features": [
            "Merge multiple DOCX and PDF files into a single document",
            "Convert PDF pages to images and embed them in Word documents",
            "Pure Python implementation - no external dependencies",
            "Handles large documents efficiently",
            "Secure temporary file handling"
        ],
        "usage": {
            "endpoint": "/merge-docx/",
            "method": "POST",
            "params": "Multiple DOCX or PDF files as form data",
            "requirements": "At least 2 files, max 40 files (supports both DOCX and PDF)"
        }
    }