# DocMerge API

A FastAPI-based server for merging multiple DOCX files into a single document using `docxcompose`.

## Features

- Pure Python implementation (no external dependencies like MS Word or LibreOffice)
- Secure temporary file handling
- Input validation for DOCX files
- Efficient merging of multiple documents
- Memory-conscious processing

## Setup

1. Make sure you have Python 3.10 installed
2. Clone or create this project directory
3. Create and activate the virtual environment:

```bash
python3.10 -m venv docmerge_env
source docmerge_env/bin/activate
pip install -r requirements.txt
```

## Running the Server

### Option 1: Using the provided script
```bash
./run_server.sh
```

### Option 2: Manual startup
```bash
source docmerge_env/bin/activate
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

The server will be available at `http://localhost:8000`

## API Endpoints

- `GET /` - Health check
- `GET /info` - API information
- `POST /merge-docx/` - Merge multiple DOCX files

## Using the Merge Endpoint

The main endpoint expects multiple DOCX files to be uploaded as form data:

```bash
curl -X POST "http://localhost:8000/merge-docx/" \
  -F "files=@document1.docx" \
  -F "files=@document2.docx" \
  -F "files=@document3.docx"
```

Requirements:
- At least 2 DOCX files
- Maximum 10 files per request
- All files must be valid DOCX documents

## Deployment Notes

This API can be deployed to:
- Linux servers
- Docker containers
- AWS / GCP / Azure
- Render
- Railway
- VPS

The implementation uses temporary files in `/tmp` which is compatible with most serverless platforms, including those with ephemeral filesystems.

## Security & Validation

- Validates that uploaded files are actually DOCX files
- Limits the number of files to prevent abuse
- Automatically cleans up temporary files after processing
- Uses UUIDs to prevent filename collisions