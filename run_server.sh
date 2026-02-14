#!/bin/bash
# Script to run the DocMerge API server

# Activate the virtual environment
source docmerge_env/bin/activate

# Run the FastAPI server with uvicorn
echo "Starting DocMerge API server..."
uvicorn main:app --host 0.0.0.0 --port 8000 --reload