#!/usr/bin/env python3
"""
Script to generate openapi.json from FastAPI application
"""
import json
from main import app

# Generate OpenAPI schema
openapi_schema = app.openapi()

# Write to openapi.json
with open("openapi.json", "w") as f:
    json.dump(openapi_schema, f, indent=2)

print("✅ openapi.json generated successfully!")

