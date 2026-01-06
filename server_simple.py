"""
PDF to PowerPoint Web Application - JSON Mode Backend (SIMPLE VERSION)
Minimal server to test download functionality
"""
import os
import shutil
import json
import uuid
import asyncio
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import fitz  # PyMuPDF

import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

app = FastAPI(
    title="PDF to PowerPoint Converter (JSON Mode) - Simple",
    description="Minimal version for testing",
    version="1.0.0"
)

# CORS for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Storage for job status
jobs = {}

# Directories
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
TEMP_DIR = BASE_DIR / "temp_processing"

# Create directories
for d in [UPLOAD_DIR, OUTPUT_DIR, TEMP_DIR]:
    d.mkdir(exist_ok=True)


class JobStatus:
    PENDING = "pending"
    PROCESSING = "processing"
    GENERATING = "generating"
    COMPLETED = "completed"
    ERROR = "error"


@app.get("/")
async def root():
    return {"message": "PDF to PPTX Converter API (JSON Mode - Simple)", "status": "running"}


@app.post("/api/upload")
async def upload_files(
    pdf_file: UploadFile = File(...),
    json_file: UploadFile = File(...),
    mode: str = Form("precision")
):
    """Upload a PDF file and its corresponding JSON analysis file"""
    if not pdf_file.filename.endswith('.pdf'):
        raise HTTPException(status_code=400, detail="First file must be a PDF")
    
    if not json_file.filename.endswith('.json'):
        raise HTTPException(status_code=400, detail="Second file must be a JSON file")
    
    job_id = str(uuid.uuid4())
    job_dir = TEMP_DIR / job_id
    job_dir.mkdir(exist_ok=True)
    
    # Save PDF
    pdf_path = job_dir / "input.pdf"
    with open(pdf_path, "wb") as f:
        content = await pdf_file.read()
        f.write(content)
    
    # Save JSON
    json_path = job_dir / "image_analysis.json"
    with open(json_path, "wb") as f:
        content = await json_file.read()
        f.write(content)
    
    jobs[job_id] = {
        "status": JobStatus.PENDING,
        "progress": 0,
        "message": "Files uploaded successfully",
        "pdf_path": str(pdf_path),
        "json_path": str(json_path),
        "job_dir": str(job_dir),
        "original_filename": pdf_file.filename,
        "total_pages": 0,
        "current_page": 0,
        "mode": mode
    }
    
    return {"job_id": job_id, "message": "Upload successful", "mode": mode}


@app.post("/api/process/{job_id}")
async def start_processing(job_id: str, background_tasks: BackgroundTasks):
    """Start processing the uploaded PDF with the provided JSON"""
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    background_tasks.add_task(process_pdf_with_json, job_id)
    jobs[job_id]["status"] = JobStatus.PROCESSING
    jobs[job_id]["message"] = "Processing started"
    
    return {"status": "processing", "message": "Processing started"}


async def process_pdf_with_json(job_id: str):
    """Background task to process PDF with pre-analyzed JSON"""
    try:
        job = jobs[job_id]
        job_dir = Path(job["job_dir"])
        pdf_path = Path(job["pdf_path"])
        json_path = Path(job["json_path"])
        
        # Step 1: Convert PDF to images
        job["status"] = JobStatus.PROCESSING
        job["message"] = "Converting PDF to images..."
        
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        job["total_pages"] = total_pages
        
        pages_dir = job_dir / "pages"
        pages_dir.mkdir(exist_ok=True)
        
        page_width = doc[0].rect.width
        page_height = doc[0].rect.height
        job["page_width"] = page_width
        job["page_height"] = page_height
        
        for page_num in range(total_pages):
            page = doc[page_num]
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            img_path = pages_dir / f"page_{page_num + 1}.png"
            pix.save(str(img_path))
            
            job["current_page"] = page_num + 1
            job["progress"] = int((page_num + 1) / total_pages * 30)
        
        doc.close()
        
        job["progress"] = 35
        
        # Step 2: Generate PowerPoint using standalone converter
        job["status"] = JobStatus.GENERATING
        job["message"] = "Generating PowerPoint..."
        
        output_filename = Path(job["original_filename"]).stem + ".pptx"
        output_path = OUTPUT_DIR / f"{job_id}_{output_filename}"
        
        # ONLY use standalone_convert.py (no mode selection)
        converter_script = BASE_DIR / "standalone_convert.py"
        print(f"Using converter: {converter_script}")
        
        log_path = job_dir / "conversion_log.txt"
        
        # Run converter as subprocess
        import subprocess
        cmd = [
            sys.executable,
            str(converter_script),
            "--pdf", str(pdf_path),
            "--output", str(output_path),
            "--json", str(json_path),
            "--log", str(log_path)
        ]
        
        print(f"Running converter: {' '.join(cmd)}")
        
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            cwd=str(BASE_DIR)
        )
        
        # Wait for completion with progress updates
        while process.poll() is None:
            job["progress"] = min(95, job["progress"] + 1)
            await asyncio.sleep(2)
        
        returncode = process.returncode
        if returncode != 0:
            stderr = process.stderr.read().decode("utf-8", errors="replace")
            raise Exception(f"Converter failed with code {returncode}: {stderr[:500]}")
        
        job["status"] = JobStatus.COMPLETED
        job["progress"] = 100
        job["message"] = "Conversion completed!"
        job["output_path"] = str(output_path)
        job["output_filename"] = output_filename
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        jobs[job_id]["status"] = JobStatus.ERROR
        jobs[job_id]["message"] = str(e)
        jobs[job_id]["progress"] = 0


@app.get("/api/status/{job_id}")
async def get_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    return jobs[job_id]


@app.get("/api/download/{job_id}")
async def download_result(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if job["status"] != JobStatus.COMPLETED:
        raise HTTPException(status_code=400, detail="Processing not completed")
    
    output_path = Path(job["output_path"])
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Output file not found")
    
    return FileResponse(
        path=output_path,
        filename=job["output_filename"],
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


@app.delete("/api/job/{job_id}")
async def cleanup_job(job_id: str):
    if job_id in jobs:
        job_dir = TEMP_DIR / job_id
        if job_dir.exists():
            shutil.rmtree(job_dir)
        del jobs[job_id]
    return {"message": "Job cleaned up"}


# Serve static files
if (BASE_DIR / "static").exists():
    app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")


if __name__ == "__main__":
    import uvicorn
    print("Starting PDF to PPTX Converter (JSON Mode - SIMPLE VERSION)")
    uvicorn.run(app, host="0.0.0.0", port=8001)
