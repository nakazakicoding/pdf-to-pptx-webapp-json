"""
PDF to PowerPoint Web Application - JSON Mode Backend (LIGHTWEIGHT VERSION)
FastAPI server for PDF conversion with pre-analyzed JSON (no Gemini API required)
Optimized for low-memory environments (512MB RAM)
"""
import os
import shutil
import json
import uuid
import asyncio
import gc
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import fitz  # PyMuPDF

# Import conversion modules
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

app = FastAPI(
    title="PDF to PowerPoint Converter (JSON Mode)",
    description="Convert PDF files to PowerPoint using pre-analyzed JSON files (no API required)",
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
    return {"message": "PDF to PowerPoint Converter API (JSON Mode)", "status": "running", "mode": "json_upload"}


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
    
    # Validate JSON structure
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            json_data = json.load(f)
        if not isinstance(json_data, dict):
            raise ValueError("JSON must be an object with page keys")
    except json.JSONDecodeError as e:
        shutil.rmtree(job_dir)
        raise HTTPException(status_code=400, detail=f"Invalid JSON format: {str(e)}")
    except ValueError as e:
        shutil.rmtree(job_dir)
        raise HTTPException(status_code=400, detail=str(e))
    
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
        
        # LIGHTWEIGHT: Use 1.5x scale instead of 2.0x (saves ~40% memory)
        for page_num in range(total_pages):
            page = doc[page_num]
            mat = fitz.Matrix(1.5, 1.5)  # Reduced from 2.0
            pix = page.get_pixmap(matrix=mat)
            img_path = pages_dir / f"page_{page_num + 1}.png"
            pix.save(str(img_path))
            
            # Clean up pixmap to free memory immediately
            del pix
            gc.collect()
            
            job["current_page"] = page_num + 1
            job["progress"] = int((page_num + 1) / total_pages * 30)
        
        doc.close()
        del doc
        gc.collect()
        
        job["progress"] = 35
        
        # Step 2: Generate PowerPoint using standalone converter
        job["status"] = JobStatus.GENERATING
        job["message"] = "Generating PowerPoint..."
        
        output_filename = Path(job["original_filename"]).stem + ".pptx"
        output_path = OUTPUT_DIR / f"{job_id}_{output_filename}"
        
        # Select converter based on mode - LIGHTWEIGHT 2x versions (full quality + page cleanup)
        mode = job.get("mode", "precision")
        if mode == "safeguard":
            converter_script = BASE_DIR / "standalone_convert_v4_v43_light_2x.py"
            print(f"Using Safeguard Mode converter (v43 LIGHT 2x)")
        else:
            converter_script = BASE_DIR / "standalone_convert_v43_light_2x.py"
            print(f"Using Precision Mode converter (v43 LIGHT 2x)")
        
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
    from urllib.parse import quote
    
    # Try to find job in memory first
    if job_id in jobs:
        job = jobs[job_id]
        if job["status"] != JobStatus.COMPLETED:
            raise HTTPException(status_code=400, detail="Processing not completed")
        
        output_path = Path(job["output_path"])
        output_filename = job["output_filename"]
    else:
        # Fallback: Search for file in output directory matching job_id
        possible_files = list(OUTPUT_DIR.glob(f"{job_id}*"))
        if not possible_files:
            raise HTTPException(status_code=404, detail="Job not found and no matching file in output directory")
        
        output_path = possible_files[0]
        output_filename = output_path.name
    
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Output file not found")
    
    # Handle Japanese filenames for download display
    encoded_filename = quote(output_filename)
    
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
        }
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
    print("Starting PDF to PPTX Converter (JSON Mode)")
    uvicorn.run(app, host="0.0.0.0", port=8001)
