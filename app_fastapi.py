import os
import sys
import tempfile
import subprocess
import base64
import pandas as pd
import numpy as np
from typing import List

from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# Try to import log_universal from poc_db, failing gracefully if not present
try:
    from database.poc_db import log_universal as _log_uni
except ImportError:
    _log_uni = None

app = FastAPI(title="Invoice to Excel POC API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def api_upload_files(
    request: Request,
    files: List[UploadFile] = File(..., alias="files[]")
):
    if not files or files[0].filename == '':
        return JSONResponse({"error": "No selected files"}, status_code=400)
        
    processed_by = request.headers.get("X-User-Email") or request.headers.get("x-user-email") or "SYSTEM"
    file_names = ", ".join([f.filename for f in files])
    
    if _log_uni:
        _log_uni(
            module="Invoice-to-Excel",
            action="extract",
            file_name=file_names,
            status="STARTED",
            details="Starting invoice extraction",
            processed_by=processed_by
        )
        
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        script_path = os.path.join(base_dir, "extract_invoices_to_excel (1).py")
        
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = os.path.join(temp_dir, "input")
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(input_dir, exist_ok=True)
            os.makedirs(output_dir, exist_ok=True)
            
            for file in files:
                file_path = os.path.join(input_dir, file.filename)
                content = await file.read()
                with open(file_path, "wb") as f:
                    f.write(content)
                
            output_excel = os.path.join(output_dir, "Astrya_Invoices.xlsx")
            
            # Subprocess runs the extraction script
            workspace_venv = os.path.join(base_dir, "..", "..", "venv", "Scripts", "python.exe")
            python_exe = workspace_venv if os.path.exists(workspace_venv) else sys.executable
            
            result = subprocess.run(
                [python_exe, script_path, "--pdf_dir", input_dir, "--output", output_excel],
                capture_output=True,
                text=True
            )
            
            if result.returncode != 0:
                if _log_uni:
                    _log_uni(
                        module="Invoice-to-Excel",
                        action="extract",
                        file_name=file_names,
                        status="FAILED",
                        details=f"Subprocess failed: {result.stderr[:200]}",
                        processed_by=processed_by
                    )
                return JSONResponse({"error": f"Error processing files. Script output: {result.stderr}"}, status_code=500)
                
            if not os.path.exists(output_excel):
                if _log_uni:
                    _log_uni(
                        module="Invoice-to-Excel",
                        action="extract",
                        file_name=file_names,
                        status="FAILED",
                        details="Output Excel file was not generated.",
                        processed_by=processed_by
                    )
                return JSONResponse({"error": "Output Excel file was not generated."}, status_code=500)
                
            df = pd.read_excel(output_excel)
            df = df.replace({np.nan: None})
            records = df.to_dict(orient="records")
            
            with open(output_excel, "rb") as f:
                excel_base64 = base64.b64encode(f.read()).decode('utf-8')
                
            if _log_uni:
                _log_uni(
                    module="Invoice-to-Excel",
                    action="extract",
                    file_name=file_names,
                    status="SUCCESS",
                    details=f"Successfully extracted {len(records)} records",
                    processed_by=processed_by
                )
                
            return {
                "records": records,
                "excel_base64": excel_base64
            }
            
    except Exception as e:
        if _log_uni:
            _log_uni(
                module="Invoice-to-Excel",
                action="extract",
                file_name=file_names,
                status="FAILED",
                details=f"Exception: {str(e)[:200]}",
                processed_by=processed_by
            )
        return JSONResponse({"error": str(e)}, status_code=500)
