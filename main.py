from fastapi import FastAPI, UploadFile, File, Form, BackgroundTasks, Header, HTTPException
from fastapi.responses import FileResponse, JSONResponse
import pandas as pd
import tempfile
import os
import json
from typing import List
from pptx import Presentation
from fastapi.middleware.cors import CORSMiddleware

API_KEY = os.getenv("POWERIO_API_KEY")

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def read_root():
    return {"message": "PowerIO API is running"}

@app.post("/process")
async def process_files(
    background_tasks: BackgroundTasks,
    api_key: str = Header(None, alias="X-API-KEY"),
    pptxFile: UploadFile = File(...),
    dataFiles: List[UploadFile] = File(...),
    instructions: str = Form(...)
):
    if api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

    pptx_temp = None
    data_temps = []
    output_temp = None
    
    try:
        # Save uploaded PPTX to temporary file
        pptx_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        pptx_content = await pptxFile.read()
        pptx_temp.write(pptx_content)
        pptx_temp.close()
        
        # Load the PowerPoint using Presentation
        prs = Presentation(pptx_temp.name)
        
        # Parse the instruction metadata
        instructions_data = json.loads(instructions)
        
        # Validate that we have matching counts
        if len(instructions_data) != len(dataFiles):
            raise ValueError(f"Number of instructions ({len(instructions_data)}) does not match number of data files ({len(dataFiles)})")
        
        # Build instruction objects with DataFrames
        instruction_objects = []
        for idx, instruction_meta in enumerate(instructions_data):
            # Get the corresponding data file
            data_file = dataFiles[idx]
            
            # Save data file to temporary file
            data_temp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(data_file.filename)[1])
            data_content = await data_file.read()
            data_temp.write(data_content)
            data_temp.close()
            data_temps.append(data_temp.name)
            
            # Load data file into pandas DataFrame (automatically detect CSV vs Excel)
            data_filename = data_file.filename.lower()
            if data_filename.endswith(".csv"):
                df = pd.read_csv(data_temp.name)
            else:
                # Excel files (.xlsx, .xls)
                df = pd.read_excel(data_temp.name)
            
            # Build instruction object
            instruction_obj = {
                "chart_type": instruction_meta["chart_type"],
                "slide_number": instruction_meta["slide_number"],
                "dataframe": df
            }
            instruction_objects.append(instruction_obj)
        
        
        # Create output temporary file
        output_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.pptx')
        output_temp.close()
        
        # Save the modified PPTX once
        prs.save(output_temp.name)
        
        # Schedule cleanup of output file after response is sent
        background_tasks.add_task(os.unlink, output_temp.name)
        
        # Return the file back to the client with FileResponse
        return FileResponse(
            output_temp.name,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="output.pptx"
        )
        
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )
    finally:
        # Cleanup temporary input files
        if pptx_temp and os.path.exists(pptx_temp.name):
            os.unlink(pptx_temp.name)
        for data_temp_path in data_temps:
            if os.path.exists(data_temp_path):
                os.unlink(data_temp_path)
        # Note: output_temp is cleaned up by FileResponse after download
